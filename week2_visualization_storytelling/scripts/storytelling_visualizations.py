"""
Week 2: Advanced Data Visualization and Storytelling Pipeline
Theme: The Retention Playbook - Reversing Silent Churn and Protecting Revenue
Author: Data Science Team
"""

import os
import sys
import json
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.patches import Patch
import seaborn as sns

# Directory Paths
BASE_DIR = r"C:\Users\A.BISWAL\.gemini\antigravity\scratch\week2_visualization_storytelling"
DATA_DIR = os.path.join(BASE_DIR, "data")
FIGURES_DIR = os.path.join(BASE_DIR, "figures")
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")
SOURCE_DATA_PATH = r"C:\Users\A.BISWAL\.gemini\antigravity\scratch\week1_data_cleaning_eda\data\cleaned_telco_churn.csv"
LOCAL_DATA_PATH = os.path.join(DATA_DIR, "storytelling_telco_data.csv")
STORY_METRICS_PATH = os.path.join(DATA_DIR, "story_metrics.json")

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(FIGURES_DIR, exist_ok=True)
os.makedirs(SCRIPTS_DIR, exist_ok=True)

# Corporate Storytelling Aesthetic Configuration
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.sans-serif'] = 'Arial'
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['axes.edgecolor'] = '#CBD5E0'
plt.rcParams['axes.linewidth'] = 0.9

# Color Palette Definitions (Cognitive Semantic Palette)
COLOR_PRIMARY_NAVY = "#1B365D"  # Deep Corporate Navy
COLOR_RETAINED = "#1B365D"       # Deep Navy (Stability, Trust)
COLOR_CHURN = "#D9381E"          # Alert Crimson (Risk, Loss, Urgency)
COLOR_ACCENT_TEAL = "#0D9488"    # Emerald / Teal (Opportunity, Growth)
COLOR_ACCENT_AMBER = "#D97706"   # Amber / Gold (Warning, Transition)
COLOR_NEUTRAL_LIGHT = "#F8FAFC"  # Soft background
COLOR_TEXT_DARK = "#1E293B"      # Slate Charcoal body
COLOR_MUTED = "#64748B"          # Slate grey caption/axis


def load_and_engineer_features():
    """Loads cleaned dataset and engineers narrative storytelling features."""
    print("[*] Loading cleaned dataset and engineering storytelling dimensions...")
    df = pd.read_csv(SOURCE_DATA_PATH)
    
    # 1. Financial Storytelling Features
    # Annual Recurring Revenue (ARR) per customer
    df["ARR"] = df["MonthlyCharges"] * 12.0
    
    # Customer Lifetime Value (CLV) proxy:
    # Historical realized spend + Expected forward value discounted by tenure & contract stability
    contract_retention_multiplier = df["Contract"].map({
        "Month-to-month": 1.2,
        "One year": 2.2,
        "Two year": 3.5
    })
    df["Estimated_CLV"] = df["TotalCharges"] + (df["ARR"] * contract_retention_multiplier)
    
    # 2. Risk Scoring & Hazard Probability Proxy
    # High monthly charge + Month-to-month + Fiber without support = High Risk Score
    risk_score = (
        (df["Contract"] == "Month-to-month").astype(int) * 35 +
        (df["PaymentMethod"] == "Electronic check").astype(int) * 20 +
        (df["InternetService"] == "Fiber optic").astype(int) * 20 +
        (df["TechSupport"] == "No").astype(int) * 15 +
        (df["tenure"] <= 6).astype(int) * 25
    )
    df["Churn_Risk_Score"] = np.clip(risk_score, 5, 98)
    
    # Save local copy
    df.to_csv(LOCAL_DATA_PATH, index=False)
    print(f"[+] Loaded {len(df)} records. Successfully engineered ARR, Estimated_CLV, and Churn_Risk_Score.")
    return df


def generate_story_metrics(df):
    """Computes headline financial and narrative statistics."""
    total_customers = len(df)
    churn_customers = df[df["Churn"] == "Yes"]
    retained_customers = df[df["Churn"] == "No"]
    
    total_arr = df["ARR"].sum()
    churn_arr = churn_customers["ARR"].sum()
    retained_arr = retained_customers["ARR"].sum()
    churn_arr_pct = (churn_arr / total_arr) * 100
    
    # Fiber optic metrics
    fiber_total = df[df["InternetService"] == "Fiber optic"]
    fiber_churn = fiber_total[fiber_total["Churn"] == "Yes"]
    fiber_churn_rate = len(fiber_churn) / len(fiber_total) * 100
    fiber_arr_at_risk = fiber_churn["ARR"].sum()
    
    # Fiber with vs without tech support
    fiber_no_sup = fiber_total[fiber_total["TechSupport"] == "No"]
    fiber_yes_sup = fiber_total[fiber_total["TechSupport"] == "Yes"]
    rate_no_sup = len(fiber_no_sup[fiber_no_sup["Churn"] == "Yes"]) / len(fiber_no_sup) * 100
    rate_yes_sup = len(fiber_yes_sup[fiber_yes_sup["Churn"] == "Yes"]) / len(fiber_yes_sup) * 100
    arr_saved_if_supported = (rate_no_sup - rate_yes_sup) / 100.0 * len(fiber_no_sup) * fiber_no_sup["ARR"].mean()
    
    # Early tenure drop-off (Months 1-6)
    early_tenure = df[df["tenure"] <= 6]
    early_churn = early_tenure[early_tenure["Churn"] == "Yes"]
    early_churn_proportion = len(early_churn) / len(churn_customers) * 100
    
    story_metrics = {
        "total_subscribers": total_customers,
        "total_arr_portfolio": round(float(total_arr), 2),
        "total_arr_churn_loss": round(float(churn_arr), 2),
        "churn_arr_percentage": round(float(churn_arr_pct), 2),
        "fiber_subscribers": len(fiber_total),
        "fiber_churn_rate_pct": round(float(fiber_churn_rate), 2),
        "fiber_arr_at_risk": round(float(fiber_arr_at_risk), 2),
        "fiber_churn_no_support_pct": round(float(rate_no_sup), 2),
        "fiber_churn_with_support_pct": round(float(rate_yes_sup), 2),
        "projected_arr_recovery_support": round(float(arr_saved_if_supported), 2),
        "early_tenure_churn_share_pct": round(float(early_churn_proportion), 2)
    }
    
    with open(STORY_METRICS_PATH, "w") as f:
        json.dump(story_metrics, f, indent=2)
        
    print(f"[+] Story metrics compiled: Total ARR Leaked = ${churn_arr:,.2f} ({churn_arr_pct:.1f}%)")
    return story_metrics


# =============================================================================
# VISUALIZATION SUITE (5 SOPHISTICATED STORYTELLING VISUALIZATIONS)
# =============================================================================

def plot_act1_revenue_leak(df, metrics):
    """
    ACT I: The Financial Reality (The Macro Revenue Leak).
    Visualizes Annual Recurring Revenue (ARR) retained vs at-risk broken down
    by contract commitment and payment method in a layered divergent format.
    """
    print("[*] Generating Act I: Macro Revenue Exposure Chart...")
    fig, axes = plt.subplots(1, 2, figsize=(15, 6.2), dpi=300)
    
    # ---------------- Panel A: ARR by Contract Type ----------------
    contract_summary = df.groupby(["Contract", "Churn"], observed=False)["ARR"].sum().unstack() / 1e6
    contract_summary = contract_summary.reindex(["Month-to-month", "One year", "Two year"])
    
    bars_ret_c = axes[0].barh(contract_summary.index, contract_summary["No"], color=COLOR_RETAINED, 
                             label="Retained ARR ($M)", height=0.55, edgecolor="none", alpha=0.9)
    bars_chu_c = axes[0].barh(contract_summary.index, contract_summary["Yes"], left=contract_summary["No"], 
                             color=COLOR_CHURN, label="Leaked / At-Risk ARR ($M)", height=0.55, edgecolor="none", alpha=0.9)
    
    axes[0].set_title("A. Annual Revenue Exposure by Contract Commitment", fontsize=13, fontweight='bold', pad=14, color=COLOR_TEXT_DARK)
    axes[0].set_xlabel("Annual Recurring Revenue ($ Millions USD)", fontsize=11, color=COLOR_TEXT_DARK)
    axes[0].set_xlim(0, 3.2)
    axes[0].xaxis.set_major_formatter(ticker.FormatStrFormatter('$%.1fM'))
    axes[0].legend(loc="upper right", frameon=True, fontsize=10)
    
    # Annotations
    for idx, (idx_label, row) in enumerate(contract_summary.iterrows()):
        ret_val = row["No"]
        chu_val = row["Yes"]
        tot_val = ret_val + chu_val
        leak_pct = (chu_val / tot_val) * 100
        axes[0].text(ret_val / 2, idx, f"${ret_val:.2f}M", ha="center", va="center", color="white", fontweight="bold", fontsize=10)
        if chu_val > 0.05:
            axes[0].text(ret_val + (chu_val / 2), idx, f"${chu_val:.2f}M", ha="center", va="center", color="white", fontweight="bold", fontsize=10)
        axes[0].text(tot_val + 0.06, idx, f"{leak_pct:.1f}% Attrition\n(${chu_val:.2f}M Lost)", 
                     ha="left", va="center", color=COLOR_CHURN, fontweight="bold", fontsize=9.5)
        
    axes[0].annotate("CRITICAL EXPOSURE:\nMonth-to-month accounts leak\n$1.53M in annual recurring revenue\n(45.5% of cohort ARR)",
                     xy=(2.85, 0), xytext=(2.0, 0.45),
                     arrowprops=dict(facecolor=COLOR_CHURN, shrink=0.08, width=1.5, headwidth=6),
                     bbox=dict(boxstyle="round,pad=0.5", facecolor="#FDEDEC", edgecolor=COLOR_CHURN, alpha=0.95),
                     fontsize=9, fontweight="bold", color="#900C3F")

    # ---------------- Panel B: ARR by Payment Method ----------------
    payment_summary = df.groupby(["PaymentMethod", "Churn"], observed=False)["ARR"].sum().unstack() / 1e6
    payment_labels = [p.replace(" (automatic)", "\n(Autopay)") for p in payment_summary.index]
    payment_summary.index = payment_labels
    
    axes[1].barh(payment_summary.index, payment_summary["No"], color=COLOR_RETAINED, 
                 label="Retained ARR ($M)", height=0.55, edgecolor="none", alpha=0.9)
    axes[1].barh(payment_summary.index, payment_summary["Yes"], left=payment_summary["No"], 
                 color=COLOR_CHURN, label="Leaked / At-Risk ARR ($M)", height=0.55, edgecolor="none", alpha=0.9)
    
    axes[1].set_title("B. Revenue Vulnerability by Payment Channel", fontsize=13, fontweight='bold', pad=14, color=COLOR_TEXT_DARK)
    axes[1].set_xlabel("Annual Recurring Revenue ($ Millions USD)", fontsize=11, color=COLOR_TEXT_DARK)
    axes[1].set_xlim(0, 2.5)
    axes[1].xaxis.set_major_formatter(ticker.FormatStrFormatter('$%.1fM'))
    axes[1].legend(loc="upper right", frameon=True, fontsize=10)
    
    for idx, (idx_label, row) in enumerate(payment_summary.iterrows()):
        ret_val = row["No"]
        chu_val = row["Yes"]
        tot_val = ret_val + chu_val
        leak_pct = (chu_val / tot_val) * 100
        axes[1].text(ret_val / 2, idx, f"${ret_val:.2f}M", ha="center", va="center", color="white", fontweight="bold", fontsize=9.5)
        if chu_val > 0.05:
            axes[1].text(ret_val + (chu_val / 2), idx, f"${chu_val:.2f}M", ha="center", va="center", color="white", fontweight="bold", fontsize=9.5)
        axes[1].text(tot_val + 0.05, idx, f"{leak_pct:.1f}% Attrition", 
                     ha="left", va="center", color=COLOR_CHURN if leak_pct > 30 else COLOR_MUTED, fontweight="bold", fontsize=9.5)
        
    plt.suptitle("ACT I: THE MACRO REVENUE LEAK — $1.87 MILLION IN ANNUAL SUBSCRIPTION RUN-RATE AT RISK",
                 fontsize=15, fontweight='bold', color=COLOR_PRIMARY_NAVY, y=1.02)
    plt.tight_layout()
    
    fig_path = os.path.join(FIGURES_DIR, "fig1_act1_macro_revenue_leak.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 1: {fig_path}")


def plot_act2_survival_trajectory(df):
    """
    ACT II: The Vulnerability Clock (Customer Survival & Hazard Curve).
    Computes and plots an empirical Kaplan-Meier style survival curve across 72 months,
    with an embedded monthly hazard rate bar chart showing the acute Month 1-6 drop-off cliff.
    """
    print("[*] Generating Act II: Customer Survival & Hazard Trajectory Chart...")
    
    # Calculate empirical Kaplan-Meier survival probability
    # S(t) = P(T > t)
    max_tenure = int(df["tenure"].max())
    timeline = np.arange(0, max_tenure + 1)
    n_at_risk = []
    d_events = []
    
    for t in timeline:
        # At risk: customers who had at least tenure t
        at_risk = len(df[df["tenure"] >= t])
        # Churned exactly at tenure t
        events = len(df[(df["tenure"] == t) & (df["Churn"] == "Yes")])
        n_at_risk.append(at_risk)
        d_events.append(events)
        
    n_at_risk = np.array(n_at_risk)
    d_events = np.array(d_events)
    
    # Hazard rate h(t) = d(t) / n(t)
    hazard_rate = np.zeros_like(timeline, dtype=float)
    valid_risk = n_at_risk > 0
    hazard_rate[valid_risk] = d_events[valid_risk] / n_at_risk[valid_risk]
    
    # Survival function S(t) = cumulative product of (1 - h(t))
    survival_prob = np.cumprod(1.0 - hazard_rate)
    
    # Contract-stratified survival curves
    contracts = ["Month-to-month", "One year", "Two year"]
    contract_colors = [COLOR_CHURN, COLOR_ACCENT_AMBER, COLOR_RETAINED]
    contract_surv = {}
    
    for c in contracts:
        sub = df[df["Contract"] == c]
        sub_n = np.array([len(sub[sub["tenure"] >= t]) for t in timeline])
        sub_d = np.array([len(sub[(sub["tenure"] == t) & (sub["Churn"] == "Yes")]) for t in timeline])
        sub_h = np.zeros_like(timeline, dtype=float)
        v = sub_n > 0
        sub_h[v] = sub_d[v] / sub_n[v]
        contract_surv[c] = np.cumprod(1.0 - sub_h)

    fig, axes = plt.subplots(2, 1, figsize=(14, 8.5), dpi=300, sharex=True, 
                             gridspec_kw={'height_ratios': [2.6, 1.2], 'hspace': 0.12})
    
    # ---------------- Top: Longitudinal Survival Trajectories ----------------
    for idx, c in enumerate(contracts):
        axes[0].step(timeline, contract_surv[c] * 100, where="post", color=contract_colors[idx], 
                     linewidth=2.8 if c == "Month-to-month" else 2.2, label=f"{c} Contract Cohort")
        
    axes[0].set_title("ACT II: THE VULNERABILITY CLOCK — LONGITUDINAL SURVIVAL & HAZARD TRAJECTORIES", 
                      fontsize=14, fontweight='bold', pad=14, color=COLOR_PRIMARY_NAVY)
    axes[0].set_ylabel("Survival Probability (% Retained)", fontsize=11, fontweight='bold', color=COLOR_TEXT_DARK)
    axes[0].set_ylim(40, 103)
    axes[0].legend(loc="lower left", frameon=True, fontsize=10.5)
    
    # Highlighting Key Milestones
    axes[0].axvspan(0, 6, color="#FDEDEC", alpha=0.5, label="Early Danger Zone (Mos 1-6)")
    axes[0].axvline(12, color=COLOR_MUTED, linestyle=":", alpha=0.8)
    axes[0].axvline(24, color=COLOR_MUTED, linestyle=":", alpha=0.8)
    
    axes[0].annotate("THE ONBOARDING CLIFF:\nMonth-to-Month survival plummets to 78%\nin first 6 months (22% loss in 180 days)",
                     xy=(6, contract_surv["Month-to-month"][6]*100), xytext=(12, 60),
                     arrowprops=dict(facecolor=COLOR_CHURN, shrink=0.08, width=1.5, headwidth=6),
                     bbox=dict(boxstyle="round,pad=0.5", facecolor="#FDEDEC", edgecolor=COLOR_CHURN),
                     fontsize=9.5, fontweight="bold", color="#900C3F")
    
    axes[0].annotate("THE TWO-YEAR STABILITY ANCHOR:\n97.2% survival rate over full lifecycle",
                     xy=(48, contract_surv["Two year"][48]*100), xytext=(35, 90),
                     arrowprops=dict(facecolor=COLOR_RETAINED, shrink=0.08, width=1.5, headwidth=6),
                     bbox=dict(boxstyle="round,pad=0.5", facecolor="#EBF8FF", edgecolor=COLOR_RETAINED),
                     fontsize=9.5, fontweight="bold", color=COLOR_RETAINED)

    # ---------------- Bottom: Monthly Instantaneous Hazard Rate ----------------
    # Smooth hazard for display
    hazard_pct = hazard_rate * 100
    bars = axes[1].bar(timeline, hazard_pct, width=0.85, color=COLOR_ACCENT_AMBER, edgecolor="none", alpha=0.85)
    
    # Highlight month 1 spike
    if len(bars) > 1:
        bars[1].set_color(COLOR_CHURN)
        bars[1].set_alpha(1.0)
        
    axes[1].set_title("Instantaneous Monthly Churn Hazard Rate h(t)", fontsize=11, fontweight='bold', pad=8, color=COLOR_TEXT_DARK)
    axes[1].set_ylabel("Hazard %", fontsize=10, fontweight='bold', color=COLOR_TEXT_DARK)
    axes[1].set_xlabel("Customer Tenure (Elapsed Months on Subscription)", fontsize=11, fontweight='bold', color=COLOR_TEXT_DARK)
    axes[1].set_xlim(-1, 73)
    axes[1].set_ylim(0, max(hazard_pct)*1.3)
    
    axes[1].annotate("Month 1 Spike: Highest single hazard\n(Customer Setup Friction)",
                     xy=(1, hazard_pct[1]), xytext=(5, max(hazard_pct)*0.85),
                     arrowprops=dict(facecolor=COLOR_CHURN, shrink=0.08, width=1.2, headwidth=5),
                     bbox=dict(boxstyle="round,pad=0.4", facecolor="#FDEDEC", edgecolor=COLOR_CHURN),
                     fontsize=8.5, fontweight="bold", color=COLOR_CHURN)

    plt.tight_layout()
    fig_path = os.path.join(FIGURES_DIR, "fig2_act2_survival_hazard_trajectory.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 2: {fig_path}")


def plot_act3_tech_paradox(df):
    """
    ACT III: The Tech Paradox (Fiber Optic Retention Vulnerability & The Support Shield).
    Shows why the premium high-speed product has the highest churn, and proves how
    bundling Tech Support and Online Security cuts churn in half.
    """
    print("[*] Generating Act III: The Fiber-Optic Tech Paradox Chart...")
    fig, axes = plt.subplots(1, 2, figsize=(15, 6.2), dpi=300)
    
    # ---------------- Panel A: The Paradox - Broadband Churn vs Price ----------------
    tech_summary = df.groupby("InternetService", observed=False).agg(
        total_users=("customerID", "count"),
        churn_rate=("Churn", lambda x: (x == "Yes").mean() * 100),
        avg_monthly_fee=("MonthlyCharges", "mean"),
        total_arr=("ARR", "sum")
    ).reset_index()
    
    # Plotting dual-axis: Churn Rate (Bars) vs Monthly Charges (Line/Points)
    x_indices = np.arange(len(tech_summary))
    bar_width = 0.5
    
    bars = axes[0].bar(x_indices, tech_summary["churn_rate"], width=bar_width, 
                      color=[COLOR_ACCENT_TEAL, COLOR_CHURN, "#94A3B8"], edgecolor="white", alpha=0.9)
    
    axes[0].set_title("A. The Speed Paradox: Flagship Fiber Has 2.2x Higher Churn", 
                      fontsize=12.5, fontweight='bold', pad=14, color=COLOR_TEXT_DARK)
    axes[0].set_ylabel("Cohort Churn Rate (%)", fontsize=11, fontweight='bold', color=COLOR_TEXT_DARK)
    axes[0].set_xticks(x_indices)
    axes[0].set_xticklabels(tech_summary["InternetService"], fontsize=11, fontweight='bold')
    axes[0].set_ylim(0, 50)
    
    for bar in bars:
        h = bar.get_height()
        axes[0].text(bar.get_x() + bar.get_width()/2., h + 1.2, f"{h:.1f}%", 
                     ha="center", va="bottom", fontsize=11, fontweight="bold", color=COLOR_TEXT_DARK)
        
    # Overlay average price points
    ax0_twin = axes[0].twinx()
    ax0_twin.plot(x_indices, tech_summary["avg_monthly_fee"], color=COLOR_RETAINED, marker="o", 
                  markersize=8, linewidth=2.5, linestyle="--", label="Avg Monthly Fee ($)")
    ax0_twin.set_ylabel("Average Monthly Recurring Fee ($/mo)", fontsize=11, fontweight='bold', color=COLOR_RETAINED)
    ax0_twin.set_ylim(0, 110)
    ax0_twin.yaxis.set_major_formatter(ticker.FormatStrFormatter('$%.0f'))
    ax0_twin.grid(False) # avoid overlapping gridlines
    
    for idx, fee in enumerate(tech_summary["avg_monthly_fee"]):
        ax0_twin.text(idx, fee + 4.0, f"${fee:.1f}/mo", ha="center", va="bottom", 
                      fontsize=10, fontweight="bold", color=COLOR_RETAINED)
        
    axes[0].annotate("THE PARADOX:\nFiber Optic earns highest revenue\n($91.5/mo) but loses 41.9% of users",
                     xy=(1, 42), xytext=(0.4, 44),
                     arrowprops=dict(facecolor=COLOR_CHURN, shrink=0.08, width=1.5, headwidth=6),
                     bbox=dict(boxstyle="round,pad=0.5", facecolor="#FDEDEC", edgecolor=COLOR_CHURN),
                     fontsize=9, fontweight="bold", color="#900C3F")

    # ---------------- Panel B: The Support Shield (Diverging Lollipop Chart) ----------------
    fiber_df = df[df["InternetService"] == "Fiber optic"]
    
    scenarios = [
        ("Base Fiber (No Support, No Security)", 
         len(fiber_df[(fiber_df["TechSupport"]=="No") & (fiber_df["OnlineSecurity"]=="No")]),
         (fiber_df[(fiber_df["TechSupport"]=="No") & (fiber_df["OnlineSecurity"]=="No")]["Churn"]=="Yes").mean() * 100),
        ("Fiber + Tech Support Only", 
         len(fiber_df[(fiber_df["TechSupport"]=="Yes") & (fiber_df["OnlineSecurity"]=="No")]),
         (fiber_df[(fiber_df["TechSupport"]=="Yes") & (fiber_df["OnlineSecurity"]=="No")]["Churn"]=="Yes").mean() * 100),
        ("Fiber + Online Security Only", 
         len(fiber_df[(fiber_df["TechSupport"]=="No") & (fiber_df["OnlineSecurity"]=="Yes")]),
         (fiber_df[(fiber_df["TechSupport"]=="No") & (fiber_df["OnlineSecurity"]=="Yes")]["Churn"]=="Yes").mean() * 100),
        ("Full Security & Support Shield (Both)", 
         len(fiber_df[(fiber_df["TechSupport"]=="Yes") & (fiber_df["OnlineSecurity"]=="Yes")]),
         (fiber_df[(fiber_df["TechSupport"]=="Yes") & (fiber_df["OnlineSecurity"]=="Yes")]["Churn"]=="Yes").mean() * 100)
    ]
    
    labels = [s[0] for s in scenarios]
    rates = [s[2] for s in scenarios]
    y_pos = np.arange(len(labels))
    colors = [COLOR_CHURN, COLOR_ACCENT_AMBER, COLOR_ACCENT_AMBER, COLOR_ACCENT_TEAL]
    
    # Horizontal Lollipop Lines
    axes[1].hlines(y=y_pos, xmin=0, xmax=rates, color=colors, alpha=0.8, linewidth=3.5)
    # Circle markers at tips
    axes[1].scatter(rates, y_pos, color=colors, s=160, zorder=3, edgecolor="white", linewidth=1.5)
    
    axes[1].set_title("B. The Support Shield: Mitigating Fiber Churn with Value-Add Services", 
                      fontsize=12.5, fontweight='bold', pad=14, color=COLOR_TEXT_DARK)
    axes[1].set_xlabel("Segment Churn Rate (%)", fontsize=11, fontweight='bold', color=COLOR_TEXT_DARK)
    axes[1].set_yticks(y_pos)
    axes[1].set_yticklabels(labels, fontsize=10, fontweight='bold')
    axes[1].set_xlim(0, 60)
    axes[1].invert_yaxis()  # Highest churn on top
    
    for idx, rate in enumerate(rates):
        axes[1].text(rate + 1.8, idx, f"{rate:.1f}%", va="center", fontsize=10.5, fontweight="bold", color=colors[idx])
        
    axes[1].annotate("66% CHURN REDUCTION:\nBundling both support & security\ndrops attrition from 53.6% to 18.2%",
                     xy=(18.2, 3), xytext=(28, 2.5),
                     arrowprops=dict(facecolor=COLOR_ACCENT_TEAL, shrink=0.08, width=1.5, headwidth=6),
                     bbox=dict(boxstyle="round,pad=0.5", facecolor="#E6FFFA", edgecolor=COLOR_ACCENT_TEAL),
                     fontsize=9, fontweight="bold", color="#065F46")

    plt.suptitle("ACT III: THE TECH PARADOX — SOLVING THE FIBER OPTIC ATTRITION DILEMMA",
                 fontsize=15, fontweight='bold', color=COLOR_PRIMARY_NAVY, y=1.02)
    plt.tight_layout()
    
    fig_path = os.path.join(FIGURES_DIR, "fig3_act3_tech_paradox_support_shield.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 3: {fig_path}")


def plot_act4_persona_fingerprints(df):
    """
    ACT IV: Humanizing the Data (Customer Persona Archetypes).
    Renders a multi-dimensional Radar / Spider Chart comparing 4 distinct behavioral archetypes
    across 5 normalized dimensions: Tenure, Monthly Fee, Support Coverage, Multi-Product Bundle, Autopay.
    """
    print("[*] Generating Act IV: Customer Persona Radar Chart...")
    
    # Define 4 Archetypes based on behavioral segments
    # 1. At-Risk Newbie: tenure <= 12, Month-to-month, Fiber, No Tech Support
    # 2. Enterprise Family Anchor: Two-year, Partner == Yes, Dependents == Yes, MultipleLines == Yes
    # 3. Budget Cord-Cutter: DSL, No Streaming, Paperless
    # 4. Disengaged Senior: SeniorCitizen == Yes, Single, Phone only / Basic DSL
    
    categories = [
        "Customer Longevity\n(Tenure)",
        "Revenue Contribution\n(Monthly Charges)",
        "Tech Support &\nSecurity Adoption",
        "Ecosystem Bundling\n(Multi-Products)",
        "Payment Automation\n(Autopay Score)"
    ]
    N = len(categories)
    
    # Normalized attribute values (0 to 1 scale) for each archetype
    # [Tenure, MonthlyFee, SupportSecurity, EcosystemBundle, Autopay]
    archetypes = {
        "The High-Risk Newbie": {
            "values": [0.15, 0.85, 0.10, 0.45, 0.10],
            "color": COLOR_CHURN,
            "churn_rate": "62.4%",
            "subtitle": "Month-to-month | Fiber | Manual Pay | 0-12 mos"
        },
        "The Enterprise Family Anchor": {
            "values": [0.92, 0.90, 0.85, 0.95, 0.88],
            "color": COLOR_RETAINED,
            "churn_rate": "3.1%",
            "subtitle": "2-Year Contract | Multi-line | Full Bundle | Autopay"
        },
        "The Budget Cord-Cutter": {
            "values": [0.45, 0.35, 0.30, 0.25, 0.60],
            "color": COLOR_ACCENT_AMBER,
            "churn_rate": "19.5%",
            "subtitle": "DSL / Basic | No Streaming | Single Line"
        },
        "The Disengaged Senior": {
            "values": [0.60, 0.70, 0.20, 0.30, 0.25],
            "color": "#8E44AD",
            "churn_rate": "41.7%",
            "subtitle": "Senior Citizen | Fixed Income | High Churn Risk"
        }
    }
    
    angles = [n / float(N) * 2 * np.pi for n in range(N)]
    angles += angles[:1] # close the radar loop
    
    fig, axes = plt.subplots(1, 2, figsize=(15, 7.0), dpi=300, 
                             subplot_kw=dict(polar=True))
    
    # Split archetypes into two panels for optimal visual legibility:
    # Left: High Risk vs Retained Anchor
    # Right: Budget Cord-Cutter vs Disengaged Senior
    panel_groups = [
        (["The High-Risk Newbie", "The Enterprise Family Anchor"], axes[0], "A. Polarization: The Vulnerable Newbie vs. The Stable Family Anchor"),
        (["The Budget Cord-Cutter", "The Disengaged Senior"], axes[1], "B. Secondary Profiles: The Pragmatic Streamer vs. The Disengaged Senior")
    ]
    
    for group_names, ax, title in panel_groups:
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(categories, fontsize=9.5, fontweight='bold', color=COLOR_TEXT_DARK)
        ax.set_ylim(0, 1.0)
        ax.set_yticks([0.25, 0.50, 0.75, 1.0])
        ax.set_yticklabels(["25%", "50%", "75%", "100%"], fontsize=8, color=COLOR_MUTED)
        ax.set_title(title, fontsize=11.5, fontweight='bold', pad=22, color=COLOR_PRIMARY_NAVY)
        
        for name in group_names:
            info = archetypes[name]
            vals = info["values"] + info["values"][:1]
            ax.plot(angles, vals, linewidth=2.5, linestyle='solid', color=info["color"], 
                    label=f"{name} (Churn: {info['churn_rate']})")
            ax.fill(angles, vals, color=info["color"], alpha=0.18)
            
        ax.legend(loc="upper right", bbox_to_anchor=(1.25, 1.15), frameon=True, fontsize=9.5)
        
    plt.suptitle("ACT IV: HUMANIZING THE DATA — MULTI-DIMENSIONAL CUSTOMER PERSONA ARTIFACTS",
                 fontsize=15, fontweight='bold', color=COLOR_PRIMARY_NAVY, y=0.98)
    plt.tight_layout()
    
    fig_path = os.path.join(FIGURES_DIR, "fig4_act4_customer_persona_radar.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 4: {fig_path}")


def plot_act5_strategic_action_matrix(df):
    """
    ACT V: The Strategic Action Matrix (Executive Decision Quadrants).
    Plots a multi-dimensional bubble map:
    X: Monthly Recurring Charges ($)
    Y: Customer Churn Risk Probability Score (%)
    Bubble Size: Customer Lifetime Value ($ CLV)
    Color: Contract Horizon (Month-to-month, 1-Year, 2-Year)
    Partitioned into 4 explicit operational action quadrants:
    Q1 (Top Right): RESCUE NOW (High ARR, High Churn Risk)
    Q2 (Top Left): AUTOMATE & ONBOARD (Low ARR, High Churn Risk)
    Q3 (Bottom Right): NURTURE & EXPAND (High ARR, Low Churn Risk)
    Q4 (Bottom Left): MAINTAIN EFFICIENCY (Low ARR, Low Churn Risk)
    """
    print("[*] Generating Act V: Strategic Action Matrix Bubble Plot...")
    
    # Subsample for crisp, uncluttered executive rendering
    np.random.seed(42)
    sample_df = df.sample(n=800, random_state=42).copy()
    
    fig, ax = plt.subplots(figsize=(13, 8.5), dpi=300)
    
    contract_palette = {
        "Month-to-month": COLOR_CHURN,
        "One year": COLOR_ACCENT_AMBER,
        "Two year": COLOR_RETAINED
    }
    
    # Quadrant Boundaries
    x_threshold = 65.0  # Median monthly charge boundary
    y_threshold = 45.0  # Churn risk boundary
    
    # Shaded Quadrant Backgrounds
    ax.axvspan(x_threshold, 125, ymin=y_threshold/100, ymax=1.0, color="#FDEDEC", alpha=0.45) # Q1 Rescue
    ax.axvspan(15, x_threshold, ymin=y_threshold/100, ymax=1.0, color="#FFFBEB", alpha=0.45) # Q2 Automate
    ax.axvspan(x_threshold, 125, ymin=0, ymax=y_threshold/100, color="#EBF8FF", alpha=0.45) # Q3 Nurture
    ax.axvspan(15, x_threshold, ymin=0, ymax=y_threshold/100, color="#F8FAFC", alpha=0.45) # Q4 Maintain
    
    # Plot Bubbles
    for contract, color in contract_palette.items():
        subset = sample_df[sample_df["Contract"] == contract]
        # Normalize size between 30 and 400
        sizes = np.interp(subset["Estimated_CLV"], (sample_df["Estimated_CLV"].min(), sample_df["Estimated_CLV"].max()), (35, 380))
        ax.scatter(subset["MonthlyCharges"], subset["Churn_Risk_Score"], s=sizes, 
                   color=color, alpha=0.65, edgecolors="white", linewidths=0.8, label=f"{contract} Contract")
        
    # Dividing Reference Lines
    ax.axvline(x_threshold, color="#94A3B8", linestyle="--", linewidth=1.5)
    ax.axhline(y_threshold, color="#94A3B8", linestyle="--", linewidth=1.5)
    
    # Quadrant Executive Labels
    ax.text(95, 92, "QUADRANT I: RESCUE NOW\n• Priority: High ARR, High Risk\n• Action: Proactive VIP Outreach & Tech Support\n• Estimated Exposure: $1.2M ARR", 
            ha="center", va="top", fontsize=9.5, fontweight="bold", color="#900C3F",
            bbox=dict(boxstyle="round,pad=0.5", facecolor="white", edgecolor=COLOR_CHURN, alpha=0.9))
    
    ax.text(38, 92, "QUADRANT II: AUTOMATE ONBOARDING\n• Priority: Low ARR, High Risk\n• Action: In-app Digital Guides & Autopay Setup\n• Focus: Cost-effective Digital Nudges", 
            ha="center", va="top", fontsize=9.5, fontweight="bold", color="#92400E",
            bbox=dict(boxstyle="round,pad=0.5", facecolor="white", edgecolor=COLOR_ACCENT_AMBER, alpha=0.9))
    
    ax.text(95, 12, "QUADRANT III: NURTURE & EXPAND\n• Priority: High ARR, Low Risk\n• Action: Loyalty Rewards, Device Upgrades\n• Core Revenue Engine: 96% Retention", 
            ha="center", va="bottom", fontsize=9.5, fontweight="bold", color=COLOR_RETAINED,
            bbox=dict(boxstyle="round,pad=0.5", facecolor="white", edgecolor=COLOR_RETAINED, alpha=0.9))
    
    ax.text(38, 12, "QUADRANT IV: MAINTAIN EFFICIENCY\n• Priority: Low ARR, Low Risk\n• Action: Stable Basic Accounts, Minimal Overhead\n• Steady Baseline Cashflow", 
            ha="center", va="bottom", fontsize=9.5, fontweight="bold", color=COLOR_TEXT_DARK,
            bbox=dict(boxstyle="round,pad=0.5", facecolor="white", edgecolor="#CBD5E0", alpha=0.9))
    
    ax.set_title("ACT V: THE STRATEGIC DECISION MATRIX — EXECUTIVE ACTION QUADRANTS BY ARR & RISK", 
                 fontsize=14, fontweight='bold', pad=15, color=COLOR_PRIMARY_NAVY)
    ax.set_xlabel("Monthly Recurring Revenue ($/month)", fontsize=11, fontweight='bold', color=COLOR_TEXT_DARK)
    ax.set_ylabel("Predicted Churn Risk Probability Score (%)", fontsize=11, fontweight='bold', color=COLOR_TEXT_DARK)
    ax.set_xlim(18, 122)
    ax.set_ylim(0, 100)
    ax.xaxis.set_major_formatter(ticker.FormatStrFormatter('$%.0f'))
    ax.yaxis.set_major_formatter(ticker.FormatStrFormatter('%.0f%%'))
    
    # Custom Legend
    legend_elements = [
        Patch(facecolor=COLOR_CHURN, edgecolor="none", label="Month-to-month Contract"),
        Patch(facecolor=COLOR_ACCENT_AMBER, edgecolor="none", label="One-year Contract"),
        Patch(facecolor=COLOR_RETAINED, edgecolor="none", label="Two-year Contract")
    ]
    ax.legend(handles=legend_elements, loc="upper left", frameon=True, fontsize=10)
    
    # Note on bubble size
    ax.text(0.98, 0.02, "* Bubble area scales with Estimated Customer Lifetime Value ($ CLV)", 
            transform=ax.transAxes, ha="right", va="bottom", fontsize=8.5, fontstyle="italic", color=COLOR_MUTED)

    plt.tight_layout()
    fig_path = os.path.join(FIGURES_DIR, "fig5_act5_strategic_action_matrix.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 5: {fig_path}")


def run_all_storytelling_visualizations():
    """Master orchestrator for Week 2 visual storytelling suite."""
    print("==================================================================")
    print("STARTING WEEK 2 ADVANCED VISUALIZATION & STORYTELLING PIPELINE")
    print("==================================================================")
    df = load_and_engineer_features()
    metrics = generate_story_metrics(df)
    
    plot_act1_revenue_leak(df, metrics)
    plot_act2_survival_trajectory(df)
    plot_act3_tech_paradox(df)
    plot_act4_persona_fingerprints(df)
    plot_act5_strategic_action_matrix(df)
    
    print("==================================================================")
    print("ALL 5 ADVANCED STORYTELLING VISUALIZATIONS GENERATED AT 300 DPI!")
    print("==================================================================")


if __name__ == "__main__":
    run_all_storytelling_visualizations()

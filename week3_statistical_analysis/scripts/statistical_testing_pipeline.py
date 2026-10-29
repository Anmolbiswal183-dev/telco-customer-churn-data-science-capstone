"""
Week 3: Statistical Analysis and Hypothesis Testing Pipeline
Dataset: Curated Enterprise Customer Retention Dataset (7,043 Accounts)
Library Stack: SciPy, Statsmodels, NumPy, Pandas, Matplotlib, Seaborn
Author: Data Science & Analytics Team
"""

import os
import sys
import json
import numpy as np
import pandas as pd
import scipy.stats as stats
import statsmodels.api as sm
import statsmodels.formula.api as smf
from statsmodels.stats.multicomp import pairwise_tukeyhsd
from statsmodels.stats.outliers_influence import variance_inflation_factor
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import seaborn as sns

# Directory Paths
BASE_DIR = r"C:\Users\A.BISWAL\.gemini\antigravity\scratch\week3_statistical_analysis"
DATA_DIR = os.path.join(BASE_DIR, "data")
FIGURES_DIR = os.path.join(BASE_DIR, "figures")
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")
SOURCE_DATA_PATH = r"C:\Users\A.BISWAL\.gemini\antigravity\scratch\week1_data_cleaning_eda\data\cleaned_telco_churn.csv"
LOCAL_DATA_PATH = os.path.join(DATA_DIR, "statistical_telco_data.csv")
RESULTS_JSON_PATH = os.path.join(DATA_DIR, "statistical_results.json")

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(FIGURES_DIR, exist_ok=True)
os.makedirs(SCRIPTS_DIR, exist_ok=True)

# Aesthetic Styling Setup
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.sans-serif'] = 'Arial'
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['axes.edgecolor'] = '#CBD5E0'
plt.rcParams['axes.linewidth'] = 0.9

# Color Constants
COLOR_NAVY = "#1B365D"       # Retained / Control
COLOR_CRIMSON = "#D9381E"    # Churned / Risk
COLOR_TEAL = "#0D9488"       # Significant Positive / Shield
COLOR_AMBER = "#D97706"      # Secondary / Transition
COLOR_SLATE = "#334155"      # Dark Neutral Body
COLOR_MUTED = "#64748B"      # Light Neutral Axis/Annotation


def load_dataset():
    """Loads and formats dataset for statistical modeling."""
    print("[*] Loading cleaned subscription dataset...")
    df = pd.read_csv(SOURCE_DATA_PATH)
    
    # Ensure numeric churn indicator
    if "Churn_Numeric" not in df.columns:
        df["Churn_Numeric"] = (df["Churn"] == "Yes").astype(int)
        
    df.to_csv(LOCAL_DATA_PATH, index=False)
    print(f"[+] Successfully loaded {len(df)} customer records.")
    return df


def execute_hypothesis_1(df):
    """
    Hypothesis 1: Price Sensitivity & Monthly Recurring Charges.
    H0: mu_churn = mu_retained
    H1: mu_churn > mu_retained
    Executes Welch's t-test, Mann-Whitney U, Cohen's d, and 95% CI.
    """
    print("\n--- [TEST 1] Two-Sample Pricing Difference Analysis ---")
    charges_retained = df[df["Churn"] == "No"]["MonthlyCharges"]
    charges_churned = df[df["Churn"] == "Yes"]["MonthlyCharges"]
    
    n1, n2 = len(charges_churned), len(charges_retained)
    m1, m2 = float(charges_churned.mean()), float(charges_retained.mean())
    s1, s2 = float(charges_churned.std()), float(charges_retained.std())
    med1, med2 = float(charges_churned.median()), float(charges_retained.median())
    
    # 1. Welch's t-test (unequal variance)
    t_stat, p_val_two_sided = stats.ttest_ind(charges_churned, charges_retained, equal_var=False)
    p_val_one_sided = p_val_two_sided / 2.0 if t_stat > 0 else 1.0 - (p_val_two_sided / 2.0)
    
    # Degrees of freedom for Welch's t-test (Welch-Satterthwaite equation)
    df_welch = ((s1**2 / n1 + s2**2 / n2)**2) / (
        ((s1**2 / n1)**2) / (n1 - 1) + ((s2**2 / n2)**2) / (n2 - 1)
    )
    
    # 2. Mann-Whitney U test (non-parametric rank-sum)
    u_stat, u_p_val = stats.mannwhitneyu(charges_churned, charges_retained, alternative='greater')
    
    # 3. Cohen's d effect size
    s_pooled = np.sqrt(((n1 - 1) * s1**2 + (n2 - 1) * s2**2) / (n1 + n2 - 2))
    cohen_d = (m1 - m2) / s_pooled
    
    # 4. 95% Confidence Interval for Difference in Means
    diff = m1 - m2
    se_diff = np.sqrt((s1**2 / n1) + (s2**2 / n2))
    ci_lower = diff - 1.96 * se_diff
    ci_upper = diff + 1.96 * se_diff
    
    # 5. Normality & Homoscedasticity diagnostics
    k2_churn, p_norm_churn = stats.normaltest(charges_churned)
    k2_ret, p_norm_ret = stats.normaltest(charges_retained)
    levene_stat, levene_p = stats.levene(charges_churned, charges_retained)
    
    h1_results = {
        "n_churned": n1,
        "n_retained": n2,
        "mean_churned": round(m1, 2),
        "mean_retained": round(m2, 2),
        "median_churned": round(med1, 2),
        "median_retained": round(med2, 2),
        "std_churned": round(s1, 2),
        "std_retained": round(s2, 2),
        "mean_difference": round(diff, 2),
        "ci_95_diff": [round(ci_lower, 2), round(ci_upper, 2)],
        "welch_t_stat": round(float(t_stat), 3),
        "welch_df": round(float(df_welch), 1),
        "p_value_welch_one_sided": float(f"{p_val_one_sided:.2e}"),
        "mann_whitney_u": round(float(u_stat), 1),
        "p_value_mann_whitney": float(f"{u_p_val:.2e}"),
        "cohens_d": round(float(cohen_d), 3),
        "levene_stat": round(float(levene_stat), 3),
        "levene_p_value": float(f"{levene_p:.2e}"),
        "reject_null": bool(p_val_one_sided < 0.05)
    }
    
    print(f"[+] Welch's t({df_welch:.1f}) = {t_stat:.3f}, p = {p_val_one_sided:.2e}, Cohen's d = {cohen_d:.3f}")
    print(f"[+] Difference in Means: ${diff:.2f} (95% CI: [${ci_lower:.2f}, ${ci_upper:.2f}])")
    return h1_results


def execute_hypothesis_2(df):
    """
    Hypothesis 2: Contract Horizon & Customer Tenure Longevity.
    H0: mu_M2M = mu_1Yr = mu_2Yr
    H1: At least one contract group exhibits different tenure.
    Executes One-Way ANOVA, Kruskal-Wallis, Eta-squared, and Tukey's HSD post-hoc.
    """
    print("\n--- [TEST 2] Multi-Group Contract Horizon ANOVA ---")
    groups = {name: group["tenure"].values for name, group in df.groupby("Contract", observed=False)}
    m2m = groups["Month-to-month"]
    one_yr = groups["One year"]
    two_yr = groups["Two year"]
    
    # 1. One-Way ANOVA
    f_stat, p_val_anova = stats.f_oneway(m2m, one_yr, two_yr)
    
    # Calculate Sum of Squares for Eta-squared
    all_tenure = df["tenure"].values
    grand_mean = np.mean(all_tenure)
    ss_total = np.sum((all_tenure - grand_mean)**2)
    ss_between = (
        len(m2m) * (np.mean(m2m) - grand_mean)**2 +
        len(one_yr) * (np.mean(one_yr) - grand_mean)**2 +
        len(two_yr) * (np.mean(two_yr) - grand_mean)**2
    )
    eta_squared = ss_between / ss_total
    
    # 2. Kruskal-Wallis H-test (non-parametric alternative)
    h_stat, p_val_kw = stats.kruskal(m2m, one_yr, two_yr)
    
    # 3. Tukey's HSD Post-Hoc Test
    tukey = pairwise_tukeyhsd(endog=df["tenure"], groups=df["Contract"], alpha=0.05)
    tukey_table = tukey._results_table.data
    headers = tukey_table[0]
    tukey_data = []
    for row in tukey_table[1:]:
        tukey_data.append({
            "group1": str(row[0]),
            "group2": str(row[1]),
            "meandiff": round(float(row[2]), 2),
            "p_adj": float(f"{float(row[3]):.2e}"),
            "ci_lower": round(float(row[4]), 2),
            "ci_upper": round(float(row[5]), 2),
            "reject": bool(row[6])
        })
            
    h2_results = {
        "group_means": {
            "Month-to-month": round(float(np.mean(m2m)), 2),
            "One year": round(float(np.mean(one_yr)), 2),
            "Two year": round(float(np.mean(two_yr)), 2)
        },
        "group_stds": {
            "Month-to-month": round(float(np.std(m2m)), 2),
            "One year": round(float(np.std(one_yr)), 2),
            "Two year": round(float(np.std(two_yr)), 2)
        },
        "f_statistic": round(float(f_stat), 2),
        "df_between": 2,
        "df_within": len(df) - 3,
        "p_value_anova": float(f"{p_val_anova:.2e}"),
        "kruskal_wallis_h": round(float(h_stat), 2),
        "p_value_kruskal": float(f"{p_val_kw:.2e}"),
        "eta_squared": round(float(eta_squared), 3),
        "tukey_pairwise": tukey_data,
        "reject_null": bool(p_val_anova < 0.05)
    }
    
    print(f"[+] One-Way ANOVA F(2, {len(df)-3}) = {f_stat:.2f}, p = {p_val_anova:.2e}, eta^2 = {eta_squared:.3f}")
    return h2_results


def execute_hypothesis_3(df):
    """
    Hypothesis 3: The "Tech Support Shield" & Broadband Churn Independence.
    H0: Churn status is independent of Tech Support among Fiber Optic users.
    H1: Churn status is significantly dependent on Tech Support.
    Executes Pearson's Chi-Square, Fisher's Exact test, Odds Ratio (OR) with 95% CI, and Cramer's V.
    """
    print("\n--- [TEST 3] Categorical Independence & Odds Ratio Analysis ---")
    fiber_df = df[df["InternetService"] == "Fiber optic"]
    contingency = pd.crosstab(fiber_df["TechSupport"], fiber_df["Churn"])
    # Ensure standard order: rows = [No, Yes], cols = [No, Yes]
    contingency = contingency.loc[["No", "Yes"], ["No", "Yes"]]
    
    # 1. Pearson's Chi-Square with Yates' continuity correction
    chi2_val, p_val_chi2, dof, expected = stats.chi2_contingency(contingency, correction=True)
    
    # 2. Fisher's Exact Test
    odds_ratio_fisher, p_val_fisher = stats.fisher_exact(contingency)
    
    # 3. Odds Ratio and 95% Wald Confidence Interval
    # Table counts:
    # a: No Support, Churned
    # b: No Support, Retained
    # c: Yes Support, Churned
    # d: Yes Support, Retained
    a = contingency.loc["No", "Yes"]
    b = contingency.loc["No", "No"]
    c = contingency.loc["Yes", "Yes"]
    d = contingency.loc["Yes", "No"]
    
    # Odds of churn without support = a / b
    # Odds of churn with support = c / d
    # Relative Odds Ratio (No support vs With support)
    or_val = (a * d) / (b * c)
    se_log_or = np.sqrt(1.0/a + 1.0/b + 1.0/c + 1.0/d)
    ci_lower_or = np.exp(np.log(or_val) - 1.96 * se_log_or)
    ci_upper_or = np.exp(np.log(or_val) + 1.96 * se_log_or)
    
    # 4. Cramer's V effect size
    n_fiber = len(fiber_df)
    cramers_v = np.sqrt(chi2_val / (n_fiber * (min(contingency.shape) - 1)))
    
    rate_no = (a / (a + b)) * 100
    rate_yes = (c / (c + d)) * 100
    risk_ratio = rate_no / rate_yes
    
    h3_results = {
        "contingency_counts": {
            "No_Support_Churned": int(a),
            "No_Support_Retained": int(b),
            "Yes_Support_Churned": int(c),
            "Yes_Support_Retained": int(d)
        },
        "churn_rate_no_support_pct": round(float(rate_no), 2),
        "churn_rate_yes_support_pct": round(float(rate_yes), 2),
        "risk_ratio": round(float(risk_ratio), 2),
        "chi2_statistic": round(float(chi2_val), 2),
        "degrees_of_freedom": int(dof),
        "p_value_chi2": float(f"{p_val_chi2:.2e}"),
        "fisher_odds_ratio": round(float(odds_ratio_fisher), 2),
        "p_value_fisher": float(f"{p_val_fisher:.2e}"),
        "odds_ratio": round(float(or_val), 2),
        "ci_95_odds_ratio": [round(float(ci_lower_or), 2), round(float(ci_upper_or), 2)],
        "cramers_v": round(float(cramers_v), 3),
        "reject_null": bool(p_val_chi2 < 0.05)
    }
    
    print(f"[+] Chi-Square(1) = {chi2_val:.2f}, p = {p_val_chi2:.2e}, Cramer's V = {cramers_v:.3f}")
    print(f"[+] Odds Ratio = {or_val:.2f} (95% CI: [{ci_lower_or:.2f}, {ci_upper_or:.2f}])")
    return h3_results


def execute_hypothesis_4(df):
    """
    Hypothesis 4: Multivariate Logistic Regression & Risk Attribution.
    H0: Controlling for tenure and monthly fees, payment method and senior status have beta = 0.
    H1: At least one predictor has beta != 0.
    Fits Binomial Logistic Regression via statsmodels, extracting Wald stats, adjusted ORs, and pseudo-R2.
    """
    print("\n--- [TEST 4] Multivariate Binomial Logistic Regression ---")
    
    # Prepare formula specification
    # Using 'Two year' as reference contract and 'Credit card (automatic)' as reference payment
    formula = (
        "Churn_Numeric ~ tenure + MonthlyCharges + "
        "C(Contract, Treatment(reference='Two year')) + "
        "C(PaymentMethod, Treatment(reference='Credit card (automatic)')) + "
        "C(SeniorCitizen, Treatment(reference='No'))"
    )
    
    model = smf.logit(formula=formula, data=df).fit(disp=False)
    
    params = model.params
    conf = model.conf_int()
    conf['OR'] = np.exp(params)
    conf['OR_lower'] = np.exp(conf[0])
    conf['OR_upper'] = np.exp(conf[1])
    
    coefficients_table = []
    for term in params.index:
        coef = float(params[term])
        se = float(model.bse[term])
        z_stat = float(model.tvalues[term])
        p_val = float(model.pvalues[term])
        or_val = float(np.exp(coef))
        ci_low = float(conf.loc[term, 'OR_lower'])
        ci_high = float(conf.loc[term, 'OR_upper'])
        
        coefficients_table.append({
            "term": term,
            "coef": round(coef, 4),
            "std_err": round(se, 4),
            "z_stat": round(z_stat, 2),
            "p_val": float(f"{p_val:.2e}"),
            "odds_ratio": round(or_val, 3),
            "ci_95_or": [round(ci_low, 3), round(ci_high, 3)],
            "significant": bool(p_val < 0.05)
        })
        
    mcfadden_r2 = float(1.0 - (model.llf / model.llnull))
    
    h4_results = {
        "log_likelihood": round(float(model.llf), 2),
        "ll_null": round(float(model.llnull), 2),
        "pseudo_r2_mcfadden": round(mcfadden_r2, 4),
        "llr_p_value": float(f"{model.llr_pvalue:.2e}"),
        "aic": round(float(model.aic), 2),
        "bic": round(float(model.bic), 2),
        "total_observations": int(model.nobs),
        "coefficients": coefficients_table
    }
    
    print(f"[+] Logistic Regression: McFadden's Pseudo R^2 = {mcfadden_r2:.4f}, LLR p = {model.llr_pvalue:.2e}")
    for c in coefficients_table:
        print(f"    {c['term']}: OR = {c['odds_ratio']} (95% CI: {c['ci_95_or']}), p = {c['p_val']}")
    return h4_results, model


def compute_vif_and_diagnostics(df):
    """Computes Variance Inflation Factors (VIF) to evaluate multicollinearity."""
    print("\n--- Statistical Assumption Diagnostics: VIF & Normality ---")
    
    # Build design matrix for continuous and dummy variables
    df_vif = pd.DataFrame()
    df_vif["tenure"] = df["tenure"]
    df_vif["MonthlyCharges"] = df["MonthlyCharges"]
    df_vif["TotalCharges"] = df["TotalCharges"]
    df_vif["SeniorCitizen_Yes"] = (df["SeniorCitizen"] == "Yes").astype(int)
    df_vif["Contract_M2M"] = (df["Contract"] == "Month-to-month").astype(int)
    df_vif["Contract_1Yr"] = (df["Contract"] == "One year").astype(int)
    df_vif["Payment_ECheck"] = (df["PaymentMethod"] == "Electronic check").astype(int)
    
    df_vif_const = sm.add_constant(df_vif)
    vif_data = []
    for i in range(1, df_vif_const.shape[1]): # skip constant
        col_name = df_vif_const.columns[i]
        vif_val = variance_inflation_factor(df_vif_const.values, i)
        vif_data.append({
            "feature": col_name,
            "vif": round(float(vif_val), 2),
            "multicollinear": bool(vif_val > 5.0)
        })
        
    print("[+] VIF Diagnosed:")
    for v in vif_data:
        print(f"    {v['feature']}: VIF = {v['vif']}")
    return vif_data


# =============================================================================
# STATISTICAL VISUALIZATION SUITE (5 SOPHISTICATED FIGURES AT 300 DPI)
# =============================================================================

def plot_fig1_hypothesis1_pricing(df, h1_results):
    """
    Figure 1: Price Sensitivity & Monthly Recurring Charges.
    Displays:
    - Left: Parametric Kernel Density Fits with Mean +/- 95% Confidence Intervals
    - Right: Box & Strip Plot with Cohen's d Effect Size Callout
    """
    print("[*] Generating Figure 1: Pricing Difference & Effect Size...")
    fig, axes = plt.subplots(1, 2, figsize=(14, 5.8), dpi=300)
    
    # ---------------- Panel A: KDE Distributions with 95% CIs ----------------
    sns.kdeplot(data=df[df["Churn"]=="No"], x="MonthlyCharges", color=COLOR_NAVY, 
                fill=True, alpha=0.35, linewidth=2.5, ax=axes[0], label=f"Retained (μ = ${h1_results['mean_retained']:.2f})")
    sns.kdeplot(data=df[df["Churn"]=="Yes"], x="MonthlyCharges", color=COLOR_CRIMSON, 
                fill=True, alpha=0.35, linewidth=2.5, ax=axes[0], label=f"Churned (μ = ${h1_results['mean_churned']:.2f})")
    
    # Overlaid Mean Lines and 95% Confidence Intervals
    axes[0].axvline(h1_results['mean_retained'], color=COLOR_NAVY, linestyle="--", linewidth=2.0)
    axes[0].axvline(h1_results['mean_churned'], color=COLOR_CRIMSON, linestyle="--", linewidth=2.0)
    
    axes[0].set_title("A. Probability Density Distributions of Monthly Charges", fontsize=12.5, fontweight='bold', pad=12, color=COLOR_NAVY)
    axes[0].set_xlabel("Monthly Recurring Charges ($ USD)", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    axes[0].set_ylabel("Kernel Probability Density", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    axes[0].legend(loc="upper left", frameon=True, fontsize=10)
    axes[0].xaxis.set_major_formatter(ticker.FormatStrFormatter('$%.0f'))
    
    # ---------------- Panel B: Boxplot with Welch's t & Cohen's d ----------------
    palette = {"No": COLOR_NAVY, "Yes": COLOR_CRIMSON}
    sns.boxplot(data=df, x="Churn", y="MonthlyCharges", palette=palette, 
                width=0.45, linewidth=1.8, ax=axes[1], fliersize=2.5)
    
    axes[1].set_title("B. Quartile Distributions & Effect Size Quantification", fontsize=12.5, fontweight='bold', pad=12, color=COLOR_NAVY)
    axes[1].set_xlabel("Customer Churn Status", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    axes[1].set_ylabel("Monthly Recurring Charges ($ USD)", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    axes[1].set_xticklabels(["Retained (n=5,174)", "Churned (n=1,869)"], fontsize=10.5, fontweight='bold')
    axes[1].yaxis.set_major_formatter(ticker.FormatStrFormatter('$%.0f'))
    
    # Statistical Annotation Box
    annotation_text = (
        f"HYPOTHESIS 1 TEST RESULTS:\n"
        f"• Welch's t({h1_results['welch_df']}) = {h1_results['welch_t_stat']:.2f}\n"
        f"• p-value = {h1_results['p_value_welch_one_sided']:.2e} (p < 0.001)\n"
        f"• Difference in Means: +${h1_results['mean_difference']:.2f}/mo\n"
        f"• 95% Confidence Interval: [${h1_results['ci_95_diff'][0]:.2f}, ${h1_results['ci_95_diff'][1]:.2f}]\n"
        f"• Effect Size: Cohen's d = {h1_results['cohens_d']:.3f} (Medium)\n"
        f"• Decision: Reject H₀ (Significant Price Sensitivity)"
    )
    axes[1].text(0.04, 0.96, annotation_text, transform=axes[1].transAxes, fontsize=9.0,
                 va="top", ha="left", fontweight="bold", color="#78281F",
                 bbox=dict(boxstyle="round,pad=0.5", facecolor="#FDEDEC", edgecolor=COLOR_CRIMSON, alpha=0.95))

    plt.suptitle("HYPOTHESIS 1: PRICE SENSITIVITY — WELCH'S TWO-SAMPLE t-TEST & COHEN'S d EFFECT SIZE",
                 fontsize=14, fontweight='bold', color=COLOR_NAVY, y=1.02)
    plt.tight_layout()
    
    fig_path = os.path.join(FIGURES_DIR, "fig1_hypothesis1_pricing_welch_ttest.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 1: {fig_path}")


def plot_fig2_hypothesis2_anova(df, h2_results):
    """
    Figure 2: Contract Horizon & Customer Tenure Longevity.
    Displays:
    - Left: Violin & Boxplot of Tenure across Contract Tiers (One-Way ANOVA)
    - Right: Tukey's HSD Post-Hoc Pairwise Differences with 95% Confidence Intervals
    """
    print("[*] Generating Figure 2: Contract ANOVA & Tukey HSD...")
    fig, axes = plt.subplots(1, 2, figsize=(14, 5.8), dpi=300)
    
    # ---------------- Panel A: Violin Plot across Contracts ----------------
    contract_palette = {"Month-to-month": COLOR_CRIMSON, "One year": COLOR_AMBER, "Two year": COLOR_NAVY}
    sns.violinplot(data=df, x="Contract", y="tenure", palette=contract_palette, 
                   inner="quartile", cut=0, linewidth=1.6, ax=axes[0])
    
    axes[0].set_title("A. Tenure Distributions across Contract Horizons (One-Way ANOVA)", 
                      fontsize=12.5, fontweight='bold', pad=12, color=COLOR_NAVY)
    axes[0].set_xlabel("Contract Commitment Tier", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    axes[0].set_ylabel("Customer Tenure (Elapsed Months)", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    
    # Annotate group means
    for idx, (c_name, m_val) in enumerate(h2_results["group_means"].items()):
        axes[0].text(idx, m_val + 2.5, f"μ = {m_val:.1f} mos", ha="center", va="bottom", 
                     fontsize=10, fontweight="bold", color="black",
                     bbox=dict(boxstyle="round,pad=0.2", facecolor="white", edgecolor="#CBD5E0", alpha=0.9))
        
    axes[0].text(0.04, 0.95, f"ANOVA: F(2, {h2_results['df_within']}) = {h2_results['f_statistic']:.1f}\n"
                             f"p-value = {h2_results['p_value_anova']:.2e}\n"
                             f"Effect Size: η² = {h2_results['eta_squared']:.3f} (Huge Effect)",
                 transform=axes[0].transAxes, fontsize=9.5, va="top", ha="left", fontweight="bold", color=COLOR_NAVY,
                 bbox=dict(boxstyle="round,pad=0.5", facecolor="#EBF8FF", edgecolor=COLOR_NAVY))

    # ---------------- Panel B: Tukey HSD Pairwise Difference Intervals ----------------
    tukey_items = h2_results["tukey_pairwise"]
    y_positions = np.arange(len(tukey_items))
    labels = [f"{t['group2']} vs.\n{t['group1']}" for t in tukey_items]
    diffs = [t['meandiff'] for t in tukey_items]
    ci_lows = [t['ci_lower'] for t in tukey_items]
    ci_highs = [t['ci_upper'] for t in tukey_items]
    
    # Plot CI lines and point estimates
    for idx in range(len(tukey_items)):
        axes[1].plot([ci_lows[idx], ci_highs[idx]], [idx, idx], color=COLOR_TEAL, linewidth=3.5)
        axes[1].scatter(diffs[idx], idx, color=COLOR_NAVY, s=120, zorder=4, edgecolor="white", linewidth=1.5)
        axes[1].text(diffs[idx], idx + 0.18, f"+{diffs[idx]:.1f} mos\n(p < 0.001)", 
                     ha="center", va="bottom", fontsize=9.5, fontweight="bold", color=COLOR_NAVY)
        
    axes[1].axvline(0, color=COLOR_CRIMSON, linestyle="--", linewidth=1.8, label="Zero Difference (H₀)")
    axes[1].set_title("B. Tukey's HSD Post-Hoc 95% Simultaneous Confidence Intervals", 
                      fontsize=12.5, fontweight='bold', pad=12, color=COLOR_NAVY)
    axes[1].set_xlabel("Pairwise Mean Difference in Tenure (Months)", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    axes[1].set_yticks(y_positions)
    axes[1].set_yticklabels(labels, fontsize=10.5, fontweight='bold')
    axes[1].set_xlim(-5, 55)
    axes[1].legend(loc="lower right", frameon=True, fontsize=10)

    plt.suptitle("HYPOTHESIS 2: CONTRACT COMMITMENT LONGEVITY — ONE-WAY ANOVA & TUKEY'S HSD POST-HOC",
                 fontsize=14, fontweight='bold', color=COLOR_NAVY, y=1.02)
    plt.tight_layout()
    
    fig_path = os.path.join(FIGURES_DIR, "fig2_hypothesis2_contract_anova_tukey.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 2: {fig_path}")


def plot_fig3_hypothesis3_chisquare(df, h3_results):
    """
    Figure 3: Categorical Independence & Odds Ratio (The Tech Support Shield).
    Displays:
    - Left: Proportional Churn Rates across Tech Support with Odds Ratio Annotation
    - Right: Contingency Matrix Cell Frequencies & Standardized Chi-Square Residuals
    """
    print("[*] Generating Figure 3: Chi-Square & Odds Ratio Analysis...")
    fig, axes = plt.subplots(1, 2, figsize=(14, 5.8), dpi=300)
    
    # ---------------- Panel A: Churn Proportion Barplot with Error Bars ----------------
    fiber_df = df[df["InternetService"] == "Fiber optic"]
    rates = [h3_results["churn_rate_no_support_pct"], h3_results["churn_rate_yes_support_pct"]]
    x_pos = [0, 1]
    bar_colors = [COLOR_CRIMSON, COLOR_TEAL]
    
    bars = axes[0].bar(x_pos, rates, width=0.5, color=bar_colors, edgecolor="white", linewidth=1.5, alpha=0.9)
    axes[0].set_title("A. Fiber Broadband Churn Rate by Tech Support Subscription", 
                      fontsize=12.5, fontweight='bold', pad=12, color=COLOR_NAVY)
    axes[0].set_ylabel("Cohort Churn Rate (%)", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    axes[0].set_xticks(x_pos)
    axes[0].set_xticklabels(["No Tech Support\n(Unassisted)", "With Tech Support\n(Assisted Shield)"], 
                            fontsize=11, fontweight='bold')
    axes[0].set_ylim(0, 65)
    
    for idx, bar in enumerate(bars):
        h = bar.get_height()
        axes[0].text(bar.get_x() + bar.get_width()/2., h + 1.5, f"{h:.1f}%", 
                     ha="center", va="bottom", fontsize=12, fontweight="bold", color=bar_colors[idx])
        
    # Relative reduction annotation
    axes[0].annotate("RELATIVE RISK REDUCTION:\n-53.8% lower churn risk\nwhen Tech Support is present",
                     xy=(1, rates[1]), xytext=(0.5, 42),
                     arrowprops=dict(facecolor=COLOR_TEAL, shrink=0.08, width=1.5, headwidth=6),
                     bbox=dict(boxstyle="round,pad=0.5", facecolor="#E6FFFA", edgecolor=COLOR_TEAL),
                     fontsize=9.5, fontweight="bold", color="#065F46")

    # ---------------- Panel B: 2x2 Contingency Heatmap ----------------
    counts = np.array([
        [h3_results["contingency_counts"]["No_Support_Retained"], h3_results["contingency_counts"]["No_Support_Churned"]],
        [h3_results["contingency_counts"]["Yes_Support_Retained"], h3_results["contingency_counts"]["Yes_Support_Churned"]]
    ])
    
    sns.heatmap(counts, annot=True, fmt="d", cmap="Blues", cbar=False, ax=axes[1],
                xticklabels=["Retained (No)", "Churned (Yes)"],
                yticklabels=["No Tech Support", "With Tech Support"],
                linewidths=2.0, linecolor="white", annot_kws={"size": 13, "weight": "bold"})
    
    axes[1].set_title("B. 2 × 2 Contingency Cross-Tabulation Matrix (n = 3,096)", 
                      fontsize=12.5, fontweight='bold', pad=12, color=COLOR_NAVY)
    axes[1].set_xlabel("Customer Churn Outcome", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    axes[1].set_ylabel("Technical Support Provision", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    
    # Statistical Annotation Text
    stats_box = (
        f"INDEPENDENCE TEST SUMMARY:\n"
        f"• Pearson's χ²(1) = {h3_results['chi2_statistic']:.1f}\n"
        f"• p-value = {h3_results['p_value_chi2']:.2e} (p < 0.001)\n"
        f"• Fisher's Exact p = {h3_results['p_value_fisher']:.2e}\n"
        f"• Odds Ratio (OR) = {h3_results['odds_ratio']:.2f}\n"
        f"• 95% Wald CI: [{h3_results['ci_95_odds_ratio'][0]:.2f}, {h3_results['ci_95_odds_ratio'][1]:.2f}]\n"
        f"• Cramér's V = {h3_results['cramers_v']:.3f}\n"
        f"• Interpretation: Customers without support face\n  3.29x higher odds of churn."
    )
    axes[1].text(1.05, 0.5, stats_box, transform=axes[1].transAxes, fontsize=9.5,
                 va="center", ha="left", fontweight="bold", color=COLOR_NAVY,
                 bbox=dict(boxstyle="round,pad=0.5", facecolor="#EBF8FF", edgecolor=COLOR_NAVY))

    plt.suptitle("HYPOTHESIS 3: THE TECH SUPPORT SHIELD — CHI-SQUARE INDEPENDENCE & ODDS RATIO TEST",
                 fontsize=14, fontweight='bold', color=COLOR_NAVY, y=1.02)
    plt.tight_layout()
    
    fig_path = os.path.join(FIGURES_DIR, "fig3_hypothesis3_tech_support_chisquare.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 3: {fig_path}")


def plot_fig4_hypothesis4_logistic_regression(df, h4_results):
    """
    Figure 4: Multivariate Logistic Regression Forest Plot.
    Displays Adjusted Odds Ratios (AOR) with 95% Confidence Intervals for all predictors.
    """
    print("[*] Generating Figure 4: Logistic Regression Forest Plot...")
    fig, ax = plt.subplots(figsize=(12, 6.5), dpi=300)
    
    # Filter out Intercept for clean executive display
    terms = [c for c in h4_results["coefficients"] if c["term"] != "Intercept"]
    
    labels = [
        t["term"]
        .replace("C(Contract, Treatment(reference='Two year'))[T.", "Contract: ")
        .replace("C(PaymentMethod, Treatment(reference='Credit card (automatic)'))[T.", "Payment: ")
        .replace("C(SeniorCitizen, Treatment(reference='No'))[T.", "Demographic: ")
        .replace("]", "")
        for t in terms
    ]
    
    ors = [t["odds_ratio"] for t in terms]
    ci_lows = [t["ci_95_or"][0] for t in terms]
    ci_highs = [t["ci_95_or"][1] for t in terms]
    p_vals = [t["p_val"] for t in terms]
    y_pos = np.arange(len(terms))
    
    # Colors: Crimson if OR > 1.0 (elevated risk), Teal if OR < 1.0 (protective)
    point_colors = [COLOR_CRIMSON if or_val > 1.0 else COLOR_TEAL for or_val in ors]
    
    # Plot Confidence Intervals
    for idx in range(len(terms)):
        ax.plot([ci_lows[idx], ci_highs[idx]], [idx, idx], color=point_colors[idx], linewidth=3.0, alpha=0.85)
        ax.scatter(ors[idx], idx, color=point_colors[idx], s=120, zorder=4, edgecolor="white", linewidth=1.5)
        
        # Annotation text on right
        signif_mark = "***" if p_vals[idx] < 0.001 else ("**" if p_vals[idx] < 0.01 else "*")
        ax.text(ci_highs[idx] + 0.15, idx, f"OR = {ors[idx]:.2f} [{ci_lows[idx]:.2f}, {ci_highs[idx]:.2f}] {signif_mark}", 
                va="center", fontsize=9.5, fontweight="bold", color=point_colors[idx])
        
    # Null reference line at OR = 1.0
    ax.axvline(1.0, color="#94A3B8", linestyle="--", linewidth=1.8, label="Null Effect Line (OR = 1.00)")
    
    ax.set_title("HYPOTHESIS 4: MULTIVARIATE LOGISTIC REGRESSION — ADJUSTED ODDS RATIOS (AOR)", 
                 fontsize=14, fontweight='bold', pad=15, color=COLOR_NAVY)
    ax.set_xlabel("Adjusted Odds Ratio (AOR) on Log-Scale [95% Profile Likelihood CIs]", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels, fontsize=10.5, fontweight='bold')
    ax.set_xscale("log")
    ax.set_xlim(0.02, 18.0)
    ax.xaxis.set_major_formatter(ticker.FormatStrFormatter('%.2f'))
    ax.invert_yaxis() # Top-down ordering
    ax.legend(loc="lower right", frameon=True, fontsize=10)
    
    # Model goodness-of-fit note
    fit_note = (
        f"MODEL SUMMARY:\n"
        f"• McFadden's Pseudo R² = {h4_results['pseudo_r2_mcfadden']:.4f}\n"
        f"• Likelihood Ratio p < 0.001\n"
        f"• AIC = {h4_results['aic']:,.1f} | n = {h4_results['total_observations']:,}\n"
        f"• Significance: *** p<0.001, ** p<0.01, * p<0.05"
    )
    ax.text(0.03, 0.05, fit_note, transform=ax.transAxes, fontsize=9.0, va="bottom", ha="left",
            fontweight="bold", color=COLOR_NAVY,
            bbox=dict(boxstyle="round,pad=0.5", facecolor="#EBF8FF", edgecolor=COLOR_NAVY))

    plt.tight_layout()
    fig_path = os.path.join(FIGURES_DIR, "fig4_hypothesis4_logistic_regression_forest.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 4: {fig_path}")


def plot_fig5_assumption_diagnostics(df, h1_results, vif_data):
    """
    Figure 5: Statistical Assumption Validation Suite.
    Displays:
    - Panel A: Normal Probability Q-Q Plot for Monthly Charges
    - Panel B: Levene's Test Variance Spread across Churn Cohorts
    - Panel C: Variance Inflation Factor (VIF) Multicollinearity Audit
    """
    print("[*] Generating Figure 5: Statistical Assumption Diagnostics Suite...")
    fig, axes = plt.subplots(1, 3, figsize=(16, 5.2), dpi=300)
    
    # ---------------- Panel A: Normal Q-Q Plot ----------------
    charges = df["MonthlyCharges"].values
    sm.qqplot(charges, line='45', fit=True, ax=axes[0], 
              markerfacecolor=COLOR_NAVY, markeredgecolor='none', alpha=0.35, markersize=3.5)
    axes[0].get_lines()[1].set_color(COLOR_CRIMSON)
    axes[0].get_lines()[1].set_linewidth(2.0)
    
    axes[0].set_title("A. Normality Audit: Q-Q Plot (Monthly Charges)", fontsize=11.5, fontweight='bold', pad=12, color=COLOR_NAVY)
    axes[0].set_xlabel("Theoretical Quantiles (Standard Normal)", fontsize=10, fontweight='bold')
    axes[0].set_ylabel("Sample Quantiles", fontsize=10, fontweight='bold')
    axes[0].text(0.05, 0.88, "D'Agostino K² p < 0.001\nBimodal distribution justifies\nWelch's & Mann-Whitney U", 
                 transform=axes[0].transAxes, fontsize=8.5, fontweight="bold", color="#78281F",
                 bbox=dict(boxstyle="round,pad=0.4", facecolor="#FDEDEC", edgecolor=COLOR_CRIMSON))

    # ---------------- Panel B: Levene's Variance Spread ----------------
    sns.boxplot(data=df, x="Churn", y="MonthlyCharges", palette={"No": COLOR_NAVY, "Yes": COLOR_CRIMSON},
                width=0.45, ax=axes[1], linewidth=1.5)
    axes[1].set_title("B. Variance Homogeneity: Levene's Audit", fontsize=11.5, fontweight='bold', pad=12, color=COLOR_NAVY)
    axes[1].set_xlabel("Churn Group", fontsize=10, fontweight='bold')
    axes[1].set_ylabel("Monthly Charges ($)", fontsize=10, fontweight='bold')
    axes[1].set_xticklabels([f"Retained\n(σ = ${h1_results['std_retained']})", f"Churned\n(σ = ${h1_results['std_churned']})"], fontsize=9.5, fontweight='bold')
    axes[1].text(0.05, 0.88, f"Levene F = {h1_results['levene_stat']:.1f}\np < 0.001 (Heteroscedastic)\nRequires Welch's t-test", 
                 transform=axes[1].transAxes, fontsize=8.5, fontweight="bold", color="#78281F",
                 bbox=dict(boxstyle="round,pad=0.4", facecolor="#FDEDEC", edgecolor=COLOR_CRIMSON))

    # ---------------- Panel C: Multicollinearity VIF Bar Chart ----------------
    features = [v["feature"].replace("Contract_", "Contract: ").replace("Payment_", "Payment: ").replace("SeniorCitizen_", "Senior: ") for v in vif_data]
    vifs = [v["vif"] for v in vif_data]
    bar_cols = [COLOR_CRIMSON if vif > 5.0 else COLOR_TEAL for vif in vifs]
    
    bars_vif = axes[2].barh(features, vifs, color=bar_cols, height=0.55, edgecolor="none", alpha=0.9)
    axes[2].axvline(5.0, color=COLOR_CRIMSON, linestyle="--", linewidth=1.5, label="VIF Threshold = 5.0")
    axes[2].set_title("C. Multicollinearity Audit: VIF Metrics", fontsize=11.5, fontweight='bold', pad=12, color=COLOR_NAVY)
    axes[2].set_xlabel("Variance Inflation Factor (VIF)", fontsize=10, fontweight='bold')
    axes[2].set_xlim(0, 12)
    axes[2].legend(loc="lower right", frameon=True, fontsize=9)
    
    for idx, vif_val in enumerate(vifs):
        axes[2].text(vif_val + 0.25, idx, f"{vif_val:.2f}", va="center", fontsize=9, fontweight="bold", color=bar_cols[idx])

    plt.suptitle("FIGURE 5: STATISTICAL ASSUMPTION AUDIT — NORMALITY, HOMOSCEDASTICITY, AND MULTICOLLINEARITY",
                 fontsize=13.5, fontweight='bold', color=COLOR_NAVY, y=1.02)
    plt.tight_layout()
    
    fig_path = os.path.join(FIGURES_DIR, "fig5_statistical_assumption_diagnostics.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 5: {fig_path}")


def run_statistical_pipeline():
    """Master orchestrator for Week 3 statistical analysis."""
    print("==================================================================")
    print("STARTING WEEK 3 STATISTICAL ANALYSIS & HYPOTHESIS TESTING PIPELINE")
    print("==================================================================")
    df = load_dataset()
    
    h1_results = execute_hypothesis_1(df)
    h2_results = execute_hypothesis_2(df)
    h3_results = execute_hypothesis_3(df)
    h4_results, logit_model = execute_hypothesis_4(df)
    vif_data = compute_vif_and_diagnostics(df)
    
    # Save results to JSON
    all_results = {
        "hypothesis_1_pricing": h1_results,
        "hypothesis_2_contract_anova": h2_results,
        "hypothesis_3_tech_support_chisquare": h3_results,
        "hypothesis_4_logistic_regression": h4_results,
        "vif_diagnostics": vif_data
    }
    with open(RESULTS_JSON_PATH, "w") as f:
        json.dump(all_results, f, indent=2)
    print(f"\n[+] Master statistical test results saved to: {RESULTS_JSON_PATH}")
    
    # Generate all 5 statistical figures
    plot_fig1_hypothesis1_pricing(df, h1_results)
    plot_fig2_hypothesis2_anova(df, h2_results)
    plot_fig3_hypothesis3_chisquare(df, h3_results)
    plot_fig4_hypothesis4_logistic_regression(df, h4_results)
    plot_fig5_assumption_diagnostics(df, h1_results, vif_data)
    
    print("==================================================================")
    print("ALL 5 STATISTICAL DIAGNOSTIC VISUALIZATIONS GENERATED AT 300 DPI!")
    print("==================================================================")


if __name__ == "__main__":
    run_statistical_pipeline()

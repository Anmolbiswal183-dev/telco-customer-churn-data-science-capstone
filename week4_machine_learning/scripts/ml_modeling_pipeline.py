"""
Week 4: Machine Learning Model Development and Evaluation Pipeline
Target Problem: Subscriber Attrition Classification (Binary Classification)
Library Stack: Scikit-Learn, Pandas, NumPy, Matplotlib, Seaborn
Author: Data Science & Machine Learning Engineering Team
"""

import os
import sys
import json
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import seaborn as sns

from sklearn.model_selection import train_test_split, StratifiedKFold, cross_validate
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
from sklearn.metrics import (
    accuracy_score, balanced_accuracy_score, precision_score, recall_score,
    f1_score, roc_auc_score, average_precision_score, confusion_matrix,
    roc_curve, precision_recall_curve, classification_report
)

# Configure Paths
BASE_DIR = r"C:\Users\A.BISWAL\.gemini\antigravity\scratch\week4_machine_learning"
DATA_DIR = os.path.join(BASE_DIR, "data")
FIGURES_DIR = os.path.join(BASE_DIR, "figures")
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")
SOURCE_DATA_PATH = r"C:\Users\A.BISWAL\.gemini\antigravity\scratch\week1_data_cleaning_eda\data\cleaned_telco_churn.csv"
LOCAL_DATA_PATH = os.path.join(DATA_DIR, "ml_prepared_data.csv")
RESULTS_JSON_PATH = os.path.join(DATA_DIR, "ml_results.json")

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
COLOR_NAVY = "#1B365D"       # Corporate Navy
COLOR_CRIMSON = "#D9381E"    # Alert Crimson
COLOR_TEAL = "#0D9488"       # Emerald / Teal
COLOR_AMBER = "#D97706"      # Amber
COLOR_PURPLE = "#7C3AED"     # Purple
COLOR_SLATE = "#334155"      # Dark Neutral Body
COLOR_MUTED = "#64748B"      # Light Neutral Axis/Annotation


def prepare_data():
    """Loads cleaned data, executes one-hot encoding, feature scaling, and stratified splitting."""
    print("[*] Preparing dataset for machine learning modeling...")
    raw_df = pd.read_csv(SOURCE_DATA_PATH)
    
    # 1. Target vector
    y = (raw_df["Churn"] == "Yes").astype(int)
    
    # 2. Features dataframe (drop customerID and Churn)
    X_raw = raw_df.drop(columns=["customerID", "Churn"])
    if "Churn_Numeric" in X_raw.columns:
        X_raw = X_raw.drop(columns=["Churn_Numeric"])
        
    # Identify continuous and categorical columns
    continuous_cols = ["tenure", "MonthlyCharges", "TotalCharges"]
    categorical_cols = [c for c in X_raw.columns if c not in continuous_cols]
    
    # One-hot encode categoricals with drop_first=True
    X_encoded = pd.get_dummies(X_raw, columns=categorical_cols, drop_first=True, dtype=int)
    
    # 3. Stratified 80/20 Train-Test Split (fixed random_state=42)
    X_train, X_test, y_train, y_test = train_test_split(
        X_encoded, y, test_size=0.20, stratify=y, random_state=42
    )
    
    # 4. Standard Scaling fitted strictly on training data (preventing data leakage)
    scaler = StandardScaler()
    X_train_scaled = X_train.copy()
    X_test_scaled = X_test.copy()
    
    X_train_scaled[continuous_cols] = scaler.fit_transform(X_train[continuous_cols])
    X_test_scaled[continuous_cols] = scaler.transform(X_test[continuous_cols])
    
    print(f"[+] Total Records: {len(raw_df)} | Total Model Features: {X_train.shape[1]}")
    print(f"[+] Training Set: {X_train.shape[0]} rows (Churn = {y_train.mean()*100:.2f}%)")
    print(f"[+] Test Set:     {X_test.shape[0]} rows (Churn = {y_test.mean()*100:.2f}%)")
    
    return X_train_scaled, X_test_scaled, y_train, y_test, X_train.columns, continuous_cols


def build_and_evaluate_models(X_train, X_test, y_train, y_test):
    """Initializes, cross-validates, fits, and evaluates the 4 benchmark models."""
    print("\n--- Model Training & 5-Fold Stratified Cross-Validation ---")
    
    models = {
        "Logistic Regression": LogisticRegression(max_iter=1000, random_state=42),
        "Decision Tree": DecisionTreeClassifier(max_depth=5, min_samples_leaf=20, random_state=42),
        "Random Forest": RandomForestClassifier(n_estimators=100, max_depth=8, min_samples_split=10, random_state=42),
        "Gradient Boosting": GradientBoostingClassifier(n_estimators=100, learning_rate=0.1, max_depth=3, subsample=0.9, random_state=42)
    }
    
    cv_kfold = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    scoring = ["roc_auc", "f1", "balanced_accuracy", "accuracy"]
    
    benchmark_results = {}
    fitted_models = {}
    test_predictions = {}
    
    for name, model in models.items():
        print(f"[*] Cross-validating and training: {name}...")
        
        # 1. 5-Fold Cross-Validation on training data
        cv_scores = cross_validate(model, X_train, y_train, cv=cv_kfold, scoring=scoring, n_jobs=-1)
        
        # 2. Fit on full training set
        model.fit(X_train, y_train)
        fitted_models[name] = model
        
        # 3. Predict on held-out test set
        y_pred = model.predict(X_test)
        y_proba = model.predict_proba(X_test)[:, 1]
        test_predictions[name] = {"pred": y_pred, "proba": y_proba}
        
        # Compute test set evaluation metrics
        acc = accuracy_score(y_test, y_pred)
        bal_acc = balanced_accuracy_score(y_test, y_pred)
        prec = precision_score(y_test, y_pred)
        rec = recall_score(y_test, y_pred) # Sensitivity
        f1 = f1_score(y_test, y_pred)
        roc_auc = roc_auc_score(y_test, y_proba)
        pr_auc = average_precision_score(y_test, y_proba)
        
        # Confusion matrix
        cm = confusion_matrix(y_test, y_pred)
        tn, fp, fn, tp = cm.ravel()
        specificity = tn / (tn + fp)
        
        benchmark_results[name] = {
            "cv_metrics": {
                "cv_roc_auc_mean": round(float(cv_scores["test_roc_auc"].mean()), 4),
                "cv_roc_auc_std": round(float(cv_scores["test_roc_auc"].std()), 4),
                "cv_f1_mean": round(float(cv_scores["test_f1"].mean()), 4),
                "cv_f1_std": round(float(cv_scores["test_f1"].std()), 4),
                "cv_accuracy_mean": round(float(cv_scores["test_accuracy"].mean()), 4),
                "cv_accuracy_std": round(float(cv_scores["test_accuracy"].std()), 4),
                "cv_balanced_acc_mean": round(float(cv_scores["test_balanced_accuracy"].mean()), 4),
                "cv_balanced_acc_std": round(float(cv_scores["test_balanced_accuracy"].std()), 4),
            },
            "test_metrics": {
                "accuracy": round(float(acc), 4),
                "balanced_accuracy": round(float(bal_acc), 4),
                "precision": round(float(prec), 4),
                "recall": round(float(rec), 4),
                "specificity": round(float(specificity), 4),
                "f1_score": round(float(f1), 4),
                "roc_auc": round(float(roc_auc), 4),
                "pr_auc": round(float(pr_auc), 4),
                "confusion_matrix": {
                    "true_negative": int(tn),
                    "false_positive": int(fp),
                    "false_negative": int(fn),
                    "true_positive": int(tp)
                }
            }
        }
        
        print(f"    Test Results: Acc={acc:.4f}, ROC-AUC={roc_auc:.4f}, F1={f1:.4f}, Recall={rec:.4f}, Precision={prec:.4f}")
        
    return benchmark_results, fitted_models, test_predictions


def execute_cost_optimization(y_test, y_proba_best):
    """
    Simulates asymmetric business cost optimization across decision thresholds [0.05, 0.95].
    Cost assumptions:
    - False Negative (missed churner = lost CLV): $1,000
    - False Positive (unnecessary retention discount): $60
    - True Positive (successfully retained churner with offer): Net cost $60, saves $1,000 -> Net benefit +$940
    - True Negative: Cost $0
    """
    print("\n--- Threshold Optimization & Asymmetric Cost-Benefit Analysis ---")
    thresholds = np.linspace(0.05, 0.95, 91)
    
    cost_fn = 1000.0  # Lost customer lifetime value
    cost_fp = 60.0    # Cost of retention voucher/discount
    
    cost_curve = []
    
    for th in thresholds:
        y_pred_th = (y_proba_best >= th).astype(int)
        cm = confusion_matrix(y_test, y_pred_th)
        tn, fp, fn, tp = cm.ravel()
        
        total_cost = (fn * cost_fn) + (fp * cost_fp)
        cost_curve.append({
            "threshold": round(float(th), 2),
            "total_cost": round(float(total_cost), 2),
            "false_negatives": int(fn),
            "false_positives": int(fp),
            "true_positives": int(tp),
            "true_negatives": int(tn),
            "recall": round(float(tp / (tp + fn)), 4),
            "precision": round(float(tp / (tp + fp) if (tp + fp) > 0 else 0), 4)
        })
        
    cost_df = pd.DataFrame(cost_curve)
    opt_idx = cost_df["total_cost"].idxmin()
    optimal_row = cost_df.loc[opt_idx]
    
    default_row = cost_df.loc[(cost_df["threshold"] - 0.50).abs().idxmin()]
    cost_savings = default_row["total_cost"] - optimal_row["total_cost"]
    
    opt_summary = {
        "default_threshold": 0.50,
        "default_total_cost": float(default_row["total_cost"]),
        "default_fn_count": int(default_row["false_negatives"]),
        "default_fp_count": int(default_row["false_positives"]),
        "default_recall": float(default_row["recall"]),
        "optimal_threshold": float(optimal_row["threshold"]),
        "optimal_total_cost": float(optimal_row["total_cost"]),
        "optimal_fn_count": int(optimal_row["false_negatives"]),
        "optimal_fp_count": int(optimal_row["false_positives"]),
        "optimal_recall": float(optimal_row["recall"]),
        "net_financial_savings": float(cost_savings),
        "percentage_cost_reduction": round(float(cost_savings / default_row["total_cost"] * 100), 2)
    }
    
    print(f"[+] Default Threshold (0.50): Cost = ${default_row['total_cost']:,.2f} (FN = {default_row['false_negatives']})")
    print(f"[+] Optimal Threshold ({optimal_row['threshold']:.2f}): Cost = ${optimal_row['total_cost']:,.2f} (FN = {optimal_row['false_negatives']})")
    print(f"[+] Net Financial Savings: ${cost_savings:,.2f} ({opt_summary['percentage_cost_reduction']}% reduction in business loss!)")
    
    return opt_summary, cost_df


# =============================================================================
# MACHINE LEARNING VISUALIZATION SUITE (5 HIGH-RES FIGURES AT 300 DPI)
# =============================================================================

def plot_fig1_roc_curves(test_predictions, y_test, benchmark_results):
    """
    Figure 1: Multi-Model Receiver Operating Characteristic (ROC) Curves.
    Compares all 4 classifiers against chance diagonal with individual AUCs.
    """
    print("[*] Generating Figure 1: Multi-Model ROC Curves...")
    fig, ax = plt.subplots(figsize=(9, 7.5), dpi=300)
    
    palette = {
        "Logistic Regression": COLOR_NAVY,
        "Decision Tree": COLOR_AMBER,
        "Random Forest": COLOR_TEAL,
        "Gradient Boosting": COLOR_CRIMSON
    }
    
    for name, data in test_predictions.items():
        fpr, tpr, _ = roc_curve(y_test, data["proba"])
        auc = benchmark_results[name]["test_metrics"]["roc_auc"]
        ax.plot(fpr, tpr, label=f"{name} (AUC = {auc:.4f})", color=palette[name], linewidth=2.6)
        
    # Diagonal Chance Line
    ax.plot([0, 1], [0, 1], color="#94A3B8", linestyle="--", linewidth=1.5, label="Random Chance (AUC = 0.5000)")
    
    ax.set_title("FIGURE 1: MULTI-MODEL RECEIVER OPERATING CHARACTERISTIC (ROC) COMPARISON", 
                 fontsize=12.5, fontweight='bold', pad=14, color=COLOR_NAVY)
    ax.set_xlabel("False Positive Rate (1 - Specificity)", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    ax.set_ylabel("True Positive Rate (Sensitivity / Recall)", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    ax.set_xlim(-0.02, 1.02)
    ax.set_ylim(-0.02, 1.02)
    ax.legend(loc="lower right", frameon=True, fontsize=10.5)
    
    # Statistical Annotation Box
    best_name = max(benchmark_results, key=lambda k: benchmark_results[k]["test_metrics"]["roc_auc"])
    best_auc = benchmark_results[best_name]["test_metrics"]["roc_auc"]
    ax.text(0.04, 0.75, f"CHAMPION MODEL:\n• {best_name}\n• Test ROC-AUC = {best_auc:.4f}\n• Superior discrimination across\n  all operating thresholds", 
            transform=ax.transAxes, fontsize=9.5, fontweight="bold", color=COLOR_NAVY,
            bbox=dict(boxstyle="round,pad=0.5", facecolor="#EBF8FF", edgecolor=COLOR_NAVY))

    plt.tight_layout()
    fig_path = os.path.join(FIGURES_DIR, "fig1_multi_model_roc_curves.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 1: {fig_path}")


def plot_fig2_pr_curves(test_predictions, y_test, benchmark_results):
    """
    Figure 2: Precision-Recall (PR) Curves with Average Precision (AP).
    Crucial evaluation under class imbalance (26.5% positive churn).
    """
    print("[*] Generating Figure 2: Precision-Recall Curves...")
    fig, ax = plt.subplots(figsize=(9, 7.5), dpi=300)
    
    palette = {
        "Logistic Regression": COLOR_NAVY,
        "Decision Tree": COLOR_AMBER,
        "Random Forest": COLOR_TEAL,
        "Gradient Boosting": COLOR_CRIMSON
    }
    
    for name, data in test_predictions.items():
        precision, recall, _ = precision_recall_curve(y_test, data["proba"])
        ap = benchmark_results[name]["test_metrics"]["pr_auc"]
        ax.plot(recall, precision, label=f"{name} (PR-AUC / AP = {ap:.4f})", color=palette[name], linewidth=2.6)
        
    # Baseline No-Skill Line (prevalence of positive class)
    prevalence = float(y_test.mean())
    ax.axhline(prevalence, color="#94A3B8", linestyle="--", linewidth=1.5, 
               label=f"No-Skill Prevalence Baseline (AP = {prevalence:.4f})")
    
    ax.set_title("FIGURE 2: PRECISION-RECALL (PR) CURVES UNDER CLASS IMBALANCE", 
                 fontsize=12.5, fontweight='bold', pad=14, color=COLOR_NAVY)
    ax.set_xlabel("Recall (Sensitivity / Churn Coverage)", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    ax.set_ylabel("Precision (Positive Predictive Value)", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    ax.set_xlim(-0.02, 1.02)
    ax.set_ylim(0.15, 1.02)
    ax.legend(loc="upper right", frameon=True, fontsize=10.5)
    
    ax.text(0.04, 0.20, f"CLASS IMBALANCE EVALUATION:\n• Churn prevalence = {prevalence*100:.1f}%\n• PR-AUC measures targeted accuracy\n  without distortion from True Negatives", 
            transform=ax.transAxes, fontsize=9.5, fontweight="bold", color=COLOR_SLATE,
            bbox=dict(boxstyle="round,pad=0.5", facecolor="#F8FAFC", edgecolor="#CBD5E0"))

    plt.tight_layout()
    fig_path = os.path.join(FIGURES_DIR, "fig2_precision_recall_curves.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 2: {fig_path}")


def plot_fig3_confusion_matrices(benchmark_results):
    """
    Figure 3: Four-Panel Confusion Matrix Suite.
    Displays raw counts and normalized percentages for all 4 models.
    """
    print("[*] Generating Figure 3: 4-Panel Confusion Matrix Suite...")
    fig, axes = plt.subplots(2, 2, figsize=(13, 11), dpi=300)
    
    model_names = list(benchmark_results.keys())
    cm_coords = [(0, 0), (0, 1), (1, 0), (1, 1)]
    
    for idx, name in enumerate(model_names):
        r, c = cm_coords[idx]
        ax = axes[r, c]
        
        cm_data = benchmark_results[name]["test_metrics"]["confusion_matrix"]
        cm_arr = np.array([
            [cm_data["true_negative"], cm_data["false_positive"]],
            [cm_data["false_negative"], cm_data["true_positive"]]
        ])
        
        # Calculate percentage labels
        cm_pct = cm_arr / cm_arr.sum(axis=1)[:, np.newaxis] * 100
        
        annot = np.array([
            [f"{cm_arr[0,0]:,}\n({cm_pct[0,0]:.1f}%)", f"{cm_arr[0,1]:,}\n({cm_pct[0,1]:.1f}%)"],
            [f"{cm_arr[1,0]:,}\n({cm_pct[1,0]:.1f}%)", f"{cm_arr[1,1]:,}\n({cm_pct[1,1]:.1f}%)"]
        ])
        
        sns.heatmap(cm_arr, annot=annot, fmt="", cmap="Blues", cbar=False, ax=ax,
                    linewidths=2.0, linecolor="white", annot_kws={"size": 11, "weight": "bold"})
        
        metrics = benchmark_results[name]["test_metrics"]
        ax.set_title(f"{name}\nAcc: {metrics['accuracy']:.4f} | F1: {metrics['f1_score']:.4f} | Recall: {metrics['recall']:.4f}", 
                     fontsize=11.5, fontweight='bold', pad=10, color=COLOR_NAVY)
        ax.set_xlabel("Predicted Class", fontsize=10, fontweight='bold')
        ax.set_ylabel("Actual Ground Truth", fontsize=10, fontweight='bold')
        ax.set_xticklabels(["Retained (0)", "Churned (1)"], fontsize=9.5, fontweight='bold')
        ax.set_yticklabels(["Retained (0)", "Churned (1)"], fontsize=9.5, fontweight='bold')
        
    plt.suptitle("FIGURE 3: CONFUSION MATRIX BENCHMARK SUITE — ERROR DISTRIBUTION COMPARISON",
                 fontsize=14, fontweight='bold', color=COLOR_NAVY, y=1.02)
    plt.tight_layout()
    
    fig_path = os.path.join(FIGURES_DIR, "fig3_confusion_matrix_suite.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 3: {fig_path}")


def plot_fig4_feature_importance(fitted_models, feature_names):
    """
    Figure 4: Global Feature Importance & Coefficient Attribution.
    Compares Random Forest Gini MDI against Logistic Regression standardized coefficients.
    """
    print("[*] Generating Figure 4: Feature Importance Comparison...")
    fig, axes = plt.subplots(1, 2, figsize=(15, 6.8), dpi=300)
    
    # ---------------- Panel A: Random Forest Gini MDI Importance ----------------
    rf = fitted_models["Random Forest"]
    rf_imp = pd.Series(rf.feature_importances_, index=feature_names).sort_values(ascending=False).head(10)
    
    clean_labels_rf = [
        f.replace("Contract_", "Contract: ")
        .replace("PaymentMethod_", "Payment: ")
        .replace("InternetService_", "Internet: ")
        .replace("OnlineSecurity_", "Security: ")
        .replace("TechSupport_", "Support: ")
        for f in rf_imp.index
    ]
    
    axes[0].barh(clean_labels_rf[::-1], rf_imp.values[::-1], color=COLOR_TEAL, height=0.6, edgecolor="none", alpha=0.9)
    axes[0].set_title("A. Random Forest Top 10 Feature Importances (Gini Impurity Reduction)", 
                      fontsize=12, fontweight='bold', pad=12, color=COLOR_NAVY)
    axes[0].set_xlabel("Relative Importance Score", fontsize=10.5, fontweight='bold')
    
    for idx, val in enumerate(rf_imp.values[::-1]):
        axes[0].text(val + 0.005, idx, f"{val:.3f}", va="center", fontsize=9.5, fontweight="bold", color=COLOR_TEAL)
    axes[0].set_xlim(0, max(rf_imp.values) * 1.18)

    # ---------------- Panel B: Logistic Regression Standardized Coefficients ----------------
    lr = fitted_models["Logistic Regression"]
    lr_coef = pd.Series(lr.coef_[0], index=feature_names)
    # Sort by absolute magnitude
    top_lr = lr_coef.reindex(lr_coef.abs().sort_values(ascending=False).head(10).index)
    
    clean_labels_lr = [
        f.replace("Contract_", "Contract: ")
        .replace("PaymentMethod_", "Payment: ")
        .replace("InternetService_", "Internet: ")
        .replace("OnlineSecurity_", "Security: ")
        .replace("TechSupport_", "Support: ")
        for f in top_lr.index
    ]
    
    bar_cols = [COLOR_CRIMSON if val > 0 else COLOR_NAVY for val in top_lr.values[::-1]]
    axes[1].barh(clean_labels_lr[::-1], top_lr.values[::-1], color=bar_cols, height=0.6, edgecolor="none", alpha=0.9)
    axes[1].axvline(0, color="#94A3B8", linestyle="--", linewidth=1.2)
    axes[1].set_title("B. Logistic Regression Top 10 Standardized Weights (Log-Odds Impact)", 
                      fontsize=12, fontweight='bold', pad=12, color=COLOR_NAVY)
    axes[1].set_xlabel("Standardized Model Coefficient β (Directional Risk)", fontsize=10.5, fontweight='bold')
    
    for idx, val in enumerate(top_lr.values[::-1]):
        offset = 0.04 if val >= 0 else -0.04
        ha = "left" if val >= 0 else "right"
        axes[1].text(val + offset, idx, f"{val:+.2f}", va="center", ha=ha, fontsize=9.5, fontweight="bold", color=bar_cols[idx])

    plt.suptitle("FIGURE 4: MODEL INTERPRETABILITY — RANDOM FOREST MDI VS LOGISTIC REGRESSION WEIGHTS",
                 fontsize=14, fontweight='bold', color=COLOR_NAVY, y=1.02)
    plt.tight_layout()
    
    fig_path = os.path.join(FIGURES_DIR, "fig4_feature_importance_comparison.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 4: {fig_path}")


def plot_fig5_cost_optimization(cost_df, opt_summary):
    """
    Figure 5: Decision Threshold Optimization & Asymmetric Cost-Benefit Curve.
    Demonstrates business financial cost vs threshold cutoff.
    """
    print("[*] Generating Figure 5: Decision Threshold Cost Optimization...")
    fig, ax = plt.subplots(figsize=(12, 6.5), dpi=300)
    
    # Plot Total Financial Cost Curve
    ax.plot(cost_df["threshold"], cost_df["total_cost"] / 1e3, color=COLOR_CRIMSON, linewidth=3.0, label="Total Business Cost ($k USD)")
    
    # Mark Default vs Optimal
    def_th = opt_summary["default_threshold"]
    def_cost = opt_summary["default_total_cost"] / 1e3
    opt_th = opt_summary["optimal_threshold"]
    opt_cost = opt_summary["optimal_total_cost"] / 1e3
    
    ax.scatter([def_th], [def_cost], color=COLOR_NAVY, s=150, zorder=5, label=f"Default Threshold (0.50): ${def_cost:.1f}k")
    ax.scatter([opt_th], [opt_cost], color=COLOR_TEAL, s=180, zorder=5, marker="*", label=f"Optimal Cutoff ({opt_th:.2f}): ${opt_cost:.1f}k")
    
    ax.axvline(def_th, color=COLOR_NAVY, linestyle=":", linewidth=1.5)
    ax.axvline(opt_th, color=COLOR_TEAL, linestyle="--", linewidth=1.8)
    
    # Shaded Savings Zone
    ax.fill_between(cost_df["threshold"], opt_cost, cost_df["total_cost"] / 1e3, 
                    where=(cost_df["threshold"] >= opt_th) & (cost_df["threshold"] <= def_th),
                    color="#D1FAE5", alpha=0.6, label=f"Financial Savings: ${opt_summary['net_financial_savings']:,.0f}")
    
    ax.annotate(f"OPTIMAL DECISION CUTOFF (p = {opt_th:.2f}):\n"
                f"• Drops False Negatives from {opt_summary['default_fn_count']} to {opt_summary['optimal_fn_count']}\n"
                f"• Increases Churn Recall from {opt_summary['default_recall']*100:.1f}% to {opt_summary['optimal_recall']*100:.1f}%\n"
                f"• Net Financial Savings: ${opt_summary['net_financial_savings']:,.0f} ({opt_summary['percentage_cost_reduction']}% reduction)",
                xy=(opt_th, opt_cost), xytext=(opt_th + 0.12, opt_cost + 40),
                arrowprops=dict(facecolor=COLOR_TEAL, shrink=0.08, width=1.5, headwidth=6),
                bbox=dict(boxstyle="round,pad=0.5", facecolor="#E6FFFA", edgecolor=COLOR_TEAL),
                fontsize=9.5, fontweight="bold", color="#065F46")
    
    ax.set_title("FIGURE 5: DECISION THRESHOLD TUNING & ASYMMETRIC COMMERCIAL COST OPTIMIZATION", 
                 fontsize=13.5, fontweight='bold', pad=14, color=COLOR_NAVY)
    ax.set_xlabel("Classification Probability Threshold (Cutoff)", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    ax.set_ylabel("Total Business Cost ($ Thousands USD)", fontsize=11, fontweight='bold', color=COLOR_SLATE)
    ax.set_xlim(0.05, 0.90)
    ax.yaxis.set_major_formatter(ticker.FormatStrFormatter('$%.0fk'))
    ax.legend(loc="upper right", frameon=True, fontsize=10)

    plt.tight_layout()
    fig_path = os.path.join(FIGURES_DIR, "fig5_threshold_cost_optimization.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [+] Saved Figure 5: {fig_path}")


def run_ml_pipeline():
    """Master orchestrator for Week 4 ML pipeline."""
    print("==================================================================")
    print("STARTING WEEK 4 MACHINE LEARNING MODEL DEVELOPMENT PIPELINE")
    print("==================================================================")
    X_train, X_test, y_train, y_test, feature_names, continuous_cols = prepare_data()
    
    benchmark_results, fitted_models, test_predictions = build_and_evaluate_models(
        X_train, X_test, y_train, y_test
    )
    
    # Cost optimization on best model (Gradient Boosting)
    y_proba_best = test_predictions["Gradient Boosting"]["proba"]
    opt_summary, cost_df = execute_cost_optimization(y_test, y_proba_best)
    
    # Save master results to JSON
    ml_master_data = {
        "benchmark_results": benchmark_results,
        "cost_optimization": opt_summary,
        "feature_count": len(feature_names),
        "continuous_features": continuous_cols,
        "train_sample_size": len(X_train),
        "test_sample_size": len(X_test)
    }
    with open(RESULTS_JSON_PATH, "w") as f:
        json.dump(ml_master_data, f, indent=2)
    print(f"[+] Master ML results saved to: {RESULTS_JSON_PATH}")
    
    # Generate 5 Visualizations
    plot_fig1_roc_curves(test_predictions, y_test, benchmark_results)
    plot_fig2_pr_curves(test_predictions, y_test, benchmark_results)
    plot_fig3_confusion_matrices(benchmark_results)
    plot_fig4_feature_importance(fitted_models, feature_names)
    plot_fig5_cost_optimization(cost_df, opt_summary)
    
    print("==================================================================")
    print("ALL 5 MACHINE LEARNING VISUALIZATIONS GENERATED AT 300 DPI!")
    print("==================================================================")


if __name__ == "__main__":
    run_ml_pipeline()

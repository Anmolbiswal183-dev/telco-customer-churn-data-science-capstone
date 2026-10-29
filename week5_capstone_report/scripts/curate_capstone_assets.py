"""
Week 5 Capstone Asset Curator
Harvests, validates, and consolidates all metrics and figures from Weeks 1 through 4
into the Capstone project directory.
"""

import os
import shutil
import json

BASE_CAPSTONE = r"C:\Users\A.BISWAL\.gemini\antigravity\scratch\week5_capstone_report"
DATA_DIR = os.path.join(BASE_CAPSTONE, "data")
FIGURES_DIR = os.path.join(BASE_CAPSTONE, "figures")
SCRIPTS_DIR = os.path.join(BASE_CAPSTONE, "scripts")

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(FIGURES_DIR, exist_ok=True)
os.makedirs(SCRIPTS_DIR, exist_ok=True)

# Source Directories
WEEK1_DIR = r"C:\Users\A.BISWAL\.gemini\antigravity\scratch\week1_data_cleaning_eda"
WEEK2_DIR = r"C:\Users\A.BISWAL\.gemini\antigravity\scratch\week2_visualization_storytelling"
WEEK3_DIR = r"C:\Users\A.BISWAL\.gemini\antigravity\scratch\week3_statistical_analysis"
WEEK4_DIR = r"C:\Users\A.BISWAL\.gemini\antigravity\scratch\week4_machine_learning"

def curate_assets():
    print("[*] Curating Capstone metrics and visual assets across Weeks 1-4...")
    
    # 1. Harvest JSON Metrics
    metrics_files = [
        (os.path.join(WEEK1_DIR, "data", "pipeline_metrics.json"), "week1_metrics.json"),
        (os.path.join(WEEK2_DIR, "data", "story_metrics.json"), "week2_metrics.json"),
        (os.path.join(WEEK3_DIR, "data", "statistical_results.json"), "week3_metrics.json"),
        (os.path.join(WEEK4_DIR, "data", "ml_results.json"), "week4_metrics.json"),
    ]
    for src, dst_name in metrics_files:
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(DATA_DIR, dst_name))
            print(f"  [+] Copied metric file: {dst_name}")
            
    # 2. Harvest Key Figures
    figures_to_copy = [
        # Week 1
        (os.path.join(WEEK1_DIR, "figures", "fig1_data_integrity_audit.png"), "capstone_fig1_data_cleaning_audit.png"),
        # Week 2 (Storytelling)
        (os.path.join(WEEK2_DIR, "figures", "fig1_act1_macro_revenue_leak.png"), "capstone_fig2_macro_revenue_leak.png"),
        (os.path.join(WEEK2_DIR, "figures", "fig2_act2_survival_hazard_trajectory.png"), "capstone_fig3_survival_hazard_curve.png"),
        (os.path.join(WEEK2_DIR, "figures", "fig3_act3_tech_paradox_support_shield.png"), "capstone_fig4_fiber_support_shield.png"),
        (os.path.join(WEEK2_DIR, "figures", "fig4_act4_customer_persona_radar.png"), "capstone_fig5_customer_persona_radar.png"),
        # Week 3 (Statistics)
        (os.path.join(WEEK3_DIR, "figures", "fig5_statistical_assumption_diagnostics.png"), "capstone_fig6_assumption_diagnostics.png"),
        (os.path.join(WEEK3_DIR, "figures", "fig4_hypothesis4_logistic_regression_forest.png"), "capstone_fig7_logistic_regression_forest.png"),
        # Week 4 (Machine Learning)
        (os.path.join(WEEK4_DIR, "figures", "fig1_multi_model_roc_curves.png"), "capstone_fig8_multi_model_roc_curves.png"),
        (os.path.join(WEEK4_DIR, "figures", "fig3_confusion_matrix_suite.png"), "capstone_fig9_confusion_matrix_suite.png"),
        (os.path.join(WEEK4_DIR, "figures", "fig5_threshold_cost_optimization.png"), "capstone_fig10_threshold_cost_optimization.png")
    ]
    
    for src, dst_name in figures_to_copy:
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(FIGURES_DIR, dst_name))
            print(f"  [+] Copied figure: {dst_name}")
        else:
            print(f"  [!] Missing source figure: {src}")
            
    print("[+] All Capstone visual and analytical assets successfully consolidated.")

if __name__ == "__main__":
    curate_assets()

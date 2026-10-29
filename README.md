# Enterprise Telco Customer Churn & Retention Analytics Portfolio
### End-to-End 5-Week Data Science, Machine Learning & Strategic Capstone

[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Scikit-Learn](https://img.shields.io/badge/scikit--learn-v1.4%2B-orange.svg)](https://scikit-learn.org/)
[![Pandas](https://img.shields.io/badge/pandas-v2.2%2B-darkblue.svg)](https://pandas.pydata.org/)
[![Status](https://img.shields.io/badge/Status-Completed%20(Weeks%201--5)-success.svg)]()

> **Master Capstone Project:** A comprehensive 5-week data science curriculum investigating customer attrition across the **IBM Telco Customer Churn ecosystem** ($N = 7,043$ accounts, 21 attributes). Synthesizes forensic data engineering, visual storytelling, inferential hypothesis testing, multi-model machine learning tournaments, asymmetric business cost optimization, and C-suite strategic playbooks protecting **$1.375 Million in Annual Recurring Revenue (ARR)**.

---

## 📑 Executive Deliverables Directory

All weekly coursework modules have been documented in publication-grade Microsoft Word (`.docx`) reports with embedded 300-DPI visual assets, formal statistical test suites, and reproducible code appendices:

| Module | Report Deliverable (.docx) | Technical Scope | Key Deliverable Highlights |
| :--- | :--- | :--- | :--- |
| **Week 1** | [Week 1 Report](deliverables/Week_1_Data_Acquisition_Cleaning_EDA_Report.docx) | **Data Cleaning & EDA** | Ingestion, unmasking 11 regex whitespace nulls, zero-tenure domain imputation, Tukey boxplot outlier fences, 5 diagnostic figures. |
| **Week 2** | [Week 2 Report](deliverables/Week_2_Advanced_Data_Visualization_Storytelling_Report.docx) | **Visualization & Storytelling** | 5-Act executive narrative arc, $1.67M ARR leak quantification, Kaplan-Meier hazard curves (180-day cliff), tech support paradox, 4 customer personas. |
| **Week 3** | [Week 3 Report](deliverables/Week_3_Statistical_Analysis_Hypothesis_Testing_Report.docx) | **Statistical Inference** | Welch's $t$-test ($p = 4.30 \times 10^{-73}$), One-Way ANOVA ($F = 2,960.02, \eta^2 = 0.457$), $\chi^2$ test ($p = 1.67 \times 10^{-41}$), Multivariate Logistic Regression ($AOR = 5.96$). |
| **Week 4** | [Week 4 Report](deliverables/Week_4_Machine_Learning_Model_Development_Report.docx) | **Machine Learning** | 4-Model tournament (Logistic, Decision Tree, Random Forest, Gradient Boosting; CV ROC-AUC = 0.8468), PR curves, MDI feature ranking, cost optimization ($p^* = 0.23$, saving $82.5K test cohort). |
| **Week 5** | [Week 5 Report](deliverables/Week_5_Comprehensive_Data_Science_Capstone_Report.docx) | **Capstone & Strategy** | **Capstone Master Synthesis:** 10 curated 300 DPI figures, 4 operational playbooks, pro-forma financial ROI model ($220K budget $\rightarrow$ $1.375M ARR protected, **+525% ROI**), MLOps/governance roadmap. |

---

## 📂 Repository File Structure

```text
telco-customer-churn-data-science-capstone/
├── README.md                                  # Executive Portfolio Documentation & Summary
├── requirements.txt                           # Python Dependencies
├── .gitignore                                 # Git Ignore Rules
├── deliverables/                              # Master Word (.docx) Reports (Weeks 1 - 5)
│   ├── Week_1_Data_Acquisition_Cleaning_EDA_Report.docx
│   ├── Week_2_Advanced_Data_Visualization_Storytelling_Report.docx
│   ├── Week_3_Statistical_Analysis_Hypothesis_Testing_Report.docx
│   ├── Week_4_Machine_Learning_Model_Development_Report.docx
│   └── Week_5_Comprehensive_Data_Science_Capstone_Report.docx
├── docs/
│   └── figures/                               # High-Resolution (300 DPI) Visual Assets
│       ├── capstone_fig1_data_cleaning_audit.png
│       ├── capstone_fig2_macro_revenue_leak.png
│       ├── capstone_fig3_survival_hazard_curve.png
│       ├── capstone_fig4_fiber_support_shield.png
│       ├── capstone_fig5_customer_persona_radar.png
│       ├── capstone_fig6_assumption_diagnostics.png
│       ├── capstone_fig7_logistic_regression_forest.png
│       ├── capstone_fig8_multi_model_roc_curves.png
│       ├── capstone_fig9_confusion_matrix_suite.png
│       └── capstone_fig10_threshold_cost_optimization.png
├── week1_data_cleaning_eda/                   # Data Ingestion, Cleaning & EDA Pipeline
│   ├── data/
│   ├── figures/
│   └── scripts/data_pipeline.py
├── week2_visualization_storytelling/          # 5-Act Visual Storytelling Engine
│   ├── data/
│   ├── figures/
│   └── scripts/storytelling_visualizations.py
├── week3_statistical_analysis/                # Inferential Hypothesis Testing Engine
│   ├── data/
│   ├── figures/
│   └── scripts/statistical_testing_pipeline.py
├── week4_machine_learning/                    # 4-Model ML Tournament & Cost Optimization
│   ├── data/
│   ├── figures/
│   └── scripts/ml_modeling_pipeline.py
└── week5_capstone_report/                     # Master Capstone Asset Harvesting & Compilation
    ├── data/
    ├── figures/
    └── scripts/generate_capstone_doc.py
```

---

## 🔬 Curriculum Modules & Empirical Findings

### Week 1: Forensic Data Engineering & Exploratory Profiling
- **Dataset:** 7,043 subscriber records, 21 variables.
- **Forensic Anomaly:** Identified 11 whitespace-padded missing values (`" "`) in `TotalCharges` that bypassed standard null checks.
- **Root Cause & Imputation:** Cross-variable validation established all 11 records had `tenure == 0` months (new accounts yet to complete a billing cycle). Rather than dropping records or imputing the mean, deterministic imputation of `$0.00` preserved sample integrity.

<p align="center">
  <img src="docs/figures/capstone_fig1_data_cleaning_audit.png" alt="Data Cleaning Audit" width="850"/>
</p>

---

### Week 2: Advanced Visual Storytelling & Narrative Economics
- **The \$1.67M ARR Leak:** 26.54% churn rate accounts for **30.5% of total gross monthly billings**, leaking **$139,131 monthly ($1.67M Annualized Recurring Revenue)**.
- **The 180-Day Cliff:** Over **38% of all churn events occur within months 1–6**. Hazard rate peaks at **12.4% in Month 1** before stabilizing under 1.5% post-Month 24.
- **The Fiber Optic Support Paradox:** Fiber optic subscribers experience an unassisted churn rate of **53.6%**. Adding Tech Support and Online Security reduces churn to **18.2% (a 66% relative risk reduction)**.

<p align="center">
  <img src="docs/figures/capstone_fig3_survival_hazard_curve.png" alt="Survival Hazard Curve" width="48%"/>
  <img src="docs/figures/capstone_fig4_fiber_support_shield.png" alt="Fiber Support Shield" width="48%"/>
</p>

---

### Week 3: Inferential Statistics & Hypothesis Testing
Four formal hypotheses were evaluated and validated:
1. **Hypothesis 1 (Pricing Impact):** Welch's $t(4,135.8) = 18.41, p = 4.30 \times 10^{-73}$, Cohen's $d = 0.446$. Churners pay a significant +\$13.18/mo premium ($M = \$74.44$ vs $\$61.27$).
2. **Hypothesis 2 (Contract Longevity):** One-Way ANOVA $F(2, 7040) = 2,960.02, p < 0.001, \eta^2 = 0.457$. Contract terms explain 45.7% of tenure variance. Tukey HSD confirmed 2-year contracts add +48.69 months over month-to-month plans.
3. **Hypothesis 3 (Tech Support Shield):** $\chi^2(1) = 182.12, p = 1.67 \times 10^{-41}$, Odds Ratio $OR = 3.33$ (95% CI: `[2.79, 3.99]`). Subscribers lacking technical support are 3.33x more likely to churn.
4. **Hypothesis 4 (Multivariate Attribution):** Multivariate Logistic Regression revealed Month-to-Month contracts yield an **Adjusted Odds Ratio ($AOR$) of 5.96** ($p < 0.001$), while Electronic Check billing increases churn odds by **1.70x** ($p < 0.001$).

<p align="center">
  <img src="docs/figures/capstone_fig6_assumption_diagnostics.png" alt="Assumption Diagnostics" width="48%"/>
  <img src="docs/figures/capstone_fig7_logistic_regression_forest.png" alt="Logistic Forest Plot" width="48%"/>
</p>

---

### Week 4: Machine Learning Tournament & Asymmetric Cost Optimization
- **Algorithm Benchmark:** Evaluated across 7,043 records via a stratified 80/20 train/test split ($n_{\text{test}} = 1,409$):

| Algorithm | Test Accuracy | Balanced Acc | Precision | Recall | F1-Score | Test ROC-AUC | Test PR-AUC | 5-Fold CV ROC-AUC |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Logistic Regression** | 80.62% | 72.74% | 65.93% | 55.88% | 0.6049 | 0.8422 | 0.6518 | 0.8415 ± 0.013 |
| **Decision Tree (Pruned)** | 79.56% | 71.59% | 63.35% | 54.55% | 0.5862 | 0.8268 | 0.6120 | 0.8210 ± 0.016 |
| **Random Forest (100 Trees)** | 80.27% | 70.45% | 67.52% | 49.47% | 0.5710 | **0.8443** | 0.6489 | 0.8432 ± 0.012 |
| **Gradient Boosting (Champion)** | 79.91% | 70.63% | 65.74% | 50.80% | 0.5732 | **0.8439** | **0.6543** | **0.8468 ± 0.011** |

- **Asymmetric Cost Curve:** Setting False Negative cost $C_{\text{FN}} = \$1,000$ and False Positive cost $C_{\text{FP}} = \$60$, lowering the decision threshold to $p^* = 0.23$ boosted recall from **50.8% to 83.2%**, slashing loss from **\$189,940 to \$107,380** (saving **\$82,560, a 43.5% reduction**).

<p align="center">
  <img src="docs/figures/capstone_fig8_multi_model_roc_curves.png" alt="Multi-Model ROC" width="48%"/>
  <img src="docs/figures/capstone_fig10_threshold_cost_optimization.png" alt="Cost Curve" width="48%"/>
</p>

---

## 💼 Capstone Strategic Recommendations: 4 Enterprise Playbooks

```mermaid
graph TD
    subgraph Data Science Insights
        I1["38% Churn in First 180 Days"]
        I2["53.6% Churn on Unassisted Fiber"]
        I3["Month-to-Month AOR = 5.96x Risk"]
        I4["Electronic Check AOR = 1.70x Risk"]
    end

    subgraph Strategic Playbooks
        P1["Playbook 1: 90-Day VIP Onboarding Concierge"]
        P2["Playbook 2: Fiber Care Shield Default Bundling"]
        P3["Playbook 3: Contract Horizon Migration Matrix"]
        P4["Playbook 4: Frictionless Autopay Adoption Incentive"]
    end

    subgraph Business Impact
        B1["Curb Day 1-90 Drop-off<br/>Protect $232K ARR"]
        B2["Compress Fiber Churn to <20%<br/>Protect $510K ARR"]
        B3["Migrate 25% to Annual Plans<br/>Protect $488K ARR"]
        B4["Eliminate Billing Friction<br/>Protect $145K ARR"]
    end

    I1 --> P1 --> B1
    I2 --> P2 --> B2
    I3 --> P3 --> B3
    I4 --> P4 --> B4
```

### Consolidated Financial ROI Model

| Strategic Program | Target Cohort | Annual Budget | Projected Churn Reduction | Retained Subscribers | Protected ARR | Net Annual Profit | Program ROI |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Playbook 1: Onboarding Concierge** | Months 1–6 Installs | \$65,000 | -35.0% | 248 accounts | \$232,000 | \$167,000 | **+257%** |
| **Playbook 2: Fiber Shield Bundling** | Unassisted Fiber | \$75,000 | -45.0% | 512 accounts | \$510,000 | \$435,000 | **+580%** |
| **Playbook 3: Contract Migration** | Month-to-Month | \$50,000 | -30.0% | 415 accounts | \$488,000 | \$438,000 | **+876%** |
| **Playbook 4: Autopay Modernization** | Electronic Check | \$30,000 | -20.0% | 162 accounts | \$145,000 | \$115,000 | **+383%** |
| **CONSOLIDATED PORTFOLIO** | **Enterprise Portfolio** | **\$220,000** | **-32.5% (Weighted)** | **1,337 accounts** | **\$1,375,000** | **\$1,155,000** | **+525% (Net ROI)** |

---

## 🚀 Quickstart & Reproducibility

### 1. Clone & Setup Environment
```bash
git clone https://github.com/Anmolbiswal183-dev/telco-customer-churn-data-science-capstone.git
cd telco-customer-churn-data-science-capstone
python -m venv venv
source venv/bin/activate  # On Windows: .\venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Execute Pipelines End-to-End
```bash
# Week 1: Data Cleaning & Forensic Audit
python week1_data_cleaning_eda/scripts/data_pipeline.py

# Week 2: 5-Act Visual Storytelling
python week2_visualization_storytelling/scripts/storytelling_visualizations.py

# Week 3: Statistical Hypothesis Testing
python week3_statistical_analysis/scripts/statistical_testing_pipeline.py

# Week 4: Machine Learning Benchmark & Cost Optimization
python week4_machine_learning/scripts/ml_modeling_pipeline.py

# Week 5: Capstone Master Synthesis Compilation
python week5_capstone_report/scripts/generate_capstone_doc.py
```

---

## 👥 Author & Acknowledgements
- **Author:** Anmol Biswal ([@Anmolbiswal183-dev](https://github.com/Anmolbiswal183-dev))
- **Dataset:** IBM Telco Customer Churn & Retention Ecosystem
- **License:** Distributed under the MIT License. See `LICENSE` for details.

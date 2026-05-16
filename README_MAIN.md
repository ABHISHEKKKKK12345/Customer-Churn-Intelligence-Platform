# Customer Churn Intelligence Platform
### MBA Capstone Project — Business Analytics
**Manipal Academy of Higher Education (MAHE) | Course Code: MBO649**

| Field | Detail |
|---|---|
| **Student** | Abhishek |
| **Registration No.** | 24154041051 |
| **Programme** | MBA — Business Analytics |
| **University** | Manipal Academy of Higher Education (MAHE) |
| **Month & Year** | May 2026 |
| **External Project Guide** | Ramendra Rai |
| **University Project Mentor** | Shrishma Rao V S |

---

## Table of Contents
1. [Project Overview](#1-project-overview)
2. [Repository and Submission Contents](#2-repository-and-submission-contents)
3. [Problem Statement](#3-problem-statement)
4. [System Architecture](#4-system-architecture)
5. [Dataset](#5-dataset)
6. [Methodology: CRISP-DM](#6-methodology-crisp-dm)
7. [Feature Engineering](#7-feature-engineering)
8. [Machine Learning Models](#8-machine-learning-models)
9. [Business Analytics Outputs](#9-business-analytics-outputs)
10. [Key Results](#10-key-results)
11. [Excel Report Structure](#11-excel-report-structure)
12. [How to Run the Code](#12-how-to-run-the-code)
13. [Dependencies](#13-dependencies)
14. [Project Findings Summary](#14-project-findings-summary)
15. [Limitations](#15-limitations)
16. [Future Scope](#16-future-scope)
17. [References](#17-references)

---

## 1. Project Overview

The **Customer Churn Intelligence Platform** is a Python-based, end-to-end analytics application that:

- Ingests raw customer datasets in multiple formats (CSV, Excel, ODS, TSV)
- Performs automated data cleaning and feature engineering
- Applies ensemble machine learning to predict individual churn probability
- Segments customers by risk level and customer lifetime value (CLV)
- Quantifies annualised revenue at risk
- Delivers a professionally formatted, multi-sheet Excel executive dashboard — requiring zero post-processing by the end user

**Core Innovation — Model Governance:** When no real churn column is found in the input data, the system generates synthetic churn labels via a percentile risk-scoring heuristic and automatically suppresses ML accuracy and AUC-ROC metrics to prevent misleading reporting. This reflects responsible analytics design.

---

## 2. Repository and Submission Contents

```
├── customer_churn_intelligence_system_Abhishek_24154041051.py   # Main platform code
├── churn.csv                                                     # Input dataset (100,000 records)
├── Churn_Analytics_Report.xlsx                                   # Output Excel report (generated)
├── Abhishek_24154041051_Final_Project_Report_Track1.pdf          # Final project report (PDF)
├── README.md                                                     # This file
```

---

## 3. Problem Statement

Subscription-based industries (telecom, SaaS, banking, OTT, insurance) face chronic revenue erosion from customer churn. Four compounding business challenges make this difficult to address:

| Challenge | Description |
|---|---|
| **Reactive Decision-Making** | Customers are identified as churned only after leaving — too late to intervene |
| **Lack of Financial Visibility** | Organisations cannot quantify revenue at risk from potential churn |
| **Limited Analytics Accessibility** | Advanced tools require data science expertise not available to business managers |
| **Insight-to-Action Gap** | Predictions exist in notebooks but never reach frontline retention teams |

**This platform directly addresses all four** through an automated, GUI-based pipeline that converts raw data into executive-ready retention intelligence.

---

## 4. System Architecture

```
Input Dataset          Data Cleaning &       Feature              Churn Prediction
(CSV / Excel)   --->   Preprocessing   --->  Engineering   --->   Model           --->   Risk Segmentation   --->   Excel Dashboard
                                                                                          & CLV Calculation          Report
```

**Processing flow (7 CRISP-DM phases):**

1. Business Understanding → KPI definition
2. Data Ingestion → Multi-format loader
3. Data Preparation → Cleaning & normalisation
4. Feature Engineering → 4 derived behavioural-financial features
5. Predictive Modelling → Random Forest → Gradient Boosting → Heuristic (tiered fallback)
6. Evaluation → Accuracy, AUC-ROC (suppressed if synthetic labels)
7. Reporting → 12-sheet Excel workbook with 14 KPIs and 6 charts

---

## 5. Dataset

**File:** `churn.csv`
**Records:** 100,000 customer records
**Variables:** 9 columns

| Column | Type | Description |
|---|---|---|
| `CustomerID` | Integer | Unique customer identifier |
| `Age` | Integer | Customer age in years (range: 18–80) |
| `Gender` | Categorical | Female / Male / Other |
| `Tenure` | Integer | Months as a customer (range: 1–72) |
| `MonthlyCharges` | Float | Monthly bill amount in Rs. (range: Rs.10–Rs.150) |
| `Contract` | Categorical | Month-to-month / One year / Two year |
| `PaymentMethod` | Categorical | Bank transfer / Credit card / Electronic check / Mailed check |
| `TotalCharges` | Float | Cumulative charges to date in Rs. (range: Rs.-118–Rs.10,831) |
| `Churn` | Categorical | Yes / No — actual churn label (real ground truth) |

**Dataset Quality:** After preprocessing — 0 missing values, 0 duplicate records. Dataset was clean and complete.

**Key Descriptive Statistics:**

| Variable | Min | Max | Mean | Key Observation |
|---|---|---|---|---|
| Age | 18 | 80 | 49.03 yrs | Churned: 48.88 vs Retained: 49.10 — negligible difference |
| Tenure | 1 | 72 months | 36.53 months | Churned avg: 30.89 — Retained avg: 39.32 |
| MonthlyCharges | Rs.10 | Rs.150 | Rs.79.97 | Churned avg: Rs.94.36 — Retained avg: Rs.72.85 |
| TotalCharges | Rs.-118 | Rs.10,831 | — | Minor negatives corrected by logical consistency check |

---

## 6. Methodology: CRISP-DM

| Phase | Implementation | Outcome |
|---|---|---|
| Phase 1 — Business Understanding | Define KPIs: Churn Rate, CLV, Revenue at Risk, Retention Rate | Alignment with revenue protection objectives |
| Phase 2 — Data Ingestion | Multi-format loader (CSV, XLSX, ODS, TSV) with 6 encoding fallbacks and 6 separator fallbacks | Universal compatibility across enterprise data exports |
| Phase 3 — Data Preparation | Missing value imputation, duplicate removal, currency normalisation, column standardisation | Clean, analysis-ready dataset from raw inputs |
| Phase 4 — Feature Engineering | 4 derived features: AvgMonthlySpend, ValueScore, LoyaltyScore, SpendVariance | Enriched feature set capturing behavioural and financial signals |
| Phase 5 — Modelling | Random Forest (primary), Gradient Boosting (fallback), Heuristic (emergency) | Layered reliability; graceful degradation on poor data |
| Phase 6 — Evaluation | Accuracy, AUC-ROC, Precision, Recall, F1; suppressed on synthetic labels | Responsible, non-misleading model reporting |
| Phase 7 — Reporting | 12-sheet Excel report with 14 KPIs, 6 charts, segmentation, audit trail | Executive-ready decision support output |

---

## 7. Feature Engineering

Four derived features are computed to capture behavioural and financial signals not directly observable in raw data:

| Feature | Formula | Business Purpose | Category |
|---|---|---|---|
| `AvgMonthlySpend` | `TotalCharges / (Tenure + 1)` | Spending consistency normalised by tenure | Behavioural |
| `ValueScore` | `log(1 + TotalCharges)` | Compressed value metric, handles skew | Financial |
| `LoyaltyScore` | `Tenure × TotalCharges / (MonthlyCharges + 1)` | Composite loyalty and engagement proxy | Behavioural |
| `SpendVariance` | `\|MonthlyCharges − AvgMonthlySpend\| ÷ (AvgMonthlySpend + 1)` | Detects spend volatility / instability | Risk |
| `Churn_Prob` | ML / heuristic output | Probability of churn (0 = safe, 1 = leaving) | ML Output |
| `Predicted_CLV` | `TotalCharges × (1 − Churn_Prob)` | Retention-adjusted value proxy | Financial |

**Smoothing:** All denominators include +1 offset to ensure numerical stability and prevent division-by-zero errors on real-world datasets.

**Code:**
```python
df['AvgMonthlySpend'] = (df['TotalCharges'] / (df['tenure'] + 1)).clip(lower=0)
df['ValueScore']      = np.log1p(df['TotalCharges']).clip(lower=0)
df['LoyaltyScore']    = (df['tenure'] * df['TotalCharges']
                         / (df['MonthlyCharges'] + 1)).clip(lower=0)
df['SpendVariance']   = ((df['MonthlyCharges'] - df['AvgMonthlySpend'])
                         / (df['AvgMonthlySpend'] + 1)).abs().clip(lower=0)
```

---

## 8. Machine Learning Models

### Three-Tier Modelling Architecture

#### Tier 1 — Random Forest (Primary Model)
```python
RandomForestClassifier(
    n_estimators=300,
    max_depth=12,
    class_weight='balanced',
    random_state=42,
    n_jobs=-1
)
```
- **300 decision trees** trained on bootstrap samples
- `class_weight='balanced'` addresses churn class imbalance automatically
- Train/test split: **80,000 / 20,000** (stratified, 80/20)

#### Tier 2 — Gradient Boosting (Secondary Fallback)
```python
GradientBoostingClassifier(
    n_estimators=150,
    max_depth=5,
    random_state=42
)
```
- Reached only if Random Forest raises an exception during training
- Sequential ensemble; 150 shallow trees correcting residual errors of predecessors

#### Tier 3 — Heuristic Fallback (Emergency)
```python
heuristic = (
    X['tenure'].rank(pct=True).rsub(1)        * 0.55  # lower tenure = higher risk
    + X['MonthlyCharges'].rank(pct=True)       * 0.25  # higher charges = higher risk
    + X['SpendVariance'].rank(pct=True)        * 0.20  # higher variance = higher risk
).clip(0, 1)
```
- Triggered when dataset has fewer than 20 rows or only one churn class
- Guarantees a valid churn score regardless of data quality

### Model Governance
The `_synthetic_churn` boolean flag automatically suppresses all accuracy and AUC-ROC metrics in both the GUI and Excel audit sheet when synthetic labels are used — preventing self-referential performance reporting.

### Feature Set (7 variables used in model)
`tenure`, `MonthlyCharges`, `TotalCharges`, `AvgMonthlySpend`, `ValueScore`, `LoyaltyScore`, `SpendVariance`

> **Note:** Demographic (Age, Gender) and categorical (Contract, PaymentMethod) variables are NOT direct model inputs. Their influence is captured indirectly through the financial and behavioural features.

---

## 9. Business Analytics Outputs

### Customer Lifetime Value (CLV)
```python
df['Predicted_CLV'] = df['TotalCharges'] * (1 - df['Churn_Prob'])
```
A retention-weighted proxy representing the proportion of a customer's historical total spend expected to be retained going forward. Explicitly documented as an approximation — not a DCF or survival-model estimate.

### Revenue at Risk
```python
revenue_at_risk = (df.loc[df['Churn_Prob'] >= 0.60, 'MonthlyCharges'] * 12).sum()
```
Annualised run-rate for all customers with predicted churn probability ≥ 60%.

### Risk Tier Classification

| Risk Tier | Threshold | Business Action |
|---|---|---|
| Low Risk | Churn_Prob < 0.30 | Stable customers; maintenance focus |
| Medium Risk | 0.30 ≤ Churn_Prob < 0.60 | Watch list; proactive engagement |
| High Risk | Churn_Prob ≥ 0.60 | Immediate intervention required |

### Customer Segmentation (2×2 Matrix)

| Segment | Churn Risk | CLV Level | Recommended Action |
|---|---|---|---|
| High Risk – High Value | ≥ 60% | ≥ Median CLV | Immediate Retention — Prioritize |
| High Risk – Low Value | ≥ 60% | < Median CLV | Monitor — Cost-Benefit Evaluation |
| Low Risk – High Value | < 60% | ≥ Median CLV | Upsell / Loyalty Programs |
| Low Risk – Low Value | < 60% | < Median CLV | Standard Service — Minimal Investment |

---

## 10. Key Results

Results from running the platform on `churn.csv` (100,000-record telecom dataset):

### Overall KPIs

| KPI | Value |
|---|---|
| Total Customers | 100,000 |
| Churn Rate | **33.14%** |
| Churned Customers | 33,144 |
| Retained Customers | 66,856 |
| Avg Tenure (Churned) | 30.89 months |
| Avg Tenure (Retained) | 39.32 months |
| Avg Monthly Charge (Churned) | Rs.94.36 |
| Avg Monthly Charge (Retained) | Rs.72.85 |
| Annualised Revenue at Risk | **Rs.54,285,942** |
| Historical TotalCharges at Risk | Rs.147,638,364 |
| High-Risk Customers (≥60%) | 41,374 (41.4%) |
| Total CLV Portfolio | Rs.152,352,263 |
| Avg Customer CLV | Rs.1,524 |
| ML Model Accuracy | **69.99%** |
| AUC-ROC Score | **0.7360** |

### Model Performance (Test Set n=20,000)

| Class | Precision | Recall | F1-Score | Support |
|---|---|---|---|---|
| Class 0 — Retained | 0.81 | 0.82 | 0.81 | 13,371 |
| Class 1 — Churned | 0.63 | 0.61 | 0.62 | 6,629 |
| Macro Average | 0.72 | 0.71 | 0.72 | 20,000 |
| Weighted Average | 0.75 | 0.75 | 0.75 | 20,000 |

### Churn Rate by Contract Type

| Contract | Churn Rate |
|---|---|
| Month-to-month | **46.56%** |
| One year | 16.75% |
| Two year | 16.88% |

### Churn Rate by Tenure Cohort

| Tenure Bucket | Churned | Retained | Churn Rate |
|---|---|---|---|
| 0–12 months | 10,655 | 5,992 | **64.0%** |
| 13–24 months | 4,471 | 12,166 | 26.9% |
| 25–36 months | 4,432 | 12,081 | 26.8% |
| 37–48 months | 4,520 | 12,299 | 26.9% |
| 49–60 months | 4,518 | 12,148 | 27.1% |
| 61–72 months | 4,548 | 12,170 | 27.2% |

### Segment Performance Summary

| Segment | Customers | Churned | Churn Rate | Avg CLV (Rs.) | Avg Tenure (mo) | Avg Monthly (Rs.) |
|---|---|---|---|---|---|---|
| High Risk – Low Value | 22,200 | 14,210 | 64.01% | 379.13 | 11.21 | 95.14 |
| High Risk – High Value | 19,174 | 9,622 | 50.18% | 2,292.05 | 50.36 | 125.78 |
| Low Risk – Low Value | 27,800 | 4,468 | 16.07% | 744.97 | 31.25 | 41.92 |
| Low Risk – High Value | 30,826 | 4,844 | 15.71% | 2,571.78 | 50.91 | 74.88 |

---

## 11. Excel Report Structure

The generated file `Churn_Analytics_Report.xlsx` contains 12 sheets:

| Sheet Name | Content |
|---|---|
| **Dashboard** | 14 KPI cards (2 rows × 7) + 6 embedded charts + ML metadata + synthetic label warning if applicable |
| **Segment Summary** | One row per segment — Churn Rate (RAG colour-coded), Avg CLV, Total CLV, Avg Tenure, Avg Monthly |
| **High Risk Customers** | Top 1,000 customers by predicted churn probability ≥ 60%, sorted highest risk first |
| **Processed Data** | Full dataset with all engineered features + Churn_Prob + Predicted_CLV + Segment + Risk_Tier |
| **Raw Data** | Original input dataset, untouched — for audit and source verification |
| **Data Quality Report** | Full pipeline audit: file details, column detection, model metrics, KPIs, methodology notes |
| **Data – Segments** | Chart source: customer count per segment |
| **Data – Risk Tiers** | Chart source: customer count per risk tier (Low: 51,198 / Med: 7,442 / High: 41,360) |
| **Data – Churn Prob** | Chart source: churn probability distribution (6 buckets) |
| **Data – CLV** | Chart source: CLV distribution (6 buckets) |
| **Data – Tenure** | Chart source: tenure distribution (6 buckets) |
| **Data – Monthly** | Chart source: monthly charges distribution (6 buckets) |

---

## 12. How to Run the Code

### Step 1 — Install dependencies
```bash
pip install pandas numpy scikit-learn openpyxl xlsxwriter xlrd odfpy
```

### Step 2 — Run the platform
```bash
python customer_churn_intelligence_system_Abhishek_24154041051.py
```

### Step 3 — Select input file
A GUI file dialog will open. Navigate to and select `churn.csv` (or any compatible customer dataset).

### Step 4 — Wait for processing
The platform will automatically:
- Detect and load the file
- Clean and preprocess the data
- Engineer features
- Train the Random Forest model
- Compute CLV, revenue at risk, and segmentation
- Generate the Excel report

### Step 5 — Retrieve output
The output file `Churn_Analytics_Report_<filename>.xlsx` will be saved in the same directory as the input file. A GUI popup will confirm completion with key metrics.

### Notes
- Works on **Windows, macOS, Linux**
- Requires Python 3.8+
- Minimum dataset size for ML: 20 rows. Below this, heuristic fallback activates automatically.
- If the input dataset has no `Churn` column, synthetic labels are generated and ML metrics are suppressed.

---

## 13. Dependencies

| Library | Version | Purpose |
|---|---|---|
| `pandas` | ≥ 1.5 | Data ingestion, manipulation, export |
| `numpy` | ≥ 1.23 | Numerical operations, array handling |
| `scikit-learn` | ≥ 1.1 | ML models (RandomForest, GradientBoosting), imputation, metrics |
| `openpyxl` | ≥ 3.0 | Excel file reading |
| `xlsxwriter` | ≥ 3.0 | Excel file writing with formatting and charts |
| `xlrd` | ≥ 2.0 | Legacy .xls file support |
| `odfpy` | ≥ 1.4 | ODS file format support |
| `tkinter` | (stdlib) | GUI file dialogs and popups |
| `os, re, sys, time, warnings, traceback` | (stdlib) | System utilities |

Install all at once:
```bash
pip install pandas numpy scikit-learn openpyxl xlsxwriter xlrd odfpy
```

---

## 14. Project Findings Summary

| # | Finding | Evidence |
|---|---|---|
| 1 | **Contract type shows the highest churn differential** | Month-to-month: 46.56% vs One year: 16.75% vs Two year: 16.88% |
| 2 | **First year is the make-or-break period** | 0–12 months: 64.0% churn vs 26.9% from month 13 onwards |
| 3 | **Churned customers pay 29.5% more per month** | Rs.94.36 (churned) vs Rs.72.85 (retained) — Rs.21.51 gap |
| 4 | **Demographic variables show minimal churn variation** | Age gap: 48.88 vs 49.10 — Gender: 32.85% vs 33.39% |
| 5 | **Payment method has minimal churn impact** | All methods cluster between 32.8% and 33.3% |
| 6 | **Model achieves commercially useful discrimination** | AUC-ROC 0.7360 — correctly ranks 73.6% of churned vs retained pairs |
| 7 | **Revenue exposure is concentrated and quantifiable** | Rs.54,285,942 annualised; Rs.147,638,364 historical TotalCharges at risk |
| 8 | **Governance mechanism works as designed** | Synthetic label suppression confirmed during development testing |
| 9 | **Model relies exclusively on financial/behavioural features** | 7 features; no demographic or categorical variables in model |

---

## 15. Limitations

| Limitation | Impact | Proposed Mitigation |
|---|---|---|
| Data Quality Dependency | Model performance depends on completeness and accuracy of input data | Implement data validation layer; enforce upstream data governance |
| CLV Approximation | CLV is a retention-weighted proxy, not a DCF or survival-model estimate | Augment with transaction-based CLV models in future iterations |
| Batch Processing Only | No real-time scoring pipeline | Deploy REST API layer using Flask or FastAPI |
| Synthetic Churn Fallback | Synthetic labels may not fully reflect real churn patterns | Mandate churn ground truth from business systems before deployment |
| Binary Churn Classification | Partial churn or product downgrade not captured | Extend to multi-class or survival analysis in next version |

---

## 16. Future Scope

| Extension | Description |
|---|---|
| **Real-Time Scoring Pipeline** | Deploy trained RF model as REST API (Flask/FastAPI) for CRM integration and point-of-interaction risk scoring |
| **Deep Learning Extension** | Apply RNNs / LSTMs to sequential transaction data to capture temporal churn patterns missed by cross-sectional models |
| **Survival Analysis Integration** | Replace binary classification with Cox Proportional Hazards model to predict *time to churn*, not just probability |
| **NLP-Enhanced Churn Signals** | Sentiment analysis on support tickets, chat logs, and call transcripts to add qualitative early-warning signals |
| **Cloud Deployment** | Package as AWS / Azure / GCP cloud service with web dashboard, multi-tenant isolation, and automated monthly reporting |

---

## 17. References

- Breiman, L. (2001). Random forests. *Machine Learning, 45*(1), 5–32. https://doi.org/10.1023/A:1010933404324
- Burez, J., & Van den Poel, D. (2009). Handling class imbalance in customer churn prediction. *Expert Systems with Applications, 36*(3), 4626–4636. https://doi.org/10.1016/j.eswa.2008.05.027
- Fader, P. S., Hardie, B. G., & Lee, K. L. (2005). RFM and CLV: Using iso-value curves for customer base analysis. *Journal of Marketing Research, 42*(4), 415–430. https://doi.org/10.1509/jmkr.2005.42.4.415
- Floridi, L., et al. (2019). An ethical framework for a good AI society. *Minds and Machines, 29*(4), 689–707. https://doi.org/10.1007/s11023-019-09497-0
- Friedman, J. H. (2001). Greedy function approximation: A gradient boosting machine. *The Annals of Statistics, 29*(5), 1189–1232. https://doi.org/10.1214/aos/1013203451
- Neslin, S. A., et al. (2006). Defection detection. *Journal of Marketing Research, 43*(2), 204–211. https://doi.org/10.1509/jmkr.43.2.204
- Power, D. J. (2002). *Decision support systems: Concepts and resources for managers.* Quorum Books.
- Reichheld, F. F., & Sasser, W. E. (1990). Zero defections: Quality comes to services. *Harvard Business Review, 68*(5), 105–111.
- Scikit-learn developers. (2024). scikit-learn: Machine learning in Python (Version 1.5). https://scikit-learn.org
- Verbeke, W., et al. (2012). New insights into churn prediction in the telecommunication sector. *European Journal of Operational Research, 218*(1), 211–229. https://doi.org/10.1016/j.ejor.2011.09.031

---

## Final Report Alignment Note

The following confirms full alignment across all 3 submitted artefacts:

| Check | Status |
|---|---|
| All KPI values in report match `churn.csv` computed values | ✅ |
| All KPI values in report match `Churn_Analytics_Report.xlsx` Dashboard | ✅ |
| All code excerpts in Annexure A match `.py` source file exactly | ✅ |
| All 12 Excel sheet names in Annexure C match output workbook | ✅ |
| All 5 Annexure B dataset rows match `churn.csv` rows 1–5 | ✅ |
| All segment figures match Excel `Segment Summary` sheet | ✅ |
| All 16 high-risk customer rows match Excel `High Risk Customers` sheet | ✅ |
| Feature engineering formulas in report match code exactly | ✅ |
| Model parameters (300 trees, max_depth=12, 80/20 split) match code | ✅ |
| Risk tier thresholds (0.30 / 0.60) match code | ✅ |
| CLV formula match code | ✅ |
| Revenue at Risk formula (≥0.60, ×12) match code | ✅ |

---

*README prepared to accompany the MBA Capstone Project submission — May 2026*
*Abhishek | Reg. No. 24154041051 | MBA — Business Analytics | MAHE*

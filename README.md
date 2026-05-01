# 🔄 Customer Churn Intelligence Platform

<div align="center">

**A Machine Learning-Based Decision Support System for Predicting Customer Attrition and Quantifying Revenue Risk in Subscription Businesses**

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![scikit-learn](https://img.shields.io/badge/scikit--learn-ML-F7931E?style=for-the-badge&logo=scikit-learn&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data-150458?style=for-the-badge&logo=pandas&logoColor=white)
![XlsxWriter](https://img.shields.io/badge/XlsxWriter-Reports-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-blueviolet?style=for-the-badge)

*An end-to-end customer churn analytics system that transforms raw business data into executive-grade intelligence — zero configuration required — works out-of-the-box with raw, unstructured business data.*

</div>

---

## 📋 Table of Contents

- [Overview](#overview)
- [Quick Start](#quick-start)
- [Key Features](#key-features)
- [How It Works](#how-it-works)
- [Installation](#installation)
- [Usage](#usage)
- [Supported File Formats](#supported-file-formats)
- [Column Auto-Detection](#column-auto-detection)
- [Machine Learning Pipeline](#machine-learning-pipeline)
- [Output Report Structure](#output-report-structure)
- [Feature Engineering](#feature-engineering)
- [Customer Segmentation](#customer-segmentation)
- [Methodology & Limitations](#methodology)
- [Build EXE](#build-exe)
- [Project Structure](#project-structure)
- [Dependencies](#dependencies)
- [Author](#author)
- [License](#license)

---

## 🎯 Overview

The **Customer Churn Intelligence Platform** is a fully automated, GUI-driven analytics tool built for business analysts and data professionals. It accepts raw customer data in virtually any tabular format, intelligently detects relevant columns, trains a machine learning model, and produces a polished, multi-sheet Excel report — all without writing a single line of code.

The platform is designed around a core business question:

> **"Which customers are about to leave, and how much revenue is at risk?"**

```
Raw Data File  →  Auto-Clean  →  ML Model  →  Segmentation  →  Executive Excel Report
  (any format)     + Repair       Training       + CLV Calc        (13 sheets + 6 charts)
```

---

## ⚡ Quick Start

```bash
# 1. Clone the repo
git clone https://github.com/your-username/churn-intelligence-platform.git
cd churn-intelligence-platform

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the tool
python customer_churn_intelligence_system.py
```

A GUI window will guide you through the rest — no configuration files, no command-line arguments.

---

## ✨ Key Features

| Feature | Description |
|---|---|
| 🗂️ **Universal File Ingestion** | Supports CSV, Excel (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`), ODS, and TSV — auto-tries 6 encodings and 6 delimiters |
| 🧹 **Intelligent Data Cleaning** | Removes duplicates, blank rows/columns, and unnamed Pandas artefacts; repairs currency symbols, US/EU/Indian locale numbers, and logical inconsistencies |
| 🔍 **Smart Column Detection** | Automatically identifies Churn, Revenue, Tenure, and Customer ID columns using keyword normalisation with conflict resolution |
| 🤖 **Dual ML Models** | Random Forest (primary, 300 trees) → Gradient Boosting (fallback, 150 estimators) → heuristic percentile score (last resort) |
| 📊 **Synthetic Label Fallback** | When no real churn column exists, labels are generated via a weighted percentile risk-score formula; ML metrics are suppressed and clearly flagged |
| 🧮 **CLV Prediction** | Retention-weighted historical proxy: `TotalCharges × (1 – Churn_Prob)` |
| 🎯 **2×2 Segmentation Matrix** | Classifies every customer into High/Low Risk × High/Low Value for action prioritisation |
| 📈 **Executive Excel Report** | 13-sheet workbook with 14 KPI cards, 6 embedded charts, colour-coded risk indicators, and a full pipeline audit trail |
| 🔒 **Safe File Saving** | Retry logic (up to 3 attempts) with automatic timestamp-based renaming when the target file is locked or already open |
| ⚠️ **Transparency Warnings** | Prominently flags synthetic labels throughout the report and suppresses misleading ML metrics for ethical, auditable reporting |

---

## 🔄 How It Works

```
┌─────────────────────────────────────────────────────────────────┐
│                        PIPELINE OVERVIEW                        │
├──────────────┬──────────────┬──────────────┬────────────────────┤
│   INGEST     │   PREPARE    │   ANALYSE    │      REPORT        │
├──────────────┼──────────────┼──────────────┼────────────────────┤
│ GUI file     │ Strip bad    │ Train Random │ Executive          │
│ picker       │ rows/cols    │ Forest or    │ Dashboard          │
│              │              │ Gradient     │                    │
│ 6 encodings  │ Detect col   │ Boosting     │ Segment Summary    │
│ tried        │ types        │              │                    │
│              │              │ Score every  │ High-Risk Action   │
│ 6 separators │ Currency /   │ customer     │ List               │
│ tried        │ locale fixes │ 0.0 → 1.0    │                    │
│              │              │              │ Processed Data     │
│ Excel, CSV,  │ Label encode │ Segment into │                    │
│ ODS, TSV     │ categoricals │ 4 groups     │ Raw Data           │
│              │              │              │                    │
│              │ Engineer     │ Compute CLV  │ Data Quality       │
│              │ features     │ & KPIs       │ Audit Report       │
└──────────────┴──────────────┴──────────────┴────────────────────┘
```

---

<a id="installation"></a>
## 🛠️ Installation

### Prerequisites

- Python **3.8 or higher**
- `pip` package manager
- Tkinter *(included with most Python distributions)*

### Step 1 — Clone or Download the Project

```bash
git clone https://github.com/your-username/churn-intelligence-platform.git
cd churn-intelligence-platform
```

### Step 2 — Install Dependencies

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install pandas numpy scikit-learn xlsxwriter openpyxl xlrd odfpy
```

### Step 3 — Run the Tool

```bash
python customer_churn_intelligence_system.py
```

> 💡 **Linux users:** If Tkinter is missing, install it via your system package manager:
> - Ubuntu/Debian: `sudo apt-get install python3-tk`
> - Fedora/RHEL: `sudo dnf install python3-tkinter`

---

## 🚀 Usage

The tool is entirely GUI-driven — no command-line arguments or configuration files required.

**Step 1 — Welcome Screen**
Click **OK** when the welcome dialog appears.

**Step 2 — Select Your Dataset**
A file picker opens. Navigate to your customer data file (see [Supported File Formats](#-supported-file-formats)).

**Step 3 — Automatic Processing**
The tool will:
- Load and clean your data
- Auto-detect all relevant columns
- Train the ML model (or fall back to heuristic scoring)
- Compute KPIs, CLV, segmentation, and revenue-at-risk

**Step 4 — Review the Summary**
A results dialog shows:
- Detected churn column (or synthetic fallback notification)
- Model accuracy and AUC-ROC *(suppressed if synthetic labels were used)*
- Key findings: churn rate, high-risk customer count, annualised revenue at risk

**Step 5 — Save the Report**
Choose a save location and filename. The fully formatted `.xlsx` report is generated immediately.

> 💡 **Tip:** If your Excel file is already open, the tool will offer to auto-save with a timestamped filename to avoid a `PermissionError`.

---

## 📁 Supported File Formats

| Format | Extensions | Notes |
|---|---|---|
| CSV | `.csv` | Auto-detects separator (`,` `;` `\t` `\|` `:`) |
| Excel (modern) | `.xlsx` `.xlsm` `.xlsb` | Uses `openpyxl` engine |
| Excel (legacy) | `.xls` | Uses `xlrd` engine |
| OpenDocument | `.ods` | Uses `odf` engine |
| Text / TSV | `.txt` `.tsv` | Same separator auto-detection as CSV |

**Encodings tried automatically:** `utf-8-sig`, `utf-8`, `latin1`, `cp1252`, `iso-8859-1`, `utf-16`

**Minimum requirement:** At least **1 data row** and **2 columns**.

---

## 🔍 Column Auto-Detection

The tool normalises all column names (lowercase, alphanumeric only) and searches for keywords in priority order. The **most specific keyword wins**, and no column can serve two roles simultaneously.

| Role | Keywords Searched (priority order) |
|---|---|
| **Churn** | `churn`, `attrition`, `churned`, `leave`, `exit`, `cancel`, `left`, `status` |
| **Tenure** | `tenure`, `seniority`, `duration`, `period`, `months`, `age` |
| **Monthly Revenue** | `monthlycharge`, `monthlyfee`, `monthly`, `charge`, `rate`, `fee`, `price`, `subscription` |
| **Total Revenue** | `totalcharge`, `totalrevenue`, `lifetimevalue`, `total`, `revenue`, `bill`, `spend`, `amount`, `sales`, `ltv` |
| **Customer ID** | `customerid`, `custid`, `userid`, `accountid`, `clientid`, `memberid`, `id` |

**Fallback behaviour:** If a column is not found, the tool falls back to numeric columns by index position, or uses sensible defaults (e.g., median tenure = 12 months, median monthly charge = Rs. 500).

### Churn Label Mapping

String values in the churn column are mapped to binary labels automatically:

| Churned → `1` | Retained → `0` |
|---|---|
| `Yes`, `Y`, `True`, `1`, `Churned`, `Left`, `Cancelled`, `Cancel`, `Exit`, `Quit`, `Inactive`, `Lost`, `Gone`, `Departed`, `Attrited`, `Closed` | `No`, `N`, `False`, `0`, `Active`, `Retained`, `Stay`, `Stayed`, `Current`, `Existing`, `Ongoing`, `Present`, `Alive`, `Good`, `Loyal` |

> ⚠️ If more than **60%** of values cannot be mapped to either side, **synthetic labels** are generated instead (see [Methodology & Limitations](#-methodology--limitations)).

---

## 🤖 Machine Learning Pipeline

### Models

The tool attempts models in this priority order, stopping at the first that succeeds:

| Priority | Model | Configuration |
|---|---|---|
| 1 | **Random Forest** | 300 trees, max depth 12, balanced class weights, `random_state=42`, `n_jobs=-1` |
| 2 | **Gradient Boosting** | 150 estimators, max depth 5, `random_state=42` |
| 3 | **Heuristic Score** | Percentile rank formula — no training required |

### Feature Set

All 7 features below are used for model training:

| Feature | Formula / Source | Description |
|---|---|---|
| `tenure` | Raw column | Months the customer has been with the company |
| `MonthlyCharges` | Raw column | Current monthly billing amount |
| `TotalCharges` | Raw column | Cumulative historical charges |
| `AvgMonthlySpend` | `TotalCharges / (tenure + 1)` | Normalised spend over customer lifetime |
| `ValueScore` | `log(1 + TotalCharges)` | Log-transformed total value — reduces right skew |
| `LoyaltyScore` | `tenure × (TotalCharges / (MonthlyCharges + 1))` | Combined tenure-spend loyalty indicator |
| `SpendVariance` | `\|MonthlyCharges – AvgMonthlySpend\| / (AvgMonthlySpend + 1)` | Billing instability signal — flags volatile payers |

### Training Conditions

ML training is attempted only when:
- Dataset has **≥ 20 rows** (constant `MIN_ROWS_FOR_ML`)
- Churn column has **≥ 2 unique values** (not all churned or all retained)
- An 80/20 stratified train/test split is applied when both classes have sufficient samples; falls back to a non-stratified split otherwise

### Missing Value Imputation

All features are imputed with **median strategy** (`sklearn.impute.SimpleImputer`) before training. Categorical columns with 2–20 unique values are label-encoded; columns with more unique values are dropped from the numeric feature set.

### Metrics Transparency

> ⚠️ When churn labels are **synthetic**, model accuracy and AUC-ROC are **intentionally suppressed**. Reporting these metrics on synthetic labels would only measure how well the model reproduced the tool's own risk-score formula — not real customer behaviour. This is clearly flagged with a warning banner on the Dashboard and in the Data Quality Report.

---

## 📊 Output Report Structure

The generated `.xlsx` file always contains **13 sheets**:

### Sheet 1 — Dashboard

- **Title banner** — file name, model used, accuracy, AUC-ROC, and customer count
- **Synthetic labels warning banner** *(shown only when applicable)*
- **14 KPI Cards** across 2 rows:

  | Row | KPIs |
  |---|---|
  | Row 1 | Total Customers · Churn Rate · Retention Rate · High Risk (≥60%) · Medium Risk (30–60%) · Low Risk (<30%) · Avg Tenure |
  | Row 2 | Avg Customer CLV · Annualised Risk Revenue · Total CLV Portfolio · Avg Monthly Charge · ML Model · Model Accuracy · AUC-ROC Score |

- **6 Embedded Charts:**

  | # | Chart Title | Type |
  |---|---|---|
  | 1 | Customer Segments by Count | Column |
  | 2 | Risk Tier Breakdown | Doughnut |
  | 3 | Churn Probability Distribution | Column |
  | 4 | Predicted CLV Distribution | Bar |
  | 5 | Customer Tenure Distribution | Column |
  | 6 | Monthly Charges Distribution | Column |

### Sheet 2 — Segment Summary

One row per customer segment with colour-coded churn rate:

- 🔴 **RED** — Churn Rate > 60%
- 🟡 **AMBER** — Churn Rate 30–60%
- 🟢 **GREEN** — Churn Rate < 30%

Columns: Segment · Customers · Churned · Churn Rate · Avg CLV · Total CLV · Avg Tenure · Avg Monthly Charge · Avg Risk Score

### Sheet 3 — High Risk Customers

Top **1,000 customers** with `Churn_Prob ≥ 60%`, sorted highest-risk first.

Columns: Customer ID *(if detected)* · Tenure · Monthly Charges · Total Charges · Predicted CLV · Churn Probability · Segment · Risk Tier

Auto-filter and frozen header row enabled.

### Sheet 4 — Processed Data

Full dataset with all engineered features and ML outputs. Per-column numeric formatting applied (decimals, currency, percentages). Auto-filter and frozen header row enabled.

### Sheet 5 — Raw Data

Original file content, completely untouched — preserved for audit and data lineage.

### Sheet 6 — Data Quality Report

Full pipeline audit trail including:
- File information (name, path, original vs cleaned row counts, rows removed)
- Column detection results and all fallback decisions
- ML model details, features used, and metrics (or suppression reason)
- Methodology notes and formula explanations
- All output KPIs with their calculation basis

### Sheets 7–13 — Chart Source Data

`Data – Segments` · `Data – Risk Tiers` · `Data – Churn Prob` · `Data – CLV` · `Data – Tenure` · `Data – Monthly`

These sheets directly feed the dashboard charts. They are accessible for reference but are not the primary deliverable.

---

## 🧮 Feature Engineering

| Engineered Column | Formula | Business Meaning |
|---|---|---|
| `AvgMonthlySpend` | `TotalCharges / (tenure + 1)` | Normalised monthly spend over customer lifetime |
| `ValueScore` | `log(1 + TotalCharges)` | Log-transformed total value — reduces right skew |
| `LoyaltyScore` | `tenure × TotalCharges / (MonthlyCharges + 1)` | Combined tenure-spend loyalty indicator |
| `SpendVariance` | `\|Monthly – AvgMonthlySpend\| / (AvgMonthlySpend + 1)` | Billing instability — flags volatile payers |
| `Churn_Prob` | ML model output (or heuristic rank) | Predicted probability of churning — range `0.0` to `1.0` |
| `Predicted_CLV` | `TotalCharges × (1 – Churn_Prob)` | Retention-weighted historical value proxy |

**Data integrity fix:** Where `TotalCharges < MonthlyCharges` (logically impossible for any active customer), TotalCharges is automatically recomputed as `tenure × MonthlyCharges`.

---

## 🎯 Customer Segmentation

Every customer is placed into one of four segments using a **2×2 Risk × Value matrix:**

|  | **HIGH VALUE** (`Predicted_CLV ≥ median`) | **LOW VALUE** (`Predicted_CLV < median`) |
|---|---|---|
| **HIGH RISK** (`Churn_Prob ≥ 0.60`) | 🔴 High Risk – High Value · **PRIORITY 1** | 🟠 High Risk – Low Value · **PRIORITY 2** |
| **LOW RISK** (`Churn_Prob < 0.60`) | 🟢 Low Risk – High Value · **NURTURE** | ⚪ Low Risk – Low Value · **MONITOR** |

**Risk Tiers** (also stored as `Risk_Tier` column):

| Tier | Churn Probability Range |
|---|---|
| Low Risk | 0% – 30% |
| Medium Risk | 30% – 60% |
| High Risk | 60% – 100% |

---

<a id="methodology"></a>
## ⚠️ Methodology & Limitations

### CLV Calculation

`Predicted_CLV` is a **retention-weighted historical proxy**, not a true Customer Lifetime Value:

```
Predicted_CLV = TotalCharges × (1 – Churn_Prob)
```

This represents the portion of historical spend expected to be retained. It is **not** a survival model, discounted cash flow (DCF) CLV, or forward-looking revenue projection. Use for **relative ranking purposes only.**

### Revenue at Risk

Annualised revenue at risk is the forward-looking run-rate for high-risk customers:

```
Risk Revenue = Σ (MonthlyCharges × 12)  for all customers where Churn_Prob ≥ 0.60
```

The historical `TotalCharges` for the same group is also separately reported in the Data Quality sheet.

### Synthetic Label Generation

When no real churn column is detected (or >60% of values cannot be mapped), labels are generated using:

```
Risk Score = tenure_rank (ascending)↓          × 0.40
           + monthly_charges_rank (ascending)↑  × 0.30
           + spend_variance_rank (ascending)↑   × 0.20
           + total_charges_rank (ascending)↓    × 0.10

Top 35% by Risk Score  →  Churned (1)
Remaining 65%          →  Retained (0)
```

ML accuracy and AUC-ROC are **suppressed** whenever synthetic labels are in use.

### Heuristic Score (No-ML Fallback)

When ML cannot run (dataset too small, single-class labels), churn probability is estimated as:

```
Heuristic Score = tenure_rank↓ × 0.55
               + monthly_charges_rank↑ × 0.25
               + spend_variance_rank↑ × 0.20
```

---

<a id="build-exe"></a>
## 🖥️ Build EXE

Convert the tool into a standalone Windows `.exe` that requires no Python installation on the target machine.

### Prerequisites

- All dependencies from `requirements.txt` must be installed
- A clean virtual environment is strongly recommended
- Build on the **same OS** as the target machine (Windows → Windows)

### Step 1 — Install PyInstaller

```bash
pip install pyinstaller
```

### Step 2 — Build the EXE

```bash
pyinstaller --onefile --windowed --noconsole \
  --name "Customer_Churn_Intelligence_Platform" \
  --collect-all sklearn \
  --collect-all openpyxl \
  --collect-all xlsxwriter \
  --collect-all odf \
  --collect-all xlrd \
  customer_churn_intelligence_system.py
```

### Step 3 — Find Your EXE

```
dist/Customer_Churn_Intelligence_Platform.exe
```

### Troubleshooting

| Issue | Fix |
|---|---|
| EXE not opening or crashing | Ensure all dependencies were installed before building; rebuild in a clean venv |
| Missing module errors | Add the missing package with `--collect-all <package>` |
| GUI not showing | Ensure `--windowed` flag is present |
| Antivirus blocking EXE | Add an exception, or rebuild with a different `--name` value |

---

## 📂 Project Structure

```
churn-intelligence-platform/
│
├── customer_churn_intelligence_system.py   # Main application — single self-contained file
├── requirements.txt                        # All Python dependencies
├── README.md                               # This documentation
└── .gitignore                              # Excludes build/, dist/, __pycache__/, *.spec
```

> 💡 `dist/`, `build/`, and `.spec` files are excluded via `.gitignore`. To generate the `.exe`, follow the [Build EXE](#-build-exe) section above.

---

## 📦 Dependencies

```bash
pip install -r requirements.txt
```

| Package | Purpose | Min Version |
|---|---|---|
| `pandas` | Data loading, cleaning, and manipulation | ≥ 1.3.0 |
| `numpy` | Numerical operations and array handling | ≥ 1.21.0 |
| `scikit-learn` | ML models, imputation, and label encoding | ≥ 0.24.0 |
| `xlsxwriter` | Excel report generation with charts and formatting | ≥ 3.0.0 |
| `openpyxl` | Reading `.xlsx`, `.xlsm`, `.xlsb` files | ≥ 3.0.0 |
| `xlrd` | Reading legacy `.xls` files | ≥ 2.0.0 |
| `odfpy` | Reading `.ods` OpenDocument spreadsheets | ≥ 1.4.0 |
| `tkinter` | GUI file dialogs and message boxes | Built-in (stdlib) |

---

<a id="author"></a>
## 👨‍💼 Author

<div align="center">

| Field | Detail |
|---|---|
| **Name** | [Abhishek](https://www.linkedin.com/in/abhishek-srivastava-1538461b1/) |
| **Project** | Customer Churn Intelligence Platform |

</div>

---

## 📄 License

This project is released under the [MIT License](LICENSE). You are free to use, modify, and distribute it with attribution.

---

<div align="center">

*Built with ❤️ for data professionals who need answers, not just data.*

</div>

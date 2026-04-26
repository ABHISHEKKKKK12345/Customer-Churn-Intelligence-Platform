# 🔄 Customer-Churn-Intelligence-Platform

<div align="center">
  
A Machine Learning-Based Decision Support System for Predicting Customer Attrition and Quantifying Revenue Risk in Subscription Businesses


![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![scikit-learn](https://img.shields.io/badge/scikit--learn-ML-F7931E?style=for-the-badge&logo=scikit-learn&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data-150458?style=for-the-badge&logo=pandas&logoColor=white)
![XlsxWriter](https://img.shields.io/badge/XlsxWriter-Reports-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![License](https://img.shields.io/badge/License-Academic-blueviolet?style=for-the-badge)

**An end-to-end customer churn analytics system that transforms raw business data into executive-grade intelligence — with zero configuration required.**


</div>

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Key Features](#-key-features)
- [How It Works](#-how-it-works)
- [Installation](#-installation)
- [Usage](#-usage)
- [Supported File Formats](#-supported-file-formats)
- [Column Auto-Detection](#-column-auto-detection)
- [Machine Learning Pipeline](#-machine-learning-pipeline)
- [Output Report Structure](#-output-report-structure)
- [Feature Engineering](#-feature-engineering)
- [Customer Segmentation](#-customer-segmentation)
- [Methodology & Limitations](#-methodology--limitations)
- [Project Structure](#-project-structure)
- [Dependencies](#-dependencies)
- [Author](#-author)

---

## 🎯 Overview

The **Customer Churn Intelligence Platform** is a fully automated, GUI-driven analytics tool built for business analysts and MBA practitioners. It accepts raw customer data in virtually any tabular format, intelligently detects relevant columns, trains a machine learning model, and produces a polished, multi-sheet Excel report — all without writing a single line of code.

The platform is designed around a core business question: **"Which customers are about to leave, and how much revenue is at risk?"**

```
Raw Data File  →  Auto-Clean  →  ML Model  →  Segmentation  →  Executive Excel Report
  (any format)     + Repair       Training       + CLV Calc        (6 sheets + 6 charts)
```

---

| Feature                          | Description                                                                                                                    |
| -------------------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| 🗂️ **Universal File Ingestion** | Supports CSV, Excel (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`), ODS, TSV with automatic handling of multiple encodings and delimiters |
| 🧹 **Intelligent Data Cleaning** | Cleans currency symbols, mixed number formats (US/EU/India), duplicates, missing values, and logical inconsistencies           |
| 🔍 **Smart Column Detection**    | Automatically detects Churn, Revenue, Tenure, and Customer ID using keyword normalization                                      |
| 🤖 **Dual ML Models**            | Uses Random Forest (primary) and Gradient Boosting (fallback), with heuristic scoring for small/invalid datasets               |
| 📊 **Synthetic Label Fallback**  | Generates churn labels using a percentile-based risk scoring method when no churn column exists                                |
| 🧮 **CLV Prediction**            | Calculates Customer Lifetime Value using: `TotalCharges × (1 – Churn_Prob)`                                                    |
| 🎯 **2×2 Segmentation Matrix**   | Classifies customers into High/Low Risk × High/Low Value segments for action prioritization                                    |
| 📈 **Executive Excel Report**    | Generates a multi-sheet Excel report with dashboards, KPIs, charts, and actionable insights                                    |
| 🔒 **Safe File Saving**          | Implements retry logic with auto-renaming (timestamp-based) for locked or open Excel files                                     |
| ⚠️ **Transparency Warnings**     | Flags synthetic data usage and suppresses misleading ML metrics for ethical reporting                                          |


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

## 🛠️ Installation

### Prerequisites

- Python **3.8 or higher**
- `pip` package manager
- Tkinter (included with most Python distributions)

### Step 1 — Clone or download the project

```bash
git clone https://github.com/your-username/churn-intelligence-platform.git
cd churn-intelligence-platform
```

### Step 2 — Install dependencies

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install pandas numpy scikit-learn xlsxwriter openpyxl xlrd odfpy
```

### Step 3 — Run the tool

```bash
python churn_analytics.py
```

A GUI window will guide you through the rest.

---

## 🚀 Usage

The tool is entirely GUI-driven — no command-line arguments or configuration files required.

**Step 1 — Welcome Screen**

Click **OK** when the welcome dialog appears.

**Step 2 — Select Your Dataset**

A file picker opens. Navigate to your customer data file (see [Supported Formats](#-supported-file-formats)).

**Step 3 — Automatic Processing**

The tool will:
- Load and clean your data
- Auto-detect relevant columns
- Train the ML model
- Compute all KPIs and segments

**Step 4 — Review the Summary**

A results dialog shows:
- Detected churn column (or synthetic fallback notification)
- Model accuracy and AUC-ROC (suppressed if synthetic labels)
- Key findings: churn rate, high-risk count, revenue at risk

**Step 5 — Save the Report**

Choose a save location and filename. The `.xlsx` report is generated immediately.

> 💡 **Tip:** If your Excel file is already open, the tool will offer to auto-save with a timestamped filename to avoid a `PermissionError`.

---

## 📁 Supported File Formats

| Format         | Extensions              | Notes                                          |
| -------------- | ----------------------- | ---------------------------------------------- |
| CSV            | `.csv`                  | Auto-detects separator (`,` `;` `\t` `\|` `:`) |
| Excel (modern) | `.xlsx` `.xlsm` `.xlsb` | Uses `openpyxl` engine                         |
| Excel (legacy) | `.xls`                  | Uses `xlrd` engine                             |
| OpenDocument   | `.ods`                  | Uses `odf` engine                              |
| Text / TSV     | `.txt` `.tsv`           | Same separator auto-detection as CSV           |


**Encodings tried automatically:** `utf-8-sig`, `utf-8`, `latin1`, `cp1252`, `iso-8859-1`, `utf-16`

**Minimum requirement:** At least **1 data row** and **2 columns**.

---

## 🔍 Column Auto-Detection

The tool normalises all column names (lowercase, alphanumeric only) and searches for keywords in priority order. The **most specific keyword wins**.

| Role                | Keywords Searched (in priority order)                                                                         |
| ------------------- | ------------------------------------------------------------------------------------------------------------- |
| **Churn**           | `churn`, `attrition`, `churned`, `leave`, `exit`, `cancel`, `left`, `status`                                  |
| **Tenure**          | `tenure`, `seniority`, `duration`, `period`, `months`, `age`                                                  |
| **Monthly Revenue** | `monthlycharge`, `monthlyfee`, `monthly`, `charge`, `rate`, `fee`, `price`, `subscription`                    |
| **Total Revenue**   | `totalcharge`, `totalrevenue`, `lifetimevalue`, `total`, `revenue`, `bill`, `spend`, `amount`, `sales`, `ltv` |
| **Customer ID**     | `customerid`, `custid`, `userid`, `accountid`, `clientid`, `memberid`, `id`                                   |


**Conflict resolution:** No column can serve two roles. If a match has already been claimed, the next-best match is used.

**Fallback:** If a column is not found, the tool falls back to numeric columns by index position, or uses sensible defaults (e.g., median tenure = 12 months).

### Churn Label Mapping

String values are mapped to binary labels automatically:

| → **Churned (1)** | → **Retained (0)** |
|------------------|--------------------|
| Yes<br>Y<br>True<br>1<br>Churned<br>Left<br>Cancelled<br>Exit<br>Quit<br>Inactive<br>Lost<br>Gone<br>Departed<br>Attrited<br>Closed | No<br>N<br>False<br>0<br>Active<br>Retained<br>Stay<br>Stayed<br>Current<br>Existing<br>Ongoing<br>Present<br>Alive<br>Good<br>Loyal |


If >60% of values cannot be mapped, **synthetic labels** are generated instead.

---

## 🤖 Machine Learning Pipeline

### Models

| Priority    | Model                 | Configuration                                   |
| ----------- | --------------------- | ----------------------------------------------- |
| Primary     | **Random Forest**     | 300 trees, max depth 12, balanced class weights |
| Fallback    | **Gradient Boosting** | 150 estimators, max depth 5                     |
| Last resort | **Heuristic Score**   | Percentile rank formula (no training required)  |


### Feature Set

All 7 features are used for training:

| Feature           | Description                                                    |
| ----------------- | -------------------------------------------------------------- |
| `tenure`          | Months with the company                                        |
| `MonthlyCharges`  | Current monthly billing amount                                 |
| `TotalCharges`    | Cumulative historical charges                                  |
| `AvgMonthlySpend` | `TotalCharges / (tenure + 1)`                                  |
| `ValueScore`      | `log(1 + TotalCharges)` — normalises skew                      |
| `LoyaltyScore`    | `tenure × (TotalCharges / (MonthlyCharges + 1))`               |
| `SpendVariance`   | `\|MonthlyCharges – AvgMonthlySpend\| / (AvgMonthlySpend + 1)` |


### Training Conditions

ML training is attempted only when:
- Dataset has **≥ 20 rows**
- Churn column has **≥ 2 unique values** (not all churned / all retained)
- Stratified split is used when both classes have sufficient samples

### Metrics Transparency

> ⚠️ When churn labels are **synthetic**, accuracy and AUC-ROC are **intentionally suppressed**. Reporting metrics on synthetic labels would only measure how well the model learned the tool's own formula — not real customer behaviour. This is clearly flagged throughout the report.

---

## 📊 Output Report Structure

The generated `.xlsx` report contains **12–13 sheets**:

### Sheet 1 — Dashboard
- **Title banner** with file info, model name, accuracy, and AUC-ROC
- **14 KPI Cards** across 2 rows:
  - Row 1: Total Customers, Churn Rate, Retention Rate, High/Medium/Low Risk counts, Avg Tenure
  - Row 2: Avg CLV, Annualised Risk Revenue, Total CLV Portfolio, Avg Monthly Charge, Model Name, Accuracy, AUC-ROC
- **6 Embedded Charts:**
  1. Customer Segments by Count (column)
  2. Risk Tier Breakdown (doughnut)
  3. Churn Probability Distribution (column)
  4. Predicted CLV Distribution (bar)
  5. Customer Tenure Distribution (column)
  6. Monthly Charges Distribution (column)
- **Synthetic labels warning banner** (shown only when applicable)

### Sheet 2 — Segment Summary
- One row per segment with colour-coded churn rate:
  - 🔴 RED: Churn Rate > 60%
  - 🟡 AMBER: Churn Rate 30–60%
  - 🟢 GREEN: Churn Rate < 30%
- Metrics: Customers, Churned, Churn Rate, Avg/Total CLV, Avg Tenure, Avg Monthly Charge, Avg Risk Score

### Sheet 3 — High Risk Customers
- Top **1,000 customers** with Churn Probability ≥ 60%
- Sorted highest-risk first
- Columns: Customer ID (if detected), Tenure, Monthly Charges, Total Charges, Predicted CLV, Churn Probability, Segment, Risk Tier
- Auto-filter and frozen header row enabled

### Sheet 4 — Processed Data
- Full dataset with all engineered features
- Formatted with numeric precision per column type
- Auto-filter and frozen header row

### Sheet 5 — Raw Data
- Original file, completely untouched
- Useful for audit and data lineage

### Sheet 6 — Data Quality Report
- Full pipeline audit trail:
  - File information (name, path, row counts)
  - Column detection results
  - ML model details and metrics
  - Methodology explanations
  - All KPIs with calculation basis

### Sheets 7–13 — Chart Source Data (hidden)
- `Data – Segments`, `Data – Risk Tiers`, `Data – Churn Prob`
- `Data – CLV`, `Data – Tenure`, `Data – Monthly`
- These feed the dashboard charts; they are accessible but not the primary focus

---

## 🧮 Feature Engineering

| Engineered Column | Formula                                        | Business Meaning                                |
| ----------------- | ---------------------------------------------- | ----------------------------------------------- |
| `AvgMonthlySpend` | `TotalCharges / (tenure + 1)`                  | Normalised monthly spend over customer lifetime |
| `ValueScore`      | `log(1 + TotalCharges)`                        | Log-transformed total value (reduces skew)      |
| `LoyaltyScore`    | `tenure × TotalCharges / (MonthlyCharges + 1)` | Combined tenure-spend loyalty indicator         |
| `SpendVariance`   | `\|Monthly – Avg\| / (Avg + 1)`                | Billing instability signal                      |
| `Churn_Prob`      | ML model output (or heuristic)                 | Predicted probability of churning (0.0 – 1.0)   |
| `Predicted_CLV`   | `TotalCharges × (1 – Churn_Prob)`              | Retention-weighted value proxy                  |


**Data integrity fix:** Where `TotalCharges < MonthlyCharges` (logically impossible), TotalCharges is recomputed as `tenure × MonthlyCharges`.

---

## 🎯 Customer Segmentation

Customers are placed into one of four segments using a **2×2 Risk × Value matrix**:

```
                    HIGH VALUE          LOW VALUE
                  ┌─────────────────┬──────────────────┐
  HIGH RISK       │  High Risk –    │  High Risk –     │
  (Churn ≥ 60%)   │  High Value     │  Low Value       │
                  │  🔴 PRIORITY 1 │  🟠 PRIORITY 2   │
                  ├─────────────────┼──────────────────┤
  LOW RISK        │  Low Risk –     │  Low Risk –      │
  (Churn < 60%)   │  High Value     │  Low Value       │
                  │  🟢 NURTURE     │  ⚪ MONITOR     │
                  └─────────────────┴──────────────────┘
```

- **High Value** = `Predicted_CLV ≥ median CLV`
- **Risk Tiers:** Low (0–30%), Medium (30–60%), High (60–100%)

---

## ⚠️ Methodology & Limitations

### CLV Calculation
The `Predicted_CLV` in this tool is a **retention-weighted historical proxy**, not a true Customer Lifetime Value:

```
Predicted_CLV = TotalCharges × (1 – Churn_Prob)
```

This represents the portion of historical spend expected to be retained. It is **not** a survival model, discounted cash flow (DCF) model, or forward-looking revenue projection. Use for **relative ranking only**.

### Revenue at Risk
Annualised revenue at risk is computed as:
```
Risk Revenue = Σ (MonthlyCharges × 12) for all customers with Churn_Prob ≥ 60%
```
This is a **forward-looking proxy**. Historical `TotalCharges` for the same group is also reported separately.

### Synthetic Label Generation
When no real churn column is detected, labels are generated using:
```
Risk Score = tenure_rank↓ × 0.40
           + monthly_charges_rank↑ × 0.30
           + spend_variance_rank↑ × 0.20
           + total_charges_rank↓ × 0.10

Top 35% by Risk Score → Labelled as Churned (1)
```
ML metrics are suppressed when synthetic labels are in use.

---

## 📂 Project Structure

```
churn-intelligence-platform/
│
├── churn_analytics.py          # Main application (single-file)
├── requirements.txt            # Python dependencies
├── README.md                   # This file
│
└── sample_data/                # (Optional) Test datasets
    ├── telecom_churn.csv
    └── retail_customers.xlsx
```

---

## 📦 Dependencies

| Package        | Purpose                              | Version  |
| -------------- | ------------------------------------ | -------- |
| `pandas`       | Data loading, cleaning, manipulation | ≥ 1.3    |
| `numpy`        | Numerical operations                 | ≥ 1.21   |
| `scikit-learn` | ML models, imputation, encoding      | ≥ 0.24   |
| `xlsxwriter`   | Excel report generation              | ≥ 3.0    |
| `openpyxl`     | Excel file reading                   | ≥ 3.0    |
| `xlrd`         | Legacy `.xls` reading                | ≥ 2.0    |
| `odfpy`        | ODS file reading                     | ≥ 1.4    |
| `tkinter`      | GUI dialogs (stdlib)                 | Built-in |


**`requirements.txt`:**
```
pandas>=1.3.0
numpy>=1.21.0
scikit-learn>=0.24.0
xlsxwriter>=3.0.0
openpyxl>=3.0.0
xlrd>=2.0.0
odfpy>=1.4.0
```

---

## 👨‍💼 Author

<div align="center">

|                 |                                            |
| --------------- | ------------------------------------------ |
| **Name**        | Abhishek                                   |
| **Program**     | MBA in Business Analytics                  |
| **Institution** | Manipal Academy of Higher Education (MAHE) |
| **Roll Number** | 24154041051                                |
| **Project**     | Customer Churn Intelligence Platform       |
| **Version**     | v3.0 (Final)                               |


</div>

---

## 📄 License

This project was developed as an academic final-year MBA project. It is intended for educational and non-commercial use.

---

<div align="center">

*Built with ❤️ for business analytics practitioners who need answers, not just data.*

</div>


# 🔄 Customer Churn Intelligence Platform

<div align="center">

**A Machine Learning-Based Decision Support System for Predicting Customer Attrition and Quantifying Revenue Risk in Subscription Businesses**

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![scikit-learn](https://img.shields.io/badge/scikit--learn-ML-F7931E?style=for-the-badge&logo=scikit-learn&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data-150458?style=for-the-badge&logo=pandas&logoColor=white)
![XlsxWriter](https://img.shields.io/badge/XlsxWriter-Reports-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![License](https://img.shields.io/badge/License-Academic-blueviolet?style=for-the-badge)

*An end-to-end customer churn analytics system that transforms raw business data into executive-grade intelligence — with zero configuration required.*

*MBA Final Project · Manipal Academy of Higher Education (MAHE)*

</div>

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Key Features](#-key-features)
- [How It Works](#-how-it-works)
- [Installation](#-installation)
- [Usage](#-usage)
- [Supported File Formats](#-supported-file-formats)
- [Column Auto-Detection](#-column-auto-detection)
- [Machine Learning Pipeline](#-machine-learning-pipeline)
- [Output Report Structure](#-output-report-structure)
- [Feature Engineering](#-feature-engineering)
- [Customer Segmentation](#-customer-segmentation)
- [Methodology & Limitations](#-methodology--limitations)
- [Project Structure](#-project-structure)
- [Dependencies](#-dependencies)
- [Author](#-author)

---

## 🎯 Overview

The **Customer Churn Intelligence Platform** is a fully automated, GUI-driven analytics tool built for business analysts and MBA practitioners. It accepts raw customer data in virtually any tabular format, intelligently detects relevant columns, trains a machine learning model, and produces a polished, multi-sheet Excel report — all without writing a single line of code.

The platform is designed around a core business question: **"Which customers are about to leave, and how much revenue is at risk?"**

```
Raw Data File  →  Auto-Clean  →  ML Model  →  Segmentation  →  Executive Excel Report
  (any format)     + Repair       Training       + CLV Calc        (6 sheets + 6 charts)
```

---

## ✨ Key Features

| Feature | Description |
|---|---|
| 🗂️ **Universal File Ingestion** | Supports CSV, Excel (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`), ODS, TSV with automatic handling of multiple encodings and delimiters |
| 🧹 **Intelligent Data Cleaning** | Cleans currency symbols, mixed number formats (US/EU/India), duplicates, missing values, and logical inconsistencies |
| 🔍 **Smart Column Detection** | Automatically detects Churn, Revenue, Tenure, and Customer ID using keyword normalisation |
| 🤖 **Dual ML Models** | Uses Random Forest (primary) and Gradient Boosting (fallback), with heuristic scoring for small or invalid datasets |
| 📊 **Synthetic Label Fallback** | Generates churn labels using a percentile-based risk scoring method when no churn column exists |
| 🧮 **CLV Prediction** | Calculates Customer Lifetime Value using: `TotalCharges × (1 – Churn_Prob)` |
| 🎯 **2×2 Segmentation Matrix** | Classifies customers into High/Low Risk × High/Low Value segments for action prioritisation |
| 📈 **Executive Excel Report** | Generates a multi-sheet Excel report with 14 KPI cards, 6 charts, and colour-coded risk indicators |
| 🔒 **Safe File Saving** | Implements retry logic with auto-renaming (timestamp-based) for locked or open Excel files |
| ⚠️ **Transparency Warnings** | Flags synthetic data usage and suppresses misleading ML metrics for ethical reporting |

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
│ tried        │ locale fixes │ 0.0 → 1.0   │                    │
│              │              │              │ Processed Data     │
│ Excel, CSV,  │ Label encode │ Segment into │                    │
│ ODS, TSV     │ categoricals │ 4 groups     │ Raw Data           │
│              │              │              │                    │
│              │ Engineer     │ Compute CLV  │ Data Quality       │
│              │ features     │ & KPIs       │ Audit Report       │
└──────────────┴──────────────┴──────────────┴────────────────────┘
```

---

## 🛠️ Installation

### Prerequisites

- Python **3.8 or higher**
- `pip` package manager
- Tkinter (included with most Python distributions)

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
python churn_analytics.py
```

A GUI window will guide you through the rest.

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
- Auto-detect relevant columns
- Train the ML model
- Compute all KPIs and segments

**Step 4 — Review the Summary**
A results dialog shows:
- Detected churn column (or synthetic fallback notification)
- Model accuracy and AUC-ROC (suppressed if synthetic labels were used)
- Key findings: churn rate, high-risk count, revenue at risk

**Step 5 — Save the Report**
Choose a save location and filename. The `.xlsx` report is generated immediately.

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

The tool normalises all column names (lowercase, alphanumeric only) and searches for keywords in priority order. The **most specific keyword wins**.

| Role | Keywords Searched (in priority order) |
|---|---|
| **Churn** | `churn`, `attrition`, `churned`, `leave`, `exit`, `cancel`, `left`, `status` |
| **Tenure** | `tenure`, `seniority`, `duration`, `period`, `months`, `age` |
| **Monthly Revenue** | `monthlycharge`, `monthlyfee`, `monthly`, `charge`, `rate`, `fee`, `price`, `subscription` |
| **Total Revenue** | `totalcharge`, `totalrevenue`, `lifetimevalue`, `total`, `revenue`, `bill`, `spend`, `amount`, `sales`, `ltv` |
| **Customer ID** | `customerid`, `custid`, `userid`, `accountid`, `clientid`, `memberid`, `id` |

**Conflict resolution:** No column can serve two roles. If a match has already been claimed, the next-best match is used.

**Fallback:** If a column is not found, the tool falls back to numeric columns by index position, or uses sensible defaults (e.g., median tenure = 12 months).

### Churn Label Mapping

String values in the churn column are mapped to binary labels automatically:

| → Churned (1) | → Retained (0) |
|---|---|
| `Yes`, `Y`, `True`, `1`, `Churned`, `Left`, `Cancelled`, `Exit`, `Quit`, `Inactive`, `Lost`, `Gone`, `Departed`, `Attrited`, `Closed` | `No`, `N`, `False`, `0`, `Active`, `Retained`, `Stay`, `Stayed`, `Current`, `Existing`, `Ongoing`, `Present`, `Alive`, `Good`, `Loyal` |

> If more than **60%** of values cannot be mapped to either side, **synthetic labels** are generated instead.

---

## 🤖 Machine Learning Pipeline

### Models

| Priority | Model | Configuration |
|---|---|---|
| Primary | **Random Forest** | 300 trees, max depth 12, balanced class weights, random state 42 |
| Fallback | **Gradient Boosting** | 150 estimators, max depth 5, random state 42 |
| Last Resort | **Heuristic Score** | Percentile rank formula — no training required |

### Feature Set

All 7 features below are used for model training:

| Feature | Formula / Source | Description |
|---|---|---|
| `tenure` | Raw column | Months the customer has been with the company |
| `MonthlyCharges` | Raw column | Current monthly billing amount |
| `TotalCharges` | Raw column | Cumulative historical charges |
| `AvgMonthlySpend` | `TotalCharges / (tenure + 1)` | Normalised spend over customer lifetime |
| `ValueScore` | `log(1 + TotalCharges)` | Log-transformed total value — reduces skew |
| `LoyaltyScore` | `tenure × (TotalCharges / (MonthlyCharges + 1))` | Combined tenure-spend loyalty indicator |
| `SpendVariance` | `\|MonthlyCharges – AvgMonthlySpend\| / (AvgMonthlySpend + 1)` | Billing instability signal |

### Training Conditions

ML training is attempted only when:
- Dataset has **≥ 20 rows**
- Churn column has **≥ 2 unique values** (not all churned or all retained)
- Stratified train/test split (80/20) is applied when both classes have sufficient samples

### Metrics Transparency

> ⚠️ When churn labels are **synthetic**, accuracy and AUC-ROC are **intentionally suppressed**. Reporting metrics on synthetic labels would only measure how well the model learned the tool's own formula — not real customer behaviour. This is clearly flagged throughout the report with a warning banner.

---

## 📊 Output Report Structure

The generated `.xlsx` report always contains **13 sheets**:

### Sheet 1 — Dashboard
- **Title banner** with file name, model used, accuracy, and AUC-ROC
- **Synthetic labels warning banner** (shown only when applicable)
- **14 KPI Cards** across 2 rows:
  - Row 1: Total Customers · Churn Rate · Retention Rate · High Risk · Medium Risk · Low Risk · Avg Tenure
  - Row 2: Avg CLV · Annualised Risk Revenue · Total CLV Portfolio · Avg Monthly Charge · ML Model · Accuracy · AUC-ROC
- **6 Embedded Charts:**

| # | Chart Title | Chart Type |
|---|---|---|
| 1 | Customer Segments by Count | Column |
| 2 | Risk Tier Breakdown | Doughnut |
| 3 | Churn Probability Distribution | Column |
| 4 | Predicted CLV Distribution | Bar |
| 5 | Customer Tenure Distribution | Column |
| 6 | Monthly Charges Distribution | Column |

### Sheet 2 — Segment Summary
- One row per customer segment with colour-coded churn rate:
  - 🔴 **RED** — Churn Rate > 60%
  - 🟡 **AMBER** — Churn Rate 30–60%
  - 🟢 **GREEN** — Churn Rate < 30%
- Columns: Customers · Churned · Churn Rate · Avg CLV · Total CLV · Avg Tenure · Avg Monthly Charge · Avg Risk Score

### Sheet 3 — High Risk Customers
- Top **1,000 customers** with Churn Probability ≥ 60%, sorted highest-risk first
- Columns: Customer ID (if detected) · Tenure · Monthly Charges · Total Charges · Predicted CLV · Churn Probability · Segment · Risk Tier
- Auto-filter and frozen header row enabled

### Sheet 4 — Processed Data
- Full dataset with all engineered features and ML outputs
- Per-column numeric formatting (decimals, currency, percentages)
- Auto-filter and frozen header row enabled

### Sheet 5 — Raw Data
- Original file content, completely untouched
- Preserved for audit, data lineage, and verification

### Sheet 6 — Data Quality Report
- Full pipeline audit trail including:
  - File information (name, path, original vs cleaned row counts)
  - Column detection results and fallback decisions
  - ML model details, features used, and metrics
  - Methodology notes and formula explanations
  - All output KPIs with their calculation basis

### Sheets 7–13 — Chart Source Data
- `Data – Segments` · `Data – Risk Tiers` · `Data – Churn Prob`
- `Data – CLV` · `Data – Tenure` · `Data – Monthly`
- These sheets feed the dashboard charts directly; accessible but not the primary deliverable

---

## 🧮 Feature Engineering

| Engineered Column | Formula | Business Meaning |
|---|---|---|
| `AvgMonthlySpend` | `TotalCharges / (tenure + 1)` | Normalised monthly spend over customer lifetime |
| `ValueScore` | `log(1 + TotalCharges)` | Log-transformed total value — reduces right skew |
| `LoyaltyScore` | `tenure × TotalCharges / (MonthlyCharges + 1)` | Combined tenure-spend loyalty indicator |
| `SpendVariance` | `\|Monthly – Avg\| / (Avg + 1)` | Billing instability signal — flags volatile payers |
| `Churn_Prob` | ML model output (or heuristic rank) | Predicted probability of churning — range 0.0 to 1.0 |
| `Predicted_CLV` | `TotalCharges × (1 – Churn_Prob)` | Retention-weighted historical value proxy |

**Data integrity fix:** Where `TotalCharges < MonthlyCharges` (logically impossible for any active customer), TotalCharges is automatically recomputed as `tenure × MonthlyCharges`.

---

## 🎯 Customer Segmentation

Customers are placed into one of four segments using a **2×2 Risk × Value matrix**:

```
                    HIGH VALUE            LOW VALUE
                 ┌──────────────────┬──────────────────┐
  HIGH RISK      │  High Risk –     │  High Risk –     │
  (Churn ≥ 60%) │  High Value      │  Low Value       │
                 │  🔴 PRIORITY 1  │  🟠 PRIORITY 2   │
                 ├──────────────────┼──────────────────┤
  LOW RISK       │  Low Risk –      │  Low Risk –      │
  (Churn < 60%) │  High Value      │  Low Value       │
                 │  🟢 NURTURE     │  ⚪ MONITOR      │
                 └──────────────────┴──────────────────┘
```

**Definitions:**
- **High Risk** = `Churn_Prob ≥ 0.60`
- **High Value** = `Predicted_CLV ≥ median Predicted_CLV`
- **Risk Tiers:** Low (0–30%) · Medium (30–60%) · High (60–100%)

---

## ⚠️ Methodology & Limitations

### CLV Calculation
The `Predicted_CLV` in this tool is a **retention-weighted historical proxy**, not a true Customer Lifetime Value:

```
Predicted_CLV = TotalCharges × (1 – Churn_Prob)
```

This represents the portion of historical spend expected to be retained. It is **not** a survival model, discounted cash flow (DCF) CLV, or forward-looking revenue projection. Use for **relative ranking purposes only**.

### Revenue at Risk
Annualised revenue at risk is computed as:

```
Risk Revenue = Σ (MonthlyCharges × 12)  for all customers where Churn_Prob ≥ 0.60
```

This is a **forward-looking run-rate proxy**. Historical `TotalCharges` for the same group is also separately reported in the Data Quality sheet for reference.

### Synthetic Label Generation
When no real churn column is detected (or >60% of values cannot be mapped), labels are generated using a percentile risk-score formula:

```
Risk Score = tenure_rank↓          × 0.40
           + monthly_charges_rank↑  × 0.30
           + spend_variance_rank↑   × 0.20
           + total_charges_rank↓    × 0.10

Customers in the top 35% by Risk Score  →  Labelled as Churned (1)
Remaining 65%                           →  Labelled as Retained (0)
```

ML accuracy and AUC-ROC are **suppressed** when synthetic labels are in use, as they would only reflect self-referential formula consistency — not true predictive power.

---

## 📂 Project Structure

```
churn-intelligence-platform/
│
├── churn_analytics.py       # Main application — single self-contained file
├── requirements.txt         # All Python dependencies with version pins
├── README.md                # This documentation file
│
└── sample_data/             # (Optional) Test datasets for validation
    ├── telecom_churn.csv
    └── retail_customers.xlsx
```

---

## 📦 Dependencies

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

Install all dependencies with:

```bash
pip install -r requirements.txt
```

---

## 👨‍💼 Author

<div align="center">

| Field | Detail |
|---|---|
| **Name** | Abhishek |
| **Program** | MBA in Business Analytics |
| **Institution** | Manipal Academy of Higher Education (MAHE) |
| **Roll Number** | 24154041051 |
| **Project Title** | Customer Churn Intelligence Platform |
| **Version** | v3.0 — Final Submission |

</div>

---

## 📄 License

This project was developed as an academic final-year MBA project at MAHE. It is intended for **educational and non-commercial use only**.

---

<div align="center">

*Built with ❤️ for business analytics practitioners who need answers, not just data.*

</div>



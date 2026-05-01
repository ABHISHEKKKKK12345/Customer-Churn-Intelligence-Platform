# 📊 Customer Churn Intelligence Platform — GUI Edition

<div align="center">

**A Fully Windowed Desktop Application for ML-Based Customer Churn Analytics**

<div align="center">

[![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![scikit-learn](https://img.shields.io/badge/scikit--learn-ML-F7931E?style=for-the-badge&logo=scikit-learn&logoColor=white)](https://scikit-learn.org/)
[![Pandas](https://img.shields.io/badge/Pandas-Data-150458?style=for-the-badge&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![Tkinter](https://img.shields.io/badge/Tkinter-GUI-007ACC?style=for-the-badge&logo=python&logoColor=white)](https://docs.python.org/3/library/tkinter.html)
[![XlsxWriter](https://img.shields.io/badge/XlsxWriter-Reports-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://xlsxwriter.readthedocs.io/)
[![License](https://img.shields.io/badge/License-MIT-blueviolet?style=for-the-badge)](../LICENSE)

</div>

*Transform raw customer data into executive-grade churn intelligence — fully windowed, zero console, zero configuration.*

</div>

---

## 📋 Table of Contents

- [Overview](#overview)
- [Quick Start](#quick-start)
- [Key Features](#key-features)
- [Application Layout](#application-layout)
- [Installation](#installation)
- [Usage](#usage)
- [Supported Input Formats](#supported-input-formats)
- [Column Auto-Detection](#column-auto-detection)
- [Machine Learning Pipeline](#machine-learning-pipeline)
- [Feature Engineering](#feature-engineering)
- [Customer Segmentation](#customer-segmentation)
- [Methodology & Limitations](#methodology)
- [Excel Report Structure](#excel-report-structure)
- [Building a Standalone EXE](#build-exe)
- [Common Issues & Fixes](#common-issues)
- [Project Structure](#project-structure)
- [Dependencies](#dependencies)
- [Author](#author)
- [License](#license)

---

<a id="overview"></a>
## 🎯 Overview

The **GUI Edition** is a fully windowed desktop application of the Customer Churn Intelligence Platform. It wraps the complete analytics pipeline — data ingestion, cleaning, ML training, segmentation, and Excel report generation — inside a polished, scroll-able Tkinter interface.

No terminal interaction is required at any point. Launch the app, pick a file, click **Run**, and receive a production-ready Excel report.

```
Launch App  →  Select File  →  Click Run  →  Progress Bar  →  KPI Cards  +  Excel Report
```

> **Application identity**
> - Title: `Customer Churn Intelligence Platform (GUI)`
> - Version: `v1.0`
> - Entry point: `churn-intelligence-platform-gui.py`

---

<a id="quick-start"></a>
## ⚡ Quick Start

```bash
# 1. Clone the repo
git clone https://github.com/your-username/churn-intelligence-platform.git
cd churn-intelligence-platform/gui

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the GUI
python churn-intelligence-platform-gui.py
```

A splash screen loads, followed by the main application window.

---

<a id="key-features"></a>
## ✨ Key Features

| Feature | Detail |
|---|---|
| 🖥️ **Fully Windowed** | Zero console interaction — splash screen, scrollable main window, live KPI cards |
| 🗂️ **Multi-format Ingestion** | CSV, Excel (`.xlsx` `.xls` `.xlsm` `.xlsb`), ODS, TSV, TXT — auto-tries 6 encodings and 6 delimiters |
| 🧹 **Auto Data Cleaning** | Removes blanks, duplicates, unnamed columns; repairs currency symbols and locale number formats |
| 🔍 **Smart Column Detection** | Auto-identifies Churn, Tenure, Revenue, and Customer ID columns via keyword normalisation |
| 🤖 **Dual ML Models** | Random Forest (primary) → Gradient Boosting (fallback) → heuristic percentile score (last resort) |
| 📉 **Churn Probability** | Per-customer score 0.0–1.0 with Low / Medium / High risk tier classification |
| 💰 **CLV Prediction** | Retention-weighted Customer Lifetime Value proxy: `TotalCharges × (1 – Churn_Prob)` |
| ⚠️ **Revenue at Risk** | Annualised run-rate (`MonthlyCharges × 12`) for all customers with `Churn_Prob ≥ 60%` |
| 🧩 **2×2 Segmentation** | High/Low Risk × High/Low Value matrix — 4 actionable customer segments |
| 📊 **Excel Report** | 12-sheet workbook with 14 KPI cards, 6 embedded charts, colour-coded risk indicators, and full audit trail |
| 🔄 **Locked File Recovery** | Retry logic (up to 3 attempts) with auto-timestamp rename when the target file is open |
| 📋 **Live Activity Log** | Timestamped, colour-coded in-app log updated in real time during analysis |
| 💡 **Synthetic Label Mode** | Generates churn labels via percentile risk score when no churn column exists; ML metrics suppressed and flagged |
| 🔒 **Safe Thread Design** | Analysis runs on a background daemon thread — UI stays responsive; quit confirmation shown if analysis is in progress |

---

<a id="application-layout"></a>
## 🖥️ Application Layout

<table>
<thead>
<tr>
<th align="left">📊 &nbsp; Customer Churn Intelligence Platform &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; v1.0</th>
</tr>
</thead>
<tbody>
<tr><td><b>Author strip</b> — Abhishek</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td>ℹ &nbsp; <b>What This Tool Does</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <i>info card</i></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td>📂 &nbsp; <b>Step 1 — Select Dataset</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <i>file browser</i><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; • File path entry + Browse… button<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; • Auto-shows: filename, size, directory</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td>💾 &nbsp; <b>Step 2 — Choose Save Location</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <i>save dialog</i><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; • Default: <code>{input_basename}_Churn_Report.xlsx</code><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; • Desktop fallback if no file selected yet</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td>🚀 &nbsp; <b>Step 3 — Run Analysis</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <i>confirm → run</i><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ▶ &nbsp; Run Churn Analysis &amp; Generate Report</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td>━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ &nbsp; <b>Progress Bar</b> &nbsp; ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Status label — live updates at each pipeline stage</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td>📈 &nbsp; <b>Results — KPI Cards</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <i>post-run</i><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; • 12 colour-coded cards × 4-column grid<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; • Model / Accuracy / AUC-ROC summary row</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td>📋 &nbsp; <b>Activity Log</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <i>always visible</i><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; • Timestamped, colour-coded scrollable log</td></tr>
<tr><td>&nbsp;</td></tr>
</tbody>
</table>

### Splash Screen

A branded splash screen is displayed on launch with an animated progress bar, stepping through 5 initialisation stages before opening the main window.

---

<a id="installation"></a>
## 🛠️ Installation

### Prerequisites

| Requirement | Version |
|---|---|
| Python | **3.8 or higher** |
| pip | Latest |
| OS | Windows 10/11 · macOS 12+ · Ubuntu 20+ |

> ⚠️ **Tkinter** ships with Python on Windows and macOS. If missing on Linux:
> ```bash
> # Ubuntu / Debian
> sudo apt-get install python3-tk
>
> # Fedora / RHEL
> sudo dnf install python3-tkinter
> ```

### Step 1 — Clone or Download

```bash
git clone https://github.com/your-username/churn-intelligence-platform.git
cd churn-intelligence-platform/gui
```

Or download the ZIP, extract it, and navigate into the `gui/` folder.

### Step 2 — Create a Virtual Environment *(Recommended)*

```bash
python -m venv .venv

# Activate on Windows:
.venv\Scripts\activate

# Activate on macOS / Linux:
source .venv/bin/activate
```

### Step 3 — Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 4 — Launch the Application

```bash
python churn-intelligence-platform-gui.py
```

---

<a id="usage"></a>
## 🚀 Usage

The application is entirely GUI-driven — no command-line arguments or config files needed.

### Step 1 — Select Your Dataset
Click **Browse…** next to *Dataset File* and pick any supported file (see [Supported Input Formats](#-supported-input-formats)).

The app displays:
- Filename and file size
- Directory path
- Auto-suggested save path: `{input_basename}_Churn_Report.xlsx`

### Step 2 — Choose Save Location
The save path is pre-filled automatically when you select a file. Override it anytime using the second **Browse…** button. If no file has been selected yet, the default is `~/Desktop/Churn_Analytics_Report.xlsx`.

### Step 3 — Run the Analysis
Click **▶ Run Churn Analysis & Generate Report**.

A confirmation dialog summarises your file and save path. Click **YES** to proceed. The analysis runs on a background thread — the UI remains fully responsive throughout.

### Step 4 — Monitor Progress
Watch the progress bar and the live **Activity Log** for timestamped status messages at each pipeline stage.

### Step 5 — Review Results
When complete:
- **12 KPI cards** populate inside the app with colour-coded metrics
- A **success dialog** shows the full results summary
- The **Excel report** is ready at your chosen save location

> 💡 **If you close the window while analysis is running**, a confirmation dialog will ask if you want to quit. The background thread is a daemon and will terminate safely.

---

<a id="supported-input-formats"></a>
## 📁 Supported Input Formats

| Format | Extensions | Notes |
|---|---|---|
| CSV | `.csv` | Auto-detects separator: `,` `;` `\t` `\|` `:` |
| Excel (modern) | `.xlsx` `.xlsm` `.xlsb` | Engine: `openpyxl` |
| Excel (legacy) | `.xls` | Engine: `xlrd` |
| OpenDocument | `.ods` | Engine: `odf` (requires `odfpy`) |
| Text / TSV | `.txt` `.tsv` | Treated as delimited text |

**Encodings tried automatically:** `utf-8-sig`, `utf-8`, `latin1`, `cp1252`, `iso-8859-1`, `utf-16`

**Currency symbols stripped automatically:** ₹ $ € £ ¥

**Number locale support:** Indian (1,23,456) · US (1,234.56) · European (1.234,56)

**Minimum requirement:** At least **1 data row** and **2 columns**.

---

<a id="column-auto-detection"></a>
## 🔍 Column Auto-Detection

All column names are normalised (lowercase, alphanumeric only) before matching. The **most specific keyword wins**, and no column can serve two roles simultaneously.

| Role | Keywords Searched (priority order) |
|---|---|
| **Churn** | `churn`, `attrition`, `churned`, `leave`, `exit`, `cancel`, `left`, `status` |
| **Tenure** | `tenure`, `seniority`, `duration`, `period`, `months`, `age` |
| **Monthly Revenue** | `monthlycharge`, `monthlyfee`, `monthly`, `charge`, `rate`, `fee`, `price`, `subscription` |
| **Total Revenue** | `totalcharge`, `totalrevenue`, `lifetimevalue`, `total`, `revenue`, `bill`, `spend`, `amount`, `sales`, `ltv` |
| **Customer ID** | `customerid`, `custid`, `userid`, `accountid`, `clientid`, `memberid`, `id` |

**Fallback behaviour:** If a column is not found, the tool falls back to numeric columns by index position, or uses hardcoded defaults (tenure = 12 months, monthly = Rs. 500, total = Rs. 6,000).

### Churn Label Mapping

| Churned → `1` | Retained → `0` |
|---|---|
| `Yes`, `Y`, `True`, `1`, `Churned`, `Left`, `Cancelled`, `Cancel`, `Exit`, `Quit`, `Inactive`, `Lost`, `Gone`, `Departed`, `Attrited`, `Closed` | `No`, `N`, `False`, `0`, `Active`, `Retained`, `Stay`, `Stayed`, `Current`, `Existing`, `Ongoing`, `Present`, `Alive`, `Good`, `Loyal` |

> ⚠️ If more than **60%** of values cannot be mapped, synthetic labels are generated automatically (see [Methodology & Limitations](#methodology)).

---

<a id="machine-learning-pipeline"></a>
## 🤖 Machine Learning Pipeline

```
Raw File
   │
   ▼
Load  →  CSV / Excel / ODS / TSV  (6 encodings × 6 delimiters tried)
   │
   ▼
Clean  →  Drop nulls · duplicates · unnamed columns
          Repair currency symbols · locale number formats
   │
   ▼
Detect Columns  →  Churn · Tenure · Revenue · Customer ID
   │
   ▼
Numeric Conversion  →  clean_num() for all non-ID columns (≥40% parse threshold)
Label Encoding      →  object columns with 2–20 unique values
   │
   ▼
Feature Engineering  (see table below)
   │
   ▼
Churn Labels
   ├── Real column found AND ≤60% unmapped  →  map YES/NO/numeric values
   └── Not found OR >60% unmapped           →  percentile risk-score formula (top 35% = churned)
   │
   ▼
Median Imputation  →  SimpleImputer(strategy="median") on all 7 features
   │
   ▼
Train (80/20 stratified split when both classes have sufficient samples)
   ├── Attempt 1 :  Random Forest        (300 trees · depth 12 · balanced class weight · n_jobs=-1)
   └── Attempt 2 :  Gradient Boosting    (150 estimators · depth 5)  ← fallback
   If rows < 20 or single class:  Heuristic percentile score only
   │
   ▼
Churn_Prob (0.0–1.0) per customer  →  CLV  →  Segmentation  →  KPIs  →  Excel Report
```

### ML Training Conditions

- Minimum **20 rows** required (`MIN_ROWS_FOR_ML = 20`)
- Churn column must have **≥ 2 unique values**
- Stratified split applied when both classes have enough samples; non-stratified otherwise
- Accuracy and AUC-ROC **suppressed** when synthetic labels are in use

---

<a id="feature-engineering"></a>
## 🧮 Feature Engineering

| Feature | Formula | Business Meaning |
|---|---|---|
| `AvgMonthlySpend` | `TotalCharges / (tenure + 1)` | Normalised spend over customer lifetime |
| `ValueScore` | `log(1 + TotalCharges)` | Log-transformed total value — reduces right skew |
| `LoyaltyScore` | `tenure × (TotalCharges / (MonthlyCharges + 1))` | Combined tenure-spend loyalty indicator |
| `SpendVariance` | `\|MonthlyCharges – AvgMonthlySpend\| / (AvgMonthlySpend + 1)` | Billing instability — flags volatile payers |
| `Churn_Prob` | ML model output (or heuristic rank) | Predicted probability of churning — range `0.0` to `1.0` |
| `Predicted_CLV` | `TotalCharges × (1 – Churn_Prob)` | Retention-weighted historical value proxy |

**Data integrity fix:** Where `TotalCharges < MonthlyCharges` (logically impossible), TotalCharges is recomputed as `tenure × MonthlyCharges`.

**Heuristic fallback score** (used when ML cannot run):
```
Score = tenure_rank↓ × 0.55
      + monthly_charges_rank↑ × 0.25
      + spend_variance_rank↑  × 0.20
```

---

<a id="customer-segmentation"></a>
## 🎯 Customer Segmentation

Every customer is assigned to one of four segments using a **2×2 Risk × Value matrix:**

|  | **High Value** (`Predicted_CLV ≥ median`) | **Low Value** (`Predicted_CLV < median`) |
|---|---|---|
| **High Risk** (`Churn_Prob ≥ 0.60`) | 🔴 High Risk – High Value · **PRIORITY 1** | 🟠 High Risk – Low Value · **PRIORITY 2** |
| **Low Risk** (`Churn_Prob < 0.60`) | 🟢 Low Risk – High Value · **NURTURE** | 🟡 Low Risk – Low Value · **MONITOR** |

**Risk Tiers** (stored as `Risk_Tier` column):

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

This is **not** a survival model or discounted cash flow (DCF) CLV. Use for **relative ranking only.**

### Revenue at Risk

```
Risk Revenue = Σ (MonthlyCharges × 12)  for all customers where Churn_Prob ≥ 0.60
```

A forward-looking annualised run-rate proxy. Historical `TotalCharges` for the same group is also reported separately in the Data Quality sheet.

### Synthetic Label Generation

When no real churn column is detected (or >60% of values cannot be mapped):

```
Risk Score = tenure_rank↓          × 0.40
           + monthly_charges_rank↑  × 0.30
           + spend_variance_rank↑   × 0.20
           + total_charges_rank↓    × 0.10

Top 35% by Risk Score  →  Churned (1)
Remaining 65%          →  Retained (0)
```

ML accuracy and AUC-ROC are **suppressed** in this mode — they would only measure how well the model reproduced the formula, not real customer behaviour. A warning banner is shown throughout the report.

---

<a id="excel-report-structure"></a>
## 📊 Excel Report Structure

The generated `.xlsx` file always contains exactly **12 sheets:**

| # | Sheet | Contents |
|---|---|---|
| 1 | **Dashboard** | 14 KPI cards across 2 rows + 6 embedded charts + synthetic label warning banner (if applicable) |
| 2 | **Segment Summary** | Per-segment aggregation — Customers · Churned · Churn Rate (RAG colour-coded) · Avg CLV · Total CLV · Avg Tenure · Avg Monthly · Avg Risk Score |
| 3 | **High Risk Customers** | Up to 1,000 customers with `Churn_Prob ≥ 60%`, sorted highest-risk first; auto-filter + frozen header |
| 4 | **Processed Data** | Full dataset with all engineered features, per-column numeric formatting, auto-filter + frozen header |
| 5 | **Raw Data** | Original file exactly as uploaded — completely untouched |
| 6 | **Data Quality Report** | Full pipeline audit: file info · column detection · model metrics · methodology notes · all KPIs |
| 7 | **Data – Segments** | Chart-source: segment distribution |
| 8 | **Data – Risk Tiers** | Chart-source: risk tier doughnut |
| 9 | **Data – Churn Prob** | Chart-source: churn probability histogram |
| 10 | **Data – CLV** | Chart-source: CLV distribution bar chart |
| 11 | **Data – Tenure** | Chart-source: tenure distribution |
| 12 | **Data – Monthly** | Chart-source: monthly charges distribution |

### Dashboard KPI Cards

**Row 1:** Total Customers · Churn Rate · Retention Rate · High Risk (≥60%) · Medium Risk (30–60%) · Low Risk (<30%) · Avg Tenure

**Row 2:** Avg Customer CLV · Annualised Risk Revenue · Total CLV Portfolio · Avg Monthly Charge · ML Model · Model Accuracy · AUC-ROC Score

### Dashboard Charts

| # | Title | Type |
|---|---|---|
| 1 | Customer Segments by Count | Column |
| 2 | Risk Tier Breakdown | Doughnut |
| 3 | Churn Probability Distribution | Column |
| 4 | Predicted CLV Distribution | Bar |
| 5 | Customer Tenure Distribution | Column |
| 6 | Monthly Charges Distribution | Column |

---

<a id="build-exe"></a>
## 🖥️ Building a Standalone EXE

Convert the app to a single executable that runs on any Windows machine without Python installed.

### Prerequisites

- All dependencies from `requirements.txt` installed
- A clean virtual environment is strongly recommended
- Build on **Windows** if distributing to Windows users

### Step 1 — Install PyInstaller

```bash
pip install pyinstaller
```

### Step 2 — Standard Build

```bash
pyinstaller --onefile --windowed --name "ChurnPlatform" \
  churn-intelligence-platform-gui.py
```

| Flag | Purpose |
|---|---|
| `--onefile` | Bundle everything into a single `.exe` |
| `--windowed` | Suppress the console window (GUI app) |
| `--name "ChurnPlatform"` | Set the output executable name |

### Step 3 — Optional: Custom Icon

```bash
pyinstaller --onefile --windowed --name "ChurnPlatform" \
  --icon="icon.ico" churn-intelligence-platform-gui.py
```

The icon must be a `.ico` file. Convert PNG → ICO using Pillow or an online converter.

### Step 4 — Extended Build (If Modules Are Missing at Runtime)

If the standard build produces a runtime `ModuleNotFoundError`, use explicit hidden imports:

```bash
pyinstaller --onefile --windowed \
  --name "Customer_Churn_Intelligence_Platform_GUI" \
  --hidden-import sklearn.ensemble \
  --hidden-import sklearn.ensemble._forest \
  --hidden-import sklearn.ensemble._gb \
  --hidden-import sklearn.tree \
  --hidden-import sklearn.tree._classes \
  --hidden-import sklearn.impute \
  --hidden-import sklearn.preprocessing \
  --hidden-import sklearn.model_selection \
  --hidden-import sklearn.metrics \
  --hidden-import openpyxl \
  --hidden-import xlrd \
  --hidden-import odf \
  --hidden-import odf.opendocument \
  --hidden-import xlsxwriter \
  --hidden-import tkinter \
  --hidden-import numpy \
  --hidden-import pandas \
  churn-intelligence-platform-gui.py
```

### Step 5 — Find Your Executable

```
gui/
└── dist/
    └── ChurnPlatform.exe      ← distribute this file
```

The `build/` folder and `.spec` file are not needed and can be deleted.

### EXE Build Troubleshooting

| Issue | Fix |
|---|---|
| Antivirus flags the `.exe` | False positive from PyInstaller packing method — add `dist/` to antivirus exclusions during testing |
| App opens briefly then closes | Build **without** `--windowed` first (`--name ChurnPlatform_debug`) and run from terminal to see the error |
| Missing module at runtime | Add `--hidden-import=<package>` for the failing module |
| Large file size (~200–400 MB) | Expected — NumPy, pandas, and scikit-learn are bundled; this is normal |

---

<a id="common-issues"></a>
## 🐛 Common Issues & Fixes

| Problem | Cause | Fix |
|---|---|---|
| `ModuleNotFoundError: tkinter` | tkinter not installed (Linux) | `sudo apt-get install python3-tk` |
| `No usable rows after cleaning` | File is empty or all-blank | Ensure the file has column headers and at least 1 data row |
| `Cannot open Excel/ODS file` | File is password-protected or corrupted | Remove password protection in Excel first |
| Report not saving | Target file already open in Excel | App prompts to auto-rename with timestamp — click **Yes** |
| ML skipped, heuristic used | Fewer than 20 rows, or single-class labels | Expected — heuristic still produces valid risk scores |
| Accuracy shows `N/A (synthetic)` | No real churn column detected | Normal — see [Methodology & Limitations](#methodology) |
| EXE missing DLLs on Windows | Visual C++ Redistributable missing | Install [VC++ Redistributable](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist) |
| UI freezes during analysis | Should not happen — analysis is threaded | If observed, check that no modal dialog is waiting behind another window |

---

<a id="project-structure"></a>
## 📂 Project Structure

```
churn-intelligence-platform/               ← repository root
│
├── docs/                                  # Additional documentation
│
├── gui/                                   ← this module lives here
│   ├── churn-intelligence-platform-gui.py # Main GUI application — single self-contained file
│   ├── requirements.txt                   # GUI-specific Python dependencies
│   ├── .gitignore                         # Excludes build/, dist/, __pycache__/, *.xlsx, etc.
│   └── README.md                          # This documentation
│
├── samples/                               # Sample datasets for testing
│
├── versions/                              # Version history / release snapshots
│
├── customer_churn_intelligence_system.py  # Core CLI/dialog version of the platform
├── requirements.txt                       # Root-level dependencies (core version)
├── .gitignore                             # Root-level ignore rules
└── README.md                              # Root-level documentation (core version)
```

> 💡 `dist/`, `build/`, and `.spec` files are excluded via `gui/.gitignore`. To generate the `.exe`, follow [Building a Standalone EXE](#build-exe) above.

---

<a id="dependencies"></a>
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
| `tkinter` | GUI framework — file dialogs, windows, widgets | Built-in (stdlib) |

---

<a id="author"></a>
## 👤 Author

<div align="center">

| Field | Detail |
|---|---|
| **Name** | [Abhishek](https://www.linkedin.com/in/abhishek-srivastava-1538461b1/) |
| **Project** | Customer Churn Intelligence Platform |

</div>

---

<a id="license"></a>
## 📄 License

This project is released under the [MIT License](../LICENSE). You are free to use, modify, and distribute it with attribution.

---

<div align="center">

*Built with ❤️ for business analysts who need answers, not just data.*

</div>

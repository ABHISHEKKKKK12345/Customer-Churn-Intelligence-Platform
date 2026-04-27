# 📊 Customer Churn Intelligence Platform
## 🖥️ GUI Application (Desktop Interface)

> Machine Learning-Based Customer Churn Analytics System

## 📁 Module Overview

This folder contains the **GUI version** of the Customer Churn Intelligence Platform.

- Fully interactive desktop application  
- Designed for non-technical users  
- Uses the same backend logic as the core system
  
---

## 📌 Overview

The **Customer Churn Intelligence Platform** is a fully windowed desktop application that transforms raw customer data into actionable business intelligence — with zero coding required from the end user.

Upload any customer dataset, click **Run**, and receive a professionally formatted Excel report complete with KPI cards, charts, ML predictions, and a prioritised retention action list.

---

## ✨ Features

| Capability | Detail |
|---|---|
| 🗂 **Multi-format Ingestion** | CSV, Excel (`.xlsx` `.xls` `.xlsm` `.xlsb`), ODS, TSV, TXT |
| 🧹 **Auto Data Cleaning** | Removes blanks, duplicates, unnamed columns; repairs mixed formats |
| 🔍 **Smart Column Detection** | Auto-detects Churn, Tenure, Revenue, Customer ID columns |
| 🤖 **Machine Learning** | Random Forest → Gradient Boosting fallback; heuristic if data insufficient |
| 📈 **Churn Probability** | Per-customer churn score (0–100%) with risk tier classification |
| 💰 **CLV Prediction** | Retention-weighted Customer Lifetime Value proxy |
| ⚠️ **Revenue at Risk** | Annualised run-rate for all high-risk customers |
| 🧩 **Segmentation** | 2×2 Risk × Value matrix (4 actionable segments) |
| 📊 **Excel Report** | 14 KPI cards, 6 embedded charts, 12+ formatted sheets |
| 🔄 **Locked File Recovery** | Auto-retry + timestamp rename if target file is open |
| 📋 **Live Activity Log** | Timestamped in-app log with colour-coded status messages |
| 💡 **Synthetic Label Mode** | Generates risk-score labels if no churn column exists |

---

## 🖥️ Application Layout

```
┌──────────────────────────────────────────────────────────┐
│  📊  Customer Churn Intelligence Platform                │
├──────────────────────────────────────────────────────────┤
│  Author strip                                            │
├──────────────────────────────────────────────────────────┤
│  ℹ  What This Tool Does          (info card)             │
│  📂  Step 1 — Select Dataset     (file browser)         │
│  💾  Step 2 — Choose Save Path   (save dialog)          │
│  🚀  Step 3 — Run Analysis       (big green button)     │
│  ━━━━━━━━━━━━  Progress Bar  ━━━━━━━━━━━━                │
│  📈  Results — KPI Cards         (live after run)       │
│  📋  Activity Log                (timestamped log)      │
└──────────────────────────────────────────────────────────┘
```

---

## 📦 Installation

### Prerequisites

| Requirement | Version |
|---|---|
| Python | 3.9 – 3.12 (3.10+ recommended) |
| pip | Latest |
| OS | Windows 10/11 · macOS 12+ · Linux (Ubuntu 20+) |

> ⚠️ **tkinter** is included with most Python installations. If missing on Linux:
> ```bash
> sudo apt-get install python3-tk
> ```

### Step 1 — Clone or Download

```bash
git clone https://github.com/your-username/churn-intelligence-platform.git
cd gui
```

Or download the ZIP, extract it, and open the `gui/` folder.

### Step 2 — Create a Virtual Environment (Recommended)

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

### Step 4 — Run the Application

```bash
python churn_analytics_gui.py
```

A splash screen will appear, followed by the main application window.

---

## 🚀 How to Use

### 1. Select Your Dataset
Click **Browse…** next to *Dataset File* and pick any supported file.
The app shows the filename, size, and auto-suggests a save path.

### 2. Choose Save Location
The default save path is auto-set next to your input file.
You can override it using the second **Browse…** button.

### 3. Run the Analysis
Click the large **▶ Run Churn Analysis & Generate Report** button.
Confirm the dialog and watch the progress bar and activity log.

### 4. Review Results
After completion:
- **KPI cards** appear inside the app (churn rate, risk counts, CLV, etc.)
- A **success dialog** summarises all key metrics
- The **Excel report** is ready at your chosen save path

---

## 📁 Excel Report Structure

The generated `.xlsx` report contains the following sheets:

| Sheet | Contents |
|---|---|
| **Dashboard** | 14 KPI cards + 6 embedded charts (segments, risk tiers, CLV, tenure, charges, churn probability) |
| **Segment Summary** | Per-segment aggregation — churn rate (RAG colour-coded), CLV, tenure, monthly charge |
| **High Risk Customers** | Up to 1,000 customers with churn probability ≥ 60%, sorted highest risk first |
| **Processed Data** | All engineered features with formatted numeric columns and auto-filter |
| **Raw Data** | Original file exactly as uploaded — untouched |
| **Data Quality Report** | Full pipeline audit: file info, column detection, model metrics, methodology notes |
| **Data – Segments** | Chart-source data for segment distribution chart |
| **Data – Risk Tiers** | Chart-source data for risk tier doughnut chart |
| **Data – Churn Prob** | Chart-source data for churn probability histogram |
| **Data – CLV** | Chart-source data for CLV distribution bar chart |
| **Data – Tenure** | Chart-source data for tenure distribution chart |
| **Data – Monthly** | Chart-source data for monthly charges chart |

---

## 🤖 Machine Learning Pipeline

```
Raw File
   │
   ▼
Load (CSV / Excel / ODS / TSV)
   │
   ▼
Clean  →  Remove nulls, duplicates, unnamed columns
   │
   ▼
Detect Columns  →  Churn · Tenure · Revenue · Customer ID
   │
   ▼
Numeric Conversion  →  Currency symbols, locale formats, label encoding
   │
   ▼
Feature Engineering
   ├── AvgMonthlySpend  =  TotalCharges / (tenure + 1)
   ├── ValueScore       =  log1p(TotalCharges)
   ├── LoyaltyScore     =  tenure × (TotalCharges / (MonthlyCharges + 1))
   └── SpendVariance    =  |MonthlyCharges − AvgMonthlySpend| / (AvgMonthlySpend + 1)
   │
   ▼
Churn Labels
   ├── Real column found  →  map YES/NO/numeric values
   └── Not found          →  synthetic percentile risk score (top 35% = churned)
   │
   ▼
Train ML Model (80/20 stratified split)
   ├── Try 1: Random Forest (300 trees, depth 12, balanced class weight)
   └── Try 2: Gradient Boosting (150 estimators, depth 5) — fallback
   │
   ▼
Churn_Prob per customer  →  CLV  →  Segments  →  KPIs  →  Excel Report
```

### Segmentation Matrix

|  | **High Value** (CLV ≥ median) | **Low Value** (CLV < median) |
|---|---|---|
| **High Risk** (prob ≥ 60%) | 🔴 High Risk – High Value | 🟠 High Risk – Low Value |
| **Low Risk** (prob < 60%) | 🟢 Low Risk – High Value | 🟡 Low Risk – Low Value |

### Synthetic Label Notice

If your dataset has **no churn column**, the app automatically generates churn labels using a weighted percentile risk formula:

```
Risk = tenure_rank↓ × 0.40
     + monthly_rank↑ × 0.30
     + variance_rank↑ × 0.20
     + total_rank↓   × 0.10

Top 35% by risk score → labelled as Churned
```

ML **accuracy and AUC-ROC are suppressed** in this mode — they would only measure how well the model learned the formula itself, not real churn. Risk scores and segmentation remain valid for relative ranking.

---

## 📋 Supported Input Formats

| Format | Extensions | Notes |
|---|---|---|
| CSV | `.csv` | Auto-detects separator (`,` `;` `\t` `\|` `:`) |
| Excel | `.xlsx` `.xlsm` | Reads first non-empty sheet |
| Legacy Excel | `.xls` | Requires `xlrd` |
| Excel Binary | `.xlsb` | Requires `openpyxl` |
| OpenDocument | `.ods` | Requires `odfpy` (optional) |
| Text / TSV | `.txt` `.tsv` | Treated as delimited text |

**Currency support:** ₹ $ € £ ¥ — all stripped automatically before numeric parsing.
**Locale support:** Indian (1,23,456), US (1,234.56), European (1.234,56) number formats.

---

## 🔧 Building a Standalone .exe (Windows)

Convert the app to a single executable that runs on any Windows machine **without Python installed**.

### Step 1 — Install PyInstaller

```bash
pip install pyinstaller
```

### Step 2 — Build the Executable

Run this command from inside the `gui/` folder:

```bash
pyinstaller --onefile --windowed --name "ChurnPlatform" churn_analytics_gui.py
```

| Flag | Purpose |
|---|---|
| `--onefile` | Bundle everything into a single `.exe` file |
| `--windowed` | Suppress the console window (GUI-only app) |
| `--name "ChurnPlatform"` | Sets the output executable name |

### Step 3 — (Optional) Add a Custom Icon

```bash
pyinstaller --onefile --windowed --name "ChurnPlatform" --icon="icon.ico" churn_analytics_gui.py
```

> The icon must be a `.ico` file. You can convert a PNG to ICO using online tools or Pillow.

### Step 4 — Locate Your Executable

After the build completes (1–3 minutes), your `.exe` is at:

```
dist/
└── ChurnPlatform.exe      ← share this file
```

The `build/` folder and `.spec` file can be deleted — they are not needed.

### Step 5 — Distribute

Copy `ChurnPlatform.exe` to any Windows PC and double-click to run.
No Python, no pip, no installation required.

---

### 🛠 Troubleshooting the .exe Build

**Antivirus flags the .exe**
PyInstaller executables are sometimes flagged by Windows Defender or antivirus software due to the packing method. This is a **false positive**. To resolve:
- Add the `dist/` folder to your antivirus exclusion list during testing
- Or submit the file to your antivirus vendor as a false positive

**App opens briefly then closes**
The `--windowed` flag suppresses the console. To debug, build **without** `--windowed` first:
```bash
pyinstaller --onefile --name "ChurnPlatform_debug" churn_analytics_gui.py
```
Run the debug `.exe` from a terminal to see error output.

**Missing module errors at runtime**
Add hidden imports if a module fails to bundle:
```bash
pyinstaller --onefile --windowed --hidden-import=sklearn.ensemble \
  --hidden-import=sklearn.impute --name "ChurnPlatform" churn_analytics_gui.py
```

**Large file size**
The `.exe` will be ~200–400 MB because it bundles NumPy, pandas, and scikit-learn. This is expected and normal.

---

## 🐛 Common Issues & Fixes

| Problem | Cause | Fix |
|---|---|---|
| `ModuleNotFoundError: tkinter` | tkinter not installed | `sudo apt-get install python3-tk` (Linux) |
| `No usable rows after cleaning` | File is empty or all-blank | Check your file has headers and data rows |
| `Cannot open Excel file` | Password-protected or corrupted | Remove password protection in Excel first |
| Report not saving | File already open in Excel | App prompts to auto-rename — click **Yes** |
| ML skipped, using heuristic | Fewer than 20 rows | Expected — heuristic still produces valid scores |
| Accuracy shows N/A | Synthetic churn labels | Normal — see Synthetic Label Notice above |
| `.exe` missing DLLs | Windows Visual C++ not installed | Install [VC++ Redistributable](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist) |

---

## 📂 GUI Module Structure
```
gui/
├── churn_analytics_gui.py      # Main GUI application
├── requirements.txt           # Dependencies
├── .gitignore                 # Ignore rules
└── README.md                  # This file
```


---

### ▶ Run Application

```bash
python churn_analytics_gui.py
```

---

## 📚 Dependencies

| Package | Version | Purpose |
|---|---|---|
| `numpy` | ≥ 1.24 | Numerical operations |
| `pandas` | ≥ 2.0 | Data loading, cleaning, transformation |
| `scikit-learn` | ≥ 1.3 | ML models, imputation, label encoding |
| `xlsxwriter` | ≥ 3.1 | Excel report generation (charts, formats) |
| `openpyxl` | ≥ 3.1 | Reading `.xlsx` / `.xlsm` input files |
| `xlrd` | ≥ 2.0 | Reading legacy `.xls` files |
| `tkinter` | stdlib | GUI framework (bundled with Python) |

---

## 👤 Author

**Abhishek**  
Business Analytics | Python | Machine Learning

---

## 📄 License

This project is for educational and portfolio purposes.

  

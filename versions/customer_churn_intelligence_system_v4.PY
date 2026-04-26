# ==============================================================================
#  CUSTOMER CHURN INTELLIGENCE PLATFORM
#
#  Overview:
#  End-to-end analytics system that transforms raw customer data into
#  actionable business insights.  Integrates data processing, machine
#  learning, and executive reporting to identify at-risk customers and
#  quantify potential revenue impact.
#
#  Core Functionality:
#    • Multi-format data ingestion (CSV, Excel, ODS, TSV)
#    • Automated data cleaning and preprocessing
#    • Intelligent column detection (Churn, Revenue, Tenure, Customer ID)
#    • Feature engineering for behavioural and financial analysis
#    • ML-based churn prediction (Random Forest / Gradient Boosting)
#    • Customer segmentation based on risk and value
#    • Customer Lifetime Value (CLV) prediction
#    • Revenue-at-risk estimation and prioritisation
#
#  Output Deliverables:
#    A fully automated Excel report including:
#      - Executive Dashboard (KPIs + Visual Insights)
#      - Segment Performance Summary
#      - High-Risk Customer Action List
#      - Processed Analytical Dataset
#      - Raw Source Data
#      - Data Quality & Model Audit Report
#
#
#  Author   : Abhishek
#  Project  : Customer Analytics / Business Intelligence
# ==============================================================================

import os
import re
import sys
import time
import warnings
import traceback

import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

warnings.filterwarnings("ignore")

from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
from sklearn.metrics import accuracy_score, roc_auc_score
from sklearn.impute import SimpleImputer
from sklearn.preprocessing import LabelEncoder

# ==============================================================================
#  CONSTANTS
# ==============================================================================
MAX_RETRIES    = 3
SHEET_NAME_LEN = 31          # xlsxwriter hard limit
LABEL_ENC_MAX  = 20          # max unique values for label-encoding (was 25 — too high)


def _trunc(s, n=SHEET_NAME_LEN):
    """Truncate sheet name to xlsxwriter's 31-char limit."""
    return str(s)[:n]


# ==============================================================================
#  GUI — always-on-top, never shows the root window
# ==============================================================================
root = tk.Tk()
root.withdraw()
try:
    root.attributes("-topmost", True)
except Exception:
    pass


def _dlg():
    w = tk.Toplevel(root)
    w.withdraw()
    try:
        w.attributes("-topmost", True)
    except Exception:
        pass
    return w


def info(title, msg):
    w = _dlg()
    messagebox.showinfo(title, msg, parent=w)
    w.destroy()


def warn(title, msg):
    w = _dlg()
    messagebox.showwarning(title, msg, parent=w)
    w.destroy()


def err(title, msg):
    w = _dlg()
    messagebox.showerror(title, msg, parent=w)
    w.destroy()


def ask_yesno(title, msg):
    w = _dlg()
    result = messagebox.askyesno(title, msg, parent=w)
    w.destroy()
    return result


def ask_file():
    w = _dlg()
    p = filedialog.askopenfilename(
        parent=w,
        title="Select Your Dataset",
        filetypes=[
            ("All Supported Files", "*.csv *.xlsx *.xls *.xlsm *.xlsb *.ods *.txt *.tsv"),
            ("CSV Files",           "*.csv"),
            ("Excel Files",         "*.xlsx *.xls *.xlsm *.xlsb"),
            ("ODS Spreadsheet",     "*.ods"),
            ("Text / TSV",          "*.txt *.tsv"),
        ],
    )
    w.destroy()
    return p or ""


def save_file(initial="Churn_Analytics_Report.xlsx"):
    w = _dlg()
    p = filedialog.asksaveasfilename(
        parent=w,
        title="Save Analytics Report",
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
        initialfile=initial,
    )
    w.destroy()
    return p or ""


# ==============================================================================
#  SAFE SAVE — handles locked / already-open files gracefully
# ==============================================================================
def _timestamped(path):
    base, ext = os.path.splitext(path)
    ts = time.strftime("%H%M%S")
    return f"{base}_{ts}{ext}"


def safe_write_excel(writer_context_fn, save_path):
    """
    Attempt to call writer_context_fn(path) up to MAX_RETRIES times.
    On PermissionError offers auto-rename or manual re-pick.
    Any exception inside writer_context_fn is re-raised after the retry loop
    so the caller's error handler can surface it to the user.
    """
    attempt_path = save_path
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            writer_context_fn(attempt_path)
            return attempt_path
        except PermissionError:
            auto_path = _timestamped(save_path)
            msg = (
                f"The file is currently open or locked:\n  {attempt_path}\n\n"
                f"Yes  →  Auto-save as: {os.path.basename(auto_path)}\n"
                f"No   →  Choose a different location manually"
            )
            use_auto = ask_yesno("File Locked — Save Conflict", msg)
            if use_auto:
                attempt_path = auto_path
            else:
                new_path = save_file(os.path.basename(auto_path))
                if not new_path:
                    raise RuntimeError("Save cancelled by user.")
                attempt_path = new_path
        except OSError as exc:
            if attempt == MAX_RETRIES:
                raise RuntimeError(
                    f"Could not save after {MAX_RETRIES} attempts.\nLast error: {exc}"
                )
            time.sleep(0.5)
    raise RuntimeError("Failed to save — all retry attempts exhausted.")


# ==============================================================================
#  UTILITIES
# ==============================================================================
def norm(s):
    """Normalise a column name: lowercase, alphanumeric only."""
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def clean_num(series):
    """
    Convert a mixed-type series to float.
    Handles currency symbols, thousand-separators, and locale decimal commas.
    """
    s = (
        series.astype(str)
        .str.strip()
        .str.replace(r"[₹$€£¥]", "", regex=True)
        .str.replace(r"\s", "", regex=True)
    )

    def _fix_locale(v):
        # European locale: 1.234,56  →  1234.56
        if re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", v):
            v = v.replace(".", "").replace(",", ".")
        # Thousand-comma with decimal comma: 1,234.56  →  1234.56
        elif re.search(r"\d,\d{3}[.,]", v):
            v = v.replace(",", "")
        else:
            v = v.replace(",", "")
        return v

    s = s.apply(_fix_locale)
    s = s.str.replace(r"[^\d.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")


def find_col(df, *keywords):
    """
    Return the first column whose normalised name contains any keyword.
    Keywords are checked in the order supplied — put the most specific
    keywords FIRST so they win over generic fallbacks.
    """
    cols_norm = {col: norm(col) for col in df.columns}
    for kw in keywords:
        for col, n in cols_norm.items():
            if kw in n:
                return col
    return None


def safe_cut(series, bins=6):
    """
    Bin a numeric series into at most `bins` equal-width buckets.
    Returns a DataFrame with columns [Range, Count] sorted by the left
    endpoint of each interval (not by category code).
    Falls back gracefully for constant or empty series.
    """
    try:
        s = pd.to_numeric(series, errors="coerce").dropna()
        if len(s) == 0 or s.nunique() < 2:
            raise ValueError("insufficient unique values")
        cuts = pd.cut(s, bins=min(bins, s.nunique()), duplicates="drop")
        res = (
            cuts.value_counts()
                .reset_index()
        )
        res.columns = ["Range", "Count"]
        # Sort by the numeric left endpoint of each interval
        res["_left"] = res["Range"].apply(
            lambda iv: iv.left if hasattr(iv, "left") else float("-inf")
        )
        res = (
            res.sort_values("_left")
               .drop(columns="_left")
               .reset_index(drop=True)
        )
        res["Range"] = res["Range"].astype(str)
        return res[res["Count"] > 0].reset_index(drop=True)
    except Exception:
        return pd.DataFrame({
            "Range": ["All Values"],
            "Count": [int(series.notna().sum())],
        })


def safe_float(v, default=0.0):
    """Safely coerce v to float; return default for NaN / inf."""
    try:
        f = float(v)
        return f if np.isfinite(f) else default
    except Exception:
        return default


# ==============================================================================
#  FILE LOADER — two-engine strategy, 6 encodings, 6 separators
# ==============================================================================
def _try_csv(path, enc, sep):
    shared = dict(encoding=enc, dtype=str, on_bad_lines="skip")
    for _engine, kw in [
        ("c",      {"engine": "c",      "low_memory": False, "sep": sep or ","}),
        ("python", {"engine": "python", "sep": sep}),
    ]:
        try:
            df = pd.read_csv(path, **shared, **kw)
            if not df.empty and len(df.columns) >= 2:
                return df
        except Exception:
            pass
    return None


def load_file(path):
    ext    = os.path.splitext(path)[1].lower()
    errors = []

    if ext in (".xlsx", ".xls", ".xlsm", ".xlsb", ".ods"):
        engines = ["openpyxl"]
        if ext == ".xls":
            engines = ["xlrd", "openpyxl"]
        if ext == ".ods":
            engines = ["odf", "openpyxl"]
        engines.append(None)

        for eng in engines:
            try:
                kw = {} if eng is None else {"engine": eng}
                xl = pd.ExcelFile(path, **kw)
                for sheet in xl.sheet_names:
                    try:
                        df = xl.parse(sheet, dtype=str)
                        df = df.dropna(how="all").dropna(axis=1, how="all")
                        if len(df) >= 1 and len(df.columns) >= 2:
                            return df
                    except Exception as exc:
                        errors.append(str(exc))
            except Exception as exc:
                errors.append(str(exc))

        raise RuntimeError(
            "Cannot open Excel/ODS file.\n"
            "Ensure it is not password-protected or corrupted.\n\n"
            f"Details: {'; '.join(errors[:3])}"
        )

    ENCODINGS  = ("utf-8-sig", "utf-8", "latin1", "cp1252", "iso-8859-1", "utf-16")
    SEPARATORS = (None, ",", ";", "\t", "|", ":")

    for enc in ENCODINGS:
        for sep in SEPARATORS:
            try:
                df = _try_csv(path, enc, sep)
                if df is not None:
                    df = df.dropna(how="all").dropna(axis=1, how="all")
                    if len(df) >= 1 and len(df.columns) >= 2:
                        return df
            except Exception as exc:
                errors.append(str(exc))

    raise RuntimeError(
        "Could not parse this file after all known format attempts.\n"
        "Please ensure:\n"
        "  • The file has column headers and tabular data\n"
        "  • It is not password-protected or binary\n"
        "  • It has at least 2 columns and 1 data row\n\n"
        f"Last errors: {'; '.join(list(dict.fromkeys(errors))[:3])}"
    )


# ==============================================================================
#  MAIN
# ==============================================================================
try:
    # ── Welcome ────────────────────────────────────────────────────────────────
    info(
        "Customer Churn Analytics  ·  Step 1 of 2",
        "Welcome to the Universal Customer Churn Analytics Tool.\n\n"
        "This tool will:\n"
        "  ✔  Auto-detect churn, tenure, revenue & ID columns\n"
        "  ✔  Clean and repair bad / missing / mixed data\n"
        "  ✔  Train a Machine Learning model (Random Forest)\n"
        "  ✔  Produce a full executive Excel report with charts\n\n"
        "Supported formats:\n"
        "  CSV · Excel (.xlsx/.xls/.xlsm/.xlsb) · ODS · Text/TSV\n\n"
        "Click OK to select your dataset.",
    )

    file_path = ask_file()
    if not file_path:
        sys.exit(0)

    info("Please Wait…", "Loading and cleaning your dataset.\nThis may take a few seconds.")

    # ── Load ───────────────────────────────────────────────────────────────────
    raw_df = load_file(file_path)
    if raw_df is None or raw_df.empty:
        raise RuntimeError("The file appears to be empty or unreadable.")

    original_rows, original_cols_n = raw_df.shape
    original_df = raw_df.copy()

    # ── Clean ──────────────────────────────────────────────────────────────────
    df = raw_df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]
    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)
    df.drop_duplicates(inplace=True)
    df.reset_index(drop=True, inplace=True)

    if df.empty:
        raise RuntimeError("No usable rows found after cleaning blanks and duplicates.")

    # ── Detect columns (specific keywords FIRST, generic last) ─────────────────
    # BUG FIX 13: more-specific keywords precede generic ones so they take priority
    churn_col   = find_col(df, "churn", "attrition", "churned", "leave",
                               "exit", "cancel", "left")
    tenure_col  = find_col(df, "tenure", "seniority", "duration",
                               "period", "months", "age")
    monthly_col = find_col(df, "monthlycharge", "monthlyfee", "monthly",
                               "charge", "rate", "fee", "price")
    total_col   = find_col(df, "totalcharge", "totalrevenue", "lifetimevalue",
                               "total", "revenue", "bill", "spend",
                               "amount", "sales")
    id_col      = find_col(df, "customerid", "custid", "userid",
                               "accountid", "clientid", "id")

    # ── Numeric conversion ─────────────────────────────────────────────────────
    skip_numeric = {c for c in [churn_col, id_col] if c}
    for col in df.columns:
        if col in skip_numeric:
            continue
        temp = clean_num(df[col])
        if temp.notna().sum() / max(len(temp), 1) >= 0.40:
            df[col] = temp

    # BUG FIX 14: cap label-encoding at LABEL_ENC_MAX unique values to avoid
    # encoding near-text columns that happen to have ≤25 unique values.
    for col in df.select_dtypes(include="object").columns:
        if col in skip_numeric:
            continue
        try:
            nu = df[col].nunique()
            if 2 <= nu <= LABEL_ENC_MAX:
                le = LabelEncoder()
                df[col] = le.fit_transform(df[col].fillna("__missing__").astype(str))
            else:
                df[col] = np.nan
        except Exception:
            df[col] = np.nan

    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()

    # ── Core columns ───────────────────────────────────────────────────────────
    def pick(col_name, fallback_idx, default_val):
        """
        Extract a numeric series for a key column.
        Priority: detected column → numeric fallback → constant default.
        BUG FIX 3: default_val is now always a finite float, never np.nan.
        """
        if col_name and col_name in df.columns:
            s = clean_num(df[col_name])
        elif len(numeric_cols) > fallback_idx:
            s = df[numeric_cols[fallback_idx]].copy()
        else:
            s = pd.Series([float(default_val)] * len(df), dtype=float)
        s = pd.to_numeric(s, errors="coerce").clip(lower=0)
        med  = s.median()
        fill = med if (pd.notna(med) and np.isfinite(med)) else float(default_val)
        return s.fillna(fill)

    df["tenure"]         = pick(tenure_col,  0, 12.0)
    df["MonthlyCharges"] = pick(monthly_col, 1, 500.0)
    # BUG FIX 3: pass 6000.0 (finite) instead of np.nan as default
    df["TotalCharges"]   = pick(total_col,   2, 6000.0)

    # Guard: TotalCharges < MonthlyCharges is logically impossible
    bad_total = df["TotalCharges"] < df["MonthlyCharges"]
    df.loc[bad_total, "TotalCharges"] = (
        df.loc[bad_total, "tenure"] * df.loc[bad_total, "MonthlyCharges"]
    )
    # Final fill for any remaining NaN (should be none, but belt-and-suspenders)
    df["TotalCharges"] = df["TotalCharges"].fillna(
        df["tenure"] * df["MonthlyCharges"]
    ).clip(lower=0)

    YES = {
        "yes", "y", "true", "1", "1.0", "churned", "left", "cancelled",
        "cancel", "exit", "quit", "inactive", "lost", "gone", "departed",
    }
    NO = {
        "no", "n", "false", "0", "0.0", "active", "retained", "stay",
        "stayed", "current", "existing", "ongoing", "present", "alive", "good",
    }

    # ── Feature engineering ────────────────────────────────────────────────────
    df["AvgMonthlySpend"] = (
        df["TotalCharges"] / (df["tenure"] + 1)
    ).clip(lower=0).fillna(0)

    df["ValueScore"] = np.log1p(df["TotalCharges"]).clip(lower=0).fillna(0)

    df["LoyaltyScore"] = (
        df["tenure"] * (df["TotalCharges"] / (df["MonthlyCharges"] + 1))
    ).clip(lower=0).fillna(0)

    df["SpendVariance"] = (
        (df["MonthlyCharges"] - df["AvgMonthlySpend"]) / (df["AvgMonthlySpend"] + 1)
    ).abs().clip(lower=0).fillna(0)

    # ── Churn mapping ──────────────────────────────────────────────────────────
    churn_source     = "Synthetic (no churn column found)"
    _synthetic_churn = True      # BUG FIX 16: always set explicitly in every branch

    if churn_col:
        raw_c  = df[churn_col].astype(str).str.strip().str.lower()
        mapped = raw_c.map(lambda v: 1 if v in YES else (0 if v in NO else np.nan))
        if mapped.notna().sum() == 0:
            mapped = pd.to_numeric(df[churn_col], errors="coerce").round().clip(0, 1)
        null_frac = mapped.isna().mean()
        if null_frac > 0.6:
            df["Churn"]      = np.where(
                df["tenure"] < df["tenure"].median(), 1, 0).astype(int)
            churn_source     = f"Synthetic ('{churn_col}' had {null_frac*100:.0f}% unmapped)"
            _synthetic_churn = True
        else:
            mode_v           = int(mapped.mode().iloc[0]) if mapped.notna().any() else 0
            df["Churn"]      = mapped.fillna(mode_v).astype(int)
            churn_source     = f"Detected: '{churn_col}'"
            _synthetic_churn = False
    else:
        risk_score = (
            (1 - df["tenure"].rank(pct=True, method="average")) * 0.40
            + df["MonthlyCharges"].rank(pct=True, method="average") * 0.30
            + df["SpendVariance"].rank(pct=True, method="average") * 0.20
            + (1 - df["LoyaltyScore"].rank(pct=True, method="average")) * 0.10
        )
        threshold        = risk_score.quantile(0.65)
        df["Churn"]      = (risk_score >= threshold).astype(int)
        churn_source     = "Synthetic — percentile risk score (top 35% flagged)"
        _synthetic_churn = True   # explicit set

    FEATURES = [
        "tenure", "MonthlyCharges", "TotalCharges",
        "AvgMonthlySpend", "ValueScore", "LoyaltyScore", "SpendVariance",
    ]

    X_raw = df[FEATURES].copy()
    y     = df["Churn"].copy()

    imputer = SimpleImputer(strategy="median")
    X = pd.DataFrame(imputer.fit_transform(X_raw), columns=FEATURES)

    # ── Heuristic baseline (used when ML cannot run) ───────────────────────────
    heuristic = (
        X["tenure"].rank(pct=True).rsub(1) * 0.55
        + X["MonthlyCharges"].rank(pct=True) * 0.25
        + X["SpendVariance"].rank(pct=True) * 0.20
    ).clip(0, 1)
    df["Churn_Prob"] = heuristic.values

    # ── ML model ───────────────────────────────────────────────────────────────
    accuracy   = None
    auc_score  = None
    model_name = (
        "Heuristic (synthetic labels — metrics suppressed)"
        if _synthetic_churn
        else "Heuristic"
    )

    if len(df) >= 20 and y.nunique() > 1:
        for Clf, name, kw in [
            (
                RandomForestClassifier,
                "Random Forest",
                dict(n_estimators=300, max_depth=12,
                     class_weight="balanced", random_state=42, n_jobs=-1),
            ),
            (
                GradientBoostingClassifier,
                "Gradient Boosting",
                dict(n_estimators=150, max_depth=5, random_state=42),
            ),
        ]:
            try:
                min_class   = y.value_counts().min()
                do_stratify = min_class >= max(2, int(len(df) * 0.20 * 0.10))
                Xt, Xe, yt, ye = train_test_split(
                    X, y,
                    test_size=0.20,
                    random_state=42,
                    stratify=(y if do_stratify else None),
                )
                clf = Clf(**kw)
                clf.fit(Xt, yt)
                prob_e = clf.predict_proba(Xe)[:, 1]

                # BUG FIX: only record metrics when labels are REAL.
                # Synthetic labels produce self-referential metrics that are
                # meaningless as model quality indicators.
                if not _synthetic_churn:
                    accuracy  = accuracy_score(ye, clf.predict(Xe))
                    auc_score = (
                        roc_auc_score(ye, prob_e) if ye.nunique() > 1 else None
                    )

                df["Churn_Prob"] = clf.predict_proba(X)[:, 1]
                model_name = (
                    f"{name} (Synthetic Labels — metrics suppressed)"
                    if _synthetic_churn
                    else name
                )
                break
            except Exception:
                continue

    df["Churn_Prob"] = df["Churn_Prob"].clip(0, 1).round(4)

    # ── CLV — retention-weighted historical proxy ──────────────────────────────
    # Formula : CLV_proxy = TotalCharges × (1 – Churn_Prob)
    # Interpretation: the portion of a customer's historical spend we expect
    # to retain, weighted by their probability of staying.  This is NOT a
    # survival-model or discounted-cash-flow CLV; it is a relative ranking
    # metric only.  Documented in the Data Quality audit sheet.
    try:
        tc_clean = (
            df["TotalCharges"]
            .replace([np.inf, -np.inf], np.nan)
            .clip(lower=0)
            .fillna(0)
        )
        df["Predicted_CLV"] = (tc_clean * (1 - df["Churn_Prob"])).clip(lower=0)
    except Exception:
        df["Predicted_CLV"] = df["TotalCharges"].clip(lower=0).fillna(0)

    # ── Segmentation ───────────────────────────────────────────────────────────
    hi_risk = df["Churn_Prob"] >= 0.60
    hi_clv  = df["Predicted_CLV"] >= df["Predicted_CLV"].median()

    df["Segment"] = np.select(
        [hi_risk & hi_clv, hi_risk & ~hi_clv, ~hi_risk & hi_clv],
        ["High Risk – High Value",
         "High Risk – Low Value",
         "Low Risk – High Value"],
        default="Low Risk – Low Value",
    )
    df["Segment"]   = df["Segment"].fillna("Low Risk – Low Value")
    df["Risk_Tier"] = pd.cut(
        df["Churn_Prob"],
        bins=[0, 0.30, 0.60, 1.001],
        labels=["Low Risk (0-30%)", "Medium Risk (30-60%)", "High Risk (60-100%)"],
        include_lowest=True,
    ).astype(str).fillna("Low Risk (0-30%)")

    # ── KPIs ───────────────────────────────────────────────────────────────────
    total_cust     = max(len(df), 1)
    churn_rate     = round(safe_float(df["Churn"].mean()) * 100, 2)
    # BUG FIX 12: explicit clamp prevents 99.99 when churn is 0.0
    retention_rate = round(max(0.0, min(100.0, 100.0 - churn_rate)), 2)
    avg_clv        = round(safe_float(df["Predicted_CLV"].mean()), 0)
    avg_tenure     = round(safe_float(df["tenure"].mean()), 1)
    avg_monthly    = round(safe_float(df["MonthlyCharges"].mean()), 0)
    high_risk_n    = int((df["Churn_Prob"] >= 0.60).sum())
    high_risk_pct  = round(high_risk_n / total_cust * 100, 1)
    med_risk_n     = int(((df["Churn_Prob"] >= 0.30) & (df["Churn_Prob"] < 0.60)).sum())
    low_risk_n     = int((df["Churn_Prob"] < 0.30).sum())
    total_clv      = round(safe_float(df["Predicted_CLV"].sum()), 0)

    # BUG FIX 19: Revenue-at-risk uses MonthlyCharges×12 for high-risk customers
    # as a forward-looking proxy (annualised run-rate), not TotalCharges
    # (which is a backward-looking historical total).
    hr_mask           = df["Churn_Prob"] >= 0.60
    revenue_at_risk   = round(safe_float(
        (df.loc[hr_mask, "MonthlyCharges"] * 12).sum()), 0)
    # Retain TotalCharges sum for contextual display in audit
    tc_at_risk_hist   = round(safe_float(
        df.loc[hr_mask, "TotalCharges"].sum()), 0)

    # Display strings — suppressed when synthetic
    if _synthetic_churn:
        acc_str = "N/A (synthetic labels)"
        auc_str = "N/A (synthetic labels)"
    else:
        acc_str = f"{round(accuracy * 100, 2)}%" if accuracy  is not None else "N/A"
        auc_str = f"{round(auc_score, 4)}"       if auc_score is not None else "N/A"

    info(
        "Dataset Ready  ✓",
        f"File          : {os.path.basename(file_path)}\n"
        f"Original size : {original_rows:,} rows  ×  {original_cols_n} columns\n"
        f"After cleaning: {total_cust:,} rows\n\n"
        f"Churn column  : {churn_source}\n"
        f"ML Model      : {model_name}\n"
        f"Accuracy      : {acc_str}     AUC-ROC: {auc_str}\n\n"
        f"── KEY FINDINGS ──────────────────────\n"
        f"  Churn Rate              : {churn_rate}%\n"
        f"  High-Risk Customers     : {high_risk_n:,}  ({high_risk_pct}%)\n"
        f"  Annualised Risk Revenue : ₹{revenue_at_risk:,.0f}\n"
        f"──────────────────────────────────────\n\n"
        "Click OK to choose where to save the report.",
    )

    save_path = save_file()
    if not save_path:
        sys.exit(0)

    # ==========================================================================
    #  BUILD EXCEL REPORT
    # ==========================================================================
    def build_report(out_path):
        with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
            wb = writer.book

            # ── Colour palette ─────────────────────────────────────────────────
            C = {
                "navy":    "#0B1F3A",  "navy2":   "#1A3A5C",
                "teal":    "#007C71",  "teal2":   "#00A896",
                "slate":   "#2B4570",  "steel":   "#3D6B99",
                "red":     "#C0392B",  "red2":    "#E74C3C",
                "orange":  "#D35400",  "amber":   "#E67E22",
                "green":   "#1E8449",  "green2":  "#27AE60",
                "blue":    "#1565C0",  "blue2":   "#1976D2",
                "purple":  "#6C3483",  "violet":  "#8E44AD",
                "gold":    "#B7950B",  "yellow":  "#F39C12",
                "lgrey":   "#F0F4F8",  "mgrey":   "#C8D0DC",
                "dgrey":   "#5D6D7E",  "white":   "#FFFFFF",
                "warn_bg": "#FFF3CD",  "warn_fc": "#856404",
            }

            CHART8 = ["#007C71", "#1565C0", "#E67E22", "#C0392B",
                      "#6C3483", "#1E8449", "#D35400", "#3D6B99"]

            # ── Format factory ──────────────────────────────────────────────────
            def MF(bold=False, italic=False, sz=10, fc=C["navy"], bg=None,
                   align="center", valign="vcenter", wrap=False,
                   border=1, bc=C["mgrey"], num_fmt=None, left=False):
                d = {
                    "font_name":  "Calibri",
                    "font_size":  sz,
                    "bold":       bold,
                    "italic":     italic,
                    "font_color": fc,
                    "valign":     valign,
                    "align":      "left" if left else align,
                    "text_wrap":  wrap,
                }
                if bg:      d["bg_color"]   = bg
                if border:  d.update({"border": border, "border_color": bc})
                if num_fmt: d["num_format"] = num_fmt
                return wb.add_format(d)

            # ── Format registry ─────────────────────────────────────────────────
            F = {
                "title":       MF(bold=True,   sz=18, fc=C["white"],   bg=C["navy"],   border=0),
                "subtitle":    MF(italic=True, sz=10, fc=C["slate"],   bg=C["lgrey"],  border=0),
                "warn_banner": MF(bold=True,   sz=10, fc=C["warn_fc"], bg=C["warn_bg"],border=1,
                                  bc=C["amber"], wrap=True),
                "col_hdr":     MF(bold=True,   sz=10, fc=C["white"],   bg=C["navy2"]),
                "col_hdr_t":   MF(bold=True,   sz=10, fc=C["white"],   bg=C["teal"]),
                "col_hdr_s":   MF(bold=True,   sz=10, fc=C["white"],   bg=C["slate"]),
                "col_hdr_p":   MF(bold=True,   sz=10, fc=C["white"],   bg=C["purple"]),
                "sec_banner":  MF(bold=True,   sz=10, fc=C["white"],   bg=C["teal2"]),

                "kpi_lbl_b":  MF(bold=True, sz=9, fc=C["white"], bg=C["navy2"],  border=0),
                "kpi_lbl_r":  MF(bold=True, sz=9, fc=C["white"], bg=C["red"],    border=0),
                "kpi_lbl_g":  MF(bold=True, sz=9, fc=C["white"], bg=C["teal"],   border=0),
                "kpi_lbl_a":  MF(bold=True, sz=9, fc=C["white"], bg=C["amber"],  border=0),
                "kpi_lbl_p":  MF(bold=True, sz=9, fc=C["white"], bg=C["purple"], border=0),
                "kpi_lbl_sl": MF(bold=True, sz=9, fc=C["white"], bg=C["slate"],  border=0),
                "kpi_lbl_w":  MF(bold=True, sz=9, fc=C["navy"],  bg=C["warn_bg"],border=0),

                "kpi_val":    MF(bold=True, sz=16, fc=C["navy"],   bg=C["white"], bc=C["mgrey"]),
                "kpi_red":    MF(bold=True, sz=16, fc=C["red"],    bg=C["white"], bc=C["mgrey"]),
                "kpi_green":  MF(bold=True, sz=16, fc=C["green"],  bg=C["white"], bc=C["mgrey"]),
                "kpi_amber":  MF(bold=True, sz=16, fc=C["orange"], bg=C["white"], bc=C["mgrey"]),
                "kpi_purple": MF(bold=True, sz=16, fc=C["purple"], bg=C["white"], bc=C["mgrey"]),
                "kpi_warn":   MF(bold=True, sz=10, fc=C["warn_fc"],bg=C["warn_bg"],
                                  bc=C["amber"], wrap=True),

                "cell":       MF(sz=10, fc=C["navy"], bg=C["white"], bc=C["mgrey"]),
                "cell_alt":   MF(sz=10, fc=C["navy"], bg=C["lgrey"], bc=C["mgrey"]),
                "cell_l":     MF(sz=10, fc=C["navy"], bg=C["white"], bc=C["mgrey"], left=True),
                "cell_l_alt": MF(sz=10, fc=C["navy"], bg=C["lgrey"], bc=C["mgrey"], left=True),

                # Typed numeric formats
                "num":      MF(sz=10, fc=C["navy"], bg=C["white"], bc=C["mgrey"], num_fmt="#,##0"),
                "num_alt":  MF(sz=10, fc=C["navy"], bg=C["lgrey"], bc=C["mgrey"], num_fmt="#,##0"),
                "dec2":     MF(sz=10, fc=C["navy"], bg=C["white"], bc=C["mgrey"], num_fmt="#,##0.00"),
                "dec2_alt": MF(sz=10, fc=C["navy"], bg=C["lgrey"], bc=C["mgrey"], num_fmt="#,##0.00"),
                "dec4":     MF(sz=10, fc=C["navy"], bg=C["white"], bc=C["mgrey"], num_fmt="0.0000"),
                "dec4_alt": MF(sz=10, fc=C["navy"], bg=C["lgrey"], bc=C["mgrey"], num_fmt="0.0000"),
                "pct2":     MF(sz=10, fc=C["navy"], bg=C["white"], bc=C["mgrey"], num_fmt="0.00%"),
                "pct2_alt": MF(sz=10, fc=C["navy"], bg=C["lgrey"], bc=C["mgrey"], num_fmt="0.00%"),

                # Conditional churn-rate colours
                "c_red":   MF(bold=True, sz=10, fc=C["red2"],   bg="#FDECEA",
                               bc=C["mgrey"], num_fmt="0.00%"),
                "c_amber": MF(bold=True, sz=10, fc=C["orange"], bg="#FDF3E7",
                               bc=C["mgrey"], num_fmt="0.00%"),
                "c_green": MF(bold=True, sz=10, fc=C["green"],  bg="#EAFAF1",
                               bc=C["mgrey"], num_fmt="0.00%"),
            }

            # ── write_table ─────────────────────────────────────────────────────
            # BUG FIX 2: set_row called once per row, not inside the column loop.
            def write_table(ws, df_t, r0, c0=0, hdr_fmt="col_hdr",
                            col_fmts=None, row_height=20):
                ws.set_row(r0, 22)   # header row height
                for ci, col in enumerate(df_t.columns):
                    ws.write(r0, c0 + ci, col, F[hdr_fmt])
                df_t = df_t.reset_index(drop=True)
                for ri in range(len(df_t)):
                    ws.set_row(r0 + 1 + ri, row_height)  # called once per data row
                    alt = (ri % 2 == 1)
                    for ci, col in enumerate(df_t.columns):
                        val = df_t.iloc[ri, ci]
                        if pd.isna(val) or str(val) == "nan":
                            val = ""
                        if col_fmts and col in col_fmts:
                            fk = col_fmts[col]
                            if callable(fk):
                                fmt = fk(val, alt)
                            else:
                                alt_key = fk + "_alt"
                                fmt = F[alt_key if (alt and alt_key in F) else fk]
                        else:
                            fmt = F["cell_alt" if alt else "cell"]
                        ws.write(r0 + 1 + ri, c0 + ci, val, fmt)

            # ── auto column widths ──────────────────────────────────────────────
            def auto_col_widths(df_src, min_w=10, max_w=60, sample=500):
                widths = []
                sample_df = df_src.iloc[:sample] if len(df_src) > sample else df_src
                for col in df_src.columns:
                    hdr_w  = len(str(col)) + 2
                    vals   = sample_df[col].astype(str).str.len()
                    data_w = int(vals.max()) + 2 if len(vals) else 0
                    widths.append(max(min_w, min(max(hdr_w, data_w), max_w)))
                return widths

            # ── data_sheet helper (chart-source sheets) ─────────────────────────
            # BUG FIX 5: header row (row 4) height explicitly set here.
            def data_sheet(name, df_t, title_text, subtitle_text,
                           hdr_color="col_hdr_t", min_w=18, max_w=55):
                sname = _trunc(name)
                ws    = wb.add_worksheet(sname)
                ws.hide_gridlines(2)
                ws.set_zoom(90)
                col_ws = auto_col_widths(df_t, min_w=min_w, max_w=max_w)
                for ci, w in enumerate(col_ws):
                    ws.set_column(ci, ci, w)
                ws.set_row(0, 5)
                ws.set_row(1, 38)
                ws.set_row(2, 18)
                ws.set_row(3, 5)
                ws.set_row(4, 22)   # BUG FIX 5: header row in data sheet
                last_col = max(len(df_t.columns) - 1, 1)
                ws.merge_range(1, 0, 1, last_col, title_text,    F["title"])
                ws.merge_range(2, 0, 2, last_col, subtitle_text, F["subtitle"])
                write_table(ws, df_t, 4, c0=0, hdr_fmt=hdr_color, row_height=22)
                return sname, len(df_t)

            # ── chart helper ────────────────────────────────────────────────────
            # BUG FIX 1: insert_chart takes a string cell reference.
            # We accept a string directly; callers compute the correct cell string.
            def add_chart(chart_type, src_sheet, nrows, title,
                          ws_dest, cell_ref,
                          fill_color=None, fill_colors=None,
                          w=500, h=295, show_legend=False):
                if nrows <= 0:
                    return
                ch = wb.add_chart({"type": chart_type})
                series = {
                    "categories":  [src_sheet, 5, 0, 4 + nrows, 0],
                    "values":      [src_sheet, 5, 1, 4 + nrows, 1],
                    "data_labels": {
                        "value": True,
                        "font":  {"bold": True, "size": 9, "color": C["navy"]},
                    },
                }
                if chart_type in ("column", "bar"):
                    series["gap"] = 60

                if fill_colors:
                    series["points"] = [
                        {"fill": {"color": fill_colors[i % len(fill_colors)]}}
                        for i in range(nrows)
                    ]
                elif fill_color:
                    series["fill"] = {"color": fill_color}

                ch.add_series(series)
                ch.set_title({
                    "name":      title,
                    "name_font": {"bold": True, "size": 11, "color": C["navy"]},
                })
                ch.set_legend(
                    {"position": "bottom"} if show_legend else {"none": True}
                )
                ch.set_chartarea({
                    "border": {"color": C["mgrey"]},
                    "fill":   {"color": C["white"]},
                })
                ch.set_plotarea({"fill": {"color": C["lgrey"]}})
                ch.set_style(2)
                ch.set_size({"width": w, "height": h})
                ws_dest.insert_chart(cell_ref, ch, {"x_offset": 5, "y_offset": 5})

            # ── write_full_sheet ────────────────────────────────────────────────
            # BUG FIX 6: per-column numeric formats applied in processed data sheet.
            # BUG FIX 11: freeze_panes applied to all full data sheets.
            PROC_NUM_FMTS = {
                "tenure":          "#,##0.0",
                "MonthlyCharges":  "#,##0.00",
                "TotalCharges":    "#,##0.00",
                "AvgMonthlySpend": "#,##0.00",
                "ValueScore":      "0.0000",
                "LoyaltyScore":    "#,##0.00",
                "SpendVariance":   "0.0000",
                "Churn":           "0",
                "Churn_Prob":      "0.00%",
                "Predicted_CLV":   "#,##0.00",
            }

            def write_full_sheet(ws, df_src, hdr_fmt_key,
                                 zoom=90, row_h=18, hdr_h=22,
                                 min_w=10, max_w=60,
                                 apply_num_fmt=False):
                ws.hide_gridlines(2)
                ws.set_zoom(zoom)
                col_widths = auto_col_widths(df_src, min_w=min_w, max_w=max_w)
                for ci, w in enumerate(col_widths):
                    ws.set_column(ci, ci, w)
                ws.set_row(0, hdr_h)
                for ci, col in enumerate(df_src.columns):
                    ws.write(0, ci, str(col), F[hdr_fmt_key])

                for ri in range(len(df_src)):
                    ws.set_row(ri + 1, row_h)
                    alt = (ri % 2 == 1)
                    for ci, col in enumerate(df_src.columns):
                        val = df_src.iloc[ri, ci]
                        if pd.isna(val) or str(val) == "nan":
                            val = ""
                        if apply_num_fmt and col in PROC_NUM_FMTS:
                            fmt = MF(sz=10, fc=C["navy"],
                                     bg=C["lgrey"] if alt else C["white"],
                                     bc=C["mgrey"],
                                     num_fmt=PROC_NUM_FMTS[col])
                        else:
                            fmt = MF(sz=10, fc=C["navy"],
                                     bg=C["lgrey"] if alt else C["white"],
                                     bc=C["mgrey"])
                        ws.write(ri + 1, ci, val, fmt)

                ws.autofilter(0, 0, len(df_src), len(df_src.columns) - 1)
                ws.freeze_panes(1, 0)   # BUG FIX 11: consistent for all full sheets

            # ── Distribution data ───────────────────────────────────────────────
            seg_data = df["Segment"].value_counts().reset_index()
            seg_data.columns = ["Segment", "Count"]

            risk_order = ["Low Risk (0-30%)", "Medium Risk (30-60%)", "High Risk (60-100%)"]
            risk_data  = df["Risk_Tier"].value_counts().reset_index()
            risk_data.columns = ["Risk Tier", "Count"]
            risk_data["Risk Tier"] = pd.Categorical(
                risk_data["Risk Tier"], categories=risk_order, ordered=True)
            risk_data = risk_data.sort_values("Risk Tier").reset_index(drop=True)

            churn_dist   = safe_cut(df["Churn_Prob"],      6)
            clv_dist     = safe_cut(df["Predicted_CLV"],   6)
            spend_dist   = safe_cut(df["AvgMonthlySpend"], 6)
            tenure_dist  = safe_cut(df["tenure"],          6)
            monthly_dist = safe_cut(df["MonthlyCharges"],  6)

            # ── Chart-source sheets ─────────────────────────────────────────────
            sn_seg,  seg_n  = data_sheet("Data – Segments",   seg_data,
                "Customer Segment Distribution",
                "Count of customers in each churn-risk vs CLV segment.",
                "col_hdr_t")
            sn_risk, risk_n = data_sheet("Data – Risk Tiers", risk_data,
                "Risk Tier Distribution",
                "Low / Medium / High churn probability buckets.",
                "col_hdr_p")
            sn_chd,  chd_n  = data_sheet("Data – Churn Prob", churn_dist,
                "Churn Probability Buckets",
                "Distribution of predicted churn probability (0 = safe, 1 = leaving).",
                "col_hdr_t")
            sn_clv,  clvd_n = data_sheet("Data – CLV",        clv_dist,
                "Customer Lifetime Value Buckets",
                "Distribution of predicted CLV (retention-weighted historical proxy).",
                "col_hdr_s")
            sn_spd,  spd_n  = data_sheet("Data – Spend",      spend_dist,
                "Average Monthly Spend Buckets",
                "Distribution of average monthly spend per customer.",
                "col_hdr_t")
            sn_ten,  tend_n = data_sheet("Data – Tenure",     tenure_dist,
                "Tenure Distribution (months)",
                "How long customers have been with the company.",
                "col_hdr_s")
            sn_mnd,  mnd_n  = data_sheet("Data – Monthly",    monthly_dist,
                "Monthly Charges Distribution",
                "Distribution of current monthly charge amounts.",
                "col_hdr_p")

            # ==================================================================
            #  SHEET 1 — DASHBOARD
            # ==================================================================
            dash = wb.add_worksheet(_trunc("Dashboard"))
            dash.hide_gridlines(2)
            dash.set_zoom(80)

            CARD_W  = 18
            MARGIN  = 1.5
            N_CARDS = 7

            dash.set_column(0, 0, MARGIN)
            for c in range(1, N_CARDS * 2 + 6):
                dash.set_column(c, c, CARD_W)

            # Row heights
            KPI_LBL_ROW1  = 4
            KPI_VAL_ROW1  = 5
            KPI_LBL_ROW2  = 7
            KPI_VAL_ROW2  = 8
            # BUG FIX 10: warn_row_start derived from KPI layout, not hardcoded
            WARN_ROW      = 10

            for r, h in {0: 5, 1: 46, 2: 20, 3: 8,
                         KPI_LBL_ROW1: 18, KPI_VAL_ROW1: 48, 6: 8,
                         KPI_LBL_ROW2: 18, KPI_VAL_ROW2: 48, 9: 12}.items():
                dash.set_row(r, h)

            # BUG FIX 17: subtitle span computed from N_CARDS
            TITLE_LAST_COL = N_CARDS * 2
            dash.merge_range(1, 1, 1, TITLE_LAST_COL,
                "  CUSTOMER CHURN ANALYTICS  ·  EXECUTIVE DASHBOARD",
                F["title"])
            dash.merge_range(2, 1, 2, TITLE_LAST_COL,
                f"File: {os.path.basename(file_path)}   |   "
                f"Model: {model_name}   |   "
                f"Accuracy: {acc_str}   |   AUC-ROC: {auc_str}   |   "
                f"Customers Analysed: {total_cust:,}",
                F["subtitle"])

            # ── Synthetic-labels warning banner ────────────────────────────────
            if _synthetic_churn:
                dash.set_row(WARN_ROW, 44)
                dash.merge_range(
                    WARN_ROW, 1, WARN_ROW, TITLE_LAST_COL,
                    "⚠  SYNTHETIC LABELS NOTICE:  No real churn column was found in "
                    "this dataset.  Churn labels were generated using a percentile "
                    "risk-score formula.  ML accuracy and AUC-ROC are suppressed "
                    "because they would only measure how well the model learned our "
                    "own synthetic formula — not true churn.  Churn probabilities "
                    "and segmentation are valid for relative risk ranking.",
                    F["warn_banner"],
                )
                chart_anchor_row = WARN_ROW + 2
            else:
                chart_anchor_row = WARN_ROW

            # ── KPI card helper ─────────────────────────────────────────────────
            def kpi_card(ws, label, value, lbl_row, val_row, card_idx,
                         lbl_fmt="kpi_lbl_b", val_fmt="kpi_val"):
                c_start = 1 + card_idx * 2
                c_end   = c_start + 1
                ws.merge_range(lbl_row, c_start, lbl_row, c_end, label, F[lbl_fmt])
                ws.merge_range(val_row, c_start, val_row, c_end, value, F[val_fmt])

            # KPI Row 1
            for i, (lbl, val, lf, vf) in enumerate([
                ("TOTAL CUSTOMERS",      f"{total_cust:,}",
                 "kpi_lbl_b",  "kpi_val"),
                ("CHURN RATE",           f"{churn_rate}%",
                 "kpi_lbl_r",  "kpi_red"),
                ("RETENTION RATE",       f"{retention_rate}%",
                 "kpi_lbl_g",  "kpi_green"),
                ("HIGH RISK  (≥60%)",    f"{high_risk_n:,} ({high_risk_pct}%)",
                 "kpi_lbl_r",  "kpi_red"),
                ("MEDIUM RISK (30-60%)", f"{med_risk_n:,}",
                 "kpi_lbl_a",  "kpi_amber"),
                ("LOW RISK  (<30%)",     f"{low_risk_n:,}",
                 "kpi_lbl_g",  "kpi_green"),
                ("AVG TENURE (months)",  f"{avg_tenure}",
                 "kpi_lbl_sl", "kpi_val"),
            ]):
                kpi_card(dash, lbl, val, KPI_LBL_ROW1, KPI_VAL_ROW1, i, lf, vf)

            # KPI Row 2
            # BUG FIX 9: model_name displayed in full (no [:28] truncation);
            # long names wrap inside the cell via kpi_warn's text_wrap=True.
            if _synthetic_churn:
                acc_lbl_fmt, acc_val_fmt = "kpi_lbl_w", "kpi_warn"
                auc_lbl_fmt, auc_val_fmt = "kpi_lbl_w", "kpi_warn"
                acc_display = "N/A — Synthetic"
                auc_display = "N/A — Synthetic"
            else:
                acc_lbl_fmt, acc_val_fmt = "kpi_lbl_g", "kpi_green"
                auc_lbl_fmt, auc_val_fmt = "kpi_lbl_p", "kpi_purple"
                acc_display = acc_str
                auc_display = auc_str

            # BUG FIX 18: revenue-at-risk label clarified as annualised run-rate
            for i, (lbl, val, lf, vf) in enumerate([
                ("AVG CUSTOMER CLV",          f"Rs.{avg_clv:,.0f}",
                 "kpi_lbl_g",  "kpi_green"),
                ("ANNUALISED RISK REVENUE",   f"Rs.{revenue_at_risk:,.0f}",
                 "kpi_lbl_r",  "kpi_red"),
                ("TOTAL CLV PORTFOLIO",       f"Rs.{total_clv:,.0f}",
                 "kpi_lbl_p",  "kpi_purple"),
                ("AVG MONTHLY CHARGE",        f"Rs.{avg_monthly:,.0f}",
                 "kpi_lbl_b",  "kpi_val"),
                ("ML MODEL",                  model_name,
                 "kpi_lbl_sl", "kpi_val"),
                ("MODEL ACCURACY",            acc_display,
                 acc_lbl_fmt,  acc_val_fmt),
                ("AUC-ROC SCORE",             auc_display,
                 auc_lbl_fmt,  auc_val_fmt),
            ]):
                kpi_card(dash, lbl, val, KPI_LBL_ROW2, KPI_VAL_ROW2, i, lf, vf)

            # ── Dashboard charts ────────────────────────────────────────────────
            # BUG FIX 1: cell references computed correctly.
            # chart_anchor_row is 0-based; Excel row() in insert_chart is also
            # 0-based.  We add 1 so charts start one row below the anchor.
            r1 = chart_anchor_row + 1
            r2 = r1 + 17
            r3 = r2 + 17

            def row_cell(r, col_letter):
                """Convert 0-based row index + column letter to Excel cell string."""
                return f"{col_letter}{r + 1}"

            add_chart("column",   sn_seg,  seg_n,
                      "Customer Segments by Count",
                      dash, row_cell(r1, "B"),
                      fill_colors=CHART8, w=540, h=310)
            add_chart("doughnut", sn_risk, risk_n,
                      "Risk Tier Breakdown",
                      dash, row_cell(r1, "J"),
                      fill_colors=["#27AE60", "#E67E22", "#C0392B"],
                      w=480, h=310, show_legend=True)

            add_chart("column", sn_chd, chd_n,
                      "Churn Probability Distribution  (0 = Safe  ·  1 = Leaving)",
                      dash, row_cell(r2, "B"),
                      fill_colors=["#1E8449", "#52BE80", "#F4D03F",
                                   "#E67E22", "#E74C3C", "#C0392B"],
                      w=540, h=310)
            add_chart("bar", sn_clv, clvd_n,
                      "Predicted CLV  (retention-weighted historical proxy)",
                      dash, row_cell(r2, "J"),
                      fill_colors=CHART8, w=480, h=310)

            add_chart("column", sn_ten, tend_n,
                      "Customer Tenure Distribution (months)",
                      dash, row_cell(r3, "B"),
                      fill_colors=["#1565C0", "#1976D2", "#42A5F5",
                                   "#64B5F6", "#90CAF9", "#BBDEFB"],
                      w=540, h=310)
            add_chart("column", sn_mnd, mnd_n,
                      "Monthly Charges Distribution",
                      dash, row_cell(r3, "J"),
                      fill_colors=["#6C3483", "#8E44AD", "#A569BD",
                                   "#C39BD3", "#D7BDE2", "#E8DAEF"],
                      w=480, h=310)

            # ==================================================================
            #  SHEET 2 — SEGMENT SUMMARY
            # ==================================================================
            ws2 = wb.add_worksheet(_trunc("Segment Summary"))
            ws2.hide_gridlines(2)
            ws2.set_zoom(95)
            ws2.set_column(0, 0,  1.5)
            ws2.set_column(1, 1,  32)
            ws2.set_column(2, 9,  19)
            for r, h in [(0, 5), (1, 46), (2, 20), (3, 8)]:
                ws2.set_row(r, h)

            # BUG FIX 17: subtitle span uses TITLE_LAST_COL
            ws2.merge_range(1, 1, 1, 9,
                "SEGMENT PERFORMANCE SUMMARY  ·  Churn Analytics", F["title"])
            ws2.merge_range(2, 1, 2, 9,
                "One row per customer segment.  "
                "Churn Rate is colour-coded: RED >60%  AMBER 30-60%  GREEN <30%",
                F["subtitle"])

            summary = df.groupby("Segment", as_index=False).agg(
                Customers     =("Churn",          "count"),
                Churned       =("Churn",          "sum"),
                Churn_Rate    =("Churn",          "mean"),
                Avg_CLV       =("Predicted_CLV",  "mean"),
                Total_CLV     =("Predicted_CLV",  "sum"),
                Avg_Tenure    =("tenure",         "mean"),
                Avg_Monthly   =("MonthlyCharges", "mean"),
                Avg_Risk_Prob =("Churn_Prob",     "mean"),
            ).sort_values("Churn_Rate", ascending=False).reset_index(drop=True)

            summary.columns = [
                "Segment", "Customers", "Churned", "Churn Rate",
                "Avg CLV (Rs.)", "Total CLV (Rs.)", "Avg Tenure (mo)",
                "Avg Monthly (Rs.)", "Avg Risk Score",
            ]

            def cr_fmt(val, alt):
                try:
                    v = float(val)
                    if v > 0.6:  return F["c_red"]
                    if v > 0.3:  return F["c_amber"]
                    return F["c_green"]
                except Exception:
                    return F["cell"]

            seg_fmts = {
                "Segment":           lambda v, a: F["cell_l_alt" if a else "cell_l"],
                "Customers":         "num",
                "Churned":           "num",
                "Churn Rate":        cr_fmt,
                "Avg CLV (Rs.)":     "dec2",
                "Total CLV (Rs.)":   "dec2",
                "Avg Tenure (mo)":   "dec2",
                "Avg Monthly (Rs.)": "dec2",
                "Avg Risk Score":    "dec4",
            }
            write_table(ws2, summary, 4, c0=1, hdr_fmt="col_hdr",
                        col_fmts=seg_fmts, row_height=22)

            # ==================================================================
            #  SHEET 3 — HIGH RISK ACTION LIST
            # ==================================================================
            ws3 = wb.add_worksheet(_trunc("High Risk Customers"))
            ws3.hide_gridlines(2)
            ws3.set_zoom(90)
            ws3.set_column(0, 0, 1.5)
            for r, h in [(0, 5), (1, 46), (2, 20), (3, 8)]:
                ws3.set_row(r, h)

            ws3.merge_range("B2:M2",
                "HIGH-RISK CUSTOMERS  ·  Priority Retention Action List",
                F["title"])
            ws3.merge_range("B3:M3",
                "Customers with predicted churn probability ≥ 60%.  "
                "Sorted highest risk first.  Take immediate action on High Value rows.",
                F["subtitle"])

            hr_base = ["tenure", "MonthlyCharges", "TotalCharges",
                       "Predicted_CLV", "Churn_Prob", "Segment", "Risk_Tier"]
            hr_cols = ([id_col] if id_col else []) + hr_base
            hr_cols = [c for c in hr_cols if c in df.columns]

            hr = (
                df[df["Churn_Prob"] >= 0.60][hr_cols]
                .sort_values("Churn_Prob", ascending=False)
                .head(1000)
                .reset_index(drop=True)
            )

            col_widths3 = {
                "tenure": 14, "MonthlyCharges": 22, "TotalCharges": 22,
                "Predicted_CLV": 22, "Churn_Prob": 18,
                "Segment": 32, "Risk_Tier": 24,
            }

            if hr.empty:
                ws3.write(4, 1,
                    "No customers have a churn probability ≥ 60%.",
                    F["c_green"])
            else:
                for ci, col in enumerate(hr.columns):
                    ws3.set_column(1 + ci, 1 + ci, col_widths3.get(col, 22))

                # BUG FIX 7: prob_fmt correctly handles alt flag for all branches;
                # dead <0.60 branch removed since table only contains ≥0.60 rows.
                def prob_fmt(val, alt):
                    try:
                        v = float(val)
                        base = F["c_red"] if v >= 0.80 else F["c_amber"]
                        return base   # conditional formats don't need alt; colour is enough
                    except Exception:
                        return F["pct2_alt" if alt else "pct2"]

                hr_fmts = {
                    "tenure":        "num",
                    "MonthlyCharges":"num",
                    "TotalCharges":  "num",
                    "Predicted_CLV": "dec2",
                    "Churn_Prob":    prob_fmt,
                    "Segment":  lambda v, a: F["cell_l_alt" if a else "cell_l"],
                    "Risk_Tier": lambda v, a: F["cell_alt"  if a else "cell"],
                }
                if id_col:
                    hr_fmts[id_col] = lambda v, a: F["cell_l_alt" if a else "cell_l"]

                write_table(ws3, hr, 4, c0=1, hdr_fmt="col_hdr",
                            col_fmts=hr_fmts, row_height=20)

                # BUG FIX 15: autofilter corrected for c0=1 offset
                ws3.autofilter(4, 1, 4 + len(hr), 1 + len(hr.columns) - 1)
                ws3.freeze_panes(5, 0)

            # ==================================================================
            #  SHEET 4 — PROCESSED DATA
            # ==================================================================
            proc_base = ["tenure", "MonthlyCharges", "TotalCharges",
                         "AvgMonthlySpend", "ValueScore", "LoyaltyScore",
                         "SpendVariance", "Churn", "Churn_Prob",
                         "Predicted_CLV", "Segment", "Risk_Tier"]
            proc_cols = ([id_col] if id_col else []) + proc_base
            proc_cols = [c for c in proc_cols if c in df.columns]
            proc_df   = df[proc_cols].copy()

            proc_df.to_excel(writer, sheet_name=_trunc("Processed Data"), index=False)
            ws4 = writer.sheets[_trunc("Processed Data")]
            # BUG FIX 6: apply_num_fmt=True → per-column numeric formatting
            write_full_sheet(ws4, proc_df, "col_hdr_t", apply_num_fmt=True)

            # ==================================================================
            #  SHEET 5 — RAW DATA
            # ==================================================================
            original_df.to_excel(writer, sheet_name=_trunc("Raw Data"), index=False)
            ws5 = writer.sheets[_trunc("Raw Data")]
            write_full_sheet(ws5, original_df, "col_hdr_s")

            # ==================================================================
            #  SHEET 6 — DATA QUALITY REPORT
            # ==================================================================
            ws6 = wb.add_worksheet(_trunc("Data Quality Report"))
            ws6.hide_gridlines(2)
            ws6.set_zoom(95)
            ws6.set_column(0, 0, 1.5)
            ws6.set_column(1, 1, 38)
            ws6.set_column(2, 2, 60)
            for r, h in [(0, 5), (1, 46), (2, 20), (3, 8)]:
                ws6.set_row(r, h)

            ws6.merge_range("B2:C2", "DATA QUALITY & PIPELINE REPORT", F["title"])
            ws6.merge_range("B3:C3",
                "Full audit trail: file details · column detection · "
                "model metrics · KPIs · methodology notes",
                F["subtitle"])

            def sec(ws, row, text):
                ws.merge_range(row, 1, row, 2, text, F["sec_banner"])
                ws.set_row(row, 20)

            # Audit notes for suppressed metrics
            metrics_note = (
                "SUPPRESSED — churn labels are synthetic; metrics would "
                "reflect formula self-consistency, not true predictive power."
            ) if _synthetic_churn else acc_str

            auc_note = (
                "SUPPRESSED — see Accuracy note above."
            ) if _synthetic_churn else auc_str

            pipeline_rows = [
                ("sec", "FILE INFORMATION"),
                ("row", "Source File",           os.path.basename(file_path)),
                ("row", "Full Path",              file_path),
                ("row", "Original Rows",          f"{original_rows:,}"),
                ("row", "Original Columns",       f"{original_cols_n}"),
                ("row", "Rows After Cleaning",    f"{total_cust:,}"),
                ("row", "Rows Removed",           f"{original_rows - total_cust:,}"),

                ("sec", "COLUMN DETECTION"),
                ("row", "Churn Column",           churn_col  or "Not found — synthetic used"),
                ("row", "Churn Source",           churn_source),
                ("row", "Churn Labels Valid",     "NO — synthetic formula" if _synthetic_churn
                                                  else "YES — real column detected"),
                ("row", "Tenure Column",          tenure_col  or "Not found — numeric fallback"),
                ("row", "Monthly Charges Column", monthly_col or "Not found — numeric fallback"),
                ("row", "Total Charges Column",   total_col   or "Not found — computed"),
                ("row", "Customer ID Column",     id_col      or "Not found"),

                ("sec", "MACHINE LEARNING MODEL"),
                ("row", "Model Used",             model_name),
                ("row", "Features Used",          ", ".join(FEATURES)),
                ("row", "Accuracy",               metrics_note),
                ("row", "AUC-ROC Score",          auc_note),
                ("row", "Metrics Suppressed",     "YES — synthetic labels" if _synthetic_churn
                                                  else "NO — real labels used"),

                ("sec", "METHODOLOGY NOTES"),
                ("row", "CLV Methodology",
                 "Retention-weighted historical proxy: TotalCharges × (1 – Churn_Prob). "
                 "NOT a survival-model or DCF CLV. Use for relative ranking only."),
                ("row", "Revenue at Risk",
                 "Annualised run-rate: MonthlyCharges × 12, summed for Churn_Prob ≥ 60% "
                 "customers. Forward-looking proxy. Historical TotalCharges at risk: "
                 f"Rs.{tc_at_risk_hist:,.0f}."),
                ("row", "Churn Probability",
                 "Predicted by ML model (Random Forest / Gradient Boosting) when data "
                 ">= 20 rows with real labels; otherwise heuristic percentile ranking."),
                ("row", "Segmentation",
                 "2×2 matrix: High/Low Risk (Churn_Prob ≥ 0.60) × "
                 "High/Low Value (Predicted_CLV ≥ median)."),

                ("sec", "OUTPUT KPIs"),
                ("row", "Total Customers",           f"{total_cust:,}"),
                ("row", "Churn Rate",                f"{churn_rate}%"),
                ("row", "Retention Rate",            f"{retention_rate}%"),
                ("row", "High-Risk (≥60%)",          f"{high_risk_n:,}  ({high_risk_pct}%)"),
                ("row", "Medium-Risk (30-60%)",      f"{med_risk_n:,}"),
                ("row", "Low-Risk (<30%)",           f"{low_risk_n:,}"),
                ("row", "Annualised Risk Revenue",
                 f"Rs.{revenue_at_risk:,.0f}  (MonthlyCharges×12 of High-Risk customers)"),
                ("row", "Historical TotalCharges at Risk",
                 f"Rs.{tc_at_risk_hist:,.0f}"),
                ("row", "Total CLV Portfolio",       f"Rs.{total_clv:,.0f}"),
                ("row", "Avg CLV per Customer",      f"Rs.{avg_clv:,.0f}"),
                ("row", "Avg Tenure",                f"{avg_tenure} months"),
                ("row", "Avg Monthly Charge",        f"Rs.{avg_monthly:,.0f}"),
            ]

            r          = 4
            data_row_i = 0
            for item in pipeline_rows:
                if item[0] == "sec":
                    sec(ws6, r, item[1])
                    data_row_i = 0
                else:
                    alt = (data_row_i % 2 == 1)
                    ws6.set_row(r, 20)
                    ws6.write(r, 1, item[1], F["cell_l_alt" if alt else "cell_l"])
                    ws6.write(r, 2, item[2], F["cell_alt"   if alt else "cell"])
                    data_row_i += 1
                r += 1

    # ── Write with full error surfacing ────────────────────────────────────────
    # BUG FIX 20: exceptions inside build_report are caught and shown clearly.
    try:
        final_path = safe_write_excel(build_report, save_path)
    except Exception as build_exc:
        raise RuntimeError(
            f"Report generation failed:\n{build_exc}\n\n"
            f"{traceback.format_exc()[-800:]}"
        )

    # ── Success ────────────────────────────────────────────────────────────────
    saved_name   = os.path.basename(final_path)
    renamed_note = (
        f"\nOriginal file was locked — saved as:\n  {saved_name}\n"
        if final_path != save_path else ""
    )
    synthetic_note = (
        "\n** NOTE: Churn labels were SYNTHETIC — ML accuracy/AUC suppressed. **\n"
        "   Risk scores and segmentation remain valid for relative ranking.\n"
    ) if _synthetic_churn else ""

    info(
        "Report Ready  ✓",
        f"Report saved successfully!{renamed_note}{synthetic_note}\n\n"
        f"  {final_path}\n\n"
        f"{'=' * 48}\n"
        f"  Customers Analysed          :  {total_cust:,}\n"
        f"  Churn Rate                  :  {churn_rate}%\n"
        f"  Retention Rate              :  {retention_rate}%\n"
        f"  High-Risk Customers         :  {high_risk_n:,}  ({high_risk_pct}%)\n"
        f"  Annualised Risk Revenue     :  Rs.{revenue_at_risk:,.0f}\n"
        f"  Historical TotalCharges     :  Rs.{tc_at_risk_hist:,.0f}\n"
        f"  Total CLV Portfolio         :  Rs.{total_clv:,.0f}\n"
        f"  ML Model                    :  {model_name}\n"
        f"  Accuracy / AUC-ROC          :  {acc_str}  /  {auc_str}\n"
        f"{'=' * 48}\n\n"
        f"Sheets in your report:\n"
        f"  Dashboard              14 KPI cards + 6 charts\n"
        f"  Segment Summary        Red/Amber/Green churn flags\n"
        f"  High Risk Customers    Top 1,000 at-risk customers\n"
        f"  Processed Data         All engineered features (formatted)\n"
        f"  Raw Data               Original file untouched\n"
        f"  Data Quality Report    Full pipeline audit + methodology\n"
        f"  Data – Segments/Risk   7 chart-source sheets",
    )

# ==============================================================================
#  ERROR HANDLER
# ==============================================================================
except SystemExit:
    pass
except Exception as e:
    tb = traceback.format_exc()
    err(
        "Unexpected Error",
        f"Something went wrong:\n\n{str(e)}\n\n"
        f"Traceback (last 1200 chars):\n{tb[-1200:]}",
    )
finally:
    try:
        root.destroy()
    except Exception:
        pass
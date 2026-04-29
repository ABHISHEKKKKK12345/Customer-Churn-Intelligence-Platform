# ==============================================================================
#  CUSTOMER CHURN INTELLIGENCE PLATFORM  ·  GUI Edition
#
#  Fully windowed — zero console interaction.
#  Launch via .exe or `python churn-intelligence-platform-gui.py`
# ==============================================================================

import os
import re
import sys
import time
import threading
import warnings
import traceback

import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

warnings.filterwarnings("ignore")

from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
from sklearn.metrics import accuracy_score, roc_auc_score
from sklearn.impute import SimpleImputer
from sklearn.preprocessing import LabelEncoder

# ==============================================================================
#  CONSTANTS
# ==============================================================================
MAX_RETRIES     = 3
SHEET_NAME_LEN  = 31
LABEL_ENC_MAX   = 20
MIN_ROWS_FOR_ML = 20

APP_TITLE   = "Customer Churn Intelligence Platform (GUI)"
APP_VERSION = "v1.0"
AUTHOR      = "Abhishek"

# Colour palette for the GUI
GUI = {
    "bg":         "#0B1F3A",
    "bg2":        "#1A3A5C",
    "bg3":        "#0F2847",
    "accent":     "#00A896",
    "accent2":    "#007C71",
    "red":        "#E74C3C",
    "orange":     "#E67E22",
    "green":      "#27AE60",
    "purple":     "#8E44AD",
    "white":      "#FFFFFF",
    "lgrey":      "#F0F4F8",
    "mgrey":      "#C8D0DC",
    "dgrey":      "#5D6D7E",
    "text":       "#ECF0F1",
    "text2":      "#BDC3C7",
    "warn_bg":    "#FFF3CD",
    "warn_fg":    "#856404",
    "card_bg":    "#142D52",
    "card_bd":    "#2E5E8E",
    "gold":       "#F39C12",
}


def _trunc(s, n=SHEET_NAME_LEN):
    return str(s)[:n]


# ==============================================================================
#  SPLASH SCREEN
# ==============================================================================
class SplashScreen:
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.configure(bg=GUI["bg"])

        W, H = 600, 380
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x  = (sw - W) // 2
        y  = (sh - H) // 2
        self.root.geometry(f"{W}x{H}+{x}+{y}")
        self.root.attributes("-topmost", True)

        # Outer frame with border
        outer = tk.Frame(self.root, bg=GUI["accent2"], padx=2, pady=2)
        outer.pack(fill="both", expand=True)

        inner = tk.Frame(outer, bg=GUI["bg"], padx=30, pady=24)
        inner.pack(fill="both", expand=True)

        # Logo area
        logo_frame = tk.Frame(inner, bg=GUI["bg"])
        logo_frame.pack(pady=(10, 0))

        tk.Label(logo_frame, text="📊", font=("Segoe UI Emoji", 42),
                 bg=GUI["bg"], fg=GUI["accent"]).pack()

        tk.Label(inner, text="Customer Churn Intelligence",
                 font=("Segoe UI", 22, "bold"),
                 bg=GUI["bg"], fg=GUI["white"]).pack(pady=(4, 0))

        tk.Label(inner, text="Platform",
                 font=("Segoe UI", 22, "bold"),
                 bg=GUI["bg"], fg=GUI["accent"]).pack()

        tk.Label(inner, text="Advanced Customer Analytics Platform",
                 font=("Segoe UI", 10),
                 bg=GUI["bg"], fg=GUI["text2"]).pack(pady=(6, 0))

        tk.Label(inner, text=AUTHOR,
                 font=("Segoe UI", 9),
                 bg=GUI["bg"], fg=GUI["dgrey"]).pack(pady=(2, 16))

        # Progress bar
        self.prog_var = tk.DoubleVar(value=0)
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Splash.Horizontal.TProgressbar",
                        troughcolor=GUI["bg2"],
                        background=GUI["accent"],
                        borderwidth=0,
                        lightcolor=GUI["accent"],
                        darkcolor=GUI["accent2"])
        self.pb = ttk.Progressbar(inner, variable=self.prog_var,
                                  style="Splash.Horizontal.TProgressbar",
                                  length=480, mode="determinate")
        self.pb.pack()

        self.status = tk.Label(inner, text="Initialising…",
                               font=("Segoe UI", 9),
                               bg=GUI["bg"], fg=GUI["text2"])
        self.status.pack(pady=(6, 0))

        tk.Label(inner, text=APP_VERSION,
                 font=("Segoe UI", 8),
                 bg=GUI["bg"], fg=GUI["dgrey"]).pack(side="bottom")

        self.root.update()

    def update(self, pct, msg):
        self.prog_var.set(pct)
        self.status.config(text=msg)
        self.root.update()

    def close(self):
        self.root.destroy()


# ==============================================================================
#  MAIN APPLICATION WINDOW
# ==============================================================================
class ChurnApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f"{APP_TITLE}  ·  {APP_VERSION}")
        self.root.configure(bg=GUI["bg"])
        self.root.resizable(True, True)

        W, H = 960, 720
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x  = (sw - W) // 2
        y  = max(0, (sh - H) // 2)
        self.root.geometry(f"{W}x{H}+{x}+{y}")
        self.root.minsize(820, 600)

        # State
        self.file_path    = tk.StringVar(value="")
        self.save_path    = tk.StringVar(value="")
        self.status_var   = tk.StringVar(value="Ready.  Select a dataset to begin.")
        self.prog_var     = tk.DoubleVar(value=0)
        self.running      = False
        self.result_data  = {}

        self._setup_styles()
        self._build_ui()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Styles ─────────────────────────────────────────────────────────────────
    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Main.Horizontal.TProgressbar",
                        troughcolor=GUI["bg2"],
                        background=GUI["accent"],
                        borderwidth=0,
                        lightcolor=GUI["accent"],
                        darkcolor=GUI["accent2"])

        style.configure("Green.TButton",
                        font=("Segoe UI", 11, "bold"),
                        padding=(18, 10))

        style.configure("TEntry",
                        fieldbackground=GUI["bg2"],
                        foreground=GUI["white"],
                        insertcolor=GUI["white"])

    # ── UI Build ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ─────────────────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=GUI["bg"], height=80)
        hdr.pack(fill="x", padx=0, pady=0)
        hdr.pack_propagate(False)

        tk.Label(hdr, text="📊  Customer Churn Intelligence Platform",
                 font=("Segoe UI", 17, "bold"),
                 bg=GUI["bg"], fg=GUI["white"]).pack(side="left", padx=24, pady=18)

        tk.Label(hdr, text=APP_VERSION,
                 font=("Segoe UI", 9),
                 bg=GUI["bg"], fg=GUI["dgrey"]).pack(side="right", padx=24)

        tk.Frame(self.root, bg=GUI["accent2"], height=2).pack(fill="x")

        # ── Author strip ────────────────────────────────────────────────────────
        auth = tk.Frame(self.root, bg=GUI["bg3"], height=28)
        auth.pack(fill="x")
        auth.pack_propagate(False)
        tk.Label(auth, text=AUTHOR,
                 font=("Segoe UI", 9),
                 bg=GUI["bg3"], fg=GUI["text2"]).pack(side="left", padx=24)

        # ── Main scroll canvas ──────────────────────────────────────────────────
        canvas = tk.Canvas(self.root, bg=GUI["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self.scroll_frame = tk.Frame(canvas, bg=GUI["bg"])
        self.win_id = canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")

        def on_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def on_resize(e):
            canvas.itemconfig(self.win_id, width=e.width)

        self.scroll_frame.bind("<Configure>", on_configure)
        canvas.bind("<Configure>", on_resize)

        # Mouse-wheel scroll
        def _mw(e):
            canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _mw)

        pad = {"padx": 28, "pady": 8}

        # ── Section: Introduction ───────────────────────────────────────────────
        self._section("What This Tool Does", "ℹ", pad)

        desc_text = (
            "  ✔  Auto-detects Churn, Tenure, Revenue and Customer ID columns\n"
            "  ✔  Cleans & repairs bad / missing / mixed-format data automatically\n"
            "  ✔  Trains a Machine Learning model (Random Forest + Gradient Boosting)\n"
            "  ✔  Produces a full executive Excel report with 14 KPI cards & 6 charts\n"
            "  ✔  Segments customers by Risk × Value matrix\n"
            "  ✔  Estimates Revenue-at-Risk and Customer Lifetime Value (CLV)\n\n"
            "  Supported Formats:   CSV  ·  Excel (.xlsx / .xls / .xlsm)  ·  ODS  ·  TSV / TXT"
        )
        desc = tk.Label(self.scroll_frame, text=desc_text,
                        font=("Segoe UI", 10), justify="left",
                        bg=GUI["card_bg"], fg=GUI["text"],
                        relief="flat", bd=0, padx=18, pady=14)
        desc.pack(fill="x", **pad)

        # ── Section: Step 1 — Select Dataset ────────────────────────────────────
        self._section("Step 1  —  Select Your Dataset", "📂", pad)

        file_frame = tk.Frame(self.scroll_frame, bg=GUI["card_bg"],
                              relief="flat", padx=18, pady=14)
        file_frame.pack(fill="x", **pad)

        tk.Label(file_frame, text="Dataset File:",
                 font=("Segoe UI", 10, "bold"),
                 bg=GUI["card_bg"], fg=GUI["text"]).grid(row=0, column=0,
                                                          sticky="w", pady=4)

        self.file_entry = tk.Entry(file_frame, textvariable=self.file_path,
                                   font=("Segoe UI", 10),
                                   bg=GUI["bg2"], fg=GUI["white"],
                                   insertbackground=GUI["white"],
                                   relief="flat", bd=0,
                                   state="readonly", width=60)
        self.file_entry.config(
            readonlybackground=GUI["bg2"],
            fg=GUI["white"]
            )
        self.file_entry.grid(row=0, column=1, padx=(10, 8), sticky="ew", ipady=6)
        file_frame.columnconfigure(1, weight=1)

        self._btn(file_frame, "Browse…", self._browse_file,
                  bg=GUI["accent2"], fg=GUI["white"]).grid(row=0, column=2)

        self.file_info_lbl = tk.Label(file_frame, text="No file selected.",
                                       font=("Segoe UI", 9), justify="left",
                                       bg=GUI["card_bg"], fg=GUI["dgrey"],
                                       wraplength=700)
        self.file_info_lbl.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # ── Section: Step 2 — Save Report ───────────────────────────────────────
        self._section("Step 2  —  Choose Save Location", "💾", pad)

        save_frame = tk.Frame(self.scroll_frame, bg=GUI["card_bg"],
                              relief="flat", padx=18, pady=14)
        save_frame.pack(fill="x", **pad)

        tk.Label(save_frame, text="Save Report As:",
                 font=("Segoe UI", 10, "bold"),
                 bg=GUI["card_bg"], fg=GUI["text"]).grid(row=0, column=0,
                                                          sticky="w", pady=4)

        self.save_entry = tk.Entry(save_frame, textvariable=self.save_path,
                                   font=("Segoe UI", 10),
                                   bg=GUI["bg2"], fg=GUI["white"],
                                   insertbackground=GUI["white"],
                                   relief="flat", bd=0, width=60)
        self.save_entry.grid(row=0, column=1, padx=(10, 8), sticky="ew", ipady=6)
        save_frame.columnconfigure(1, weight=1)

        self._btn(save_frame, "Browse…", self._browse_save,
                  bg=GUI["accent2"], fg=GUI["white"]).grid(row=0, column=2)

        tk.Label(save_frame,
                 text="Default: Churn_Analytics_Report.xlsx on your Desktop",
                 font=("Segoe UI", 9),
                 bg=GUI["card_bg"], fg=GUI["dgrey"]).grid(
                     row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))

        # Set default save path
        default_save = os.path.join(
            os.path.expanduser("~"), "Desktop", "Churn_Analytics_Report.xlsx"
        )
        self.save_path.set(default_save)

        # ── Section: Step 3 — Run Analysis ──────────────────────────────────────
        self._section("Step 3  —  Run Analysis", "🚀", pad)

        run_frame = tk.Frame(self.scroll_frame, bg=GUI["card_bg"],
                             relief="flat", padx=18, pady=18)
        run_frame.pack(fill="x", **pad)

        self.run_btn = tk.Button(
            run_frame,
            text="▶   Run Churn Analysis  &  Generate Report",
            font=("Segoe UI", 13, "bold"),
            bg=GUI["accent"], fg=GUI["white"],
            activebackground=GUI["accent2"],
            activeforeground=GUI["white"],
            relief="flat", bd=0,
            padx=32, pady=14,
            cursor="hand2",
            command=self._run_analysis
        )
        self.run_btn.pack(fill="x")

        # ── Progress area ────────────────────────────────────────────────────────
        prog_frame = tk.Frame(self.scroll_frame, bg=GUI["card_bg"],
                              relief="flat", padx=18, pady=14)
        prog_frame.pack(fill="x", **pad)

        self.pb = ttk.Progressbar(prog_frame, variable=self.prog_var,
                                  style="Main.Horizontal.TProgressbar",
                                  length=800, mode="determinate")
        self.pb.pack(fill="x")

        self.status_lbl = tk.Label(prog_frame, textvariable=self.status_var,
                                   font=("Segoe UI", 10),
                                   bg=GUI["card_bg"], fg=GUI["white"],
                                   anchor="w")
        self.status_lbl.pack(fill="x", pady=(6, 0))

        # ── Results area (KPI cards) ─────────────────────────────────────────────
        self._section("Results  —  Key Performance Indicators", "📈", pad)

        self.kpi_frame = tk.Frame(self.scroll_frame, bg=GUI["bg"])
        self.kpi_frame.pack(fill="x", **pad)

        self._kpi_placeholder()

        # ── Log area ─────────────────────────────────────────────────────────────
        self._section("Activity Log", "📋", pad)

        log_outer = tk.Frame(self.scroll_frame, bg=GUI["card_bg"],
                             relief="flat", padx=14, pady=10)
        log_outer.pack(fill="x", **pad)

        self.log_text = tk.Text(log_outer, height=10,
                                font=("Consolas", 9),
                                bg="#0A1928", fg=GUI["text"],
                                insertbackground=GUI["white"],
                                relief="flat", bd=0,
                                state="disabled",
                                wrap="word")
        log_sb = ttk.Scrollbar(log_outer, orient="vertical",
                               command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        log_sb.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        # Tag colours for log
        self.log_text.tag_config("info",    foreground=GUI["text"])
        self.log_text.tag_config("ok",      foreground=GUI["green"])
        self.log_text.tag_config("warn",    foreground=GUI["orange"])
        self.log_text.tag_config("err",     foreground=GUI["red"])
        self.log_text.tag_config("section", foreground=GUI["accent"])

        # Footer
        tk.Frame(self.scroll_frame, bg=GUI["bg"], height=20).pack()
        tk.Label(self.scroll_frame,
                 text=f"© 2025–2026  {AUTHOR}",
                 font=("Segoe UI", 8),
                 bg=GUI["bg"], fg=GUI["dgrey"]).pack(pady=(0, 16))

        self._log("section", f"{'─'*60}")
        self._log("section", f"  {APP_TITLE}  {APP_VERSION}")
        self._log("section", f"  {AUTHOR}")
        self._log("section", f"{'─'*60}")
        self._log("info", "Application ready.  Please select a dataset file to begin.")

    # ── Helpers ─────────────────────────────────────────────────────────────────
    def _section(self, text, icon, pad):
        f = tk.Frame(self.scroll_frame, bg=GUI["bg3"], height=34)
        f.pack(fill="x", padx=pad["padx"], pady=(14, 0))
        f.pack_propagate(False)
        tk.Label(f, text=f" {icon}  {text}",
                 font=("Segoe UI", 11, "bold"),
                 bg=GUI["bg3"], fg=GUI["accent"]).pack(side="left", padx=10)

    def _btn(self, parent, text, cmd, bg=None, fg=None):
        return tk.Button(parent, text=text, command=cmd,
                         font=("Segoe UI", 10, "bold"),
                         bg=bg or GUI["bg2"],
                         fg=fg or GUI["white"],
                         activebackground=GUI["accent2"],
                         activeforeground=GUI["white"],
                         relief="flat", bd=0,
                         padx=14, pady=7,
                         cursor="hand2")

    def _kpi_placeholder(self):
        for w in self.kpi_frame.winfo_children():
            w.destroy()
        ph = tk.Label(self.kpi_frame,
                      text="Run the analysis to see KPI results here.",
                      font=("Segoe UI", 10),
                      bg=GUI["card_bg"], fg=GUI["dgrey"],
                      padx=18, pady=18)
        ph.pack(fill="x")

    def _log(self, tag, msg):
        def _write():
            self.log_text.config(state="normal")
            ts = time.strftime("%H:%M:%S")
            prefix = {"info": "ℹ", "ok": "✔", "warn": "⚠",
                      "err": "✖", "section": "─"}.get(tag, " ")
            self.log_text.insert("end", f"[{ts}]  {prefix}  {msg}\n", tag)
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.root.after(0, _write)

    def _set_status(self, msg, pct=None):
        def _upd():
            self.status_var.set(msg)
            if pct is not None:
                self.prog_var.set(pct)
        self.root.after(0, _upd)

    # ── File browsing ────────────────────────────────────────────────────────────
    def _browse_file(self):
        p = filedialog.askopenfilename(
            parent=self.root,
            title="Select Your Dataset File",
            filetypes=[
                ("All Supported Files",
                 "*.csv *.xlsx *.xls *.xlsm *.xlsb *.ods *.txt *.tsv"),
                ("CSV Files",           "*.csv"),
                ("Excel Files",         "*.xlsx *.xls *.xlsm *.xlsb"),
                ("ODS Spreadsheet",     "*.ods"),
                ("Text / TSV",          "*.txt *.tsv"),
            ],
        )
        if p:
            self.file_path.set(p)
            size = os.path.getsize(p)
            size_str = (f"{size/1024/1024:.1f} MB" if size > 1024*1024
                        else f"{size/1024:.1f} KB")
            self.file_info_lbl.config(
                text=f"Selected:  {os.path.basename(p)}   |   Size: {size_str}   |   "
                     f"Path: {os.path.dirname(p)}",
                fg=GUI["accent"]
            )
            self._log("ok", f"Dataset selected: {os.path.basename(p)}  ({size_str})")

            # Auto-set save path next to file
            base = os.path.splitext(p)[0]
            self.save_path.set(base + "_Churn_Report.xlsx")

    def _browse_save(self):
        p = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save Analytics Report As",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile="Churn_Analytics_Report.xlsx",
        )
        if p:
            self.save_path.set(p)
            self._log("ok", f"Save location set: {os.path.basename(p)}")

    # ── Run Analysis ─────────────────────────────────────────────────────────────
    def _run_analysis(self):
        if self.running:
            return

        if not self.file_path.get():
            messagebox.showwarning("No File Selected",
                                   "Please select a dataset file first.",
                                   parent=self.root)
            return

        if not self.save_path.get():
            messagebox.showwarning("No Save Location",
                                   "Please choose where to save the report.",
                                   parent=self.root)
            return

        # Confirm
        proceed = messagebox.askyesno(
            "Confirm Analysis",
            f"File   :  {os.path.basename(self.file_path.get())}\n"
            f"Save to:  {os.path.basename(self.save_path.get())}\n\n"
            "Ready to run the full churn analysis?\n\n"
            "This will:\n"
            "  • Load and clean your dataset\n"
            "  • Train a Machine Learning model\n"
            "  • Generate the Excel report with charts\n\n"
            "Click YES to proceed.",
            parent=self.root,
            icon="question"
        )
        if not proceed:
            return

        self.running = True
        self.run_btn.config(state="disabled",
                            text="⏳ ANALYSIS RUNNING... PLEASE WAIT",
                            bg=GUI["bg3"])
        self.prog_var.set(0)
        self._kpi_placeholder()

        t = threading.Thread(target=self._analysis_thread, daemon=True)
        t.start()

    # ── Analysis thread ───────────────────────────────────────────────────────
    def _analysis_thread(self):
        try:
            self._run_pipeline()
        except Exception as e:
            tb = traceback.format_exc()
            self._log("err", f"FATAL ERROR: {e}")
            self._log("err", tb[-800:])
            self._set_status(f"Error: {e}", 0)
            self.root.after(0, lambda: messagebox.showerror(
                "Analysis Failed",
                f"An error occurred during analysis:\n\n{e}\n\n"
                f"Please check the Activity Log for details.",
                parent=self.root
            ))
        finally:
            self.running = False
            self.root.after(0, lambda: self.run_btn.config(
                state="normal",
                text="▶   Run Churn Analysis  &  Generate Report",
                bg=GUI["accent"]
            ))

    def _run_pipeline(self):
        fp = self.file_path.get()
        sp = self.save_path.get()

        # ── Load ────────────────────────────────────────────────────────────────
        self._set_status("Loading dataset…", 5)
        self._log("section", "─── STEP 1: Loading Dataset ───────────────────────────")
        self._log("info", f"Reading file: {os.path.basename(fp)}")
        raw_df = load_file(fp)
        original_rows, original_cols_n = raw_df.shape
        original_df = raw_df.copy()
        self._log("ok", f"Loaded {original_rows:,} rows × {original_cols_n} columns")
        self._set_status("File loaded.  Cleaning data…", 10)

        # ── Clean ───────────────────────────────────────────────────────────────
        self._log("section", "─── STEP 2: Cleaning Data ─────────────────────────────")
        df = raw_df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]
        df.dropna(how="all", inplace=True)
        df.dropna(axis=1, how="all", inplace=True)
        df.drop_duplicates(inplace=True)
        df.reset_index(drop=True, inplace=True)

        if df.empty:
            raise RuntimeError("No usable rows after cleaning.")

        removed = original_rows - len(df)
        self._log("ok", f"Cleaning complete. {removed:,} rows removed. {len(df):,} rows remain.")
        self._set_status("Data cleaned.  Detecting columns…", 18)

        # ── Column detection ─────────────────────────────────────────────────────
        self._log("section", "─── STEP 3: Column Detection ──────────────────────────")
        churn_col   = find_col(df, "churn","attrition","churned","leave",
                                   "exit","cancel","left","status")
        tenure_col  = find_col(df, "tenure","seniority","duration",
                                   "period","months","age")
        monthly_col = find_col(df, "monthlycharge","monthlyfee","monthly",
                                   "charge","rate","fee","price","subscription")
        total_col   = find_col(df, "totalcharge","totalrevenue","lifetimevalue",
                                   "total","revenue","bill","spend","amount","sales","ltv")
        id_col      = find_col(df, "customerid","custid","userid",
                                   "accountid","clientid","memberid","id")

        detected = {}
        for role, col in [("churn",churn_col),("tenure",tenure_col),
                          ("monthly",monthly_col),("total",total_col),("id",id_col)]:
            if col and col not in detected.values():
                detected[role] = col
            else:
                detected[role] = None

        churn_col   = detected["churn"]
        tenure_col  = detected["tenure"]
        monthly_col = detected["monthly"]
        total_col   = detected["total"]
        id_col      = detected["id"]

        self._log("ok", f"Churn col    : {churn_col   or 'NOT FOUND — synthetic labels will be used'}")
        self._log("ok", f"Tenure col   : {tenure_col  or 'NOT FOUND — fallback'}")
        self._log("ok", f"Revenue col  : {monthly_col or 'NOT FOUND — fallback'}")
        self._log("ok", f"Total col    : {total_col   or 'NOT FOUND — computed'}")
        self._log("ok", f"Customer ID  : {id_col      or 'NOT FOUND'}")
        self._set_status("Columns detected.  Encoding features…", 25)

        # ── Numeric conversion ──────────────────────────────────────────────────
        skip_numeric = {c for c in [churn_col, id_col] if c}
        for col in df.columns:
            if col in skip_numeric:
                continue
            temp = clean_num(df[col])
            if temp.notna().sum() / max(len(temp), 1) >= 0.40:
                df[col] = temp

        for col in df.select_dtypes(include="object").columns:
            if col in skip_numeric:
                continue
            try:
                nu = df[col].nunique()
                if 2 <= nu <= LABEL_ENC_MAX:
                    le = LabelEncoder()
                    df[col] = le.fit_transform(
                        df[col].fillna("__missing__").astype(str))
                else:
                    df[col] = np.nan
            except Exception:
                df[col] = np.nan

        numeric_cols = df.select_dtypes(include=np.number).columns.tolist()

        # ── Core columns ────────────────────────────────────────────────────────
        def pick(col_name, fallback_idx, default_val):
            if col_name and col_name in df.columns:
                s = clean_num(df[col_name])
            elif len(numeric_cols) > fallback_idx:
                s = pd.to_numeric(df[numeric_cols[fallback_idx]], errors="coerce")
            else:
                return pd.Series([float(default_val)] * len(df), dtype=float)
            s = s.clip(lower=0)
            med = s.median()
            fill = med if (pd.notna(med) and np.isfinite(med)) else float(default_val)
            return s.fillna(fill)

        df["tenure"]         = pick(tenure_col,  0, 12.0)
        df["MonthlyCharges"] = pick(monthly_col, 1, 500.0)
        df["TotalCharges"]   = pick(total_col,   2, 6000.0)

        bad_total = df["TotalCharges"] < df["MonthlyCharges"]
        df.loc[bad_total, "TotalCharges"] = (
            df.loc[bad_total, "tenure"] * df.loc[bad_total, "MonthlyCharges"])
        df["TotalCharges"] = (
            df["TotalCharges"]
            .fillna(df["tenure"] * df["MonthlyCharges"])
            .clip(lower=0))

        self._set_status("Building churn labels…", 32)
        self._log("section", "─── STEP 4: Churn Label Mapping ───────────────────────")

        # ── Churn labels ─────────────────────────────────────────────────────────
        YES = frozenset({"yes","y","true","1","1.0","churned","left","cancelled",
                         "cancel","exit","quit","inactive","lost","gone","departed",
                         "attrited","closed"})
        NO  = frozenset({"no","n","false","0","0.0","active","retained","stay",
                         "stayed","current","existing","ongoing","present","alive",
                         "good","loyal"})

        churn_source     = "Synthetic (no churn column found)"
        _synthetic_churn = True

        if churn_col and churn_col in df.columns:
            raw_c  = df[churn_col].astype(str).str.strip().str.lower()
            mapped = raw_c.map(lambda v: 1 if v in YES else (0 if v in NO else np.nan))
            if mapped.notna().sum() == 0:
                mapped = pd.to_numeric(df[churn_col], errors="coerce").round().clip(0,1)
            null_frac = mapped.isna().mean()
            if null_frac > 0.60:
                df["Churn"]      = np.where(df["tenure"] < df["tenure"].median(), 1, 0).astype(int)
                churn_source     = f"Synthetic ('{churn_col}' had {null_frac*100:.0f}% unmapped)"
                _synthetic_churn = True
                self._log("warn", f"Churn column had too many unmapped values → synthetic labels")
            else:
                mode_v           = int(mapped.mode().iloc[0]) if mapped.notna().any() else 0
                df["Churn"]      = mapped.fillna(mode_v).astype(int)
                churn_source     = f"Detected: '{churn_col}'"
                _synthetic_churn = False
                self._log("ok", f"Real churn labels found in column: {churn_col}")
        else:
            _sv = ((df["MonthlyCharges"] - (df["TotalCharges"]/(df["tenure"]+1)))
                   / (df["TotalCharges"]/(df["tenure"]+1)+1)).abs().clip(lower=0).fillna(0)
            risk_score = (
                df["tenure"].rank(pct=True,method="average").rsub(1) * 0.40
                + df["MonthlyCharges"].rank(pct=True,method="average") * 0.30
                + _sv.rank(pct=True,method="average") * 0.20
                + df["TotalCharges"].rank(pct=True,method="average").rsub(1) * 0.10)
            threshold        = risk_score.quantile(0.65)
            df["Churn"]      = (risk_score >= threshold).astype(int)
            churn_source     = "Synthetic — percentile risk score (top 35% flagged)"
            _synthetic_churn = True
            self._log("warn", "No churn column found — synthetic labels generated via risk-score formula")

        self._set_status("Engineering features…", 40)
        self._log("section", "─── STEP 5: Feature Engineering ───────────────────────")

        # ── Features ─────────────────────────────────────────────────────────────
        df["AvgMonthlySpend"] = (df["TotalCharges"]/(df["tenure"]+1)).clip(lower=0).fillna(0)
        df["ValueScore"]      = np.log1p(df["TotalCharges"]).clip(lower=0).fillna(0)
        df["LoyaltyScore"]    = (df["tenure"]*(df["TotalCharges"]/(df["MonthlyCharges"]+1))).clip(lower=0).fillna(0)
        df["SpendVariance"]   = ((df["MonthlyCharges"]-df["AvgMonthlySpend"])/(df["AvgMonthlySpend"]+1)).abs().clip(lower=0).fillna(0)
        self._log("ok", "Engineered: AvgMonthlySpend, ValueScore, LoyaltyScore, SpendVariance")

        FEATURES = ["tenure","MonthlyCharges","TotalCharges",
                    "AvgMonthlySpend","ValueScore","LoyaltyScore","SpendVariance"]
        X_raw = df[FEATURES].copy()
        y     = df["Churn"].copy()

        imputer = SimpleImputer(strategy="median")
        X = pd.DataFrame(imputer.fit_transform(X_raw), columns=FEATURES)

        heuristic = (
            X["tenure"].rank(pct=True).rsub(1) * 0.55
            + X["MonthlyCharges"].rank(pct=True) * 0.25
            + X["SpendVariance"].rank(pct=True) * 0.20
        ).clip(0, 1)
        df["Churn_Prob"] = heuristic.values

        # ── ML ───────────────────────────────────────────────────────────────────
        self._set_status("Training Machine Learning model…", 52)
        self._log("section", "─── STEP 6: Machine Learning ──────────────────────────")

        accuracy  = None
        auc_score = None
        model_name = ("Heuristic (synthetic labels)" if _synthetic_churn else "Heuristic")

        if len(df) >= MIN_ROWS_FOR_ML and y.nunique() > 1:
            for Clf, name, kw in [
                (RandomForestClassifier, "Random Forest",
                 dict(n_estimators=300, max_depth=12, class_weight="balanced",
                      random_state=42, n_jobs=-1)),
                (GradientBoostingClassifier, "Gradient Boosting",
                 dict(n_estimators=150, max_depth=5, random_state=42)),
            ]:
                try:
                    self._log("info", f"Training {name}…")
                    min_class   = int(y.value_counts().min())
                    test_size   = 0.20
                    do_stratify = min_class >= max(2, int(len(df)*test_size*0.10))
                    Xt, Xe, yt, ye = train_test_split(
                        X, y, test_size=test_size, random_state=42,
                        stratify=(y if do_stratify else None))
                    clf = Clf(**kw)
                    clf.fit(Xt, yt)
                    prob_e = clf.predict_proba(Xe)[:,1]
                    if not _synthetic_churn:
                        accuracy  = accuracy_score(ye, clf.predict(Xe))
                        auc_score = roc_auc_score(ye, prob_e) if ye.nunique()>1 else None
                    df["Churn_Prob"] = clf.predict_proba(X)[:,1]
                    model_name = (f"{name} (Synthetic)" if _synthetic_churn else name)
                    self._log("ok", f"Model trained: {name}")
                    if not _synthetic_churn and accuracy:
                        self._log("ok", f"Accuracy: {accuracy*100:.2f}%   AUC-ROC: {auc_score:.4f}" if auc_score else f"Accuracy: {accuracy*100:.2f}%")
                    break
                except Exception as ex:
                    self._log("warn", f"{name} failed: {ex}")
                    continue
        else:
            self._log("warn", f"ML skipped (rows={len(df)}, classes={y.nunique()}) — using heuristic")

        df["Churn_Prob"] = df["Churn_Prob"].clip(0,1).round(4)
        self._set_status("Calculating CLV and segments…", 64)
        self._log("section", "─── STEP 7: CLV & Segmentation ────────────────────────")

        # ── CLV & Segments ───────────────────────────────────────────────────────
        try:
            tc_clean = df["TotalCharges"].replace([np.inf,-np.inf],np.nan).clip(lower=0).fillna(0)
            df["Predicted_CLV"] = (tc_clean * (1-df["Churn_Prob"])).clip(lower=0)
        except Exception:
            df["Predicted_CLV"] = df["TotalCharges"].clip(lower=0).fillna(0)

        hi_risk = df["Churn_Prob"] >= 0.60
        hi_clv  = df["Predicted_CLV"] >= df["Predicted_CLV"].median()
        df["Segment"] = np.select(
            [hi_risk & hi_clv, hi_risk & ~hi_clv, ~hi_risk & hi_clv],
            ["High Risk – High Value","High Risk – Low Value","Low Risk – High Value"],
            default="Low Risk – Low Value")
        df["Segment"]   = df["Segment"].fillna("Low Risk – Low Value")
        df["Risk_Tier"] = pd.cut(
            df["Churn_Prob"], bins=[0,0.30,0.60,1.001],
            labels=["Low Risk (0-30%)","Medium Risk (30-60%)","High Risk (60-100%)"],
            include_lowest=True).astype(str).fillna("Low Risk (0-30%)")

        # ── KPIs ─────────────────────────────────────────────────────────────────
        total_cust     = max(len(df),1)
        churn_rate     = round(safe_float(df["Churn"].mean())*100,2)
        retention_rate = round(max(0.0,min(100.0,100.0-churn_rate)),2)
        avg_clv        = round(safe_float(df["Predicted_CLV"].mean()),0)
        avg_tenure     = round(safe_float(df["tenure"].mean()),1)
        avg_monthly    = round(safe_float(df["MonthlyCharges"].mean()),0)
        high_risk_n    = int((df["Churn_Prob"]>=0.60).sum())
        high_risk_pct  = round(high_risk_n/total_cust*100,1)
        med_risk_n     = int(((df["Churn_Prob"]>=0.30)&(df["Churn_Prob"]<0.60)).sum())
        low_risk_n     = int((df["Churn_Prob"]<0.30).sum())
        total_clv      = round(safe_float(df["Predicted_CLV"].sum()),0)
        hr_mask         = df["Churn_Prob"]>=0.60
        revenue_at_risk = round(safe_float((df.loc[hr_mask,"MonthlyCharges"]*12).sum()),0)
        tc_at_risk_hist = round(safe_float(df.loc[hr_mask,"TotalCharges"].sum()),0)

        if _synthetic_churn:
            acc_str = "N/A (synthetic)"
            auc_str = "N/A (synthetic)"
        else:
            acc_str = (f"{round(accuracy*100,2)}%" if accuracy  is not None else "N/A")
            auc_str = (f"{round(auc_score,4)}"    if auc_score is not None else "N/A")

        self._log("ok", f"KPIs computed: Churn={churn_rate}%, High-Risk={high_risk_n:,}, CLV=Rs.{total_clv:,.0f}")
        self._set_status("Building Excel report…", 72)
        self._log("section", "─── STEP 8: Building Excel Report ─────────────────────")

        # Store result
        self.result_data = dict(
            total_cust=total_cust, churn_rate=churn_rate,
            retention_rate=retention_rate, avg_clv=avg_clv,
            avg_tenure=avg_tenure, avg_monthly=avg_monthly,
            high_risk_n=high_risk_n, high_risk_pct=high_risk_pct,
            med_risk_n=med_risk_n, low_risk_n=low_risk_n,
            total_clv=total_clv, revenue_at_risk=revenue_at_risk,
            tc_at_risk_hist=tc_at_risk_hist, model_name=model_name,
            acc_str=acc_str, auc_str=auc_str, churn_source=churn_source,
            _synthetic_churn=_synthetic_churn,
            original_rows=original_rows, original_cols_n=original_cols_n,
        )

        # ── Build report ─────────────────────────────────────────────────────────
        def write_report(out_path):
            build_excel_report(
                df=df, original_df=original_df, out_path=out_path,
                file_path=fp, **self.result_data,
                id_col=id_col,
                FEATURES=FEATURES,
            )

        # Retry save
        attempt_path = sp
        for attempt in range(1, MAX_RETRIES+1):
            try:
                write_report(attempt_path)
                break
            except PermissionError:
                # Ask on main thread
                resolve = threading.Event()
                choice  = [None]

                def _ask_locked():
                    r = messagebox.askyesno(
                        "File Locked",
                        f"The file is currently open or locked:\n  {attempt_path}\n\n"
                        f"Yes  →  Auto-save with timestamp\n"
                        f"No   →  Choose a new location",
                        parent=self.root
                    )
                    choice[0] = r
                    resolve.set()

                self.root.after(0, _ask_locked)
                resolve.wait()

                if choice[0]:
                    base, ext = os.path.splitext(sp)
                    attempt_path = f"{base}_{time.strftime('%H%M%S')}{ext}"
                else:
                    new_path_event = threading.Event()
                    new_path_holder = [None]

                    def _pick_new():
                        p2 = filedialog.asksaveasfilename(
                            parent=self.root, title="Save Report",
                            defaultextension=".xlsx",
                            filetypes=[("Excel Workbook", "*.xlsx")])
                        new_path_holder[0] = p2
                        new_path_event.set()

                    self.root.after(0, _pick_new)
                    new_path_event.wait()
                    if not new_path_holder[0]:
                        raise RuntimeError("Save cancelled by user.")
                    attempt_path = new_path_holder[0]
            except Exception:
                raise

        self._set_status("Finalising…", 94)
        self._log("ok", f"Report saved: {os.path.basename(attempt_path)}")
        self._set_status(f"Done!  Report saved: {os.path.basename(attempt_path)}", 100)

        # ── Update KPI cards on main thread ──────────────────────────────────────
        self.root.after(0, lambda: self._show_kpi_cards(
            total_cust, churn_rate, retention_rate, high_risk_n,
            high_risk_pct, med_risk_n, low_risk_n, avg_tenure,
            avg_clv, revenue_at_risk, total_clv, avg_monthly,
            model_name, acc_str, auc_str, _synthetic_churn
        ))

        self._log("section", "─── COMPLETE ───────────────────────────────────────────")
        self._log("ok", f"Customers analysed       : {total_cust:,}")
        self._log("ok", f"Churn rate               : {churn_rate}%")
        self._log("ok", f"High-risk customers      : {high_risk_n:,}  ({high_risk_pct}%)")
        self._log("ok", f"Annualised risk revenue  : Rs.{revenue_at_risk:,.0f}")
        self._log("ok", f"Total CLV portfolio      : Rs.{total_clv:,.0f}")
        self._log("ok", f"Model                    : {model_name}")
        self._log("ok", f"Accuracy / AUC           : {acc_str} / {auc_str}")

        # Success dialog
        synthetic_note = ("\n⚠  NOTE: Synthetic churn labels used.\n"
                          "   ML accuracy suppressed — risk scores are valid.\n"
                          if _synthetic_churn else "")

        self.root.after(0, lambda: messagebox.showinfo(
            "✔  Report Generated Successfully!",
            f"Your report has been saved!\n\n"
            f"  {attempt_path}\n\n"
            f"{'═'*50}\n"
            f"  Customers Analysed     : {total_cust:,}\n"
            f"  Churn Rate             : {churn_rate}%\n"
            f"  Retention Rate         : {retention_rate}%\n"
            f"  High-Risk Customers    : {high_risk_n:,}  ({high_risk_pct}%)\n"
            f"  Risk Revenue (annual)  : Rs.{revenue_at_risk:,.0f}\n"
            f"  Total CLV Portfolio    : Rs.{total_clv:,.0f}\n"
            f"  ML Model               : {model_name}\n"
            f"  Accuracy / AUC-ROC     : {acc_str} / {auc_str}\n"
            f"{'═'*50}\n\n"
            f"Sheets in your report:\n"
            f"  • Dashboard  (14 KPIs + 6 charts)\n"
            f"  • Segment Summary\n"
            f"  • High Risk Customers\n"
            f"  • Processed Data\n"
            f"  • Raw Data\n"
            f"  • Data Quality Report\n"
            f"  • 7 Chart-source sheets\n"
            f"{synthetic_note}",
            parent=self.root
        ))

    # ── KPI cards ───────────────────────────────────────────────────────────────
    def _show_kpi_cards(self, total_cust, churn_rate, retention_rate,
                        high_risk_n, high_risk_pct, med_risk_n, low_risk_n,
                        avg_tenure, avg_clv, revenue_at_risk, total_clv,
                        avg_monthly, model_name, acc_str, auc_str,
                        _synthetic_churn):
        for w in self.kpi_frame.winfo_children():
            w.destroy()

        cards = [
            ("👥 Total Customers",        f"{total_cust:,}",               GUI["accent2"],  GUI["white"]),
            ("📉 Churn Rate",              f"{churn_rate}%",                GUI["red"],      GUI["white"]),
            ("💚 Retention Rate",          f"{retention_rate}%",            GUI["green"],    GUI["white"]),
            ("🔴 High Risk  (≥60%)",       f"{high_risk_n:,}\n({high_risk_pct}%)", GUI["red"], GUI["white"]),
            ("🟡 Medium Risk (30–60%)",    f"{med_risk_n:,}",               "#B7950B",       GUI["white"]),
            ("🟢 Low Risk  (<30%)",        f"{low_risk_n:,}",               GUI["green"],    GUI["white"]),
            ("⏱ Avg Tenure",              f"{avg_tenure} mo",              GUI["bg2"],      GUI["white"]),
            ("💰 Avg CLV",                 f"Rs.{avg_clv:,.0f}",            GUI["accent2"],  GUI["white"]),
            ("⚠ Risk Revenue/yr",         f"Rs.{revenue_at_risk:,.0f}",    GUI["red"],      GUI["white"]),
            ("📊 Total CLV Portfolio",    f"Rs.{total_clv:,.0f}",          GUI["purple"],   GUI["white"]),
            ("💳 Avg Monthly",             f"Rs.{avg_monthly:,.0f}",        GUI["bg2"],      GUI["white"]),
            ("🤖 ML Accuracy",            acc_str,
             GUI["warn_bg"] if _synthetic_churn else GUI["green"],
             GUI["warn_fg"] if _synthetic_churn else GUI["white"]),
        ]

        cols = 4
        for i, (lbl, val, bg, fg) in enumerate(cards):
            r, c = divmod(i, cols)
            card = tk.Frame(self.kpi_frame, bg=bg,
                            relief="flat", padx=14, pady=10)
            card.grid(row=r, column=c, padx=5, pady=5, sticky="nsew")
            self.kpi_frame.columnconfigure(c, weight=1)

            tk.Label(card, text=lbl, font=("Segoe UI", 9, "bold"),
                     bg=bg, fg=fg, anchor="w").pack(fill="x")
            tk.Label(card, text=val, font=("Segoe UI", 16, "bold"),
                     bg=bg, fg=fg, anchor="w").pack(fill="x", pady=(2, 0))

        # Model row
        model_frame = tk.Frame(self.kpi_frame, bg=GUI["bg3"],
                               relief="flat", padx=14, pady=8)
        model_frame.grid(row=(len(cards)//cols)+1, column=0,
                         columnspan=cols, padx=5, pady=5, sticky="ew")
        tk.Label(model_frame,
                 text=f"🤖  Model: {model_name}   |   Accuracy: {acc_str}   |   AUC-ROC: {auc_str}",
                 font=("Segoe UI", 10),
                 bg=GUI["bg3"], fg=GUI["text"], anchor="w").pack(fill="x")

    def _on_close(self):
        if self.running:
            if not messagebox.askyesno("Quit", "Analysis is running. Quit anyway?",
                                       parent=self.root):
                return
        self.root.destroy()

    def run(self):
        self.root.mainloop()


# ==============================================================================
#  ANALYTICS UTILITIES  (same logic as original — no GUI dependency)
# ==============================================================================
def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def clean_num(series):
    s = (series.astype(str).str.strip()
         .str.replace(r"[₹$€£¥]","",regex=True)
         .str.replace(r"\s+","",regex=True))
    def _fix(v):
        if re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", v):
            return v.replace(".","").replace(",",".")
        if re.search(r"\d,\d{3}[.,]", v):
            return v.replace(",","")
        if re.search(r"\d,\d{2},\d{3}", v):
            return v.replace(",","")
        return v.replace(",","")
    s = s.apply(_fix)
    s = s.str.replace(r"[^\d.\-]","",regex=True)
    return pd.to_numeric(s, errors="coerce")


def find_col(df, *keywords):
    cols_norm = {col: norm(col) for col in df.columns}
    for kw in keywords:
        for col, n in cols_norm.items():
            if kw in n:
                return col
    return None


def safe_cut(series, bins=6):
    try:
        s = pd.to_numeric(series, errors="coerce").dropna()
        if len(s)==0 or s.nunique()<2:
            raise ValueError("insufficient")
        cuts = pd.cut(s, bins=min(bins, s.nunique()), duplicates="drop")
        res  = cuts.value_counts().reset_index()
        res.columns = ["Range","Count"]
        res["_left"] = res["Range"].apply(
            lambda iv: iv.left if hasattr(iv,"left") else float("-inf"))
        res = res.sort_values("_left").drop(columns="_left").reset_index(drop=True)
        res["Range"] = res["Range"].astype(str)
        return res[res["Count"]>0].reset_index(drop=True)
    except Exception:
        return pd.DataFrame({"Range":["All Values"],"Count":[int(series.notna().sum())]})


def safe_float(v, default=0.0):
    try:
        f = float(v)
        return f if np.isfinite(f) else default
    except Exception:
        return default


def _try_csv(path, enc, sep):
    shared = dict(encoding=enc, dtype=str, on_bad_lines="skip")
    for kw in [{"engine":"c","low_memory":False,"sep":sep or ","},
               {"engine":"python","sep":sep}]:
        try:
            df = pd.read_csv(path, **shared, **kw)
            if not df.empty and len(df.columns)>=2:
                return df
        except Exception:
            pass
    return None


def load_file(path):
    ext = os.path.splitext(path)[1].lower()
    errors = []
    if ext in (".xlsx",".xls",".xlsm",".xlsb",".ods"):
        engines = (["xlrd","openpyxl"] if ext==".xls"
                   else ["odf","openpyxl"] if ext==".ods"
                   else ["openpyxl"])
        engines.append(None)
        for eng in engines:
            try:
                kw = {} if eng is None else {"engine":eng}
                xl = pd.ExcelFile(path,**kw)
                for sheet in xl.sheet_names:
                    try:
                        df = xl.parse(sheet,dtype=str)
                        df = df.dropna(how="all").dropna(axis=1,how="all")
                        if len(df)>=1 and len(df.columns)>=2:
                            return df
                    except Exception as ex:
                        errors.append(str(ex))
            except Exception as ex:
                errors.append(str(ex))
        raise RuntimeError("Cannot open Excel/ODS file.\n"+"; ".join(errors[:3]))

    ENCODINGS  = ("utf-8-sig","utf-8","latin1","cp1252","iso-8859-1","utf-16")
    SEPARATORS = (None,",",";","\t","|",":")
    for enc in ENCODINGS:
        for sep in SEPARATORS:
            try:
                df = _try_csv(path,enc,sep)
                if df is not None:
                    df = df.dropna(how="all").dropna(axis=1,how="all")
                    if len(df)>=1 and len(df.columns)>=2:
                        return df
            except Exception as ex:
                errors.append(str(ex))
    raise RuntimeError("Could not parse file after all attempts.\n"+"; ".join(list(dict.fromkeys(errors))[:3]))


# ==============================================================================
#  EXCEL REPORT BUILDER  (identical to original, extracted to top-level fn)
# ==============================================================================
def build_excel_report(df, original_df, out_path, file_path,
                       total_cust, churn_rate, retention_rate,
                       avg_clv, avg_tenure, avg_monthly,
                       high_risk_n, high_risk_pct, med_risk_n, low_risk_n,
                       total_clv, revenue_at_risk, tc_at_risk_hist,
                       model_name, acc_str, auc_str, churn_source,
                       _synthetic_churn, original_rows, original_cols_n,
                       id_col, FEATURES, **kwargs):

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        wb = writer.book

        C = {
            "navy":"#0B1F3A","navy2":"#1A3A5C","teal":"#007C71","teal2":"#00A896",
            "slate":"#2B4570","steel":"#3D6B99","red":"#C0392B","red2":"#E74C3C",
            "orange":"#D35400","amber":"#E67E22","green":"#1E8449","green2":"#27AE60",
            "blue":"#1565C0","blue2":"#1976D2","purple":"#6C3483","violet":"#8E44AD",
            "gold":"#B7950B","yellow":"#F39C12","lgrey":"#F0F4F8","mgrey":"#C8D0DC",
            "dgrey":"#5D6D7E","white":"#FFFFFF","warn_bg":"#FFF3CD","warn_fc":"#856404",
        }
        CHART8 = ["#007C71","#1565C0","#E67E22","#C0392B","#6C3483","#1E8449","#D35400","#3D6B99"]

        def MF(bold=False,italic=False,sz=10,fc=C["navy"],bg=None,
               align="center",valign="vcenter",wrap=False,
               border=1,bc=C["mgrey"],num_fmt=None,left=False):
            d={"font_name":"Calibri","font_size":sz,"bold":bold,"italic":italic,
               "font_color":fc,"valign":valign,"align":"left" if left else align,"text_wrap":wrap}
            if bg:      d["bg_color"]   = bg
            if border:  d.update({"border":border,"border_color":bc})
            if num_fmt: d["num_format"] = num_fmt
            return wb.add_format(d)

        F = {
            "title":      MF(bold=True,sz=18,fc=C["white"],bg=C["navy"],border=0),
            "subtitle":   MF(italic=True,sz=10,fc=C["slate"],bg=C["lgrey"],border=0),
            "warn_banner":MF(bold=True,sz=10,fc=C["warn_fc"],bg=C["warn_bg"],border=1,bc=C["amber"],wrap=True),
            "col_hdr":    MF(bold=True,sz=10,fc=C["white"],bg=C["navy2"]),
            "col_hdr_t":  MF(bold=True,sz=10,fc=C["white"],bg=C["teal"]),
            "col_hdr_s":  MF(bold=True,sz=10,fc=C["white"],bg=C["slate"]),
            "col_hdr_p":  MF(bold=True,sz=10,fc=C["white"],bg=C["purple"]),
            "sec_banner": MF(bold=True,sz=10,fc=C["white"],bg=C["teal2"]),
            "kpi_lbl_b":  MF(bold=True,sz=9,fc=C["white"],bg=C["navy2"],border=0),
            "kpi_lbl_r":  MF(bold=True,sz=9,fc=C["white"],bg=C["red"],border=0),
            "kpi_lbl_g":  MF(bold=True,sz=9,fc=C["white"],bg=C["teal"],border=0),
            "kpi_lbl_a":  MF(bold=True,sz=9,fc=C["white"],bg=C["amber"],border=0),
            "kpi_lbl_p":  MF(bold=True,sz=9,fc=C["white"],bg=C["purple"],border=0),
            "kpi_lbl_sl": MF(bold=True,sz=9,fc=C["white"],bg=C["slate"],border=0),
            "kpi_lbl_w":  MF(bold=True,sz=9,fc=C["navy"],bg=C["warn_bg"],border=0),
            "kpi_val":    MF(bold=True,sz=16,fc=C["navy"],bg=C["white"],bc=C["mgrey"]),
            "kpi_red":    MF(bold=True,sz=16,fc=C["red"],bg=C["white"],bc=C["mgrey"]),
            "kpi_green":  MF(bold=True,sz=16,fc=C["green"],bg=C["white"],bc=C["mgrey"]),
            "kpi_amber":  MF(bold=True,sz=16,fc=C["orange"],bg=C["white"],bc=C["mgrey"]),
            "kpi_purple": MF(bold=True,sz=16,fc=C["purple"],bg=C["white"],bc=C["mgrey"]),
            "kpi_warn":   MF(bold=True,sz=10,fc=C["warn_fc"],bg=C["warn_bg"],bc=C["amber"],wrap=True),
            "cell":       MF(sz=10,fc=C["navy"],bg=C["white"],bc=C["mgrey"]),
            "cell_alt":   MF(sz=10,fc=C["navy"],bg=C["lgrey"],bc=C["mgrey"]),
            "cell_l":     MF(sz=10,fc=C["navy"],bg=C["white"],bc=C["mgrey"],left=True),
            "cell_l_alt": MF(sz=10,fc=C["navy"],bg=C["lgrey"],bc=C["mgrey"],left=True),
            "num":        MF(sz=10,fc=C["navy"],bg=C["white"],bc=C["mgrey"],num_fmt="#,##0"),
            "num_alt":    MF(sz=10,fc=C["navy"],bg=C["lgrey"],bc=C["mgrey"],num_fmt="#,##0"),
            "dec2":       MF(sz=10,fc=C["navy"],bg=C["white"],bc=C["mgrey"],num_fmt="#,##0.00"),
            "dec2_alt":   MF(sz=10,fc=C["navy"],bg=C["lgrey"],bc=C["mgrey"],num_fmt="#,##0.00"),
            "dec4":       MF(sz=10,fc=C["navy"],bg=C["white"],bc=C["mgrey"],num_fmt="0.0000"),
            "dec4_alt":   MF(sz=10,fc=C["navy"],bg=C["lgrey"],bc=C["mgrey"],num_fmt="0.0000"),
            "pct2":       MF(sz=10,fc=C["navy"],bg=C["white"],bc=C["mgrey"],num_fmt="0.00%"),
            "pct2_alt":   MF(sz=10,fc=C["navy"],bg=C["lgrey"],bc=C["mgrey"],num_fmt="0.00%"),
            "c_red":      MF(bold=True,sz=10,fc=C["red2"],bg="#FDECEA",bc=C["mgrey"],num_fmt="0.00%"),
            "c_amber":    MF(bold=True,sz=10,fc=C["orange"],bg="#FDF3E7",bc=C["mgrey"],num_fmt="0.00%"),
            "c_green":    MF(bold=True,sz=10,fc=C["green"],bg="#EAFAF1",bc=C["mgrey"],num_fmt="0.00%"),
        }

        PROC_NUM_FMTS = {
            "tenure":"#,##0.0","MonthlyCharges":"#,##0.00","TotalCharges":"#,##0.00",
            "AvgMonthlySpend":"#,##0.00","ValueScore":"0.0000","LoyaltyScore":"#,##0.00",
            "SpendVariance":"0.0000","Churn":"0","Churn_Prob":"0.00%","Predicted_CLV":"#,##0.00",
        }

        def write_table(ws,df_t,r0,c0=0,hdr_fmt="col_hdr",col_fmts=None,row_height=20):
            ws.set_row(r0,22)
            for ci,col in enumerate(df_t.columns):
                ws.write(r0,c0+ci,col,F[hdr_fmt])
            df_t=df_t.reset_index(drop=True)
            for ri in range(len(df_t)):
                ws.set_row(r0+1+ri,row_height)
                alt=(ri%2==1)
                for ci,col in enumerate(df_t.columns):
                    val=df_t.iloc[ri,ci]
                    if pd.isna(val) or str(val)=="nan": val=""
                    if col_fmts and col in col_fmts:
                        fk=col_fmts[col]
                        if callable(fk):    fmt=fk(val,alt)
                        else:
                            ak=fk+"_alt"
                            fmt=F[ak if (alt and ak in F) else fk]
                    else:
                        fmt=F["cell_alt" if alt else "cell"]
                    ws.write(r0+1+ri,c0+ci,val,fmt)

        def auto_col_widths(df_src,min_w=10,max_w=60,sample=500):
            sd=df_src.iloc[:sample] if len(df_src)>sample else df_src
            w=[]
            for col in df_src.columns:
                hw=len(str(col))+2
                vals=sd[col].astype(str).str.len()
                dw=int(vals.max())+2 if len(vals) else 0
                w.append(max(min_w,min(max(hw,dw),max_w)))
            return w

        def data_sheet(name,df_t,title_text,subtitle_text,hdr_color="col_hdr_t",min_w=18,max_w=55):
            sname=_trunc(name); ws=wb.add_worksheet(sname)
            ws.hide_gridlines(2); ws.set_zoom(90)
            for ci,ww in enumerate(auto_col_widths(df_t,min_w,max_w)):
                ws.set_column(ci,ci,ww)
            for r,h in [(0,5),(1,38),(2,18),(3,5),(4,22)]:
                ws.set_row(r,h)
            lc=max(len(df_t.columns)-1,1)
            ws.merge_range(1,0,1,lc,title_text,F["title"])
            ws.merge_range(2,0,2,lc,subtitle_text,F["subtitle"])
            write_table(ws,df_t,4,c0=0,hdr_fmt=hdr_color,row_height=22)
            return sname,len(df_t)

        def add_chart(chart_type,src_sheet,nrows,title,ws_dest,cell_ref,
                      fill_color=None,fill_colors=None,w=500,h=295,show_legend=False):
            if nrows<=0: return
            ch=wb.add_chart({"type":chart_type})
            series={"categories":[src_sheet,5,0,4+nrows,0],"values":[src_sheet,5,1,4+nrows,1],
                    "data_labels":{"value":True,"font":{"bold":True,"size":9,"color":C["navy"]}}}
            if chart_type in ("column","bar"): series["gap"]=60
            if fill_colors:
                series["points"]=[{"fill":{"color":fill_colors[i%len(fill_colors)]}} for i in range(nrows)]
            elif fill_color:
                series["fill"]={"color":fill_color}
            ch.add_series(series)
            ch.set_title({"name":title,"name_font":{"bold":True,"size":11,"color":C["navy"]}})
            ch.set_legend({"position":"bottom"} if show_legend else {"none":True})
            ch.set_chartarea({"border":{"color":C["mgrey"]},"fill":{"color":C["white"]}})
            ch.set_plotarea({"fill":{"color":C["lgrey"]}})
            ch.set_style(2); ch.set_size({"width":w,"height":h})
            ws_dest.insert_chart(cell_ref,ch,{"x_offset":5,"y_offset":5})

        def write_full_sheet(ws,df_src,hdr_fmt_key,zoom=90,row_h=18,hdr_h=22,
                             min_w=10,max_w=60,apply_num_fmt=False):
            ws.hide_gridlines(2); ws.set_zoom(zoom)
            for ci,ww in enumerate(auto_col_widths(df_src,min_w,max_w)):
                ws.set_column(ci,ci,ww)
            ws.set_row(0,hdr_h)
            for ci,col in enumerate(df_src.columns):
                ws.write(0,ci,str(col),F[hdr_fmt_key])
            for ri in range(len(df_src)):
                ws.set_row(ri+1,row_h); alt=(ri%2==1)
                for ci,col in enumerate(df_src.columns):
                    val=df_src.iloc[ri,ci]
                    if pd.isna(val) or str(val)=="nan": val=""
                    if apply_num_fmt and col in PROC_NUM_FMTS:
                        fmt=MF(sz=10,fc=C["navy"],bg=C["lgrey"] if alt else C["white"],
                               bc=C["mgrey"],num_fmt=PROC_NUM_FMTS[col])
                    else:
                        fmt=MF(sz=10,fc=C["navy"],bg=C["lgrey"] if alt else C["white"],bc=C["mgrey"])
                    ws.write(ri+1,ci,val,fmt)
            ws.autofilter(0,0,len(df_src),len(df_src.columns)-1)
            ws.freeze_panes(1,0)

        # ── Distribution data ────────────────────────────────────────────────────
        seg_data=df["Segment"].value_counts().reset_index(); seg_data.columns=["Segment","Count"]
        risk_order=["Low Risk (0-30%)","Medium Risk (30-60%)","High Risk (60-100%)"]
        risk_data=df["Risk_Tier"].value_counts().reset_index(); risk_data.columns=["Risk Tier","Count"]
        risk_data["Risk Tier"]=pd.Categorical(risk_data["Risk Tier"],categories=risk_order,ordered=True)
        risk_data=risk_data.sort_values("Risk Tier").reset_index(drop=True)
        churn_dist=safe_cut(df["Churn_Prob"],6); clv_dist=safe_cut(df["Predicted_CLV"],6)
        tenure_dist=safe_cut(df["tenure"],6); monthly_dist=safe_cut(df["MonthlyCharges"],6)

        sn_seg,seg_n   = data_sheet("Data – Segments",  seg_data,"Customer Segment Distribution","Count per segment.","col_hdr_t")
        sn_risk,risk_n = data_sheet("Data – Risk Tiers",risk_data,"Risk Tier Distribution","Low/Medium/High buckets.","col_hdr_p")
        sn_chd,chd_n   = data_sheet("Data – Churn Prob",churn_dist,"Churn Probability Buckets","0=safe 1=leaving.","col_hdr_t")
        sn_clv,clvd_n  = data_sheet("Data – CLV",       clv_dist,"CLV Buckets","Retention-weighted proxy.","col_hdr_s")
        sn_ten,tend_n  = data_sheet("Data – Tenure",    tenure_dist,"Tenure Distribution","Months with company.","col_hdr_s")
        sn_mnd,mnd_n   = data_sheet("Data – Monthly",   monthly_dist,"Monthly Charges Dist","Current monthly amounts.","col_hdr_p")

        # ── Sheet 1: Dashboard ───────────────────────────────────────────────────
        dash=wb.add_worksheet(_trunc("Dashboard"))
        dash.hide_gridlines(2); dash.set_zoom(80)
        CARD_W=18; N_CARDS=7
        dash.set_column(0,0,1.5)
        for c in range(1,N_CARDS*2+6): dash.set_column(c,c,CARD_W)
        KLR1=4;KVR1=5;KLR2=7;KVR2=8;WARN_ROW=10
        for r,h in {0:5,1:46,2:20,3:8,KLR1:18,KVR1:48,6:8,KLR2:18,KVR2:48,9:12}.items():
            dash.set_row(r,h)
        TLC=N_CARDS*2
        dash.merge_range(1,1,1,TLC,"  CUSTOMER CHURN ANALYTICS  ·  EXECUTIVE DASHBOARD",F["title"])
        dash.merge_range(2,1,2,TLC,
            f"File: {os.path.basename(file_path)}   |   Model: {model_name}   |   "
            f"Accuracy: {acc_str}   |   AUC-ROC: {auc_str}   |   Customers: {total_cust:,}",
            F["subtitle"])

        if _synthetic_churn:
            dash.set_row(WARN_ROW,44)
            dash.merge_range(WARN_ROW,1,WARN_ROW,TLC,
                "⚠  SYNTHETIC LABELS NOTICE:  No real churn column detected.  "
                "Churn labels were generated using percentile risk-score formula.  "
                "ML accuracy and AUC-ROC suppressed — metrics would only reflect "
                "formula self-consistency.  Risk scores & segmentation are valid.",
                F["warn_banner"])
            car=WARN_ROW+2
        else:
            car=WARN_ROW

        def kpi_card(ws,label,value,lr,vr,ci,lf="kpi_lbl_b",vf="kpi_val"):
            cs=1+ci*2; ce=cs+1
            ws.merge_range(lr,cs,lr,ce,label,F[lf])
            ws.merge_range(vr,cs,vr,ce,value,F[vf])

        for i,(lbl,val,lf,vf) in enumerate([
            ("TOTAL CUSTOMERS",      f"{total_cust:,}",                 "kpi_lbl_b","kpi_val"),
            ("CHURN RATE",           f"{churn_rate}%",                  "kpi_lbl_r","kpi_red"),
            ("RETENTION RATE",       f"{retention_rate}%",              "kpi_lbl_g","kpi_green"),
            ("HIGH RISK  (≥60%)",    f"{high_risk_n:,} ({high_risk_pct}%)","kpi_lbl_r","kpi_red"),
            ("MEDIUM RISK (30-60%)", f"{med_risk_n:,}",                 "kpi_lbl_a","kpi_amber"),
            ("LOW RISK  (<30%)",     f"{low_risk_n:,}",                 "kpi_lbl_g","kpi_green"),
            ("AVG TENURE (months)",  f"{avg_tenure}",                   "kpi_lbl_sl","kpi_val"),
        ]): kpi_card(dash,lbl,val,KLR1,KVR1,i,lf,vf)

        if _synthetic_churn:
            af,vf2="kpi_lbl_w","kpi_warn"; ad,auc_d="N/A — Synthetic","N/A — Synthetic"
        else:
            af,vf2="kpi_lbl_g","kpi_green"; af2="kpi_lbl_p"
            ad,auc_d=acc_str,auc_str

        for i,(lbl,val,lf,vf) in enumerate([
            ("AVG CUSTOMER CLV",        f"Rs.{avg_clv:,.0f}",          "kpi_lbl_g","kpi_green"),
            ("ANNUALISED RISK REVENUE", f"Rs.{revenue_at_risk:,.0f}",  "kpi_lbl_r","kpi_red"),
            ("TOTAL CLV PORTFOLIO",     f"Rs.{total_clv:,.0f}",        "kpi_lbl_p","kpi_purple"),
            ("AVG MONTHLY CHARGE",      f"Rs.{avg_monthly:,.0f}",      "kpi_lbl_b","kpi_val"),
            ("ML MODEL",                model_name,                     "kpi_lbl_sl","kpi_val"),
            ("MODEL ACCURACY",          ad,                             af,vf2),
            ("AUC-ROC SCORE",           auc_d,                         af,vf2),
        ]): kpi_card(dash,lbl,val,KLR2,KVR2,i,lf,vf)

        r1=car+1; r2=r1+17; r3=r2+17
        def rc(r,c): return f"{c}{r+1}"
        add_chart("column",sn_seg,seg_n,"Customer Segments by Count",dash,rc(r1,"B"),fill_colors=CHART8,w=540,h=310)
        add_chart("doughnut",sn_risk,risk_n,"Risk Tier Breakdown",dash,rc(r1,"J"),fill_colors=["#27AE60","#E67E22","#C0392B"],w=480,h=310,show_legend=True)
        add_chart("column",sn_chd,chd_n,"Churn Probability Distribution",dash,rc(r2,"B"),fill_colors=["#1E8449","#52BE80","#F4D03F","#E67E22","#E74C3C","#C0392B"],w=540,h=310)
        add_chart("bar",sn_clv,clvd_n,"Predicted CLV (retention-weighted proxy)",dash,rc(r2,"J"),fill_colors=CHART8,w=480,h=310)
        add_chart("column",sn_ten,tend_n,"Customer Tenure Distribution (months)",dash,rc(r3,"B"),fill_colors=["#1565C0","#1976D2","#42A5F5","#64B5F6","#90CAF9","#BBDEFB"],w=540,h=310)
        add_chart("column",sn_mnd,mnd_n,"Monthly Charges Distribution",dash,rc(r3,"J"),fill_colors=["#6C3483","#8E44AD","#A569BD","#C39BD3","#D7BDE2","#E8DAEF"],w=480,h=310)

        # ── Sheet 2: Segment Summary ─────────────────────────────────────────────
        ws2=wb.add_worksheet(_trunc("Segment Summary"))
        ws2.hide_gridlines(2); ws2.set_zoom(95)
        ws2.set_column(0,0,1.5); ws2.set_column(1,1,32); ws2.set_column(2,9,19)
        for r,h in [(0,5),(1,46),(2,20),(3,8)]: ws2.set_row(r,h)
        ws2.merge_range(1,1,1,TLC,"SEGMENT PERFORMANCE SUMMARY  ·  Churn Analytics",F["title"])
        ws2.merge_range(2,1,2,TLC,"One row per segment.  Churn Rate: RED >60%  AMBER 30-60%  GREEN <30%",F["subtitle"])

        summary=df.groupby("Segment",as_index=False).agg(
            Customers=("Churn","count"),Churned=("Churn","sum"),
            Churn_Rate=("Churn","mean"),Avg_CLV=("Predicted_CLV","mean"),
            Total_CLV=("Predicted_CLV","sum"),Avg_Tenure=("tenure","mean"),
            Avg_Monthly=("MonthlyCharges","mean"),Avg_Risk_Prob=("Churn_Prob","mean"),
        ).sort_values("Churn_Rate",ascending=False).reset_index(drop=True)
        summary.columns=["Segment","Customers","Churned","Churn Rate","Avg CLV (Rs.)","Total CLV (Rs.)","Avg Tenure (mo)","Avg Monthly (Rs.)","Avg Risk Score"]

        def cr_fmt(val,alt):
            try:
                v=float(val)
                if v>0.6: return F["c_red"]
                if v>0.3: return F["c_amber"]
                return F["c_green"]
            except: return F["cell"]

        write_table(ws2,summary,4,c0=1,hdr_fmt="col_hdr",
            col_fmts={"Segment":lambda v,a:F["cell_l_alt" if a else "cell_l"],
                      "Customers":"num","Churned":"num","Churn Rate":cr_fmt,
                      "Avg CLV (Rs.)":"dec2","Total CLV (Rs.)":"dec2",
                      "Avg Tenure (mo)":"dec2","Avg Monthly (Rs.)":"dec2",
                      "Avg Risk Score":"dec4"},row_height=22)

        # ── Sheet 3: High Risk ───────────────────────────────────────────────────
        ws3=wb.add_worksheet(_trunc("High Risk Customers"))
        ws3.hide_gridlines(2); ws3.set_zoom(90)
        ws3.set_column(0,0,1.5)
        for r,h in [(0,5),(1,46),(2,20),(3,8)]: ws3.set_row(r,h)
        ws3.merge_range("B2:M2","HIGH-RISK CUSTOMERS  ·  Priority Retention Action List",F["title"])
        ws3.merge_range("B3:M3","Customers with churn probability ≥ 60%.  Sorted highest risk first.",F["subtitle"])

        hr_base=["tenure","MonthlyCharges","TotalCharges","Predicted_CLV","Churn_Prob","Segment","Risk_Tier"]
        hr_cols=([id_col] if id_col else [])+hr_base
        hr_cols=[c for c in hr_cols if c in df.columns]
        hr=(df[df["Churn_Prob"]>=0.60][hr_cols]
            .sort_values("Churn_Prob",ascending=False)
            .head(1000).reset_index(drop=True))

        cw3={"tenure":14,"MonthlyCharges":22,"TotalCharges":22,"Predicted_CLV":22,
             "Churn_Prob":18,"Segment":32,"Risk_Tier":24}
        if hr.empty:
            ws3.write(4,1,"No customers have churn probability ≥ 60%.",F["c_green"])
        else:
            for ci,col in enumerate(hr.columns):
                ws3.set_column(1+ci,1+ci,cw3.get(col,22))
            def pf(val,alt):
                try:
                    v=float(val)
                    return F["c_red"] if v>=0.80 else F["c_amber"]
                except: return F["pct2_alt" if alt else "pct2"]
            hf={"tenure":"num","MonthlyCharges":"num","TotalCharges":"num",
                "Predicted_CLV":"dec2","Churn_Prob":pf,
                "Segment":lambda v,a:F["cell_l_alt" if a else "cell_l"],
                "Risk_Tier":lambda v,a:F["cell_alt" if a else "cell"]}
            if id_col: hf[id_col]=lambda v,a:F["cell_l_alt" if a else "cell_l"]
            write_table(ws3,hr,4,c0=1,hdr_fmt="col_hdr",col_fmts=hf,row_height=20)
            ws3.autofilter(4,1,4+len(hr),1+len(hr.columns)-1)
            ws3.freeze_panes(5,0)

        # ── Sheet 4: Processed Data ──────────────────────────────────────────────
        pb=["tenure","MonthlyCharges","TotalCharges","AvgMonthlySpend","ValueScore",
            "LoyaltyScore","SpendVariance","Churn","Churn_Prob","Predicted_CLV","Segment","Risk_Tier"]
        pc=([id_col] if id_col else [])+pb
        pc=[c for c in pc if c in df.columns]
        pd_df=df[pc].copy()
        pd_df.to_excel(writer,sheet_name=_trunc("Processed Data"),index=False)
        write_full_sheet(writer.sheets[_trunc("Processed Data")],pd_df,"col_hdr_t",apply_num_fmt=True)

        # ── Sheet 5: Raw Data ────────────────────────────────────────────────────
        original_df.to_excel(writer,sheet_name=_trunc("Raw Data"),index=False)
        write_full_sheet(writer.sheets[_trunc("Raw Data")],original_df,"col_hdr_s")

        # ── Sheet 6: Data Quality Report ─────────────────────────────────────────
        ws6=wb.add_worksheet(_trunc("Data Quality Report"))
        ws6.hide_gridlines(2); ws6.set_zoom(95)
        ws6.set_column(0,0,1.5); ws6.set_column(1,1,38); ws6.set_column(2,2,65)
        for r,h in [(0,5),(1,46),(2,20),(3,8)]: ws6.set_row(r,h)
        ws6.merge_range("B2:C2","DATA QUALITY & PIPELINE REPORT",F["title"])
        ws6.merge_range("B3:C3","Full audit trail: file details · column detection · model metrics · KPIs · methodology",F["subtitle"])

        def sec(ws,row,text):
            ws.merge_range(row,1,row,2,text,F["sec_banner"]); ws.set_row(row,20)

        mn = ("SUPPRESSED — synthetic labels" if _synthetic_churn else acc_str)
        an = ("SUPPRESSED — see Accuracy note" if _synthetic_churn else auc_str)

        pipeline_rows=[
            ("sec","FILE INFORMATION"),
            ("row","Source File",          os.path.basename(file_path)),
            ("row","Full Path",            file_path),
            ("row","Original Rows",        f"{original_rows:,}"),
            ("row","Original Columns",     f"{original_cols_n}"),
            ("row","Rows After Cleaning",  f"{total_cust:,}"),
            ("row","Rows Removed",         f"{original_rows-total_cust:,}"),
            ("sec","COLUMN DETECTION"),
            ("row","Churn Source",         churn_source),
            ("row","Churn Labels Valid",   "NO — synthetic" if _synthetic_churn else "YES — real"),
            ("row","Tenure Column",        "detected" if "tenure" in df.columns else "fallback"),
            ("row","Customer ID Column",   id_col or "Not found"),
            ("sec","MACHINE LEARNING MODEL"),
            ("row","Model Used",           model_name),
            ("row","Features Used",        ", ".join(FEATURES)),
            ("row","Accuracy",             mn),
            ("row","AUC-ROC Score",        an),
            ("row","Metrics Suppressed",   "YES — synthetic" if _synthetic_churn else "NO"),
            ("sec","METHODOLOGY NOTES"),
            ("row","CLV Methodology","TotalCharges × (1 – Churn_Prob).  NOT DCF CLV.  Relative ranking only."),
            ("row","Revenue at Risk",f"MonthlyCharges × 12 for Churn_Prob ≥ 60%.  Historical at risk: Rs.{tc_at_risk_hist:,.0f}."),
            ("row","Segmentation","2×2: Risk (≥0.60) × Value (≥median CLV)"),
            ("row","Synthetic Formula","Risk=tenure↓×0.40+monthly↑×0.30+variance↑×0.20+total↓×0.10.  Top 35%→Churned."),
            ("sec","OUTPUT KPIs"),
            ("row","Total Customers",            f"{total_cust:,}"),
            ("row","Churn Rate",                 f"{churn_rate}%"),
            ("row","Retention Rate",             f"{retention_rate}%"),
            ("row","High-Risk (≥60%)",           f"{high_risk_n:,}  ({high_risk_pct}%)"),
            ("row","Medium-Risk (30-60%)",        f"{med_risk_n:,}"),
            ("row","Low-Risk (<30%)",             f"{low_risk_n:,}"),
            ("row","Annualised Risk Revenue",     f"Rs.{revenue_at_risk:,.0f}"),
            ("row","Historical TotalCharges@Risk",f"Rs.{tc_at_risk_hist:,.0f}"),
            ("row","Total CLV Portfolio",         f"Rs.{total_clv:,.0f}"),
            ("row","Avg CLV per Customer",        f"Rs.{avg_clv:,.0f}"),
            ("row","Avg Tenure",                  f"{avg_tenure} months"),
            ("row","Avg Monthly Charge",          f"Rs.{avg_monthly:,.0f}"),
        ]
        r=4; dri=0
        for item in pipeline_rows:
            if item[0]=="sec": sec(ws6,r,item[1]); dri=0
            else:
                alt=(dri%2==1); ws6.set_row(r,20)
                ws6.write(r,1,item[1],F["cell_l_alt" if alt else "cell_l"])
                ws6.write(r,2,item[2],F["cell_alt"   if alt else "cell"])
                dri+=1
            r+=1


# ==============================================================================
#  ENTRY POINT
# ==============================================================================
def main():
    # Splash
    splash = SplashScreen()
    msgs = [
        (20,  "Loading dependencies…"),
        (45,  "Initialising ML libraries…"),
        (70,  "Setting up analytics engine…"),
        (90,  "Preparing user interface…"),
        (100, "Ready!"),
    ]
    for pct, msg in msgs:
        splash.update(pct, msg)
        time.sleep(0.22)

    splash.close()

    # Launch main app
    app = ChurnApp()
    app.run()


if __name__ == "__main__":
    main()

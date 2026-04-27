# ==============================
# CUSTOMER ANALYTICS TOOL
# Production-Ready, Robust, Flexible
# ==============================

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import warnings
import os
import sys
from pathlib import Path
from datetime import datetime
import traceback

warnings.filterwarnings("ignore")

# Scikit-learn imports with fallback
try:
    from sklearn.model_selection import train_test_split
    from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
    from sklearn.linear_model import Ridge
    from sklearn.preprocessing import StandardScaler
    from sklearn.metrics import accuracy_score
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False


# ==============================
# CONFIGURATION
# ==============================
class Config:
    """Central configuration for the application."""
    APP_NAME = "Customer Analytics Tool"
    VERSION = "2.0"
    MIN_ROWS = 3
    MAX_ROWS_WARN = 100000
    SUPPORTED_EXTENSIONS = {
        '.csv': 'CSV',
        '.xlsx': 'Excel',
        '.xls': 'Excel (Legacy)',
        '.xlsm': 'Excel (Macro)',
        '.xlsb': 'Excel (Binary)',
        '.json': 'JSON',
        '.parquet': 'Parquet',
        '.pq': 'Parquet'
    }
    ENCODINGS = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1', 'utf-16', 'ascii']
    
    # Column name mappings (lowercase keys)
    COLUMN_MAPPINGS = {
        # Tenure
        'tenure': 'tenure', 'tenure_months': 'tenure', 'customer_tenure': 'tenure',
        'months': 'tenure', 'duration': 'tenure', 'time_as_customer': 'tenure',
        'account_length': 'tenure', 'customer_age': 'tenure', 'loyalty_months': 'tenure',
        
        # Monthly Charges
        'monthlycharges': 'MonthlyCharges', 'monthly_charges': 'MonthlyCharges',
        'monthly_charge': 'MonthlyCharges', 'monthly_fee': 'MonthlyCharges',
        'monthly_payment': 'MonthlyCharges', 'monthly_amount': 'MonthlyCharges',
        'monthly_spend': 'MonthlyCharges', 'monthly_bill': 'MonthlyCharges',
        'monthly_cost': 'MonthlyCharges', 'recurring_charge': 'MonthlyCharges',
        
        # Total Charges
        'totalcharges': 'TotalCharges', 'total_charges': 'TotalCharges',
        'total_charge': 'TotalCharges', 'total_payment': 'TotalCharges',
        'lifetime_value': 'TotalCharges', 'total_spend': 'TotalCharges',
        'total_amount': 'TotalCharges', 'cumulative_charges': 'TotalCharges',
        'total_revenue': 'TotalCharges', 'customer_value': 'TotalCharges',
        
        # Churn
        'churn': 'Churn', 'churned': 'Churn', 'is_churn': 'Churn',
        'has_churned': 'Churn', 'customer_churn': 'Churn', 'attrition': 'Churn',
        'left': 'Churn', 'exited': 'Churn', 'cancelled': 'Churn',
        'terminated': 'Churn', 'inactive': 'Churn', 'lost': 'Churn',
        'status': 'Churn', 'customer_status': 'Churn', 'active': 'Churn'
    }


# ==============================
# LOGGING UTILITY
# ==============================
class Logger:
    """Simple logging for debugging without console."""
    def __init__(self):
        self.logs = []
    
    def log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.logs.append(f"[{timestamp}] [{level}] {message}")
    
    def get_logs(self):
        return "\n".join(self.logs)
    
    def clear(self):
        self.logs = []


logger = Logger()


# ==============================
# FILE LOADING
# ==============================
class FileLoader:
    """Robust file loading with multiple format support."""
    
    @staticmethod
    def load(file_path):
        """Load file with automatic format detection and error handling."""
        path = Path(file_path)
        ext = path.suffix.lower()
        
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        if path.stat().st_size == 0:
            raise ValueError("File is empty (0 bytes)")
        
        logger.log(f"Loading file: {path.name} (Format: {ext})")
        
        try:
            if ext == '.csv':
                return FileLoader._load_csv(file_path)
            elif ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
                return FileLoader._load_excel(file_path)
            elif ext == '.json':
                return FileLoader._load_json(file_path)
            elif ext in ['.parquet', '.pq']:
                return FileLoader._load_parquet(file_path)
            else:
                # Try CSV as fallback
                logger.log(f"Unknown extension {ext}, trying CSV parser")
                return FileLoader._load_csv(file_path)
        except Exception as e:
            logger.log(f"Load error: {str(e)}", "ERROR")
            raise
    
    @staticmethod
    def _load_csv(file_path):
        """Load CSV with encoding detection and delimiter inference."""
        # Try different encodings
        for encoding in Config.ENCODINGS:
            try:
                # First, try to detect delimiter
                with open(file_path, 'r', encoding=encoding) as f:
                    sample = f.read(8192)
                
                # Count potential delimiters
                delimiters = {',': 0, ';': 0, '\t': 0, '|': 0}
                for delim in delimiters:
                    delimiters[delim] = sample.count(delim)
                
                best_delim = max(delimiters, key=delimiters.get)
                if delimiters[best_delim] == 0:
                    best_delim = ','  # Default
                
                logger.log(f"Detected delimiter: '{best_delim}', encoding: {encoding}")
                
                # Load with detected settings
                df = pd.read_csv(
                    file_path,
                    encoding=encoding,
                    delimiter=best_delim,
                    on_bad_lines='skip',
                    low_memory=False,
                    dtype=str  # Load as string first for safety
                )
                
                if not df.empty:
                    logger.log(f"CSV loaded: {len(df)} rows, {len(df.columns)} columns")
                    return df
                    
            except UnicodeDecodeError:
                continue
            except Exception as e:
                logger.log(f"CSV parse attempt failed ({encoding}): {str(e)}", "WARN")
                continue
        
        raise ValueError("Could not parse CSV with any supported encoding")
    
    @staticmethod
    def _load_excel(file_path):
        """Load Excel with sheet detection."""
        try:
            xl = pd.ExcelFile(file_path)
            sheets = xl.sheet_names
            logger.log(f"Excel file has {len(sheets)} sheet(s): {sheets}")
            
            # Find the best sheet (first non-empty one with data)
            best_df = None
            best_rows = 0
            
            for sheet in sheets:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet, dtype=str)
                    if len(df) > best_rows:
                        best_df = df
                        best_rows = len(df)
                        logger.log(f"Sheet '{sheet}': {len(df)} rows")
                except Exception as e:
                    logger.log(f"Couldn't read sheet '{sheet}': {str(e)}", "WARN")
                    continue
            
            if best_df is not None and not best_df.empty:
                logger.log(f"Using sheet with {best_rows} rows")
                return best_df
            
            raise ValueError("No valid data found in any Excel sheet")
            
        except Exception as e:
            # Fallback: try openpyxl directly
            try:
                df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
                return df
            except:
                pass
            raise
    
    @staticmethod
    def _load_json(file_path):
        """Load JSON with multiple format support."""
        try:
            # Try records orientation first
            df = pd.read_json(file_path, orient='records', dtype=str)
            if not df.empty:
                return df
        except:
            pass
        
        try:
            # Try default orientation
            df = pd.read_json(file_path, dtype=str)
            return df
        except:
            pass
        
        try:
            # Try line-delimited JSON
            df = pd.read_json(file_path, lines=True, dtype=str)
            return df
        except Exception as e:
            raise ValueError(f"Could not parse JSON: {str(e)}")
    
    @staticmethod
    def _load_parquet(file_path):
        """Load Parquet file."""
        try:
            df = pd.read_parquet(file_path)
            # Convert to string for consistent processing
            return df.astype(str)
        except Exception as e:
            raise ValueError(f"Could not read Parquet: {str(e)}")


# ==============================
# DATA PROCESSOR
# ==============================
class DataProcessor:
    """Comprehensive data processing and cleaning."""
    
    def __init__(self, df):
        self.df = df.copy()
        self.original_columns = list(df.columns)
        self.processing_notes = []
    
    def process(self):
        """Run full processing pipeline."""
        self._clean_column_names()
        self._convert_types()
        self._map_columns()
        self._infer_missing_columns()
        self._clean_data()
        self._engineer_features()
        return self.df, self.processing_notes
    
    def _clean_column_names(self):
        """Standardize column names."""
        # Remove leading/trailing whitespace and special characters
        new_cols = {}
        for col in self.df.columns:
            clean = str(col).strip()
            # Remove common problematic characters
            clean = clean.replace('\n', ' ').replace('\r', '')
            if clean != col:
                new_cols[col] = clean
        
        if new_cols:
            self.df.rename(columns=new_cols, inplace=True)
            logger.log(f"Cleaned {len(new_cols)} column names")
    
    def _convert_types(self):
        """Convert columns to appropriate types."""
        for col in self.df.columns:
            # Try numeric conversion
            try:
                # Replace common non-numeric strings
                temp = self.df[col].astype(str).str.strip()
                temp = temp.replace(['', ' ', 'NA', 'N/A', 'null', 'NULL', 'None', 'nan', 'NaN', '-'], np.nan)
                temp = temp.str.replace('$', '', regex=False)
                temp = temp.str.replace(',', '', regex=False)
                temp = temp.str.replace('%', '', regex=False)
                
                numeric_temp = pd.to_numeric(temp, errors='coerce')
                
                # If more than 50% are valid numbers, convert
                valid_ratio = numeric_temp.notna().sum() / len(numeric_temp)
                if valid_ratio > 0.5:
                    self.df[col] = numeric_temp
                    logger.log(f"Converted '{col}' to numeric ({valid_ratio*100:.0f}% valid)")
            except:
                pass
    
    def _map_columns(self):
        """Map column names to standard names."""
        mapped = {}
        for col in self.df.columns:
            col_lower = col.lower().strip().replace(' ', '_')
            if col_lower in Config.COLUMN_MAPPINGS:
                target = Config.COLUMN_MAPPINGS[col_lower]
                if target not in self.df.columns:
                    mapped[col] = target
        
        if mapped:
            self.df.rename(columns=mapped, inplace=True)
            logger.log(f"Mapped columns: {mapped}")
            self.processing_notes.append(f"Mapped columns: {list(mapped.keys())}")
    
    def _infer_missing_columns(self):
        """Intelligently infer required columns if not found."""
        required = ['tenure', 'MonthlyCharges', 'TotalCharges', 'Churn']
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns.tolist()
        
        for col in required:
            if col not in self.df.columns:
                inferred = self._infer_column(col, numeric_cols)
                if inferred:
                    self.processing_notes.append(f"Inferred '{col}' from '{inferred}'")
    
    def _infer_column(self, target, numeric_cols):
        """Infer a specific column."""
        
        if target == 'Churn':
            return self._infer_churn()
        elif target == 'tenure':
            return self._infer_tenure(numeric_cols)
        elif target == 'MonthlyCharges':
            return self._infer_monthly_charges(numeric_cols)
        elif target == 'TotalCharges':
            return self._infer_total_charges(numeric_cols)
        return None
    
    def _infer_churn(self):
        """Infer churn column."""
        # Look for binary columns
        for col in self.df.columns:
            col_lower = col.lower()
            
            # Check keywords
            churn_keywords = ['churn', 'left', 'exit', 'attrition', 'cancel', 'status', 'active']
            if any(kw in col_lower for kw in churn_keywords):
                unique = self.df[col].dropna().unique()
                if len(unique) <= 5:  # Likely categorical
                    self.df['Churn'] = self._convert_to_binary(self.df[col])
                    logger.log(f"Inferred Churn from '{col}'")
                    return col
        
        # Look for any binary column
        for col in self.df.columns:
            if self.df[col].nunique() == 2:
                self.df['Churn'] = self._convert_to_binary(self.df[col])
                logger.log(f"Inferred Churn from binary column '{col}'")
                return col
        
        # Default: no churn data
        self.df['Churn'] = 0
        logger.log("No churn column found, defaulting to 0")
        self.processing_notes.append("No churn data found - using placeholder values")
        return None
    
    def _convert_to_binary(self, series):
        """Convert a series to binary (0/1)."""
        series = series.astype(str).str.lower().str.strip()
        
        positive = ['yes', 'y', 'true', '1', 'churned', 'left', 'exited', 'cancelled', 'inactive', 'lost']
        negative = ['no', 'n', 'false', '0', 'active', 'stayed', 'retained', 'current']
        
        def convert(val):
            if pd.isna(val) or val in ['nan', 'none', '']:
                return 0
            if val in positive:
                return 1
            if val in negative:
                return 0
            # For 'active' column, invert logic
            return 0
        
        return series.apply(convert)
    
    def _infer_tenure(self, numeric_cols):
        """Infer tenure column."""
        keywords = ['month', 'tenure', 'time', 'period', 'duration', 'age', 'length', 'days', 'weeks', 'years']
        
        for col in numeric_cols:
            if any(kw in col.lower() for kw in keywords):
                self.df['tenure'] = self.df[col].copy()
                
                # Convert if appears to be days/years
                if 'day' in col.lower():
                    self.df['tenure'] = self.df['tenure'] / 30
                elif 'year' in col.lower():
                    self.df['tenure'] = self.df['tenure'] * 12
                elif 'week' in col.lower():
                    self.df['tenure'] = self.df['tenure'] / 4
                
                logger.log(f"Inferred tenure from '{col}'")
                return col
        
        # Use first reasonable numeric column
        for col in numeric_cols:
            if self.df[col].min() >= 0 and self.df[col].max() <= 1000:
                self.df['tenure'] = self.df[col].copy()
                logger.log(f"Inferred tenure from '{col}' (first suitable numeric)")
                return col
        
        # Default
        self.df['tenure'] = 12
        logger.log("No tenure column found, defaulting to 12")
        return None
    
    def _infer_monthly_charges(self, numeric_cols):
        """Infer monthly charges column."""
        keywords = ['monthly', 'charge', 'fee', 'payment', 'price', 'amount', 'bill', 'cost', 'rate']
        exclude = ['total', 'cumulative', 'lifetime', 'sum']
        
        for col in numeric_cols:
            col_lower = col.lower()
            if any(kw in col_lower for kw in keywords) and not any(ex in col_lower for ex in exclude):
                self.df['MonthlyCharges'] = self.df[col].copy()
                logger.log(f"Inferred MonthlyCharges from '{col}'")
                return col
        
        # Use a column with typical monthly charge range
        for col in numeric_cols:
            if col not in ['tenure', 'Churn']:
                median = self.df[col].median()
                if 10 <= median <= 500:
                    self.df['MonthlyCharges'] = self.df[col].copy()
                    logger.log(f"Inferred MonthlyCharges from '{col}' (value range match)")
                    return col
        
        # Default
        self.df['MonthlyCharges'] = 50
        logger.log("No monthly charges column found, defaulting to 50")
        return None
    
    def _infer_total_charges(self, numeric_cols):
        """Infer total charges column."""
        keywords = ['total', 'lifetime', 'cumulative', 'sum', 'overall', 'aggregate']
        
        for col in numeric_cols:
            if any(kw in col.lower() for kw in keywords):
                self.df['TotalCharges'] = self.df[col].copy()
                logger.log(f"Inferred TotalCharges from '{col}'")
                return col
        
        # Calculate from tenure and monthly
        if 'tenure' in self.df.columns and 'MonthlyCharges' in self.df.columns:
            self.df['TotalCharges'] = self.df['MonthlyCharges'] * self.df['tenure']
            logger.log("Calculated TotalCharges from tenure × MonthlyCharges")
            self.processing_notes.append("Calculated TotalCharges from tenure × MonthlyCharges")
            return 'calculated'
        
        # Default
        self.df['TotalCharges'] = 600
        logger.log("No total charges column found, defaulting to 600")
        return None
    
    def _clean_data(self):
        """Comprehensive data cleaning."""
        key_cols = ['tenure', 'MonthlyCharges', 'TotalCharges', 'Churn']
        
        for col in key_cols:
            if col in self.df.columns:
                # Convert to numeric
                self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
                
                # Fill missing with median (or 0 for Churn)
                if col == 'Churn':
                    self.df[col] = self.df[col].fillna(0).astype(int)
                    # Ensure binary
                    self.df[col] = (self.df[col] > 0).astype(int)
                else:
                    median = self.df[col].median()
                    if pd.isna(median):
                        median = 1 if col == 'tenure' else 50
                    self.df[col] = self.df[col].fillna(median)
                    
                    # Ensure positive
                    self.df[col] = self.df[col].abs()
        
        # Handle zero tenure
        if 'tenure' in self.df.columns:
            self.df['tenure'] = self.df['tenure'].replace(0, 1)
        
        # Remove completely empty rows
        self.df.dropna(how='all', inplace=True)
        
        # Remove duplicate rows
        initial_len = len(self.df)
        self.df.drop_duplicates(inplace=True)
        if len(self.df) < initial_len:
            logger.log(f"Removed {initial_len - len(self.df)} duplicate rows")
        
        # Reset index
        self.df.reset_index(drop=True, inplace=True)
        
        logger.log("Data cleaning complete")
    
    def _engineer_features(self):
        """Create analytical features."""
        # Average monthly spend
        if 'TotalCharges' in self.df.columns and 'tenure' in self.df.columns:
            self.df['AvgMonthlySpend'] = self.df['TotalCharges'] / (self.df['tenure'] + 1)
            self.df['AvgMonthlySpend'] = self.df['AvgMonthlySpend'].clip(lower=0)
        else:
            self.df['AvgMonthlySpend'] = self.df.get('MonthlyCharges', 50)
        
        # Tenure groups
        if 'tenure' in self.df.columns:
            try:
                self.df['TenureGroup'] = pd.cut(
                    self.df['tenure'],
                    bins=[0, 12, 24, 48, 72, float('inf')],
                    labels=['0-1yr', '1-2yr', '2-4yr', '4-6yr', '6+yr'],
                    include_lowest=True
                )
            except:
                self.df['TenureGroup'] = 'Unknown'
        
        # Charge ratio
        if 'MonthlyCharges' in self.df.columns:
            avg_charge = self.df['MonthlyCharges'].mean()
            if avg_charge > 0:
                self.df['ChargeRatio'] = self.df['MonthlyCharges'] / avg_charge
            else:
                self.df['ChargeRatio'] = 1
        else:
            self.df['ChargeRatio'] = 1
        
        # Spending variance
        if all(c in self.df.columns for c in ['MonthlyCharges', 'tenure', 'TotalCharges']):
            self.df['ExpectedTotal'] = self.df['MonthlyCharges'] * self.df['tenure']
            self.df['SpendingVariance'] = self.df['TotalCharges'] - self.df['ExpectedTotal']
        
        # Value score
        if 'TotalCharges' in self.df.columns:
            max_total = self.df['TotalCharges'].max()
            if max_total > 0:
                self.df['ValueScore'] = (self.df['TotalCharges'] / max_total) * 100
            else:
                self.df['ValueScore'] = 50
        
        logger.log("Feature engineering complete")


# ==============================
# MODEL BUILDER
# ==============================
class ModelBuilder:
    """Build predictive models with fallback logic."""
    
    def __init__(self, df):
        self.df = df.copy()
        self.accuracy = None
        self.model_note = ""
    
    def build(self):
        """Build churn and CLV models."""
        self._build_churn_model()
        self._build_clv_model()
        self._create_segments()
        return self.df, self.accuracy, self.model_note
    
    def _build_churn_model(self):
        """Build churn prediction model."""
        features = ['tenure', 'MonthlyCharges', 'TotalCharges', 'AvgMonthlySpend', 'ChargeRatio']
        features = [f for f in features if f in self.df.columns]
        
        if len(features) < 2:
            self.df['Churn_Prob'] = 0.5
            self.model_note = "Insufficient features for ML model"
            logger.log("Not enough features for churn model")
            return
        
        X = self.df[features].copy()
        y = self.df['Churn'].copy()
        
        # Check if sklearn is available
        if not SKLEARN_AVAILABLE:
            self._heuristic_churn_scoring(X)
            self.model_note = "Using heuristic scoring (sklearn not available)"
            return
        
        # Check data validity
        unique_classes = y.nunique()
        churn_sum = y.sum()
        
        if unique_classes < 2 or len(self.df) < 10:
            self._heuristic_churn_scoring(X)
            self.model_note = "Using heuristic scoring (insufficient class variation)"
            return
        
        try:
            # Scale features
            scaler = StandardScaler()
            X_scaled = pd.DataFrame(scaler.fit_transform(X), columns=features)
            
            # Determine test size
            test_size = min(0.25, max(0.1, 10 / len(self.df)))
            
            # Stratify if possible
            stratify = y if churn_sum >= 2 and (len(y) - churn_sum) >= 2 else None
            
            X_train, X_test, y_train, y_test = train_test_split(
                X_scaled, y, test_size=test_size, stratify=stratify, random_state=42
            )
            
            # Train models
            models = [
                ('RF', RandomForestClassifier(n_estimators=100, max_depth=8, min_samples_split=5, random_state=42)),
                ('GB', GradientBoostingClassifier(n_estimators=80, max_depth=4, random_state=42))
            ]
            
            best_model = None
            best_accuracy = 0
            
            for name, model in models:
                try:
                    model.fit(X_train, y_train)
                    acc = accuracy_score(y_test, model.predict(X_test))
                    logger.log(f"Model {name}: accuracy = {acc:.3f}")
                    if acc > best_accuracy:
                        best_accuracy = acc
                        best_model = model
                except Exception as e:
                    logger.log(f"Model {name} failed: {str(e)}", "WARN")
                    continue
            
            if best_model is not None:
                # Predict probabilities
                self.df['Churn_Prob'] = best_model.predict_proba(scaler.transform(X))[:, 1]
                self.accuracy = best_accuracy
                self.model_note = f"ML model trained (accuracy: {best_accuracy:.1%})"
                logger.log(f"Best model accuracy: {best_accuracy:.3f}")
            else:
                self._heuristic_churn_scoring(X)
                self.model_note = "Using heuristic scoring (model training failed)"
                
        except Exception as e:
            logger.log(f"Model building error: {str(e)}", "ERROR")
            self._heuristic_churn_scoring(X)
            self.model_note = "Using heuristic scoring (model error)"
    
    def _heuristic_churn_scoring(self, X):
        """Calculate churn probability using heuristics."""
        # Lower tenure = higher risk
        if 'tenure' in X.columns:
            tenure_score = 1 - (X['tenure'] / X['tenure'].max()).clip(0, 1)
        else:
            tenure_score = 0.5
        
        # Higher charges relative to average = slightly higher risk
        if 'ChargeRatio' in X.columns:
            charge_score = (X['ChargeRatio'] - 1).clip(-1, 1) * 0.2 + 0.5
        else:
            charge_score = 0.5
        
        # Combine scores
        self.df['Churn_Prob'] = (tenure_score * 0.6 + charge_score * 0.4).clip(0.05, 0.95)
        logger.log("Using heuristic churn scoring")
    
    def _build_clv_model(self):
        """Build Customer Lifetime Value model."""
        features = ['tenure', 'MonthlyCharges', 'AvgMonthlySpend']
        features = [f for f in features if f in self.df.columns]
        
        if len(features) < 2 or not SKLEARN_AVAILABLE:
            # Simple heuristic
            self.df['Predicted_CLV'] = self.df['MonthlyCharges'] * 24
            logger.log("Using heuristic CLV calculation")
            return
        
        try:
            X = self.df[features]
            y = self.df['TotalCharges']
            
            model = Ridge(alpha=1.0)
            model.fit(X, y)
            
            self.df['Predicted_CLV'] = model.predict(X)
            self.df['Predicted_CLV'] = self.df['Predicted_CLV'].clip(lower=0)
            
            # Adjust for churn risk
            self.df['Predicted_CLV'] = self.df['Predicted_CLV'] * (1 - self.df['Churn_Prob'] * 0.5)
            
            logger.log("CLV model built successfully")
        except Exception as e:
            logger.log(f"CLV model error: {str(e)}", "WARN")
            self.df['Predicted_CLV'] = self.df['MonthlyCharges'] * 24
    
    def _create_segments(self):
        """Create customer segments."""
        # Use percentiles for robust thresholds
        churn_high = self.df['Churn_Prob'].quantile(0.7)
        clv_median = self.df['Predicted_CLV'].median()
        
        def segment(row):
            high_risk = row['Churn_Prob'] > churn_high
            high_value = row['Predicted_CLV'] > clv_median
            
            if high_risk and high_value:
                return "High Risk - High Value"
            elif high_risk:
                return "High Risk - Low Value"
            elif high_value:
                return "Low Risk - High Value"
            else:
                return "Low Risk - Low Value"
        
        self.df['Segment'] = self.df.apply(segment, axis=1)
        
        # Priority score
        self.df['Priority'] = (self.df['Churn_Prob'] * self.df['Predicted_CLV']).rank(pct=True) * 100
        
        logger.log("Customer segmentation complete")


# ==============================
# REPORT GENERATOR
# ==============================
class ReportGenerator:
    """Generate comprehensive Excel report with dashboard."""
    
    def __init__(self, df, accuracy, model_note, processing_notes):
        self.df = df
        self.accuracy = accuracy
        self.model_note = model_note
        self.processing_notes = processing_notes
    
    def generate(self, save_path):
        """Generate the full report."""
        # Create distributions
        distributions = self._create_distributions()
        
        # Write Excel
        self._write_excel(save_path, distributions)
        
        logger.log(f"Report saved to: {save_path}")
    
    def _create_distributions(self):
        """Create distribution dataframes."""
        distributions = {}
        
        # Segment distribution
        seg = self.df['Segment'].value_counts().reset_index()
        seg.columns = ['Segment', 'Count']
        seg['Percentage'] = (seg['Count'] / seg['Count'].sum() * 100).round(1)
        distributions['Segments'] = seg
        
        # Churn probability distribution
        try:
            churn_bins = pd.cut(self.df['Churn_Prob'], bins=[0, 0.2, 0.4, 0.6, 0.8, 1.0], 
                               labels=['0-20%', '20-40%', '40-60%', '60-80%', '80-100%'])
            churn_dist = churn_bins.value_counts().sort_index().reset_index()
            churn_dist.columns = ['Risk Level', 'Count']
        except:
            churn_dist = pd.DataFrame({'Risk Level': ['All'], 'Count': [len(self.df)]})
        distributions['ChurnDist'] = churn_dist
        
        # CLV distribution
        try:
            clv_labels = ['Very Low', 'Low', 'Medium', 'High', 'Very High']
            clv_bins = pd.qcut(self.df['Predicted_CLV'], q=5, labels=clv_labels, duplicates='drop')
            clv_dist = clv_bins.value_counts().sort_index().reset_index()
            clv_dist.columns = ['CLV Tier', 'Count']
        except:
            clv_dist = pd.DataFrame({'CLV Tier': ['All'], 'Count': [len(self.df)]})
        distributions['CLVDist'] = clv_dist
        
        # Spend distribution
        try:
            spend_labels = ['Very Low', 'Low', 'Medium', 'High', 'Very High']
            spend_bins = pd.qcut(self.df['AvgMonthlySpend'], q=5, labels=spend_labels, duplicates='drop')
            spend_dist = spend_bins.value_counts().sort_index().reset_index()
            spend_dist.columns = ['Spend Tier', 'Count']
        except:
            spend_dist = pd.DataFrame({'Spend Tier': ['All'], 'Count': [len(self.df)]})
        distributions['SpendDist'] = spend_dist
        
        # Tenure distribution
        if 'TenureGroup' in self.df.columns:
            tenure_dist = self.df['TenureGroup'].value_counts().reset_index()
            tenure_dist.columns = ['Tenure Group', 'Count']
        else:
            tenure_dist = pd.DataFrame({'Tenure Group': ['All'], 'Count': [len(self.df)]})
        distributions['TenureDist'] = tenure_dist
        
        return distributions
    
    def _write_excel(self, save_path, distributions):
        """Write the Excel report."""
        with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Create formats
            formats = self._create_formats(workbook)
            
            # Write data sheet
            self._write_data_sheet(writer, formats)
            
            # Write distribution sheets
            self._write_distribution_sheets(writer, distributions, formats)
            
            # Write summary sheet
            self._write_summary_sheet(writer, formats)
            
            # Write dashboard
            self._write_dashboard(writer, distributions, formats)
    
    def _create_formats(self, workbook):
        """Create Excel formats."""
        return {
            'header': workbook.add_format({
                'bold': True, 'font_color': 'white', 'bg_color': '#006400',
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True
            }),
            'cell': workbook.add_format({
                'align': 'center', 'valign': 'vcenter'
            }),
            'number': workbook.add_format({
                'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0.00'
            }),
            'percent': workbook.add_format({
                'align': 'center', 'valign': 'vcenter', 'num_format': '0.0%'
            }),
            'currency': workbook.add_format({
                'align': 'center', 'valign': 'vcenter', 'num_format': '$#,##0.00'
            }),
            'title': workbook.add_format({
                'bold': True, 'font_size': 20, 'font_color': '#006400',
                'align': 'left', 'valign': 'vcenter'
            }),
            'kpi_blue': workbook.add_format({
                'bold': True, 'font_size': 14, 'font_color': 'white', 'bg_color': '#1E90FF',
                'align': 'center', 'valign': 'vcenter', 'border': 2
            }),
            'kpi_green': workbook.add_format({
                'bold': True, 'font_size': 14, 'font_color': 'white', 'bg_color': '#32CD32',
                'align': 'center', 'valign': 'vcenter', 'border': 2
            }),
            'kpi_red': workbook.add_format({
                'bold': True, 'font_size': 14, 'font_color': 'white', 'bg_color': '#FF4500',
                'align': 'center', 'valign': 'vcenter', 'border': 2
            }),
            'kpi_purple': workbook.add_format({
                'bold': True, 'font_size': 14, 'font_color': 'white', 'bg_color': '#9370DB',
                'align': 'center', 'valign': 'vcenter', 'border': 2
            }),
            'kpi_value': workbook.add_format({
                'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'border': 2
            }),
            'kpi_value_currency': workbook.add_format({
                'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter',
                'border': 2, 'num_format': '$#,##0.00'
            }),
            'kpi_value_percent': workbook.add_format({
                'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter',
                'border': 2, 'num_format': '0.0%'
            })
        }
    
    def _write_data_sheet(self, writer, formats):
        """Write the main data sheet."""
        # Select columns to export
        exclude_cols = ['ExpectedTotal', 'ChargeRatio']
        export_cols = [c for c in self.df.columns if c not in exclude_cols]
        df_export = self.df[export_cols].copy()
        
        # Convert categorical to string for Excel
        for col in df_export.columns:
            if df_export[col].dtype.name == 'category':
                df_export[col] = df_export[col].astype(str)
        
        df_export.to_excel(writer, sheet_name='Data', index=False)
        ws = writer.sheets['Data']
        
        # Format headers and columns
        for i, col in enumerate(df_export.columns):
            ws.write(0, i, col, formats['header'])
            max_len = min(max(df_export[col].astype(str).str.len().max(), len(col)) + 2, 30)
            ws.set_column(i, i, max_len, formats['cell'])
        
        ws.autofilter(0, 0, len(df_export), len(df_export.columns) - 1)
        ws.freeze_panes(1, 0)
    
    def _write_distribution_sheets(self, writer, distributions, formats):
        """Write distribution sheets."""
        for name, data in distributions.items():
            if data is not None and not data.empty:
                # Convert any interval/categorical to string
                for col in data.columns:
                    if data[col].dtype.name == 'category' or 'interval' in str(data[col].dtype).lower():
                        data[col] = data[col].astype(str)
                
                data.to_excel(writer, sheet_name=name, index=False)
                ws = writer.sheets[name]
                
                for i, col in enumerate(data.columns):
                    ws.write(0, i, col, formats['header'])
                    ws.set_column(i, i, 20, formats['cell'])
                
                ws.autofilter(0, 0, len(data), len(data.columns) - 1)
    
    def _write_summary_sheet(self, writer, formats):
        """Write summary sheet with KPIs."""
        # Calculate KPIs
        kpis = {
            'Total Customers': len(self.df),
            'Average CLV': self.df['Predicted_CLV'].mean(),
            'Median CLV': self.df['Predicted_CLV'].median(),
            'Total CLV': self.df['Predicted_CLV'].sum(),
            'Churn Rate': self.df['Churn'].mean(),
            'Average Tenure (months)': self.df['tenure'].mean(),
            'Avg Monthly Charges': self.df['MonthlyCharges'].mean(),
            'High Risk Customers': (self.df['Churn_Prob'] > 0.7).sum(),
            'High Value Customers': (self.df['Predicted_CLV'] > self.df['Predicted_CLV'].median()).sum(),
            'Model Accuracy': self.accuracy if self.accuracy else 'N/A',
            'Model Note': self.model_note
        }
        
        # Create summary dataframe
        summary_df = pd.DataFrame([
            ['Metric', 'Value'],
            *[[k, v] for k, v in kpis.items()]
        ])
        
        ws = writer.book.add_worksheet('Summary')
        
        # Write with formatting
        for row_num, row_data in enumerate(summary_df.values):
            for col_num, value in enumerate(row_data):
                if row_num == 0:
                    ws.write(row_num, col_num, value, formats['header'])
                elif isinstance(value, float):
                    if 'CLV' in str(summary_df.values[row_num][0]) or 'Charges' in str(summary_df.values[row_num][0]):
                        ws.write(row_num, col_num, value, formats['currency'])
                    elif 'Rate' in str(summary_df.values[row_num][0]):
                        ws.write(row_num, col_num, value, formats['percent'])
                    else:
                        ws.write(row_num, col_num, value, formats['number'])
                else:
                    ws.write(row_num, col_num, value, formats['cell'])
        
        ws.set_column('A:A', 30)
        ws.set_column('B:B', 25)
        
        # Add processing notes
        if self.processing_notes:
            ws.write(len(kpis) + 3, 0, 'Processing Notes:', formats['header'])
            for i, note in enumerate(self.processing_notes):
                ws.write(len(kpis) + 4 + i, 0, f"• {note}", formats['cell'])
    
    def _write_dashboard(self, writer, distributions, formats):
        """Write dashboard with charts."""
        workbook = writer.book
        dashboard = workbook.add_worksheet('Dashboard')
        
        # Column sizing
        dashboard.set_column('A:A', 35)
        dashboard.set_column('B:B', 25)
        dashboard.set_column('C:C', 5)
        for col in range(3, 25):
            dashboard.set_column(col, col, 12)
        
        for row in range(3, 10):
            dashboard.set_row(row, 45)
        
        # Title
        dashboard.merge_range('A1:H1', f'📊 Customer Analytics Dashboard', formats['title'])
        dashboard.set_row(0, 50)
        
        # Subtitle with timestamp
        subtitle_format = workbook.add_format({
            'italic': True, 'font_size': 10, 'font_color': '#666666'
        })
        dashboard.write('A2', f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", subtitle_format)
        
        # KPI Cards
        total_customers = len(self.df)
        avg_clv = self.df['Predicted_CLV'].mean()
        churn_rate = self.df['Churn'].mean()
        high_risk = (self.df['Churn_Prob'] > 0.7).sum()
        avg_tenure = self.df['tenure'].mean()
        avg_monthly = self.df['MonthlyCharges'].mean()
        
        # Row 1 KPIs
        dashboard.write('A4', 'Total Customers', formats['kpi_blue'])
        dashboard.write('B4', total_customers, formats['kpi_value'])
        
        dashboard.write('A5', 'Average CLV', formats['kpi_green'])
        dashboard.write('B5', avg_clv, formats['kpi_value_currency'])
        
        dashboard.write('A6', 'Churn Rate', formats['kpi_red'])
        dashboard.write('B6', churn_rate, formats['kpi_value_percent'])
        
        dashboard.write('A7', 'High Risk Customers', formats['kpi_purple'])
        dashboard.write('B7', high_risk, formats['kpi_value'])
        
        dashboard.write('A8', 'Avg Tenure (months)', formats['kpi_blue'])
        dashboard.write('B8', round(avg_tenure, 1), formats['kpi_value'])
        
        dashboard.write('A9', 'Avg Monthly Charges', formats['kpi_green'])
        dashboard.write('B9', avg_monthly, formats['kpi_value_currency'])
        
        # Charts
        chart_row = 3
        chart_col = 3
        
        # Segment Distribution Chart
        seg_data = distributions.get('Segments')
        if seg_data is not None and len(seg_data) > 0:
            chart = workbook.add_chart({'type': 'column'})
            chart.add_series({
                'categories': ['Segments', 1, 0, len(seg_data), 0],
                'values': ['Segments', 1, 1, len(seg_data), 1],
                'fill': {'color': '#2E8B57'},
                'data_labels': {'value': True, 'font': {'bold': True, 'size': 9}}
            })
            chart.set_title({'name': 'Customer Segments', 'name_font': {'size': 12, 'bold': True}})
            chart.set_style(10)
            chart.set_legend({'none': True})
            chart.set_x_axis({'name_font': {'size': 9}})
            dashboard.insert_chart(chart_row, chart_col, chart, {'x_scale': 1.3, 'y_scale': 1.2})
        
        # Churn Risk Distribution
        churn_data = distributions.get('ChurnDist')
        if churn_data is not None and len(churn_data) > 0:
            chart = workbook.add_chart({'type': 'column'})
            chart.add_series({
                'categories': ['ChurnDist', 1, 0, len(churn_data), 0],
                'values': ['ChurnDist', 1, 1, len(churn_data), 1],
                'fill': {'color': '#FF6347'},
                'data_labels': {'value': True, 'font': {'bold': True, 'size': 9}}
            })
            chart.set_title({'name': 'Churn Risk Distribution', 'name_font': {'size': 12, 'bold': True}})
            chart.set_style(10)
            chart.set_legend({'none': True})
            dashboard.insert_chart(chart_row + 17, chart_col, chart, {'x_scale': 1.3, 'y_scale': 1.2})
        
        # CLV Distribution
        clv_data = distributions.get('CLVDist')
        if clv_data is not None and len(clv_data) > 0:
            chart = workbook.add_chart({'type': 'column'})
            chart.add_series({
                'categories': ['CLVDist', 1, 0, len(clv_data), 0],
                'values': ['CLVDist', 1, 1, len(clv_data), 1],
                'fill': {'color': '#1E90FF'},
                'data_labels': {'value': True, 'font': {'bold': True, 'size': 9}}
            })
            chart.set_title({'name': 'Customer Lifetime Value Tiers', 'name_font': {'size': 12, 'bold': True}})
            chart.set_style(10)
            chart.set_legend({'none': True})
            dashboard.insert_chart(chart_row, chart_col + 10, chart, {'x_scale': 1.3, 'y_scale': 1.2})
        
        # Spend Distribution
        spend_data = distributions.get('SpendDist')
        if spend_data is not None and len(spend_data) > 0:
            chart = workbook.add_chart({'type': 'column'})
            chart.add_series({
                'categories': ['SpendDist', 1, 0, len(spend_data), 0],
                'values': ['SpendDist', 1, 1, len(spend_data), 1],
                'fill': {'color': '#FFD700'},
                'data_labels': {'value': True, 'font': {'bold': True, 'size': 9}}
            })
            chart.set_title({'name': 'Monthly Spending Pattern', 'name_font': {'size': 12, 'bold': True}})
            chart.set_style(10)
            chart.set_legend({'none': True})
            dashboard.insert_chart(chart_row + 17, chart_col + 10, chart, {'x_scale': 1.3, 'y_scale': 1.2})
        
        # Tenure Distribution (Pie)
        tenure_data = distributions.get('TenureDist')
        if tenure_data is not None and len(tenure_data) > 1:
            chart = workbook.add_chart({'type': 'pie'})
            chart.add_series({
                'categories': ['TenureDist', 1, 0, len(tenure_data), 0],
                'values': ['TenureDist', 1, 1, len(tenure_data), 1],
                'data_labels': {'percentage': True, 'category': True, 'font': {'size': 9}}
            })
            chart.set_title({'name': 'Tenure Distribution', 'name_font': {'size': 12, 'bold': True}})
            chart.set_style(10)
            dashboard.insert_chart(chart_row, chart_col + 20, chart, {'x_scale': 1.2, 'y_scale': 1.2})


# ==============================
# PROGRESS DIALOG
# ==============================
class ProgressDialog:
    """Simple progress dialog for feedback."""
    
    def __init__(self, parent, title="Processing"):
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("400x150")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()
        
        # Center the window
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() - 400) // 2
        y = (self.window.winfo_screenheight() - 150) // 2
        self.window.geometry(f"+{x}+{y}")
        
        # Status label
        self.status_var = tk.StringVar(value="Initializing...")
        self.status_label = ttk.Label(self.window, textvariable=self.status_var, font=('Segoe UI', 10))
        self.status_label.pack(pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(self.window, mode='indeterminate', length=350)
        self.progress.pack(pady=10)
        self.progress.start(10)
        
        # Detail label
        self.detail_var = tk.StringVar(value="")
        self.detail_label = ttk.Label(self.window, textvariable=self.detail_var, font=('Segoe UI', 8), foreground='gray')
        self.detail_label.pack(pady=5)
    
    def update(self, status, detail=""):
        self.status_var.set(status)
        self.detail_var.set(detail)
        self.window.update()
    
    def close(self):
        self.progress.stop()
        self.window.destroy()


# ==============================
# MAIN APPLICATION
# ==============================
class CustomerAnalyticsApp:
    """Main application class."""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        
        # Set DPI awareness for Windows
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
    
    def run(self):
        """Run the application."""
        try:
            # Welcome message
            result = messagebox.askyesno(
                f"{Config.APP_NAME} v{Config.VERSION}",
                "Welcome to Customer Analytics Tool!\n\n"
                "This tool will:\n"
                "• Analyze your customer data\n"
                "• Predict churn probability\n"
                "• Calculate customer lifetime value\n"
                "• Create actionable segments\n"
                "• Generate a visual dashboard\n\n"
                "Supported formats: CSV, Excel, JSON, Parquet\n\n"
                "Would you like to continue?",
                icon='info'
            )
            
            if not result:
                return
            
            # File selection
            file_path = self._select_input_file()
            if not file_path:
                return
            
            # Process data
            df, accuracy, model_note, processing_notes = self._process_data(file_path)
            
            if df is None:
                return
            
            # Save location
            save_path = self._select_save_location(file_path)
            if not save_path:
                return
            
            # Generate report
            self._generate_report(df, save_path, accuracy, model_note, processing_notes)
            
            # Success message
            self._show_success(df, save_path, accuracy, model_note)
            
        except Exception as e:
            self._show_error(e)
        finally:
            self.root.destroy()
    
    def _select_input_file(self):
        """Select input file."""
        filetypes = [
            ("All Supported", "*.csv;*.xlsx;*.xls;*.xlsm;*.xlsb;*.json;*.parquet;*.pq"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx;*.xls;*.xlsm;*.xlsb"),
            ("JSON files", "*.json"),
            ("Parquet files", "*.parquet;*.pq"),
            ("All files", "*.*")
        ]
        
        file_path = filedialog.askopenfilename(
            title="Select Customer Data File",
            filetypes=filetypes
        )
        
        if not file_path:
            messagebox.showinfo("Cancelled", "No file selected. Exiting.")
            return None
        
        return file_path
    
    def _process_data(self, file_path):
        """Process the data file."""
        progress = ProgressDialog(self.root, "Processing Data")
        
        try:
            # Load file
            progress.update("Loading file...", Path(file_path).name)
            df = FileLoader.load(file_path)
            
            # Validate
            if df.empty:
                raise ValueError("File is empty - no data found")
            
            if len(df) < Config.MIN_ROWS:
                raise ValueError(f"Dataset too small. Need at least {Config.MIN_ROWS} rows, found {len(df)}")
            
            # Warn for large files
            if len(df) > Config.MAX_ROWS_WARN:
                progress.close()
                proceed = messagebox.askyesno(
                    "Large Dataset",
                    f"This file has {len(df):,} rows.\n"
                    "Processing may take a while.\n\n"
                    "Continue?",
                    icon='warning'
                )
                if not proceed:
                    return None, None, None, None
                progress = ProgressDialog(self.root, "Processing Data")
            
            # Process
            progress.update("Cleaning data...", f"{len(df):,} rows loaded")
            processor = DataProcessor(df)
            df, processing_notes = processor.process()
            
            # Build models
            progress.update("Building models...", "Training ML algorithms")
            builder = ModelBuilder(df)
            df, accuracy, model_note = builder.build()
            
            progress.close()
            
            # Show preview
            messagebox.showinfo(
                "Processing Complete",
                f"✅ Data processed successfully!\n\n"
                f"Rows: {len(df):,}\n"
                f"Columns: {len(df.columns)}\n"
                f"Segments: {df['Segment'].nunique()}\n\n"
                f"Model: {model_note}\n\n"
                f"Click OK to choose save location."
            )
            
            return df, accuracy, model_note, processing_notes
            
        except Exception as e:
            progress.close()
            raise
    
    def _select_save_location(self, input_path):
        """Select where to save the report."""
        default_name = Path(input_path).stem + "_analysis.xlsx"
        
        save_path = filedialog.asksaveasfilename(
            title="Save Analysis Report",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not save_path:
            messagebox.showinfo("Cancelled", "Save cancelled. No report generated.")
            return None
        
        return save_path
    
    def _generate_report(self, df, save_path, accuracy, model_note, processing_notes):
        """Generate the Excel report."""
        progress = ProgressDialog(self.root, "Generating Report")
        
        try:
            progress.update("Creating report...", "Building sheets and charts")
            
            generator = ReportGenerator(df, accuracy, model_note, processing_notes)
            generator.generate(save_path)
            
            progress.close()
            
        except Exception as e:
            progress.close()
            raise
    
    def _show_success(self, df, save_path, accuracy, model_note):
        """Show success message."""
        accuracy_str = f"{accuracy:.1%}" if accuracy else "N/A"
        
        high_risk = (df['Churn_Prob'] > 0.7).sum()
        high_value = (df['Predicted_CLV'] > df['Predicted_CLV'].median()).sum()
        
        messagebox.showinfo(
            "Success! 🎉",
            f"Analysis Complete!\n\n"
            f"📊 Customers Analyzed: {len(df):,}\n"
            f"🎯 Model Accuracy: {accuracy_str}\n"
            f"⚠️ High Risk Customers: {high_risk:,}\n"
            f"💎 High Value Customers: {high_value:,}\n\n"
            f"📁 Report saved to:\n{save_path}\n\n"
            f"Sheets included:\n"
            f"• Dashboard - Visual overview with KPIs\n"
            f"• Summary - Key metrics and notes\n"
            f"• Data - Full dataset with predictions\n"
            f"• Segments - Customer segment breakdown\n"
            f"• ChurnDist - Churn risk distribution\n"
            f"• CLVDist - Lifetime value tiers\n"
            f"• SpendDist - Spending patterns\n"
            f"• TenureDist - Customer tenure groups"
        )
    
    def _show_error(self, error):
        """Show error message."""
        error_msg = str(error)
        
        # Add helpful context
        hints = []
        if 'encoding' in error_msg.lower():
            hints.append("Try saving your file with UTF-8 encoding")
        if 'permission' in error_msg.lower():
            hints.append("Make sure the file is not open in another program")
        if 'memory' in error_msg.lower():
            hints.append("Try with a smaller dataset")
        if 'empty' in error_msg.lower():
            hints.append("Ensure your file contains data")
        
        hint_text = "\n\nSuggestions:\n• " + "\n• ".join(hints) if hints else ""
        
        # Log the full error for debugging
        logger.log(f"Error: {error_msg}", "ERROR")
        logger.log(traceback.format_exc(), "DEBUG")
        
        messagebox.showerror(
            "Error",
            f"⚠️ An error occurred:\n\n{error_msg}{hint_text}\n\n"
            f"Please ensure your file:\n"
            f"• Is not corrupted or open elsewhere\n"
            f"• Contains valid tabular data\n"
            f"• Has at least {Config.MIN_ROWS} rows"
        )


# ==============================
# ENTRY POINT
# ==============================
def main():
    """Main entry point."""
    app = CustomerAnalyticsApp()
    app.run()


if __name__ == "__main__":
    main()

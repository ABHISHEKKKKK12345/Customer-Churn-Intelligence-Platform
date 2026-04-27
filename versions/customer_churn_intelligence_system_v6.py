# ==============================
# CUSTOMER ANALYTICS TOOL 
# Fully Robust, Production-Ready
# Handles ALL Edge Cases
# ==============================

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import warnings
import os
import sys
import re
from pathlib import Path
from datetime import datetime
import traceback
import hashlib

warnings.filterwarnings("ignore")

# Scikit-learn imports with fallback
try:
    from sklearn.model_selection import train_test_split, cross_val_score
    from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier, IsolationForest
    from sklearn.linear_model import Ridge, LogisticRegression
    from sklearn.preprocessing import StandardScaler, RobustScaler
    from sklearn.metrics import accuracy_score, roc_auc_score
    from sklearn.impute import SimpleImputer
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False


# ==============================
# CONFIGURATION
# ==============================
class Config:
    """Central configuration for the application."""
    APP_NAME = "Customer Analytics Tool"
    VERSION = "3.0"
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
        '.pq': 'Parquet',
        '.tsv': 'TSV',
        '.txt': 'Text'
    }
    ENCODINGS = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1', 'utf-16', 'ascii', 'utf-8-sig']
    
    # Realistic value ranges for validation
    VALID_RANGES = {
        'tenure': {'min': 0, 'max': 600, 'unit': 'months'},  # 0-50 years
        'MonthlyCharges': {'min': 0, 'max': 50000, 'unit': 'currency'},
        'TotalCharges': {'min': 0, 'max': 10000000, 'unit': 'currency'},
        'Churn': {'min': 0, 'max': 1, 'unit': 'binary'}
    }
    
    # Extended column name mappings (lowercase keys)
    COLUMN_MAPPINGS = {
        # Tenure variations
        'tenure': 'tenure', 'tenure_months': 'tenure', 'customer_tenure': 'tenure',
        'months': 'tenure', 'duration': 'tenure', 'time_as_customer': 'tenure',
        'account_length': 'tenure', 'customer_age_months': 'tenure', 'loyalty_months': 'tenure',
        'months_as_customer': 'tenure', 'subscription_length': 'tenure', 'membership_duration': 'tenure',
        'account_age': 'tenure', 'customer_since_months': 'tenure', 'service_tenure': 'tenure',
        'tenure_days': 'tenure_days', 'tenure_years': 'tenure_years', 'days_as_customer': 'tenure_days',
        'years_as_customer': 'tenure_years', 'customer_lifetime_days': 'tenure_days',
        
        # Monthly Charges variations
        'monthlycharges': 'MonthlyCharges', 'monthly_charges': 'MonthlyCharges',
        'monthly_charge': 'MonthlyCharges', 'monthly_fee': 'MonthlyCharges',
        'monthly_payment': 'MonthlyCharges', 'monthly_amount': 'MonthlyCharges',
        'monthly_spend': 'MonthlyCharges', 'monthly_bill': 'MonthlyCharges',
        'monthly_cost': 'MonthlyCharges', 'recurring_charge': 'MonthlyCharges',
        'mrr': 'MonthlyCharges', 'monthly_revenue': 'MonthlyCharges',
        'subscription_fee': 'MonthlyCharges', 'plan_price': 'MonthlyCharges',
        'monthly_rate': 'MonthlyCharges', 'billing_amount': 'MonthlyCharges',
        'avg_monthly_spend': 'MonthlyCharges', 'average_monthly_charge': 'MonthlyCharges',
        
        # Annual charges (need conversion)
        'annual_charges': 'AnnualCharges', 'yearly_charges': 'AnnualCharges',
        'annual_fee': 'AnnualCharges', 'yearly_fee': 'AnnualCharges',
        'annual_payment': 'AnnualCharges', 'yearly_payment': 'AnnualCharges',
        'annual_spend': 'AnnualCharges', 'yearly_spend': 'AnnualCharges',
        'arr': 'AnnualCharges', 'annual_revenue': 'AnnualCharges',
        
        # Total Charges variations
        'totalcharges': 'TotalCharges', 'total_charges': 'TotalCharges',
        'total_charge': 'TotalCharges', 'total_payment': 'TotalCharges',
        'lifetime_value': 'TotalCharges', 'total_spend': 'TotalCharges',
        'total_amount': 'TotalCharges', 'cumulative_charges': 'TotalCharges',
        'total_revenue': 'TotalCharges', 'customer_value': 'TotalCharges',
        'ltv': 'TotalCharges', 'clv': 'TotalCharges', 'cltv': 'TotalCharges',
        'total_billings': 'TotalCharges', 'all_time_spend': 'TotalCharges',
        'cumulative_spend': 'TotalCharges', 'historical_charges': 'TotalCharges',
        
        # Churn variations (including tricky ones)
        'churn': 'Churn', 'churned': 'Churn', 'is_churn': 'Churn',
        'has_churned': 'Churn', 'customer_churn': 'Churn', 'attrition': 'Churn',
        'left': 'Churn', 'exited': 'Churn', 'cancelled': 'Churn',
        'terminated': 'Churn', 'inactive': 'Churn', 'lost': 'Churn',
        'customer_status': 'CustomerStatus', 'status': 'CustomerStatus',
        'account_status': 'CustomerStatus', 'subscription_status': 'CustomerStatus',
        'active': 'Active', 'is_active': 'Active', 'current_customer': 'Active',
        'retained': 'Active', 'is_retained': 'Active',
        
        # Payment status (often confused with churn)
        'payment_status': 'PaymentStatus', 'billing_status': 'PaymentStatus',
        'invoice_status': 'PaymentStatus', 'transaction_status': 'PaymentStatus'
    }
    
    # Status values mapping
    STATUS_MAPPINGS = {
        # Positive/Active (churn = 0)
        'active': 0, 'current': 0, 'retained': 0, 'staying': 0, 'continued': 0,
        'subscribed': 0, 'enrolled': 0, 'member': 0, 'yes_active': 0,
        'in_good_standing': 0, 'good': 0, 'ok': 0, 'valid': 0, 'live': 0,
        
        # Negative/Churned (churn = 1)
        'churned': 1, 'cancelled': 1, 'canceled': 1, 'terminated': 1,
        'inactive': 1, 'left': 1, 'exited': 1, 'lost': 1, 'closed': 1,
        'suspended': 1, 'expired': 1, 'lapsed': 1, 'dormant': 1,
        'unsubscribed': 1, 'withdrawn': 1, 'dropped': 1
    }
    
    # Domain detection patterns
    DOMAIN_PATTERNS = {
        'customer_analytics': [
            'customer', 'subscriber', 'member', 'client', 'user', 'account',
            'churn', 'tenure', 'monthly', 'charges', 'subscription', 'plan'
        ],
        'hr_payroll': [
            'employee', 'salary', 'department', 'hire_date', 'job_title',
            'payroll', 'benefits', 'manager', 'performance', 'hr'
        ],
        'sales': [
            'order', 'product', 'quantity', 'price', 'discount', 'sale',
            'invoice', 'sku', 'inventory', 'shipping'
        ],
        'financial': [
            'transaction', 'balance', 'credit', 'debit', 'interest',
            'loan', 'mortgage', 'investment', 'portfolio'
        ]
    }


# ==============================
# LOGGING UTILITY
# ==============================
class Logger:
    """Enhanced logging for debugging and audit trail."""
    
    def __init__(self):
        self.logs = []
        self.warnings = []
        self.errors = []
    
    def log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        entry = f"[{timestamp}] [{level}] {message}"
        self.logs.append(entry)
        
        if level == "WARN":
            self.warnings.append(message)
        elif level == "ERROR":
            self.errors.append(message)
    
    def get_logs(self):
        return "\n".join(self.logs)
    
    def get_warnings(self):
        return self.warnings.copy()
    
    def get_errors(self):
        return self.errors.copy()
    
    def clear(self):
        self.logs = []
        self.warnings = []
        self.errors = []


logger = Logger()


# ==============================
# DATA QUALITY ANALYZER
# ==============================
class DataQualityAnalyzer:
    """Comprehensive data quality assessment."""
    
    def __init__(self, df):
        self.df = df
        self.issues = []
        self.quality_score = 100
        self.is_valid = True
        self.domain = None
        self.recommendations = []
    
    def analyze(self):
        """Run full quality analysis."""
        self._check_basic_validity()
        self._detect_domain()
        self._check_data_quality()
        self._check_column_relevance()
        self._calculate_quality_score()
        
        return {
            'is_valid': self.is_valid,
            'quality_score': self.quality_score,
            'domain': self.domain,
            'issues': self.issues,
            'recommendations': self.recommendations
        }
    
    def _check_basic_validity(self):
        """Check basic file validity."""
        # Empty check
        if self.df is None or self.df.empty:
            self.issues.append("CRITICAL: Dataset is empty")
            self.is_valid = False
            self.quality_score = 0
            return
        
        # Row count check
        if len(self.df) < Config.MIN_ROWS:
            self.issues.append(f"CRITICAL: Only {len(self.df)} rows (minimum {Config.MIN_ROWS} required)")
            self.is_valid = False
            self.quality_score = 0
            return
        
        # Column count check
        if len(self.df.columns) < 2:
            self.issues.append("CRITICAL: Only 1 column found - not tabular data")
            self.is_valid = False
            self.quality_score = 0
            return
        
        # Check for corrupted data (all same value)
        unique_counts = self.df.nunique()
        if (unique_counts == 1).all():
            self.issues.append("CRITICAL: All columns have single values - corrupted or placeholder data")
            self.is_valid = False
            self.quality_score = 0
            return
        
        # Check for random text (no numeric columns at all)
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        potentially_numeric = 0
        for col in self.df.columns:
            try:
                converted = pd.to_numeric(self.df[col].astype(str).str.replace(r'[$,€£%]', '', regex=True), errors='coerce')
                if converted.notna().sum() > len(self.df) * 0.3:
                    potentially_numeric += 1
            except:
                pass
        
        if len(numeric_cols) == 0 and potentially_numeric == 0:
            self.issues.append("WARNING: No numeric columns found - limited analysis possible")
            self.quality_score -= 30
    
    def _detect_domain(self):
        """Detect the data domain/type."""
        column_text = ' '.join(self.df.columns).lower()
        
        domain_scores = {}
        for domain, patterns in Config.DOMAIN_PATTERNS.items():
            score = sum(1 for pattern in patterns if pattern in column_text)
            domain_scores[domain] = score
        
        best_domain = max(domain_scores, key=domain_scores.get)
        best_score = domain_scores[best_domain]
        
        if best_score >= 3:
            self.domain = best_domain
            logger.log(f"Detected domain: {best_domain} (confidence: {best_score})")
        elif best_score >= 1:
            self.domain = best_domain
            self.issues.append(f"WARNING: Dataset appears to be {best_domain.replace('_', ' ')} data - may not be ideal for customer analytics")
            self.quality_score -= 15
            logger.log(f"Possible domain: {best_domain} (low confidence: {best_score})")
        else:
            self.domain = 'unknown'
            self.issues.append("WARNING: Could not identify dataset type - will attempt generic analysis")
            self.quality_score -= 20
            logger.log("Unknown domain - attempting generic analysis")
        
        # Special handling for completely unrelated datasets
        if self.domain == 'hr_payroll':
            self.recommendations.append("This appears to be HR/payroll data. Customer analytics may not be applicable.")
            self.recommendations.append("Consider using employee ID as customer ID and salary-related fields for value analysis.")
        elif self.domain == 'sales':
            self.recommendations.append("This appears to be sales/order data. Will aggregate by customer if possible.")
    
    def _check_data_quality(self):
        """Check overall data quality."""
        # Missing value analysis
        missing_pct = (self.df.isna().sum().sum() / (len(self.df) * len(self.df.columns))) * 100
        
        if missing_pct > 50:
            self.issues.append(f"CRITICAL: {missing_pct:.1f}% missing values - very poor data quality")
            self.quality_score -= 40
        elif missing_pct > 20:
            self.issues.append(f"WARNING: {missing_pct:.1f}% missing values")
            self.quality_score -= 20
        elif missing_pct > 5:
            self.issues.append(f"INFO: {missing_pct:.1f}% missing values (acceptable)")
            self.quality_score -= 5
        
        # Duplicate row analysis
        dup_pct = (self.df.duplicated().sum() / len(self.df)) * 100
        if dup_pct > 30:
            self.issues.append(f"WARNING: {dup_pct:.1f}% duplicate rows")
            self.quality_score -= 15
        elif dup_pct > 10:
            self.issues.append(f"INFO: {dup_pct:.1f}% duplicate rows")
            self.quality_score -= 5
        
        # Check for columns that are entirely null
        null_cols = self.df.columns[self.df.isna().all()].tolist()
        if null_cols:
            self.issues.append(f"WARNING: {len(null_cols)} completely empty columns: {null_cols[:5]}")
            self.quality_score -= 10
    
    def _check_column_relevance(self):
        """Check if key customer analytics columns are present or inferable."""
        column_lower = [c.lower() for c in self.df.columns]
        column_text = ' '.join(column_lower)
        
        # Check for customer identifier
        has_customer_id = any(kw in column_text for kw in ['customer', 'user', 'account', 'member', 'subscriber', 'client', 'id'])
        if not has_customer_id:
            self.issues.append("INFO: No obvious customer identifier column found")
            self.recommendations.append("First column or row index will be used as customer identifier")
        
        # Check for value-related columns
        has_value = any(kw in column_text for kw in ['charge', 'payment', 'spend', 'revenue', 'amount', 'price', 'fee', 'bill', 'cost', 'value', 'salary', 'income'])
        if not has_value:
            self.issues.append("WARNING: No monetary/value columns detected")
            self.recommendations.append("Will attempt to use available numeric columns for value estimation")
            self.quality_score -= 10
        
        # Check for time/tenure columns
        has_time = any(kw in column_text for kw in ['tenure', 'month', 'day', 'year', 'date', 'time', 'duration', 'age', 'length', 'since', 'period'])
        if not has_time:
            self.issues.append("INFO: No tenure/time columns detected")
            self.recommendations.append("Will estimate tenure from available data or use defaults")
        
        # Check for churn/status columns
        has_churn = any(kw in column_text for kw in ['churn', 'status', 'active', 'cancelled', 'terminated', 'left', 'exited', 'attrition', 'retained'])
        if not has_churn:
            self.issues.append("INFO: No churn/status column detected")
            self.recommendations.append("Will create synthetic churn probability based on available patterns")
    
    def _calculate_quality_score(self):
        """Calculate final quality score."""
        self.quality_score = max(0, min(100, self.quality_score))
        
        if self.quality_score < 30:
            self.is_valid = False
            self.issues.append("CRITICAL: Data quality too low for meaningful analysis")
        
        logger.log(f"Data quality score: {self.quality_score}/100")


# ==============================
# OUTLIER HANDLER
# ==============================
class OutlierHandler:
    """Handle outliers and unrealistic values."""
    
    def __init__(self, df):
        self.df = df.copy()
        self.outlier_report = {}
    
    def process(self):
        """Process outliers in key columns."""
        self._handle_tenure_outliers()
        self._handle_charge_outliers()
        self._handle_generic_outliers()
        
        return self.df, self.outlier_report
    
    def _handle_tenure_outliers(self):
        """Handle unrealistic tenure values."""
        if 'tenure' not in self.df.columns:
            return
        
        col = 'tenure'
        original_stats = self.df[col].describe()
        outlier_count = 0
        
        # Check for negative values
        neg_mask = self.df[col] < 0
        if neg_mask.any():
            outlier_count += neg_mask.sum()
            self.df.loc[neg_mask, col] = self.df[col].abs()
            logger.log(f"Fixed {neg_mask.sum()} negative tenure values")
        
        # Check for unrealistic high values (e.g., 99999)
        max_reasonable = Config.VALID_RANGES['tenure']['max']
        high_mask = self.df[col] > max_reasonable
        
        if high_mask.any():
            outlier_count += high_mask.sum()
            # Cap at reasonable maximum or use median
            median_val = self.df.loc[~high_mask, col].median()
            if pd.notna(median_val):
                self.df.loc[high_mask, col] = median_val
            else:
                self.df.loc[high_mask, col] = 24  # Default 2 years
            logger.log(f"Capped {high_mask.sum()} extreme tenure values (>{max_reasonable})")
        
        # Check for values that look like days instead of months
        if self.df[col].median() > 365:
            # Likely in days, convert to months
            self.df[col] = self.df[col] / 30
            logger.log("Tenure appears to be in days - converted to months")
            self.outlier_report['tenure_conversion'] = 'days_to_months'
        elif self.df[col].median() > 50 and self.df[col].max() < 100:
            # Might be in years
            if self.df[col].max() <= 80:  # Reasonable year range
                self.df[col] = self.df[col] * 12
                logger.log("Tenure appears to be in years - converted to months")
                self.outlier_report['tenure_conversion'] = 'years_to_months'
        
        self.outlier_report['tenure_outliers'] = outlier_count
    
    def _handle_charge_outliers(self):
        """Handle unrealistic charge values."""
        charge_cols = ['MonthlyCharges', 'TotalCharges', 'AnnualCharges']
        
        for col in charge_cols:
            if col not in self.df.columns:
                continue
            
            outlier_count = 0
            
            # Handle negative values
            neg_mask = self.df[col] < 0
            if neg_mask.any():
                outlier_count += neg_mask.sum()
                self.df.loc[neg_mask, col] = self.df[col].abs()
            
            # Handle extreme outliers using IQR method
            Q1 = self.df[col].quantile(0.25)
            Q3 = self.df[col].quantile(0.75)
            IQR = Q3 - Q1
            
            lower_bound = max(0, Q1 - 3 * IQR)  # Use 3*IQR for less aggressive clipping
            upper_bound = Q3 + 3 * IQR
            
            # Don't clip if the range is too small (could be legitimate)
            if upper_bound > lower_bound * 100:  # If upper is 100x lower, likely outliers
                extreme_mask = (self.df[col] > upper_bound) | (self.df[col] < lower_bound)
                if extreme_mask.any() and extreme_mask.sum() < len(self.df) * 0.1:  # Less than 10%
                    outlier_count += extreme_mask.sum()
                    median_val = self.df.loc[~extreme_mask, col].median()
                    self.df.loc[extreme_mask, col] = np.clip(
                        self.df.loc[extreme_mask, col],
                        lower_bound,
                        upper_bound
                    )
                    logger.log(f"Handled {extreme_mask.sum()} extreme values in {col}")
            
            self.outlier_report[f'{col}_outliers'] = outlier_count
    
    def _handle_generic_outliers(self):
        """Handle outliers in other numeric columns using Isolation Forest."""
        if not SKLEARN_AVAILABLE:
            return
        
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns.tolist()
        key_cols = ['tenure', 'MonthlyCharges', 'TotalCharges', 'Churn', 'Churn_Prob', 'Predicted_CLV']
        other_cols = [c for c in numeric_cols if c not in key_cols]
        
        if len(other_cols) < 2 or len(self.df) < 50:
            return
        
        try:
            # Use Isolation Forest for multivariate outlier detection
            X = self.df[other_cols].fillna(self.df[other_cols].median())
            
            iso_forest = IsolationForest(contamination=0.05, random_state=42, n_jobs=-1)
            outlier_labels = iso_forest.fit_predict(X)
            
            outlier_mask = outlier_labels == -1
            outlier_count = outlier_mask.sum()
            
            if outlier_count > 0 and outlier_count < len(self.df) * 0.1:
                # Flag outliers but don't remove them
                self.df['_is_outlier'] = outlier_mask.astype(int)
                self.outlier_report['multivariate_outliers'] = outlier_count
                logger.log(f"Flagged {outlier_count} multivariate outliers")
        except Exception as e:
            logger.log(f"Outlier detection skipped: {str(e)}", "WARN")


# ==============================
# SEMANTIC COLUMN MAPPER
# ==============================
class SemanticColumnMapper:
    """Intelligent column mapping with semantic understanding."""
    
    def __init__(self, df):
        self.df = df.copy()
        self.mapping_report = {}
        self.confidence_scores = {}
    
    def map_columns(self):
        """Perform intelligent column mapping."""
        self._apply_name_based_mapping()
        self._apply_value_based_inference()
        self._handle_status_columns()
        self._convert_annual_to_monthly()
        self._validate_mappings()
        
        return self.df, self.mapping_report, self.confidence_scores
    
    def _apply_name_based_mapping(self):
        """Map columns based on name matching."""
        mapped = {}
        
        for col in self.df.columns:
            col_lower = col.lower().strip().replace(' ', '_').replace('-', '_')
            
            # Direct match
            if col_lower in Config.COLUMN_MAPPINGS:
                target = Config.COLUMN_MAPPINGS[col_lower]
                if target not in self.df.columns and target not in mapped.values():
                    mapped[col] = target
                    self.confidence_scores[target] = 0.9
                    continue
            
            # Partial match
            for pattern, target in Config.COLUMN_MAPPINGS.items():
                if pattern in col_lower or col_lower in pattern:
                    if target not in self.df.columns and target not in mapped.values():
                        mapped[col] = target
                        self.confidence_scores[target] = 0.7
                        break
        
        if mapped:
            self.df.rename(columns=mapped, inplace=True)
            self.mapping_report['name_based'] = mapped
            logger.log(f"Name-based mapping: {mapped}")
    
    def _apply_value_based_inference(self):
        """Infer column meanings from value patterns."""
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns.tolist()
        
        # Infer tenure if not found
        if 'tenure' not in self.df.columns:
            best_col, best_score = self._find_tenure_column(numeric_cols)
            if best_col and best_score > 0.5:
                self.df['tenure'] = self.df[best_col].copy()
                self.mapping_report['tenure_inferred_from'] = best_col
                self.confidence_scores['tenure'] = best_score
                logger.log(f"Inferred tenure from '{best_col}' (confidence: {best_score:.2f})")
        
        # Infer MonthlyCharges if not found
        if 'MonthlyCharges' not in self.df.columns:
            best_col, best_score = self._find_monthly_charges_column(numeric_cols)
            if best_col and best_score > 0.5:
                self.df['MonthlyCharges'] = self.df[best_col].copy()
                self.mapping_report['MonthlyCharges_inferred_from'] = best_col
                self.confidence_scores['MonthlyCharges'] = best_score
                logger.log(f"Inferred MonthlyCharges from '{best_col}' (confidence: {best_score:.2f})")
        
        # Infer TotalCharges if not found
        if 'TotalCharges' not in self.df.columns:
            best_col, best_score = self._find_total_charges_column(numeric_cols)
            if best_col and best_score > 0.5:
                self.df['TotalCharges'] = self.df[best_col].copy()
                self.mapping_report['TotalCharges_inferred_from'] = best_col
                self.confidence_scores['TotalCharges'] = best_score
                logger.log(f"Inferred TotalCharges from '{best_col}' (confidence: {best_score:.2f})")
    
    def _find_tenure_column(self, numeric_cols):
        """Find the most likely tenure column."""
        best_col = None
        best_score = 0
        
        excluded = ['MonthlyCharges', 'TotalCharges', 'Churn', 'AnnualCharges']
        candidates = [c for c in numeric_cols if c not in excluded]
        
        for col in candidates:
            score = 0
            col_lower = col.lower()
            
            # Name hints
            tenure_hints = ['tenure', 'month', 'duration', 'time', 'age', 'length', 'period']
            if any(hint in col_lower for hint in tenure_hints):
                score += 0.4
            
            # Value range check (0-600 months is reasonable)
            col_data = self.df[col].dropna()
            if len(col_data) == 0:
                continue
            
            min_val, max_val, median_val = col_data.min(), col_data.max(), col_data.median()
            
            # Typical tenure characteristics
            if 0 <= min_val and max_val <= 600 and 1 <= median_val <= 120:
                score += 0.3
            
            # Integer-like values (tenure is often whole months)
            if (col_data == col_data.astype(int)).mean() > 0.9:
                score += 0.2
            
            # Distribution check (tenure usually has some spread)
            if col_data.std() > 1:
                score += 0.1
            
            if score > best_score:
                best_score = score
                best_col = col
        
        return best_col, best_score
    
    def _find_monthly_charges_column(self, numeric_cols):
        """Find the most likely monthly charges column."""
        best_col = None
        best_score = 0
        
        excluded = ['tenure', 'TotalCharges', 'Churn', 'AnnualCharges']
        candidates = [c for c in numeric_cols if c not in excluded]
        
        for col in candidates:
            score = 0
            col_lower = col.lower()
            
            # Name hints
            monthly_hints = ['monthly', 'charge', 'fee', 'payment', 'bill', 'cost', 'price', 'rate']
            annual_hints = ['annual', 'yearly', 'total', 'lifetime', 'cumulative']
            
            if any(hint in col_lower for hint in monthly_hints):
                score += 0.3
            if any(hint in col_lower for hint in annual_hints):
                score -= 0.2  # Penalize annual-looking columns
            
            # Value range check
            col_data = self.df[col].dropna()
            if len(col_data) == 0:
                continue
            
            median_val = col_data.median()
            
            # Typical monthly charge range ($10-$500)
            if 10 <= median_val <= 500:
                score += 0.4
            elif 1 <= median_val <= 1000:
                score += 0.2
            
            # Decimal values (charges often have cents)
            has_decimals = (col_data != col_data.astype(int)).mean() > 0.1
            if has_decimals:
                score += 0.1
            
            if score > best_score:
                best_score = score
                best_col = col
        
        return best_col, best_score
    
    def _find_total_charges_column(self, numeric_cols):
        """Find the most likely total charges column."""
        best_col = None
        best_score = 0
        
        excluded = ['tenure', 'MonthlyCharges', 'Churn', 'AnnualCharges']
        candidates = [c for c in numeric_cols if c not in excluded]
        
        for col in candidates:
            score = 0
            col_lower = col.lower()
            
            # Name hints
            total_hints = ['total', 'lifetime', 'cumulative', 'ltv', 'clv', 'value', 'overall']
            if any(hint in col_lower for hint in total_hints):
                score += 0.4
            
            # Value range check
            col_data = self.df[col].dropna()
            if len(col_data) == 0:
                continue
            
            median_val = col_data.median()
            
            # Total charges should be larger than monthly
            if 'MonthlyCharges' in self.df.columns:
                monthly_median = self.df['MonthlyCharges'].median()
                if median_val > monthly_median * 3:
                    score += 0.3
            
            # Typical total charge range ($100-$100000)
            if 100 <= median_val <= 100000:
                score += 0.2
            
            if score > best_score:
                best_score = score
                best_col = col
        
        return best_col, best_score
    
    def _handle_status_columns(self):
        """Handle status columns that might indicate churn."""
        if 'Churn' in self.df.columns:
            # Already have churn, validate it
            self._validate_churn_column()
            return
        
        # Look for status-like columns
        status_cols = ['CustomerStatus', 'Active', 'PaymentStatus']
        
        for col in status_cols:
            if col not in self.df.columns:
                continue
            
            values = self.df[col].astype(str).str.lower().str.strip()
            
            if col == 'Active':
                # Active column: active=0 churn, inactive=1 churn (inverted)
                self.df['Churn'] = values.apply(
                    lambda x: 0 if x in ['yes', 'true', '1', 'active', 'y'] else 1
                )
                self.mapping_report['churn_from'] = f'{col} (inverted)'
                self.confidence_scores['Churn'] = 0.8
                logger.log(f"Created Churn from '{col}' (inverted logic)")
                return
            
            elif col == 'CustomerStatus':
                # Map status values to churn
                def map_status(val):
                    val = str(val).lower().strip()
                    if val in Config.STATUS_MAPPINGS:
                        return Config.STATUS_MAPPINGS[val]
                    # Default mapping for unknown values
                    if any(pos in val for pos in ['activ', 'current', 'good', 'valid']):
                        return 0
                    if any(neg in val for neg in ['churn', 'cancel', 'termin', 'inactiv', 'left', 'exit']):
                        return 1
                    return 0  # Default to not churned
                
                self.df['Churn'] = values.apply(map_status)
                self.mapping_report['churn_from'] = f'{col} (status mapping)'
                self.confidence_scores['Churn'] = 0.7
                logger.log(f"Created Churn from '{col}' status mapping")
                return
            
            elif col == 'PaymentStatus':
                # Payment status is NOT churn - don't use it
                logger.log(f"Skipped '{col}' - payment status ≠ churn status", "WARN")
                self.mapping_report['payment_status_warning'] = 'PaymentStatus column found but not used for churn'
        
        # No status column found, look for binary columns
        self._infer_churn_from_binary()
    
    def _validate_churn_column(self):
        """Validate existing churn column."""
        churn_values = self.df['Churn'].astype(str).str.lower().str.strip()
        unique_values = churn_values.unique()
        
        # Check if it's already binary numeric
        try:
            numeric_churn = pd.to_numeric(self.df['Churn'], errors='coerce')
            if numeric_churn.notna().all() and set(numeric_churn.unique()).issubset({0, 1, 0.0, 1.0}):
                self.df['Churn'] = numeric_churn.astype(int)
                self.confidence_scores['Churn'] = 0.95
                return
        except:
            pass
        
        # Convert string values to binary
        def convert_churn(val):
            val = str(val).lower().strip()
            if val in ['yes', 'y', 'true', '1', 'churned', 'left', 'exited', 'cancelled']:
                return 1
            elif val in ['no', 'n', 'false', '0', 'stayed', 'active', 'retained']:
                return 0
            else:
                return 0  # Default
        
        self.df['Churn'] = churn_values.apply(convert_churn)
        self.confidence_scores['Churn'] = 0.8
        logger.log("Converted Churn column to binary")
    
    def _infer_churn_from_binary(self):
        """Infer churn from any binary column."""
        for col in self.df.columns:
            unique_vals = self.df[col].nunique()
            if unique_vals == 2:
                values = self.df[col].astype(str).str.lower().str.strip().unique()
                
                # Check if values suggest churn
                churn_positive = ['yes', 'y', 'true', '1', 'churned']
                churn_negative = ['no', 'n', 'false', '0', 'stayed']
                
                if any(v in churn_positive for v in values) and any(v in churn_negative for v in values):
                    self.df['Churn'] = self.df[col].astype(str).str.lower().str.strip().apply(
                        lambda x: 1 if x in churn_positive else 0
                    )
                    self.mapping_report['churn_from'] = f'{col} (binary inference)'
                    self.confidence_scores['Churn'] = 0.5
                    logger.log(f"Inferred Churn from binary column '{col}'")
                    return
        
        # No suitable column found - create placeholder
        self.df['Churn'] = 0
        self.mapping_report['churn_from'] = 'default (no churn data found)'
        self.confidence_scores['Churn'] = 0.1
        logger.log("No churn column found - using placeholder", "WARN")
    
    def _convert_annual_to_monthly(self):
        """Convert annual charges to monthly if needed."""
        if 'AnnualCharges' in self.df.columns and 'MonthlyCharges' not in self.df.columns:
            self.df['MonthlyCharges'] = self.df['AnnualCharges'] / 12
            self.mapping_report['annual_to_monthly'] = True
            logger.log("Converted AnnualCharges to MonthlyCharges (÷12)")
    
    def _validate_mappings(self):
        """Validate that mapped columns make sense."""
        # Check tenure vs total charges relationship
        if 'tenure' in self.df.columns and 'TotalCharges' in self.df.columns and 'MonthlyCharges' in self.df.columns:
            expected_total = self.df['MonthlyCharges'] * self.df['tenure']
            actual_total = self.df['TotalCharges']
            
            # Calculate correlation
            correlation = expected_total.corr(actual_total)
            
            if correlation < 0.3 and not pd.isna(correlation):
                logger.log(f"Warning: tenure × MonthlyCharges doesn't correlate well with TotalCharges (r={correlation:.2f})", "WARN")
                self.mapping_report['validation_warning'] = 'Inconsistent relationship between tenure, monthly, and total charges'


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
        
        file_size = path.stat().st_size
        if file_size == 0:
            raise ValueError("File is empty (0 bytes)")
        
        if file_size < 10:
            raise ValueError("File too small to contain valid data")
        
        logger.log(f"Loading file: {path.name} (Format: {ext}, Size: {file_size/1024:.1f} KB)")
        
        try:
            if ext == '.csv':
                return FileLoader._load_csv(file_path)
            elif ext == '.tsv':
                return FileLoader._load_csv(file_path, delimiter='\t')
            elif ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
                return FileLoader._load_excel(file_path)
            elif ext == '.json':
                return FileLoader._load_json(file_path)
            elif ext in ['.parquet', '.pq']:
                return FileLoader._load_parquet(file_path)
            elif ext == '.txt':
                return FileLoader._load_csv(file_path)  # Try CSV parser
            else:
                # Try CSV as fallback
                logger.log(f"Unknown extension {ext}, trying CSV parser")
                return FileLoader._load_csv(file_path)
        except Exception as e:
            logger.log(f"Load error: {str(e)}", "ERROR")
            raise
    
    @staticmethod
    def _load_csv(file_path, delimiter=None):
        """Load CSV with encoding detection and delimiter inference."""
        last_error = None
        
        for encoding in Config.ENCODINGS:
            try:
                # Detect delimiter if not specified
                if delimiter is None:
                    with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                        sample = f.read(16384)
                    
                    # Skip if sample is too short
                    if len(sample) < 10:
                        continue
                    
                    # Count potential delimiters
                    delimiters = {',': 0, ';': 0, '\t': 0, '|': 0, ':': 0}
                    for delim in delimiters:
                        delimiters[delim] = sample.count(delim)
                    
                    best_delim = max(delimiters, key=delimiters.get)
                    if delimiters[best_delim] == 0:
                        best_delim = ','
                else:
                    best_delim = delimiter
                
                logger.log(f"Trying delimiter: '{best_delim}', encoding: {encoding}")
                
                # Load with detected settings
                df = pd.read_csv(
                    file_path,
                    encoding=encoding,
                    delimiter=best_delim,
                    on_bad_lines='skip',
                    low_memory=False,
                    dtype=str,
                    skipinitialspace=True
                )
                
                # Validate loaded data
                if df.empty:
                    continue
                
                # Check if we got actual columns (not just one column with all data)
                if len(df.columns) == 1 and df.columns[0].count(best_delim) > 2:
                    # Wrong delimiter, try again
                    continue
                
                # Check for header row issues
                if df.columns[0].startswith('Unnamed'):
                    # Try reading without header
                    df_no_header = pd.read_csv(
                        file_path,
                        encoding=encoding,
                        delimiter=best_delim,
                        on_bad_lines='skip',
                        low_memory=False,
                        dtype=str,
                        header=None
                    )
                    if not df_no_header.empty:
                        # Use first row as header
                        df_no_header.columns = df_no_header.iloc[0]
                        df = df_no_header[1:].reset_index(drop=True)
                
                logger.log(f"CSV loaded: {len(df)} rows, {len(df.columns)} columns")
                return df
                
            except UnicodeDecodeError:
                last_error = f"Encoding error with {encoding}"
                continue
            except Exception as e:
                last_error = str(e)
                logger.log(f"CSV parse attempt failed ({encoding}): {str(e)}", "WARN")
                continue
        
        raise ValueError(f"Could not parse CSV with any supported encoding. Last error: {last_error}")
    
    @staticmethod
    def _load_excel(file_path):
        """Load Excel with sheet detection."""
        try:
            xl = pd.ExcelFile(file_path)
            sheets = xl.sheet_names
            logger.log(f"Excel file has {len(sheets)} sheet(s): {sheets[:5]}")
            
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
            try:
                df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
                return df
            except:
                pass
            raise
    
    @staticmethod
    def _load_json(file_path):
        """Load JSON with multiple format support."""
        import json
        
        # First, try to understand the JSON structure
        with open(file_path, 'r', encoding='utf-8') as f:
            try:
                data = json.load(f)
            except json.JSONDecodeError:
                # Try line-by-line
                f.seek(0)
                lines = f.readlines()
                data = [json.loads(line) for line in lines if line.strip()]
        
        if isinstance(data, list):
            if len(data) > 0 and isinstance(data[0], dict):
                return pd.DataFrame(data)
            else:
                return pd.DataFrame({'value': data})
        elif isinstance(data, dict):
            if all(isinstance(v, list) for v in data.values()):
                return pd.DataFrame(data)
            else:
                return pd.DataFrame([data])
        else:
            raise ValueError("Unsupported JSON structure")
    
    @staticmethod
    def _load_parquet(file_path):
        """Load Parquet file."""
        try:
            df = pd.read_parquet(file_path)
            return df.astype(str)
        except Exception as e:
            raise ValueError(f"Could not read Parquet: {str(e)}")


# ==============================
# DATA PROCESSOR
# ==============================
class DataProcessor:
    """Comprehensive data processing and cleaning."""
    
    def __init__(self, df, quality_report):
        self.df = df.copy()
        self.quality_report = quality_report
        self.original_columns = list(df.columns)
        self.processing_notes = []
    
    def process(self):
        """Run full processing pipeline."""
        self._clean_column_names()
        self._remove_empty_rows_cols()
        self._convert_types()
        
        # Semantic mapping
        mapper = SemanticColumnMapper(self.df)
        self.df, mapping_report, confidence_scores = mapper.map_columns()
        
        if mapping_report:
            self.processing_notes.append(f"Column mappings: {mapping_report}")
        
        # Ensure required columns exist
        self._ensure_required_columns()
        
        # Handle outliers
        outlier_handler = OutlierHandler(self.df)
        self.df, outlier_report = outlier_handler.process()
        
        if outlier_report:
            self.processing_notes.append(f"Outlier handling: {outlier_report}")
        
        # Clean data
        self._clean_data()
        
        # Engineer features
        self._engineer_features()
        
        return self.df, self.processing_notes
    
    def _clean_column_names(self):
        """Standardize column names."""
        new_cols = {}
        for col in self.df.columns:
            clean = str(col).strip()
            clean = clean.replace('\n', ' ').replace('\r', '').replace('\t', ' ')
            clean = re.sub(r'\s+', ' ', clean)  # Multiple spaces to single
            if clean != col:
                new_cols[col] = clean
        
        if new_cols:
            self.df.rename(columns=new_cols, inplace=True)
            logger.log(f"Cleaned {len(new_cols)} column names")
    
    def _remove_empty_rows_cols(self):
        """Remove completely empty rows and columns."""
        # Remove columns that are entirely null
        null_cols = self.df.columns[self.df.isna().all()].tolist()
        if null_cols:
            self.df.drop(columns=null_cols, inplace=True)
            logger.log(f"Removed {len(null_cols)} empty columns")
        
        # Remove rows that are entirely null
        initial_len = len(self.df)
        self.df.dropna(how='all', inplace=True)
        if len(self.df) < initial_len:
            logger.log(f"Removed {initial_len - len(self.df)} empty rows")
    
    def _convert_types(self):
        """Convert columns to appropriate types."""
        for col in self.df.columns:
            try:
                temp = self.df[col].astype(str).str.strip()
                temp = temp.replace(['', ' ', 'NA', 'N/A', 'null', 'NULL', 'None', 'nan', 'NaN', '-', '#N/A', '#REF!', '#VALUE!'], np.nan)
                temp = temp.str.replace(r'[$€£¥₹,]', '', regex=True)
                temp = temp.str.replace(r'%$', '', regex=True)
                
                numeric_temp = pd.to_numeric(temp, errors='coerce')
                valid_ratio = numeric_temp.notna().sum() / max(len(numeric_temp), 1)
                
                if valid_ratio > 0.5:
                    self.df[col] = numeric_temp
                    logger.log(f"Converted '{col}' to numeric ({valid_ratio*100:.0f}% valid)")
            except Exception as e:
                logger.log(f"Type conversion skipped for '{col}': {str(e)}", "WARN")
    
    def _ensure_required_columns(self):
        """Ensure all required columns exist with sensible defaults."""
        # Tenure
        if 'tenure' not in self.df.columns:
            numeric_cols = self.df.select_dtypes(include=[np.number]).columns.tolist()
            if numeric_cols:
                # Use first suitable numeric column
                for col in numeric_cols:
                    if col not in ['Churn', 'MonthlyCharges', 'TotalCharges']:
                        if self.df[col].min() >= 0 and self.df[col].max() <= 1000:
                            self.df['tenure'] = self.df[col].copy()
                            self.processing_notes.append(f"Used '{col}' as tenure")
                            break
            
            if 'tenure' not in self.df.columns:
                self.df['tenure'] = 12  # Default 1 year
                self.processing_notes.append("No tenure data - using default (12 months)")
        
        # Monthly Charges
        if 'MonthlyCharges' not in self.df.columns:
            numeric_cols = self.df.select_dtypes(include=[np.number]).columns.tolist()
            for col in numeric_cols:
                if col not in ['tenure', 'Churn', 'TotalCharges']:
                    median = self.df[col].median()
                    if 10 <= median <= 500:
                        self.df['MonthlyCharges'] = self.df[col].copy()
                        self.processing_notes.append(f"Used '{col}' as MonthlyCharges")
                        break
            
            if 'MonthlyCharges' not in self.df.columns:
                self.df['MonthlyCharges'] = 50  # Default
                self.processing_notes.append("No monthly charges data - using default ($50)")
        
        # Total Charges
        if 'TotalCharges' not in self.df.columns:
            if 'tenure' in self.df.columns and 'MonthlyCharges' in self.df.columns:
                self.df['TotalCharges'] = self.df['MonthlyCharges'] * self.df['tenure']
                self.processing_notes.append("Calculated TotalCharges from tenure × MonthlyCharges")
            else:
                self.df['TotalCharges'] = self.df.get('MonthlyCharges', 50) * 12
                self.processing_notes.append("No total charges data - using estimate")
        
        # Churn
        if 'Churn' not in self.df.columns:
            self.df['Churn'] = 0
            self.processing_notes.append("No churn data - defaulting to 0 (no churn)")
    
    def _clean_data(self):
        """Comprehensive data cleaning."""
        key_cols = ['tenure', 'MonthlyCharges', 'TotalCharges', 'Churn']
        
        for col in key_cols:
            if col not in self.df.columns:
                continue
            
            self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
            
            if col == 'Churn':
                self.df[col] = self.df[col].fillna(0).astype(int)
                self.df[col] = (self.df[col] > 0).astype(int)
            else:
                median = self.df[col].median()
                if pd.isna(median) or median <= 0:
                    median = 1 if col == 'tenure' else 50
                self.df[col] = self.df[col].fillna(median)
                self.df[col] = self.df[col].abs()
        
        # Ensure tenure is at least 1
        if 'tenure' in self.df.columns:
            self.df['tenure'] = self.df['tenure'].clip(lower=1)
        
        # Remove duplicates
        initial_len = len(self.df)
        self.df.drop_duplicates(inplace=True)
        if len(self.df) < initial_len:
            removed = initial_len - len(self.df)
            logger.log(f"Removed {removed} duplicate rows")
            self.processing_notes.append(f"Removed {removed} duplicate rows")
        
        self.df.reset_index(drop=True, inplace=True)
        logger.log("Data cleaning complete")
    
    def _engineer_features(self):
        """Create analytical features."""
        # Average monthly spend
        if 'TotalCharges' in self.df.columns and 'tenure' in self.df.columns:
            self.df['AvgMonthlySpend'] = self.df['TotalCharges'] / (self.df['tenure'].clip(lower=1))
            self.df['AvgMonthlySpend'] = self.df['AvgMonthlySpend'].clip(lower=0)
        else:
            self.df['AvgMonthlySpend'] = self.df.get('MonthlyCharges', 50)
        
        # Tenure groups
        if 'tenure' in self.df.columns:
            try:
                max_tenure = self.df['tenure'].max()
                if max_tenure <= 12:
                    bins = [0, 3, 6, 9, 12, float('inf')]
                    labels = ['0-3mo', '3-6mo', '6-9mo', '9-12mo', '12+mo']
                elif max_tenure <= 60:
                    bins = [0, 6, 12, 24, 36, float('inf')]
                    labels = ['0-6mo', '6-12mo', '1-2yr', '2-3yr', '3+yr']
                else:
                    bins = [0, 12, 24, 48, 72, float('inf')]
                    labels = ['0-1yr', '1-2yr', '2-4yr', '4-6yr', '6+yr']
                
                self.df['TenureGroup'] = pd.cut(
                    self.df['tenure'],
                    bins=bins,
                    labels=labels,
                    include_lowest=True
                )
            except Exception as e:
                logger.log(f"Tenure grouping failed: {str(e)}", "WARN")
                self.df['TenureGroup'] = 'Unknown'
        else:
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
        
        # Spending consistency
        if all(c in self.df.columns for c in ['MonthlyCharges', 'tenure', 'TotalCharges']):
            self.df['ExpectedTotal'] = self.df['MonthlyCharges'] * self.df['tenure']
            self.df['SpendingVariance'] = (self.df['TotalCharges'] - self.df['ExpectedTotal']).abs()
            
            max_variance = self.df['SpendingVariance'].quantile(0.99)
            if max_variance > 0:
                self.df['SpendingConsistency'] = 1 - (self.df['SpendingVariance'] / max_variance).clip(0, 1)
            else:
                self.df['SpendingConsistency'] = 1
        
        # Value score
        if 'TotalCharges' in self.df.columns:
            max_total = self.df['TotalCharges'].quantile(0.99)
            if max_total > 0:
                self.df['ValueScore'] = (self.df['TotalCharges'] / max_total * 100).clip(0, 100)
            else:
                self.df['ValueScore'] = 50
        else:
            self.df['ValueScore'] = 50
        
        logger.log("Feature engineering complete")


# ==============================
# MODEL BUILDER
# ==============================
class ModelBuilder:
    """Build predictive models with comprehensive fallback logic."""
    
    def __init__(self, df, confidence_scores=None):
        self.df = df.copy()
        self.accuracy = None
        self.auc = None
        self.model_note = ""
        self.confidence_scores = confidence_scores or {}
    
    def build(self):
        """Build churn and CLV models."""
        self._build_churn_model()
        self._build_clv_model()
        self._create_segments()
        self._add_recommendations()
        
        return self.df, self.accuracy, self.model_note
    
    def _build_churn_model(self):
        """Build churn prediction model with multiple fallbacks."""
        features = ['tenure', 'MonthlyCharges', 'TotalCharges', 'AvgMonthlySpend', 'ChargeRatio']
        features = [f for f in features if f in self.df.columns]
        
        if len(features) < 2:
            self._heuristic_churn_scoring()
            self.model_note = "Heuristic scoring (insufficient features)"
            return
        
        X = self.df[features].copy()
        y = self.df['Churn'].copy()
        
        # Check data quality for ML
        churn_rate = y.mean()
        unique_classes = y.nunique()
        sample_size = len(self.df)
        
        # Log churn statistics
        logger.log(f"Churn rate: {churn_rate:.1%}, Classes: {unique_classes}, Samples: {sample_size}")
        
        # Determine if ML is feasible
        ml_feasible = (
            SKLEARN_AVAILABLE and
            unique_classes >= 2 and
            sample_size >= 20 and
            y.sum() >= 3 and
            (len(y) - y.sum()) >= 3
        )
        
        if not ml_feasible:
            self._heuristic_churn_scoring()
            reason = []
            if not SKLEARN_AVAILABLE:
                reason.append("sklearn not available")
            if unique_classes < 2:
                reason.append("single class")
            if sample_size < 20:
                reason.append("too few samples")
            if y.sum() < 3:
                reason.append("too few churned")
            if (len(y) - y.sum()) < 3:
                reason.append("too few retained")
            
            self.model_note = f"Heuristic scoring ({', '.join(reason)})"
            return
        
        try:
            # Use RobustScaler for outlier resistance
            scaler = RobustScaler()
            X_scaled = pd.DataFrame(scaler.fit_transform(X), columns=features)
            
            # Impute any remaining missing values
            X_scaled = X_scaled.fillna(X_scaled.median())
            
            # Determine stratification
            test_size = min(0.25, max(0.1, 10 / len(self.df)))
            
            min_class_count = min(y.sum(), len(y) - y.sum())
            can_stratify = min_class_count >= 2
            
            if can_stratify:
                X_train, X_test, y_train, y_test = train_test_split(
                    X_scaled, y, test_size=test_size, stratify=y, random_state=42
                )
            else:
                X_train, X_test, y_train, y_test = train_test_split(
                    X_scaled, y, test_size=test_size, random_state=42
                )
            
            # Train models
            models = [
                ('RF', RandomForestClassifier(
                    n_estimators=100, max_depth=6, min_samples_split=5,
                    min_samples_leaf=2, class_weight='balanced', random_state=42, n_jobs=-1
                )),
                ('GB', GradientBoostingClassifier(
                    n_estimators=80, max_depth=4, min_samples_split=5,
                    min_samples_leaf=2, random_state=42
                )),
                ('LR', LogisticRegression(
                    class_weight='balanced', max_iter=1000, random_state=42
                ))
            ]
            
            best_model = None
            best_score = 0
            best_name = ""
            
            for name, model in models:
                try:
                    model.fit(X_train, y_train)
                    
                    # Calculate AUC if possible
                    try:
                        y_prob = model.predict_proba(X_test)[:, 1]
                        auc = roc_auc_score(y_test, y_prob)
                    except:
                        auc = 0.5
                    
                    acc = accuracy_score(y_test, model.predict(X_test))
                    
                    # Use combined score
                    combined_score = (auc + acc) / 2
                    
                    logger.log(f"Model {name}: accuracy={acc:.3f}, AUC={auc:.3f}")
                    
                    if combined_score > best_score:
                        best_score = combined_score
                        best_model = model
                        best_name = name
                        self.accuracy = acc
                        self.auc = auc
                        
                except Exception as e:
                    logger.log(f"Model {name} failed: {str(e)}", "WARN")
                    continue
            
            if best_model is not None:
                # Predict probabilities
                X_full_scaled = scaler.transform(X)
                self.df['Churn_Prob'] = best_model.predict_proba(X_full_scaled)[:, 1]
                
                # Clip probabilities to reasonable range
                self.df['Churn_Prob'] = self.df['Churn_Prob'].clip(0.01, 0.99)
                
                self.model_note = f"ML model ({best_name}) - Accuracy: {self.accuracy:.1%}"
                if self.auc:
                    self.model_note += f", AUC: {self.auc:.2f}"
                
                logger.log(f"Best model: {best_name} with score {best_score:.3f}")
            else:
                self._heuristic_churn_scoring()
                self.model_note = "Heuristic scoring (all models failed)"
                
        except Exception as e:
            logger.log(f"Model building error: {str(e)}", "ERROR")
            self._heuristic_churn_scoring()
            self.model_note = f"Heuristic scoring (error: {str(e)[:50]})"
    
    def _heuristic_churn_scoring(self):
        """Calculate churn probability using domain heuristics."""
        scores = pd.Series(0.5, index=self.df.index)
        
        # Tenure factor (lower tenure = higher risk)
        if 'tenure' in self.df.columns:
            tenure = self.df['tenure'].clip(1, None)
            max_tenure = tenure.quantile(0.95)
            if max_tenure > 0:
                tenure_score = 1 - (tenure / max_tenure).clip(0, 1)
                scores += tenure_score * 0.25
        
        # Charge ratio factor (extreme values = higher risk)
        if 'ChargeRatio' in self.df.columns:
            ratio = self.df['ChargeRatio']
            # Very high or very low ratios indicate risk
            ratio_deviation = (ratio - 1).abs()
            ratio_score = (ratio_deviation / ratio_deviation.quantile(0.95)).clip(0, 1)
            scores += ratio_score * 0.15
        
        # Spending consistency factor
        if 'SpendingConsistency' in self.df.columns:
            consistency_score = 1 - self.df['SpendingConsistency']
            scores += consistency_score * 0.1
        
        # If we have actual churn data, use it to calibrate
        if 'Churn' in self.df.columns and self.df['Churn'].nunique() > 1:
            actual_rate = self.df['Churn'].mean()
            # Adjust scores to match actual rate distribution
            scores = scores.rank(pct=True)
            # Scale to have mean close to actual rate
            scores = scores * actual_rate * 2
        
        self.df['Churn_Prob'] = scores.clip(0.05, 0.95)
        logger.log("Using heuristic churn scoring")
    
    def _build_clv_model(self):
        """Build Customer Lifetime Value model."""
        features = ['tenure', 'MonthlyCharges', 'AvgMonthlySpend']
        features = [f for f in features if f in self.df.columns]
        
        if len(features) < 2 or not SKLEARN_AVAILABLE:
            self._heuristic_clv()
            return
        
        try:
            X = self.df[features].fillna(self.df[features].median())
            y = self.df['TotalCharges']
            
            model = Ridge(alpha=1.0)
            model.fit(X, y)
            
            self.df['Predicted_CLV'] = model.predict(X)
            self.df['Predicted_CLV'] = self.df['Predicted_CLV'].clip(lower=0)
            
            # Adjust for churn risk
            retention_factor = 1 - (self.df['Churn_Prob'] * 0.5)
            self.df['Predicted_CLV'] = self.df['Predicted_CLV'] * retention_factor
            
            # Project future value (assume 24 months forward)
            if 'MonthlyCharges' in self.df.columns:
                future_months = 24
                self.df['Future_CLV'] = self.df['MonthlyCharges'] * future_months * retention_factor
            
            logger.log("CLV model built successfully")
            
        except Exception as e:
            logger.log(f"CLV model error: {str(e)}", "WARN")
            self._heuristic_clv()
    
    def _heuristic_clv(self):
        """Calculate CLV using heuristics."""
        base_months = 24  # Assume 24 month projection
        
        if 'MonthlyCharges' in self.df.columns:
            monthly = self.df['MonthlyCharges']
        else:
            monthly = 50
        
        retention_factor = 1 - (self.df['Churn_Prob'] * 0.5)
        
        self.df['Predicted_CLV'] = monthly * base_months * retention_factor
        self.df['Future_CLV'] = self.df['Predicted_CLV']
        
        logger.log("Using heuristic CLV calculation")
    
    def _create_segments(self):
        """Create customer segments using adaptive thresholds."""
        # Use percentile-based thresholds for robustness
        churn_threshold = self.df['Churn_Prob'].quantile(0.7)
        clv_threshold = self.df['Predicted_CLV'].median()
        
        def segment(row):
            high_risk = row['Churn_Prob'] > churn_threshold
            high_value = row['Predicted_CLV'] > clv_threshold
            
            if high_risk and high_value:
                return "High Risk - High Value"
            elif high_risk:
                return "High Risk - Low Value"
            elif high_value:
                return "Low Risk - High Value"
            else:
                return "Low Risk - Low Value"
        
        self.df['Segment'] = self.df.apply(segment, axis=1)
        
        # Priority score (higher = needs more attention)
        churn_rank = self.df['Churn_Prob'].rank(pct=True)
        value_rank = self.df['Predicted_CLV'].rank(pct=True)
        
        # Priority = high risk AND high value customers
        self.df['Priority'] = (churn_rank * 0.6 + value_rank * 0.4) * 100
        
        logger.log("Customer segmentation complete")
    
    def _add_recommendations(self):
        """Add action recommendations for each customer."""
        def get_recommendation(row):
            segment = row['Segment']
            churn_prob = row['Churn_Prob']
            
            if segment == "High Risk - High Value":
                if churn_prob > 0.8:
                    return "URGENT: Immediate retention outreach required"
                else:
                    return "Priority retention program - offer incentives"
            elif segment == "High Risk - Low Value":
                if churn_prob > 0.8:
                    return "Evaluate cost of retention vs acquisition"
                else:
                    return "Standard retention communication"
            elif segment == "Low Risk - High Value":
                return "Loyalty program - upsell opportunities"
            else:
                return "Monitor - standard engagement"
        
        self.df['Recommendation'] = self.df.apply(get_recommendation, axis=1)


# ==============================
# REPORT GENERATOR
# ==============================
class ReportGenerator:
    """Generate comprehensive Excel report with dashboard."""
    
    def __init__(self, df, accuracy, model_note, processing_notes, quality_report):
        self.df = df
        self.accuracy = accuracy
        self.model_note = model_note
        self.processing_notes = processing_notes
        self.quality_report = quality_report
    
    def generate(self, save_path):
        """Generate the full report."""
        distributions = self._create_distributions()
        self._write_excel(save_path, distributions)
        logger.log(f"Report saved to: {save_path}")
    
    def _create_distributions(self):
        """Create distribution dataframes."""
        distributions = {}
        
        # Segment distribution
        if 'Segment' in self.df.columns:
            seg = self.df['Segment'].value_counts().reset_index()
            seg.columns = ['Segment', 'Count']
            seg['Percentage'] = (seg['Count'] / seg['Count'].sum() * 100).round(1)
            distributions['Segments'] = seg
        
        # Churn probability distribution
        if 'Churn_Prob' in self.df.columns:
            try:
                churn_bins = pd.cut(
                    self.df['Churn_Prob'],
                    bins=[0, 0.2, 0.4, 0.6, 0.8, 1.0],
                    labels=['0-20%', '20-40%', '40-60%', '60-80%', '80-100%'],
                    include_lowest=True
                )
                churn_dist = churn_bins.value_counts().sort_index().reset_index()
                churn_dist.columns = ['Risk Level', 'Count']
                distributions['ChurnDist'] = churn_dist
            except:
                pass
        
        # CLV distribution
        if 'Predicted_CLV' in self.df.columns:
            try:
                clv_labels = ['Very Low', 'Low', 'Medium', 'High', 'Very High']
                clv_bins = pd.qcut(
                    self.df['Predicted_CLV'],
                    q=5,
                    labels=clv_labels,
                    duplicates='drop'
                )
                clv_dist = clv_bins.value_counts().sort_index().reset_index()
                clv_dist.columns = ['CLV Tier', 'Count']
                distributions['CLVDist'] = clv_dist
            except:
                pass
        
        # Spend distribution
        if 'AvgMonthlySpend' in self.df.columns:
            try:
                spend_labels = ['Very Low', 'Low', 'Medium', 'High', 'Very High']
                spend_bins = pd.qcut(
                    self.df['AvgMonthlySpend'],
                    q=5,
                    labels=spend_labels,
                    duplicates='drop'
                )
                spend_dist = spend_bins.value_counts().sort_index().reset_index()
                spend_dist.columns = ['Spend Tier', 'Count']
                distributions['SpendDist'] = spend_dist
            except:
                pass
        
        # Tenure distribution
        if 'TenureGroup' in self.df.columns:
            tenure_dist = self.df['TenureGroup'].value_counts().reset_index()
            tenure_dist.columns = ['Tenure Group', 'Count']
            distributions['TenureDist'] = tenure_dist
        
        return distributions
    
    def _write_excel(self, save_path, distributions):
        """Write the Excel report."""
        with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            formats = self._create_formats(workbook)
            
            self._write_dashboard(writer, distributions, formats)
            self._write_summary_sheet(writer, formats)
            self._write_data_sheet(writer, formats)
            self._write_distribution_sheets(writer, distributions, formats)
            self._write_quality_sheet(writer, formats)
    
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
            'subtitle': workbook.add_format({
                'italic': True, 'font_size': 10, 'font_color': '#666666'
            }),
            'kpi_blue': workbook.add_format({
                'bold': True, 'font_size': 12, 'font_color': 'white', 'bg_color': '#1E90FF',
                'align': 'center', 'valign': 'vcenter', 'border': 2
            }),
            'kpi_green': workbook.add_format({
                'bold': True, 'font_size': 12, 'font_color': 'white', 'bg_color': '#32CD32',
                'align': 'center', 'valign': 'vcenter', 'border': 2
            }),
            'kpi_red': workbook.add_format({
                'bold': True, 'font_size': 12, 'font_color': 'white', 'bg_color': '#FF4500',
                'align': 'center', 'valign': 'vcenter', 'border': 2
            }),
            'kpi_purple': workbook.add_format({
                'bold': True, 'font_size': 12, 'font_color': 'white', 'bg_color': '#9370DB',
                'align': 'center', 'valign': 'vcenter', 'border': 2
            }),
            'kpi_orange': workbook.add_format({
                'bold': True, 'font_size': 12, 'font_color': 'white', 'bg_color': '#FF8C00',
                'align': 'center', 'valign': 'vcenter', 'border': 2
            }),
            'kpi_value': workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 2
            }),
            'kpi_value_currency': workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
                'border': 2, 'num_format': '$#,##0'
            }),
            'kpi_value_percent': workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
                'border': 2, 'num_format': '0.0%'
            }),
            'warning': workbook.add_format({
                'font_color': '#FF6600', 'italic': True
            }),
            'note': workbook.add_format({
                'font_color': '#666666', 'text_wrap': True
            })
        }
    
    def _write_dashboard(self, writer, distributions, formats):
        """Write dashboard with charts and KPIs."""
        workbook = writer.book
        dashboard = workbook.add_worksheet('Dashboard')
        
        # Column sizing
        dashboard.set_column('A:A', 30)
        dashboard.set_column('B:B', 20)
        dashboard.set_column('C:C', 5)
        for col in range(3, 25):
            dashboard.set_column(col, col, 12)
        
        for row in range(3, 12):
            dashboard.set_row(row, 40)
        
        # Title
        dashboard.merge_range('A1:J1', '📊 Customer Analytics Dashboard', formats['title'])
        dashboard.set_row(0, 45)
        
        # Subtitle
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
        dashboard.write('A2', f"Generated: {timestamp} | Model: {self.model_note}", formats['subtitle'])
        
        # Data Quality Indicator
        quality_score = self.quality_report.get('quality_score', 0)
        quality_color = '#32CD32' if quality_score >= 70 else '#FF8C00' if quality_score >= 50 else '#FF4500'
        quality_format = workbook.add_format({
            'bold': True, 'font_size': 10, 'font_color': quality_color
        })
        dashboard.write('H2', f"Data Quality: {quality_score}/100", quality_format)
        
        # KPI Cards
        total_customers = len(self.df)
        avg_clv = self.df['Predicted_CLV'].mean() if 'Predicted_CLV' in self.df.columns else 0
        churn_rate = self.df['Churn'].mean() if 'Churn' in self.df.columns else 0
        high_risk = (self.df['Churn_Prob'] > 0.7).sum() if 'Churn_Prob' in self.df.columns else 0
        avg_tenure = self.df['tenure'].mean() if 'tenure' in self.df.columns else 0
        avg_monthly = self.df['MonthlyCharges'].mean() if 'MonthlyCharges' in self.df.columns else 0
        total_revenue = self.df['TotalCharges'].sum() if 'TotalCharges' in self.df.columns else 0
        at_risk_revenue = self.df[self.df['Churn_Prob'] > 0.7]['Predicted_CLV'].sum() if 'Churn_Prob' in self.df.columns else 0
        
        # Row 1 KPIs
        kpis = [
            ('Total Customers', total_customers, formats['kpi_blue'], formats['kpi_value']),
            ('Average CLV', avg_clv, formats['kpi_green'], formats['kpi_value_currency']),
            ('Churn Rate', churn_rate, formats['kpi_red'], formats['kpi_value_percent']),
            ('High Risk Count', high_risk, formats['kpi_purple'], formats['kpi_value']),
            ('Avg Tenure (mo)', round(avg_tenure, 1), formats['kpi_blue'], formats['kpi_value']),
            ('Avg Monthly', avg_monthly, formats['kpi_green'], formats['kpi_value_currency']),
            ('Total Revenue', total_revenue, formats['kpi_orange'], formats['kpi_value_currency']),
            ('At-Risk Revenue', at_risk_revenue, formats['kpi_red'], formats['kpi_value_currency']),
        ]
        
        row = 3
        for i, (label, value, label_fmt, value_fmt) in enumerate(kpis):
            dashboard.write(row + (i // 2) * 2, 0 if i % 2 == 0 else 1, label, label_fmt)
            dashboard.write(row + (i // 2) * 2 + 1, 0 if i % 2 == 0 else 1, value, value_fmt)
        
        # Charts
        chart_row = 3
        chart_col = 3
        
        # Segment Chart
        if 'Segments' in distributions:
            seg_data = distributions['Segments']
            if len(seg_data) > 0:
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'categories': ['Segments', 1, 0, len(seg_data), 0],
                    'values': ['Segments', 1, 1, len(seg_data), 1],
                    'fill': {'color': '#2E8B57'},
                    'data_labels': {'value': True, 'font': {'bold': True, 'size': 9}}
                })
                chart.set_title({'name': 'Customer Segments', 'name_font': {'size': 11, 'bold': True}})
                chart.set_legend({'none': True})
                chart.set_style(10)
                dashboard.insert_chart(chart_row, chart_col, chart, {'x_scale': 1.2, 'y_scale': 1.1})
        
        # Churn Distribution
        if 'ChurnDist' in distributions:
            churn_data = distributions['ChurnDist']
            if len(churn_data) > 0:
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'categories': ['ChurnDist', 1, 0, len(churn_data), 0],
                    'values': ['ChurnDist', 1, 1, len(churn_data), 1],
                    'fill': {'color': '#FF6347'},
                    'data_labels': {'value': True, 'font': {'bold': True, 'size': 9}}
                })
                chart.set_title({'name': 'Churn Risk Distribution', 'name_font': {'size': 11, 'bold': True}})
                chart.set_legend({'none': True})
                chart.set_style(10)
                dashboard.insert_chart(chart_row + 15, chart_col, chart, {'x_scale': 1.2, 'y_scale': 1.1})
        
        # CLV Distribution
        if 'CLVDist' in distributions:
            clv_data = distributions['CLVDist']
            if len(clv_data) > 0:
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'categories': ['CLVDist', 1, 0, len(clv_data), 0],
                    'values': ['CLVDist', 1, 1, len(clv_data), 1],
                    'fill': {'color': '#1E90FF'},
                    'data_labels': {'value': True, 'font': {'bold': True, 'size': 9}}
                })
                chart.set_title({'name': 'Customer Lifetime Value', 'name_font': {'size': 11, 'bold': True}})
                chart.set_legend({'none': True})
                chart.set_style(10)
                dashboard.insert_chart(chart_row, chart_col + 9, chart, {'x_scale': 1.2, 'y_scale': 1.1})
        
        # Tenure Distribution (Pie)
        if 'TenureDist' in distributions:
            tenure_data = distributions['TenureDist']
            if len(tenure_data) > 1:
                chart = workbook.add_chart({'type': 'pie'})
                chart.add_series({
                    'categories': ['TenureDist', 1, 0, len(tenure_data), 0],
                    'values': ['TenureDist', 1, 1, len(tenure_data), 1],
                    'data_labels': {'percentage': True, 'category': True, 'font': {'size': 8}}
                })
                chart.set_title({'name': 'Tenure Distribution', 'name_font': {'size': 11, 'bold': True}})
                chart.set_style(10)
                dashboard.insert_chart(chart_row + 15, chart_col + 9, chart, {'x_scale': 1.1, 'y_scale': 1.1})
    
    def _write_summary_sheet(self, writer, formats):
        """Write summary sheet."""
        ws = writer.book.add_worksheet('Summary')
        
        # KPIs
        kpis = {
            'Total Customers': len(self.df),
            'Average CLV': self.df['Predicted_CLV'].mean() if 'Predicted_CLV' in self.df.columns else 'N/A',
            'Median CLV': self.df['Predicted_CLV'].median() if 'Predicted_CLV' in self.df.columns else 'N/A',
            'Total Portfolio Value': self.df['Predicted_CLV'].sum() if 'Predicted_CLV' in self.df.columns else 'N/A',
            'Actual Churn Rate': self.df['Churn'].mean() if 'Churn' in self.df.columns else 'N/A',
            'Predicted High Risk (>70%)': (self.df['Churn_Prob'] > 0.7).sum() if 'Churn_Prob' in self.df.columns else 'N/A',
            'Average Tenure (months)': self.df['tenure'].mean() if 'tenure' in self.df.columns else 'N/A',
            'Average Monthly Charges': self.df['MonthlyCharges'].mean() if 'MonthlyCharges' in self.df.columns else 'N/A',
            'Model Accuracy': f"{self.accuracy:.1%}" if self.accuracy else 'Heuristic',
            'Model Type': self.model_note,
            'Data Quality Score': self.quality_report.get('quality_score', 'N/A')
        }
        
        # Write header
        ws.write(0, 0, 'Metric', formats['header'])
        ws.write(0, 1, 'Value', formats['header'])
        
        row = 1
        for metric, value in kpis.items():
            ws.write(row, 0, metric, formats['cell'])
            
            if isinstance(value, float):
                if 'CLV' in metric or 'Charges' in metric or 'Value' in metric:
                    ws.write(row, 1, value, formats['currency'])
                elif 'Rate' in metric:
                    ws.write(row, 1, value, formats['percent'])
                else:
                    ws.write(row, 1, value, formats['number'])
            else:
                ws.write(row, 1, str(value), formats['cell'])
            row += 1
        
        ws.set_column('A:A', 35)
        ws.set_column('B:B', 45)
        
        # Processing notes
        if self.processing_notes:
            row += 2
            ws.write(row, 0, 'Processing Notes:', formats['header'])
            row += 1
            for note in self.processing_notes:
                ws.write(row, 0, f"• {note}", formats['note'])
                row += 1
    
    def _write_data_sheet(self, writer, formats):
        """Write main data sheet."""
        # Select columns to export
        exclude_cols = ['ExpectedTotal', 'ChargeRatio', 'SpendingVariance', '_is_outlier']
        export_cols = [c for c in self.df.columns if c not in exclude_cols]
        df_export = self.df[export_cols].copy()
        
        # Convert categorical
        for col in df_export.columns:
            if df_export[col].dtype.name == 'category':
                df_export[col] = df_export[col].astype(str)
        
        df_export.to_excel(writer, sheet_name='Data', index=False)
        ws = writer.sheets['Data']
        
        # Format
        for i, col in enumerate(df_export.columns):
            ws.write(0, i, col, formats['header'])
            max_len = min(max(df_export[col].astype(str).str.len().max(), len(col)) + 2, 25)
            ws.set_column(i, i, max_len, formats['cell'])
        
        ws.autofilter(0, 0, len(df_export), len(df_export.columns) - 1)
        ws.freeze_panes(1, 0)
    
    def _write_distribution_sheets(self, writer, distributions, formats):
        """Write distribution sheets."""
        for name, data in distributions.items():
            if data is not None and not data.empty:
                for col in data.columns:
                    if data[col].dtype.name == 'category' or 'interval' in str(data[col].dtype).lower():
                        data[col] = data[col].astype(str)
                
                data.to_excel(writer, sheet_name=name, index=False)
                ws = writer.sheets[name]
                
                for i, col in enumerate(data.columns):
                    ws.write(0, i, col, formats['header'])
                    ws.set_column(i, i, 18, formats['cell'])
                
                ws.autofilter(0, 0, len(data), len(data.columns) - 1)
    
    def _write_quality_sheet(self, writer, formats):
        """Write data quality report sheet."""
        ws = writer.book.add_worksheet('DataQuality')
        
        ws.write(0, 0, 'Data Quality Report', formats['title'])
        ws.set_row(0, 30)
        
        row = 2
        
        # Quality score
        score = self.quality_report.get('quality_score', 0)
        ws.write(row, 0, 'Quality Score:', formats['header'])
        ws.write(row, 1, f"{score}/100", formats['kpi_value'])
        row += 2
        
        # Domain
        domain = self.quality_report.get('domain', 'Unknown')
        ws.write(row, 0, 'Detected Domain:', formats['header'])
        ws.write(row, 1, domain.replace('_', ' ').title(), formats['cell'])
        row += 2
        
        # Issues
        issues = self.quality_report.get('issues', [])
        if issues:
            ws.write(row, 0, 'Issues Found:', formats['header'])
            row += 1
            for issue in issues:
                ws.write(row, 0, f"• {issue}", formats['warning'] if 'WARNING' in issue or 'CRITICAL' in issue else formats['note'])
                row += 1
        row += 1
        
        # Recommendations
        recommendations = self.quality_report.get('recommendations', [])
        if recommendations:
            ws.write(row, 0, 'Recommendations:', formats['header'])
            row += 1
            for rec in recommendations:
                ws.write(row, 0, f"• {rec}", formats['note'])
                row += 1
        
        ws.set_column('A:A', 80)
        ws.set_column('B:B', 20)


# ==============================
# PROGRESS DIALOG
# ==============================
class ProgressDialog:
    """Progress dialog with detailed status."""
    
    def __init__(self, parent, title="Processing"):
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("450x180")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()
        
        # Center
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() - 450) // 2
        y = (self.window.winfo_screenheight() - 180) // 2
        self.window.geometry(f"+{x}+{y}")
        
        # Status
        self.status_var = tk.StringVar(value="Initializing...")
        self.status_label = ttk.Label(self.window, textvariable=self.status_var, font=('Segoe UI', 11))
        self.status_label.pack(pady=15)
        
        # Progress
        self.progress = ttk.Progressbar(self.window, mode='indeterminate', length=400)
        self.progress.pack(pady=10)
        self.progress.start(10)
        
        # Detail
        self.detail_var = tk.StringVar(value="")
        self.detail_label = ttk.Label(self.window, textvariable=self.detail_var, font=('Segoe UI', 9), foreground='gray')
        self.detail_label.pack(pady=5)
        
        # Sub-detail
        self.subdetail_var = tk.StringVar(value="")
        self.subdetail_label = ttk.Label(self.window, textvariable=self.subdetail_var, font=('Segoe UI', 8), foreground='#999999')
        self.subdetail_label.pack(pady=2)
    
    def update(self, status, detail="", subdetail=""):
        self.status_var.set(status)
        self.detail_var.set(detail)
        self.subdetail_var.set(subdetail)
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
        
        # DPI awareness
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
    
    def run(self):
        """Run the application."""
        try:
            # Welcome
            result = messagebox.askyesno(
                f"{Config.APP_NAME} v{Config.VERSION}",
                "Welcome to Customer Analytics Tool v3.0!\n\n"
                "This robust tool will:\n"
                "• Automatically detect and adapt to your data format\n"
                "• Handle outliers and data quality issues\n"
                "• Intelligently map columns to required fields\n"
                "• Predict churn probability and customer lifetime value\n"
                "• Create actionable customer segments\n"
                "• Generate a comprehensive visual dashboard\n\n"
                "Supported formats: CSV, Excel, JSON, Parquet, TSV\n\n"
                "Continue?",
                icon='info'
            )
            
            if not result:
                return
            
            # File selection
            file_path = self._select_input_file()
            if not file_path:
                return
            
            # Process
            result = self._process_data(file_path)
            if result is None:
                return
            
            df, accuracy, model_note, processing_notes, quality_report = result
            
            # Save location
            save_path = self._select_save_location(file_path)
            if not save_path:
                return
            
            # Generate report
            self._generate_report(df, save_path, accuracy, model_note, processing_notes, quality_report)
            
            # Success
            self._show_success(df, save_path, accuracy, model_note, quality_report)
            
        except Exception as e:
            self._show_error(e)
        finally:
            self.root.destroy()
    
    def _select_input_file(self):
        """Select input file."""
        filetypes = [
            ("All Supported", "*.csv;*.xlsx;*.xls;*.xlsm;*.xlsb;*.json;*.parquet;*.pq;*.tsv;*.txt"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx;*.xls;*.xlsm;*.xlsb"),
            ("JSON files", "*.json"),
            ("Parquet files", "*.parquet;*.pq"),
            ("TSV files", "*.tsv"),
            ("Text files", "*.txt"),
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
        progress = ProgressDialog(self.root, "Analyzing Data")
        
        try:
            # Load
            progress.update("Loading file...", Path(file_path).name)
            df = FileLoader.load(file_path)
            
            # Quality check
            progress.update("Analyzing data quality...", f"{len(df):,} rows loaded")
            analyzer = DataQualityAnalyzer(df)
            quality_report = analyzer.analyze()
            
            # Check if data is usable
            if not quality_report['is_valid']:
                progress.close()
                issues = '\n'.join(f"• {i}" for i in quality_report['issues'][:5])
                messagebox.showerror(
                    "Data Quality Issue",
                    f"Cannot process this dataset:\n\n{issues}\n\n"
                    f"Quality Score: {quality_report['quality_score']}/100\n\n"
                    "Please provide a dataset with customer-related data."
                )
                return None
            
            # Warn about quality issues
            if quality_report['quality_score'] < 70:
                progress.close()
                issues = '\n'.join(f"• {i}" for i in quality_report['issues'][:5])
                recommendations = '\n'.join(f"• {r}" for r in quality_report['recommendations'][:3])
                
                proceed = messagebox.askyesno(
                    "Data Quality Warning",
                    f"Data quality is below optimal:\n\n"
                    f"Quality Score: {quality_report['quality_score']}/100\n\n"
                    f"Issues:\n{issues}\n\n"
                    f"Recommendations:\n{recommendations}\n\n"
                    f"Continue anyway? (Results may be less accurate)",
                    icon='warning'
                )
                
                if not proceed:
                    return None
                
                progress = ProgressDialog(self.root, "Analyzing Data")
            
            # Large file warning
            if len(df) > Config.MAX_ROWS_WARN:
                progress.close()
                proceed = messagebox.askyesno(
                    "Large Dataset",
                    f"This file has {len(df):,} rows.\n"
                    "Processing may take a while.\n\nContinue?",
                    icon='warning'
                )
                if not proceed:
                    return None
                progress = ProgressDialog(self.root, "Analyzing Data")
            
            # Process
            progress.update("Processing data...", "Cleaning and mapping columns")
            processor = DataProcessor(df, quality_report)
            df, processing_notes = processor.process()
            
            # Build models
            progress.update("Building models...", "Training prediction algorithms", "This may take a moment...")
            builder = ModelBuilder(df)
            df, accuracy, model_note = builder.build()
            
            progress.close()
            
            # Summary
            messagebox.showinfo(
                "Processing Complete",
                f"✅ Data processed successfully!\n\n"
                f"Rows: {len(df):,}\n"
                f"Columns: {len(df.columns)}\n"
                f"Segments: {df['Segment'].nunique()}\n"
                f"Quality Score: {quality_report['quality_score']}/100\n\n"
                f"Model: {model_note}\n\n"
                "Click OK to choose save location."
            )
            
            return df, accuracy, model_note, processing_notes, quality_report
            
        except Exception as e:
            progress.close()
            raise
    
    def _select_save_location(self, input_path):
        """Select save location."""
        default_name = Path(input_path).stem + "_analysis.xlsx"
        
        save_path = filedialog.asksaveasfilename(
            title="Save Analysis Report",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not save_path:
            messagebox.showinfo("Cancelled", "Save cancelled.")
            return None
        
        return save_path
    
    def _generate_report(self, df, save_path, accuracy, model_note, processing_notes, quality_report):
        """Generate report."""
        progress = ProgressDialog(self.root, "Generating Report")
        
        try:
            progress.update("Creating report...", "Building sheets and charts")
            generator = ReportGenerator(df, accuracy, model_note, processing_notes, quality_report)
            generator.generate(save_path)
            progress.close()
        except Exception as e:
            progress.close()
            raise
    
    def _show_success(self, df, save_path, accuracy, model_note, quality_report):
        """Show success message."""
        accuracy_str = f"{accuracy:.1%}" if accuracy else "Heuristic"
        
        high_risk = (df['Churn_Prob'] > 0.7).sum() if 'Churn_Prob' in df.columns else 0
        high_value = (df['Predicted_CLV'] > df['Predicted_CLV'].median()).sum() if 'Predicted_CLV' in df.columns else 0
        
        messagebox.showinfo(
            "Success! 🎉",
            f"Analysis Complete!\n\n"
            f"📊 Customers: {len(df):,}\n"
            f"🎯 Model: {accuracy_str}\n"
            f"⚠️ High Risk: {high_risk:,}\n"
            f"💎 High Value: {high_value:,}\n"
            f"📈 Quality: {quality_report['quality_score']}/100\n\n"
            f"📁 Saved to:\n{save_path}\n\n"
            "Sheets: Dashboard, Summary, Data, Segments,\n"
            "ChurnDist, CLVDist, SpendDist, TenureDist, DataQuality"
        )
    
    def _show_error(self, error):
        """Show error with helpful context."""
        error_msg = str(error)
        
        hints = []
        if 'encoding' in error_msg.lower():
            hints.append("Try saving with UTF-8 encoding")
        if 'permission' in error_msg.lower():
            hints.append("Close the file in other programs")
        if 'memory' in error_msg.lower():
            hints.append("Try a smaller dataset")
        if 'empty' in error_msg.lower():
            hints.append("Ensure file contains data")
        if 'column' in error_msg.lower():
            hints.append("Check column names and data format")
        
        hint_text = "\n\nSuggestions:\n• " + "\n• ".join(hints) if hints else ""
        
        logger.log(f"Error: {error_msg}", "ERROR")
        logger.log(traceback.format_exc(), "DEBUG")
        
        messagebox.showerror(
            "Error",
            f"⚠️ An error occurred:\n\n{error_msg[:200]}{hint_text}\n\n"
            f"Ensure your file:\n"
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

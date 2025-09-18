"""
Data processing pipeline for Excel-Ollama AI Plugin.
Handles data validation, cleaning, preprocessing, and format conversion.
"""

import pandas as pd
import numpy as np
import json
import xml.etree.ElementTree as ET
from typing import Dict, Any, List, Optional, Iterator, Union, Tuple
from dataclasses import dataclass
from enum import Enum
import io
import re
from pathlib import Path

from .interfaces import IDataProcessor, ValidationResult
from ..utils.config import config_manager


class DataType(Enum):
    """Supported data types for automatic detection."""
    NUMERIC = "numeric"
    CATEGORICAL = "categorical"
    DATETIME = "datetime"
    TEXT = "text"
    BOOLEAN = "boolean"
    UNKNOWN = "unknown"


class CleaningStrategy(Enum):
    """Strategies for handling missing values."""
    DROP = "drop"
    FILL_MEAN = "fill_mean"
    FILL_MEDIAN = "fill_median"
    FILL_MODE = "fill_mode"
    FILL_FORWARD = "fill_forward"
    FILL_BACKWARD = "fill_backward"
    FILL_ZERO = "fill_zero"
    INTERPOLATE = "interpolate"


@dataclass
class DataQualityMetrics:
    """Metrics for assessing data quality."""
    total_rows: int
    total_columns: int
    missing_values: int
    duplicate_rows: int
    outliers: int
    data_types: Dict[str, str]
    completeness_score: float
    consistency_score: float
    validity_score: float


@dataclass
class CleaningRules:
    """Rules for data cleaning operations."""
    missing_value_strategy: CleaningStrategy = CleaningStrategy.DROP
    remove_duplicates: bool = True
    handle_outliers: bool = True
    outlier_method: str = "iqr"  # iqr, zscore, isolation_forest
    outlier_threshold: float = 3.0
    normalize_text: bool = True
    convert_data_types: bool = True
    custom_rules: Dict[str, Any] = None
    
    def __post_init__(self):
        if self.custom_rules is None:
            self.custom_rules = {}


class DataProcessor(IDataProcessor):
    """Main data processing class."""
    
    def __init__(self):
        """Initialize data processor."""
        config = config_manager.get_config()
        self.chunk_size = config.excel_settings.max_rows_per_chunk
        self.validators = self._initialize_validators()
        self.cleaners = self._initialize_cleaners()
        self.type_detectors = self._initialize_type_detectors()
    
    def _initialize_validators(self) -> Dict[str, callable]:
        """Initialize data validation functions."""
        return {
            'numeric': self._validate_numeric,
            'datetime': self._validate_datetime,
            'categorical': self._validate_categorical,
            'text': self._validate_text,
            'boolean': self._validate_boolean
        }
    
    def _initialize_cleaners(self) -> Dict[CleaningStrategy, callable]:
        """Initialize data cleaning functions."""
        return {
            CleaningStrategy.DROP: self._clean_drop_missing,
            CleaningStrategy.FILL_MEAN: self._clean_fill_mean,
            CleaningStrategy.FILL_MEDIAN: self._clean_fill_median,
            CleaningStrategy.FILL_MODE: self._clean_fill_mode,
            CleaningStrategy.FILL_FORWARD: self._clean_fill_forward,
            CleaningStrategy.FILL_BACKWARD: self._clean_fill_backward,
            CleaningStrategy.FILL_ZERO: self._clean_fill_zero,
            CleaningStrategy.INTERPOLATE: self._clean_interpolate
        }
    
    def _initialize_type_detectors(self) -> Dict[DataType, callable]:
        """Initialize data type detection functions."""
        return {
            DataType.NUMERIC: self._detect_numeric,
            DataType.DATETIME: self._detect_datetime,
            DataType.BOOLEAN: self._detect_boolean,
            DataType.CATEGORICAL: self._detect_categorical,
            DataType.TEXT: self._detect_text
        }
    
    def validate_data(self, data: pd.DataFrame) -> ValidationResult:
        """Validate data quality and structure."""
        errors = []
        warnings = []
        
        # Basic structure validation
        if data.empty:
            errors.append("Dataset is empty")
            return ValidationResult(
                is_valid=False,
                errors=errors,
                warnings=warnings,
                data_types={},
                row_count=0,
                column_count=0
            )
        
        # Check for minimum data requirements
        if len(data) < 2:
            warnings.append("Dataset has fewer than 2 rows")
        
        if len(data.columns) < 1:
            errors.append("Dataset has no columns")
        
        # Detect data types
        detected_types = self.detect_data_types(data)
        
        # Validate each column based on detected type
        for column, data_type in detected_types.items():
            if column in data.columns:
                validator = self.validators.get(data_type)
                if validator:
                    column_errors = validator(data[column])
                    errors.extend([f"Column '{column}': {error}" for error in column_errors])
        
        # Check for excessive missing values
        missing_percentage = (data.isnull().sum().sum() / (len(data) * len(data.columns))) * 100
        if missing_percentage > 50:
            warnings.append(f"High percentage of missing values: {missing_percentage:.1f}%")
        elif missing_percentage > 20:
            warnings.append(f"Moderate percentage of missing values: {missing_percentage:.1f}%")
        
        # Check for duplicate rows
        duplicate_count = data.duplicated().sum()
        if duplicate_count > 0:
            warnings.append(f"Found {duplicate_count} duplicate rows")
        
        # Check column names
        invalid_columns = [col for col in data.columns if not isinstance(col, str) or not col.strip()]
        if invalid_columns:
            errors.append(f"Invalid column names: {invalid_columns}")
        
        return ValidationResult(
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=warnings,
            data_types=detected_types,
            row_count=len(data),
            column_count=len(data.columns)
        )
    
    def clean_data(self, data: pd.DataFrame, cleaning_rules: CleaningRules) -> pd.DataFrame:
        """Clean and preprocess data according to rules."""
        cleaned_data = data.copy()
        
        # Handle missing values
        if cleaning_rules.missing_value_strategy in self.cleaners:
            cleaner = self.cleaners[cleaning_rules.missing_value_strategy]
            cleaned_data = cleaner(cleaned_data)
        
        # Remove duplicates
        if cleaning_rules.remove_duplicates:
            initial_count = len(cleaned_data)
            cleaned_data = cleaned_data.drop_duplicates()
            removed_count = initial_count - len(cleaned_data)
            if removed_count > 0:
                print(f"Removed {removed_count} duplicate rows")
        
        # Handle outliers
        if cleaning_rules.handle_outliers:
            cleaned_data = self._handle_outliers(
                cleaned_data,
                method=cleaning_rules.outlier_method,
                threshold=cleaning_rules.outlier_threshold
            )
        
        # Normalize text columns
        if cleaning_rules.normalize_text:
            text_columns = self._get_text_columns(cleaned_data)
            for col in text_columns:
                cleaned_data[col] = self._normalize_text_column(cleaned_data[col])
        
        # Convert data types
        if cleaning_rules.convert_data_types:
            cleaned_data = self._convert_data_types(cleaned_data)
        
        # Apply custom rules
        if cleaning_rules.custom_rules:
            cleaned_data = self._apply_custom_rules(cleaned_data, cleaning_rules.custom_rules)
        
        return cleaned_data
    
    def chunk_large_dataset(self, data: pd.DataFrame) -> Iterator[pd.DataFrame]:
        """Split large dataset into manageable chunks."""
        if len(data) <= self.chunk_size:
            yield data
            return
        
        for start_idx in range(0, len(data), self.chunk_size):
            end_idx = min(start_idx + self.chunk_size, len(data))
            yield data.iloc[start_idx:end_idx].copy()
    
    def detect_data_types(self, data: pd.DataFrame) -> Dict[str, str]:
        """Detect data types for each column."""
        detected_types = {}
        
        for column in data.columns:
            series = data[column].dropna()
            
            if len(series) == 0:
                detected_types[column] = DataType.UNKNOWN.value
                continue
            
            # Try each detector in order of specificity
            for data_type, detector in self.type_detectors.items():
                if detector(series):
                    detected_types[column] = data_type.value
                    break
            else:
                detected_types[column] = DataType.UNKNOWN.value
        
        return detected_types
    
    def handle_missing_values(self, data: pd.DataFrame, strategy: str = 'auto') -> pd.DataFrame:
        """Handle missing values with specified strategy."""
        if strategy == 'auto':
            # Choose strategy based on data characteristics
            missing_percentage = (data.isnull().sum().sum() / (len(data) * len(data.columns))) * 100
            
            if missing_percentage < 5:
                strategy = CleaningStrategy.DROP
            elif missing_percentage < 20:
                strategy = CleaningStrategy.FILL_MEAN
            else:
                strategy = CleaningStrategy.INTERPOLATE
        else:
            strategy = CleaningStrategy(strategy)
        
        if strategy in self.cleaners:
            return self.cleaners[strategy](data)
        else:
            return data
    
    def get_data_quality_metrics(self, data: pd.DataFrame) -> DataQualityMetrics:
        """Calculate comprehensive data quality metrics."""
        total_cells = len(data) * len(data.columns)
        missing_values = data.isnull().sum().sum()
        duplicate_rows = data.duplicated().sum()
        
        # Detect outliers
        outliers = 0
        numeric_columns = data.select_dtypes(include=[np.number]).columns
        for col in numeric_columns:
            outliers += self._count_outliers(data[col])
        
        # Calculate quality scores
        completeness_score = 1.0 - (missing_values / total_cells) if total_cells > 0 else 0.0
        consistency_score = 1.0 - (duplicate_rows / len(data)) if len(data) > 0 else 0.0
        validity_score = self._calculate_validity_score(data)
        
        return DataQualityMetrics(
            total_rows=len(data),
            total_columns=len(data.columns),
            missing_values=missing_values,
            duplicate_rows=duplicate_rows,
            outliers=outliers,
            data_types=self.detect_data_types(data),
            completeness_score=completeness_score,
            consistency_score=consistency_score,
            validity_score=validity_score
        )
    
    # Data format converters
    def import_csv(self, file_path: str, **kwargs) -> pd.DataFrame:
        """Import data from CSV file."""
        try:
            return pd.read_csv(file_path, **kwargs)
        except Exception as e:
            raise ValueError(f"Error importing CSV: {e}")
    
    def import_json(self, file_path: str, **kwargs) -> pd.DataFrame:
        """Import data from JSON file."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if isinstance(data, list):
                return pd.DataFrame(data)
            elif isinstance(data, dict):
                return pd.DataFrame([data])
            else:
                raise ValueError("JSON data must be a list or dictionary")
        except Exception as e:
            raise ValueError(f"Error importing JSON: {e}")
    
    def import_xml(self, file_path: str, root_element: str = None) -> pd.DataFrame:
        """Import data from XML file."""
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            # If root_element is specified, find those elements
            if root_element:
                elements = root.findall(f".//{root_element}")
            else:
                # Use direct children of root
                elements = list(root)
            
            data = []
            for element in elements:
                row = {}
                for child in element:
                    row[child.tag] = child.text
                data.append(row)
            
            return pd.DataFrame(data)
        except Exception as e:
            raise ValueError(f"Error importing XML: {e}")
    
    # Private helper methods
    def _validate_numeric(self, series: pd.Series) -> List[str]:
        """Validate numeric column."""
        errors = []
        try:
            pd.to_numeric(series, errors='raise')
        except (ValueError, TypeError):
            errors.append("Contains non-numeric values")
        return errors
    
    def _validate_datetime(self, series: pd.Series) -> List[str]:
        """Validate datetime column."""
        errors = []
        try:
            pd.to_datetime(series, errors='raise')
        except (ValueError, TypeError):
            errors.append("Contains invalid datetime values")
        return errors
    
    def _validate_categorical(self, series: pd.Series) -> List[str]:
        """Validate categorical column."""
        errors = []
        unique_ratio = len(series.unique()) / len(series)
        if unique_ratio > 0.5:
            errors.append("High cardinality for categorical data")
        return errors
    
    def _validate_text(self, series: pd.Series) -> List[str]:
        """Validate text column."""
        errors = []
        if series.dtype != 'object':
            errors.append("Text column should have object dtype")
        return errors
    
    def _validate_boolean(self, series: pd.Series) -> List[str]:
        """Validate boolean column."""
        errors = []
        unique_values = set(series.dropna().astype(str).str.lower())
        valid_boolean = {'true', 'false', '1', '0', 'yes', 'no', 'y', 'n'}
        if not unique_values.issubset(valid_boolean):
            errors.append("Contains invalid boolean values")
        return errors
    
    def _detect_numeric(self, series: pd.Series) -> bool:
        """Detect if series is numeric."""
        try:
            pd.to_numeric(series, errors='raise')
            return True
        except (ValueError, TypeError):
            return False
    
    def _detect_datetime(self, series: pd.Series) -> bool:
        """Detect if series is datetime."""
        if series.dtype.name.startswith('datetime'):
            return True
        
        try:
            pd.to_datetime(series, errors='raise')
            return True
        except (ValueError, TypeError):
            return False
    
    def _detect_boolean(self, series: pd.Series) -> bool:
        """Detect if series is boolean."""
        if series.dtype == bool:
            return True
        
        unique_values = set(series.dropna().astype(str).str.lower())
        boolean_values = {'true', 'false', '1', '0', 'yes', 'no', 'y', 'n'}
        return len(unique_values) <= 2 and unique_values.issubset(boolean_values)
    
    def _detect_categorical(self, series: pd.Series) -> bool:
        """Detect if series is categorical."""
        if series.dtype.name == 'category':
            return True
        
        unique_ratio = len(series.unique()) / len(series)
        return unique_ratio < 0.1 and len(series.unique()) < 50
    
    def _detect_text(self, series: pd.Series) -> bool:
        """Detect if series is text."""
        return series.dtype == 'object' and not self._detect_datetime(series)
    
    # Cleaning methods
    def _clean_drop_missing(self, data: pd.DataFrame) -> pd.DataFrame:
        """Drop rows with missing values."""
        return data.dropna()
    
    def _clean_fill_mean(self, data: pd.DataFrame) -> pd.DataFrame:
        """Fill missing values with mean for numeric columns."""
        numeric_columns = data.select_dtypes(include=[np.number]).columns
        for col in numeric_columns:
            data[col].fillna(data[col].mean(), inplace=True)
        return data
    
    def _clean_fill_median(self, data: pd.DataFrame) -> pd.DataFrame:
        """Fill missing values with median for numeric columns."""
        numeric_columns = data.select_dtypes(include=[np.number]).columns
        for col in numeric_columns:
            data[col].fillna(data[col].median(), inplace=True)
        return data
    
    def _clean_fill_mode(self, data: pd.DataFrame) -> pd.DataFrame:
        """Fill missing values with mode."""
        for col in data.columns:
            if data[col].isnull().any():
                mode_value = data[col].mode()
                if not mode_value.empty:
                    data[col].fillna(mode_value.iloc[0], inplace=True)
        return data
    
    def _clean_fill_forward(self, data: pd.DataFrame) -> pd.DataFrame:
        """Forward fill missing values."""
        return data.fillna(method='ffill')
    
    def _clean_fill_backward(self, data: pd.DataFrame) -> pd.DataFrame:
        """Backward fill missing values."""
        return data.fillna(method='bfill')
    
    def _clean_fill_zero(self, data: pd.DataFrame) -> pd.DataFrame:
        """Fill missing values with zero."""
        return data.fillna(0)
    
    def _clean_interpolate(self, data: pd.DataFrame) -> pd.DataFrame:
        """Interpolate missing values."""
        numeric_columns = data.select_dtypes(include=[np.number]).columns
        for col in numeric_columns:
            data[col] = data[col].interpolate()
        return data
    
    def _handle_outliers(self, data: pd.DataFrame, method: str = 'iqr', threshold: float = 3.0) -> pd.DataFrame:
        """Handle outliers in numeric columns."""
        numeric_columns = data.select_dtypes(include=[np.number]).columns
        
        for col in numeric_columns:
            if method == 'iqr':
                Q1 = data[col].quantile(0.25)
                Q3 = data[col].quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                data[col] = data[col].clip(lower_bound, upper_bound)
            
            elif method == 'zscore':
                z_scores = np.abs((data[col] - data[col].mean()) / data[col].std())
                data[col] = data[col][z_scores < threshold]
        
        return data
    
    def _get_text_columns(self, data: pd.DataFrame) -> List[str]:
        """Get list of text columns."""
        return data.select_dtypes(include=['object']).columns.tolist()
    
    def _normalize_text_column(self, series: pd.Series) -> pd.Series:
        """Normalize text in a column."""
        return series.astype(str).str.strip().str.lower()
    
    def _convert_data_types(self, data: pd.DataFrame) -> pd.DataFrame:
        """Convert columns to appropriate data types."""
        detected_types = self.detect_data_types(data)
        
        for column, data_type in detected_types.items():
            try:
                if data_type == DataType.NUMERIC.value:
                    data[column] = pd.to_numeric(data[column], errors='coerce')
                elif data_type == DataType.DATETIME.value:
                    data[column] = pd.to_datetime(data[column], errors='coerce')
                elif data_type == DataType.BOOLEAN.value:
                    data[column] = data[column].astype(bool)
                elif data_type == DataType.CATEGORICAL.value:
                    data[column] = data[column].astype('category')
            except Exception as e:
                print(f"Warning: Could not convert column '{column}' to {data_type}: {e}")
        
        return data
    
    def _apply_custom_rules(self, data: pd.DataFrame, custom_rules: Dict[str, Any]) -> pd.DataFrame:
        """Apply custom cleaning rules."""
        # This can be extended based on specific requirements
        for rule_name, rule_config in custom_rules.items():
            if rule_name == 'replace_values':
                for column, replacements in rule_config.items():
                    if column in data.columns:
                        data[column] = data[column].replace(replacements)
        
        return data
    
    def _count_outliers(self, series: pd.Series) -> int:
        """Count outliers in a numeric series."""
        if series.dtype not in [np.number]:
            return 0
        
        Q1 = series.quantile(0.25)
        Q3 = series.quantile(0.75)
        IQR = Q3 - Q1
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR
        
        return ((series < lower_bound) | (series > upper_bound)).sum()
    
    def _calculate_validity_score(self, data: pd.DataFrame) -> float:
        """Calculate validity score based on data type consistency."""
        detected_types = self.detect_data_types(data)
        valid_cells = 0
        total_cells = 0
        
        for column, expected_type in detected_types.items():
            if column in data.columns:
                series = data[column].dropna()
                total_cells += len(series)
                
                if expected_type == DataType.NUMERIC.value:
                    valid_cells += pd.to_numeric(series, errors='coerce').notna().sum()
                elif expected_type == DataType.DATETIME.value:
                    valid_cells += pd.to_datetime(series, errors='coerce').notna().sum()
                else:
                    valid_cells += len(series)  # Assume text/categorical are always valid
        
        return valid_cells / total_cells if total_cells > 0 else 0.0
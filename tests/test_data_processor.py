"""
Unit tests for data processing functionality.
"""

import pytest
import pandas as pd
import numpy as np
from unittest.mock import Mock, patch
import tempfile
import json
import xml.etree.ElementTree as ET

from src.core.data_processor import (
    DataProcessor, DataType, CleaningStrategy, CleaningRules,
    DataQualityMetrics
)


class TestDataProcessor:
    """Test data processing functionality."""
    
    @pytest.fixture
    def processor(self):
        """Create test data processor."""
        return DataProcessor()
    
    @pytest.fixture
    def sample_data(self):
        """Create sample test data."""
        return pd.DataFrame({
            'numeric': [1, 2, 3, 4, 5],
            'text': ['a', 'b', 'c', 'd', 'e'],
            'datetime': pd.date_range('2023-01-01', periods=5),
            'boolean': [True, False, True, False, True],
            'categorical': ['cat1', 'cat2', 'cat1', 'cat2', 'cat1']
        })
    
    @pytest.fixture
    def messy_data(self):
        """Create messy test data with issues."""
        return pd.DataFrame({
            'numeric': [1, 2, np.nan, 4, 100],  # Missing value and outlier
            'text': ['a', 'b', None, 'd', 'e'],  # Missing value
            'duplicate_col': [1, 1, 2, 2, 3],  # Duplicates
            'mixed': [1, 'text', 3, 4, 5]  # Mixed types
        })
    
    def test_initialization(self, processor):
        """Test processor initialization."""
        assert processor.chunk_size > 0
        assert len(processor.validators) > 0
        assert len(processor.cleaners) > 0
        assert len(processor.type_detectors) > 0
    
    def test_detect_data_types(self, processor, sample_data):
        """Test data type detection."""
        detected_types = processor.detect_data_types(sample_data)
        
        assert detected_types['numeric'] == DataType.NUMERIC.value
        assert detected_types['text'] == DataType.TEXT.value
        assert detected_types['datetime'] == DataType.DATETIME.value
        assert detected_types['boolean'] == DataType.BOOLEAN.value
        assert detected_types['categorical'] == DataType.CATEGORICAL.value
    
    def test_validate_data_valid(self, processor, sample_data):
        """Test validation of valid data."""
        result = processor.validate_data(sample_data)
        
        assert result.is_valid is True
        assert len(result.errors) == 0
        assert result.row_count == 5
        assert result.column_count == 5
    
    def test_validate_data_empty(self, processor):
        """Test validation of empty data."""
        empty_data = pd.DataFrame()
        result = processor.validate_data(empty_data)
        
        assert result.is_valid is False
        assert "Dataset is empty" in result.errors
        assert result.row_count == 0
        assert result.column_count == 0
    
    def test_validate_data_with_issues(self, processor, messy_data):
        """Test validation of data with issues."""
        result = processor.validate_data(messy_data)
        
        # Should have warnings about missing values
        assert any("missing values" in warning.lower() for warning in result.warnings)
        assert result.row_count == 5
        assert result.column_count == 4
    
    def test_clean_data_drop_missing(self, processor, messy_data):
        """Test cleaning data by dropping missing values."""
        rules = CleaningRules(missing_value_strategy=CleaningStrategy.DROP)
        cleaned = processor.clean_data(messy_data, rules)
        
        # Should have fewer rows after dropping missing values
        assert len(cleaned) < len(messy_data)
        assert not cleaned.isnull().any().any()
    
    def test_clean_data_fill_mean(self, processor):
        """Test cleaning data by filling with mean."""
        data = pd.DataFrame({
            'numeric': [1, 2, np.nan, 4, 5]
        })
        
        rules = CleaningRules(missing_value_strategy=CleaningStrategy.FILL_MEAN)
        cleaned = processor.clean_data(data, rules)
        
        # Missing value should be filled with mean
        assert not cleaned['numeric'].isnull().any()
        assert cleaned['numeric'].iloc[2] == 3.0  # Mean of [1,2,4,5]
    
    def test_clean_data_remove_duplicates(self, processor):
        """Test removing duplicate rows."""
        data = pd.DataFrame({
            'a': [1, 2, 1, 3],
            'b': [4, 5, 4, 6]
        })
        
        rules = CleaningRules(remove_duplicates=True)
        cleaned = processor.clean_data(data, rules)
        
        assert len(cleaned) == 3  # One duplicate removed
    
    def test_chunk_large_dataset_small(self, processor, sample_data):
        """Test chunking small dataset."""
        chunks = list(processor.chunk_large_dataset(sample_data))
        
        assert len(chunks) == 1
        assert chunks[0].equals(sample_data)
    
    def test_chunk_large_dataset_large(self, processor):
        """Test chunking large dataset."""
        # Create large dataset
        large_data = pd.DataFrame({
            'col1': range(25000),
            'col2': range(25000, 50000)
        })
        
        # Set small chunk size for testing
        processor.chunk_size = 10000
        chunks = list(processor.chunk_large_dataset(large_data))
        
        assert len(chunks) == 3  # 25000 / 10000 = 2.5, so 3 chunks
        assert len(chunks[0]) == 10000
        assert len(chunks[1]) == 10000
        assert len(chunks[2]) == 5000
    
    def test_handle_missing_values_auto(self, processor):
        """Test automatic missing value handling."""
        # Low missing percentage - should drop
        data_low_missing = pd.DataFrame({
            'a': [1, 2, 3, 4, 5],
            'b': [1, np.nan, 3, 4, 5]  # 20% missing
        })
        
        result = processor.handle_missing_values(data_low_missing, 'auto')
        assert len(result) == 4  # One row dropped
    
    def test_get_data_quality_metrics(self, processor, messy_data):
        """Test data quality metrics calculation."""
        metrics = processor.get_data_quality_metrics(messy_data)
        
        assert isinstance(metrics, DataQualityMetrics)
        assert metrics.total_rows == 5
        assert metrics.total_columns == 4
        assert metrics.missing_values > 0
        assert 0 <= metrics.completeness_score <= 1
        assert 0 <= metrics.consistency_score <= 1
        assert 0 <= metrics.validity_score <= 1
    
    def test_import_csv(self, processor):
        """Test CSV import."""
        # Create temporary CSV file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
            f.write("a,b,c\n1,2,3\n4,5,6\n")
            csv_path = f.name
        
        try:
            data = processor.import_csv(csv_path)
            assert len(data) == 2
            assert list(data.columns) == ['a', 'b', 'c']
            assert data.iloc[0, 0] == 1
        finally:
            import os
            os.unlink(csv_path)
    
    def test_import_json(self, processor):
        """Test JSON import."""
        # Create temporary JSON file
        test_data = [
            {"name": "Alice", "age": 30},
            {"name": "Bob", "age": 25}
        ]
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(test_data, f)
            json_path = f.name
        
        try:
            data = processor.import_json(json_path)
            assert len(data) == 2
            assert list(data.columns) == ['name', 'age']
            assert data.iloc[0]['name'] == 'Alice'
        finally:
            import os
            os.unlink(json_path)
    
    def test_import_xml(self, processor):
        """Test XML import."""
        # Create temporary XML file
        xml_content = """<?xml version="1.0"?>
        <root>
            <person>
                <name>Alice</name>
                <age>30</age>
            </person>
            <person>
                <name>Bob</name>
                <age>25</age>
            </person>
        </root>"""
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False) as f:
            f.write(xml_content)
            xml_path = f.name
        
        try:
            data = processor.import_xml(xml_path, 'person')
            assert len(data) == 2
            assert list(data.columns) == ['name', 'age']
            assert data.iloc[0]['name'] == 'Alice'
        finally:
            import os
            os.unlink(xml_path)
    
    def test_detect_numeric_type(self, processor):
        """Test numeric type detection."""
        numeric_series = pd.Series([1, 2, 3, 4, 5])
        assert processor._detect_numeric(numeric_series) is True
        
        text_series = pd.Series(['a', 'b', 'c'])
        assert processor._detect_numeric(text_series) is False
    
    def test_detect_datetime_type(self, processor):
        """Test datetime type detection."""
        datetime_series = pd.Series(pd.date_range('2023-01-01', periods=5))
        assert processor._detect_datetime(datetime_series) is True
        
        text_series = pd.Series(['a', 'b', 'c'])
        assert processor._detect_datetime(text_series) is False
    
    def test_detect_boolean_type(self, processor):
        """Test boolean type detection."""
        boolean_series = pd.Series([True, False, True])
        assert processor._detect_boolean(boolean_series) is True
        
        text_boolean_series = pd.Series(['yes', 'no', 'yes'])
        assert processor._detect_boolean(text_boolean_series) is True
        
        numeric_series = pd.Series([1, 2, 3])
        assert processor._detect_boolean(numeric_series) is False
    
    def test_detect_categorical_type(self, processor):
        """Test categorical type detection."""
        categorical_series = pd.Series(['cat1', 'cat2', 'cat1', 'cat2'] * 10)
        assert processor._detect_categorical(categorical_series) is True
        
        high_cardinality_series = pd.Series(range(100))
        assert processor._detect_categorical(high_cardinality_series) is False
    
    def test_validate_numeric_column(self, processor):
        """Test numeric column validation."""
        valid_numeric = pd.Series([1, 2, 3, 4, 5])
        errors = processor._validate_numeric(valid_numeric)
        assert len(errors) == 0
        
        invalid_numeric = pd.Series([1, 'text', 3])
        errors = processor._validate_numeric(invalid_numeric)
        assert len(errors) > 0
    
    def test_count_outliers(self, processor):
        """Test outlier counting."""
        # Series with outliers
        series_with_outliers = pd.Series([1, 2, 3, 4, 5, 100])  # 100 is outlier
        count = processor._count_outliers(series_with_outliers)
        assert count > 0
        
        # Series without outliers
        normal_series = pd.Series([1, 2, 3, 4, 5])
        count = processor._count_outliers(normal_series)
        assert count == 0
    
    def test_normalize_text_column(self, processor):
        """Test text normalization."""
        text_series = pd.Series(['  Hello  ', 'WORLD', 'Test'])
        normalized = processor._normalize_text_column(text_series)
        
        expected = pd.Series(['hello', 'world', 'test'])
        pd.testing.assert_series_equal(normalized, expected)
    
    def test_cleaning_rules_defaults(self):
        """Test cleaning rules default values."""
        rules = CleaningRules()
        
        assert rules.missing_value_strategy == CleaningStrategy.DROP
        assert rules.remove_duplicates is True
        assert rules.handle_outliers is True
        assert rules.custom_rules == {}
    
    def test_error_handling_invalid_csv(self, processor):
        """Test error handling for invalid CSV."""
        with pytest.raises(ValueError, match="Error importing CSV"):
            processor.import_csv("nonexistent_file.csv")
    
    def test_error_handling_invalid_json(self, processor):
        """Test error handling for invalid JSON."""
        with pytest.raises(ValueError, match="Error importing JSON"):
            processor.import_json("nonexistent_file.json")
    
    def test_error_handling_invalid_xml(self, processor):
        """Test error handling for invalid XML."""
        with pytest.raises(ValueError, match="Error importing XML"):
            processor.import_xml("nonexistent_file.xml")


if __name__ == "__main__":
    pytest.main([__file__])
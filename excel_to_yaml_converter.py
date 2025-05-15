"""
Excel to YAML Data Source Configuration Converter

This script converts Excel files to YAML configuration files for the QA Analytics data source system.
It analyzes Excel files and generates data source configuration YAML with appropriate metadata,
column definitions, validation rules, and relationships.

Usage:
    python excel_to_yaml_converter.py input_file.xlsx [output_file.yaml]

If output_file is not specified, it will use the input filename with a .yaml extension.
"""

import os
import sys
import re
import argparse
import pandas as pd
import numpy as np
import yaml
import datetime
from typing import Dict, List, Any, Optional, Tuple, Set
import logging
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("excel_to_yaml")

# Type hints for clarity
DataSourceConfig = Dict[str, Any]
ColumnConfig = Dict[str, Any]
ValidationRule = Dict[str, Any]


class ExcelAnalyzer:
    """Analyzes Excel files and extracts metadata, structure, and relationships."""
    
    def __init__(self, file_path: str):
        """
        Initialize the Excel analyzer with a file path.
        
        Args:
            file_path: Path to the Excel file to analyze
        """
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.extension = os.path.splitext(file_path)[1].lower()
        
        # File info
        self.last_modified = None
        self.file_size_mb = None
        
        # Data containers
        self.sheets = {}  # Dict of sheet_name -> DataFrame
        self.sheet_info = {}  # Dict of sheet_name -> metadata
        self.relationships = []  # List of detected relationships
        
        # Load the file
        self._load_file()
        
    def _load_file(self) -> None:
        """Load the Excel file and collect basic metadata."""
        try:
            # Get file stats
            file_stats = os.stat(self.file_path)
            self.last_modified = datetime.datetime.fromtimestamp(file_stats.st_mtime)
            self.file_size_mb = file_stats.st_size / (1024 * 1024)
            
            # Load based on file type
            if self.extension in ['.xlsx', '.xls']:
                # Load all sheets
                excel = pd.ExcelFile(self.file_path)
                for sheet_name in excel.sheet_names:
                    self.sheets[sheet_name] = pd.read_excel(excel, sheet_name)
                    logger.info(f"Loaded sheet '{sheet_name}' with {len(self.sheets[sheet_name])} rows")
            elif self.extension == '.csv':
                # For CSV, create a single "data" sheet
                self.sheets["data"] = pd.read_csv(self.file_path)
                logger.info(f"Loaded CSV with {len(self.sheets['data'])} rows")
            else:
                raise ValueError(f"Unsupported file type: {self.extension}")
                
        except Exception as e:
            logger.error(f"Error loading file: {e}")
            raise
            
    def analyze(self) -> Dict[str, Any]:
        """
        Analyze the Excel file and return comprehensive metadata.
        
        Returns:
            Dictionary with all the extracted metadata
        """
        # Analyze each sheet
        for sheet_name, df in self.sheets.items():
            self._analyze_sheet(sheet_name, df)
            
        # Detect relationships between sheets
        if len(self.sheets) > 1:
            self._detect_relationships()
            
        # Compile all metadata
        metadata = {
            'file_info': {
                'file_name': self.file_name,
                'file_type': self.extension.lstrip('.'),
                'last_modified': self.last_modified.isoformat(),
                'file_size_mb': round(self.file_size_mb, 2),
                'sheets': list(self.sheets.keys()),
                'total_rows': sum(len(df) for df in self.sheets.values())
            },
            'sheets': self.sheet_info,
            'relationships': self.relationships
        }
        
        return metadata
        
    def _analyze_sheet(self, sheet_name: str, df: pd.DataFrame) -> None:
        """
        Analyze a single sheet and extract its metadata.
        
        Args:
            sheet_name: Name of the sheet
            df: DataFrame containing the sheet data
        """
        # Basic sheet info
        row_count = len(df)
        col_count = len(df.columns)
        
        # Analyze columns
        columns = []
        key_candidates = []
        
        for col_name in df.columns:
            col_info = self._analyze_column(df, col_name)
            columns.append(col_info)
            
            # Check if column could be a key
            if col_info.get('unique_values', 0) == row_count and not col_info.get('has_nulls', False):
                key_candidates.append(col_name)
                
        # Detect data validation rules
        validation_rules = self._detect_validation_rules(df)
        
        # Store sheet info
        self.sheet_info[sheet_name] = {
            'row_count': row_count,
            'column_count': col_count,
            'columns': columns,
            'key_candidates': key_candidates,
            'validation_rules': validation_rules
        }
        
    def _analyze_column(self, df: pd.DataFrame, column_name: str) -> Dict[str, Any]:
        """
        Analyze a single column and extract its metadata.
        
        Args:
            df: DataFrame containing the data
            column_name: Name of the column to analyze
            
        Returns:
            Dictionary with column metadata
        """
        # Get the column
        col = df[column_name]
        
        # Basic stats
        stats = {
            'name': column_name,
            'non_null_count': col.count(),
            'has_nulls': col.isna().any(),
            'unique_values': col.nunique(),
            'inferred_type': str(col.dtype)
        }
        
        # Calculate null percentage
        if len(df) > 0:
            stats['null_percentage'] = round((1 - col.count() / len(df)) * 100, 2)
        else:
            stats['null_percentage'] = 0
            
        # Determine more specific data type
        stats['data_type'] = self._infer_data_type(col)
        
        # Add more type-specific information
        if stats['data_type'] == 'date':
            # For dates, add min and max dates
            try:
                non_null = col.dropna()
                if len(non_null) > 0:
                    stats['min_date'] = non_null.min().isoformat().split('T')[0]
                    stats['max_date'] = non_null.max().isoformat().split('T')[0]
            except Exception as e:
                logger.debug(f"Could not extract date range for {column_name}: {e}")
                
        elif stats['data_type'] in ['integer', 'float']:
            # For numeric columns, add min, max, avg
            non_null = col.dropna()
            if len(non_null) > 0:
                stats['min_value'] = float(non_null.min())
                stats['max_value'] = float(non_null.max())
                stats['avg_value'] = float(non_null.mean())
                
        elif stats['data_type'] == 'categorical':
            # For categorical data, list common values
            value_counts = col.value_counts(normalize=True)
            if len(value_counts) <= 10:  # Only if we have a reasonable number of categories
                stats['categories'] = value_counts.index.tolist()
                stats['category_counts'] = value_counts.values.tolist()
                
        # Detect column name patterns
        stats['name_patterns'] = self._detect_column_patterns(column_name)
                
        return stats
    
    def _infer_data_type(self, series: pd.Series) -> str:
        """
        Infer the data type of a pandas Series more precisely than the default dtype.
        
        Args:
            series: The pandas Series to analyze
            
        Returns:
            String with the inferred data type
        """
        # Handle obvious cases first
        if pd.api.types.is_integer_dtype(series):
            return 'integer'
        elif pd.api.types.is_float_dtype(series):
            return 'float'
        elif pd.api.types.is_bool_dtype(series):
            return 'boolean'
            
        # For object types, we need to dig deeper
        if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
            # Try to convert to datetime
            try:
                non_null = series.dropna()
                if len(non_null) > 0:
                    # Check if it looks like a date
                    pd.to_datetime(non_null, errors='raise')
                    return 'date'
            except (ValueError, TypeError):
                pass
                
            # Check if it could be categorical (few unique values)
            unique_ratio = series.nunique() / series.count() if series.count() > 0 else 0
            if unique_ratio < 0.1 or (series.nunique() <= 10 and series.count() >= 20):
                return 'categorical'
                
            # Check if it resembles an ID field
            if self._is_id_column(series.name, series):
                return 'id'
                
            # Default to string
            return 'string'
            
        # Handle datetime types
        if pd.api.types.is_datetime64_any_dtype(series):
            return 'date'
            
        # Default case
        return 'unknown'
    
    def _is_id_column(self, col_name: str, series: pd.Series) -> bool:
        """
        Check if a column seems to be an ID column.
        
        Args:
            col_name: Column name
            series: Series with column data
            
        Returns:
            True if the column appears to be an ID column
        """
        # Check name pattern
        name_lower = col_name.lower()
        id_patterns = ['id', '_id', 'code', 'key', 'num', 'number']
        
        name_match = any(pattern in name_lower for pattern in id_patterns)
        
        # Check data pattern (if it has a consistent format like all numeric or UUID-like)
        data_pattern = False
        if series.dtype == 'object' and series.nunique() > 0.5 * series.count():
            # Sample some non-null values
            sample = series.dropna().sample(min(10, series.count()))
            
            # Check if all are numeric strings
            numeric_pattern = all(str(x).isdigit() for x in sample)
            
            # Check for UUID-like pattern
            uuid_pattern = all(
                isinstance(x, str) and 
                (len(x) > 8) and 
                bool(re.match(r'^[a-zA-Z0-9_-]+$', x))
                for x in sample
            )
            
            data_pattern = numeric_pattern or uuid_pattern
            
        return name_match or data_pattern
        
    def _detect_column_patterns(self, column_name: str) -> List[str]:
        """
        Detect common patterns in column names that indicate their purpose.
        
        Args:
            column_name: Name of the column
            
        Returns:
            List of pattern matches
        """
        patterns = []
        name_lower = column_name.lower()
        
        # ID patterns
        if any(x in name_lower for x in ['id', 'key', 'code', 'number', 'num']):
            patterns.append('identifier')
            
        # Date patterns
        if any(x in name_lower for x in ['date', 'time', 'when', 'created', 'modified', 'updated']):
            patterns.append('date')
            
        # Status patterns
        if any(x in name_lower for x in ['status', 'state', 'flag', 'complete', 'active']):
            patterns.append('status')
            
        # User patterns
        if any(x in name_lower for x in ['user', 'name', 'person', 'submitter', 'approver', 'reviewer']):
            patterns.append('user')
            
        # Numeric value patterns
        if any(x in name_lower for x in ['amount', 'value', 'cost', 'price', 'score', 'rating']):
            patterns.append('numeric_value')
            
        return patterns
    
    def _detect_validation_rules(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """
        Detect potential validation rules from the data.
        
        Args:
            df: DataFrame to analyze
            
        Returns:
            List of potential validation rules
        """
        validation_rules = []
        
        # Check for required columns (no nulls)
        for col in df.columns:
            if not df[col].isna().any():
                validation_rules.append({
                    'type': 'required_column',
                    'column': col,
                    'description': f"Column '{col}' has no null values and appears to be required"
                })
                
        # Check for numeric range constraints
        for col in df.select_dtypes(include=[np.number]).columns:
            min_val = df[col].min()
            max_val = df[col].max()
            
            # Only add if we have a significant range that's not just 0-1 (boolean-like)
            if not (min_val >= 0 and max_val <= 1 and max_val - min_val <= 1):
                validation_rules.append({
                    'type': 'numeric_range',
                    'column': col,
                    'min_value': float(min_val),
                    'max_value': float(max_val),
                    'description': f"Numeric values in '{col}' fall between {min_val} and {max_val}"
                })
                
        # Check for categorical constraints
        for col in df.columns:
            unique_ratio = df[col].nunique() / df[col].count() if df[col].count() > 0 else 0
            # If column has a small number of unique values, it may be categorical
            if 0 < unique_ratio < 0.1 or (df[col].nunique() <= 10 and df[col].count() >= 20):
                values = df[col].dropna().unique().tolist()
                # Limit to showing at most 10 values
                if len(values) <= 10:
                    validation_rules.append({
                        'type': 'categorical',
                        'column': col,
                        'allowed_values': values,
                        'description': f"Column '{col}' appears to be categorical with {len(values)} distinct values"
                    })
                    
        return validation_rules
    
    def _detect_relationships(self) -> None:
        """Detect potential relationships between sheets."""
        # Only proceed if we have multiple sheets
        if len(self.sheets) <= 1:
            return
            
        # Look for matching column names across sheets
        for sheet1, info1 in self.sheet_info.items():
            for sheet2, info2 in self.sheet_info.items():
                if sheet1 == sheet2:
                    continue
                    
                # Get column names for each sheet
                columns1 = [col['name'] for col in info1['columns']]
                columns2 = [col['name'] for col in info2['columns']]
                
                # Find columns that appear in both sheets
                common_columns = set(columns1).intersection(set(columns2))
                
                # Focus on ID-like or key columns
                key_columns = [col for col in common_columns if 
                               any(pattern in col.lower() for pattern in ['id', 'key', 'code'])]
                
                for key_col in key_columns:
                    # Get values from both sheets
                    values1 = set(self.sheets[sheet1][key_col].dropna())
                    values2 = set(self.sheets[sheet2][key_col].dropna())
                    
                    # Check for value overlap
                    overlap = values1.intersection(values2)
                    if len(overlap) > 0:
                        overlap_ratio = len(overlap) / min(len(values1), len(values2)) if min(len(values1), len(values2)) > 0 else 0
                        
                        if overlap_ratio > 0.1:  # Only if meaningful overlap
                            relationship = {
                                'from_sheet': sheet1,
                                'to_sheet': sheet2,
                                'join_column': key_col,
                                'overlap_count': len(overlap),
                                'overlap_ratio': round(overlap_ratio, 2),
                                'description': f"Sheets '{sheet1}' and '{sheet2}' share values in column '{key_col}'"
                            }
                            self.relationships.append(relationship)


class YAMLGenerator:
    """Generates YAML data source configuration from Excel metadata."""
    
    def __init__(self, metadata: Dict[str, Any], original_file_path: str):
        """
        Initialize the YAML generator with metadata.
        
        Args:
            metadata: Dictionary with Excel metadata
            original_file_path: Path to the original Excel file
        """
        self.metadata = metadata
        self.original_file_path = original_file_path
        self.file_name = os.path.basename(original_file_path)
        self.data_source_name = self._generate_data_source_name()
        
    def _generate_data_source_name(self) -> str:
        """
        Generate a clean data source name from the file name.
        
        Returns:
            Clean data source name
        """
        # Strip extension
        base_name = os.path.splitext(self.file_name)[0]
        
        # Replace spaces and special chars with underscores
        clean_name = re.sub(r'[^a-zA-Z0-9]', '_', base_name)
        
        # Convert to lowercase for consistency
        clean_name = clean_name.lower()
        
        # Remove consecutive underscores
        clean_name = re.sub(r'_+', '_', clean_name)
        
        # Remove leading/trailing underscores
        clean_name = clean_name.strip('_')
        
        return clean_name
        
    def generate_config(self) -> Dict[str, Any]:
        """
        Generate a YAML data source configuration.
        
        Returns:
            Dictionary with the YAML configuration
        """
        # Start with basic structure
        config = {
            'data_sources': {
                self.data_source_name: {
                    'type': 'report',
                    'description': self._generate_description(),
                    'version': '1.0',
                    'owner': 'QA Analytics',
                    'refresh_frequency': self._infer_refresh_frequency(),
                    'last_updated': self.metadata['file_info']['last_modified'],
                    'file_type': self.metadata['file_info']['file_type'],
                    'file_pattern': self._generate_file_pattern(),
                    'validation_rules': self._generate_validation_rules(),
                    'columns_mapping': self._generate_column_mapping()
                }
            }
        }
        
        # Add sheet-specific information
        if len(self.metadata['sheets']) > 1:
            config['data_sources'][self.data_source_name]['components'] = self._generate_components()
        else:
            # For single-sheet files, we'll use the first sheet
            sheet_name = list(self.metadata['sheets'].keys())[0]
            sheet_info = self.metadata['sheets'][sheet_name]
            
            # Extract key columns
            config['data_sources'][self.data_source_name]['key_columns'] = sheet_info['key_candidates']
            
            # Set sheet name
            if self.metadata['file_info']['file_type'] in ['xlsx', 'xls']:
                config['data_sources'][self.data_source_name]['sheet_name'] = sheet_name
                
        # Add analytics mapping
        config['analytics_mapping'] = [{
            'data_source': self.data_source_name,
            'analytics': []  # This will need to be filled in manually
        }]
        
        return config
    
    def _generate_description(self) -> str:
        """
        Generate a description for the data source.
        
        Returns:
            Description string
        """
        file_type = self.metadata['file_info']['file_type'].upper()
        sheet_count = len(self.metadata['sheets'])
        total_rows = self.metadata['file_info']['total_rows']
        
        if sheet_count > 1:
            return f"{file_type} file with {sheet_count} sheets and {total_rows} total rows"
        else:
            sheet_name = list(self.metadata['sheets'].keys())[0]
            return f"{file_type} file with {total_rows} rows in sheet '{sheet_name}'"
            
    def _infer_refresh_frequency(self) -> str:
        """
        Infer the refresh frequency based on file name and content.
        
        Returns:
            Refresh frequency string
        """
        file_name_lower = self.file_name.lower()
        
        # Check for date patterns in filename
        if re.search(r'daily|day', file_name_lower):
            return 'Daily'
        elif re.search(r'weekly|week', file_name_lower):
            return 'Weekly'
        elif re.search(r'monthly|month', file_name_lower):
            return 'Monthly'
        elif re.search(r'quarterly|quarter', file_name_lower):
            return 'Quarterly'
        elif re.search(r'annual|yearly|year', file_name_lower):
            return 'Annually'
            
        # Check for date columns that might indicate frequency
        has_date_columns = False
        for sheet_info in self.metadata['sheets'].values():
            for col in sheet_info['columns']:
                if col['data_type'] == 'date':
                    has_date_columns = True
                    break
                    
        # Default based on presence of date columns
        if has_date_columns:
            return 'Weekly'  # Conservative default if dates are present
        else:
            return 'Monthly'  # More relaxed default if no dates
            
    def _generate_file_pattern(self) -> str:
        """
        Generate a file pattern for matching similar files.
        
        Returns:
            File pattern string
        """
        # Extract base name without extension
        base_name = os.path.splitext(self.file_name)[0]
        extension = os.path.splitext(self.file_name)[1]
        
        # Check for date patterns in filename
        date_matches = re.findall(r'(20\d{2})[_-]?(\d{2})[_-]?(\d{2})', base_name)
        
        if date_matches:
            # Replace date with placeholder
            for year, month, day in date_matches:
                date_str = f"{year}{month}{day}"
                short_date = f"{year}{month}"
                
                # Replace with appropriate placeholder
                if date_str in base_name:
                    base_name = base_name.replace(date_str, '{YYYY}{MM}{DD}')
                elif short_date in base_name:
                    base_name = base_name.replace(short_date, '{YYYY}{MM}')
        else:
            # Check for other numeric patterns that might be dates
            # For example, pattern like Report_20230131
            numeric_matches = re.findall(r'[_-](\d{8})[_-]?', base_name)
            if numeric_matches:
                for match in numeric_matches:
                    base_name = base_name.replace(match, '{YYYY}{MM}{DD}')
                    
        # Final pattern
        return f"{base_name}{extension}"
        
    def _generate_validation_rules(self) -> List[Dict[str, Any]]:
        """
        Generate validation rules for the data source.
        
        Returns:
            List of validation rule dictionaries
        """
        validation_rules = []
        
        # Collect all unique validation rules from all sheets
        all_required_columns = []
        
        for sheet_name, sheet_info in self.metadata['sheets'].items():
            for rule in sheet_info['validation_rules']:
                if rule['type'] == 'required_column':
                    all_required_columns.append(rule['column'])
                    
        # Add row count validation
        total_rows = self.metadata['file_info']['total_rows']
        min_threshold = max(10, int(total_rows * 0.5))  # At least 10 rows or 50% of current
        
        validation_rules.append({
            'type': 'row_count_min',
            'threshold': min_threshold,
            'description': f"Should have at least {min_threshold} rows"
        })
        
        # Add required columns validation if we have any
        if all_required_columns:
            # Limit to a reasonable number of columns
            if len(all_required_columns) > 10:
                # Prioritize columns that look like IDs or keys
                id_cols = [col for col in all_required_columns if 
                           any(pattern in col.lower() for pattern in ['id', 'key', 'code'])]
                # Add some non-ID columns to round out the list
                other_cols = [col for col in all_required_columns if col not in id_cols]
                selected_cols = id_cols + other_cols[:10-len(id_cols)]
            else:
                selected_cols = all_required_columns
                
            validation_rules.append({
                'type': 'required_columns',
                'columns': selected_cols,
                'description': "Critical columns that must be present"
            })
            
        return validation_rules
        
    def _generate_column_mapping(self) -> List[Dict[str, Any]]:
        """
        Generate column mapping for the data source.
        
        Returns:
            List of column mapping dictionaries
        """
        column_mappings = []
        processed_columns = set()
        
        # Process each sheet
        for sheet_name, sheet_info in self.metadata['sheets'].items():
            for col in sheet_info['columns']:
                col_name = col['name']
                
                # Skip if already processed
                if col_name in processed_columns:
                    continue
                    
                processed_columns.add(col_name)
                
                # Create mapping
                mapping = {
                    'source': col_name,
                    'target': col_name,  # Same name by default
                    'data_type': col['data_type']
                }
                
                # Add aliases if the same column appears with different names
                aliases = []
                for other_sheet, other_info in self.metadata['sheets'].items():
                    if other_sheet == sheet_name:
                        continue
                        
                    # Look for similar columns in other sheets
                    for other_col in other_info['columns']:
                        if self._are_columns_similar(col, other_col) and other_col['name'] != col_name:
                            aliases.append(other_col['name'])
                            processed_columns.add(other_col['name'])
                            
                if aliases:
                    mapping['aliases'] = aliases
                    
                # For categorical columns, add valid values
                if col['data_type'] == 'categorical' and 'categories' in col:
                    # Limit to 15 categories to avoid excessive size
                    if len(col['categories']) <= 15:
                        mapping['valid_values'] = col['categories']
                        
                column_mappings.append(mapping)
                
        return column_mappings
        
    def _are_columns_similar(self, col1: Dict[str, Any], col2: Dict[str, Any]) -> bool:
        """
        Check if two columns are likely to be the same based on name and content.
        
        Args:
            col1: First column metadata
            col2: Second column metadata
            
        Returns:
            True if columns appear to be the same
        """
        # If names match exactly, they're the same
        if col1['name'] == col2['name']:
            return True
            
        # Check if the names are very similar
        name1 = col1['name'].lower()
        name2 = col2['name'].lower()
        
        # Simple fuzzy matching on names
        if name1 in name2 or name2 in name1:
            # If one is substring of the other, they're likely related
            return True
            
        # Check for common abbreviations
        abbrev_patterns = [
            (r'id$', r'identifier$'),
            (r'^desc', r'^description'),
            (r'num$', r'number$'),
            (r'^qty', r'^quantity'),
            (r'^amt', r'^amount'),
            (r'^val', r'^value')
        ]
        
        for pattern1, pattern2 in abbrev_patterns:
            if (re.search(pattern1, name1) and re.search(pattern2, name2)) or \
               (re.search(pattern2, name1) and re.search(pattern1, name2)):
                return True
                
        # Not similar enough
        return False
        
    def _generate_components(self) -> List[Dict[str, Any]]:
        """
        Generate components definition for multi-sheet files.
        
        Returns:
            List of component dictionaries
        """
        components = []
        
        # For each sheet, create a component
        for sheet_name, sheet_info in self.metadata['sheets'].items():
            component = {
                'name': self._clean_component_name(sheet_name),
                'sheet_name': sheet_name,
                'key_columns': sheet_info['key_candidates'][:3] if sheet_info['key_candidates'] else []
            }
            
            components.append(component)
            
        # Add relationships
        if self.metadata['relationships']:
            # For each relationship, add join info
            for idx, rel in enumerate(self.metadata['relationships']):
                # Find the components
                from_component = next(
                    (c for c in components if c['sheet_name'] == rel['from_sheet']), 
                    None
                )
                to_component = next(
                    (c for c in components if c['sheet_name'] == rel['to_sheet']), 
                    None
                )
                
                if from_component and to_component:
                    # Add join info to the "to" component
                    if 'join_to' not in to_component:
                        to_component['join_to'] = from_component['name']
                        to_component['join_key'] = rel['join_column']
                        
        return components
        
    def _clean_component_name(self, sheet_name: str) -> str:
        """
        Clean a sheet name to be used as a component name.
        
        Args:
            sheet_name: Sheet name
            
        Returns:
            Clean component name
        """
        # Replace spaces and special chars with underscores
        clean_name = re.sub(r'[^a-zA-Z0-9]', '_', sheet_name)
        
        # Convert to lowercase for consistency
        clean_name = clean_name.lower()
        
        # Remove consecutive underscores
        clean_name = re.sub(r'_+', '_', clean_name)
        
        # Remove leading/trailing underscores
        clean_name = clean_name.strip('_')
        
        return clean_name
        
    def to_yaml(self) -> str:
        """
        Convert the configuration to YAML format.
        
        Returns:
            YAML string
        """
        config = self.generate_config()
        
        # Custom YAML formatting
        return yaml.dump(config, default_flow_style=False, sort_keys=False)


def main():
    """Main function to handle command-line usage."""
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description="Convert Excel files to YAML data source configurations")
    parser.add_argument("input_file", help="Path to the Excel file to analyze")
    parser.add_argument("output_file", nargs="?", help="Path for the output YAML file (default: input_name.yaml)")
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose logging")
    
    args = parser.parse_args()
    
    # Set logging level
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    # Validate input file
    if not os.path.exists(args.input_file):
        logger.error(f"Input file not found: {args.input_file}")
        return 1
        
    # Determine output file if not specified
    if not args.output_file:
        input_base = os.path.splitext(args.input_file)[0]
        args.output_file = f"{input_base}.yaml"
        
    try:
        # Analyze Excel file
        logger.info(f"Analyzing file: {args.input_file}")
        analyzer = ExcelAnalyzer(args.input_file)
        metadata = analyzer.analyze()
        
        # Generate YAML
        logger.info("Generating YAML configuration")
        generator = YAMLGenerator(metadata, args.input_file)
        yaml_content = generator.to_yaml()
        
        # Save output
        with open(args.output_file, 'w') as f:
            f.write(yaml_content)
            
        logger.info(f"YAML configuration saved to: {args.output_file}")
        
        # Print summary
        data_source_name = generator.data_source_name
        sheet_count = len(metadata['sheets'])
        relationship_count = len(metadata['relationships'])
        
        print(f"\nSummary:")
        print(f"- Input file: {args.input_file}")
        print(f"- Output file: {args.output_file}")
        print(f"- Data source name: {data_source_name}")
        print(f"- Sheets analyzed: {sheet_count}")
        print(f"- Relationships detected: {relationship_count}")
        print(f"\nReview the YAML configuration and adjust as needed before using.")
        
        return 0
        
    except Exception as e:
        logger.error(f"Error processing file: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())

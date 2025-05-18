import os
import yaml
import pandas as pd
import datetime
from typing import Dict, List, Optional, Tuple

from qa_analytics.utils.logging_config import setup_logging

logger = setup_logging()

logger = setup_logging()


class DataSourceManager:
    """
    Manages loading and validation of data sources defined in the registry
    """

    def __init__(self, registry_path: str = "configs/data_sources.yaml"):
        """
        Initialize with data source registry

        Args:
            registry_path: Path to data source registry file
        """
        self.registry_path = registry_path
        self.registry = {}
        self.analytics_mapping = {}
        self.settings = {}
        self.loaded_sources = {}

        # Load registry
        self._load_registry()

    def _load_registry(self) -> None:
        """Load data source registry"""
        try:
            if os.path.exists(self.registry_path):
                with open(self.registry_path, 'r', encoding='utf-8') as f:
                    registry_data = yaml.safe_load(f)

                # Extract settings, data sources, and analytics mapping
                self.settings = registry_data.get('settings', {})
                self.registry = registry_data.get('data_sources', {})

                # Process analytics mapping for easier lookup
                for mapping in registry_data.get('analytics_mapping', []):
                    data_source = mapping.get('data_source')
                    analytics = mapping.get('analytics', [])

                    if data_source and analytics:
                        for analytic_id in analytics:
                            self.analytics_mapping[analytic_id] = data_source

                logger.info(f"Loaded data source registry from {self.registry_path}")
                logger.info(
                    f"Found {len(self.registry)} data sources and {len(self.analytics_mapping)} analytic mappings")
            else:
                logger.warning(f"Data source registry not found at {self.registry_path}")
                self.registry = {}
                self.analytics_mapping = {}
                self.settings = {}
        except Exception as e:
            logger.error(f"Error loading data source registry: {e}")
            self.registry = {}
            self.analytics_mapping = {}
            self.settings = {}

    def get_data_source_for_analytic(self, analytic_id: str) -> Optional[str]:
        """
        Get the data source name for a given analytic ID

        Args:
            analytic_id: Analytic ID

        Returns:
            Data source name or None if not found
        """
        return self.analytics_mapping.get(analytic_id)

    def get_data_source_config(self, source_name: str) -> Optional[Dict]:
        """
        Get configuration for a data source

        Args:
            source_name: Data source name

        Returns:
            Data source configuration or None if not found
        """
        return self.registry.get(source_name)

    def load_data_source(self, source_name: str, file_path: str) -> Tuple[bool, Optional[pd.DataFrame], List[str]]:
        """
        Load and validate a data source

        Args:
            source_name: Data source name
            file_path: Path to data file

        Returns:
            Tuple of (success, DataFrame or None, list of warnings)
        """
        warnings = []

        # Check if source exists in registry
        if source_name not in self.registry:
            logger.error(f"Data source '{source_name}' not found in registry")
            return False, None, [f"Data source '{source_name}' not found in registry"]

        # Get source configuration
        source_config = self.registry[source_name]

        # Check if file exists
        if not os.path.exists(file_path):
            logger.error(f"Data file not found: {file_path}")
            return False, None, [f"Data file not found: {file_path}"]

        try:
            # Check file type
            file_ext = os.path.splitext(file_path)[1].lower()
            expected_type = source_config.get('file_type', 'xlsx').lower()

            if expected_type == 'xlsx' and file_ext not in ['.xlsx', '.xls']:
                warning = f"File type mismatch: expected .xlsx but got {file_ext}"
                logger.warning(warning)
                warnings.append(warning)

            # Load the data
            if file_ext in ['.xlsx', '.xls']:
                if 'sheet_name' in source_config:
                    df = pd.read_excel(file_path, sheet_name=source_config['sheet_name'])
                else:
                    df = pd.read_excel(file_path)
            elif file_ext == '.csv':
                df = pd.read_csv(file_path)
            else:
                logger.error(f"Unsupported file type: {file_ext}")
                return False, None, [f"Unsupported file type: {file_ext}"]

            # Check row count
            if not len(df):
                logger.error(f"Data source '{source_name}' contains no rows")
                return False, None, [f"Data source '{source_name}' contains no rows"]

            # Apply validation rules
            validation_warnings = self._validate_data(df, source_config)
            warnings.extend(validation_warnings)

            # Map columns if needed
            df = self._map_columns(df, source_config)

            # Convert date columns
            df = self._convert_date_columns(df, source_config)

            # Store in loaded sources with metadata
            mod_time = os.path.getmtime(file_path)
            last_modified = datetime.datetime.fromtimestamp(mod_time)

            self.loaded_sources[source_name] = {
                'data': df,
                'file_path': file_path,
                'last_modified': last_modified,
                'row_count': len(df),
                'columns': list(df.columns),
                'loaded_at': datetime.datetime.now(),
                'warnings': warnings
            }

            logger.info(f"Loaded data source '{source_name}' from {file_path} ({len(df)} rows)")

            # Check freshness
            freshness_days = self.settings.get('data_freshness_warning', 7)
            data_age = (datetime.datetime.now() - last_modified).days

            if data_age > freshness_days:
                freshness_warning = f"Data source '{source_name}' is {data_age} days old (warning threshold: {freshness_days} days)"
                logger.warning(freshness_warning)
                warnings.append(freshness_warning)

            return True, df, warnings

        except Exception as e:
            logger.error(f"Error loading data source '{source_name}': {e}")
            return False, None, [f"Error loading data source: {str(e)}"]

    def _validate_data(self, df: pd.DataFrame, source_config: Dict) -> List[str]:
        """
        Validate data against rules in configuration

        Args:
            df: DataFrame to validate
            source_config: Data source configuration

        Returns:
            List of validation warnings
        """
        warnings = []

        # Check validation rules
        for rule in source_config.get('validation_rules', []):
            rule_type = rule.get('type')

            if rule_type == 'row_count_min':
                threshold = rule.get('threshold', 0)
                if len(df) < threshold:
                    warning = f"Row count ({len(df)}) is below minimum threshold ({threshold})"
                    logger.warning(warning)
                    warnings.append(warning)

            elif rule_type == 'required_columns':
                required_cols = rule.get('columns', [])
                missing_cols = [col for col in required_cols if col not in df.columns]

                if missing_cols:
                    warning = f"Missing required columns: {', '.join(missing_cols)}"
                    logger.warning(warning)
                    warnings.append(warning)

        return warnings

    def _map_columns(self, df: pd.DataFrame, source_config: Dict) -> pd.DataFrame:
        """
        Map columns based on configuration

        Args:
            df: DataFrame to process
            source_config: Data source configuration

        Returns:
            DataFrame with mapped columns
        """
        column_mapping = {}

        # Process each column mapping
        for col_mapping in source_config.get('columns_mapping', []):
            source_col = col_mapping.get('source')
            aliases = col_mapping.get('aliases', [])
            target_col = col_mapping.get('target')

            if not source_col or not target_col:
                continue

            # If the source column exists, rename to target
            if source_col in df.columns:
                column_mapping[source_col] = target_col
                continue

            # If the source doesn't exist, check for aliases
            for alias in aliases:
                if alias in df.columns:
                    column_mapping[alias] = target_col
                    break

        # Apply column mapping if any
        if column_mapping:
            df = df.rename(columns=column_mapping)

        return df

    def _convert_date_columns(self, df: pd.DataFrame, source_config: Dict) -> pd.DataFrame:
        """
        Convert date columns to datetime based on configuration

        Args:
            df: DataFrame to process
            source_config: Data source configuration

        Returns:
            DataFrame with converted date columns
        """
        # Extract columns that should be date type
        date_columns = []
        for col_mapping in source_config.get('columns_mapping', []):
            if col_mapping.get('data_type') == 'date':
                target_col = col_mapping.get('target')
                if target_col and target_col in df.columns:
                    date_columns.append(target_col)

        # Convert each date column
        for col in date_columns:
            try:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            except Exception as e:
                logger.warning(f"Error converting column '{col}' to datetime: {e}")

        return df

    def get_data_source_info(self) -> Dict:
        """
        Get information about available data sources

        Returns:
            Dictionary with data source information
        """
        info = {}

        # Information about registry
        info['registry'] = {
            'path': self.registry_path,
            'data_sources': len(self.registry),
            'analytics_mapped': len(self.analytics_mapping)
        }

        # Information about each data source
        info['sources'] = {}
        for name, config in self.registry.items():
            source_info = {
                'description': config.get('description', ''),
                'owner': config.get('owner', 'Unknown'),
                'version': config.get('version', 'Unknown'),
                'refresh_frequency': config.get('refresh_frequency', 'Unknown'),
                'file_type': config.get('file_type', 'xlsx'),
                'is_loaded': name in self.loaded_sources,
                'validation_rules': len(config.get('validation_rules', [])),
                'mappable_columns': len(config.get('columns_mapping', []))
            }

            # Add analytics that use this source
            analytics = []
            for analytic_id, source_name in self.analytics_mapping.items():
                if source_name == name:
                    analytics.append(analytic_id)

            source_info['analytics'] = analytics

            # Add loaded data info if available
            if name in self.loaded_sources:
                loaded_info = self.loaded_sources[name]
                source_info.update({
                    'file_path': loaded_info['file_path'],
                    'last_modified': loaded_info['last_modified'],
                    'row_count': loaded_info['row_count'],
                    'loaded_at': loaded_info['loaded_at'],
                    'warnings': len(loaded_info['warnings'])
                })

            info['sources'][name] = source_info

        return info
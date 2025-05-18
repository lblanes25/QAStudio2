import os
import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional

from qa_analytics.core.validation_rules import ValidationRules
from qa_analytics.core.reference_data_manager import ReferenceDataManager
from qa_analytics.core.data_source_manager import DataSourceManager
from qa_analytics.utils.logging_config import setup_logging
from qa_analytics.core.excel_utils import extract_column_names, is_valid_excel_formula
# Import the ExcelFormulaProcessor
from qa_analytics.core.excel_engine import ExcelFormulaProcessor

logger = setup_logging()


class EnhancedDataProcessor:
    """Processes data files according to configuration rules with enhanced capabilities"""

    def __init__(self, config: Dict):
        """Initialize with configuration dictionary"""
        self.config = config
        self.validation_rules = ValidationRules()
        self.reference_data_manager = ReferenceDataManager()
        self.data_source_manager = DataSourceManager()
        self.reference_data = {}
        self.source_data = None
        self.results = None
        self.warnings = []  # Track any warnings during processing
        self.excel_processor = None  # Will hold the Excel processor instance

        # Check if we need Excel formula processing based on validation rules
        self._init_excel_processor_if_needed()

    def _init_excel_processor_if_needed(self) -> None:
        """Initialize Excel processor if any validation rules use custom formulas"""
        if not self.config:
            return

        # Check if any validation rule uses custom formulas
        needs_excel = False
        for validation in self.config.get('validations', []):
            if validation.get('rule') == 'custom_formula':
                needs_excel = True
                break

        if needs_excel:
            try:
                logger.info("Initializing Excel Formula Processor for custom formula validations")
                # Initialize the Excel processor (visible=False for production)
                self.excel_processor = ExcelFormulaProcessor(visible=False, add_to_mru=False)

                # Set the Excel processor in the validation rules
                self.validation_rules.set_excel_processor(self.excel_processor)

                logger.info("Excel Formula Processor initialized successfully")
            except Exception as e:
                logger.error(f"Failed to initialize Excel Formula Processor: {e}")
                self.warnings.append(f"Excel formula validation unavailable: {str(e)}")

    def load_source_data(self, file_path: str) -> bool:
        """
        Load source data from file using data source registry when possible

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Check if this analytic has a registered data source
            analytic_id = str(self.config.get('analytic_id'))

            # Handle both old and new format
            if 'data_source' in self.config:
                # New format - get from registry
                data_source_name = self.config['data_source'].get('name')
                if not data_source_name:
                    # Try the mapping
                    data_source_name = self.data_source_manager.get_data_source_for_analytic(analytic_id)
            else:
                # Old format - check registry mapping
                data_source_name = self.data_source_manager.get_data_source_for_analytic(analytic_id)

            if data_source_name:
                # Load via data source manager
                logger.info(f"Loading data from registered data source '{data_source_name}'")
                success, df, source_warnings = self.data_source_manager.load_data_source(data_source_name, file_path)

                if not success:
                    logger.error(f"Failed to load data from registered source: {source_warnings}")
                    return False

                self.source_data = df
                self.warnings.extend(source_warnings)

            else:
                # Fall back to direct loading
                logger.info(f"No registered data source for analytic {analytic_id}, loading directly")

                # Load data from Excel
                self.source_data = pd.read_excel(file_path)

                # Map column aliases
                self._map_column_aliases()

                # Validate required columns
                missing_columns = self._check_required_columns()
                if missing_columns:
                    logger.error(f"Missing required columns: {', '.join(missing_columns)}")
                    return False

                # Clean and prepare the data
                self._clean_data()

            logger.info(f"Successfully loaded source data with {len(self.source_data)} rows")
            return True

        except Exception as e:
            logger.error(f"Error loading source data: {e}")
            return False

    def _map_column_aliases(self) -> None:
        """Map column aliases to standard names based on configuration"""
        if self.source_data is None:
            return

        column_mapping = {}
        for column_info in self.config['source']['required_columns']:
            std_name = column_info['name']
            aliases = column_info.get('alias', [])

            # Check if standard name exists in DataFrame
            if std_name in self.source_data.columns:
                continue

            # Check if any alias exists in DataFrame
            for alias in aliases:
                if alias in self.source_data.columns:
                    column_mapping[alias] = std_name
                    break

        # Rename columns if needed
        if column_mapping:
            self.source_data = self.source_data.rename(columns=column_mapping)

    def _check_required_columns(self) -> List[str]:
        """Check that all required columns are present"""
        if self.source_data is None:
            return []

        # Get required columns (handle both old and new formats)
        if 'source' in self.config:
            # Old format
            required_columns = [col['name'] for col in self.config['source']['required_columns']]
        elif 'data_source' in self.config:
            # New format
            required_columns = self.config['data_source']['required_fields']
        else:
            logger.error("No source or data_source configuration found")
            return ["No source configuration found"]

        missing = [col for col in required_columns if col not in self.source_data.columns]
        return missing

    def _clean_data(self) -> None:
        """Clean and prepare data for analysis"""
        if self.source_data is None:
            return

        # Convert date columns to datetime
        for col_info in self.config['source']['required_columns']:
            col_name = col_info['name']
            if 'date' in col_name.lower() and col_name in self.source_data.columns:
                try:
                    self.source_data[col_name] = pd.to_datetime(
                        self.source_data[col_name],
                        errors='coerce'
                    )
                except Exception as e:
                    logger.warning(f"Error converting {col_name} to datetime: {e}")

        # Strip whitespace from string columns
        for col in self.source_data.columns:
            if self.source_data[col].dtype == 'object':
                self.source_data[col] = self.source_data[col].str.strip()

    def load_reference_data(self) -> bool:
        """
        Load reference data files specified in configuration using the reference data manager

        Returns:
            bool: True if successful or no reference data needed
        """
        reference_section = self.config.get('reference_data', self.config.get('reference_files', {}))
        if not reference_section:
            return True

        success = True
        self.reference_data = {}

        for ref_name in reference_section:
            # Load reference data through manager
            data = self.reference_data_manager.get_reference_data(ref_name)

            if data is None:
                # Fall back to legacy loading if available in config
                if 'reference_files' in self.config:
                    legacy_ref = next((ref for ref in self.config['reference_files'] if ref['name'] == ref_name), None)
                    if legacy_ref:
                        logger.warning(f"Falling back to legacy loading for reference data '{ref_name}'")
                        try:
                            path = legacy_ref['path']
                            key_col = legacy_ref['key_column']
                            value_col = legacy_ref['value_column']

                            if not os.path.exists(path):
                                logger.error(f"Reference file not found: {path}")
                                success = False
                                continue

                            # Load the reference file
                            ref_df = pd.read_excel(path)

                            # Create a dictionary from key->value columns
                            ref_dict = dict(zip(ref_df[key_col], ref_df[value_col]))

                            # Store in reference data dictionary
                            self.reference_data[ref_name] = ref_dict

                            logger.info(f"Loaded legacy reference data '{ref_name}' with {len(ref_dict)} entries")
                            continue

                        except Exception as e:
                            logger.error(f"Error loading legacy reference data '{ref_name}': {e}")
                            success = False
                            continue

                logger.error(f"Failed to load reference data '{ref_name}'")
                success = False
                continue

            # Store in reference data dictionary
            self.reference_data[ref_name] = data

            # Check freshness
            if not self.reference_data_manager.check_freshness(ref_name):
                warning = f"Reference data '{ref_name}' may be stale"
                logger.warning(warning)
                self.warnings.append(warning)

        return success

    def run_validations(self) -> None:
        """Run all validation rules and compile results"""
        if self.source_data is None:
            logger.error("Cannot run validations - no source data loaded")
            return

        # Reinitialize Excel processor if needed but not yet initialized
        if not self.excel_processor:
            self._init_excel_processor_if_needed()

        # Create result columns for each validation
        validation_results = {}

        for validation in self.config['validations']:
            rule_name = validation['rule']
            params = validation.get('parameters', {})

            # Handle custom formula validation if needed
            if rule_name == 'custom_formula':
                # Make sure we have an Excel processor for custom formulas
                if not self.excel_processor:
                    logger.error("Excel processor not available for custom formula validation")
                    validation_results[rule_name] = pd.Series(False, index=self.source_data.index)
                    self.warnings.append("Excel formula validation unavailable - results marked as non-conforming")
                    continue

                try:
                    # Ensure the original_formula parameter is present
                    if 'original_formula' not in params:
                        logger.error("Missing original_formula parameter for custom_formula")
                        validation_results[rule_name] = pd.Series(False, index=self.source_data.index)
                        continue

                    # Get the original formula
                    original_formula = params['original_formula']

                    # Ensure formula starts with equals sign
                    if not original_formula.startswith('='):
                        original_formula = f"={original_formula}"

                    # Basic validation
                    if not is_valid_excel_formula(original_formula):
                        logger.error(f"Invalid Excel formula: {original_formula}")
                        validation_results[rule_name] = pd.Series(False, index=self.source_data.index)
                        continue

                    # Extract fields used in the formula (for documentation)
                    fields_used = extract_column_names(original_formula)

                    # Log the formula details
                    logger.info(f"Using Excel formula: {original_formula}")
                    logger.info(f"Fields referenced: {fields_used}")

                except Exception as e:
                    logger.error(f"Error processing custom formula: {e}")
                    validation_results[rule_name] = pd.Series(False, index=self.source_data.index)
                    continue

            # Get the validation method by name
            if hasattr(self.validation_rules, rule_name) and callable(getattr(self.validation_rules, rule_name)):
                validation_method = getattr(self.validation_rules, rule_name)

                # Run the validation
                try:
                    if rule_name == 'title_based_approval':
                        result = validation_method(self.source_data, params, self.reference_data)
                    else:
                        result = validation_method(self.source_data, params)

                    validation_results[rule_name] = result
                    logger.info(f"Validation '{rule_name}' completed - {result.sum()} of {len(result)} records conform")

                except Exception as e:
                    logger.error(f"Error running validation '{rule_name}': {e}")
                    validation_results[rule_name] = pd.Series(False, index=self.source_data.index)
            else:
                logger.error(f"Validation rule '{rule_name}' not found")
                validation_results[rule_name] = pd.Series(False, index=self.source_data.index)

        # Calculate overall result - "GC", "PC", or "DNC"
        if validation_results:
            # Create a DataFrame with all validation results
            result_df = pd.DataFrame(validation_results)

            # Add to source data
            for col in result_df.columns:
                self.source_data[f"Valid_{col}"] = result_df[col]

            # Determine if ALL validations pass - Generally Conforms (GC)
            all_valid = result_df.all(axis=1)

            # Calculate overall compliance result
            self.source_data['Compliance'] = np.where(
                all_valid,
                'GC',  # All validations pass - Generally Conforms
                'DNC'  # Some validations fail - Does Not Conform
            )

            # Add a column to help when manually validating DNCs
            self.source_data['DNC_Validated'] = np.where(
                self.source_data['Compliance'] == 'DNC',
                'TBD',  # To be validated manually
                'N/A'  # Not applicable for GC items
            )

            logger.info(f"Validation complete: {all_valid.sum()} GC, {len(all_valid) - all_valid.sum()} DNC")
        else:
            logger.warning("No validation results calculated")
            self.source_data['Compliance'] = 'N/A'

    def generate_summary(self) -> Optional[pd.DataFrame]:
        """Generate summary statistics by group"""
        if self.source_data is None or 'Compliance' not in self.source_data:
            logger.error("Cannot generate summary - validation not complete")
            return None

        group_by_field = self.config['reporting']['group_by']

        if group_by_field not in self.source_data.columns:
            logger.error(f"Group by field '{group_by_field}' not found in data")
            return None

        # Count records by group and compliance status
        summary = self.source_data.groupby([group_by_field, 'Compliance']).size().unstack(fill_value=0)

        # Ensure all compliance categories exist
        for category in ['GC', 'PC', 'DNC']:
            if category not in summary.columns:
                summary[category] = 0

        # Calculate totals and percentages
        summary['Total'] = summary.sum(axis=1)
        summary['DNC_Percentage'] = (summary['DNC'] / summary['Total'] * 100).round(2)

        # Compare against threshold
        threshold = self.config['thresholds']['error_percentage']
        summary['Exceeds_Threshold'] = summary['DNC_Percentage'] > threshold

        # Order columns
        ordered_cols = ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage', 'Exceeds_Threshold']
        summary = summary[ordered_cols]

        # Reset index to make the group field a column
        summary = summary.reset_index()

        # Filter out any rows where the group value is the same as a column name
        # This removes any header rows that might have been mistakenly included
        summary = summary[summary[group_by_field] != group_by_field]

        # Also filter out any rows where the group value is the same as ANY column name
        for col in summary.columns:
            summary = summary[summary[group_by_field] != col]

        # Log the summary creation
        logger.info(f"Generated summary with {len(summary)} groups")

        return summary

    def process_data(self, source_file: str) -> Tuple[bool, str]:
        """
        Process data file according to configuration

        Args:
            source_file: Path to source data file

        Returns:
            Tuple of (success, message)
        """
        try:
            # Step 1: Load source data
            if not self.load_source_data(source_file):
                return False, "Failed to load source data"

            # Step 2: Load reference data if needed
            if not self.load_reference_data():
                return False, "Failed to load reference data"

            # Step 3: Run validations
            self.run_validations()

            # Step 4: Generate summary
            summary = self.generate_summary()
            if summary is None:
                return False, "Failed to generate summary"

            self.results = {
                'detail': self.source_data,
                'summary': summary,
                'warnings': self.warnings
            }

            # Return success with warning count if any
            if self.warnings:
                return True, f"Processing complete with {len(self.warnings)} warnings"
            else:
                return True, "Processing complete"

        except Exception as e:
            logger.error(f"Error during data processing: {e}")
            return False, f"Processing error: {str(e)}"

        finally:
            # Clean up Excel resources if used
            self._cleanup_resources()

    def _cleanup_resources(self) -> None:
        """Clean up any resources used during processing"""
        # Clean up Excel resources if initialized
        if self.excel_processor:
            try:
                logger.info("Cleaning up Excel resources")
                self.excel_processor.cleanup()
                self.excel_processor = None
            except Exception as e:
                logger.warning(f"Error cleaning up Excel resources: {e}")

    def __del__(self):
        """Destructor to ensure cleanup of resources"""
        self._cleanup_resources()
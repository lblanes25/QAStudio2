import os
import yaml
from typing import Dict, List, Tuple, Set
from qa_analytics.utils.logging_config import setup_logging

logger = setup_logging()


class ConfigManager:
    """Manages loading and validation of configuration files with enhanced Excel formula support"""

    def __init__(self, config_dir: str = "configs"):
        """Initialize config manager with directory of config files"""
        self.config_dir = config_dir
        self.configs = {}
        self.load_all_configs()

    def load_all_configs(self) -> None:
        """Load all configuration files from the config directory"""
        try:
            # Clear existing configurations to ensure fresh load
            self.configs = {}

            if not os.path.exists(self.config_dir):
                logger.info(f"Creating config directory: {self.config_dir}")
                os.makedirs(self.config_dir)
                self._create_sample_config()

            # Count configuration files for logging
            config_files = [f for f in os.listdir(self.config_dir) if f.endswith(('.yaml', '.yml'))]
            logger.info(f"Found {len(config_files)} configuration files in {self.config_dir}")

            for filename in os.listdir(self.config_dir):
                if filename.endswith(('.yaml', '.yml')):
                    config_path = os.path.join(self.config_dir, filename)
                    try:
                        with open(config_path, 'r', encoding='utf-8') as file:
                            config = yaml.safe_load(file)
                            if self._validate_config(config):
                                analytic_id = str(config.get('analytic_id'))
                                self.configs[analytic_id] = config
                                logger.info(f"Loaded config for QA-ID {analytic_id} from {filename}")
                            else:
                                logger.warning(f"Config file {filename} failed validation")
                    except Exception as e:
                        logger.error(f"Error loading config {filename}: {e}")
        except Exception as e:
            logger.error(f"Error accessing config directory {self.config_dir}: {e}")
            import traceback
            logger.error(traceback.format_exc())

    def _validate_config(self, config: Dict) -> bool:
        """
        Validate that a configuration has all required elements
        and that Excel formula validation rules are properly configured

        Args:
            config: Configuration dictionary to validate

        Returns:
            bool: True if valid, False otherwise
        """
        required_keys = ['analytic_id', 'analytic_name', 'validations', 'thresholds', 'reporting']

        # Check required top-level keys
        for key in required_keys:
            if key not in config:
                logger.error(f"Missing required config key: {key}")
                return False

        # Check source configuration (support both old and new format)
        if 'source' in config:
            # Old format
            if 'required_columns' not in config['source']:
                logger.error("Source config missing required_columns")
                return False
        elif 'data_source' in config:
            # New format
            if 'name' not in config['data_source']:
                logger.error("Data source config missing name")
                return False
            if 'required_fields' not in config['data_source']:
                logger.error("Data source config missing required_fields")
                return False
        else:
            logger.error("Missing either 'source' or 'data_source' configuration")
            return False

        # Validate Excel formula validation rules if present
        if 'validations' in config:
            for validation in config['validations']:
                # Check for custom formula validation rule
                if validation.get('rule') == 'custom_formula':
                    if not self._validate_formula_rule(validation):
                        return False

        return True

    def _validate_formula_rule(self, validation: Dict) -> bool:
        """
        Validate an Excel formula validation rule

        Args:
            validation: Validation rule dictionary

        Returns:
            bool: True if valid, False otherwise
        """
        # Check required parameters
        if 'parameters' not in validation:
            logger.error("Excel formula validation missing parameters")
            return False

        parameters = validation.get('parameters', {})

        # Check for original formula
        if 'original_formula' not in parameters:
            logger.error("Excel formula validation missing original_formula parameter")
            return False

        original_formula = parameters.get('original_formula', '')

        # Basic validation of formula syntax
        if not original_formula:
            logger.error("Empty Excel formula")
            return False

        # Ensure formula starts with equals sign
        if not original_formula.startswith('='):
            logger.warning(f"Excel formula doesn't start with equals sign: {original_formula}")
            # Not a fatal error, will be corrected during processing

        # Additional validation could be added here

        return True

    def _extract_fields_from_formula(self, formula: str) -> Set[str]:
        """
        Extract field names from an Excel formula

        Args:
            formula: Excel formula string

        Returns:
            Set of field names referenced in the formula
        """
        # Remove equals sign if present
        if formula.startswith('='):
            formula = formula[1:]

        fields = set()

        # Extract fields enclosed in backticks (for names with spaces)
        # Example: `First Name` = `Last Name`
        import re
        backtick_fields = re.findall(r'`([^`]+)`', formula)
        fields.update(backtick_fields)

        # Extract fields enclosed in brackets (Excel's field notation)
        # Example: [First Name] = [Last Name]
        bracket_fields = re.findall(r'\[([^\]]+)\]', formula)
        fields.update(bracket_fields)

        # Extract other potential field names (simple identifiers)
        # This is basic and may pick up functions or other non-fields
        # Example: FirstName = LastName
        # Exclude known Excel functions to reduce false positives
        excel_functions = {
            'IF', 'AND', 'OR', 'NOT', 'SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN',
            'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'ISBLANK', 'ISERROR',
            'TODAY', 'NOW', 'DATE', 'LEN', 'LEFT', 'RIGHT', 'MID', 'TRIM',
            'UPPER', 'LOWER', 'PROPER', 'TEXT', 'VALUE', 'TRUE', 'FALSE'
        }

        # Find potential identifiers - words not preceded by ' or "
        words = re.findall(r'(?<![\'"])\b([A-Za-z][A-Za-z0-9_]*)\b', formula)

        # Filter out Excel functions and common keywords
        potential_fields = {word for word in words if word not in excel_functions}
        fields.update(potential_fields)

        return fields

    def _update_required_fields(self, config: Dict) -> Dict:
        """
        Update required fields in configuration based on Excel formulas

        Args:
            config: Configuration dictionary

        Returns:
            Updated configuration dictionary
        """
        # Check for Excel formula validations
        if 'validations' not in config:
            return config

        required_fields = set()

        # Extract fields from existing required_fields
        if 'data_source' in config and 'required_fields' in config['data_source']:
            required_fields.update(config['data_source']['required_fields'])
        elif 'source' in config and 'required_columns' in config['source']:
            # Handle old format with column objects
            for col in config['source']['required_columns']:
                if isinstance(col, dict) and 'name' in col:
                    required_fields.add(col['name'])
                elif isinstance(col, str):
                    required_fields.add(col)

        # Extract fields from Excel formulas
        for validation in config['validations']:
            if validation.get('rule') == 'custom_formula':
                parameters = validation.get('parameters', {})
                formula = parameters.get('original_formula', '')

                if formula:
                    formula_fields = self._extract_fields_from_formula(formula)
                    required_fields.update(formula_fields)

        # Update the configuration with the combined fields
        if 'data_source' in config:
            config['data_source']['required_fields'] = sorted(list(required_fields))
        elif 'source' in config:
            # Handle old format - convert to new format
            logger.warning(f"Converting old source format to new format for QA-ID {config.get('analytic_id')}")
            config['data_source'] = {
                'name': f"data_source_for_qa_{config.get('analytic_id')}",
                'required_fields': sorted(list(required_fields))
            }

        return config

    def _create_sample_config(self) -> None:
        """Create a sample configuration file with enhanced fields for Excel formulas"""
        sample_config = {
            'analytic_id': 77,
            'analytic_name': 'Audit Test Workpaper Approvals',
            'analytic_description': 'This analytic evaluates workpaper approvals to ensure proper segregation of duties, correct approval sequences, and appropriate approval authority based on job titles.',
            'data_source': {
                'name': 'audit_workpaper_approvals',
                'required_fields': [
                    'Audit TW ID',
                    'TW submitter',
                    'TL approver',
                    'AL approver',
                    'Submit Date',
                    'TL Approval Date',
                    'AL Approval Date'
                ]
            },
            'reference_data': {
                'HR_Titles': {}
            },
            'validations': [
                {
                    'rule': 'segregation_of_duties',
                    'description': 'Submitter cannot be TL or AL',
                    'rationale': 'Ensures independent review by preventing the submitter from also being an approver.',
                    'parameters': {
                        'submitter_field': 'TW submitter',
                        'approver_fields': ['TL approver', 'AL approver']
                    }
                },
                {
                    'rule': 'approval_sequence',
                    'description': 'Approvals must be in order: Submit -> TL -> AL',
                    'rationale': 'Maintains proper workflow sequence to ensure the Team Lead reviews before the Audit Leader.',
                    'parameters': {
                        'date_fields_in_order': ['Submit Date', 'TL Approval Date', 'AL Approval Date']
                    }
                },
                {
                    'rule': 'title_based_approval',
                    'description': 'AL must have appropriate title',
                    'rationale': 'Ensures approval authority is limited to those with appropriate job titles.',
                    'parameters': {
                        'approver_field': 'AL approver',
                        'allowed_titles': ['Audit Leader', 'Executive Auditor', 'Audit Manager'],
                        'title_reference': 'HR_Titles'
                    }
                },
                # Add sample Excel formula validation
                {
                    'rule': 'custom_formula',
                    'description': 'Custom validation using Excel formula',
                    'rationale': 'Allows complex validation logic using familiar Excel syntax.',
                    'parameters': {
                        'original_formula': '=AND(NOT(ISBLANK(`TW submitter`)), `Submit Date` <= `TL Approval Date`)',
                        'display_name': 'Custom Validation'
                    }
                }
            ],
            'thresholds': {
                'error_percentage': 5.0,
                'rationale': 'Industry standard for audit workpapers allows for up to 5% error rate.'
            },
            'reporting': {
                'group_by': 'AL approver',
                'summary_fields': ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage'],
                'detail_required': True
            },
            'report_metadata': {
                'owner': 'Quality Assurance Team',
                'review_frequency': 'Monthly',
                'last_revised': '2025-05-01',
                'version': '1.0',
                'contact_email': 'qa_analytics@example.com'
            }
        }

        sample_path = os.path.join(self.config_dir, 'sample_qa_77.yaml')
        with open(sample_path, 'w', encoding='utf-8') as file:
            yaml.dump(sample_config, file, default_flow_style=False)

        logger.info(f"Created enhanced sample config at {sample_path}")

    def get_config(self, analytic_id: str) -> Dict:
        """
        Get configuration for a specific analytic ID
        with additional processing for Excel formula validations

        Args:
            analytic_id: Analytics ID

        Returns:
            Configuration dictionary
        """
        if analytic_id in self.configs:
            # Get a copy of the configuration
            config = self.configs[analytic_id].copy()

            # Update required fields based on Excel formulas
            config = self._update_required_fields(config)

            return config
        else:
            logger.error(f"No configuration found for QA-ID {analytic_id}")
            raise ValueError(f"No configuration found for QA-ID {analytic_id}")

    def save_config(self, config: Dict) -> bool:
        """
        Save configuration to file
        with additional validation for Excel formula rules

        Args:
            config: Configuration dictionary

        Returns:
            bool: Success
        """
        if 'analytic_id' not in config:
            logger.error("Cannot save config: missing analytic_id")
            return False

        try:
            # Validate configuration before saving
            if not self._validate_config(config):
                logger.error("Cannot save config: validation failed")
                return False

            # Update required fields based on Excel formulas
            config = self._update_required_fields(config)

            analytic_id = str(config['analytic_id'])
            filename = f"qa_{analytic_id}.yaml"
            file_path = os.path.join(self.config_dir, filename)

            with open(file_path, 'w', encoding='utf-8') as file:
                yaml.dump(config, file, default_flow_style=False)

            # Update in-memory config
            self.configs[analytic_id] = config
            logger.info(f"Saved config for QA-ID {analytic_id} to {file_path}")
            return True

        except Exception as e:
            logger.error(f"Error saving config: {e}")
            return False

    def get_available_analytics(self) -> List[Tuple[str, str]]:
        """Get list of available analytics as (id, name) tuples"""
        return [(analytic_id, config.get('analytic_name', 'Unnamed'))
                for analytic_id, config in self.configs.items()]
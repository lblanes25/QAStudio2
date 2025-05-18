import os
import yaml
import logging
from typing import Dict, List, Any, Optional, Tuple

# Set up logging
logger = logging.getLogger("qa_analytics")


class TemplateManager:
    """Manages the loading, validation, and application of templates"""

    def __init__(self, templates_dir: str = "templates"):
        """
        Initialize the template manager
        
        Args:
            templates_dir: Directory containing template files
        """
        self.templates_dir = templates_dir
        self.templates = {}
        self.metadata = {}
        
        # Load templates and metadata
        self._load_templates()
        self._load_metadata()

    def _load_templates(self) -> None:
        """Load all template files from the templates directory"""
        if not os.path.exists(self.templates_dir):
            logger.warning(f"Templates directory not found: {self.templates_dir}")
            os.makedirs(self.templates_dir)
            self.create_sample_templates()  # Create sample templates
            return

        # Count template files
        template_files = [f for f in os.listdir(self.templates_dir)
                          if f.endswith('.yaml') and f != 'metadata.yaml']

        # If no template files found, create samples
        if not template_files:
            logger.warning(f"No template files found in {self.templates_dir}")
            self.create_sample_templates()

        # Now load the templates (either existing or newly created)
        for filename in os.listdir(self.templates_dir):
            if filename.endswith('.yaml') and filename != 'metadata.yaml':
                template_path = os.path.join(self.templates_dir, filename)
                try:
                    with open(template_path, 'r', encoding='utf-8') as f:
                        template = yaml.safe_load(f)

                    # Check if this is a valid template
                    if 'template_id' in template:
                        template_id = template['template_id']
                        self.templates[template_id] = template
                        logger.info(f"Loaded template '{template_id}' from {filename}")
                except Exception as e:
                    logger.error(f"Error loading template {filename}: {e}")

    def _load_metadata(self) -> None:
        """Load template metadata file"""
        metadata_path = os.path.join(self.templates_dir, 'metadata.yaml')

        if not os.path.exists(metadata_path):
            logger.warning(f"Template metadata file not found: {metadata_path}")
            self.create_sample_templates()  # Create sample templates and metadata

        try:
            with open(metadata_path, 'r', encoding='utf-8') as f:
                self.metadata = yaml.safe_load(f)
            logger.info("Loaded template metadata")
        except Exception as e:
            logger.error(f"Error loading template metadata: {e}")
    
    def get_template(self, template_id: str) -> Optional[Dict]:
        """
        Get a template by ID
        
        Args:
            template_id: Template identifier
            
        Returns:
            Template dictionary or None if not found
        """
        return self.templates.get(template_id)
    
    def get_all_templates(self) -> List[Dict]:
        """
        Get all available templates with metadata
        
        Returns:
            List of template info dictionaries
        """
        result = []
        
        for template_id, template in self.templates.items():
            template_info = {
                'id': template_id,
                'name': template.get('template_name', 'Unnamed Template'),
                'description': template.get('template_description', ''),
                'version': template.get('template_version', '1.0'),
                'category': template.get('template_category', 'Uncategorized'),
                'parameter_count': len(template.get('template_parameters', []))
            }
            
            # Add metadata if available
            if 'templates' in self.metadata and template_id in self.metadata['templates']:
                meta = self.metadata['templates'][template_id]
                template_info.update({
                    'suitable_for': meta.get('suitable_for', []),
                    'difficulty': meta.get('difficulty', 'Medium'),
                    'validation_rules': meta.get('validation_rules', [])
                })
            
            result.append(template_info)
        
        return result
    
    def get_template_parameters(self, template_id: str) -> List[Dict]:
        """
        Get parameters for a specific template
        
        Args:
            template_id: Template identifier
            
        Returns:
            List of parameter dictionaries
        """
        template = self.get_template(template_id)
        if not template:
            return []
        
        return template.get('template_parameters', [])
    
    def get_template_categories(self) -> List[Dict]:
        """
        Get all template categories with descriptions
        
        Returns:
            List of category dictionaries
        """
        if 'categories' not in self.metadata:
            return []
        
        return [
            {'id': cat_id, 'name': cat_id, **cat_info}
            for cat_id, cat_info in self.metadata['categories'].items()
        ]
    
    def get_validation_rules(self) -> Dict:
        """
        Get information about all validation rules
        
        Returns:
            Dictionary of validation rule information
        """
        if 'validation_rules' not in self.metadata:
            return {}
        
        return self.metadata['validation_rules']

    def apply_template(self, template_id: str, parameter_values: Dict) -> Tuple[bool, Optional[Dict], Optional[str]]:
        """
        Apply a template with parameter values to generate a configuration

        Args:
            template_id: Template identifier
            parameter_values: Dictionary of parameter values

        Returns:
            Tuple of (success, config, error_message)
        """
        template = self.get_template(template_id)
        if not template:
            return False, None, f"Template '{template_id}' not found"

        # Validate that all required parameters are provided
        missing_params = []
        for param in template.get('template_parameters', []):
            if param.get('required', False) and param['name'] not in parameter_values:
                missing_params.append(param['name'])

        if missing_params:
            return False, None, f"Missing required parameters: {', '.join(missing_params)}"

        # Create the configuration
        try:
            # Start with basic configuration
            config = {
                'analytic_id': parameter_values.get('analytic_id', ''),
                'analytic_name': parameter_values.get('analytic_name', ''),
                'analytic_description': parameter_values.get('analytic_description',
                                                             template.get('template_description', '')),
            }

            # Add data source configuration
            if 'data_source' in parameter_values and parameter_values['data_source']:
                # Collect fields that will be used in validations
                required_fields = []

                # Identify fields from validation parameters
                for validation in template.get('generated_validations', []):
                    for param_name, param_template in validation.get('parameters_mapping', {}).items():
                        if isinstance(param_template, str) and param_template.startswith(
                                '{') and param_template.endswith('}'):
                            # Extract the parameter name from {param_name}
                            template_param = param_template[1:-1]
                            if template_param in parameter_values:
                                # If this parameter refers to a field name, add it to required fields
                                if any(field_keyword in param_name.lower() for field_keyword in ['field', 'column']):
                                    field_value = parameter_values[template_param]
                                    if field_value and field_value not in required_fields:
                                        required_fields.append(field_value)

                config['data_source'] = {
                    'name': parameter_values['data_source'],
                    'required_fields': required_fields
                }

            # Add reference data
            reference_params = [p for p in template.get('template_parameters', [])
                                if p.get('data_type') == 'reference' and p['name'] in parameter_values]

            if reference_params:
                config['reference_data'] = {}
                for param in reference_params:
                    ref_name = parameter_values[param['name']]
                    if ref_name:  # Only add if not empty
                        config['reference_data'][ref_name] = {}

            # Add validations
            config['validations'] = []
            for val in template.get('generated_validations', []):
                validation = {
                    'rule': val['rule'],
                    'description': val['description'],
                    'parameters': {}
                }

                # Map parameters
                for param_name, param_template in val.get('parameters_mapping', {}).items():
                    # Handle direct parameter mapping
                    if isinstance(param_template, str) and param_template.startswith('{') and param_template.endswith(
                            '}'):
                        # Extract the parameter name from {param_name}
                        template_param = param_template[1:-1]
                        if template_param in parameter_values:
                            # For lists, evaluate the string to a list
                            if isinstance(parameter_values[template_param], str) and parameter_values[
                                template_param].startswith('['):
                                try:
                                    validation['parameters'][param_name] = eval(parameter_values[template_param])
                                except Exception as e:
                                    logger.warning(f"Failed to evaluate parameter {template_param}: {e}")
                                    validation['parameters'][param_name] = parameter_values[template_param]
                            else:
                                validation['parameters'][param_name] = parameter_values[template_param]
                    else:
                        # Handle static values or complex templates
                        validation['parameters'][param_name] = param_template

                config['validations'].append(validation)

            # Add thresholds
            if 'threshold_percentage' in parameter_values:
                try:
                    threshold_value = float(parameter_values['threshold_percentage'])
                    config['thresholds'] = {
                        'error_percentage': threshold_value,
                        'rationale': parameter_values.get('threshold_rationale',
                                                          template.get('default_thresholds', {}).get('rationale', ''))
                    }
                except ValueError:
                    # If conversion fails, use default
                    config['thresholds'] = template.get('default_thresholds', {})
            else:
                config['thresholds'] = template.get('default_thresholds', {})

            # Add reporting config
            if 'group_by' in parameter_values and parameter_values['group_by']:
                config['reporting'] = {
                    'group_by': parameter_values['group_by'],
                    'summary_fields': template.get('default_reporting', {}).get('summary_fields',
                                                                                ['GC', 'PC', 'DNC', 'Total',
                                                                                 'DNC_Percentage']),
                    'detail_required': template.get('default_reporting', {}).get('detail_required', True)
                }
            else:
                report_config = template.get('default_reporting', {})
                if 'group_by' in report_config:
                    if isinstance(report_config['group_by'], str) and report_config['group_by'].startswith('{') and \
                            report_config['group_by'].endswith('}'):
                        param_name = report_config['group_by'][1:-1]
                        if param_name in parameter_values and parameter_values[param_name]:
                            config['reporting'] = {
                                'group_by': parameter_values[param_name],
                                'summary_fields': report_config.get('summary_fields',
                                                                    ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage']),
                                'detail_required': report_config.get('detail_required', True)
                            }
                    else:
                        # Use the literal value from the template
                        config['reporting'] = {
                            'group_by': report_config['group_by'],
                            'summary_fields': report_config.get('summary_fields',
                                                                ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage']),
                            'detail_required': report_config.get('detail_required', True)
                        }
                else:
                    # Default reporting if nothing specified
                    config['reporting'] = {
                        'group_by': 'Audit Leader',  # Safe default
                        'summary_fields': ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage'],
                        'detail_required': True
                    }

            # Ensure analytic_id is numeric if possible
            if 'analytic_id' in config and config['analytic_id']:
                try:
                    config['analytic_id'] = int(config['analytic_id'])
                except (ValueError, TypeError):
                    # If it can't be converted to int, keep as is
                    pass

            # Log the generated configuration
            logger.info(f"Successfully generated configuration for template {template_id}")
            return True, config, None

        except Exception as e:
            logger.error(f"Error applying template: {e}")
            return False, None, f"Error applying template: {e}"
    
    def get_example_values(self, template_id: str, mapping_name: str = None) -> Dict:
        """
        Get example parameter values for a template
        
        Args:
            template_id: Template identifier
            mapping_name: Optional name of example mapping to use
            
        Returns:
            Dictionary of example parameter values
        """
        template = self.get_template(template_id)
        if not template:
            return {}
        
        # Start with empty values
        example_values = {}
        
        # Add default example from parameters
        for param in template.get('template_parameters', []):
            if 'example' in param:
                example_values[param['name']] = param['example']
        
        # If a specific mapping is requested and available, use it
        if mapping_name and 'example_mappings' in template and mapping_name in template['example_mappings']:
            example_values.update(template['example_mappings'][mapping_name])
        
        return example_values

    # In template_manager.py
    def save_config(self, config: Dict, analytics_id: str) -> Tuple[bool, Optional[str]]:
        """
        Save a generated configuration to file

        Args:
            config: Configuration dictionary
            analytics_id: Analytics ID for filename

        Returns:
            Tuple of (success, error_message or file_path)
        """
        if not config:
            return False, "No configuration to save"

        # Change from "../../configs" to "configs" to match where ConfigManager loads from
        configs_dir = "configs"
        if not os.path.exists(configs_dir):
            os.makedirs(configs_dir)

        # Create filename
        filename = f"qa_{analytics_id}.yaml"
        file_path = os.path.join(configs_dir, filename)

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                yaml.dump(config, f, default_flow_style=False)

            logger.info(f"Saved configuration to {file_path}")
            return True, file_path
        except Exception as e:
            logger.error(f"Error saving configuration: {e}")
            return False, f"Error saving configuration: {e}"

    def create_sample_templates(self) -> None:
        """Create sample templates and metadata if none exist"""
        logger.info("Creating sample templates directory and files")

        # Create templates directory if it doesn't exist
        if not os.path.exists(self.templates_dir):
            os.makedirs(self.templates_dir)

        # Create metadata file
        metadata_path = os.path.join(self.templates_dir, 'metadata.yaml')
        metadata = {
            'categories': {
                'audit_validation': {
                    'name': 'Audit Validation',
                    'description': 'Templates for validating audit processes and workpapers',
                    'icon': 'check-circle'
                },
                'risk_assessment': {
                    'name': 'Risk Assessment',
                    'description': 'Templates for risk assessment validations',
                    'icon': 'alert-triangle'
                },
                'compliance': {
                    'name': 'Compliance',
                    'description': 'Templates for regulatory compliance checks',
                    'icon': 'shield'
                }
            },
            'validation_rules': {
                'segregation_of_duties': {
                    'name': 'Segregation of Duties',
                    'description': 'Validates that submitter and approver are different people',
                    'complexity': 'Medium'
                },
                'approval_sequence': {
                    'name': 'Approval Sequence',
                    'description': 'Validates that dates follow the correct sequence',
                    'complexity': 'Medium'
                },
                'custom_formula': {
                    'name': 'Custom Excel Formula',
                    'description': 'Uses Excel formula for custom validation logic',
                    'complexity': 'Advanced'
                }
            },
            'templates': {
                'audit_workpaper_template': {
                    'suitable_for': [
                        'Audit workpaper validations',
                        'Team member segregation of duties',
                        'Approval workflow validation'
                    ],
                    'difficulty': 'Medium',
                    'validation_rules': [
                        'segregation_of_duties',
                        'approval_sequence'
                    ]
                },
                'risk_assessment_template': {
                    'suitable_for': [
                        'Risk assessment validations',
                        'Third-party risk evaluations'
                    ],
                    'difficulty': 'Medium',
                    'validation_rules': [
                        'custom_formula'
                    ]
                }
            }
        }

        with open(metadata_path, 'w', encoding='utf-8') as f:
            yaml.dump(metadata, f, default_flow_style=False)

        # Create audit workpaper template
        audit_template = {
            'template_id': 'audit_workpaper_template',
            'template_name': 'Audit Workpaper Approvals',
            'template_description': 'Validates audit workpaper approvals for proper segregation of duties and approval sequences',
            'template_category': 'audit_validation',
            'template_version': '1.0',
            'template_parameters': [
                {
                    'name': 'analytic_id',
                    'description': 'Unique identifier for this analytic',
                    'data_type': 'string',
                    'required': True,
                    'example': '77'
                },
                {
                    'name': 'analytic_name',
                    'description': 'Descriptive name for this analytic',
                    'data_type': 'string',
                    'required': True,
                    'example': 'Audit Test Workpaper Approvals'
                },
                {
                    'name': 'data_source',
                    'description': 'Data source containing approval data',
                    'data_type': 'data_source',
                    'required': True,
                    'example': 'audit_workpaper_approvals'
                },
                {
                    'name': 'submitter_field',
                    'description': 'Field containing the submitter name',
                    'data_type': 'string',
                    'required': True,
                    'example': 'TW submitter'
                },
                {
                    'name': 'tl_approver_field',
                    'description': 'Field containing the team lead approver name',
                    'data_type': 'string',
                    'required': True,
                    'example': 'TL approver'
                },
                {
                    'name': 'al_approver_field',
                    'description': 'Field containing the audit leader approver name',
                    'data_type': 'string',
                    'required': True,
                    'example': 'AL approver'
                },
                {
                    'name': 'submit_date_field',
                    'description': 'Field containing the submission date',
                    'data_type': 'string',
                    'required': True,
                    'example': 'Submit Date'
                },
                {
                    'name': 'tl_approval_date_field',
                    'description': 'Field containing the team lead approval date',
                    'data_type': 'string',
                    'required': True,
                    'example': 'TL Approval Date'
                },
                {
                    'name': 'al_approval_date_field',
                    'description': 'Field containing the audit leader approval date',
                    'data_type': 'string',
                    'required': True,
                    'example': 'AL Approval Date'
                },
                {
                    'name': 'group_by',
                    'description': 'Field to group results by',
                    'data_type': 'string',
                    'required': True,
                    'example': 'AL approver'
                },
                {
                    'name': 'threshold_percentage',
                    'description': 'Maximum acceptable error percentage',
                    'data_type': 'number',
                    'required': True,
                    'example': '5.0'
                }
            ],
            'generated_validations': [
                {
                    'rule': 'segregation_of_duties',
                    'description': 'Submitter cannot be TL or AL',
                    'parameters_mapping': {
                        'submitter_field': '{submitter_field}',
                        'approver_fields': "['{tl_approver_field}', '{al_approver_field}']"
                    }
                },
                {
                    'rule': 'approval_sequence',
                    'description': 'Approvals must be in order: Submit -> TL -> AL',
                    'parameters_mapping': {
                        'date_fields_in_order': "['{submit_date_field}', '{tl_approval_date_field}', '{al_approval_date_field}']"
                    }
                }
            ],
            'default_thresholds': {
                'error_percentage': 5.0,
                'rationale': 'Industry standard for audit workpapers allows for up to 5% error rate.'
            },
            'default_reporting': {
                'group_by': '{group_by}',
                'summary_fields': ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage'],
                'detail_required': True
            },
            'example_mappings': {
                'workpaper_approvals': {
                    'analytic_id': '77',
                    'analytic_name': 'Audit Test Workpaper Approvals',
                    'data_source': 'audit_workpaper_approvals',
                    'submitter_field': 'TW submitter',
                    'tl_approver_field': 'TL approver',
                    'al_approver_field': 'AL approver',
                    'submit_date_field': 'Submit Date',
                    'tl_approval_date_field': 'TL Approval Date',
                    'al_approval_date_field': 'AL Approval Date',
                    'group_by': 'AL approver',
                    'threshold_percentage': '5.0'
                }
            }
        }

        audit_template_path = os.path.join(self.templates_dir, 'audit_workpaper_template.yaml')
        with open(audit_template_path, 'w', encoding='utf-8') as f:
            yaml.dump(audit_template, f, default_flow_style=False)

        # Create risk assessment template
        risk_template = {
            'template_id': 'risk_assessment_template',
            'template_name': 'Third Party Risk Assessment',
            'template_description': 'Validates third party risk assessments for proper risk evaluation and documentation',
            'template_category': 'risk_assessment',
            'template_version': '1.0',
            'template_parameters': [
                {
                    'name': 'analytic_id',
                    'description': 'Unique identifier for this analytic',
                    'data_type': 'string',
                    'required': True,
                    'example': '78'
                },
                {
                    'name': 'analytic_name',
                    'description': 'Descriptive name for this analytic',
                    'data_type': 'string',
                    'required': True,
                    'example': 'Third Party Risk Assessment Validation'
                },
                {
                    'name': 'data_source',
                    'description': 'Data source containing risk assessment data',
                    'data_type': 'data_source',
                    'required': True,
                    'example': 'third_party_risk'
                },
                {
                    'name': 'vendor_field',
                    'description': 'Field containing the third party vendor name',
                    'data_type': 'string',
                    'required': True,
                    'example': 'Third Party Vendors'
                },
                {
                    'name': 'risk_field',
                    'description': 'Field containing the risk rating',
                    'data_type': 'string',
                    'required': True,
                    'example': 'Vendor Risk Rating'
                },
                {
                    'name': 'group_by',
                    'description': 'Field to group results by',
                    'data_type': 'string',
                    'required': True,
                    'example': 'Assessment Owner'
                },
                {
                    'name': 'original_formula',
                    'description': 'Excel formula for validation',
                    'data_type': 'string',
                    'required': True,
                    'example': '=IF(NOT(ISBLANK(Third Party Vendors)), Vendor Risk Rating<>"N/A", Vendor Risk Rating="N/A")'
                },
                {
                    'name': 'threshold_percentage',
                    'description': 'Maximum acceptable error percentage',
                    'data_type': 'number',
                    'required': True,
                    'example': '5.0'
                }
            ],
            'generated_validations': [
                {
                    'rule': 'custom_formula',
                    'description': 'Third party risk validation',
                    'parameters_mapping': {
                        'original_formula': '{original_formula}',
                        'display_name': 'Third Party Risk Validation'
                    }
                }
            ],
            'default_thresholds': {
                'error_percentage': 5.0,
                'rationale': 'Industry standard for risk assessment error threshold.'
            },
            'default_reporting': {
                'group_by': '{group_by}',
                'summary_fields': ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage'],
                'detail_required': True
            },
            'example_mappings': {
                'third_party_risk': {
                    'analytic_id': '78',
                    'analytic_name': 'Third Party Risk Assessment Validation',
                    'data_source': 'third_party_risk',
                    'vendor_field': 'Third Party Vendors',
                    'risk_field': 'Vendor Risk Rating',
                    'group_by': 'Assessment Owner',
                    'original_formula': '=IF(NOT(ISBLANK(Third Party Vendors)), Vendor Risk Rating<>"N/A", Vendor Risk Rating="N/A")',
                    'threshold_percentage': '5.0'
                }
            }
        }

        risk_template_path = os.path.join(self.templates_dir, 'risk_assessment_template.yaml')
        with open(risk_template_path, 'w', encoding='utf-8') as f:
            yaml.dump(risk_template, f, default_flow_style=False)

        logger.info(f"Created sample templates in {self.templates_dir}")
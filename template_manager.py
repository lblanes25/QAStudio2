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
            return
        
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
            return
        
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
            # Start with basic configuration (Comes up in a different order due to YAML likely.)
            config = {
                'analytic_id': int(parameter_values.get('analytic_id', '0')) if parameter_values.get('analytic_id', '').isdigit() else parameter_values.get('analytic_id', ''),
                'analytic_name': parameter_values.get('analytic_name', ''),
                'analytic_description': parameter_values.get('analytic_description', 
                                                          template.get('template_description', '')),
            }
            
            # Add data source configuration
            if 'data_source' in parameter_values:
                config['data_source'] = {'name': parameter_values['data_source']}
                
            # Add reference data
            reference_params = [p for p in template.get('template_parameters', []) 
                             if p.get('data_type') == 'reference' and p['name'] in parameter_values]
            
            if reference_params:
                config['reference_data'] = {}
                for param in reference_params:
                    ref_name = parameter_values[param['name']]
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
                                except:
                                    validation['parameters'][param_name] = parameter_values[template_param]
                            else:
                                validation['parameters'][param_name] = parameter_values[template_param]
                    else:
                        # Handle static values or complex templates
                        validation['parameters'][param_name] = param_template
                
                config['validations'].append(validation)
            
            # Add thresholds
            if 'threshold_percentage' in parameter_values:
                config['thresholds'] = {
                    'error_percentage': float(parameter_values['threshold_percentage']),
                    'rationale': parameter_values.get('threshold_rationale', 
                                                   template.get('default_thresholds', {}).get('rationale', ''))
                }
            else:
                config['thresholds'] = template.get('default_thresholds', {})
            
            # Add reporting config
            if 'group_by' in parameter_values:
                config['reporting'] = {
                    'group_by': parameter_values['group_by'],
                    'summary_fields': template.get('default_reporting', {}).get('summary_fields', 
                                                                             ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage']),
                    'detail_required': template.get('default_reporting', {}).get('detail_required', True)
                }
            else:
                report_config = template.get('default_reporting', {})
                if 'group_by' in report_config and report_config['group_by'].startswith('{') and report_config['group_by'].endswith('}'):
                    param_name = report_config['group_by'][1:-1]
                    if param_name in parameter_values:
                        config['reporting'] = {
                            'group_by': parameter_values[param_name],
                            'summary_fields': report_config.get('summary_fields', 
                                                            ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage']),
                            'detail_required': report_config.get('detail_required', True)
                        }
            
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
        
        # Ensure configs directory exists
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
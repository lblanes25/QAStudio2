import os
import tkinter as tk
from tkinter import ttk, messagebox
import yaml
import logging
from typing import Dict, List, Any, Optional, Callable

from template_manager import TemplateManager

# Set up logging
logger = logging.getLogger("qa_analytics")


class ConfigWizard:
    """GUI wizard for creating and editing analytics configurations"""

    def __init__(self, parent_frame, config_manager, template_manager=None, on_config_saved=None):
        """
        Initialize the configuration wizard
        
        Args:
            parent_frame: Parent tkinter frame
            config_manager: ConfigManager instance for loading/saving configs
            template_manager: Optional TemplateManager instance
            on_config_saved: Optional callback function to call when a config is saved
        """
        self.parent = parent_frame
        self.config_manager = config_manager
        self.template_manager = template_manager or TemplateManager()
        self.on_config_saved = on_config_saved
        
        # Initialize state variables
        self.current_template_id = None
        self.parameter_entries = {}
        self.current_config = None
        self.analytics_id_var = tk.StringVar()
        self.analytics_name_var = tk.StringVar()
        self.analytics_desc_var = tk.StringVar()
        self.data_source_var = tk.StringVar()
        self.threshold_var = tk.StringVar(value="5.0")
        self.group_by_var = tk.StringVar()
        
        # Set up the wizard interface
        self._setup_ui()
    
    def _setup_ui(self):
        """Set up the wizard user interface"""
        main_frame = ttk.Frame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create notebook for wizard steps
        self.wizard_notebook = ttk.Notebook(main_frame)
        self.wizard_notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create wizard steps
        self.step1_frame = ttk.Frame(self.wizard_notebook)  # Template Selection
        self.step2_frame = ttk.Frame(self.wizard_notebook)  # Basic Configuration
        self.step3_frame = ttk.Frame(self.wizard_notebook)  # Template Parameters
        self.step4_frame = ttk.Frame(self.wizard_notebook)  # Review & Save
        
        self.wizard_notebook.add(self.step1_frame, text="1. Select Template")
        self.wizard_notebook.add(self.step2_frame, text="2. Basic Configuration")
        self.wizard_notebook.add(self.step3_frame, text="3. Template Parameters")
        self.wizard_notebook.add(self.step4_frame, text="4. Review & Save")
        
        # Set up each step
        self._setup_step1()
        self._setup_step2()
        self._setup_step3()
        self._setup_step4()
        
        # Add navigation buttons at the bottom
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.prev_btn = ttk.Button(button_frame, text="Previous", command=self._go_prev_step)
        self.prev_btn.pack(side=tk.LEFT)
        
        self.next_btn = ttk.Button(button_frame, text="Next", command=self._go_next_step)
        self.next_btn.pack(side=tk.RIGHT)
        
        # Initialize button states
        self._update_button_states()
    
    def _setup_step1(self):
        """Set up Step 1: Template Selection"""
        frame = ttk.LabelFrame(self.step1_frame, text="Available Templates")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create template filter options
        filter_frame = ttk.Frame(frame)
        filter_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(filter_frame, text="Filter by Category:").pack(side=tk.LEFT)
        
        # Get categories from template manager
        categories = self.template_manager.get_template_categories()
        category_names = ["All Categories"] + [c.get('name', '') for c in categories]
        
        self.category_var = tk.StringVar(value="All Categories")
        category_combo = ttk.Combobox(filter_frame, textvariable=self.category_var, 
                                     values=category_names, state="readonly", width=20)
        category_combo.pack(side=tk.LEFT, padx=(5, 0))
        category_combo.bind("<<ComboboxSelected>>", lambda e: self._populate_template_tree())
        
        # Create template treeview
        columns = ("Name", "Category", "Description", "Parameters", "Difficulty")
        self.template_tree = ttk.Treeview(frame, columns=columns, show="headings", height=10)
        
        # Configure columns
        self.template_tree.column("Name", width=150)
        self.template_tree.column("Category", width=100)
        self.template_tree.column("Description", width=250)
        self.template_tree.column("Parameters", width=80, anchor=tk.CENTER)
        self.template_tree.column("Difficulty", width=80, anchor=tk.CENTER)
        
        # Configure headings
        self.template_tree.heading("Name", text="Template Name")
        self.template_tree.heading("Category", text="Category")
        self.template_tree.heading("Description", text="Description")
        self.template_tree.heading("Parameters", text="Parameters")
        self.template_tree.heading("Difficulty", text="Difficulty")
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(frame, orient="vertical", command=self.template_tree.yview)
        self.template_tree.configure(yscrollcommand=tree_scroll.set)
        
        # Pack tree and scrollbar
        self.template_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind selection event
        self.template_tree.bind("<<TreeviewSelect>>", self._on_template_selected)
        
        # Create template details frame
        details_frame = ttk.LabelFrame(self.step1_frame, text="Template Details")
        details_frame.pack(fill=tk.BOTH, padx=10, pady=(10, 0))
        
        # Template details text
        self.details_text = tk.Text(details_frame, wrap=tk.WORD, height=8)
        self.details_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.details_text.config(state=tk.DISABLED)
        
        # Add scrollbar for details
        details_scroll = ttk.Scrollbar(details_frame, orient="vertical", command=self.details_text.yview)
        details_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        self.details_text.config(yscrollcommand=details_scroll.set)
        
        # Populate the template tree
        self._populate_template_tree()
    
    def _setup_step2(self):
        """Set up Step 2: Basic Configuration"""
        frame = ttk.Frame(self.step2_frame)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Analytics ID
        ttk.Label(frame, text="Analytics ID:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        ttk.Entry(frame, textvariable=self.analytics_id_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=(0, 5))
        ttk.Label(frame, text="(Required - numeric identifier for this analytic)").grid(row=0, column=2, sticky=tk.W, padx=(10, 0), pady=(0, 5))
        
        # Analytics Name
        ttk.Label(frame, text="Analytics Name:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        ttk.Entry(frame, textvariable=self.analytics_name_var, width=40).grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # Analytics Description
        ttk.Label(frame, text="Description:").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        description_entry = ttk.Entry(frame, textvariable=self.analytics_desc_var, width=60)
        description_entry.grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # Data Source
        ttk.Label(frame, text="Data Source:").grid(row=3, column=0, sticky=tk.W, pady=(10, 5))
        
        # Get data sources from config
        try:
            from data_source_manager import DataSourceManager
            data_source_manager = DataSourceManager()
            data_sources = list(data_source_manager.registry.keys())
        except:
            data_sources = []
        
        data_source_combo = ttk.Combobox(frame, textvariable=self.data_source_var, 
                                       values=data_sources, width=40)
        data_source_combo.grid(row=3, column=1, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        # Threshold
        ttk.Label(frame, text="Error Threshold %:").grid(row=4, column=0, sticky=tk.W, pady=(0, 5))
        threshold_entry = ttk.Entry(frame, textvariable=self.threshold_var, width=10)
        threshold_entry.grid(row=4, column=1, sticky=tk.W, pady=(0, 5))
        ttk.Label(frame, text="(Maximum acceptable error percentage)").grid(row=4, column=2, sticky=tk.W, padx=(10, 0), pady=(0, 5))
        
        # Group By
        ttk.Label(frame, text="Group By Field:").grid(row=5, column=0, sticky=tk.W, pady=(0, 5))
        group_by_entry = ttk.Entry(frame, textvariable=self.group_by_var, width=40)
        group_by_entry.grid(row=5, column=1, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # Explanation text
        explanation_frame = ttk.LabelFrame(self.step2_frame, text="Information")
        explanation_frame.pack(fill=tk.X, padx=10, pady=(10, 10))
        
        explanation_text = """
        Enter the basic configuration for your analytics:
        
        - Analytics ID: Unique identifier for this analytic (required)
        - Analytics Name: Descriptive name for the analytic
        - Description: Explanation of what this analytic validates
        - Data Source: The registered data source to use for this analytic
        - Error Threshold: Maximum acceptable percentage of non-conforming records
        - Group By Field: Field used to group results in reports
        """
        
        explanation_label = ttk.Label(explanation_frame, text=explanation_text, wraplength=500, justify=tk.LEFT)
        explanation_label.pack(padx=10, pady=10)
    
    def _setup_step3(self):
        """Set up Step 3: Template Parameters"""
        # Create container frame with scrollbar
        container = ttk.Frame(self.step3_frame)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add canvas for scrolling
        self.canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        
        # Configure canvas
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        # Pack canvas and scrollbar
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create frame inside canvas for parameters
        self.params_container = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.params_container, anchor="nw")
        
        # Add instructions at the top
        instructions = ttk.Label(
            self.params_container, 
            text="Configure the parameters for your template. These values will be used to generate the validation rules.",
            wraplength=500,
            justify=tk.LEFT
        )
        instructions.grid(row=0, column=0, columnspan=3, sticky=tk.W, padx=10, pady=10)
        
        # Example values button
        example_frame = ttk.Frame(self.params_container)
        example_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=10, pady=(0, 10))
        
        ttk.Label(example_frame, text="Quick Fill:").pack(side=tk.LEFT)
        
        self.example_var = tk.StringVar()
        self.example_combo = ttk.Combobox(example_frame, textvariable=self.example_var, state="readonly", width=30)
        self.example_combo.pack(side=tk.LEFT, padx=(5, 5))
        
        example_btn = ttk.Button(example_frame, text="Apply Example Values", command=self._apply_example_values)
        example_btn.pack(side=tk.LEFT)
        
        # Parameter fields will be dynamically added when a template is selected
    
    def _setup_step4(self):
        """Set up Step 4: Review & Save"""
        frame = ttk.Frame(self.step4_frame)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create preview pane
        preview_frame = ttk.LabelFrame(frame, text="Configuration Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.preview_text = tk.Text(preview_frame, wrap=tk.WORD)
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add scrollbar for preview
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_text.yview)
        preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_text.config(yscrollcommand=preview_scroll.set)
        
        # Add save button
        save_frame = ttk.Frame(frame)
        save_frame.pack(fill=tk.X, pady=10)
        
        save_btn = ttk.Button(save_frame, text="Save Configuration", command=self._save_configuration)
        save_btn.pack(side=tk.RIGHT)
        
        preview_btn = ttk.Button(save_frame, text="Refresh Preview", command=self._refresh_preview)
        preview_btn.pack(side=tk.RIGHT, padx=(0, 10))
    
    def _populate_template_tree(self):
        """Populate the template treeview with available templates"""
        # Clear existing items
        for item in self.template_tree.get_children():
            self.template_tree.delete(item)
        
        # Get templates
        templates = self.template_manager.get_all_templates()
        
        # Apply category filter if needed
        selected_category = self.category_var.get()
        if selected_category != "All Categories":
            templates = [t for t in templates if t.get('category') == selected_category]
        
        # Add to tree
        for template in templates:
            self.template_tree.insert("", tk.END, iid=template['id'], values=(
                template.get('name', 'Unnamed'),
                template.get('category', 'Uncategorized'),
                template.get('description', ''),
                template.get('parameter_count', 0),
                template.get('difficulty', 'Medium')
            ))
    
    def _on_template_selected(self, event):
        """Handle template selection from the treeview"""
        selection = self.template_tree.selection()
        if not selection:
            return
        
        # Get the selected template ID
        template_id = selection[0]
        self.current_template_id = template_id
        
        # Update template details
        template = self.template_manager.get_template(template_id)
        if not template:
            return
        
        # Update details text
        self.details_text.config(state=tk.NORMAL)
        self.details_text.delete(1.0, tk.END)
        
        details = f"Template: {template.get('template_name', 'Unnamed')}\n"
        details += f"Category: {template.get('template_category', 'Uncategorized')}\n"
        details += f"Version: {template.get('template_version', '1.0')}\n\n"
        details += f"Description: {template.get('template_description', '')}\n\n"
        
        # Add information about validations
        if 'generated_validations' in template:
            details += "Validation Rules:\n"
            for validation in template['generated_validations']:
                details += f"- {validation.get('description', validation.get('rule', ''))}\n"
        
        # Get suitable for list from metadata
        if 'templates' in self.template_manager.metadata and template_id in self.template_manager.metadata['templates']:
            meta = self.template_manager.metadata['templates'][template_id]
            if 'suitable_for' in meta:
                details += "\nSuitable for:\n"
                for item in meta['suitable_for']:
                    details += f"- {item}\n"
        
        self.details_text.insert(tk.END, details)
        self.details_text.config(state=tk.DISABLED)
        
        # Get example mappings for this template
        if 'example_mappings' in template:
            example_names = list(template['example_mappings'].keys())
            self.example_combo['values'] = example_names
            if example_names:
                self.example_combo.current(0)
        else:
            self.example_combo['values'] = []

    def _update_parameter_fields(self):
        """Update parameter fields based on selected template"""
        # Clear existing parameter fields
        for widget in self.params_container.winfo_children()[2:]:  # Skip instructions and example frame
            widget.destroy()

        self.parameter_entries = {}

        # If no template selected, return
        if not self.current_template_id:
            return

        # Get parameters for the selected template
        template = self.template_manager.get_template(self.current_template_id)
        if not template:
            return

        parameters = template.get('template_parameters', [])

        # Create fields for each parameter
        for i, param in enumerate(parameters):
            row = i + 2  # Start after instructions and example frame
            param_name = param.get('name', 'Parameter')
            required = param.get('required', False)

            # Parameter name (with required indicator)
            name_text = f"{param_name}{'*' if required else ''}:"
            name_label = ttk.Label(self.params_container, text=name_text)
            name_label.grid(row=row, column=0, sticky=tk.W, padx=10, pady=(10 if i == 0 else 5, 5))

            # Parameter value field
            var = tk.StringVar(value=param.get('example', ''))
            entry = ttk.Entry(self.params_container, textvariable=var, width=40)
            entry.grid(row=row, column=1, sticky=tk.W, pady=(10 if i == 0 else 5, 5))

            # Store the variable in the dictionary
            self.parameter_entries[param_name] = var

            # Parameter description
            desc_label = ttk.Label(self.params_container, text=param.get('description', ''),
                                   wraplength=250, justify=tk.LEFT)
            desc_label.grid(row=row, column=2, sticky=tk.W, padx=10, pady=(10 if i == 0 else 5, 5))

        # Update canvas scroll region
        self.params_container.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def _apply_example_values(self):
        """Apply example values to parameter fields"""
        if not self.current_template_id:
            return
        
        mapping_name = self.example_var.get()
        if not mapping_name:
            return
        
        # Get example values
        example_values = self.template_manager.get_example_values(self.current_template_id, mapping_name)
        
        # Apply values to fields
        for param_name, value in example_values.items():
            if param_name in self.parameter_entries:
                # Convert lists to string representation
                if isinstance(value, list):
                    value = str(value)
                self.parameter_entries[param_name].set(value)
    
    def _get_parameter_values(self):
        """Get values from parameter fields"""
        values = {
            'analytic_id': self.analytics_id_var.get(),
            'analytic_name': self.analytics_name_var.get(),
            'analytic_description': self.analytics_desc_var.get(),
            'data_source': self.data_source_var.get(),
            'threshold_percentage': self.threshold_var.get(),
            'group_by': self.group_by_var.get()
        }
        
        # Add template parameters
        for name, var in self.parameter_entries.items():
            values[name] = var.get()
        
        return values
    
    def _generate_preview_config(self):
        """Generate preview configuration from entered values"""
        if not self.current_template_id:
            return None
        
        # Get all parameter values
        parameter_values = self._get_parameter_values()

        if 'analytic_id' in parameter_values and parameter_values['analytic_id']:
            try:
                # Convert string ID to integer
                parameter_values['analytic_id'] = int(parameter_values['analytic_id'])
            except ValueError:
                pass  # If conversion fails, keep as string
        
        # Apply template
        success, config, error = self.template_manager.apply_template(
            self.current_template_id, parameter_values)
        
        if not success:
            messagebox.showerror("Error", f"Failed to generate configuration: {error}")
            return None
        
        return config
    
    def _refresh_preview(self):
        """Refresh the configuration preview"""
        config = self._generate_preview_config()
        if not config:
            return
        
        self.current_config = config
        
        # Update preview text
        self.preview_text.delete(1.0, tk.END)
        
        # Convert to YAML for display
        config_yaml = yaml.dump(config, default_flow_style=False)
        self.preview_text.insert(tk.END, config_yaml)

    def _save_configuration(self):
        """Save the configuration to file"""
        # Validate analytics ID
        analytics_id = self.analytics_id_var.get().strip()
        if not analytics_id:
            messagebox.showerror("Error", "Analytics ID is required")
            return

        # Generate configuration if needed
        if not self.current_config:
            self._refresh_preview()
            if not self.current_config:
                return

        # Save configuration
        success, result = self.template_manager.save_config(self.current_config, analytics_id)

        if success:
            messagebox.showinfo("Success", f"Configuration saved to {result}")

            # Reload configurations in the config manager
            if hasattr(self.config_manager, 'load_all_configs'):
                self.config_manager.load_all_configs()

            # Call the callback if provided
            if self.on_config_saved:
                self.on_config_saved()

        else:
            messagebox.showerror("Error", result)
    
    def _go_next_step(self):
        """Navigate to the next wizard step"""
        current_tab = self.wizard_notebook.index(self.wizard_notebook.select())
        
        # Validate before moving to next step
        if current_tab == 0:  # From Step 1 to Step 2
            if not self.current_template_id:
                messagebox.showinfo("Select Template", "Please select a template to continue")
                return

        elif current_tab == 1:  # From Step 2 to Step 3
            # Validate basic configuration
            if not self.analytics_id_var.get().strip():
                messagebox.showinfo("Required Field", "Analytics ID is required")
                return

            if not self.analytics_name_var.get().strip():
                messagebox.showinfo("Required Field", "Analytics Name is required")
                return

            # Update parameter fields for Step 3
            self._update_parameter_fields()
        
        elif current_tab == 2:  # From Step 3 to Step 4
            # Update preview for Step 4
            self._refresh_preview()
        
        # Go to next tab
        if current_tab < 3:  # 3 is the last tab
            self.wizard_notebook.select(current_tab + 1)
            self._update_button_states()
    
    def _go_prev_step(self):
        """Navigate to the previous wizard step"""
        current_tab = self.wizard_notebook.index(self.wizard_notebook.select())
        
        if current_tab > 0:
            self.wizard_notebook.select(current_tab - 1)
            self._update_button_states()
    
    def _update_button_states(self):
        """Update the state of navigation buttons based on current step"""
        current_tab = self.wizard_notebook.index(self.wizard_notebook.select())
        
        # Update Previous button
        if current_tab == 0:
            self.prev_btn.config(state=tk.DISABLED)
        else:
            self.prev_btn.config(state=tk.NORMAL)
        
        # Update Next button
        if current_tab == 3:  # Last step
            self.next_btn.config(state=tk.DISABLED)
        else:
            self.next_btn.config(state=tk.NORMAL)
    
    def load_existing_config(self, analytics_id):
        """
        Load an existing configuration for editing
        
        Args:
            analytics_id: Analytics ID to load
        """
        try:
            # Get configuration from config manager
            config = self.config_manager.get_config(analytics_id)
            if not config:
                messagebox.showerror("Error", f"Configuration for QA-{analytics_id} not found")
                return False
            
            # TODO: Implement reverse mapping from config to template parameters
            messagebox.showinfo("Not Implemented", "Editing existing configurations is not yet implemented")
            return False
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load configuration: {e}")
            return False
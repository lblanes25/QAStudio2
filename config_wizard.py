import os
import tkinter as tk
from tkinter import ttk, messagebox
import yaml
import logging
from typing import Dict, List, Any, Optional, Callable

from template_manager import TemplateManager
from excel_formula_ui import ExcelFormulaFrame
from excel_formula_parser import ExcelFormulaParser
from step_tracker import StepTracker  # Import the new StepTracker component

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
            on_config_saved: Optional callback for when config is saved
        """
        self.parent = parent_frame
        self.config_manager = config_manager
        self.template_manager = template_manager or TemplateManager()
        self.on_config_saved = on_config_saved  # Callback when config is saved

        # Initialize state variables
        self.current_step = 0  # Start at first step
        self.current_template_id = None
        self.template_parameters = []
        self.parameter_entries = {}
        self.current_config = None
        self.analytics_id_var = tk.StringVar()
        self.analytics_name_var = tk.StringVar()
        self.analytics_desc_var = tk.StringVar()
        self.data_source_var = tk.StringVar()
        self.threshold_var = tk.StringVar(value="5.0")
        self.group_by_var = tk.StringVar()

        # Variables for validation rules
        self.validation_params = {}
        self.validation_metadata = {}

        # Set up the wizard interface
        self._setup_ui()

    def _setup_ui(self):
        """Set up the wizard user interface with StepTracker"""
        main_frame = ttk.Frame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Define the steps for our wizard
        steps = ["Select Template", "Basic Configuration", "Template Parameters", "Review & Save"]

        # Add StepTracker at the top
        self.step_tracker = StepTracker(main_frame, steps, initial_step=self.current_step)
        self.step_tracker.pack(fill=tk.X, pady=(0, 20))

        # Create a frame for step content that will change based on current step
        self.content_frame = ttk.Frame(main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        # Add navigation buttons at the bottom
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        self.prev_btn = ttk.Button(button_frame, text="Previous", command=self._go_prev_step)
        self.prev_btn.pack(side=tk.LEFT)

        self.next_btn = ttk.Button(button_frame, text="Next", command=self._go_next_step)
        self.next_btn.pack(side=tk.RIGHT)

        # Display the initial step
        self._display_current_step()

        # Initialize button states
        self._update_button_states()

    def _display_current_step(self):
        """Display the content for the current step"""
        # Clear current content
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        # Display content based on current step
        if self.current_step == 0:
            self._display_step1()  # Template Selection
        elif self.current_step == 1:
            self._display_step2()  # Basic Configuration
        elif self.current_step == 2:
            self._display_step3()  # Template Parameters
        elif self.current_step == 3:
            self._display_step4()  # Review & Save

    def _display_step1(self):
        """Display Step 1: Template Selection"""
        # Create frames
        frame = ttk.LabelFrame(self.content_frame, text="Available Templates")
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
        details_frame = ttk.LabelFrame(self.content_frame, text="Template Details")
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

    def _display_step2(self):
        """Display Step 2: Basic Configuration"""
        frame = ttk.Frame(self.content_frame)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Analytics ID
        ttk.Label(frame, text="Analytics ID:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        ttk.Entry(frame, textvariable=self.analytics_id_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=(0, 5))
        ttk.Label(frame, text="(Required - numeric identifier for this analytic)").grid(row=0, column=2, sticky=tk.W,
                                                                                        padx=(10, 0), pady=(0, 5))

        # Analytics Name
        ttk.Label(frame, text="Analytics Name:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        ttk.Entry(frame, textvariable=self.analytics_name_var, width=40).grid(row=1, column=1, columnspan=2,
                                                                              sticky=tk.W, pady=(0, 5))

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
        ttk.Label(frame, text="(Maximum acceptable error percentage)").grid(row=4, column=2, sticky=tk.W, padx=(10, 0),
                                                                            pady=(0, 5))

        # Group By
        ttk.Label(frame, text="Group By Field:").grid(row=5, column=0, sticky=tk.W, pady=(0, 5))
        group_by_entry = ttk.Entry(frame, textvariable=self.group_by_var, width=40)
        group_by_entry.grid(row=5, column=1, columnspan=2, sticky=tk.W, pady=(0, 5))

        # Explanation text
        explanation_frame = ttk.LabelFrame(self.content_frame, text="Information")
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

    def _display_step3(self):
        """Display Step 3: Template Parameters"""
        # Create container frame with scrollbar
        container = ttk.Frame(self.content_frame)
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

        # Initialize next_row counter for dynamic parameter fields
        self.next_row = 1

        # Update parameter fields based on selected template
        self._update_parameter_fields()

    def _display_step4(self):
        """Display Step 4: Review & Save"""
        frame = ttk.Frame(self.content_frame)
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

        # Refresh the preview when entering this step
        self._refresh_preview()

    def _update_parameter_fields(self):
        """Update parameter fields based on selected template"""
        # Clear existing parameter fields beyond the instructions
        for widget in self.params_container.winfo_children()[1:]:
            widget.destroy()

        self.parameter_entries = {}
        self.next_row = 1

        # If no template selected, return
        if not self.current_template_id:
            return

        # Get parameters for the selected template
        template = self.template_manager.get_template(self.current_template_id)
        if not template:
            return

        self.template_parameters = template.get('template_parameters', [])

        # Example values button if there are example mappings
        if 'example_mappings' in template and template['example_mappings']:
            example_frame = ttk.Frame(self.params_container)
            example_frame.grid(row=self.next_row, column=0, columnspan=3, sticky=tk.W, padx=10, pady=(0, 10))
            self.next_row += 1

            ttk.Label(example_frame, text="Quick Fill:").pack(side=tk.LEFT)

            self.example_var = tk.StringVar()
            example_combo = ttk.Combobox(
                example_frame,
                textvariable=self.example_var,
                values=list(template['example_mappings'].keys()),
                state="readonly",
                width=30
            )
            example_combo.pack(side=tk.LEFT, padx=(5, 5))

            example_btn = ttk.Button(example_frame, text="Apply Example Values", command=self._apply_example_values)
            example_btn.pack(side=tk.LEFT)

        # Check if this is a custom formula template
        is_custom_formula_template = (self.current_template_id == 'custom_formula' or
                                      any(val.get('rule', '') == 'custom_formula'
                                          for val in template.get('generated_validations', [])))

        # Add Excel Formula UI if this is a custom formula template
        if is_custom_formula_template:
            # Add a separator
            ttk.Separator(self.params_container, orient='horizontal').grid(
                row=self.next_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
            self.next_row += 1

            # Add title for Excel Formula section
            ttk.Label(
                self.params_container,
                text="Excel Formula Validation",
                font=("Arial", 12, "bold")
            ).grid(row=self.next_row, column=0, columnspan=3, sticky=tk.W, pady=(10, 0), padx=10)
            self.next_row += 1

            # Create a custom Excel formula panel
            custom_formula_frame = ttk.LabelFrame(self.params_container, text="Custom Excel Formula")
            custom_formula_frame.grid(row=self.next_row, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=10, pady=5)
            self.next_row += 1

            # Add description
            description_label = ttk.Label(
                custom_formula_frame,
                text="Enter your validation logic using familiar Excel syntax. Reference field names exactly as they appear in your data.",
                wraplength=600,
                justify=tk.LEFT
            )
            description_label.pack(fill=tk.X, padx=10, pady=10)

            # Formula input
            formula_input_frame = ttk.Frame(custom_formula_frame)
            formula_input_frame.pack(fill=tk.X, padx=10)

            ttk.Label(formula_input_frame, text="Formula:").grid(row=0, column=0, sticky=tk.W, pady=5)
            self.formula_var = tk.StringVar()
            formula_entry = ttk.Entry(formula_input_frame, textvariable=self.formula_var, width=60)
            formula_entry.grid(row=0, column=1, sticky=tk.EW, padx=(5, 0), pady=5)

            # Parsed formula display
            ttk.Label(formula_input_frame, text="Parsed Formula:").grid(row=1, column=0, sticky=tk.W, pady=5)
            self.parsed_formula_text = tk.Text(formula_input_frame, height=3, wrap=tk.WORD)
            self.parsed_formula_text.grid(row=1, column=1, sticky=tk.EW, padx=(5, 0), pady=5)
            self.parsed_formula_text.config(state=tk.DISABLED)

            # Configure column weights
            formula_input_frame.columnconfigure(1, weight=1)

            # Status indicator
            self.formula_status_var = tk.StringVar(value="Enter a formula")
            status_label = ttk.Label(
                custom_formula_frame,
                textvariable=self.formula_status_var,
                foreground="gray"
            )
            status_label.pack(fill=tk.X, padx=10, pady=(0, 10))

            # Set up test section
            test_frame = ttk.LabelFrame(custom_formula_frame, text="Test Formula")
            test_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

            test_description = ttk.Label(
                test_frame,
                text="Generate sample data or upload a file to test your formula.",
                wraplength=600
            )
            test_description.pack(fill=tk.X, padx=10, pady=(5, 10))

            # Test options
            test_options_frame = ttk.Frame(test_frame)
            test_options_frame.pack(fill=tk.X, padx=10)

            self.data_source_var = tk.StringVar(value="generate")
            generate_radio = ttk.Radiobutton(
                test_options_frame,
                text="Generate Sample Data",
                variable=self.data_source_var,
                value="generate",
                command=self._update_test_options
            )
            generate_radio.pack(side=tk.LEFT)

            existing_radio = ttk.Radiobutton(
                test_options_frame,
                text="Use Existing Data",
                variable=self.data_source_var,
                value="existing",
                command=self._update_test_options
            )
            existing_radio.pack(side=tk.LEFT, padx=(20, 0))

            # Sample data options
            self.sample_frame = ttk.Frame(test_frame)
            self.sample_frame.pack(fill=tk.X, padx=10, pady=(10, 0))

            ttk.Label(self.sample_frame, text="Records:").pack(side=tk.LEFT)
            self.record_count_var = tk.StringVar(value="100")
            ttk.Entry(self.sample_frame, textvariable=self.record_count_var, width=8).pack(side=tk.LEFT, padx=(5, 20))

            ttk.Label(self.sample_frame, text="Error %:").pack(side=tk.LEFT)
            self.error_pct_var = tk.StringVar(value="20")
            ttk.Entry(self.sample_frame, textvariable=self.error_pct_var, width=8).pack(side=tk.LEFT, padx=(5, 0))

            # File selection frame
            self.file_frame = ttk.Frame(test_frame)
            ttk.Label(self.file_frame, text="Data File:").pack(side=tk.LEFT)
            self.file_var = tk.StringVar()
            ttk.Entry(self.file_frame, textvariable=self.file_var, width=40).pack(side=tk.LEFT, padx=5)
            ttk.Button(self.file_frame, text="Browse...", command=self._browse_test_file).pack(side=tk.LEFT)

            # Progress bar and test button
            progress_frame = ttk.Frame(test_frame)
            progress_frame.pack(fill=tk.X, padx=10, pady=10)

            self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="indeterminate", length=200)
            self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

            self.test_btn = ttk.Button(progress_frame, text="Test Formula", command=self._test_formula)
            self.test_btn.pack(side=tk.RIGHT)

            # Results frame (initially hidden)
            self.results_frame = ttk.LabelFrame(custom_formula_frame, text="Test Results")

            # Set up formula change callback
            self.formula_var.trace_add("write", self._on_formula_changed)

            # Show the correct frame based on the data source option
            self._update_test_options()

            # Import parser if not already present
            if not hasattr(self, 'formula_parser'):
                try:
                    from excel_formula_parser import ExcelFormulaParser
                    self.formula_parser = ExcelFormulaParser()
                except ImportError:
                    logger.error("Could not import ExcelFormulaParser")

        # Create fields for each parameter
        for param in self.template_parameters:
            param_name = param.get('name', 'Parameter')
            required = param.get('required', False)
            data_type = param.get('data_type', 'string')

            # Skip formula parameters if we already have the formula UI
            if data_type == 'formula' and is_custom_formula_template and hasattr(self, 'formula_var'):
                # Store a variable for this parameter, but don't show an entry field
                self.parameter_entries[param_name] = self.formula_var
                continue

            # Parameter name (with required indicator)
            name_text = f"{param_name}{'*' if required else ''}:"
            name_label = ttk.Label(self.params_container, text=name_text)
            name_label.grid(row=self.next_row, column=0, sticky=tk.W, padx=10,
                            pady=(10 if self.next_row == 1 else 5, 5))

            # Parameter value field
            if data_type == 'formula':
                # Store a variable for this parameter
                var = tk.StringVar(value=param.get('example', ''))
                self.parameter_entries[param_name] = var

                # Create formula entry
                formula_entry = ttk.Entry(self.params_container, textvariable=var, width=60)
                formula_entry.grid(row=self.next_row, column=1, sticky=tk.W, pady=(10 if self.next_row == 1 else 5, 5))
            else:
                # Standard entry field for other types
                var = tk.StringVar(value=param.get('example', ''))
                entry = ttk.Entry(self.params_container, textvariable=var, width=40)
                entry.grid(row=self.next_row, column=1, sticky=tk.W, pady=(10 if self.next_row == 1 else 5, 5))

                # Store the variable in the dictionary
                self.parameter_entries[param_name] = var

            # Parameter description
            desc_label = ttk.Label(
                self.params_container,
                text=param.get('description', ''),
                wraplength=250,
                justify=tk.LEFT
            )
            desc_label.grid(row=self.next_row, column=2, sticky=tk.W, padx=10,
                            pady=(10 if self.next_row == 1 else 5, 5))

            self.next_row += 1

        # Update canvas scroll region
        self.params_container.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _update_test_options(self):
        """Update test options based on selection"""
        if not hasattr(self, 'data_source_var'):
            return

        source = self.data_source_var.get()

        if source == "generate":
            if hasattr(self, 'sample_frame'):
                self.sample_frame.pack(fill=tk.X, padx=10, pady=(10, 0))
            if hasattr(self, 'file_frame'):
                self.file_frame.pack_forget()
        else:
            if hasattr(self, 'sample_frame'):
                self.sample_frame.pack_forget()
            if hasattr(self, 'file_frame'):
                self.file_frame.pack(fill=tk.X, padx=10, pady=(10, 0))

    def _browse_test_file(self):
        """Browse for a data file"""
        from tkinter import filedialog

        filename = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")],
            title="Select Data File"
        )

        if filename and hasattr(self, 'file_var'):
            self.file_var.set(filename)

    def _on_formula_changed(self, *args):
        """Handle formula changes and validate in real-time"""
        if not hasattr(self, 'formula_var') or not hasattr(self, 'formula_parser'):
            return

        formula = self.formula_var.get()

        if not formula:
            self._update_formula_status("Enter a formula", "gray")
            self._update_parsed_display("")
            return

        try:
            # Parse the formula
            parsed_formula, fields_used = self.formula_parser.parse(formula)

            # Update UI
            self._update_formula_status("Formula is valid", "green")
            self._update_parsed_display(parsed_formula)

            # Store for validation
            self.validation_params = {
                'original_formula': formula,
                'formula': parsed_formula,
                'fields_used': fields_used
            }

        except Exception as e:
            # Update UI
            self._update_formula_status(f"Error: {str(e)}", "red")
            self._update_parsed_display("")

    def _update_formula_status(self, message, color="black"):
        """Update formula status message"""
        if hasattr(self, 'formula_status_var'):
            self.formula_status_var.set(message)

            # Find the status label and update its color
            for widget in self.params_container.winfo_children():
                if isinstance(widget, ttk.LabelFrame) and widget.winfo_children():
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Label) and child.cget('textvariable') == str(self.formula_status_var):
                            child.configure(foreground=color)
                            break

    def _update_parsed_display(self, text):
        """Update the parsed formula display"""
        if not hasattr(self, 'parsed_formula_text'):
            return

        self.parsed_formula_text.config(state=tk.NORMAL)
        self.parsed_formula_text.delete(1.0, tk.END)
        if text:
            self.parsed_formula_text.insert(tk.END, text)
        self.parsed_formula_text.config(state=tk.DISABLED)

    def _test_formula(self):
        """Test the formula with sample data"""
        if not hasattr(self, 'formula_var') or not self.formula_var.get():
            messagebox.showinfo("Formula Required", "Please enter a formula to test")
            return

        # Implementation of formula testing would go here
        messagebox.showinfo("Test Formula", "Formula testing not implemented yet")

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
            if hasattr(self, 'example_combo'):
                self.example_combo['values'] = example_names
                if example_names:
                    self.example_combo.current(0)

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

        # If we have a formula frame, set the formula if applicable
        if hasattr(self, 'formula_frame') and 'original_formula' in example_values:
            self.formula_frame.set_formula(example_values['original_formula'])

    def _handle_formula_change(self, original_formula, parsed_formula, fields_used):
        """Handle changes to the formula in the main formula UI"""
        # Update validation parameters
        self.validation_params['original_formula'] = original_formula
        self.validation_params['formula'] = parsed_formula

        # Store fields used for validation and documentation
        self.validation_metadata['fields_used'] = fields_used

        # If we have a parameter for the formula, update it
        for param in self.template_parameters:
            if param.get('name') == 'original_formula' or param.get('data_type') == 'formula':
                if param.get('name') in self.parameter_entries:
                    self.parameter_entries[param.get('name')].set(original_formula)

    def _handle_formula_parameter(self, param_name, original_formula, parsed_formula, fields_used):
        """Handle changes to a formula parameter field"""
        if param_name in self.parameter_entries:
            # Store the original formula in the parameter entry
            self.parameter_entries[param_name].set(original_formula)

            # If we have a dedicated place to store the parsed formula, use it
            parsed_param = f"{param_name}_parsed"
            if parsed_param in self.parameter_entries:
                self.parameter_entries[parsed_param].set(parsed_formula)

            # Store fields for validation and documentation
            fields_param = f"{param_name}_fields"
            if fields_param in self.parameter_entries:
                self.parameter_entries[fields_param].set(",".join(fields_used))

    def _get_parameter_values(self):
        """Get values from parameter fields"""
        values = {}

        # Add basic fields
        values['analytic_id'] = self.analytics_id_var.get()
        values['analytic_name'] = self.analytics_name_var.get()
        values['analytic_description'] = self.analytics_desc_var.get()
        values['data_source'] = self.data_source_var.get()
        values['threshold_percentage'] = self.threshold_var.get()
        values['group_by'] = self.group_by_var.get()

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

        # Apply template
        success, config, error = self.template_manager.apply_template(
            self.current_template_id, parameter_values)

        if not success:
            messagebox.showerror("Error", f"Failed to generate configuration: {error}")
            return None

        # If this is a custom formula configuration, add the validation
        is_custom_formula = (self.current_template_id == 'custom_formula' or
                             any(v.get('rule', '') == 'custom_formula'
                                 for v in config.get('validations', [])))

        if is_custom_formula and 'validation_params' in self.__dict__ and self.validation_params:
            # Find and update the custom formula validation
            for validation in config.get('validations', []):
                if validation.get('rule') == 'custom_formula':
                    # Update parameters
                    validation['parameters'] = self.validation_params

                    # Add metadata if available
                    if 'validation_metadata' in self.__dict__ and self.validation_metadata:
                        validation['metadata'] = self.validation_metadata

                    break

        return config

    def _refresh_preview(self):
        """Refresh the configuration preview"""
        # Generate the configuration
        config = self._generate_preview_config()
        if not config:
            return

        self.current_config = config

        # Make sure preview_text exists before updating it
        if not hasattr(self, 'preview_text') or not self.preview_text.winfo_exists():
            # If we're called before _display_step4, don't try to update the preview
            return

        # Update preview text
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete(1.0, tk.END)

        # Convert to YAML for display
        try:
            import yaml
            config_yaml = yaml.dump(config, default_flow_style=False)
            self.preview_text.insert(tk.END, config_yaml)
        except Exception as e:
            self.preview_text.insert(tk.END, f"Error generating preview: {str(e)}\n\n")
            self.preview_text.insert(tk.END, str(config))

        self.preview_text.config(state=tk.DISABLED)

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

            # Call the callback if provided to refresh UI in other tabs
            if callable(self.on_config_saved):
                self.on_config_saved()

            # Show a helpful message about the new configuration being available
            messagebox.showinfo(
                "Configuration Available",
                f"The new configuration 'QA-{analytics_id}' is now available in the Run Analytics tab."
            )
        else:
            messagebox.showerror("Error", result)

    def _go_next_step(self):
        """Navigate to the next wizard step"""
        # Validate before moving to next step
        if not self._validate_current_step():
            return

        # Move to next step if not at the end
        if self.current_step < 3:  # 3 is the last step (0-indexed)
            self.current_step += 1
            self.step_tracker.set_current_step(self.current_step)
            self._display_current_step()
            self._update_button_states()

    def _go_prev_step(self):
        """Navigate to the previous wizard step"""
        if self.current_step > 0:
            self.current_step -= 1
            self.step_tracker.set_current_step(self.current_step)
            self._display_current_step()
            self._update_button_states()

    def _update_button_states(self):
        """Update the state of navigation buttons based on current step"""
        # Update Previous button
        if self.current_step == 0:
            self.prev_btn.config(state=tk.DISABLED)
        else:
            self.prev_btn.config(state=tk.NORMAL)

        # Update Next button
        if self.current_step == 3:  # Last step
            self.next_btn.config(state=tk.DISABLED)
        else:
            self.next_btn.config(state=tk.NORMAL)

    def _validate_current_step(self):
        """Validate the current step before allowing navigation to the next step"""
        if self.current_step == 0:  # Template Selection
            if not self.current_template_id:
                messagebox.showinfo("Select Template", "Please select a template to continue")
                return False
            return True

        elif self.current_step == 1:  # Basic Configuration
            # Validate required fields
            if not self.analytics_id_var.get().strip():
                messagebox.showinfo("Required Field", "Analytics ID is required")
                return False

            if not self.analytics_name_var.get().strip():
                messagebox.showinfo("Required Field", "Analytics Name is required")
                return False

            return True

        elif self.current_step == 2:  # Parameters
            # Only try to validate parameters if params_container exists
            if hasattr(self, 'params_container'):
                # Make sure all required parameters are filled
                for param in self.template_parameters:
                    if param.get('required', False):
                        param_name = param.get('name', '')
                        if param_name in self.parameter_entries:
                            value = self.parameter_entries[param_name].get()
                            if not value:
                                messagebox.showinfo("Required Parameter", f"Parameter '{param_name}' is required")
                                return False

            # We'll generate the preview when step 4 is displayed, not here
            # This prevents the error when preview_text doesn't exist yet
            return True

        return True  # Default to allow navigation

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
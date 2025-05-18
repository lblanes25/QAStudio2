"""
Enhanced ConfigWizardTab with Excel Formula Integration.

This module provides an enhanced version of the ConfigWizardTab with full
integration of the Excel formula testing capabilities.
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox
import logging
import yaml
from typing import Callable, Dict, List, Optional, Any

from qa_analytics.templates.template_manager import TemplateManager
from qa_analytics.utils.step_tracker import StepTracker
from qa_analytics.ui.components.formula_tester import FormulaTester
from qa_analytics.core.excel_utils import is_valid_excel_formula, extract_column_names

# Set up logging
logger = logging.getLogger("qa_analytics")


class ConfigWizardTab(ttk.Frame):
    """
    Enhanced tab for creating and editing analytics configurations with
    integrated Excel formula testing capabilities.
    """

    def __init__(self, parent, config_manager, template_manager=None, on_config_saved=None):
        """
        Initialize the Configuration Wizard tab with formula testing.

        Args:
            parent: Parent widget
            config_manager: ConfigManager instance for loading/saving configs
            template_manager: Optional TemplateManager instance
            on_config_saved: Optional callback for when config is saved
        """
        super().__init__(parent, padding="20 15 20 15")
        self.parent = parent
        self.config_manager = config_manager
        self.template_manager = template_manager or TemplateManager()
        self.on_config_saved = on_config_saved  # Callback when config is saved

        # Define wizard steps
        self.steps = ["Select Template", "Basic Settings", "Define Validations", "Review & Save"]
        self.current_step = 0

        # Initialize state variables
        self.source_var = tk.StringVar()
        self.data_source_name_var = tk.StringVar()
        self.file_type_var = tk.StringVar(value="XLSX")
        self.column_mapping_var = tk.StringVar()
        self.validation_rules_var = tk.StringVar()
        self.analytics_id_var = tk.StringVar()
        self.analytics_name_var = tk.StringVar()
        self.analytics_desc_var = tk.StringVar()
        self.threshold_var = tk.StringVar(value="5.0")
        self.group_by_var = tk.StringVar()
        self.current_template_id = None
        self.template_parameters = []
        self.parameter_entries = {}
        self.current_config = None

        # Formula validation state
        self.formula_is_valid = False
        self.formula_fields = set()
        self.formula_tester = None  # Will hold the FormulaTester component

        # Set up UI components
        self._create_widgets()

    def _create_widgets(self):
        """Create all widgets for this tab"""
        # Configure grid layout
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)  # Content area should expand

        # Add StepTracker at the top
        self.step_tracker = StepTracker(self, self.steps, initial_step=self.current_step)
        self.step_tracker.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))

        # Create a card-like container for step content
        self.content_card = ttk.Frame(self, style="Card.TFrame")
        self.content_card.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=2, pady=2)
        self.content_card.columnconfigure(0, weight=1)
        self.content_card.rowconfigure(0, weight=1)

        # Container for step content with padding
        self.content_frame = ttk.Frame(self.content_card, padding=15)
        self.content_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.rowconfigure(0, weight=1)

        # Button frame at bottom
        button_frame = ttk.Frame(self)
        button_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(20, 0))

        self.prev_btn = ttk.Button(
            button_frame,
            text="Previous",
            command=self._go_previous
        )
        self.prev_btn.pack(side=tk.LEFT)

        self.next_btn = ttk.Button(
            button_frame,
            text="Next",
            style="Primary.TButton",
            command=self._go_next
        )
        self.next_btn.pack(side=tk.RIGHT)

        # Display the initial step content
        self._display_current_step()

        # Update button states
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
            self._display_step3()  # Validation Rules
        elif self.current_step == 3:
            self._display_step4()  # Review & Save

    # Path: qa_analytics/tabs/config_wizard_tab.py
    # Add to the ConfigWizardTab class

    def _display_step1(self):
        """Display Step 1: Template Selection with Create New option"""
        # Create template selection container
        step_frame = ttk.Frame(self.content_frame)
        step_frame.pack(fill=tk.BOTH, expand=True)
        step_frame.columnconfigure(0, weight=1)
        step_frame.rowconfigure(1, weight=1)  # Tree should expand

        # Step title
        title_frame = ttk.Frame(step_frame)
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        title_frame.columnconfigure(1, weight=1)  # Push button to right

        ttk.Label(
            title_frame,
            text="Select a Template",
            style="Header.TLabel"
        ).grid(row=0, column=0, sticky=tk.W)

        # Add Create New Template button
        create_btn = ttk.Button(
            title_frame,
            text="Create New Template",
            command=self._create_new_template,
            style="Secondary.TButton"
        )
        create_btn.grid(row=0, column=1, sticky=tk.E)

        # Filter frame
        filter_frame = ttk.Frame(step_frame)
        filter_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Label(filter_frame, text="Filter by Category:").pack(side=tk.LEFT)

        # Get categories from template manager
        categories = self.template_manager.get_template_categories()
        category_names = ["All Categories"] + [c.get('name', '') for c in categories]

        self.category_var = tk.StringVar(value="All Categories")
        category_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.category_var,
            values=category_names,
            state="readonly",
            width=20
        )
        category_combo.pack(side=tk.LEFT, padx=(8, 0))
        category_combo.bind("<<ComboboxSelected>>", lambda e: self._populate_template_tree())

        # Create container for treeview and scrollbar
        tree_container = ttk.Frame(step_frame)
        tree_container.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        # Create template treeview
        columns = ("Name", "Category", "Description", "Parameters", "Difficulty")
        self.template_tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show="headings",
            height=10
        )

        # Configure columns
        self.template_tree.column("Name", width=150)
        self.template_tree.column("Category", width=100)
        self.template_tree.column("Description", width=250)
        self.template_tree.column("Parameters", width=80, anchor=tk.CENTER)
        self.template_tree.column("Difficulty", width=80, anchor=tk.CENTER)

        # Configure headings
        for col in columns:
            self.template_tree.heading(col, text=col)

        # Add scrollbars
        y_scrollbar = ttk.Scrollbar(
            tree_container,
            orient="vertical",
            command=self.template_tree.yview,
            style="Vertical.TScrollbar"
        )
        self.template_tree.configure(yscrollcommand=y_scrollbar.set)

        # Pack tree and scrollbar
        self.template_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        y_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Bind selection event
        self.template_tree.bind("<<TreeviewSelect>>", self._on_template_selected)

        # Template details section
        details_frame = ttk.LabelFrame(step_frame, text="Template Details", padding=10)
        details_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        details_frame.columnconfigure(0, weight=1)

        # Template details text
        details_container = ttk.Frame(details_frame)
        details_container.pack(fill=tk.BOTH, expand=True)
        details_container.columnconfigure(0, weight=1)
        details_container.rowconfigure(0, weight=1)

        self.details_text = tk.Text(
            details_container,
            wrap=tk.WORD,
            height=6,
            background="#F9F9F9",
            relief=tk.FLAT,
            padx=10,
            pady=10,
            borderwidth=0
        )
        self.details_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Add scrollbar for details
        details_scroll = ttk.Scrollbar(
            details_container,
            orient="vertical",
            command=self.details_text.yview,
            style="Vertical.TScrollbar"
        )
        details_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.details_text.config(yscrollcommand=details_scroll.set)

        # Make details text read-only
        self.details_text.config(state=tk.DISABLED)

        # Populate the template tree
        self._populate_template_tree()

    def _create_new_template(self):
        """Open a dialog to create a new template"""
        # Create template creation dialog
        dialog = tk.Toplevel(self)
        dialog.title("Create New Template")
        dialog.geometry("700x600")
        dialog.transient(self)  # Set to be on top of the parent window
        dialog.grab_set()  # Modal dialog

        # Make dialog appear in center of parent window
        x = self.winfo_rootx() + (self.winfo_width() // 2) - (700 // 2)
        y = self.winfo_rooty() + (self.winfo_height() // 2) - (600 // 2)
        dialog.geometry(f"+{x}+{y}")

        # Apply padding to dialog content
        content_frame = ttk.Frame(dialog, padding=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(1, weight=1)

        # Dialog title
        ttk.Label(
            content_frame,
            text="Create New Template",
            style="Header.TLabel"
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 20))

        # Basic template info section
        info_frame = ttk.LabelFrame(content_frame, text="Template Information", padding=10)
        info_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        info_frame.columnconfigure(1, weight=1)

        # Template ID
        ttk.Label(info_frame, text="Template ID:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 10))
        template_id_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=template_id_var, width=30).grid(row=0, column=1, sticky=(tk.W, tk.E),
                                                                           pady=(0, 10))

        # Template Name
        ttk.Label(info_frame, text="Template Name:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 10))
        template_name_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=template_name_var, width=40).grid(row=1, column=1, sticky=(tk.W, tk.E),
                                                                             pady=(0, 10))

        # Template Description
        ttk.Label(info_frame, text="Description:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 10))
        template_desc_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=template_desc_var, width=60).grid(row=2, column=1, sticky=(tk.W, tk.E),
                                                                             pady=(0, 10))

        # Template Category
        ttk.Label(info_frame, text="Category:").grid(row=3, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 10))

        # Get categories from template manager
        categories = self.template_manager.get_template_categories()
        category_names = [c.get('name', '') for c in categories]

        template_category_var = tk.StringVar()
        if category_names:
            template_category_var.set(category_names[0])

        ttk.Combobox(
            info_frame,
            textvariable=template_category_var,
            values=category_names,
            state="readonly"
        ).grid(row=3, column=1, sticky=tk.W, pady=(0, 10))

        # Create buttons section
        buttons_frame = ttk.Frame(content_frame)
        buttons_frame.grid(row=5, column=0, columnspan=2, sticky=tk.E, pady=(20, 0))

        ttk.Button(
            buttons_frame,
            text="Cancel",
            command=dialog.destroy
        ).pack(side=tk.RIGHT, padx=(10, 0))

        ttk.Button(
            buttons_frame,
            text="Start with Blank Template",
            command=lambda: self._create_blank_template(
                dialog,
                template_id_var.get(),
                template_name_var.get(),
                template_desc_var.get(),
                template_category_var.get()
            ),
            style="Secondary.TButton"
        ).pack(side=tk.RIGHT, padx=(10, 0))

        ttk.Button(
            buttons_frame,
            text="Clone Selected Template",
            command=lambda: self._clone_template(
                dialog,
                template_id_var.get(),
                template_name_var.get(),
                template_desc_var.get(),
                template_category_var.get()
            ),
            style="Primary.TButton"
        ).pack(side=tk.RIGHT)

    def _display_step2(self):
        """Display Step 2: Basic Configuration"""
        step_frame = ttk.Frame(self.content_frame)
        step_frame.pack(fill=tk.BOTH, expand=True)
        step_frame.columnconfigure(1, weight=1)

        # Step title
        ttk.Label(
            step_frame,
            text="Basic Configuration",
            style="Header.TLabel"
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 20))

        # Form layout with consistent spacing
        # Analytics ID
        ttk.Label(
            step_frame,
            text="Analytics ID*:"
        ).grid(row=1, column=0, sticky=tk.W, padx=(0, 15), pady=(0, 12))

        id_frame = ttk.Frame(step_frame)
        id_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 12))

        ttk.Entry(
            id_frame,
            textvariable=self.analytics_id_var,
            width=10
        ).pack(side=tk.LEFT)

        ttk.Label(
            id_frame,
            text="(Required - numeric identifier for this analytic)",
            style="Small.TLabel"
        ).pack(side=tk.LEFT, padx=(10, 0))

        # Analytics Name
        ttk.Label(
            step_frame,
            text="Analytics Name*:"
        ).grid(row=2, column=0, sticky=tk.W, padx=(0, 15), pady=(0, 12))

        ttk.Entry(
            step_frame,
            textvariable=self.analytics_name_var,
            width=40
        ).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=(0, 12))

        # Description
        ttk.Label(
            step_frame,
            text="Description:"
        ).grid(row=3, column=0, sticky=tk.W, padx=(0, 15), pady=(0, 12))

        ttk.Entry(
            step_frame,
            textvariable=self.analytics_desc_var,
            width=60
        ).grid(row=3, column=1, sticky=(tk.W, tk.E), pady=(0, 12))

        # Data Source
        ttk.Label(
            step_frame,
            text="Data Source:"
        ).grid(row=4, column=0, sticky=tk.W, padx=(0, 15), pady=(0, 12))

        # Get data sources from config
        try:
            from qa_analytics.core.data_source_manager import DataSourceManager
            data_source_manager = DataSourceManager()
            data_sources = list(data_source_manager.registry.keys())
        except:
            data_sources = []

        data_source_combo = ttk.Combobox(
            step_frame,
            textvariable=self.data_source_name_var,
            values=data_sources,
            width=40
        )
        data_source_combo.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=(0, 12))

        # Error Threshold
        ttk.Label(
            step_frame,
            text="Error Threshold %:"
        ).grid(row=5, column=0, sticky=tk.W, padx=(0, 15), pady=(0, 12))

        threshold_frame = ttk.Frame(step_frame)
        threshold_frame.grid(row=5, column=1, sticky=(tk.W, tk.E), pady=(0, 12))

        ttk.Entry(
            threshold_frame,
            textvariable=self.threshold_var,
            width=10
        ).pack(side=tk.LEFT)

        ttk.Label(
            threshold_frame,
            text="(Maximum acceptable error percentage)",
            style="Small.TLabel"
        ).pack(side=tk.LEFT, padx=(10, 0))

        # Group By Field
        ttk.Label(
            step_frame,
            text="Group By Field:"
        ).grid(row=6, column=0, sticky=tk.W, padx=(0, 15), pady=(0, 12))

        ttk.Entry(
            step_frame,
            textvariable=self.group_by_var,
            width=40
        ).grid(row=6, column=1, sticky=(tk.W, tk.E), pady=(0, 12))

        # Information box
        info_frame = ttk.Frame(step_frame, style="Card.TFrame")
        info_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(20, 0))

        info_label = ttk.Label(
            info_frame,
            text=(
                "Complete the basic configuration for your analytics:\n\n"
                "• Analytics ID: Unique identifier for this analytic (required)\n"
                "• Analytics Name: Descriptive name for the analytic\n"
                "• Description: Explanation of what this analytic validates\n"
                "• Data Source: The registered data source to use\n"
                "• Error Threshold: Maximum acceptable percentage of non-conforming records\n"
                "• Group By Field: Field used to group results in reports"
            ),
            style="Info.TLabel",
            wraplength=550,
            justify=tk.LEFT
        )
        info_label.pack(padx=15, pady=15, fill=tk.X)

    def _display_step3(self):
        """Display Step 3: Validation Rules"""
        # Create scrollable canvas for validation rules
        canvas_container = ttk.Frame(self.content_frame)
        canvas_container.pack(fill=tk.BOTH, expand=True)
        canvas_container.columnconfigure(0, weight=1)
        canvas_container.rowconfigure(0, weight=1)

        # Create canvas and scrollbar
        canvas = tk.Canvas(
            canvas_container,
            background="#FFFFFF",
            highlightthickness=0,
            borderwidth=0
        )
        scrollbar = ttk.Scrollbar(
            canvas_container,
            orient="vertical",
            command=canvas.yview,
            style="Vertical.TScrollbar"
        )

        # Configure canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Create a frame inside the canvas for all content
        validation_frame = ttk.Frame(canvas, padding=5)
        validation_window = canvas.create_window((0, 0), window=validation_frame, anchor="nw")
        validation_frame.columnconfigure(0, weight=1)

        # Step title
        ttk.Label(
            validation_frame,
            text="Define Validation Rules",
            style="Header.TLabel"
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 15))

        # Description
        ttk.Label(
            validation_frame,
            text="Specify the rules that determine whether records conform to requirements.",
            wraplength=600
        ).grid(row=1, column=0, sticky=tk.W, pady=(0, 20))

        # Rule Type Selection
        rule_frame = ttk.LabelFrame(validation_frame, text="Rule Type", padding=10)
        rule_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15))

        self.rule_type_var = tk.StringVar(value="excel_formula")

        ttk.Radiobutton(
            rule_frame,
            text="Excel Formula",
            variable=self.rule_type_var,
            value="excel_formula",
            command=self._update_rule_section
        ).pack(anchor=tk.W, padx=10, pady=8)

        ttk.Radiobutton(
            rule_frame,
            text="Segregation of Duties",
            variable=self.rule_type_var,
            value="segregation",
            command=self._update_rule_section
        ).pack(anchor=tk.W, padx=10, pady=8)

        ttk.Radiobutton(
            rule_frame,
            text="Approval Sequence",
            variable=self.rule_type_var,
            value="approval",
            command=self._update_rule_section
        ).pack(anchor=tk.W, padx=10, pady=8)

        # Rule configuration section (will be populated based on selection)
        self.rule_config_frame = ttk.LabelFrame(
            validation_frame,
            text="Rule Configuration",
            padding=10
        )
        self.rule_config_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # Initialize rule configuration
        self._update_rule_section()

        # Update scroll region when frame size changes
        def update_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(validation_window, width=event.width)

        validation_frame.bind("<Configure>", update_scroll_region)

    def _update_rule_section(self):
        """Update the rule configuration section based on the selected rule type"""
        # Clear current content
        for widget in self.rule_config_frame.winfo_children():
            widget.destroy()

        rule_type = self.rule_type_var.get()

        if rule_type == "excel_formula":
            # Excel Formula configuration
            ttk.Label(
                self.rule_config_frame,
                text="Use an Excel-style formula that evaluates to TRUE for records that conform:",
                wraplength=550
            ).pack(anchor=tk.W, padx=10, pady=(10, 15))

            # Add the FormulaTester component
            self.formula_tester = FormulaTester(
                self.rule_config_frame,
                callback=self._handle_formula_change,
                initial_formula=getattr(self, 'formula_var', tk.StringVar()).get() if hasattr(self, 'formula_var') else "",
                description=getattr(self, 'display_name_var', tk.StringVar()).get() if hasattr(self, 'display_name_var') else "Custom Validation"
            )
            self.formula_tester.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        elif rule_type == "segregation":
            # Segregation of Duties configuration
            ttk.Label(
                self.rule_config_frame,
                text="Specify fields that should contain different users:",
                wraplength=550
            ).pack(anchor=tk.W, padx=10, pady=(10, 15))

            # Submitter field
            field_frame1 = ttk.Frame(self.rule_config_frame)
            field_frame1.pack(fill=tk.X, padx=10, pady=(0, 10))

            ttk.Label(
                field_frame1,
                text="Submitter Field:",
                width=15
            ).pack(side=tk.LEFT)

            self.submitter_var = tk.StringVar()
            ttk.Entry(
                field_frame1,
                textvariable=self.submitter_var,
                width=30
            ).pack(side=tk.LEFT, fill=tk.X, expand=True)

            # Approver field
            field_frame2 = ttk.Frame(self.rule_config_frame)
            field_frame2.pack(fill=tk.X, padx=10, pady=(0, 10))

            ttk.Label(
                field_frame2,
                text="Approver Field:",
                width=15
            ).pack(side=tk.LEFT)

            self.approver_var = tk.StringVar()
            ttk.Entry(
                field_frame2,
                textvariable=self.approver_var,
                width=30
            ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        elif rule_type == "approval":
            # Approval Sequence configuration
            ttk.Label(
                self.rule_config_frame,
                text="Specify date fields that should be in chronological sequence:",
                wraplength=550
            ).pack(anchor=tk.W, padx=10, pady=(10, 15))

            # Submit date field
            field_frame1 = ttk.Frame(self.rule_config_frame)
            field_frame1.pack(fill=tk.X, padx=10, pady=(0, 10))

            ttk.Label(
                field_frame1,
                text="Submit Date Field:",
                width=20
            ).pack(side=tk.LEFT)

            self.submit_date_var = tk.StringVar()
            ttk.Entry(
                field_frame1,
                textvariable=self.submit_date_var,
                width=30
            ).pack(side=tk.LEFT, fill=tk.X, expand=True)

            # Approval date field
            field_frame2 = ttk.Frame(self.rule_config_frame)
            field_frame2.pack(fill=tk.X, padx=10, pady=(0, 10))

            ttk.Label(
                field_frame2,
                text="Approval Date Field:",
                width=20
            ).pack(side=tk.LEFT)

            self.approval_date_var = tk.StringVar()
            ttk.Entry(
                field_frame2,
                textvariable=self.approval_date_var,
                width=30
            ).pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _handle_formula_change(self, formula: str, display_name: str, is_valid: bool, fields: set):
        """
        Handle formula changes from the FormulaTester component

        Args:
            formula: Excel formula
            display_name: Display name for the formula
            is_valid: Whether the formula is valid
            fields: Fields used in the formula
        """
        # Store formula information for later use
        self.formula_is_valid = is_valid
        self.formula_fields = fields

        # If we have parameter entries, update them
        for param in self.template_parameters:
            if param.get('name') in ['formula', 'original_formula'] and param.get('name') in self.parameter_entries:
                self.parameter_entries[param.get('name')].set(formula)

        logger.debug(f"Formula updated: {formula}, valid: {is_valid}, fields: {fields}")

    def _display_step4(self):
        """Display Step 4: Review & Save"""
        step_frame = ttk.Frame(self.content_frame)
        step_frame.pack(fill=tk.BOTH, expand=True)
        step_frame.columnconfigure(0, weight=1)
        step_frame.rowconfigure(2, weight=1)  # Preview should expand

        # Step title
        ttk.Label(
            step_frame,
            text="Review and Save Configuration",
            style="Header.TLabel"
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 15))

        # Description
        ttk.Label(
            step_frame,
            text="Review the generated configuration before saving it. You can make adjustments by returning to previous steps.",
            wraplength=600
        ).grid(row=1, column=0, sticky=tk.W, pady=(0, 15))

        # Preview pane with card-like styling
        preview_frame = ttk.LabelFrame(
            step_frame,
            text="Configuration Preview",
            padding=5
        )
        preview_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        # Preview container for text and scrollbar
        preview_container = ttk.Frame(preview_frame)
        preview_container.pack(fill=tk.BOTH, expand=True)
        preview_container.columnconfigure(0, weight=1)
        preview_container.rowconfigure(0, weight=1)

        self.preview_text = tk.Text(
            preview_container,
            wrap=tk.WORD,
            font=("Consolas", 10),
            background="#F9F9F9",
            relief=tk.FLAT,
            padx=10,
            pady=10,
            borderwidth=0
        )
        self.preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Add scrollbar for preview
        preview_scroll = ttk.Scrollbar(
            preview_container,
            orient="vertical",
            command=self.preview_text.yview,
            style="Vertical.TScrollbar"
        )
        preview_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.preview_text.config(yscrollcommand=preview_scroll.set)

        # Action buttons
        action_frame = ttk.Frame(step_frame)
        action_frame.grid(row=3, column=0, sticky=tk.E, pady=(10, 0))

        refresh_btn = ttk.Button(
            action_frame,
            text="Refresh Preview",
            command=self._refresh_preview
        )
        refresh_btn.pack(side=tk.LEFT, padx=(0, 10))

        save_btn = ttk.Button(
            action_frame,
            text="Save Configuration",
            style="Primary.TButton",
            command=self._save_configuration
        )
        save_btn.pack(side=tk.LEFT)

        # Refresh the preview when entering this step
        self._refresh_preview()

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

    def _go_next(self):
        """Navigate to the next wizard step"""
        # Validate before moving to next step
        if not self._validate_current_step():
            return

        # Move to next step if not at the end
        if self.current_step < len(self.steps) - 1:
            self.current_step += 1
            self.step_tracker.set_current_step(self.current_step)
            self._display_current_step()
            self._update_button_states()

    def _go_previous(self):
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
        if self.current_step == len(self.steps) - 1:
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

        elif self.current_step == 2:  # Validation Rules
            # Validate based on rule type
            rule_type = self.rule_type_var.get()

            if rule_type == "excel_formula":
                # Check formula tester if available
                if hasattr(self, 'formula_tester') and self.formula_tester:
                    formula = self.formula_tester.get_formula()
                    if not formula:
                        messagebox.showinfo("Required Field", "Please enter a formula")
                        return False

                    if not self.formula_tester.is_valid():
                        if messagebox.askyesno("Invalid Formula",
                                         "The formula appears to be invalid. Continue anyway?"):
                            return True
                        return False
                else:
                    # Fall back to formula_var if available
                    if hasattr(self, 'formula_var') and not self.formula_var.get().strip():
                        messagebox.showinfo("Required Field", "Please enter a formula")
                        return False

            elif rule_type == "segregation":
                if not hasattr(self, 'submitter_var') or not self.submitter_var.get().strip():
                    messagebox.showinfo("Required Field", "Submitter Field is required")
                    return False
                if not hasattr(self, 'approver_var') or not self.approver_var.get().strip():
                    messagebox.showinfo("Required Field", "Approver Field is required")
                    return False

            elif rule_type == "approval":
                if not hasattr(self, 'submit_date_var') or not self.submit_date_var.get().strip():
                    messagebox.showinfo("Required Field", "Submit Date Field is required")
                    return False
                if not hasattr(self, 'approval_date_var') or not self.approval_date_var.get().strip():
                    messagebox.showinfo("Required Field", "Approval Date Field is required")
                    return False

            return True

        return True  # Default to allow navigation

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
            config_yaml = yaml.dump(config, default_flow_style=False, sort_keys=False)
            self.preview_text.insert(tk.END, config_yaml)
        except Exception as e:
            self.preview_text.insert(tk.END, f"Error generating preview: {str(e)}\n\n")
            self.preview_text.insert(tk.END, str(config))

        self.preview_text.config(state=tk.DISABLED)

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

        # Process rule type specific configurations
        rule_type = self.rule_type_var.get()

        if rule_type == "excel_formula":
            # Add custom formula validation
            if 'validations' not in config:
                config['validations'] = []

            # Get formula and display name from the FormulaTester component
            if hasattr(self, 'formula_tester') and self.formula_tester:
                formula = self.formula_tester.get_formula()
                display_name = self.formula_tester.get_display_name()
                fields_used = list(self.formula_tester.get_fields_used())
            else:
                # Fallback to formula_var if available
                formula = getattr(self, 'formula_var', tk.StringVar()).get()
                display_name = getattr(self, 'display_name_var', tk.StringVar()).get() or "Custom Validation"
                fields_used = []

            config['validations'].append({
                'rule': 'custom_formula',
                'description': display_name or 'User-defined Excel formula validation',
                'parameters': {
                    'original_formula': formula,
                    'display_name': display_name
                },
                'metadata': {
                    'fields_used': fields_used
                }
            })

            # Ensure these fields are added to required_fields
            if fields_used and 'data_source' in config:
                if 'required_fields' not in config['data_source']:
                    config['data_source']['required_fields'] = []

                for field in fields_used:
                    if field not in config['data_source']['required_fields']:
                        config['data_source']['required_fields'].append(field)

        elif rule_type == "segregation":
            # Add segregation of duties validation
            if 'validations' not in config:
                config['validations'] = []

            config['validations'].append({
                'rule': 'segregation_of_duties',
                'description': 'Validates proper segregation of duties',
                'parameters': {
                    'submitter_field': self.submitter_var.get(),
                    'approver_fields': [self.approver_var.get()]
                }
            })

            # Add fields to required_fields
            if 'data_source' in config:
                if 'required_fields' not in config['data_source']:
                    config['data_source']['required_fields'] = []

                fields_to_add = [self.submitter_var.get(), self.approver_var.get()]
                for field in fields_to_add:
                    if field and field not in config['data_source']['required_fields']:
                        config['data_source']['required_fields'].append(field)

        elif rule_type == "approval":
            # Add approval sequence validation
            if 'validations' not in config:
                config['validations'] = []

            config['validations'].append({
                'rule': 'approval_sequence',
                'description': 'Validates that approvals happened in correct sequence',
                'parameters': {
                    'date_fields_in_order': [
                        self.submit_date_var.get(),
                        self.approval_date_var.get()
                    ]
                }
            })

            # Add fields to required_fields
            if 'data_source' in config:
                if 'required_fields' not in config['data_source']:
                    config['data_source']['required_fields'] = []

                fields_to_add = [self.submit_date_var.get(), self.approval_date_var.get()]
                for field in fields_to_add:
                    if field and field not in config['data_source']['required_fields']:
                        config['data_source']['required_fields'].append(field)

        return config

    def _get_parameter_values(self):
        """Get values from all parameter fields"""
        values = {}

        # Add basic fields
        values['analytic_id'] = self.analytics_id_var.get()
        values['analytic_name'] = self.analytics_name_var.get()
        values['analytic_description'] = self.analytics_desc_var.get()
        values['data_source'] = self.data_source_name_var.get()
        values['threshold_percentage'] = self.threshold_var.get()
        values['group_by'] = self.group_by_var.get()

        # Add template parameters
        for name, var in self.parameter_entries.items():
            values[name] = var.get()

        return values

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

            # Reload configurations in the config manager - ensure this is done correctly
            if hasattr(self.config_manager, 'load_all_configs'):
                logger.info("Reloading configurations after save")
                self.config_manager.load_all_configs()

                # Verify if configurations were loaded correctly
                config_ids = list(self.config_manager.configs.keys())
                logger.info(f"Loaded configurations: {config_ids}")

            # Call the callback if provided to refresh UI in other tabs
            if callable(self.on_config_saved):
                logger.info("Calling on_config_saved callback")
                self.on_config_saved()

            # Show a helpful message about the new configuration being available
            messagebox.showinfo(
                "Configuration Available",
                f"The new configuration 'QA-{analytics_id}' is now available in the Run Analytics tab."
            )
        else:
            messagebox.showerror("Error", result)

    def cleanup(self):
        """Clean up resources, especially the Excel processor in the FormulaTester"""
        if hasattr(self, 'formula_tester') and self.formula_tester:
            self.formula_tester.cleanup()

    def __del__(self):
        """Destructor to ensure resources are cleaned up"""
        self.cleanup()

    def debug_initialization(self):
        """Debug initialization issues with the tab"""
        try:
            # Check if template manager is properly initialized
            templates = self.template_manager.get_all_templates()
            print(f"Found {len(templates)} templates")

            # Check if step tracking is working
            print(f"Current step: {self.current_step}")

            # Force display of step 1
            self._display_step1()

            # Log widget hierarchy to see what's created
            self._print_widget_hierarchy(self)

        except Exception as e:
            import traceback
            print(f"Error during initialization: {e}")
            print(traceback.format_exc())

    def _print_widget_hierarchy(self, widget, level=0):
        """Print the widget hierarchy to help debug UI issues"""
        print(" " * level + f"Widget: {widget} ({widget.winfo_class()})")
        try:
            children = widget.winfo_children()
            for child in children:
                self._print_widget_hierarchy(child, level + 2)
        except:
            pass

    def _create_blank_template(self, dialog, template_id, template_name, template_desc, template_category):
        """Create a blank template with basic structure"""
        # Validate inputs
        if not template_id or not template_name:
            messagebox.showinfo("Missing Information", "Template ID and Name are required")
            return

        # Create a new blank template
        template = {
            'template_id': template_id,
            'template_name': template_name,
            'template_description': template_desc,
            'template_category': template_category,
            'template_version': '1.0',
            'template_parameters': [
                {
                    'name': 'analytic_id',
                    'description': 'Unique identifier for this analytic',
                    'data_type': 'string',
                    'required': True,
                    'example': ''
                },
                {
                    'name': 'analytic_name',
                    'description': 'Descriptive name for this analytic',
                    'data_type': 'string',
                    'required': True,
                    'example': ''
                },
                {
                    'name': 'data_source',
                    'description': 'Data source containing the data',
                    'data_type': 'data_source',
                    'required': True,
                    'example': ''
                },
                {
                    'name': 'group_by',
                    'description': 'Field to group results by',
                    'data_type': 'string',
                    'required': True,
                    'example': ''
                },
                {
                    'name': 'threshold_percentage',
                    'description': 'Maximum acceptable error percentage',
                    'data_type': 'number',
                    'required': True,
                    'example': '5.0'
                }
            ],
            'generated_validations': [],
            'default_thresholds': {
                'error_percentage': 5.0,
                'rationale': 'Standard error threshold.'
            },
            'default_reporting': {
                'group_by': '{group_by}',
                'summary_fields': ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage'],
                'detail_required': True
            }
        }

        # Save the template
        self._save_new_template(dialog, template)

    def _clone_template(self, dialog, template_id, template_name, template_desc, template_category):
        """Clone an existing template with new information"""
        # Validate inputs
        if not template_id or not template_name:
            messagebox.showinfo("Missing Information", "Template ID and Name are required")
            return

        # Check if a template is selected
        selection = self.template_tree.selection()
        if not selection:
            messagebox.showinfo("No Selection", "Please select a template to clone")
            return

        # Get the selected template ID
        selected_template_id = selection[0]

        # Get the source template
        source_template = self.template_manager.get_template(selected_template_id)
        if not source_template:
            messagebox.showinfo("Error", "Could not load the selected template")
            return

        # Clone the template with new information
        template = source_template.copy()
        template['template_id'] = template_id
        template['template_name'] = template_name
        template['template_description'] = template_desc if template_desc else template['template_description']
        template['template_category'] = template_category if template_category else template['template_category']

        # Save the template
        self._save_new_template(dialog, template)

    def _save_new_template(self, dialog, template):
        """Save a new template to the templates directory"""
        try:
            # Ensure templates directory exists
            if not os.path.exists(self.template_manager.templates_dir):
                os.makedirs(self.template_manager.templates_dir)

            # Save template file
            template_path = os.path.join(
                self.template_manager.templates_dir,
                f"{template['template_id']}.yaml"
            )

            with open(template_path, 'w', encoding='utf-8') as f:
                yaml.dump(template, f, default_flow_style=False)

            # Update metadata to include the new template
            metadata_path = os.path.join(self.template_manager.templates_dir, 'metadata.yaml')
            if os.path.exists(metadata_path):
                try:
                    with open(metadata_path, 'r', encoding='utf-8') as f:
                        metadata = yaml.safe_load(f)

                    # Add template to metadata if not already present
                    if 'templates' not in metadata:
                        metadata['templates'] = {}

                    metadata['templates'][template['template_id']] = {
                        'suitable_for': [],
                        'difficulty': 'Medium',
                        'validation_rules': []
                    }

                    with open(metadata_path, 'w', encoding='utf-8') as f:
                        yaml.dump(metadata, f, default_flow_style=False)
                except Exception as e:
                    logger.warning(f"Error updating metadata: {e}")

            # Reload templates
            self.template_manager._load_templates()
            self.template_manager._load_metadata()

            # Update the template tree
            self._populate_template_tree()

            # Select the new template
            for item in self.template_tree.get_children():
                if item == template['template_id']:
                    self.template_tree.selection_set(item)
                    self.template_tree.focus(item)
                    self.template_tree.see(item)
                    self._on_template_selected(None)
                    break

            # Close the dialog
            dialog.destroy()

            # Show success message
            messagebox.showinfo(
                "Template Created",
                f"Template '{template['template_name']}' created successfully."
            )

        except Exception as e:
            logger.error(f"Error saving template: {e}")
            messagebox.showerror("Error", f"Failed to save template: {str(e)}")
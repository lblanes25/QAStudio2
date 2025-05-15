import os
import sys
import json
import logging
import threading
import tkinter as tk
import pandas as pd
from tkinter import filedialog, messagebox, ttk
import datetime

from typing import Dict, List

from config_manager import ConfigManager
from enhanced_data_processor import EnhancedDataProcessor
from enhanced_report_generator import EnhancedReportGenerator
from reference_data_manager import ReferenceDataManager
from data_source_manager import DataSourceManager
from template_manager import TemplateManager
from logging_config import setup_logging

# Import new components
from config_wizard import ConfigWizard
from testing_environment import TestingEnvironment
from automation_scheduler import AutomationScheduler, SchedulerUI

logger = setup_logging()

class EnhancedQAAnalyticsApp:
    """Enhanced application with GUI interface and simplified, progressive UI"""

    def __init__(self, root):
        """Initialize the application"""
        self.root = root
        self.root.title("Enhanced QA Analytics Automation")
        self.root.geometry("900x700")

        # Load user preferences
        self.user_preferences = self._load_user_preferences()
        self.is_advanced_mode = self.user_preferences.get("advanced_mode", False)

        # Load configuration
        self.config_manager = ConfigManager()
        self.available_analytics = self.config_manager.get_available_analytics()

        # Initialize managers
        self.reference_data_manager = ReferenceDataManager()
        self.data_source_manager = DataSourceManager()
        self.template_manager = TemplateManager()

        # Initialize scheduler
        self.scheduler = AutomationScheduler(
            config_manager=self.config_manager,
            data_processor_class=EnhancedDataProcessor,
            report_generator_class=EnhancedReportGenerator
        )

        # Set up UI components
        self._setup_ui()

    def _load_user_preferences(self):
        """Load saved user preferences or use defaults"""
        try:
            # Create user_data directory if it doesn't exist
            os.makedirs("user_data", exist_ok=True)

            preferences_path = os.path.join("user_data", "preferences.json")
            if os.path.exists(preferences_path):
                with open(preferences_path, "r") as f:
                    return json.load(f)
            else:
                # Default preferences
                default_prefs = {"advanced_mode": False}
                # Save defaults
                with open(preferences_path, "w") as f:
                    json.dump(default_prefs, f, indent=2)
                return default_prefs
        except Exception as e:
            logger.error(f"Error loading preferences: {e}")
            return {"advanced_mode": False}

    def _save_user_preferences(self):
        """Save current user preferences"""
        try:
            preferences_path = os.path.join("user_data", "preferences.json")
            with open(preferences_path, "w") as f:
                json.dump(self.user_preferences, f, indent=2)
            logger.info("User preferences saved")
        except Exception as e:
            logger.error(f"Error saving preferences: {e}")

    def _setup_ui(self):
        """Set up the user interface with notebook tabs"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Create tabs - all tabs are created but not all are added initially
        self.main_tab = ttk.Frame(self.notebook)
        self.config_wizard_tab = ttk.Frame(self.notebook)
        self.testing_tab = ttk.Frame(self.notebook)
        self.scheduler_tab = ttk.Frame(self.notebook)
        self.data_source_tab = ttk.Frame(self.notebook)
        self.reference_data_tab = ttk.Frame(self.notebook)

        # Always add main tab
        self.notebook.add(self.main_tab, text="Run Analytics")

        # Conditionally add other tabs based on mode
        if self.is_advanced_mode:
            self.notebook.add(self.config_wizard_tab, text="Configuration Wizard")
            self.notebook.add(self.testing_tab, text="Testing")
            self.notebook.add(self.scheduler_tab, text="Scheduler")
            self.notebook.add(self.data_source_tab, text="Data Sources")
            self.notebook.add(self.reference_data_tab, text="Reference Data")

        # Set up each tab
        self._setup_main_tab()
        self._setup_config_wizard_tab()
        self._setup_testing_tab()
        self._setup_scheduler_tab()
        self._setup_data_source_tab()
        self._setup_reference_data_tab()

        # Status bar with mode toggle
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Frame(self.root, relief=tk.SUNKEN)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Status label on left
        status_label = ttk.Label(self.status_bar, textvariable=self.status_var, anchor=tk.W)
        status_label.pack(side=tk.LEFT, padx=5)

        # Mode toggle on right
        self.mode_btn = ttk.Button(
            self.status_bar,
            text="Switch to Advanced Mode" if not self.is_advanced_mode else "Switch to Simple Mode",
            command=self._toggle_ui_mode,
            width=20
        )
        self.mode_btn.pack(side=tk.RIGHT, padx=5, pady=2)

        # Set up log handler
        self._setup_log_handler()

    def _toggle_ui_mode(self):
        """Toggle between simple and advanced UI modes"""
        self.is_advanced_mode = not self.is_advanced_mode

        # Update button text
        self.mode_btn.config(
            text="Switch to Advanced Mode" if not self.is_advanced_mode else "Switch to Simple Mode"
        )

        # Update user preferences
        self.user_preferences["advanced_mode"] = self.is_advanced_mode
        self._save_user_preferences()

        # Update UI
        self._update_ui_mode()

    def _update_ui_mode(self):
        """Update UI based on current mode"""
        # Remove all tabs except main tab
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)

        # Always add main tab
        self.notebook.add(self.main_tab, text="Run Analytics")

        # Add other tabs in advanced mode
        if self.is_advanced_mode:
            self.notebook.add(self.config_wizard_tab, text="Configuration Wizard")
            self.notebook.add(self.testing_tab, text="Testing")
            self.notebook.add(self.scheduler_tab, text="Scheduler")
            self.notebook.add(self.data_source_tab, text="Data Sources")
            self.notebook.add(self.reference_data_tab, text="Reference Data")

        # Update status message
        self.status_var.set(f"Ready - {'Advanced' if self.is_advanced_mode else 'Simple'} Mode")

    def _setup_main_tab(self):
        """Set up the main analytics tab with a simplified, step-based interface"""
        # Clean up any existing widgets
        for widget in self.main_tab.winfo_children():
            widget.destroy()

        # Create main frame with padding
        main_frame = ttk.Frame(self.main_tab, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Add a clear title
        title_label = ttk.Label(
            main_frame,
            text="Run Quality Assurance Analytics",
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 20))

        # Step 1: Select Analytics
        step1_frame = ttk.LabelFrame(main_frame, text="Step 1: Select Analytics")
        step1_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15), padx=5)

        # Add "Quick Select" and "Recently Used" sections
        quick_select_frame = ttk.Frame(step1_frame)
        quick_select_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(quick_select_frame, text="Quick Select:").grid(row=0, column=0, sticky=tk.W)

        # Define quick select buttons
        quick_select_buttons = [
            ("Audit Approvals", "77"),
            ("Risk Assessment", "78"),
            ("Issue Management", "02")
        ]

        # Add quick select buttons
        for i, (name, qa_id) in enumerate(quick_select_buttons):
            ttk.Button(
                quick_select_frame,
                text=name,
                command=lambda id=qa_id: self._quick_select_analytics(id)
            ).grid(row=0, column=i + 1, padx=5)

        # Standard dropdown for analytics selection
        ttk.Label(step1_frame, text="Or select from list:").pack(anchor=tk.W, padx=10, pady=(0, 5))

        self.analytic_var = tk.StringVar()
        self.analytic_combo = ttk.Combobox(
            step1_frame,
            textvariable=self.analytic_var,
            state="readonly",
            width=50
        )
        self.analytic_combo["values"] = [f"{id} - {name}" for id, name in self.available_analytics]
        if self.available_analytics:
            self.analytic_combo.current(0)
        self.analytic_combo.pack(fill=tk.X, padx=10, pady=(0, 10))

        # Step 2: Select Data
        step2_frame = ttk.LabelFrame(main_frame, text="Step 2: Select Data Source")
        step2_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15), padx=5)

        # Simplified file selection
        file_frame = ttk.Frame(step2_frame)
        file_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(file_frame, text="Data File:").grid(row=0, column=0, sticky=tk.W)

        self.source_var = tk.StringVar()
        self.source_entry = ttk.Entry(file_frame, textvariable=self.source_var, width=50)
        self.source_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)

        ttk.Button(file_frame, text="Browse...", command=self._browse_source).grid(row=0, column=2)

        # Make source entry expandable
        file_frame.columnconfigure(1, weight=1)

        # Step 3: Run and View Results
        step3_frame = ttk.LabelFrame(main_frame, text="Step 3: Run Analysis")
        step3_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15), padx=5)

        action_frame = ttk.Frame(step3_frame)
        action_frame.pack(fill=tk.X, padx=10, pady=10)

        self.progress = ttk.Progressbar(action_frame, orient="horizontal", length=200, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.run_btn = ttk.Button(action_frame, text="Run Analysis", command=self._run_analysis)
        self.run_btn.pack(side=tk.RIGHT)

        # Results section (initially collapsed)
        self.results_section = ttk.LabelFrame(main_frame, text="Results")
        self.results_section.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10), padx=5)

        # Results will be populated after running analysis
        self.results_placeholder = ttk.Label(
            self.results_section,
            text="Run an analysis to see results",
            font=("Arial", 10, "italic")
        )
        self.results_placeholder.pack(padx=20, pady=20)

        # Log section at bottom with toggle
        log_header = ttk.Frame(main_frame)
        log_header.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))

        ttk.Label(log_header, text="Status Log:").pack(side=tk.LEFT)

        self.log_toggle_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            log_header,
            text="Show Log",
            variable=self.log_toggle_var,
            command=self._toggle_log_visibility
        ).pack(side=tk.RIGHT)

        # Log container (initially hidden)
        self.log_container = ttk.Frame(main_frame)
        self.log_container.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.log_text = tk.Text(self.log_container, height=15, width=80, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        self.log_text.config(state=tk.DISABLED)

        log_scroll = ttk.Scrollbar(self.log_container, orient="vertical", command=self.log_text.yview)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=log_scroll.set)

        # Hide log initially
        self.log_container.grid_remove()

        # Configure resizing
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Make results section expandable
        main_frame.rowconfigure(4, weight=1)

    def _toggle_log_visibility(self):
        """Toggle the visibility of the log section"""
        if self.log_toggle_var.get():
            self.log_container.grid()
        else:
            self.log_container.grid_remove()

    def _quick_select_analytics(self, analytics_id):
        """Quickly select an analytics configuration by ID"""
        # Find the matching item in the dropdown and select it
        for i, (id, _) in enumerate(self.available_analytics):
            if id == analytics_id:
                self.analytic_combo.current(i)
                break

        # Set focus to the data file selection
        self.source_entry.focus_set()

    def _setup_config_wizard_tab(self):
        """Set up the configuration wizard tab"""
        # Create the configuration wizard
        self.config_wizard = ConfigWizard(
            parent_frame=self.config_wizard_tab,
            config_manager=self.config_manager,
            template_manager=self.template_manager
        )

    def _setup_testing_tab(self):
        """Set up the testing tab"""
        # Create the testing environment
        self.testing_environment = TestingEnvironment(
            parent_frame=self.testing_tab,
            config_manager=self.config_manager,
            data_processor_class=EnhancedDataProcessor,
            report_generator_class=EnhancedReportGenerator
        )

    def _setup_scheduler_tab(self):
        """Set up the scheduler tab"""
        # Create the scheduler UI
        self.scheduler_ui = SchedulerUI(
            parent_frame=self.scheduler_tab,
            scheduler=self.scheduler
        )

    def _setup_data_source_tab(self):
        """Set up the data source management tab"""
        # Create frames
        frame = ttk.LabelFrame(self.data_source_tab, text="Data Source Registry")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create treeview
        columns = ("Name", "Type", "Owner", "Version", "Last Updated", "Analytics")
        self.data_source_tree = ttk.Treeview(frame, columns=columns, show="headings", height=15)

        # Configure columns
        self.data_source_tree.column("Name", width=100)
        self.data_source_tree.column("Type", width=80)
        self.data_source_tree.column("Owner", width=150)
        self.data_source_tree.column("Version", width=80)
        self.data_source_tree.column("Last Updated", width=100)
        self.data_source_tree.column("Analytics", width=80)

        # Configure headings
        self.data_source_tree.heading("Name", text="Data Source")
        self.data_source_tree.heading("Type", text="Type")
        self.data_source_tree.heading("Owner", text="Owner")
        self.data_source_tree.heading("Version", text="Version")
        self.data_source_tree.heading("Last Updated", text="Last Updated")
        self.data_source_tree.heading("Analytics", text="# Analytics")

        # Add scrollbar
        tree_scroll = ttk.Scrollbar(frame, orient="vertical", command=self.data_source_tree.yview)
        self.data_source_tree.configure(yscrollcommand=tree_scroll.set)

        # Pack tree and scrollbar
        self.data_source_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Populate data source tree
        self._populate_data_source_tree()

        # Add buttons
        button_frame = ttk.Frame(self.data_source_tab)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        refresh_btn = ttk.Button(button_frame, text="Refresh Registry",
                                 command=self._refresh_data_source_registry)
        refresh_btn.pack(side=tk.LEFT, padx=5)

        view_details_btn = ttk.Button(button_frame, text="View Details",
                                      command=self._view_data_source_details)
        view_details_btn.pack(side=tk.RIGHT, padx=5)

    def _setup_reference_data_tab(self):
        """Set up the reference data management tab"""
        # Create frames
        frame = ttk.LabelFrame(self.reference_data_tab, text="Reference Data")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create treeview
        columns = ("Name", "Format", "Version", "Last Modified", "Rows", "Freshness")
        self.reference_tree = ttk.Treeview(frame, columns=columns, show="headings", height=15)

        # Configure columns
        self.reference_tree.column("Name", width=150)
        self.reference_tree.column("Format", width=80)
        self.reference_tree.column("Version", width=80)
        self.reference_tree.column("Last Modified", width=150)
        self.reference_tree.column("Rows", width=80, anchor=tk.CENTER)
        self.reference_tree.column("Freshness", width=100, anchor=tk.CENTER)

        # Configure headings
        self.reference_tree.heading("Name", text="Reference Data")
        self.reference_tree.heading("Format", text="Format")
        self.reference_tree.heading("Version", text="Version")
        self.reference_tree.heading("Last Modified", text="Last Modified")
        self.reference_tree.heading("Rows", text="# Rows")
        self.reference_tree.heading("Freshness", text="Freshness")

        # Add scrollbar
        tree_scroll = ttk.Scrollbar(frame, orient="vertical", command=self.reference_tree.yview)
        self.reference_tree.configure(yscrollcommand=tree_scroll.set)

        # Pack tree and scrollbar
        self.reference_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Populate reference tree
        self._populate_reference_tree()

        # Add buttons
        button_frame = ttk.Frame(self.reference_data_tab)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        refresh_btn = ttk.Button(button_frame, text="Refresh Status",
                                 command=self._refresh_reference_status)
        refresh_btn.pack(side=tk.LEFT, padx=5)

        update_btn = ttk.Button(button_frame, text="Update Reference File",
                                command=self._update_reference_file)
        update_btn.pack(side=tk.LEFT, padx=5)

        history_btn = ttk.Button(button_frame, text="View Update History",
                                 command=self._view_reference_history)
        history_btn.pack(side=tk.RIGHT, padx=5)

    def _setup_log_handler(self):
        """Set up log handler to redirect to text widget"""

        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                logging.Handler.__init__(self)
                self.text_widget = text_widget

            def emit(self, record):
                msg = self.format(record)

                def append():
                    self.text_widget.config(state=tk.NORMAL)
                    self.text_widget.insert(tk.END, msg + "\n")
                    self.text_widget.see(tk.END)
                    self.text_widget.config(state=tk.DISABLED)

                # Schedule to be executed in the main thread
                self.text_widget.after(0, append)

        # Create a handler and add it to the logger
        text_handler = TextHandler(self.log_text)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%H:%M:%S')
        text_handler.setFormatter(formatter)
        logger.addHandler(text_handler)

    def _browse_source(self):
        """Browse for source data file"""
        filename = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")],
            title="Select Source Data File"
        )
        if filename:
            self.source_var.set(filename)

    def _browse_output(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directory:
            self.output_var.set(directory)

    def _update_results_display(self, processor_results):
        """Update the results display area with analysis results"""
        # Clear placeholder
        for widget in self.results_section.winfo_children():
            widget.destroy()

        # Create notebook for result tabs
        results_notebook = ttk.Notebook(self.results_section)
        results_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Create result tabs
        summary_tab = ttk.Frame(results_notebook)
        detail_tab = ttk.Frame(results_notebook)

        results_notebook.add(summary_tab, text="Summary")
        results_notebook.add(detail_tab, text="Detail")

        # Summary tab
        summary_frame = ttk.Frame(summary_tab)
        summary_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Statistics at the top
        stats_frame = ttk.LabelFrame(summary_frame, text="Results Summary")
        stats_frame.pack(fill=tk.X, pady=(0, 10))

        # Get statistics from results
        detail_data = processor_results.get('detail')
        if detail_data is not None and 'Compliance' in detail_data:
            total = len(detail_data)
            gc_count = sum(detail_data['Compliance'] == 'GC')
            dnc_count = sum(detail_data['Compliance'] == 'DNC')
            pc_count = sum(detail_data['Compliance'] == 'PC')

            error_pct = (dnc_count / total * 100) if total > 0 else 0

            # Display stats in a grid
            stats_grid = ttk.Frame(stats_frame)
            stats_grid.pack(padx=10, pady=10, fill=tk.X)

            # Create a 2-column grid for stats
            ttk.Label(stats_grid, text="Total Records:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
            ttk.Label(stats_grid, text=str(total)).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

            ttk.Label(stats_grid, text="Generally Conforms (GC):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
            ttk.Label(stats_grid, text=f"{gc_count} ({gc_count / total * 100:.1f}%)").grid(row=1, column=1, sticky=tk.W,
                                                                                           padx=5, pady=2)

            ttk.Label(stats_grid, text="Does Not Conform (DNC):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
            ttk.Label(stats_grid, text=f"{dnc_count} ({dnc_count / total * 100:.1f}%)").grid(row=2, column=1,
                                                                                             sticky=tk.W, padx=5,
                                                                                             pady=2)

            ttk.Label(stats_grid, text="Partially Conforms (PC):").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
            ttk.Label(stats_grid, text=f"{pc_count} ({pc_count / total * 100:.1f}%)").grid(row=3, column=1, sticky=tk.W,
                                                                                           padx=5, pady=2)

            ttk.Label(stats_grid, text="Error Percentage:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)
            ttk.Label(stats_grid, text=f"{error_pct:.2f}%").grid(row=4, column=1, sticky=tk.W, padx=5, pady=2)

        # Group results if available
        summary_data = processor_results.get('summary')
        if summary_data is not None:
            # Group results table
            group_frame = ttk.LabelFrame(summary_frame, text="Group Summary")
            group_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

            # Create treeview for group results
            columns = ["Group", "GC", "PC", "DNC", "Total", "Error %", "Status"]
            group_tree = ttk.Treeview(group_frame, columns=columns, show="headings", height=8)

            # Configure columns
            group_tree.column("Group", width=150)
            group_tree.column("GC", width=70, anchor=tk.CENTER)
            group_tree.column("PC", width=70, anchor=tk.CENTER)
            group_tree.column("DNC", width=70, anchor=tk.CENTER)
            group_tree.column("Total", width=70, anchor=tk.CENTER)
            group_tree.column("Error %", width=70, anchor=tk.CENTER)
            group_tree.column("Status", width=100, anchor=tk.CENTER)

            # Configure headings
            for col in columns:
                group_tree.heading(col, text=col)

            # Add scrollbar
            tree_scroll = ttk.Scrollbar(group_frame, orient="vertical", command=group_tree.yview)
            group_tree.configure(yscrollcommand=tree_scroll.set)

            # Pack tree and scrollbar
            group_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

            # Add data to tree
            for _, row in summary_data.iterrows():
                # Get group by field name
                group_field = self.current_config.get('reporting', {}).get('group_by', '')
                group_value = row[group_field] if group_field in row else "Unknown"

                # Get counts
                gc = row.get('GC', 0)
                pc = row.get('PC', 0)
                dnc = row.get('DNC', 0)
                total = row.get('Total', 0)
                dnc_pct = row.get('DNC_Percentage', 0)
                exceeds = row.get('Exceeds_Threshold', False)

                # Add to tree
                item_id = group_tree.insert("", tk.END, values=(
                    group_value,
                    gc,
                    pc,
                    dnc,
                    total,
                    f"{dnc_pct:.2f}%",
                    "EXCEEDS" if exceeds else "Within Threshold"
                ))

                # Color code based on threshold
                if exceeds:
                    group_tree.item(item_id, tags=("exceeds",))

            # Configure tag colors
            group_tree.tag_configure("exceeds", background="#FFE6E6")  # Light red

        # Detail tab
        detail_frame = ttk.Frame(detail_tab)
        detail_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Add filter options
        filter_frame = ttk.Frame(detail_frame)
        filter_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(filter_frame, text="Show:").pack(side=tk.LEFT)

        filter_var = tk.StringVar(value="all")

        ttk.Radiobutton(filter_frame, text="All", variable=filter_var,
                        value="all").pack(side=tk.LEFT, padx=(5, 10))
        ttk.Radiobutton(filter_frame, text="GC Only", variable=filter_var,
                        value="gc").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(filter_frame, text="DNC Only", variable=filter_var,
                        value="dnc").pack(side=tk.LEFT, padx=(0, 10))

        # Detail data table
        if detail_data is not None:
            # Create a container frame for the treeview and scrollbars
            tree_container = ttk.Frame(detail_frame)
            tree_container.pack(fill=tk.BOTH, expand=True)

            # Create treeview for detail results
            detail_tree = ttk.Treeview(tree_container, show="headings")

            # Set up columns
            columns = list(detail_data.columns)
            detail_tree['columns'] = columns

            # Configure columns and headings
            for col in columns:
                detail_tree.column(col, width=100, stretch=True)
                detail_tree.heading(col, text=col)

            # Add scrollbars
            detail_y_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=detail_tree.yview)
            detail_x_scroll = ttk.Scrollbar(tree_container, orient="horizontal", command=detail_tree.xview)
            detail_tree.configure(yscrollcommand=detail_y_scroll.set, xscrollcommand=detail_x_scroll.set)

            # Pack tree and scrollbars
            detail_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            detail_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            detail_x_scroll.pack(side=tk.BOTTOM, fill=tk.X)

            # Add data rows
            for idx, row in detail_data.iterrows():
                values = []
                for col in columns:
                    val = row[col]
                    # Format date values
                    if isinstance(val, pd.Timestamp) or isinstance(val, datetime.datetime):
                        val = val.strftime('%Y-%m-%d %H:%M')
                    values.append(val)

                item_id = detail_tree.insert("", tk.END, values=values)

                # Color code based on compliance
                compliance = row.get('Compliance')
                if compliance == 'GC':
                    detail_tree.item(item_id, tags=("gc",))
                elif compliance == 'DNC':
                    detail_tree.item(item_id, tags=("dnc",))
                elif compliance == 'PC':
                    detail_tree.item(item_id, tags=("pc",))

            # Configure tag colors
            detail_tree.tag_configure("gc", background="#E6FFE6")  # Light green
            detail_tree.tag_configure("dnc", background="#FFE6E6")  # Light red
            detail_tree.tag_configure("pc", background="#FFF8E6")  # Light yellow

            # Connect filter buttons to the tree
            def apply_filter():
                filter_value = filter_var.get()

                # Show all rows first
                for item in detail_tree.get_children():
                    detail_tree.item(item, open=True)

                # Hide rows that don't match the filter
                if filter_value != "all":
                    for item in detail_tree.get_children():
                        tags = detail_tree.item(item, "tags")
                        if filter_value not in tags:
                            detail_tree.detach(item)

            # Connect radiobuttons to filter function
            for rb in filter_frame.winfo_children():
                if isinstance(rb, ttk.Radiobutton):
                    rb.config(command=apply_filter)

        # Add buttons for report generation
        button_frame = ttk.Frame(self.results_section)
        button_frame.pack(fill=tk.X, pady=(0, 10), padx=5)

        ttk.Button(
            button_frame,
            text="Generate Excel Report",
            command=self._generate_excel_report
        ).pack(side=tk.RIGHT)

        ttk.Button(
            button_frame,
            text="Export Results",
            command=self._export_results
        ).pack(side=tk.RIGHT, padx=(0, 10))

    def _run_analysis(self):
        """Run the analysis process"""
        # Validate inputs
        if not self.analytic_var.get():
            messagebox.showerror("Error", "Please select a QA-ID")
            return

        if not self.source_var.get():
            messagebox.showerror("Error", "Please select a source data file")
            return

        if not os.path.exists(self.source_var.get()):
            messagebox.showerror("Error", "Source data file does not exist")
            return

        # Get the analytic ID from selection
        analytic_id = self.analytic_var.get().split(" - ")[0]

        # Start progress bar
        self.progress.start()
        self.status_var.set("Processing...")
        self.run_btn.config(state=tk.DISABLED)

        # Run in a separate thread to avoid freezing the UI
        threading.Thread(target=self._process_data, args=(analytic_id,), daemon=True).start()

    def _process_data(self, analytic_id):
        """Process data in a separate thread"""
        try:
            # Get configuration
            self.current_config = self.config_manager.get_config(analytic_id)

            # Create output directory if needed
            output_dir = "output"
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # Initialize enhanced processor
            processor = EnhancedDataProcessor(self.current_config)

            # Process data
            logger.info(f"Starting processing for QA-ID {analytic_id}")
            success, message = processor.process_data(self.source_var.get())

            if not success:
                self.root.after(0, lambda: messagebox.showerror("Error", message))
                self.root.after(0, self._reset_ui_after_processing)
                return

            # Store results for reporting
            self.processor_results = processor.results

            # Update UI with results
            self.root.after(0, lambda: self._update_results_display(processor.results))

            # Show success message
            self.root.after(0, lambda: messagebox.showinfo("Success", "Analysis completed successfully"))
            self.root.after(0, lambda: self.status_var.set("Ready - Analysis complete"))

        except Exception as e:
            logger.error(f"Error in processing: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {e}"))
            self.root.after(0, lambda: self.status_var.set("Ready - Error occurred"))

        finally:
            # Reset UI
            self.root.after(0, self._reset_ui_after_processing)

    def _reset_ui_after_processing(self):
        """Reset UI after processing completes"""
        self.progress.stop()
        self.run_btn.config(state=tk.NORMAL)

    def _generate_excel_report(self):
        """Generate a formatted Excel report"""
        if not hasattr(self, 'processor_results') or not self.processor_results:
            messagebox.showinfo("Info", "No results available to generate report")
            return

        if not hasattr(self, 'current_config') or not self.current_config:
            messagebox.showinfo("Info", "No configuration loaded")
            return

        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Excel Report"
        )

        if not file_path:
            return

        try:
            # Create report generator
            report_generator = EnhancedReportGenerator(self.current_config, self.processor_results)

            # Generate report
            report_path = report_generator.generate_main_report(file_path)

            if report_path:
                messagebox.showinfo("Success", f"Report generated: {report_path}")
            else:
                messagebox.showerror("Error", "Failed to generate report")
        except Exception as e:
            messagebox.showerror("Error", f"Error generating report: {e}")
            logger.error(f"Error generating report: {e}", exc_info=True)

    def _export_results(self):
        """Export the raw results to Excel files"""
        if not hasattr(self, 'processor_results') or not self.processor_results:
            messagebox.showinfo("Info", "No results available to export")
            return

        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Export Results"
        )

        if not file_path:
            return

        try:
            # Create a workbook
            with pd.ExcelWriter(file_path) as writer:
                # Save summary
                if 'summary' in self.processor_results and self.processor_results['summary'] is not None:
                    self.processor_results['summary'].to_excel(writer, sheet_name='Summary', index=False)

                # Save detail
                if 'detail' in self.processor_results and self.processor_results['detail'] is not None:
                    self.processor_results['detail'].to_excel(writer, sheet_name='Detail', index=False)

            messagebox.showinfo("Success", f"Results exported to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export results: {e}")

    def _populate_data_source_tree(self):
        """Populate data source tree with registry information"""
        # Clear existing items
        for item in self.data_source_tree.get_children():
            self.data_source_tree.delete(item)

        # Get data source info
        source_info = self.data_source_manager.get_data_source_info()

        # Add sources to tree
        for name, info in source_info.get('sources', {}).items():
            # Format last updated date
            last_updated = info.get('last_modified', info.get('last_updated', 'Unknown'))
            if isinstance(last_updated, datetime.datetime):
                last_updated = last_updated.strftime('%Y-%m-%d')

            # Add to tree
            self.data_source_tree.insert("", tk.END, values=(
                name,
                info.get('type', 'Unknown'),
                info.get('owner', 'Unknown'),
                info.get('version', 'Unknown'),
                last_updated,
                len(info.get('analytics', []))
            ))

    def _refresh_data_source_registry(self):
        """Refresh the data source registry display"""
        # Reload registry
        self.data_source_manager = DataSourceManager()
        # Update tree
        self._populate_data_source_tree()
        # Show confirmation
        self.status_var.set("Data source registry refreshed")

    def _view_data_source_details(self):
        """View detailed information for the selected data source"""
        # Get selected item
        item = self.data_source_tree.focus()
        if not item:
            messagebox.showinfo("No Selection", "Please select a data source first")
            return

        # Get the data source name
        values = self.data_source_tree.item(item, "values")
        if not values:
            return

        data_source_name = values[0]

        # Get source config
        source_config = self.data_source_manager.get_data_source_config(data_source_name)
        if not source_config:
            messagebox.showerror("Error", f"Data source '{data_source_name}' not found in registry")
            return

        # Create a dialog to display details
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Data Source Details: {data_source_name}")
        dialog.geometry("600x500")

        # Create frame with scrollbar
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create text widget
        text = tk.Text(frame, wrap=tk.WORD)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add scrollbar
        scroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text.config(yscrollcommand=scroll.set)

        # Format and display details
        text.insert(tk.END, f"DATA SOURCE: {data_source_name}\n")
        text.insert(tk.END, "=" * 50 + "\n\n")

        text.insert(tk.END, f"Description: {source_config.get('description', 'N/A')}\n")
        text.insert(tk.END, f"Type: {source_config.get('type', 'N/A')}\n")
        text.insert(tk.END, f"Owner: {source_config.get('owner', 'N/A')}\n")
        text.insert(tk.END, f"Version: {source_config.get('version', 'N/A')}\n")
        text.insert(tk.END, f"Last Updated: {source_config.get('last_updated', 'N/A')}\n")
        text.insert(tk.END, f"Refresh Frequency: {source_config.get('refresh_frequency', 'N/A')}\n")
        text.insert(tk.END, f"File Type: {source_config.get('file_type', 'N/A')}\n")
        text.insert(tk.END, f"File Pattern: {source_config.get('file_pattern', 'N/A')}\n\n")

        # Key columns
        text.insert(tk.END, "KEY COLUMNS:\n")
        for col in source_config.get('key_columns', []):
            text.insert(tk.END, f"  - {col}\n")
        text.insert(tk.END, "\n")

        # Validation rules
        text.insert(tk.END, "VALIDATION RULES:\n")
        for rule in source_config.get('validation_rules', []):
            text.insert(tk.END, f"  - {rule.get('type')}: {rule.get('description', '')}\n")
            if 'threshold' in rule:
                text.insert(tk.END, f"    Threshold: {rule['threshold']}\n")
            if 'columns' in rule:
                text.insert(tk.END, f"    Columns: {', '.join(rule['columns'])}\n")
        text.insert(tk.END, "\n")

        # Column mappings
        text.insert(tk.END, "COLUMN MAPPINGS:\n")
        for mapping in source_config.get('columns_mapping', []):
            text.insert(tk.END,
                        f"  - {mapping.get('source')} -> {mapping.get('target')} ({mapping.get('data_type', 'no type')})\n")
            if mapping.get('aliases'):
                text.insert(tk.END, f"    Aliases: {', '.join(mapping['aliases'])}\n")
        text.insert(tk.END, "\n")

        # Associated analytics
        analytics = []
        for analytic_id, source in self.data_source_manager.analytics_mapping.items():
            if source == data_source_name:
                analytics.append(analytic_id)

        text.insert(tk.END, "ASSOCIATED ANALYTICS:\n")
        if analytics:
            for analytic_id in analytics:
                # Get analytic name if available
                analytic_name = next((name for id, name in self.available_analytics if id == analytic_id), "Unknown")
                text.insert(tk.END, f"  - QA-{analytic_id}: {analytic_name}\n")
        else:
            text.insert(tk.END, "  No analytics associated with this data source\n")

        # Make text read-only
        text.config(state=tk.DISABLED)

        # Add close button
        close_btn = ttk.Button(dialog, text="Close", command=dialog.destroy)
        close_btn.pack(pady=10)

    def _populate_reference_tree(self):
        """Populate reference data tree with status information"""
        # Clear existing items
        for item in self.reference_tree.get_children():
            self.reference_tree.delete(item)

        # Get reference data info
        reference_info = self.reference_data_manager.get_reference_data_info()

        # Add to tree
        for name, info in reference_info.items():
            # Format last modified date
            last_modified = info.get('last_modified', 'Not loaded')
            if last_modified != 'Not loaded':
                last_modified = last_modified.strftime('%Y-%m-%d %H:%M')

            # Format freshness
            if 'is_fresh' in info:
                freshness = "✓ Fresh" if info['is_fresh'] else "⚠ Stale"
                tag = "fresh" if info['is_fresh'] else "stale"
            else:
                freshness = "Not loaded"
                tag = "not_loaded"

            # Add to tree
            item = self.reference_tree.insert("", tk.END, values=(
                name,
                info.get('format', 'Unknown'),
                info.get('version', 'Unknown'),
                last_modified,
                info.get('row_count', '-'),
                freshness
            ), tags=(tag,))

        # Configure tags for color coding
        self.reference_tree.tag_configure("fresh", background="#e6ffe6")  # Light green
        self.reference_tree.tag_configure("stale", background="#fff0e6")  # Light orange
        self.reference_tree.tag_configure("not_loaded", background="#f0f0f0")  # Light gray

    def _refresh_reference_status(self):
        """Refresh the reference data status display"""
        # Update tree
        self._populate_reference_tree()
        # Show confirmation
        self.status_var.set("Reference data status refreshed")

    def _update_reference_file(self):
        """Update a reference data file"""
        # Get selected reference data
        item = self.reference_tree.focus()
        if not item:
            messagebox.showinfo("No Selection", "Please select a reference data entry first")
            return

        values = self.reference_tree.item(item, "values")
        ref_name = values[0]

        # Show file dialog
        filename = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")],
            title=f"Select New File for Reference Data '{ref_name}'"
        )

        if not filename:
            return

        # Confirm update
        if messagebox.askyesno("Confirm Update",
                               f"Are you sure you want to update reference data '{ref_name}' with file:\n{filename}?"):
            # Get username
            username = os.environ.get('USERNAME', 'unknown')

            # Update reference data
            success = self.reference_data_manager.update_reference_data(ref_name, filename, username)

            if success:
                messagebox.showinfo("Success", f"Reference data '{ref_name}' updated successfully")
                # Refresh display
                self._populate_reference_tree()
            else:
                messagebox.showerror("Error", f"Failed to update reference data '{ref_name}'")

    def _view_reference_history(self):
        """View reference data update history"""
        # Create a dialog to display history
        dialog = tk.Toplevel(self.root)
        dialog.title("Reference Data Update History")
        dialog.geometry("800x500")

        # Create text widget for display
        text = tk.Text(dialog, wrap=tk.WORD)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Add scrollbar
        scroll = ttk.Scrollbar(dialog, orient="vertical", command=text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=10)
        text.config(yscrollcommand=scroll.set)

        # Format and display history
        if hasattr(self.reference_data_manager, 'audit_log'):
            if not self.reference_data_manager.audit_log:
                text.insert(tk.END, "No history records found.")
            else:
                # Sort by timestamp, newest first
                sorted_log = sorted(self.reference_data_manager.audit_log,
                                    key=lambda x: x.get('timestamp', ''), reverse=True)

                for entry in sorted_log:
                    # Format timestamp
                    timestamp = entry.get('timestamp', 'Unknown')
                    if isinstance(timestamp, str):
                        formatted_time = timestamp
                    else:
                        try:
                            formatted_time = timestamp.strftime('%Y-%m-%d %H:%M:%S')
                        except:
                            formatted_time = str(timestamp)

                    # Format entry
                    text.insert(tk.END, f"Time: {formatted_time}\n")
                    text.insert(tk.END, f"User: {entry.get('user', 'Unknown')}\n")
                    text.insert(tk.END, f"Action: {entry.get('action', 'Unknown')}\n")
                    text.insert(tk.END, f"Reference Data: {entry.get('name', 'Unknown')}\n")

                    # Previous version info
                    prev = entry.get('previous_version')
                    if prev:
                        prev_version = prev.get('version', 'Unknown')
                        prev_modified = prev.get('last_modified', 'Unknown')
                        if not isinstance(prev_modified, str):
                            try:
                                prev_modified = prev_modified.strftime('%Y-%m-%d')
                            except:
                                prev_modified = str(prev_modified)
                        text.insert(tk.END, f"Previous Version: {prev_version} (Modified: {prev_modified})\n")

                    # New version info
                    new = entry.get('new_version')
                    if new:
                        new_version = new.get('version', 'Unknown')
                        new_modified = new.get('last_modified', 'Unknown')
                        if not isinstance(new_modified, str):
                            try:
                                new_modified = new_modified.strftime('%Y-%m-%d')
                            except:
                                new_modified = str(new_modified)
                        text.insert(tk.END, f"New Version: {new_version} (Modified: {new_modified})\n")

                    text.insert(tk.END, "\n" + "-" * 50 + "\n\n")
        else:
            text.insert(tk.END, "Audit logging is not enabled for reference data.")

        # Make text read-only
        text.config(state=tk.DISABLED)

        # Add close button
        close_btn = ttk.Button(dialog, text="Close", command=dialog.destroy)
        close_btn.pack(pady=10)


# Application entry point
if __name__ == "__main__":
    root = tk.Tk()
    app = EnhancedQAAnalyticsApp(root)
    root.mainloop()
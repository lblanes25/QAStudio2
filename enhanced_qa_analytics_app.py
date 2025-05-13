import os
import sys
import logging
import threading
import tkinter as tk
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
    """Enhanced application with GUI interface and data management tabs"""

    def __init__(self, root):
        """Initialize the application"""
        self.root = root
        self.root.title("Enhanced QA Analytics Automation")
        self.root.geometry("900x700")

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

    def _setup_ui(self):
        """Set up the user interface with notebook tabs"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Create tabs
        self.main_tab = ttk.Frame(self.notebook)
        self.config_wizard_tab = ttk.Frame(self.notebook)
        self.testing_tab = ttk.Frame(self.notebook)
        self.scheduler_tab = ttk.Frame(self.notebook)
        self.data_source_tab = ttk.Frame(self.notebook)
        self.reference_data_tab = ttk.Frame(self.notebook)

        # Add tabs to notebook
        self.notebook.add(self.main_tab, text="Run Analytics")
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

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Set up log handler
        self._setup_log_handler()

    def _setup_main_tab(self):
        """Set up the main analytics tab"""
        # Create main frame
        main_frame = ttk.Frame(self.main_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Analytics selection
        ttk.Label(main_frame, text="Select QA-ID:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        self.analytic_var = tk.StringVar()
        self.analytic_combo = ttk.Combobox(main_frame, textvariable=self.analytic_var, state="readonly", width=50)
        self.analytic_combo["values"] = [f"{id} - {name}" for id, name in self.available_analytics]
        if self.available_analytics:
            self.analytic_combo.current(0)
        self.analytic_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5))

        # Source file selection
        ttk.Label(main_frame, text="Source Data File:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))

        source_frame = ttk.Frame(main_frame)
        source_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 5))

        self.source_var = tk.StringVar()
        self.source_entry = ttk.Entry(source_frame, textvariable=self.source_var, width=50)
        self.source_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        source_btn = ttk.Button(source_frame, text="Browse...", command=self._browse_source)
        source_btn.pack(side=tk.RIGHT, padx=(5, 0))

        # Output directory
        ttk.Label(main_frame, text="Output Directory:").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))

        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=(0, 5))

        self.output_var = tk.StringVar(value="output")
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_var, width=50)
        self.output_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        output_btn = ttk.Button(output_frame, text="Browse...", command=self._browse_output)
        output_btn.pack(side=tk.RIGHT, padx=(5, 0))

        # Execution frame
        exec_frame = ttk.Frame(main_frame)
        exec_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))

        self.progress = ttk.Progressbar(exec_frame, orient="horizontal", length=200, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        exec_btn = ttk.Button(exec_frame, text="Run Analysis", command=self._run_analysis)
        exec_btn.pack(side=tk.RIGHT)

        # Status log
        ttk.Label(main_frame, text="Status Log:").grid(row=4, column=0, sticky=tk.W, pady=(10, 5))

        self.log_text = tk.Text(main_frame, height=15, width=80, wrap=tk.WORD)
        self.log_text.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_text.config(state=tk.DISABLED)

        # Add scrollbar to log
        log_scroll = ttk.Scrollbar(main_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=5, column=2, sticky=(tk.N, tk.S))
        self.log_text.config(yscrollcommand=log_scroll.set)

        # Configure resizing
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)

    def _setup_config_wizard_tab(self):
        """Set up the configuration wizard tab"""
        # Create the configuration wizard with a callback
        self.config_wizard = ConfigWizard(
            parent_frame=self.config_wizard_tab,
            config_manager=self.config_manager,
            template_manager=self.template_manager,
            on_config_saved=self._refresh_available_analytics  # Add this callback
        )

    def _refresh_available_analytics(self):
        """Refresh the list of available analytics"""
        # Reload available analytics
        self.available_analytics = self.config_manager.get_available_analytics()

        # Update the dropdown in the main tab
        self.analytic_combo["values"] = [f"{id} - {name}" for id, name in self.available_analytics]

        # Show a confirmation message
        self.status_var.set("Analytics list refreshed with new configuration")

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

        # Run in a separate thread to avoid freezing the UI
        threading.Thread(target=self._process_data, args=(analytic_id,), daemon=True).start()

    def _process_data(self, analytic_id):
        """Process data in a separate thread"""
        try:
            # Get configuration
            config = self.config_manager.get_config(analytic_id)

            # Create output directory if needed
            output_dir = self.output_var.get()
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # Initialize enhanced processor
            processor = EnhancedDataProcessor(config)

            # Process data
            logger.info(f"Starting processing for QA-ID {analytic_id}")
            success, message = processor.process_data(self.source_var.get())

            if not success:
                self.root.after(0, lambda: messagebox.showerror("Error", message))
                return

            # Generate reports
            logger.info("Generating reports...")
            report_generator = EnhancedReportGenerator(config, processor.results)

            # Generate main report
            main_report_path = os.path.join(
                output_dir,
                f"QA_{analytic_id}_Main_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            main_report = report_generator.generate_main_report(main_report_path)

            # Generate individual reports
            individual_reports = report_generator.generate_individual_reports()

            # Show completion message
            report_count = 1 + len(individual_reports)
            completion_msg = f"Processing complete. Generated {report_count} reports."
            logger.info(completion_msg)

            self.root.after(0, lambda: messagebox.showinfo("Success", completion_msg))
            self.root.after(0, lambda: self.status_var.set("Ready"))

        except Exception as e:
            logger.error(f"Error in processing: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {e}"))

        finally:
            # Stop progress bar
            self.root.after(0, self.progress.stop)
            self.root.after(0, lambda: self.status_var.set("Ready"))

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

"""
Enhanced TestingTab with Excel Formula Integration.

This module provides an enhanced version of the TestingTab with full
integration of the Excel formula testing capabilities.
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pandas as pd
import numpy as np
from typing import Callable, Dict, List, Optional, Any

from qa_analytics.ui.components.formula_tester import FormulaTester
from qa_analytics.core.excel_utils import is_valid_excel_formula, extract_column_names
from qa_analytics.core.excel_engine import ExcelFormulaProcessor
from qa_analytics.utils.logging_config import setup_logging

logger = setup_logging()


class TestingTab(ttk.Frame):
    """
    Enhanced tab for testing analytics configurations with integrated
    Excel formula testing capabilities.
    """

    def __init__(self, parent, status_callback: Callable):
        """
        Initialize the Testing tab with formula testing.

        Args:
            parent: Parent widget
            status_callback: Function to call to update status bar
        """
        super().__init__(parent, padding="20 15 20 15")
        self.parent = parent
        self.update_status = status_callback

        # State variables
        self.analytics_var = tk.StringVar()
        self.data_source_var = tk.StringVar(value="generate")
        self.record_count_var = tk.StringVar(value="100")
        self.error_pct_var = tk.StringVar(value="20")
        self.file_var = tk.StringVar()
        self.filter_var = tk.StringVar(value="all")

        # Data processor and report generator classes
        self.data_processor_class = None
        self.report_generator_class = None

        # Sample data and results storage
        self.sample_data = None
        self.test_results = None

        # Formula tester component reference
        self.formula_tester = None
        self.excel_processor = None

        # Create widgets
        self._create_widgets()

    def _create_widgets(self):
        """Create all widgets for this tab"""
        # Use Grid layout for better control
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)  # Make results section expandable

        # Analytics Selection Section
        selection_frame = ttk.Frame(self)
        selection_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        selection_frame.columnconfigure(1, weight=1)

        ttk.Label(
            selection_frame,
            text="Select Analytics:",
            style="Header.TLabel"
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 15))

        # Load available analytics configurations
        analytics_options = [
            "QA-123 - Data Quality Analysis",
            "QA-77 - Audit Test Workpaper Approvals",
            "QA-78 - Third Party Risk Assessment",
            "QA-99 - Audit Workpaper Review Validation"
        ]

        analytics_combo = ttk.Combobox(
            selection_frame,
            textvariable=self.analytics_var,
            values=analytics_options,
            state="readonly",
            width=50
        )
        if analytics_options:
            analytics_combo.current(0)
        analytics_combo.grid(row=0, column=1, sticky=(tk.W, tk.E))

        # Button to create new Excel formula validation
        formula_test_btn = ttk.Button(
            selection_frame,
            text="Create Excel Formula Validation",
            command=self._create_formula_validation
        )
        formula_test_btn.grid(row=0, column=2, padx=(10, 0))

        # Test Data Options Section
        data_frame = ttk.LabelFrame(self, text="Test Data Options", padding=10)
        data_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        data_frame.columnconfigure(0, weight=1)

        # Option selection (radio buttons)
        option_frame = ttk.Frame(data_frame)
        option_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Radiobutton(
            option_frame,
            text="Generate Sample Data",
            variable=self.data_source_var,
            value="generate",
            command=self._update_data_options
        ).pack(side=tk.LEFT, padx=(0, 30))

        ttk.Radiobutton(
            option_frame,
            text="Use Existing Data",
            variable=self.data_source_var,
            value="existing",
            command=self._update_data_options
        ).pack(side=tk.LEFT)

        # Sample data generation options
        self.sample_frame = ttk.Frame(data_frame)
        self.sample_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(self.sample_frame, text="Number of Records:").pack(side=tk.LEFT)
        ttk.Entry(
            self.sample_frame,
            textvariable=self.record_count_var,
            width=8
        ).pack(side=tk.LEFT, padx=(8, 30))

        ttk.Label(self.sample_frame, text="Error Percentage:").pack(side=tk.LEFT)
        ttk.Entry(
            self.sample_frame,
            textvariable=self.error_pct_var,
            width=8
        ).pack(side=tk.LEFT, padx=(8, 0))

        # File selection options
        self.file_frame = ttk.Frame(data_frame)
        self.file_frame.columnconfigure(1, weight=1)

        ttk.Label(self.file_frame, text="Data File:").grid(row=0, column=0, sticky=tk.W)

        file_input_frame = ttk.Frame(self.file_frame)
        file_input_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(8, 0))
        file_input_frame.columnconfigure(0, weight=1)

        ttk.Entry(
            file_input_frame,
            textvariable=self.file_var,
            width=40
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Add a container for the button to control its size
        button_container = ttk.Frame(file_input_frame, width=40, height=36)
        button_container.pack(side=tk.LEFT, padx=(8, 0))
        button_container.pack_propagate(False)

        ttk.Button(
            button_container,
            text="📂",
            style="Icon.TButton",
            command=self._browse_file
        ).pack(fill=tk.BOTH, expand=True)

        # Run Test Action Section
        action_frame = ttk.Frame(self)
        action_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        action_frame.columnconfigure(0, weight=1)

        # Progress indicator with run button
        progress_frame = ttk.Frame(action_frame)
        progress_frame.pack(fill=tk.X)
        progress_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(
            progress_frame,
            orient=tk.HORIZONTAL,
            mode='indeterminate',
            length=200
        )
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 15))

        self.run_btn = ttk.Button(
            progress_frame,
            text="Run Test",
            style="Primary.TButton",
            command=self._run_test
        )
        self.run_btn.pack(side=tk.RIGHT)

        # Results Section
        results_frame = ttk.LabelFrame(self, text="Test Results", padding=10)
        results_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)

        # Create notebook for result tabs
        self.results_notebook = ttk.Notebook(results_frame)
        self.results_notebook.pack(fill=tk.BOTH, expand=True)

        # Summary tab
        self.summary_tab = ttk.Frame(self.results_notebook, padding=10)
        self.results_notebook.add(self.summary_tab, text="Summary")
        self.summary_tab.columnconfigure(0, weight=1)
        self.summary_tab.rowconfigure(1, weight=1)  # Group summary should expand

        # Create placeholder for summary tab
        self.summary_placeholder = ttk.Label(
            self.summary_tab,
            text="Run a test to see results summary here.",
            style="Info.TLabel"
        )
        self.summary_placeholder.pack(expand=True)

        # Detail tab
        self.detail_tab = ttk.Frame(self.results_notebook, padding=10)
        self.results_notebook.add(self.detail_tab, text="Detail")
        self.detail_tab.columnconfigure(0, weight=1)
        self.detail_tab.rowconfigure(1, weight=1)  # Detail view should expand

        # Create placeholder for detail tab
        self.detail_placeholder = ttk.Label(
            self.detail_tab,
            text="Run a test to see detailed results here.",
            style="Info.TLabel"
        )
        self.detail_placeholder.pack(expand=True)

        # Sample data tab
        self.sample_tab = ttk.Frame(self.results_notebook, padding=10)
        self.results_notebook.add(self.sample_tab, text="Sample Data")
        self.sample_tab.columnconfigure(0, weight=1)
        self.sample_tab.rowconfigure(0, weight=1)  # Sample data should expand

        # Create placeholder for sample data tab
        self.sample_placeholder = ttk.Label(
            self.sample_tab,
            text="Run a test to see sample data here.",
            style="Info.TLabel"
        )
        self.sample_placeholder.pack(expand=True)

        # Export Buttons Section
        export_frame = ttk.Frame(self)
        export_frame.grid(row=4, column=0, sticky=(tk.W, tk.E))

        self.export_sample_btn = ttk.Button(
            export_frame,
            text="Export Sample Data",
            command=self._export_sample_data,
            state=tk.DISABLED  # Initially disabled until we have results
        )
        self.export_sample_btn.pack(side=tk.LEFT)

        self.export_results_btn = ttk.Button(
            export_frame,
            text="Export Results",
            command=self._export_results,
            state=tk.DISABLED  # Initially disabled until we have results
        )
        self.export_results_btn.pack(side=tk.LEFT, padx=(10, 0))

        self.report_btn = ttk.Button(
            export_frame,
            text="Generate Report",
            style="Secondary.TButton",
            command=self._generate_report,
            state=tk.DISABLED  # Initially disabled until we have results
        )
        self.report_btn.pack(side=tk.RIGHT)

        # Update initial data options
        self._update_data_options()

    def _update_data_options(self):
        """Update the data options based on the selected data source option"""
        if self.data_source_var.get() == "generate":
            self.sample_frame.pack(fill=tk.X, pady=(0, 10))
            self.file_frame.pack_forget()
        else:
            self.sample_frame.pack_forget()
            self.file_frame.pack(fill=tk.X, pady=(0, 10))

    def _browse_file(self):
        """Browse for a data file"""
        filename = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("CSV Files", "*.csv"),
                ("All Files", "*.*")
            ]
        )
        if filename:
            self.file_var.set(filename)

    def _create_formula_validation(self):
        """Create a new Excel formula validation in a dialog"""
        # Create dialog window
        dialog = tk.Toplevel(self)
        dialog.title("Excel Formula Validation")
        dialog.geometry("800x700")
        dialog.transient(self)  # Set to be on top of the parent window
        dialog.grab_set()  # Modal dialog

        # Configure dialog
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(0, weight=1)

        # Create main frame with padding
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)  # Formula tester should expand

        # Title and description
        ttk.Label(
            main_frame,
            text="Create and Test Excel Formula Validation",
            style="Header.TLabel"
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 10))

        # Description
        ttk.Label(
            main_frame,
            text=(
                "Create an Excel formula that evaluates to TRUE for records that conform to requirements.\n"
                "Test it with sample data before adding to your configuration."
            ),
            wraplength=700
        ).grid(row=1, column=0, sticky=tk.W, pady=(0, 15))

        # Add FormulaTester component
        self.formula_tester = FormulaTester(main_frame)
        self.formula_tester.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))

        # Buttons at bottom
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

        ttk.Button(
            button_frame,
            text="Close",
            command=dialog.destroy
        ).pack(side=tk.RIGHT, padx=(10, 0))

        ttk.Button(
            button_frame,
            text="Add to Configuration",
            style="Primary.TButton",
            command=lambda: self._add_formula_to_config(dialog)
        ).pack(side=tk.RIGHT)

        # Clean up resources when dialog closes
        dialog.protocol("WM_DELETE_WINDOW", lambda: self._cleanup_formula_tester(dialog))

    def _add_formula_to_config(self, dialog):
        """
        Add the formula to the configuration

        Args:
            dialog: Dialog window to close
        """
        # Get formula and display name
        if not self.formula_tester:
            return

        formula = self.formula_tester.get_formula()
        display_name = self.formula_tester.get_display_name()
        is_valid = self.formula_tester.is_valid()
        fields = self.formula_tester.get_fields_used()

        if not formula:
            messagebox.showinfo("Formula Required", "Please enter a formula")
            return

        if not is_valid:
            if not messagebox.askyesno(
                "Invalid Formula",
                "The formula appears to be invalid. Add it anyway?"
            ):
                return

        # Here you would normally update the config
        # For now, just show a success message
        messagebox.showinfo(
            "Formula Added",
            f"The formula validation '{display_name}' has been created.\n\n"
            f"Formula: {formula}\n"
            f"Fields used: {', '.join(fields)}"
        )

        # Close dialog
        self._cleanup_formula_tester(dialog)

    def _cleanup_formula_tester(self, dialog):
        """
        Clean up formula tester resources and close dialog

        Args:
            dialog: Dialog window to close
        """
        if self.formula_tester:
            self.formula_tester.cleanup()
            self.formula_tester = None

        dialog.destroy()

    def _run_test(self):
        """Run the test with current settings"""
        selected_analytics = self.analytics_var.get()

        # Validate that an analytics configuration is selected
        if not selected_analytics:
            messagebox.showinfo("Analytics Selection", "Please select an analytics configuration")
            return

        # Validate data source
        if self.data_source_var.get() == "existing" and not self.file_var.get():
            messagebox.showinfo("Data Source", "Please select a data file")
            return

        # Start progress
        self.progress.start()
        self.run_btn.config(state=tk.DISABLED)

        # Update status
        self.update_status(f"Running test for {selected_analytics}")

        # Run in a separate thread
        threading.Thread(target=self._execute_test, daemon=True).start()

    def _execute_test(self):
        """Execute the test in a separate thread"""
        try:
            # Generate or load sample data
            if self.data_source_var.get() == "generate":
                self.sample_data = self._generate_sample_data()
            else:
                self.sample_data = self._load_data_file()

            if self.sample_data is None:
                self.after(0, lambda: self.update_status("Failed to prepare test data"))
                return

            # Initialize Excel processor if needed for formulas
            self._init_excel_processor()

            # Run the validation
            self.test_results = self._validate_data(self.sample_data)

            # Update the results UI
            self.after(0, self._update_results_ui)

            # Enable export buttons
            self.after(0, lambda: self.export_sample_btn.config(state=tk.NORMAL))
            self.after(0, lambda: self.export_results_btn.config(state=tk.NORMAL))
            self.after(0, lambda: self.report_btn.config(state=tk.NORMAL))

            # Update status
            self.after(0, lambda: self.update_status("Test completed successfully"))

        except Exception as e:
            # Handle errors
            import traceback
            logger.error(f"Error running test: {traceback.format_exc()}")
            self.after(0, lambda: self.update_status(f"Error running test: {str(e)}"))

        finally:
            # Reset UI
            self.after(0, lambda: self.progress.stop())
            self.after(0, lambda: self.run_btn.config(state=tk.NORMAL))

            # Clean up Excel processor
            self._cleanup_excel_processor()

    def _init_excel_processor(self):
        """Initialize Excel processor if needed for formulas"""
        if self.excel_processor is None:
            try:
                self.excel_processor = ExcelFormulaProcessor(visible=False)
                logger.info("Excel Formula Processor initialized")
            except Exception as e:
                logger.error(f"Error initializing Excel Formula Processor: {e}")
                raise

    def _cleanup_excel_processor(self):
        """Clean up Excel processor resources"""
        if self.excel_processor:
            try:
                self.excel_processor.cleanup()
                self.excel_processor = None
                logger.info("Excel Formula Processor cleaned up")
            except Exception as e:
                logger.warning(f"Error cleaning up Excel Formula Processor: {e}")

    def _generate_sample_data(self):
        """
        Generate sample data for testing.

        Returns:
            DataFrame with sample data or None if an error occurs
        """
        try:
            # Get parameters
            try:
                record_count = int(self.record_count_var.get())
                error_pct = float(self.error_pct_var.get()) / 100
            except ValueError:
                self.after(0, lambda: messagebox.showinfo(
                    "Invalid Input",
                    "Please enter valid numbers for record count and error percentage"
                ))
                return None

            # Create sample data structure based on the selected analytics
            data = {}

            # Get analytics ID from selection
            analytics_id = self.analytics_var.get().split(" - ")[0].replace("QA-", "")

            if analytics_id == "77":  # Audit Test Workpaper Approvals
                # Create basic columns for audit workpaper approvals
                import datetime

                # Workpaper IDs
                data["Audit TW ID"] = [f"TW-{i:06d}" for i in range(1, record_count + 1)]

                # Submitters and approvers
                submitters = ["John Smith", "Emma Johnson", "Michael Brown", "Sarah Davis", "David Wilson"]
                tl_approvers = ["Alex Rodriguez", "Michelle Lee", "Richard White", "Patricia Moore", "James Martin"]
                al_approvers = ["William Thomas", "Elizabeth Thompson", "Charles Garcia", "Susan Clark", "Joseph Lewis"]

                data["TW submitter"] = np.random.choice(submitters, record_count)
                data["TL approver"] = np.random.choice(tl_approvers, record_count)
                data["AL approver"] = np.random.choice(al_approvers, record_count)

                # Dates
                base_date = datetime.datetime.now() - datetime.timedelta(days=30)

                data["Submit Date"] = [base_date + datetime.timedelta(days=np.random.randint(0, 10))
                                       for _ in range(record_count)]

                # For valid records (1 - error_pct)
                valid_count = int(record_count * (1 - error_pct))
                error_count = record_count - valid_count

                # TL approval dates
                tl_dates = []
                for i in range(record_count):
                    if i < valid_count:
                        # Valid: TL date after submit date
                        tl_dates.append(data["Submit Date"][i] + datetime.timedelta(days=np.random.randint(1, 5)))
                    else:
                        # Invalid: TL date before submit date (for some error records)
                        if np.random.random() < 0.5:
                            tl_dates.append(data["Submit Date"][i] - datetime.timedelta(days=np.random.randint(1, 5)))
                        else:
                            # Another type of error: TL is same as submitter
                            data["TL approver"][i] = data["TW submitter"][i]
                            tl_dates.append(data["Submit Date"][i] + datetime.timedelta(days=np.random.randint(1, 5)))

                data["TL Approval Date"] = tl_dates

                # AL approval dates
                al_dates = []
                for i in range(record_count):
                    if i < valid_count:
                        # Valid: AL date after TL date
                        al_dates.append(tl_dates[i] + datetime.timedelta(days=np.random.randint(1, 5)))
                    else:
                        # For error records not already with date errors, make AL same as TL
                        if data["Submit Date"][i] < tl_dates[i]:
                            al_dates.append(tl_dates[i])
                        else:
                            # If already has date error, make AL date valid
                            al_dates.append(tl_dates[i] + datetime.timedelta(days=np.random.randint(1, 5)))

                data["AL Approval Date"] = al_dates

            elif analytics_id == "78":  # Third Party Risk Assessment
                # Create basic columns for third party risk assessment
                # Assessment IDs and names
                data["Assessment ID"] = [f"RA-{i:06d}" for i in range(1, record_count + 1)]
                data["Assessment Name"] = [f"Risk Assessment {i}" for i in range(1, record_count + 1)]

                # Owners
                owners = ["John Manager", "Emma Director", "Michael Leader", "Sarah Executive", "David Officer"]
                data["Assessment Owner"] = np.random.choice(owners, record_count)

                # Third Party Vendors
                vendors = [
                    "",  # Empty for some records
                    "Vendor A",
                    "Vendor B",
                    "Vendor C",
                    "Vendor A, Vendor B",
                    "Vendor A, Vendor C",
                    "Vendor B, Vendor C"
                ]

                data["Third Party Vendors"] = np.random.choice(vendors, record_count)

                # Risk Ratings
                ratings = ["Critical", "High", "Medium", "Low", "N/A"]

                # For valid records (1 - error_pct)
                valid_count = int(record_count * (1 - error_pct))

                risk_ratings = []
                for i in range(record_count):
                    if i < valid_count:
                        # Valid: If there are vendors, risk should not be N/A
                        if data["Third Party Vendors"][i]:
                            risk_ratings.append(np.random.choice(ratings[:4]))  # Non-N/A
                        else:
                            risk_ratings.append("N/A")
                    else:
                        # Invalid: If there are vendors, risk is N/A, or if no vendors, risk is not N/A
                        if data["Third Party Vendors"][i]:
                            risk_ratings.append("N/A")
                        else:
                            risk_ratings.append(np.random.choice(ratings[:4]))  # Non-N/A

                data["Vendor Risk Rating"] = risk_ratings

            else:
                # Generic sample data for other analytics types
                data["ID"] = [f"ID-{i:06d}" for i in range(1, record_count + 1)]
                data["Name"] = [f"Item {i}" for i in range(1, record_count + 1)]
                data["Status"] = ["Active" if i % 5 != 0 else "Inactive" for i in range(1, record_count + 1)]
                data["Value"] = [round(100 * i / record_count, 2) for i in range(1, record_count + 1)]

                # Add some random valid/invalid records markers
                import datetime

                data["Date"] = [datetime.datetime.now() - datetime.timedelta(days=i % 30)
                                for i in range(1, record_count + 1)]

                data["IsValid"] = [True if i < (record_count * (1 - error_pct)) else False
                                   for i in range(1, record_count + 1)]

            # Convert to DataFrame
            df = pd.DataFrame(data)
            return df

        except Exception as e:
            import traceback
            logger.error(f"Error generating sample data: {traceback.format_exc()}")
            self.after(0, lambda: self.update_status(f"Error generating sample data: {str(e)}"))
            return None

    def _load_data_file(self):
        """
        Load data from file for testing.

        Returns:
            DataFrame with loaded data or None if an error occurs
        """
        file_path = self.file_var.get()

        if not file_path or not os.path.exists(file_path):
            self.after(0, lambda: messagebox.showinfo("File Selection", "Please select a valid file"))
            return None

        try:
            # Determine file type
            if file_path.lower().endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path)
            elif file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                self.after(0, lambda: messagebox.showinfo("File Error", "Unsupported file type"))
                return None

            return df

        except Exception as e:
            self.after(0, lambda: messagebox.showinfo("File Error", f"Failed to load file: {str(e)}"))
            return None

    def _validate_data(self, data):
        """
        Validate the data based on the selected analytics configuration.
        In a real implementation, this would use the actual validation rules.

        Args:
            data: DataFrame to validate

        Returns:
            Dictionary with validation results
        """
        # Get analytics ID
        analytics_id = self.analytics_var.get().split(" - ")[0].replace("QA-", "")

        try:
            # Apply validation rules based on analytics type
            if analytics_id == "77":  # Audit Test Workpaper Approvals
                # Rule 1: Submitter cannot be TL or AL (segregation of duties)
                segregation_rule = (data["TW submitter"] != data["TL approver"]) & (
                        data["TW submitter"] != data["AL approver"])

                # Rule 2: Dates in sequence (Submit -> TL -> AL)
                sequence_rule = (data["Submit Date"] <= data["TL Approval Date"]) & (
                        data["TL Approval Date"] <= data["AL Approval Date"])

                # Combined compliance
                compliance = segregation_rule & sequence_rule

                # Create compliance column
                data["Compliance"] = compliance.map({True: "GC", False: "DNC"})

                # Prepare summary by approver
                summary = data.groupby("AL approver").agg(
                    GC=("Compliance", lambda x: sum(x == "GC")),
                    DNC=("Compliance", lambda x: sum(x == "DNC")),
                    Total=("Compliance", "count")
                ).reset_index()

                # Calculate percentages
                summary["DNC_Percentage"] = (summary["DNC"] / summary["Total"] * 100).round(2)

                # Add compliance status column
                threshold = 5.0  # Default threshold
                summary["Exceeds_Threshold"] = summary["DNC_Percentage"] > threshold

                return {
                    "summary": summary,
                    "detail": data
                }

            elif analytics_id == "78":  # Third Party Risk Assessment
                # Rule: If Third Party Vendors exists, Risk Rating should not be N/A
                # And if no vendors, Risk Rating should be N/A
                has_vendors = data["Third Party Vendors"] != ""
                has_risk = data["Vendor Risk Rating"] != "N/A"

                compliance = (has_vendors & has_risk) | (~has_vendors & ~has_risk)

                # Create compliance column
                data["Compliance"] = compliance.map({True: "GC", False: "DNC"})

                # Prepare summary by owner
                summary = data.groupby("Assessment Owner").agg(
                    GC=("Compliance", lambda x: sum(x == "GC")),
                    DNC=("Compliance", lambda x: sum(x == "DNC")),
                    Total=("Compliance", "count")
                ).reset_index()

                # Calculate percentages
                summary["DNC_Percentage"] = (summary["DNC"] / summary["Total"] * 100).round(2)

                # Add compliance status column
                threshold = 5.0  # Default threshold
                summary["Exceeds_Threshold"] = summary["DNC_Percentage"] > threshold

                return {
                    "summary": summary,
                    "detail": data
                }

            else:
                # Generic validation for other analytics types
                if "IsValid" in data.columns:
                    data["Compliance"] = data["IsValid"].map({True: "GC", False: "DNC"})
                else:
                    # Random validation for demo purposes
                    valid_count = int(len(data) * (1 - float(self.error_pct_var.get()) / 100))
                    compliance = ["GC"] * valid_count + ["DNC"] * (len(data) - valid_count)
                    np.random.shuffle(compliance)
                    data["Compliance"] = compliance

                # Use first string column as group by field
                group_col = next((col for col in data.columns if data[col].dtype == 'object'
                                  and col not in ["Compliance", "ID"]), data.columns[0])

                # Prepare summary
                summary = data.groupby(group_col).agg(
                    GC=("Compliance", lambda x: sum(x == "GC")),
                    DNC=("Compliance", lambda x: sum(x == "DNC")),
                    Total=("Compliance", "count")
                ).reset_index()

                # Calculate percentages
                summary["DNC_Percentage"] = (summary["DNC"] / summary["Total"] * 100).round(2)

                # Add compliance status column
                threshold = 5.0  # Default threshold
                summary["Exceeds_Threshold"] = summary["DNC_Percentage"] > threshold

                return {
                    "summary": summary,
                    "detail": data
                }

        except Exception as e:
            import traceback
            logger.error(f"Error validating data: {traceback.format_exc()}")
            self.after(0, lambda: self.update_status(f"Error validating data: {str(e)}"))
            return None

    def _update_results_ui(self):
        """Update the UI with test results"""
        if not self.test_results:
            return

        # Clear placeholders
        for placeholder in [self.summary_placeholder, self.detail_placeholder, self.sample_placeholder]:
            if placeholder.winfo_exists():
                placeholder.destroy()

        # Summary tab
        self._update_summary_tab()

        # Detail tab
        self._update_detail_tab()

        # Sample data tab
        self._update_sample_tab()

    def _update_summary_tab(self):
        """Update the summary tab with test results"""
        summary_data = self.test_results.get("summary")
        if summary_data is None:
            return

        # Clear existing widgets
        for widget in self.summary_tab.winfo_children():
            widget.destroy()

        # Summary statistics section
        stats_frame = ttk.LabelFrame(self.summary_tab, text="Results Summary", padding=10)
        stats_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))

        # Calculate overall statistics
        detail_data = self.test_results.get("detail")
        if detail_data is not None and "Compliance" in detail_data:
            total = len(detail_data)
            gc_count = (detail_data["Compliance"] == "GC").sum()
            dnc_count = (detail_data["Compliance"] == "DNC").sum()
            pc_count = (detail_data["Compliance"] == "PC").sum() if "PC" in detail_data["Compliance"].values else 0

            error_pct = (dnc_count / total * 100) if total > 0 else 0

            # Create grid for stats
            stats_grid = ttk.Frame(stats_frame)
            stats_grid.pack(fill=tk.X, padx=10, pady=5)

            row = 0
            # Total Records
            ttk.Label(stats_grid, text="Total Records:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0, 20))
            ttk.Label(stats_grid, text=str(total)).grid(row=row, column=1, sticky=tk.W, pady=5)
            row += 1

            # Generally Conforms
            ttk.Label(stats_grid, text="Generally Conforms (GC):").grid(row=row, column=0, sticky=tk.W, pady=5,
                                                                        padx=(0, 20))
            ttk.Label(
                stats_grid,
                text=f"{gc_count} ({gc_count / total * 100:.1f}%)",
                style="Success.TLabel"
            ).grid(row=row, column=1, sticky=tk.W, pady=5)
            row += 1

            # Does Not Conform
            ttk.Label(stats_grid, text="Does Not Conform (DNC):").grid(row=row, column=0, sticky=tk.W, pady=5,
                                                                       padx=(0, 20))
            ttk.Label(
                stats_grid,
                text=f"{dnc_count} ({dnc_count / total * 100:.1f}%)",
                style="Error.TLabel" if dnc_count > 0 else None
            ).grid(row=row, column=1, sticky=tk.W, pady=5)
            row += 1

            # Partially Conforms (if applicable)
            if pc_count > 0:
                ttk.Label(stats_grid, text="Partially Conforms (PC):").grid(row=row, column=0, sticky=tk.W, pady=5,
                                                                            padx=(0, 20))
                ttk.Label(
                    stats_grid,
                    text=f"{pc_count} ({pc_count / total * 100:.1f}%)",
                    style="Warning.TLabel"
                ).grid(row=row, column=1, sticky=tk.W, pady=5)
                row += 1

            # Error Percentage
            ttk.Label(stats_grid, text="Error Percentage:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0, 20))
            ttk.Label(
                stats_grid,
                text=f"{error_pct:.2f}%",
                style="Error.TLabel" if error_pct > 5.0 else "Success.TLabel"
            ).grid(row=row, column=1, sticky=tk.W, pady=5)

            # Threshold indicator
            threshold = 5.0
            threshold_frame = ttk.Frame(stats_frame, padding=(10, 10, 10, 5))
            threshold_frame.pack(fill=tk.X)

            if error_pct > threshold:
                ttk.Label(
                    threshold_frame,
                    text=f"⚠️ Error percentage exceeds threshold of {threshold:.1f}%",
                    style="Error.TLabel"
                ).pack(side=tk.LEFT)
            else:
                ttk.Label(
                    threshold_frame,
                    text=f"✓ Error percentage is below threshold of {threshold:.1f}%",
                    style="Success.TLabel"
                ).pack(side=tk.LEFT)

        # Group summary section
        if not summary_data.empty:
            group_frame = ttk.LabelFrame(self.summary_tab, text="Group Summary", padding=10)
            group_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
            group_frame.columnconfigure(0, weight=1)
            group_frame.rowconfigure(0, weight=1)

            # Create container for treeview and scrollbar
            tree_container = ttk.Frame(group_frame)
            tree_container.pack(fill=tk.BOTH, expand=True)
            tree_container.columnconfigure(0, weight=1)
            tree_container.rowconfigure(0, weight=1)

            # Create treeview for group data
            tree_columns = list(summary_data.columns)

            group_tree = ttk.Treeview(
                tree_container,
                columns=tree_columns,
                show="headings",
                selectmode="browse"
            )
            group_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

            # Add scrollbars
            y_scroll = ttk.Scrollbar(
                tree_container,
                orient="vertical",
                command=group_tree.yview,
                style="Vertical.TScrollbar"
            )
            y_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))

            x_scroll = ttk.Scrollbar(
                tree_container,
                orient="horizontal",
                command=group_tree.xview,
                style="Horizontal.TScrollbar"
            )
            x_scroll.grid(row=1, column=0, sticky=(tk.W, tk.E))

            group_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

            # Configure columns
            col_width = 100  # Default width

            for col in tree_columns:
                if col == "Exceeds_Threshold":
                    # Don't show this column
                    group_tree.column(col, width=0, stretch=False)
                    group_tree.heading(col, text="")
                else:
                    # Format the column name for display
                    display_name = col.replace("_", " ")
                    if col == "DNC_Percentage":
                        display_name = "Error %"

                    anchor = tk.CENTER if col in ["GC", "DNC", "PC", "Total", "DNC_Percentage"] else tk.W
                    group_tree.column(col, width=col_width, anchor=anchor)
                    group_tree.heading(col, text=display_name)

            # Configure tags for status colors
            group_tree.tag_configure("exceeds", background="#FFE6E6")  # Light red
            group_tree.tag_configure("within", background="#E6FFE6")  # Light green

            # Add data to tree
            for _, row in summary_data.iterrows():
                values = [row[col] if col != "DNC_Percentage" else f"{row[col]:.2f}%" for col in tree_columns]

                item_id = group_tree.insert("", tk.END, values=values)

                # Apply color tag based on threshold status
                if row.get("Exceeds_Threshold", False):
                    group_tree.item(item_id, tags=("exceeds",))
                else:
                    group_tree.item(item_id, tags=("within",))

    def _update_detail_tab(self):
        """Update the detail tab with test results"""
        detail_data = self.test_results.get("detail")
        if detail_data is None:
            return

        # Clear existing widgets
        for widget in self.detail_tab.winfo_children():
            widget.destroy()

        # Create filter section
        filter_frame = ttk.Frame(self.detail_tab)
        filter_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))

        ttk.Label(filter_frame, text="Show:").pack(side=tk.LEFT, padx=(0, 10))

        ttk.Radiobutton(
            filter_frame,
            text="All Records",
            variable=self.filter_var,
            value="all",
            command=self._apply_filter
        ).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Radiobutton(
            filter_frame,
            text="GC Only",
            variable=self.filter_var,
            value="gc",
            command=self._apply_filter
        ).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Radiobutton(
            filter_frame,
            text="DNC Only",
            variable=self.filter_var,
            value="dnc",
            command=self._apply_filter
        ).pack(side=tk.LEFT)

        # Add PC filter if there are any PC records
        if "PC" in detail_data["Compliance"].values:
            ttk.Radiobutton(
                filter_frame,
                text="PC Only",
                variable=self.filter_var,
                value="pc",
                command=self._apply_filter
            ).pack(side=tk.LEFT, padx=(15, 0))

        # Create container for the treeview
        detail_container = ttk.Frame(self.detail_tab)
        detail_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        detail_container.columnconfigure(0, weight=1)
        detail_container.rowconfigure(0, weight=1)

        # Get columns with priority on Compliance column
        columns = ["Compliance"] + [col for col in detail_data.columns if col != "Compliance"]

        self.detail_tree = ttk.Treeview(
            detail_container,
            columns=columns,
            show="headings",
            selectmode="browse"
        )
        self.detail_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Add scrollbars
        y_scroll = ttk.Scrollbar(
            detail_container,
            orient="vertical",
            command=self.detail_tree.yview,
            style="Vertical.TScrollbar"
        )
        y_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))

        x_scroll = ttk.Scrollbar(
            detail_container,
            orient="horizontal",
            command=self.detail_tree.xview,
            style="Horizontal.TScrollbar"
        )
        x_scroll.grid(row=1, column=0, sticky=(tk.W, tk.E))

        self.detail_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        # Configure columns
        col_width = 100  # Default width

        for col in columns:
            anchor = tk.CENTER if col in ["Compliance", "GC", "DNC", "PC"] else tk.W
            self.detail_tree.column(col, width=col_width, anchor=anchor)
            self.detail_tree.heading(col, text=col)

        # Configure tags for status colors
        self.detail_tree.tag_configure("gc", background="#E6FFE6")  # Light green
        self.detail_tree.tag_configure("dnc", background="#FFE6E6")  # Light red
        self.detail_tree.tag_configure("pc", background="#FFF8E6")  # Light yellow

        # Apply initial filter
        self._apply_filter()

    def _apply_filter(self):
        """Apply filter to detail data"""
        if not hasattr(self, 'detail_tree') or not self.test_results:
            return

        detail_data = self.test_results.get("detail")
        if detail_data is None:
            return

        # Clear existing items
        for item in self.detail_tree.get_children():
            self.detail_tree.delete(item)

        # Apply filter
        filter_value = self.filter_var.get()

        if filter_value == "all":
            filtered_data = detail_data
        elif filter_value == "gc":
            filtered_data = detail_data[detail_data["Compliance"] == "GC"]
        elif filter_value == "dnc":
            filtered_data = detail_data[detail_data["Compliance"] == "DNC"]
        elif filter_value == "pc":
            filtered_data = detail_data[detail_data["Compliance"] == "PC"]
        else:
            filtered_data = detail_data

        # Get columns
        columns = [col for col in self.detail_tree["columns"]]

        # Add rows to treeview
        for idx, row in filtered_data.iterrows():
            values = []
            for col in columns:
                val = row.get(col, "")

                # Format date values
                if pd.api.types.is_datetime64_any_dtype(val) or isinstance(val, pd.Timestamp):
                    val = val.strftime('%Y-%m-%d %H:%M')

                values.append(val)

            item_id = self.detail_tree.insert("", tk.END, values=values)

            # Apply color tag based on compliance
            compliance = row.get("Compliance")
            if compliance == "GC":
                self.detail_tree.item(item_id, tags=("gc",))
            elif compliance == "DNC":
                self.detail_tree.item(item_id, tags=("dnc",))
            elif compliance == "PC":
                self.detail_tree.item(item_id, tags=("pc",))

    def _update_sample_tab(self):
        """Update the sample data tab"""
        if self.sample_data is None:
            return

        # Clear existing widgets
        for widget in self.sample_tab.winfo_children():
            widget.destroy()

        # Create container for the treeview
        sample_container = ttk.Frame(self.sample_tab)
        sample_container.pack(fill=tk.BOTH, expand=True)
        sample_container.columnconfigure(0, weight=1)
        sample_container.rowconfigure(0, weight=1)

        # Get columns
        columns = list(self.sample_data.columns)

        sample_tree = ttk.Treeview(
            sample_container,
            columns=columns,
            show="headings",
            selectmode="browse"
        )

        # Add scrollbars
        y_scroll = ttk.Scrollbar(
            sample_container,
            orient="vertical",
            command=sample_tree.yview,
            style="Vertical.TScrollbar"
        )

        x_scroll = ttk.Scrollbar(
            sample_container,
            orient="horizontal",
            command=sample_tree.xview,
            style="Horizontal.TScrollbar"
        )

        # Position tree and scrollbars
        sample_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        y_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        x_scroll.grid(row=1, column=0, sticky=(tk.W, tk.E))

        sample_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        # Configure columns
        col_width = 100  # Default width

        for col in columns:
            sample_tree.column(col, width=col_width)
            sample_tree.heading(col, text=col)

        # Add rows to treeview
        for idx, row in self.sample_data.iterrows():
            values = []
            for col in columns:
                val = row.get(col, "")

                # Format date values
                if pd.api.types.is_datetime64_any_dtype(val) or isinstance(val, pd.Timestamp):
                    val = val.strftime('%Y-%m-%d %H:%M')

                values.append(val)

            sample_tree.insert("", tk.END, values=values)

    def _export_sample_data(self):
        """Export the sample data to a file"""
        if self.sample_data is None:
            messagebox.showinfo("Export Error", "No sample data to export")
            return

        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Files", "*.xlsx"),
                ("CSV Files", "*.csv"),
                ("All Files", "*.*")
            ],
            title="Export Sample Data"
        )

        if not file_path:
            return

        try:
            # Save data
            if file_path.endswith('.csv'):
                self.sample_data.to_csv(file_path, index=False)
            else:
                self.sample_data.to_excel(file_path, index=False)

            self.update_status(f"Sample data exported to {file_path}")

        except Exception as e:
            messagebox.showerror("Export Error", f"Error exporting sample data: {str(e)}")

    def _export_results(self):
        """Export the test results to a file"""
        if not self.test_results:
            messagebox.showinfo("Export Error", "No test results to export")
            return

        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Files", "*.xlsx"),
                ("All Files", "*.*")
            ],
            title="Export Test Results"
        )

        if not file_path:
            return

        try:
            # Save data to Excel with multiple sheets
            with pd.ExcelWriter(file_path) as writer:
                if 'summary' in self.test_results and self.test_results['summary'] is not None:
                    self.test_results['summary'].to_excel(writer, sheet_name='Summary', index=False)

                if 'detail' in self.test_results and self.test_results['detail'] is not None:
                    self.test_results['detail'].to_excel(writer, sheet_name='Detail', index=False)

            self.update_status(f"Test results exported to {file_path}")

        except Exception as e:
            messagebox.showerror("Export Error", f"Error exporting test results: {str(e)}")

    def _generate_report(self):
        """Generate a formatted report"""
        if not self.test_results:
            messagebox.showinfo("Report Generation", "No test results for report")
            return

        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Files", "*.xlsx"),
                ("All Files", "*.*")
            ],
            title="Save Report"
        )

        if not file_path:
            return

        try:
            # Get analytics details
            analytics_id = self.analytics_var.get().split(" - ")[0].replace("QA-", "")
            analytics_name = self.analytics_var.get().split(" - ")[1] if " - " in self.analytics_var.get() else ""

            # Create config to simulate real report generation
            config = {
                'analytic_id': analytics_id,
                'analytic_name': analytics_name,
                'thresholds': {
                    'error_percentage': 5.0
                },
                'reporting': {
                    'group_by': 'Default Group'
                }
            }

            # If the report generator class is set, use it
            if self.report_generator_class:
                # Initialize report generator
                report_generator = self.report_generator_class(config, self.test_results)

                # Generate report
                report_path = report_generator.generate_main_report(output_path=file_path)

                if report_path:
                    self.update_status(f"Report generated at {report_path}")
                    messagebox.showinfo(
                        "Report Generated",
                        f"Report has been successfully generated at:\n{report_path}"
                    )
                else:
                    messagebox.showerror("Report Error", "Failed to generate report")
            else:
                # Basic report generation without the report generator class
                with pd.ExcelWriter(file_path) as writer:
                    # Write summary sheet
                    if 'summary' in self.test_results and self.test_results['summary'] is not None:
                        self.test_results['summary'].to_excel(writer, sheet_name='Summary', index=False)

                    # Write detail sheet
                    if 'detail' in self.test_results and self.test_results['detail'] is not None:
                        self.test_results['detail'].to_excel(writer, sheet_name='Detail', index=False)

                    # Create configuration sheet data
                    import datetime

                    config_data = [
                        {'Parameter': 'Analytic ID', 'Value': analytics_id},
                        {'Parameter': 'Analytic Name', 'Value': analytics_name},
                        {'Parameter': 'Run Date', 'Value': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
                        {'Parameter': 'Threshold (%)', 'Value': "5.0"},
                        {'Parameter': '--- TEST INFORMATION ---', 'Value': ''},
                        {'Parameter': 'Test Mode',
                         'Value': 'Generated Data' if self.data_source_var.get() == 'generate' else 'Existing Data'},
                        {'Parameter': 'Record Count',
                         'Value': len(self.sample_data) if self.sample_data is not None else 0},
                        {'Parameter': 'Error Percentage',
                         'Value': self.error_pct_var.get() if self.data_source_var.get() == 'generate' else 'N/A'}
                    ]

                    # Write configuration data
                    pd.DataFrame(config_data).to_excel(writer, sheet_name='Configuration', index=False)

                self.update_status(f"Report generated at {file_path}")

                # Show success message
                messagebox.showinfo(
                    "Report Generated",
                    f"Report has been successfully generated at:\n{file_path}"
                )

        except Exception as e:
            messagebox.showerror("Report Error", f"Error generating report: {str(e)}")

    def cleanup(self):
        """Clean up resources"""
        self._cleanup_excel_processor()
        if self.formula_tester:
            self.formula_tester.cleanup()
            self.formula_tester = None

    def __del__(self):
        """Destructor to ensure resources are cleaned up"""
        self.cleanup()
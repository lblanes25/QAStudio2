"""
FormulaTester Component for QA Analytics Framework.

This module provides a reusable UI component for testing Excel formulas
against sample data. It allows users to:
1. Enter and validate Excel formulas
2. Generate sample data or load existing data for testing
3. View formula results in real-time
4. Get detailed feedback on formula syntax and execution
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pandas as pd
import numpy as np
from typing import Callable, Dict, List, Optional, Any

from qa_analytics.core.excel_utils import (
    is_valid_excel_formula, 
    extract_column_names, 
    simplify_formula,
    get_excel_formula_description
)
from qa_analytics.core.excel_engine import ExcelFormulaProcessor
from qa_analytics.utils.logging_config import setup_logging

logger = setup_logging()


class FormulaTester(ttk.Frame):
    """
    Reusable UI component for testing Excel formulas.
    
    This component provides a complete interface for entering, validating,
    and testing Excel formulas against real or generated data.
    """

    def __init__(self, parent, callback: Optional[Callable] = None, 
                 initial_formula: str = "", description: str = "Formula Validation"):
        """
        Initialize the FormulaTester component.
        
        Args:
            parent: Parent widget
            callback: Optional callback function to receive formula changes
                     Called with (formula, display_name, is_valid, fields)
            initial_formula: Initial formula to display
            description: Initial display name for the formula
        """
        super().__init__(parent, padding=10)
        self.parent = parent
        self.callback = callback
        
        # Initialize state variables
        self.formula_var = tk.StringVar(value=initial_formula)
        self.display_name_var = tk.StringVar(value=description)
        self.data_source_var = tk.StringVar(value="generate")
        self.record_count_var = tk.StringVar(value="100")
        self.error_pct_var = tk.StringVar(value="20")
        self.file_var = tk.StringVar()
        self.formula_status_var = tk.StringVar(value="Enter a formula")
        
        # Excel processor instance will be created on-demand
        self.excel_processor = None
        self.sample_data = None
        self.formula_result = None
        self.fields_used = set()
        self.is_formula_valid = False
        
        # Setup the UI
        self._create_widgets()
        
        # Set up formula change callback
        self.formula_var.trace_add("write", self._on_formula_changed)
        
        # If there's an initial formula, validate it
        if initial_formula:
            self._validate_formula(initial_formula)

    def _create_widgets(self):
        """Create all widgets for the formula tester"""
        # Configure grid layout
        self.columnconfigure(0, weight=1)
        
        # Create frame for formula input
        formula_frame = ttk.LabelFrame(self, text="Excel Formula", padding=10)
        formula_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        formula_frame.columnconfigure(1, weight=1)
        
        # Formula input
        ttk.Label(formula_frame, text="Formula:").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 5), pady=(0, 5))
        
        formula_entry = ttk.Entry(
            formula_frame, 
            textvariable=self.formula_var,
            width=60
        )
        formula_entry.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Display name input
        ttk.Label(formula_frame, text="Display Name:").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(0, 5))
        
        ttk.Entry(
            formula_frame,
            textvariable=self.display_name_var,
            width=60
        ).grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Formula status indicator
        ttk.Label(formula_frame, text="Status:").grid(
            row=2, column=0, sticky=tk.W, padx=(0, 5), pady=(0, 5))
        
        status_label = ttk.Label(
            formula_frame,
            textvariable=self.formula_status_var,
            foreground="gray"
        )
        status_label.grid(row=2, column=1, sticky=tk.W, pady=(0, 5))
        
        # Create frame for test data options
        test_frame = ttk.LabelFrame(self, text="Test Data", padding=10)
        test_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        test_frame.columnconfigure(0, weight=1)
        
        # Test data options
        options_frame = ttk.Frame(test_frame)
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Radiobutton(
            options_frame,
            text="Generate Sample Data",
            variable=self.data_source_var,
            value="generate",
            command=self._update_data_options
        ).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Radiobutton(
            options_frame,
            text="Use Existing Data",
            variable=self.data_source_var,
            value="existing",
            command=self._update_data_options
        ).pack(side=tk.LEFT)
        
        # Sample data options frame
        self.sample_frame = ttk.Frame(test_frame)
        self.sample_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(self.sample_frame, text="Records:").pack(side=tk.LEFT)
        ttk.Entry(
            self.sample_frame,
            textvariable=self.record_count_var,
            width=8
        ).pack(side=tk.LEFT, padx=(5, 20))
        
        ttk.Label(self.sample_frame, text="Error %:").pack(side=tk.LEFT)
        ttk.Entry(
            self.sample_frame,
            textvariable=self.error_pct_var,
            width=8
        ).pack(side=tk.LEFT, padx=(5, 0))
        
        # File selection frame
        self.file_frame = ttk.Frame(test_frame)
        ttk.Label(self.file_frame, text="Data File:").pack(side=tk.LEFT)
        ttk.Entry(
            self.file_frame,
            textvariable=self.file_var,
            width=40
        ).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Button(
            self.file_frame,
            text="Browse...",
            command=self._browse_file
        ).pack(side=tk.LEFT)
        
        # Test button and progress bar
        buttons_frame = ttk.Frame(test_frame)
        buttons_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.progress_bar = ttk.Progressbar(
            buttons_frame,
            orient="horizontal",
            mode="indeterminate",
            length=200
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        self.test_btn = ttk.Button(
            buttons_frame,
            text="Test Formula",
            command=self._test_formula
        )
        self.test_btn.pack(side=tk.RIGHT)
        
        # Create frame for results
        self.results_frame = ttk.LabelFrame(self, text="Test Results")
        
        # Apply initial data source option
        self._update_data_options()

    def _update_data_options(self):
        """Update test data options based on selected option"""
        data_source = self.data_source_var.get()
        
        if data_source == "generate":
            if hasattr(self, 'file_frame'):
                self.file_frame.pack_forget()
            self.sample_frame.pack(fill=tk.X, pady=(0, 10))
        else:
            self.sample_frame.pack_forget()
            self.file_frame.pack(fill=tk.X, pady=(0, 10))

    def _browse_file(self):
        """Open file dialog to select data file"""
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

    def _on_formula_changed(self, *args):
        """Handle formula changes and validate in real-time"""
        formula = self.formula_var.get()
        
        # Update formula validation
        self._validate_formula(formula)
        
        # Call the callback if provided
        if self.callback and callable(self.callback):
            self.callback(
                formula, 
                self.display_name_var.get(),
                self.is_formula_valid,
                self.fields_used
            )

    def _validate_formula(self, formula: str):
        """Validate Excel formula and update status"""
        if not formula:
            self._update_status("Enter a formula", "gray")
            self.is_formula_valid = False
            self.fields_used = set()
            return
        
        try:
            # Ensure formula starts with equals sign
            if not formula.startswith('='):
                formula = f"={formula}"
            
            # Validate the formula using excel_utils
            is_valid = is_valid_excel_formula(formula)
            
            if is_valid:
                # Extract fields used in the formula
                fields_used = extract_column_names(formula)
                self.fields_used = fields_used
                
                # Get simplified version (optional)
                simplified_formula = simplify_formula(formula)
                
                # Get formula description
                formula_desc = get_excel_formula_description(formula)
                
                # Update status
                if fields_used:
                    fields_str = ", ".join(f"'{f}'" for f in fields_used)
                    self._update_status(
                        f"Valid formula using {fields_str}",
                        "green"
                    )
                else:
                    self._update_status(f"Valid formula: {formula_desc}", "green")
                
                self.is_formula_valid = True
            else:
                # Update status
                self._update_status("Invalid formula syntax", "red")
                self.is_formula_valid = False
                self.fields_used = set()
        
        except Exception as e:
            # Update status
            self._update_status(f"Error: {str(e)}", "red")
            self.is_formula_valid = False
            self.fields_used = set()
            logger.error(f"Error validating formula: {e}")

    def _update_status(self, message: str, color: str = "black"):
        """Update formula status message and color"""
        self.formula_status_var.set(message)
        
        # Find the status label and update its foreground color
        for child in self.winfo_children():
            if isinstance(child, ttk.LabelFrame) and child.winfo_children():
                for widget in child.winfo_children():
                    if isinstance(widget, ttk.Label) and widget.cget('textvariable') == str(self.formula_status_var):
                        widget.configure(foreground=color)
                        break

    def _test_formula(self):
        """Test the formula with sample or existing data"""
        formula = self.formula_var.get()
        
        # Validate formula first
        if not formula:
            messagebox.showinfo("Formula Required", "Please enter a formula to test")
            return
        
        # Ensure formula starts with equals sign
        if not formula.startswith('='):
            formula = f"={formula}"
        
        # Verify formula is valid
        if not is_valid_excel_formula(formula):
            messagebox.showinfo("Invalid Formula", "Please enter a valid Excel formula")
            return
        
        # Start progress bar
        self.progress_bar.start()
        self.test_btn.config(state=tk.DISABLED)
        
        # Run test in a separate thread to keep UI responsive
        threading.Thread(
            target=self._run_formula_test,
            args=(formula,),
            daemon=True
        ).start()

    def _run_formula_test(self, formula: str):
        """
        Run formula test in a separate thread
        
        Args:
            formula: Excel formula to test
        """
        try:
            # Initialize Excel processor if needed
            if not self.excel_processor:
                self.excel_processor = ExcelFormulaProcessor(visible=False)
            
            # Prepare data for testing
            if self.data_source_var.get() == "generate":
                self._generate_sample_data()
            else:
                self._load_data_file()
            
            # Check if data was successfully prepared
            if self.sample_data is None:
                self._finish_test(False, "Failed to prepare test data")
                return
            
            # Create a result column name based on the display name
            result_column = self.display_name_var.get()
            if not result_column:
                result_column = "Formula_Result"
            
            # Set up formulas dictionary for Excel processor
            formulas = {result_column: formula}
            
            # Process the data with Excel formula
            result_df, warnings = self.excel_processor.process_data_with_formulas(
                self.sample_data, formulas
            )
            
            # Process warnings if any
            if warnings:
                for warning in warnings:
                    logger.warning(f"Excel formula warning: {warning}")
            
            if result_df is None:
                self._finish_test(False, "Excel formula processing failed")
                return
            
            # Extract the results column
            if result_column in result_df.columns:
                # Store the results
                self.formula_result = result_df[result_column]
                
                # Count true/false values
                try:
                    true_count = sum(result_df[result_column] == True)
                    false_count = sum(result_df[result_column] == False)
                    
                    summary = (
                        f"Formula tested successfully on {len(result_df)} records:\n"
                        f"- {true_count} records conform ({true_count/len(result_df)*100:.1f}%)\n"
                        f"- {false_count} records do not conform ({false_count/len(result_df)*100:.1f}%)"
                    )
                    
                    if warnings:
                        summary += f"\n\nWarnings ({len(warnings)}):\n"
                        for warning in warnings[:3]:  # Show first 3 warnings
                            summary += f"- {warning}\n"
                        if len(warnings) > 3:
                            summary += f"- And {len(warnings) - 3} more..."
                    
                    self._finish_test(True, summary)
                except Exception as e:
                    logger.error(f"Error processing results: {e}")
                    self._finish_test(False, f"Error processing results: {str(e)}")
            else:
                self._finish_test(False, f"Result column '{result_column}' not found in output")
        
        except Exception as e:
            logger.error(f"Error testing formula: {e}")
            self._finish_test(False, f"Error testing formula: {str(e)}")

    def _finish_test(self, success: bool, message: str):
        """
        Finish formula test and update UI
        
        Args:
            success: Whether the test was successful
            message: Message to display in results
        """
        # Stop progress bar and re-enable test button
        self.after(0, lambda: self.progress_bar.stop())
        self.after(0, lambda: self.test_btn.config(state=tk.NORMAL))
        
        # Show results
        if success:
            self._show_test_results(message)
        else:
            messagebox.showerror("Test Failed", message)

    def _show_test_results(self, summary: str):
        """
        Show test results in dialog
        
        Args:
            summary: Summary text to display
        """
        # Create and display results
        if hasattr(self, 'results_frame'):
            self.results_frame.destroy()
        
        self.results_frame = ttk.LabelFrame(self, text="Test Results")
        self.results_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Results text
        results_text = tk.Text(
            self.results_frame,
            wrap=tk.WORD,
            width=60,
            height=8,
            background='#F9F9F9',
            relief=tk.FLAT,
            padx=10,
            pady=10
        )
        results_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Insert result summary
        results_text.insert(tk.END, summary)
        
        # Make read-only
        results_text.config(state=tk.DISABLED)

    def _generate_sample_data(self):
        """Generate sample data for testing"""
        try:
            # Get parameters
            try:
                record_count = int(self.record_count_var.get())
                error_pct = float(self.error_pct_var.get()) / 100
            except ValueError:
                messagebox.showinfo(
                    "Invalid Input",
                    "Please enter valid numbers for record count and error percentage"
                )
                self.sample_data = None
                return
            
            # Basic validation
            if record_count <= 0 or record_count > 10000:
                messagebox.showinfo(
                    "Invalid Input", 
                    "Record count must be between 1 and 10,000"
                )
                self.sample_data = None
                return
            
            if error_pct < 0 or error_pct > 1:
                messagebox.showinfo(
                    "Invalid Input",
                    "Error percentage must be between 0 and 100"
                )
                self.sample_data = None
                return
            
            # Create sample data structure based on fields in the formula
            data = {}
            
            # Add standard fields for common test scenarios
            data["ID"] = [f"ID-{i:06d}" for i in range(1, record_count + 1)]
            
            # Add fields used in the formula
            for field in self.fields_used:
                if field not in data:
                    # Generate data based on field name
                    if "date" in field.lower():
                        # Date field
                        import datetime
                        data[field] = [
                            datetime.datetime.now() - datetime.timedelta(days=i % 30)
                            for i in range(1, record_count + 1)
                        ]
                    elif any(term in field.lower() for term in ["amount", "value", "price", "cost"]):
                        # Numeric field
                        data[field] = [
                            round(100 * i / record_count, 2) 
                            for i in range(1, record_count + 1)
                        ]
                    elif any(term in field.lower() for term in ["flag", "indicator", "valid", "enabled"]):
                        # Boolean field
                        data[field] = [
                            i < (record_count * (1 - error_pct))
                            for i in range(1, record_count + 1)
                        ]
                        # Shuffle to randomize
                        np.random.shuffle(data[field])
                    else:
                        # Text field
                        # For fields like Owner, Approver, etc. make it a person name
                        if any(term in field.lower() for term in ["name", "owner", "approver", "person", "user"]):
                            people = ["John Smith", "Emma Johnson", "Olivia Garcia", 
                                     "James Anderson", "Michael Brown", "Sarah Davis", 
                                     "William Thomas", "Patricia Moore"]
                            data[field] = np.random.choice(people, record_count)
                        else:
                            # Generic text field
                            data[field] = [f"{field}-{i}" for i in range(1, record_count + 1)]
            
            # Convert to DataFrame
            self.sample_data = pd.DataFrame(data)
            logger.info(f"Generated sample data with {record_count} rows")
        
        except Exception as e:
            logger.error(f"Error generating sample data: {e}")
            messagebox.showerror("Error", f"Failed to generate sample data: {str(e)}")
            self.sample_data = None

    def _load_data_file(self):
        """Load data from file for testing"""
        file_path = self.file_var.get()
        
        if not file_path or not os.path.exists(file_path):
            messagebox.showinfo("File Selection", "Please select a valid file")
            self.sample_data = None
            return
        
        try:
            # Determine file type
            if file_path.lower().endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path)
            elif file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                messagebox.showinfo("File Error", "Unsupported file type")
                self.sample_data = None
                return
            
            # Check if file contains data
            if df.empty:
                messagebox.showinfo("File Error", "The file contains no data")
                self.sample_data = None
                return
            
            # Check if all fields from formula exist in the data
            missing_fields = [field for field in self.fields_used if field not in df.columns]
            if missing_fields:
                if messagebox.askyesno(
                    "Missing Fields", 
                    f"The following fields used in the formula are missing from the data:\n"
                    f"{', '.join(missing_fields)}\n\n"
                    f"Would you like to add these fields with sample data?"
                ):
                    # Add missing fields with sample data
                    for field in missing_fields:
                        # Use different data types based on field name
                        if "date" in field.lower():
                            import datetime
                            df[field] = pd.NaT
                        elif any(term in field.lower() for term in ["amount", "value", "price", "cost"]):
                            df[field] = np.nan
                        else:
                            df[field] = None
                else:
                    self.sample_data = None
                    return
            
            self.sample_data = df
            logger.info(f"Loaded data file with {len(df)} rows")
        
        except Exception as e:
            logger.error(f"Error loading data file: {e}")
            messagebox.showerror("File Error", f"Failed to load file: {str(e)}")
            self.sample_data = None

    def get_formula(self) -> str:
        """
        Get the current formula
        
        Returns:
            str: Current formula
        """
        formula = self.formula_var.get()
        
        # Ensure formula starts with equals sign
        if formula and not formula.startswith('='):
            formula = f"={formula}"
            
        return formula
    
    def get_display_name(self) -> str:
        """
        Get the display name for the formula
        
        Returns:
            str: Display name
        """
        return self.display_name_var.get()
    
    def get_fields_used(self) -> set:
        """
        Get the fields used in the formula
        
        Returns:
            set: Set of field names used in the formula
        """
        return self.fields_used
    
    def is_valid(self) -> bool:
        """
        Check if the current formula is valid
        
        Returns:
            bool: True if formula is valid
        """
        return self.is_formula_valid
    
    def set_formula(self, formula: str) -> None:
        """
        Set the formula
        
        Args:
            formula: Excel formula to set
        """
        self.formula_var.set(formula)
    
    def set_display_name(self, name: str) -> None:
        """
        Set the display name
        
        Args:
            name: Display name for the formula
        """
        self.display_name_var.set(name)
    
    def cleanup(self) -> None:
        """Clean up resources, especially Excel processor"""
        if self.excel_processor:
            try:
                self.excel_processor.cleanup()
                self.excel_processor = None
            except Exception as e:
                logger.warning(f"Error cleaning up Excel processor: {e}")


# Example usage if run directly
if __name__ == "__main__":
    # Create a basic window to test the component
    root = tk.Tk()
    root.title("Formula Tester Component")
    root.geometry("800x600")
    
    def on_formula_changed(formula, display_name, is_valid, fields):
        print(f"Formula changed: {formula}")
        print(f"Display name: {display_name}")
        print(f"Valid: {is_valid}")
        print(f"Fields: {fields}")
    
    tester = FormulaTester(
        root,
        callback=on_formula_changed,
        initial_formula="=IF(Amount > 0, IF(Status='Active', TRUE, FALSE), FALSE)",
        description="Positive Amount Check"
    )
    tester.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
    root.protocol("WM_DELETE_WINDOW", lambda: (tester.cleanup(), root.destroy()))
    root.mainloop()
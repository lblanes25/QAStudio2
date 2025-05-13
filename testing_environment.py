import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import yaml
import random
import datetime
from typing import Dict, List, Any, Optional, Tuple
import logging

# Set up logging
logger = logging.getLogger("qa_analytics")


class TestingEnvironment:
    """
    Interactive testing environment for validating analytics configurations
    with sample data generation and visualization
    """

    def __init__(self, parent_frame, config_manager, data_processor_class=None, report_generator_class=None):
        """
        Initialize the testing environment
        
        Args:
            parent_frame: Parent tkinter frame
            config_manager: ConfigManager instance for loading configurations
            data_processor_class: EnhancedDataProcessor class (not instance)
            report_generator_class: EnhancedReportGenerator class (not instance)
        """
        self.parent = parent_frame
        self.config_manager = config_manager
        self.data_processor_class = data_processor_class
        self.report_generator_class = report_generator_class
        
        # State variables
        self.current_config = None
        self.current_sample_data = None
        self.test_results = None
        
        # Setup variables
        self.analytics_var = tk.StringVar()
        self.data_source_var = tk.StringVar(value="generate")
        self.record_count_var = tk.StringVar(value="100")
        self.error_pct_var = tk.StringVar(value="20")
        self.source_file_var = tk.StringVar()
        
        # Set up the UI
        self._setup_ui()
    
    def _setup_ui(self):
        """Set up the testing environment UI"""
        main_frame = ttk.Frame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Top section: Configuration selection and test setup
        top_frame = ttk.LabelFrame(main_frame, text="Test Configuration")
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Analytics selection
        selection_frame = ttk.Frame(top_frame)
        selection_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(selection_frame, text="Select Analytics:").grid(row=0, column=0, sticky=tk.W)
        
        # Get available analytics
        self.available_analytics = self.config_manager.get_available_analytics()
        analytics_values = [f"{id} - {name}" for id, name in self.available_analytics]
        
        self.analytics_combo = ttk.Combobox(selection_frame, textvariable=self.analytics_var, 
                                         values=analytics_values, state="readonly", width=50)
        self.analytics_combo.grid(row=0, column=1, sticky=tk.W, padx=(5, 0))
        self.analytics_combo.bind("<<ComboboxSelected>>", self._on_analytics_selected)
        
        # Load button
        load_btn = ttk.Button(selection_frame, text="Load Configuration", command=self._load_configuration)
        load_btn.grid(row=0, column=2, padx=10)
        
        # Data source options
        data_frame = ttk.LabelFrame(top_frame, text="Test Data")
        data_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        # Data source option radios
        option_frame = ttk.Frame(data_frame)
        option_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Radiobutton(option_frame, text="Generate Sample Data", variable=self.data_source_var, 
                      value="generate", command=self._update_data_options).pack(side=tk.LEFT)
        
        ttk.Radiobutton(option_frame, text="Use Existing File", variable=self.data_source_var,
                      value="existing", command=self._update_data_options).pack(side=tk.LEFT, padx=(20, 0))
        
        # Sample data options
        self.sample_frame = ttk.Frame(data_frame)
        self.sample_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(self.sample_frame, text="Number of Records:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.sample_frame, textvariable=self.record_count_var, width=10).grid(row=0, column=1, sticky=tk.W, padx=(5, 20))
        
        ttk.Label(self.sample_frame, text="Error Percentage:").grid(row=0, column=2, sticky=tk.W)
        ttk.Entry(self.sample_frame, textvariable=self.error_pct_var, width=10).grid(row=0, column=3, sticky=tk.W, padx=5)
        
        # Existing file options
        self.file_frame = ttk.Frame(data_frame)
        
        ttk.Label(self.file_frame, text="Data File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.file_frame, textvariable=self.source_file_var, width=50).grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Button(self.file_frame, text="Browse...", command=self._browse_file).grid(row=0, column=2, padx=5)
        
        # Update which frame is visible
        self._update_data_options()
        
        # Run Test button
        button_frame = ttk.Frame(top_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        self.run_btn = ttk.Button(button_frame, text="Run Test", command=self._run_test)
        self.run_btn.pack(side=tk.RIGHT)
        
        # Progress bar
        self.progress = ttk.Progressbar(button_frame, orient="horizontal", length=200, mode="indeterminate")
        self.progress.pack(side=tk.RIGHT, padx=(0, 10))
        
        # Results section
        results_frame = ttk.LabelFrame(main_frame, text="Test Results")
        results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create notebook for result tabs
        self.results_notebook = ttk.Notebook(results_frame)
        self.results_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create result tabs
        self.summary_tab = ttk.Frame(self.results_notebook)
        self.detail_tab = ttk.Frame(self.results_notebook)
        self.sample_tab = ttk.Frame(self.results_notebook)
        
        self.results_notebook.add(self.summary_tab, text="Summary")
        self.results_notebook.add(self.detail_tab, text="Detail")
        self.results_notebook.add(self.sample_tab, text="Sample Data")
        
        # Summary tab content
        self._setup_summary_tab()
        
        # Detail tab content
        self._setup_detail_tab()
        
        # Sample data tab content
        self._setup_sample_tab()
        
        # Add export buttons at the bottom
        export_frame = ttk.Frame(main_frame)
        export_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(export_frame, text="Export Sample Data", command=self._export_sample_data).pack(side=tk.LEFT)
        ttk.Button(export_frame, text="Export Results", command=self._export_results).pack(side=tk.LEFT, padx=10)
        ttk.Button(export_frame, text="Generate Report", command=self._generate_report).pack(side=tk.RIGHT)
    
    def _setup_summary_tab(self):
        """Set up the summary tab content"""
        # Create frame for summary stats
        stats_frame = ttk.LabelFrame(self.summary_tab, text="Validation Statistics")
        stats_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Grid for stats
        self.stats_labels = {}
        row = 0
        
        # Total records
        ttk.Label(stats_frame, text="Total Records:").grid(row=row, column=0, sticky=tk.W, padx=10, pady=5)
        self.stats_labels['total'] = ttk.Label(stats_frame, text="--")
        self.stats_labels['total'].grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        row += 1
        
        # Generally Conforms (GC)
        ttk.Label(stats_frame, text="Generally Conforms (GC):").grid(row=row, column=0, sticky=tk.W, padx=10, pady=5)
        self.stats_labels['gc'] = ttk.Label(stats_frame, text="--")
        self.stats_labels['gc'].grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        row += 1
        
        # Does Not Conform (DNC)
        ttk.Label(stats_frame, text="Does Not Conform (DNC):").grid(row=row, column=0, sticky=tk.W, padx=10, pady=5)
        self.stats_labels['dnc'] = ttk.Label(stats_frame, text="--")
        self.stats_labels['dnc'].grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        row += 1
        
        # Partially Conforms (PC)
        ttk.Label(stats_frame, text="Partially Conforms (PC):").grid(row=row, column=0, sticky=tk.W, padx=10, pady=5)
        self.stats_labels['pc'] = ttk.Label(stats_frame, text="--")
        self.stats_labels['pc'].grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        row += 1
        
        # Error Percentage
        ttk.Label(stats_frame, text="Error Percentage:").grid(row=row, column=0, sticky=tk.W, padx=10, pady=5)
        self.stats_labels['error_pct'] = ttk.Label(stats_frame, text="--")
        self.stats_labels['error_pct'].grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        row += 1
        
        # Threshold Status
        ttk.Label(stats_frame, text="Threshold Status:").grid(row=row, column=0, sticky=tk.W, padx=10, pady=5)
        self.stats_labels['threshold'] = ttk.Label(stats_frame, text="--")
        self.stats_labels['threshold'].grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        
        # Create frame for group summary results
        group_frame = ttk.LabelFrame(self.summary_tab, text="Group Summary")
        group_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Treeview for group results
        columns = ("Group", "GC", "PC", "DNC", "Total", "Error %", "Status")
        self.group_tree = ttk.Treeview(group_frame, columns=columns, show="headings", height=8)
        
        # Configure columns
        self.group_tree.column("Group", width=150)
        self.group_tree.column("GC", width=70, anchor=tk.CENTER)
        self.group_tree.column("PC", width=70, anchor=tk.CENTER)
        self.group_tree.column("DNC", width=70, anchor=tk.CENTER)
        self.group_tree.column("Total", width=70, anchor=tk.CENTER)
        self.group_tree.column("Error %", width=70, anchor=tk.CENTER)
        self.group_tree.column("Status", width=100, anchor=tk.CENTER)
        
        # Configure headings
        for col in columns:
            self.group_tree.heading(col, text=col)
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(group_frame, orient="vertical", command=self.group_tree.yview)
        self.group_tree.configure(yscrollcommand=tree_scroll.set)
        
        # Pack tree and scrollbar
        self.group_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def _setup_detail_tab(self):
        """Set up the detail tab content"""
        # Create container for details
        detail_frame = ttk.Frame(self.detail_tab)
        detail_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Filter options
        filter_frame = ttk.Frame(detail_frame)
        filter_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(filter_frame, text="Show:").pack(side=tk.LEFT)

        self.filter_var = tk.StringVar(value="all")
        ttk.Radiobutton(filter_frame, text="All", variable=self.filter_var,
                        value="all", command=self._apply_filter).pack(side=tk.LEFT, padx=(5, 10))
        ttk.Radiobutton(filter_frame, text="GC Only", variable=self.filter_var,
                        value="gc", command=self._apply_filter).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(filter_frame, text="DNC Only", variable=self.filter_var,
                        value="dnc", command=self._apply_filter).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(filter_frame, text="PC Only", variable=self.filter_var,
                        value="pc", command=self._apply_filter).pack(side=tk.LEFT)

        # Create a container frame for the treeview and scrollbars
        tree_container = ttk.Frame(detail_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        # Create treeview for detail results
        self.detail_tree = ttk.Treeview(tree_container, show="headings")

        # Add scrollbars
        detail_y_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self.detail_tree.yview)
        detail_x_scroll = ttk.Scrollbar(tree_container, orient="horizontal", command=self.detail_tree.xview)
        self.detail_tree.configure(yscrollcommand=detail_y_scroll.set, xscrollcommand=detail_x_scroll.set)

        # Pack tree and scrollbars
        self.detail_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        detail_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        detail_x_scroll.pack(side=tk.BOTTOM, fill=tk.X)

    def _setup_sample_tab(self):
        """Set up the sample data tab content"""
        # Create container for sample data
        sample_frame = ttk.Frame(self.sample_tab)
        sample_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create a container frame for the treeview and scrollbars
        tree_container = ttk.Frame(sample_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        # Create treeview for sample data
        self.sample_tree = ttk.Treeview(tree_container, show="headings")

        # Add scrollbars
        sample_y_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self.sample_tree.yview)
        sample_x_scroll = ttk.Scrollbar(tree_container, orient="horizontal", command=self.sample_tree.xview)
        self.sample_tree.configure(yscrollcommand=sample_y_scroll.set, xscrollcommand=sample_x_scroll.set)

        # Pack tree and scrollbars
        self.sample_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sample_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        sample_x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
    
    def _update_data_options(self):
        """Update the data options based on the selected data source"""
        data_source = self.data_source_var.get()
        
        if data_source == "generate":
            self.sample_frame.pack(fill=tk.X, padx=10, pady=5)
            self.file_frame.pack_forget()
        else:
            self.sample_frame.pack_forget()
            self.file_frame.pack(fill=tk.X, padx=10, pady=5)
    
    def _browse_file(self):
        """Browse for a data file"""
        filename = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")],
            title="Select Data File"
        )
        
        if filename:
            self.source_file_var.set(filename)
    
    def _on_analytics_selected(self, event):
        """Handle analytics selection event"""
        self._load_configuration()
    
    def _load_configuration(self):
        """Load the selected analytics configuration"""
        selection = self.analytics_var.get()
        if not selection:
            return
        
        try:
            # Extract analytics ID from the selection
            analytics_id = selection.split(" - ")[0]
            
            # Load configuration
            config = self.config_manager.get_config(analytics_id)
            if not config:
                messagebox.showerror("Error", f"Failed to load configuration for QA-{analytics_id}")
                return
            
            self.current_config = config
            messagebox.showinfo("Success", f"Loaded configuration for QA-{analytics_id}: {config.get('analytic_name', '')}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading configuration: {e}")
    
    def _generate_sample_data(self):
        """Generate sample data based on the current configuration"""
        if not self.current_config:
            messagebox.showerror("Error", "No configuration loaded")
            return None
        
        try:
            # Get required parameters
            record_count = int(self.record_count_var.get())
            error_pct = float(self.error_pct_var.get()) / 100
            
            # Extract required columns from config
            required_columns = []
            
            # Check different config formats
            if 'source' in self.current_config and 'required_columns' in self.current_config['source']:
                # Old format with source.required_columns
                for col in self.current_config['source']['required_columns']:
                    if isinstance(col, dict):
                        required_columns.append(col['name'])
                    else:
                        required_columns.append(col)
            elif 'data_source' in self.current_config and 'required_fields' in self.current_config['data_source']:
                # New format with data_source.required_fields
                required_columns = self.current_config['data_source']['required_fields']
            
            if not required_columns:
                messagebox.showerror("Error", "No required columns found in configuration")
                return None
            
            # Create a DataFrame with the required columns
            df = pd.DataFrame(columns=required_columns)
            
            # Create data generators for different types of columns
            generators = {}
            
            for col in required_columns:
                # Determine the type of column based on the name
                if 'ID' in col or 'id' in col:
                    generators[col] = lambda i: f"ID-{i:06d}"
                elif 'Date' in col or 'date' in col:
                    generators[col] = self._generate_date
                elif 'submitter' in col.lower() or 'preparer' in col.lower():
                    generators[col] = self._generate_submitter
                elif 'approver' in col.lower() or 'reviewer' in col.lower():
                    generators[col] = self._generate_approver
                elif 'risk' in col.lower() and 'rating' in col.lower():
                    generators[col] = self._generate_risk_rating
                elif 'third' in col.lower() and 'party' in col.lower():
                    generators[col] = self._generate_third_party
                else:
                    generators[col] = self._generate_text
            
            # Now generate the data
            # First, generate valid data (will conform to validations)
            valid_count = int(record_count * (1 - error_pct))
            valid_data = []
            
            for i in range(valid_count):
                record = {}
                for col, generator in generators.items():
                    record[col] = generator(i)
                
                # Ensure approvals are in sequence for valid records
                if 'Submit Date' in record and 'TL Approval Date' in record and 'AL Approval Date' in record:
                    base_date = datetime.datetime.now() - datetime.timedelta(days=random.randint(30, 60))
                    record['Submit Date'] = base_date
                    record['TL Approval Date'] = base_date + datetime.timedelta(days=random.randint(1, 3))
                    record['AL Approval Date'] = record['TL Approval Date'] + datetime.timedelta(days=random.randint(1, 3))
                
                # Ensure submitter is not approver for valid records
                if 'TW submitter' in record and 'TL approver' in record and 'AL approver' in record:
                    # Make sure they're different
                    while record['TW submitter'] == record['TL approver'] or record['TW submitter'] == record['AL approver']:
                        record['TW submitter'] = self._generate_submitter(i)
                
                valid_data.append(record)
            
            # Now generate invalid data (will not conform to validations)
            invalid_count = record_count - valid_count
            invalid_data = []
            
            for i in range(invalid_count):
                record = {}
                for col, generator in generators.items():
                    record[col] = generator(i + valid_count)
                
                # Introduce errors based on validation rules in config
                self._introduce_errors(record)
                
                invalid_data.append(record)
            
            # Combine valid and invalid data
            all_data = valid_data + invalid_data
            
            # Shuffle to mix valid and invalid
            random.shuffle(all_data)
            
            # Create DataFrame
            sample_df = pd.DataFrame(all_data)
            
            # Return the generated data
            return sample_df
            
        except Exception as e:
            messagebox.showerror("Error", f"Error generating sample data: {e}")
            logger.error(f"Error generating sample data: {e}")
            return None
    
    def _introduce_errors(self, record):
        """Introduce errors based on validation rules in the configuration"""
        if not self.current_config or 'validations' not in self.current_config:
            return
        
        # Get validation rules
        validations = self.current_config['validations']
        
        # Randomly select which validation to violate
        if not validations:
            return
        
        violation = random.choice(validations)
        rule = violation.get('rule')
        
        if rule == 'segregation_of_duties':
            # Make submitter same as one of the approvers
            params = violation.get('parameters', {})
            submitter_field = params.get('submitter_field')
            approver_fields = params.get('approver_fields', [])
            
            if submitter_field and approver_fields and submitter_field in record:
                # Pick a random approver field that exists in the record
                existing_approvers = [f for f in approver_fields if f in record]
                if existing_approvers:
                    approver_field = random.choice(existing_approvers)
                    # Make submitter same as approver
                    record[submitter_field] = record[approver_field]
        
        elif rule == 'approval_sequence':
            # Mess up the approval sequence
            params = violation.get('parameters', {})
            date_fields = params.get('date_fields_in_order', [])
            
            # Check if we have at least 2 date fields
            date_fields_in_record = [f for f in date_fields if f in record]
            if len(date_fields_in_record) >= 2:
                # Pick two random date fields
                field1, field2 = random.sample(date_fields_in_record, 2)
                # Swap their order if they should be in sequence
                idx1 = date_fields.index(field1)
                idx2 = date_fields.index(field2)
                
                if idx1 < idx2:  # field1 should be before field2
                    # Make field1 later than field2
                    if isinstance(record[field1], datetime.datetime) and isinstance(record[field2], datetime.datetime):
                        # Add days to field1 to make it later
                        record[field1] = record[field2] + datetime.timedelta(days=random.randint(1, 5))
        
        elif rule == 'third_party_risk_validation':
            # Either add third parties but set risk to N/A, or remove third parties but keep risk rating
            params = violation.get('parameters', {})
            third_party_field = params.get('third_party_field')
            risk_level_field = params.get('risk_level_field')
            
            if third_party_field in record and risk_level_field in record:
                if random.choice([True, False]):
                    # Add third parties but set risk to N/A
                    record[third_party_field] = "Vendor A, Vendor B"
                    record[risk_level_field] = "N/A"
                else:
                    # Remove third parties but keep risk rating
                    record[third_party_field] = ""
                    record[risk_level_field] = random.choice(["Critical", "High", "Medium", "Low"])
    
    def _generate_submitter(self, index):
        """Generate a submitter name"""
        submitters = [
            "John Smith", "Emma Johnson", "Michael Brown", "Sarah Davis", "David Wilson",
            "Jennifer Miller", "Robert Taylor", "Jessica Anderson", "William Thomas", "Lisa Jackson"
        ]
        return random.choice(submitters)
    
    def _generate_approver(self, index):
        """Generate an approver name"""
        approvers = [
            "Alex Rodriguez", "Michelle Lee", "Richard White", "Patricia Moore", "James Martin",
            "Elizabeth Thompson", "Charles Garcia", "Susan Clark", "Joseph Lewis", "Donna Hall"
        ]
        return random.choice(approvers)
    
    def _generate_date(self, index):
        """Generate a random date in the past 90 days"""
        days_ago = random.randint(0, 90)
        return datetime.datetime.now() - datetime.timedelta(days=days_ago)
    
    def _generate_risk_rating(self, index):
        """Generate a risk rating"""
        ratings = ["Critical", "High", "Medium", "Low", "N/A"]
        return random.choice(ratings)
    
    def _generate_third_party(self, index):
        """Generate third party information"""
        third_parties = [
            "", "", "Vendor A", "Vendor B", "Vendor C", 
            "Vendor A, Vendor B", "Vendor C, Vendor D", "Vendor E"
        ]
        return random.choice(third_parties)
    
    def _generate_text(self, index):
        """Generate generic text"""
        words = ["Alpha", "Beta", "Gamma", "Delta", "Epsilon", "Zeta", "Eta", "Theta", "Iota", "Kappa"]
        return f"{random.choice(words)} {random.choice(words)}"
    
    def _run_test(self):
        """Run the test with the current configuration"""
        if not self.current_config:
            messagebox.showerror("Error", "No configuration loaded")
            return
        
        # Check that we have the necessary processor class
        if not self.data_processor_class:
            messagebox.showerror("Error", "Data processor class not available")
            return
        
        try:
            # Start progress
            self.progress.start()
            self.run_btn.config(state=tk.DISABLED)
            
            # Get data source
            data_source = self.data_source_var.get()
            
            if data_source == "generate":
                # Generate sample data
                sample_data = self._generate_sample_data()
                if sample_data is None:
                    return
                
                # Save to temporary file
                temp_dir = "temp"
                if not os.path.exists(temp_dir):
                    os.makedirs(temp_dir)
                
                temp_file = os.path.join(temp_dir, f"sample_data_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                sample_data.to_excel(temp_file, index=False)
                
                self.current_sample_data = sample_data
                source_file = temp_file
                
            else:  # Use existing file
                source_file = self.source_file_var.get()
                if not source_file or not os.path.exists(source_file):
                    messagebox.showerror("Error", "Please select a valid data file")
                    self.progress.stop()
                    self.run_btn.config(state=tk.NORMAL)
                    return
                
                # Load the file to display in sample tab
                try:
                    self.current_sample_data = pd.read_excel(source_file)
                except Exception as e:
                    logger.error(f"Error loading sample data: {e}")
                    self.current_sample_data = None
            
            # Create processor instance
            processor = self.data_processor_class(self.current_config)
            
            # Process data
            success, message = processor.process_data(source_file)
            
            if not success:
                messagebox.showerror("Error", f"Processing failed: {message}")
                self.progress.stop()
                self.run_btn.config(state=tk.NORMAL)
                return
            
            # Store results
            self.test_results = processor.results
            
            # Update UI with results
            self._update_results_display()
            
            # Display sample data
            self._display_sample_data()
            
            messagebox.showinfo("Success", "Test completed successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error running test: {e}")
            logger.error(f"Error running test: {e}", exc_info=True)
        
        finally:
            # Stop progress
            self.progress.stop()
            self.run_btn.config(state=tk.NORMAL)
    
    def _update_results_display(self):
        """Update the UI with test results"""
        if not self.test_results:
            return
        
        # Update summary statistics
        detail_data = self.test_results.get('detail')
        if detail_data is not None and 'Compliance' in detail_data:
            total = len(detail_data)
            gc_count = sum(detail_data['Compliance'] == 'GC')
            dnc_count = sum(detail_data['Compliance'] == 'DNC')
            pc_count = sum(detail_data['Compliance'] == 'PC')
            
            error_pct = (dnc_count / total * 100) if total > 0 else 0
            
            # Update labels
            self.stats_labels['total'].config(text=str(total))
            self.stats_labels['gc'].config(text=f"{gc_count} ({gc_count/total*100:.1f}%)")
            self.stats_labels['dnc'].config(text=f"{dnc_count} ({dnc_count/total*100:.1f}%)")
            self.stats_labels['pc'].config(text=f"{pc_count} ({pc_count/total*100:.1f}%)")
            self.stats_labels['error_pct'].config(text=f"{error_pct:.2f}%")
            
            # Check threshold
            threshold = self.current_config.get('thresholds', {}).get('error_percentage', 5.0)
            if error_pct > threshold:
                self.stats_labels['threshold'].config(text=f"EXCEEDS {threshold}%", foreground="red")
            else:
                self.stats_labels['threshold'].config(text=f"Within {threshold}%", foreground="green")
        
        # Update group summary
        summary_data = self.test_results.get('summary')
        if summary_data is not None:
            # Clear existing items
            for item in self.group_tree.get_children():
                self.group_tree.delete(item)
            
            # Add new data
            threshold = self.current_config.get('thresholds', {}).get('error_percentage', 5.0)
            
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
                item_id = self.group_tree.insert("", tk.END, values=(
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
                    self.group_tree.item(item_id, tags=("exceeds",))
            
            # Configure tag colors
            self.group_tree.tag_configure("exceeds", background="#FFE6E6")  # Light red
        
        # Update detail view
        self._update_detail_view()
    
    def _update_detail_view(self):
        """Update the detail view with filtered results"""
        if not self.test_results:
            return
        
        detail_data = self.test_results.get('detail')
        if detail_data is None:
            return
        
        # Apply filter
        filter_value = self.filter_var.get()
        if filter_value == "all":
            filtered_data = detail_data
        elif filter_value == "gc":
            filtered_data = detail_data[detail_data['Compliance'] == 'GC']
        elif filter_value == "dnc":
            filtered_data = detail_data[detail_data['Compliance'] == 'DNC']
        elif filter_value == "pc":
            filtered_data = detail_data[detail_data['Compliance'] == 'PC']
        else:
            filtered_data = detail_data
        
        # Clear existing items
        for col in self.detail_tree['columns']:
            self.detail_tree.heading(col, text="")
        
        self.detail_tree['columns'] = []
        
        for item in self.detail_tree.get_children():
            self.detail_tree.delete(item)
        
        if filtered_data.empty:
            return
        
        # Set up columns
        columns = list(filtered_data.columns)
        self.detail_tree['columns'] = columns
        
        # Configure columns and headings
        for col in columns:
            self.detail_tree.column(col, width=100, stretch=True)
            self.detail_tree.heading(col, text=col)
        
        # Add data rows
        for idx, row in filtered_data.iterrows():
            values = []
            for col in columns:
                val = row[col]
                # Format date values
                if isinstance(val, pd.Timestamp) or isinstance(val, datetime.datetime):
                    val = val.strftime('%Y-%m-%d %H:%M')
                values.append(val)
            
            item_id = self.detail_tree.insert("", tk.END, values=values)
            
            # Color code based on compliance
            compliance = row.get('Compliance')
            if compliance == 'GC':
                self.detail_tree.item(item_id, tags=("gc",))
            elif compliance == 'DNC':
                self.detail_tree.item(item_id, tags=("dnc",))
            elif compliance == 'PC':
                self.detail_tree.item(item_id, tags=("pc",))
        
        # Configure tag colors
        self.detail_tree.tag_configure("gc", background="#E6FFE6")  # Light green
        self.detail_tree.tag_configure("dnc", background="#FFE6E6")  # Light red
        self.detail_tree.tag_configure("pc", background="#FFF8E6")  # Light yellow
    
    def _display_sample_data(self):
        """Display the sample data in the sample tab"""
        if self.current_sample_data is None:
            return
        
        # Clear existing items
        for col in self.sample_tree['columns']:
            self.sample_tree.heading(col, text="")
        
        self.sample_tree['columns'] = []
        
        for item in self.sample_tree.get_children():
            self.sample_tree.delete(item)
        
        # Set up columns
        columns = list(self.current_sample_data.columns)
        self.sample_tree['columns'] = columns
        
        # Configure columns and headings
        for col in columns:
            self.sample_tree.column(col, width=100, stretch=True)
            self.sample_tree.heading(col, text=col)
        
        # Add data rows
        for idx, row in self.current_sample_data.iterrows():
            values = []
            for col in columns:
                val = row[col]
                # Format date values
                if isinstance(val, pd.Timestamp) or isinstance(val, datetime.datetime):
                    val = val.strftime('%Y-%m-%d %H:%M')
                values.append(val)
            
            self.sample_tree.insert("", tk.END, values=values)
    
    def _apply_filter(self):
        """Apply filter to detail view"""
        self._update_detail_view()
    
    def _export_sample_data(self):
        """Export the sample data to an Excel file"""
        if self.current_sample_data is None:
            messagebox.showinfo("Info", "No sample data available")
            return
        
        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Sample Data"
        )
        
        if not file_path:
            return
        
        try:
            # Save to Excel
            self.current_sample_data.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Sample data saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save sample data: {e}")
    
    def _export_results(self):
        """Export the test results to an Excel file"""
        if self.test_results is None:
            messagebox.showinfo("Info", "No test results available")
            return
        
        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Test Results"
        )
        
        if not file_path:
            return
        
        try:
            # Create a workbook
            with pd.ExcelWriter(file_path) as writer:
                # Save summary
                if 'summary' in self.test_results and self.test_results['summary'] is not None:
                    self.test_results['summary'].to_excel(writer, sheet_name='Summary', index=False)
                
                # Save detail
                if 'detail' in self.test_results and self.test_results['detail'] is not None:
                    self.test_results['detail'].to_excel(writer, sheet_name='Detail', index=False)
            
            messagebox.showinfo("Success", f"Test results saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save test results: {e}")
    
    def _generate_report(self):
        """Generate a formatted report using the ReportGenerator"""
        if self.test_results is None:
            messagebox.showinfo("Info", "No test results available")
            return
        
        if not self.report_generator_class:
            messagebox.showerror("Error", "Report generator class not available")
            return
        
        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Report"
        )
        
        if not file_path:
            return
        
        try:
            # Create report generator
            report_generator = self.report_generator_class(self.current_config, self.test_results)
            
            # Generate report
            report_path = report_generator.generate_main_report(file_path)
            
            if report_path:
                messagebox.showinfo("Success", f"Report generated: {report_path}")
            else:
                messagebox.showerror("Error", "Failed to generate report")
        except Exception as e:
            messagebox.showerror("Error", f"Error generating report: {e}")
            logger.error(f"Error generating report: {e}", exc_info=True)

import tkinter as tk
from tkinter import ttk, font
import os
from PIL import Image, ImageTk  # You'll need to install Pillow: pip install Pillow

class ModernQAUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QA Analytics Automation")
        self.geometry("1100x700")
        self.configure(bg="white")
        
        # Set icon
        # self.iconphoto(True, tk.PhotoImage(file="logo.png"))  # Uncomment to use an actual logo
        
        # Create style
        self.style = ttk.Style(self)
        self._configure_styles()
        
        # Create main container with padding
        self.main_frame = ttk.Frame(self, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook for tabs
        self._create_notebook()
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(self, textvariable=self.status_var, 
                                    relief=tk.SUNKEN, anchor=tk.W, padding=(10, 2))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _configure_styles(self):
        """Configure styling for widgets to create a modern, clean look"""
        # Determine best font
        available_fonts = font.families()
        preferred_fonts = ['Inter', 'Helvetica Neue', 'Segoe UI', 'SF UI Text', 'Arial']
        
        # Find the first available preferred font
        ui_font = next((f for f in preferred_fonts if f in available_fonts), None)
        if not ui_font:
            ui_font = "TkDefaultFont"
        
        # Configure font sizes
        header_font = (ui_font, 12, 'bold')
        normal_font = (ui_font, 10)
        small_font = (ui_font, 9)
        
        # Configure ttk theme - start with a clean base theme
        self.style.theme_use('clam')
        
        # Configure colors
        bg_color = '#FFFFFF'  # White background
        accent_color = '#000000'  # Black for primary elements
        disabled_bg = '#F5F5F5'  # Light gray for disabled elements
        hover_color = '#EEEEEE'  # Lighter gray for hover states
        selected_bg = '#E0E0E0'  # Medium gray for selected items
        
        # Fresh/stale/not loaded colors for reference data tab
        self.fresh_color = '#e6ffe6'     # Light green
        self.stale_color = '#fff0e6'     # Light orange
        self.not_loaded_color = '#f0f0f0'  # Light gray
        
        # Configure widget styles
        
        # TFrame - regular frames
        self.style.configure('TFrame', background=bg_color)
        
        # TLabel - text labels
        self.style.configure('TLabel', background=bg_color, font=normal_font)
        self.style.configure('Header.TLabel', font=header_font)
        self.style.configure('Small.TLabel', font=small_font)
        
        # TButton - buttons with rounded corners (as much as ttk allows)
        self.style.configure('TButton', 
                             font=normal_font,
                             padding=(10, 5))
        
        # Primary button (for main actions)
        self.style.configure('Primary.TButton',
                             background=accent_color,
                             foreground='white',
                             padding=(15, 8))
        
        # TEntry - text entry fields
        self.style.configure('TEntry', 
                             font=normal_font,
                             padding=5,
                             fieldbackground=bg_color)
        
        # TCombobox - dropdown fields
        self.style.configure('TCombobox',
                             font=normal_font,
                             padding=5,
                             fieldbackground=bg_color)
        
        # TNotebook - tabbed interface
        self.style.configure('TNotebook',
                             background=bg_color)
        
        self.style.configure('TNotebook.Tab',
                             font=normal_font,
                             padding=(15, 5))
        
        # TProgressbar - progress indicators
        self.style.configure('TProgressbar',
                             background=accent_color,
                             troughcolor='#EEEEEE')
        
        # TTreeview - table views
        self.style.configure('Treeview',
                             font=normal_font,
                             background=bg_color,
                             fieldbackground=bg_color)
        
        self.style.configure('Treeview.Heading',
                             font=(ui_font, 10, 'bold'),
                             background='#F0F0F0')
        
        # LabelFrame
        self.style.configure('TLabelframe', 
                             background=bg_color)
        self.style.configure('TLabelframe.Label', 
                             font=normal_font,
                             background=bg_color)

    def _create_notebook(self):
        """Create the tabbed interface"""
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Add tabs
        self._add_run_analytics_tab()
        self._add_config_wizard_tab()
        self._add_testing_tab()
        self._add_scheduler_tab()
        self._add_data_sources_tab()
        self._add_reference_data_tab()
    
    def _add_run_analytics_tab(self):
        """Create the Run Analytics tab"""
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Run Analytics")
        
        # Configure grid layout
        tab.columnconfigure(0, weight=0)  # Label column
        tab.columnconfigure(1, weight=1)  # Input column
        tab.rowconfigure(5, weight=1)     # Log area should expand
        
        # QA-ID Selection
        ttk.Label(tab, text="QA-ID", style="Header.TLabel").grid(
            row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Create sample analytics list for dropdown
        analytics_options = [
            "QA-123 - Data Quality Analysis",
            "QA-77 - Audit Test Workpaper Approvals",
            "QA-78 - Third Party Risk Assessment",
            "QA-99 - Audit Workpaper Review Validation"
        ]
        
        analytics_combo = ttk.Combobox(
            tab,
            values=analytics_options,
            state="readonly",
            width=50
        )
        analytics_combo.current(0)
        analytics_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Source Data File
        ttk.Label(tab, text="Source Data File", style="Header.TLabel").grid(
            row=1, column=0, sticky=tk.W, pady=(10, 10))
        
        file_frame = ttk.Frame(tab)
        file_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(10, 10))
        file_frame.columnconfigure(0, weight=1)
        
        source_entry = ttk.Entry(
            file_frame,
            width=40
        )
        source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        source_entry.insert(0, "C:\\Data\\source_data.xlsx")
        
        source_button = ttk.Button(
            file_frame,
            text="📂",
            width=3
        )
        source_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Output Directory
        ttk.Label(tab, text="Output Directory", style="Header.TLabel").grid(
            row=2, column=0, sticky=tk.W, pady=(10, 10))
        
        output_frame = ttk.Frame(tab)
        output_frame.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=(10, 10))
        output_frame.columnconfigure(0, weight=1)
        
        output_entry = ttk.Entry(
            output_frame,
            width=40
        )
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        output_entry.insert(0, "C:\\Output")
        
        output_button = ttk.Button(
            output_frame,
            text="📂",
            width=3
        )
        output_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Run Analysis Button
        run_button = ttk.Button(
            tab,
            text="▶ Run Analysis",
            style="Primary.TButton"
        )
        run_button.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                       pady=(20, 10))
        
        # Progress Bar
        progress = ttk.Progressbar(
            tab,
            orient=tk.HORIZONTAL,
            mode='determinate',
            length=100,
            value=25  # Set to 25% for demonstration
        )
        progress.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                     pady=(0, 10))
        
        # Status Log
        ttk.Label(tab, text="Status Log", style="Header.TLabel").grid(
            row=5, column=0, sticky=tk.NW, pady=(10, 0))
        
        # Create frame for status log with scrollbar
        log_frame = ttk.Frame(tab)
        log_frame.grid(row=5, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), 
                      pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        log_text = tk.Text(
            log_frame,
            height=10,
            wrap=tk.WORD,
            font=('Courier', 10),
            background='#F9F9F9',
            relief=tk.FLAT
        )
        log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        log_scrollbar = ttk.Scrollbar(
            log_frame,
            orient=tk.VERTICAL,
            command=log_text.yview
        )
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        log_text.configure(yscrollcommand=log_scrollbar.set)
        
        # Add sample log entries
        log_text.insert(tk.END, "Starting Enhanced QA Analytics in GUI mode\n")
        log_text.insert(tk.END, "Processing QA-ID QA-123 with source file\n")
        log_text.insert(tk.END, "C:\\Data\\source_data.xlsx\n")
        log_text.insert(tk.END, "Processing completed: 100 records processed\n")
        log_text.insert(tk.END, "Processing complete. Generated 1 report.\n")
        
        # Make log read-only
        log_text.config(state=tk.DISABLED)

    def _add_config_wizard_tab(self):
        """Create the Configuration Wizard tab"""
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Configuration Wizard")
        
        # Create a stepper header for the wizard
        stepper_frame = ttk.Frame(tab)
        stepper_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Create stepper circles
        steps = ["Data Source", "Basic Settings", "Validation Rules", "Review & Save"]
        step_frames = []
        
        for i, step_name in enumerate(steps):
            step_frame = ttk.Frame(stepper_frame)
            
            # Create the circle with Canvas
            canvas = tk.Canvas(step_frame, width=30, height=30, 
                              highlightthickness=0, bg='white')
            canvas.pack(side=tk.LEFT)
            
            # Draw circle
            if i == 0:  # Current step
                circle_color = '#CCE5FF'  # Light blue
                text_color = 'black'
                outline_color = '#0066CC'  # Medium blue
            elif i < 0:  # Completed steps
                circle_color = '#90EE90'  # Light green
                text_color = 'black'
                outline_color = '#228B22'  # Dark green
            else:  # Future steps
                circle_color = 'white'
                text_color = 'gray'
                outline_color = 'gray'
            
            canvas.create_oval(5, 5, 25, 25, fill=circle_color, outline=outline_color, width=2)
            canvas.create_text(15, 15, text=str(i+1), fill=text_color)
            
            # Add step label
            label = ttk.Label(step_frame, text=step_name)
            if i == 0:  # Current step
                label.configure(font=('Segoe UI', 9, 'bold'))
            else:
                label.configure(foreground='gray')
            label.pack(side=tk.LEFT, padx=(5, 0))
            
            step_frame.pack(side=tk.LEFT)
            step_frames.append(step_frame)
            
            # Add connector line between steps (except after the last step)
            if i < len(steps) - 1:
                connector = ttk.Separator(stepper_frame, orient="horizontal")
                connector.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10, pady=15)
        
        # Content for Data Source step
        content_frame = ttk.Frame(tab)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Data Source Name
        ttk.Label(content_frame, text="Data Source Name", style="Header.TLabel").pack(
            anchor=tk.W, pady=(0, 5))
        
        # Available data sources
        data_sources = [
            "audit_workpaper_approvals",
            "third_party_risk",
            "audit_planning_approvals",
            "risk_assessment_validation",
            "audit_workpapers_2025q2"
        ]
        
        # Dropdown for data source selection
        source_combo = ttk.Combobox(
            content_frame,
            values=data_sources,
            state="readonly",
            width=50
        )
        source_combo.current(0)
        source_combo.pack(fill=tk.X, pady=(0, 15))
        
        # File Type Selection
        ttk.Label(content_frame, text="File Type", style="Header.TLabel").pack(
            anchor=tk.W, pady=(10, 5))
        
        # Radio buttons for file type
        file_type_frame = ttk.Frame(content_frame)
        file_type_frame.pack(fill=tk.X, pady=(0, 15))
        
        file_type_var = tk.StringVar(value="XLSX")
        
        ttk.Radiobutton(
            file_type_frame,
            text="XLSX",
            variable=file_type_var,
            value="XLSX"
        ).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Radiobutton(
            file_type_frame,
            text="CSV",
            variable=file_type_var,
            value="CSV"
        ).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Radiobutton(
            file_type_frame,
            text="JSON",
            variable=file_type_var,
            value="JSON"
        ).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Radiobutton(
            file_type_frame,
            text="Other",
            variable=file_type_var,
            value="Other"
        ).pack(side=tk.LEFT)
        
        # Column Mapping
        ttk.Label(content_frame, text="Column Mapping", style="Header.TLabel").pack(
            anchor=tk.W, pady=(10, 5))
        
        column_entry = ttk.Entry(
            content_frame,
            width=50
        )
        column_entry.pack(fill=tk.X, pady=(0, 15))
        column_entry.insert(0, "Optional")
        
        # Validation Rules
        ttk.Label(content_frame, text="Validation Rules", style="Header.TLabel").pack(
            anchor=tk.W, pady=(10, 5))
        
        validation_entry = ttk.Entry(
            content_frame,
            width=50
        )
        validation_entry.pack(fill=tk.X)
        validation_entry.insert(0, "Optional")
        
        # Navigation buttons
        nav_frame = ttk.Frame(tab)
        nav_frame.pack(fill=tk.X, pady=(20, 0))
        
        back_btn = ttk.Button(
            nav_frame,
            text="Back",
            state=tk.DISABLED  # Disabled for first step
        )
        back_btn.pack(side=tk.LEFT)
        
        next_btn = ttk.Button(
            nav_frame,
            text="Next"
        )
        next_btn.pack(side=tk.RIGHT)

    def _add_testing_tab(self):
        """Create the Testing tab"""
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Testing")
        
        # Analytics selection
        ttk.Label(tab, text="Select Analytics:", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 5))
        
        analytics_options = [
            "QA-123 - Data Quality Analysis",
            "QA-77 - Audit Test Workpaper Approvals",
            "QA-78 - Third Party Risk Assessment",
            "QA-99 - Audit Workpaper Review Validation"
        ]
        
        analytics_combo = ttk.Combobox(
            tab,
            values=analytics_options,
            state="readonly",
            width=50
        )
        analytics_combo.current(0)
        analytics_combo.pack(fill=tk.X, pady=(0, 15))
        
        # Test data options
        ttk.Label(tab, text="Test Data:", style="Header.TLabel").pack(anchor=tk.W, pady=(10, 5))
        
        option_frame = ttk.Frame(tab)
        option_frame.pack(fill=tk.X, pady=(0, 10))
        
        data_source_var = tk.StringVar(value="generate")
        
        ttk.Radiobutton(
            option_frame,
            text="Generate Sample Data",
            variable=data_source_var,
            value="generate"
        ).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Radiobutton(
            option_frame,
            text="Use Existing Data",
            variable=data_source_var,
            value="existing"
        ).pack(side=tk.LEFT)
        
        # Sample data options
        sample_frame = ttk.Frame(tab)
        sample_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(sample_frame, text="Number of Records:").pack(side=tk.LEFT)
        
        record_entry = ttk.Entry(
            sample_frame,
            width=8
        )
        record_entry.pack(side=tk.LEFT, padx=(5, 20))
        record_entry.insert(0, "100")
        
        ttk.Label(sample_frame, text="Error Percentage:").pack(side=tk.LEFT)
        
        error_entry = ttk.Entry(
            sample_frame,
            width=8
        )
        error_entry.pack(side=tk.LEFT, padx=(5, 0))
        error_entry.insert(0, "20")
        
        # Test actions
        action_frame = ttk.Frame(tab)
        action_frame.pack(fill=tk.X, pady=(10, 15))
        
        progress = ttk.Progressbar(
            action_frame,
            orient=tk.HORIZONTAL,
            mode='indeterminate',
            length=200
        )
        progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        run_btn = ttk.Button(
            action_frame,
            text="Run Test",
            style="Primary.TButton"
        )
        run_btn.pack(side=tk.RIGHT)
        
        # Results notebook
        ttk.Label(tab, text="Test Results:", style="Header.TLabel").pack(anchor=tk.W, pady=(10, 5))
        
        results_notebook = ttk.Notebook(tab)
        results_notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Summary tab
        summary_tab = ttk.Frame(results_notebook, padding=10)
        results_notebook.add(summary_tab, text="Summary")
        
        ttk.Label(
            summary_tab,
            text="Test Results Summary",
            style="Header.TLabel"
        ).pack(anchor=tk.W, pady=(0, 10))
        
        ttk.Label(
            summary_tab,
            text="Analytics: QA-123 - Data Quality Analysis"
        ).pack(anchor=tk.W, pady=(0, 5))
        
        ttk.Label(
            summary_tab,
            text="Total Records: 100"
        ).pack(anchor=tk.W, pady=(0, 5))
        
        ttk.Label(
            summary_tab,
            text="Generally Conforms (GC): 75 (75.0%)"
        ).pack(anchor=tk.W, pady=(0, 5))
        
        ttk.Label(
            summary_tab,
            text="Does Not Conform (DNC): 25 (25.0%)"
        ).pack(anchor=tk.W, pady=(0, 5))
        
        # Detail tab
        detail_tab = ttk.Frame(results_notebook, padding=10)
        results_notebook.add(detail_tab, text="Detail")
        
        ttk.Label(
            detail_tab,
            text="Detailed Results",
            style="Header.TLabel"
        ).pack(anchor=tk.W, pady=(0, 10))
        
        # Sample data tab
        sample_tab = ttk.Frame(results_notebook, padding=10)
        results_notebook.add(sample_tab, text="Sample Data")
        
        # Export buttons
        export_frame = ttk.Frame(tab)
        export_frame.pack(fill=tk.X, pady=(0, 0))
        
        ttk.Button(
            export_frame,
            text="Export Sample Data"
        ).pack(side=tk.LEFT)
        
        ttk.Button(
            export_frame,
            text="Export Results"
        ).pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(
            export_frame,
            text="Generate Report"
        ).pack(side=tk.RIGHT)

    def _add_scheduler_tab(self):
        """Create the Scheduler tab"""
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Scheduler")
        
        # Schedule Settings section
        settings_frame = ttk.LabelFrame(tab, text="Schedule Settings")
        settings_frame.pack(fill=tk.X, pady=(0, 15))
        
        settings_frame.columnconfigure(1, weight=1)
        settings_frame.columnconfigure(3, weight=1)
        
        # Default Run Time
        ttk.Label(settings_frame, text="Default Run Time:").grid(
            row=0, column=0, sticky=tk.W, padx=10, pady=10)
        
        time_entry = ttk.Entry(
            settings_frame,
            width=10
        )
        time_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=10)
        time_entry.insert(0, "00:00")
        
        # Run Day
        ttk.Label(settings_frame, text="Day:").grid(
            row=0, column=2, sticky=tk.W, padx=10, pady=10)
        
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday", "Daily"]
        
        day_combo = ttk.Combobox(
            settings_frame,
            values=days,
            state="readonly",
            width=15
        )
        day_combo.current(0)
        day_combo.grid(row=0, column=3, sticky=tk.W, padx=5, pady=10)
        
        # Email Configuration section
        email_frame = ttk.LabelFrame(tab, text="Email Configuration")
        email_frame.pack(fill=tk.X, pady=(0, 15))
        
        email_frame.columnconfigure(1, weight=1)
        
        # From Email
        ttk.Label(email_frame, text="From Email:").grid(
            row=0, column=0, sticky=tk.W, padx=10, pady=10)
        
        email_entry = ttk.Entry(
            email_frame,
            width=30
        )
        email_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=10)
        email_entry.insert(0, "qa.analytics@example.com")
        
        # SMTP Server
        ttk.Label(email_frame, text="SMTP Server:").grid(
            row=1, column=0, sticky=tk.W, padx=10, pady=10)
        
        smtp_entry = ttk.Entry(
            email_frame,
            width=30
        )
        smtp_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=10)
        smtp_entry.insert(0, "smtp.example.com")
        
        # Use TLS checkbox
        ttk.Checkbutton(
            email_frame,
            text="Use TLS"
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=10, pady=10)
        
        # Control Buttons
        control_frame = ttk.Frame(tab)
        control_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Button(
            control_frame,
            text="Start Scheduler",
            style="Primary.TButton"
        ).pack(side=tk.LEFT)
        
        ttk.Button(
            control_frame,
            text="Test Email"
        ).pack(side=tk.RIGHT)
        
        # Status section
        status_frame = ttk.LabelFrame(tab, text="Scheduler Status")
        status_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(
            status_frame,
            text="Scheduler is not running",
            foreground="red"
        ).pack(anchor=tk.W, padx=10, pady=10)
        
        # Scheduled tasks
        ttk.Label(tab, text="Scheduled Tasks:", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 5))
        
        # Create treeview for scheduled tasks
        columns = ("Task", "Schedule", "Last Run", "Next Run", "Status")
        tasks_tree = ttk.Treeview(
            tab,
            columns=columns,
            show="headings",
            height=5
        )
        
        # Configure columns
        tasks_tree.column("Task", width=150)
        tasks_tree.column("Schedule", width=150)
        tasks_tree.column("Last Run", width=150)
        tasks_tree.column("Next Run", width=150)
        tasks_tree.column("Status", width=100)
        
        # Configure headings
        for col in columns:
            tasks_tree.heading(col, text=col)
        
        # Add sample data
        tasks_tree.insert("", tk.END, values=(
            "QA-77 - Audit Workpaper Approvals",
            "Monday at 00:00",
            "N/A",
            "Next Monday at 00:00",
            "Pending"
        ))
        
        # Pack tree
        tasks_tree.pack(fill=tk.X, pady=(0, 10))

    def _add_data_sources_tab(self):
        """Create the Data Sources tab"""
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Data Sources")
        
        # Create treeview for data sources
        columns = ("Name", "Type", "Owner", "Version", "Last Updated", "Analytics")
        source_tree = ttk.Treeview(
            tab,
            columns=columns,
            show="headings",
            height=15
        )
        
        # Configure columns
        source_tree.column("Name", width=150)
        source_tree.column("Type", width=80)
        source_tree.column("Owner", width=150)
        source_tree.column("Version", width=80)
        source_tree.column("Last Updated", width=150)
        source_tree.column("Analytics", width=80, anchor=tk.CENTER)
        
        # Configure headings
        for col in columns:
            source_tree.heading(col, text=col)
        
        # Add sample data
        sample_sources = [
            ("audit_workpaper_approvals", "report", "Quality Assurance Team", "1.0", "2025-05-01", 1),
            ("third_party_risk", "report", "Risk Management", "1.1", "2025-04-15", 1),
            ("audit_planning_approvals", "report", "QA Team", "1.0", "2025-05-01", 3),
            ("risk_assessment_validation", "report", "Risk Management Team", "1.1", "2025-04-15", 3),
            ("audit_workpapers_2025q2", "report", "QA Analytics", "1.0", "2025-05-15", 1)
        ]
        
        for source in sample_sources:
            source_tree.insert("", tk.END, values=source)
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(tab, orient=tk.VERTICAL, command=source_tree.yview)
        source_tree.configure(yscrollcommand=tree_scroll.set)
        
        # Pack tree and scrollbar
        source_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Button frame
        button_frame = ttk.Frame(tab)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(
            button_frame,
            text="Refresh Registry"
        ).pack(side=tk.LEFT)
        
        ttk.Button(
            button_frame,
            text="View Details"
        ).pack(side=tk.RIGHT)

    def _add_reference_data_tab(self):
        """Create the Reference Data tab"""
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Reference Data")
        
        # Create treeview for reference data
        columns = ("Name", "Format", "Version", "Last Modified", "Rows", "Freshness")
        ref_tree = ttk.Treeview(
            tab,
            columns=columns,
            show="headings",
            height=15
        )
        
        # Configure columns
        ref_tree.column("Name", width=150)
        ref_tree.column("Format", width=80)
        ref_tree.column("Version", width=80)
        ref_tree.column("Last Modified", width=150)
        ref_tree.column("Rows", width=70, anchor=tk.CENTER)
        ref_tree.column("Freshness", width=100)
        
        # Configure headings
        for col in columns:
            ref_tree.heading(col, text=col)
        
        # Define tag colors
        ref_tree.tag_configure("fresh", background="#e6ffe6")  # Light green
        ref_tree.tag_configure("stale", background="#fff0e6")  # Light orange
        ref_tree.tag_configure("not_loaded", background="#f0f0f0")  # Light gray
        
        # Add sample data
        sample_data = [
            ("HR_Titles", "dictionary", "2025-Q2", "2025-04-15 10:30", 250, "✓ Fresh", "fresh"),
            ("Risk_Categories", "dataframe", "2025-01", "2025-01-20 14:22", 15, "⚠ Stale", "stale"),
            ("Control_Standards", "dataframe", "2025-05", "2025-05-01 09:15", 75, "✓ Fresh", "fresh"),
            ("Audit_Leaders", "dictionary", "1.0", "Not loaded", "-", "Not loaded", "not_loaded")
        ]
        
        for data in sample_data:
            ref_tree.insert("", tk.END, values=data[:-1], tags=(data[-1],))
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(tab, orient=tk.VERTICAL, command=ref_tree.yview)
        ref_tree.configure(yscrollcommand=tree_scroll.set)
        
        # Pack tree and scrollbar
        ref_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Button frame
        button_frame = ttk.Frame(tab)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(
            button_frame,
            text="Refresh Status"
        ).pack(side=tk.LEFT)
        
        ttk.Button(
            button_frame,
            text="Update Reference File"
        ).pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(
            button_frame,
            text="View Update History"
        ).pack(side=tk.RIGHT)


if __name__ == "__main__":
    app = ModernQAUI()
    app.mainloop()

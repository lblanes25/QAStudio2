# tabs/run_analytics_tab.py
import os
import tkinter as tk
from tkinter import ttk, filedialog
import threading
from typing import Callable


class RunAnalyticsTab(ttk.Frame):
    """
    Tab for running QA analytics with a clean, modern interface design.
    Allows users to select an analytics configuration, specify input/output
    files, and run the analysis.
    """

    def __init__(self, parent, status_callback: Callable):
        """
        Initialize the Run Analytics tab.

        Args:
            parent: Parent widget
            status_callback: Function to call to update status bar
        """
        super().__init__(parent, padding="20 15 20 15")
        self.parent = parent
        self.update_status = status_callback

        # Load available analytics configurations (would be loaded from actual config manager)
        self.available_analytics = [
            ('123', 'Data Quality Analysis'),
            ('77', 'Audit Test Workpaper Approvals'),
            ('78', 'Third Party Risk Assessment'),
            ('99', 'Audit Workpaper Review Validation')
        ]

        self._create_widgets()

    def _create_widgets(self):
        """Create all widgets for this tab"""
        # Use Grid layout for better control
        self.columnconfigure(0, weight=0)  # Label column
        self.columnconfigure(1, weight=1)  # Input field column
        self.rowconfigure(5, weight=1)  # Status log row should expand

        # QA-ID Selection Section
        qa_section = ttk.Frame(self)
        qa_section.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        qa_section.columnconfigure(1, weight=1)

        ttk.Label(qa_section, text="QA-ID", style="Header.TLabel").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 15))

        # Create formatted analytics list for dropdown
        analytics_options = [f"QA-{id} - {name}" for id, name in self.available_analytics]

        self.analytics_var = tk.StringVar()
        self.analytics_combo = ttk.Combobox(
            qa_section,
            textvariable=self.analytics_var,
            values=analytics_options,
            state="readonly",
            width=50
        )
        if analytics_options:
            self.analytics_combo.current(0)
        self.analytics_combo.grid(row=0, column=1, sticky=(tk.W, tk.E))

        # Source Data File Section
        file_section = ttk.Frame(self)
        file_section.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        file_section.columnconfigure(1, weight=1)

        ttk.Label(file_section, text="Source Data File", style="Header.TLabel").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 15))

        file_input_frame = ttk.Frame(file_section)
        file_input_frame.grid(row=0, column=1, sticky=(tk.W, tk.E))
        file_input_frame.columnconfigure(0, weight=1)

        self.source_file_var = tk.StringVar()
        source_entry = ttk.Entry(
            file_input_frame,
            textvariable=self.source_file_var,
            width=40
        )
        source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Add a container for the button to control its size and vertical alignment
        button_container = ttk.Frame(file_input_frame, width=40, height=36)
        button_container.pack(side=tk.LEFT, padx=(8, 0))
        button_container.pack_propagate(False)  # Prevent the button from changing the frame size

        source_button = ttk.Button(
            button_container,
            text="📂",
            style="Icon.TButton",
            command=self._browse_source_file
        )
        source_button.pack(fill=tk.BOTH, expand=True)

        # Output Directory Section
        output_section = ttk.Frame(self)
        output_section.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        output_section.columnconfigure(1, weight=1)

        ttk.Label(output_section, text="Output Directory", style="Header.TLabel").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 15))

        output_input_frame = ttk.Frame(output_section)
        output_input_frame.grid(row=0, column=1, sticky=(tk.W, tk.E))
        output_input_frame.columnconfigure(0, weight=1)

        self.output_dir_var = tk.StringVar()
        self.output_dir_var.set(os.path.join(os.getcwd(), "output"))

        output_entry = ttk.Entry(
            output_input_frame,
            textvariable=self.output_dir_var,
            width=40
        )
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Add a container for the button to control its size and vertical alignment
        button_container = ttk.Frame(output_input_frame, width=40, height=36)
        button_container.pack(side=tk.LEFT, padx=(8, 0))
        button_container.pack_propagate(False)

        output_button = ttk.Button(
            button_container,
            text="📂",
            style="Icon.TButton",
            command=self._browse_output_dir
        )
        output_button.pack(fill=tk.BOTH, expand=True)

        # Run Analysis Button Section
        action_section = ttk.Frame(self)
        action_section.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))

        run_button = ttk.Button(
            action_section,
            text="▶  Run Analysis",
            style="Primary.TButton",
            command=self._run_analysis
        )
        run_button.pack(side=tk.RIGHT)

        # Progress Bar Section
        progress_section = ttk.Frame(self)
        progress_section.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        progress_section.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(
            progress_section,
            orient=tk.HORIZONTAL,
            mode='determinate',
            length=100,
            style="TProgressbar"
        )
        self.progress.pack(fill=tk.X)

        # Status Log Section
        log_section = ttk.Frame(self)
        log_section.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_section.columnconfigure(0, weight=1)
        log_section.rowconfigure(1, weight=1)

        ttk.Label(log_section, text="Status Log", style="Header.TLabel").grid(
            row=0, column=0, sticky=tk.W, pady=(0, 10))

        # Create a card-like frame for the log with proper styling
        log_card = ttk.Frame(log_section, style="Card.TFrame")
        log_card.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(0, weight=1)

        # Create a container for the log text and scrollbar
        log_container = ttk.Frame(log_card, padding=2)
        log_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=1, pady=1)
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)

        # Create text widget with monospace font and subtle background
        self.log_text = tk.Text(
            log_container,
            height=10,
            wrap=tk.WORD,
            font=('Consolas', 10),
            background='#F9F9F9',
            relief=tk.FLAT,
            padx=15,
            pady=15,
            borderwidth=0
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Add a modern, thin scrollbar
        log_scrollbar = ttk.Scrollbar(
            log_container,
            orient=tk.VERTICAL,
            command=self.log_text.yview,
            style="Vertical.TScrollbar"
        )
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

        # Make log read-only and add placeholder text
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, "Ready to run analytics. Select options and click 'Run Analysis'...\n")
        self.log_text.config(state=tk.DISABLED)

    def _browse_source_file(self):
        """Open file dialog to select source data file"""
        filename = filedialog.askopenfilename(
            title="Select Source Data File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("CSV Files", "*.csv"),
                ("All Files", "*.*")
            ]
        )
        if filename:
            self.source_file_var.set(filename)
            self._log(f"Source file selected: {filename}")

    def _browse_output_dir(self):
        """Open directory dialog to select output directory"""
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directory:
            self.output_dir_var.set(directory)
            self._log(f"Output directory selected: {directory}")

    def reload_analytics(self):
        """Reload available analytics configurations"""
        try:
            # Get available configurations from config manager
            from qa_analytics.core.config_manager import ConfigManager
            config_manager = ConfigManager()

            # Get available analytics as (id, name) tuples
            available_analytics = config_manager.get_available_analytics()

            # Format for display in combobox
            self.analytics_options = [f"QA-{id} - {name}" for id, name in available_analytics]

            # Update combobox values
            if hasattr(self, 'analytics_combo') and self.analytics_combo.winfo_exists():
                self.analytics_combo['values'] = self.analytics_options

                # If we have options, select the first one
                if self.analytics_options:
                    self.analytics_combo.current(0)

            logger.info(f"Reloaded analytics configurations: {len(self.analytics_options)} configurations available")

        except Exception as e:
            logger.error(f"Error reloading analytics configurations: {e}")
            import traceback
            logger.error(traceback.format_exc())

    def _run_analysis(self):
        """Run the selected analytics configuration"""
        # Validate inputs
        if not self.analytics_var.get():
            self._log("Error: Please select a QA-ID")
            return

        if not self.source_file_var.get():
            self._log("Error: Please select a source data file")
            return

        if not os.path.exists(self.source_file_var.get()):
            self._log(f"Error: Source file does not exist: {self.source_file_var.get()}")
            return

        # Extract QA-ID from the selection
        qa_id = self.analytics_var.get().split(" - ")[0].replace("QA-", "")

        # Update status
        self.update_status(f"Running analysis: {self.analytics_var.get()}")

        # Reset progress
        self.progress["value"] = 0

        # Clear log and add initial message
        self._clear_log()
        self._log(f"Starting Enhanced QA Analytics in GUI mode")
        self._log(f"Processing QA-ID {qa_id} with source file")
        self._log(f"{self.source_file_var.get()}")

        # Run analysis in a separate thread
        threading.Thread(
            target=self._process_data,
            args=(qa_id,),
            daemon=True
        ).start()

    def _process_data(self, qa_id):
        """
        Process data in a separate thread

        Args:
            qa_id: Analytics ID to process
        """
        try:
            # Simulate processing
            # In a real implementation, this would call the actual data processor
            import time

            # Start progress animation
            for i in range(0, 101, 10):
                time.sleep(0.3)  # Simulate processing time
                self.progress["value"] = i
                if i == 50:
                    self._log("Processing completed: 100 records processed")

            # Update status on completion
            self._log("Processing complete. Generated 1 report.")
            self.update_status("Analysis complete")

        except Exception as e:
            self._log(f"Error during processing: {str(e)}")
            self.update_status("Analysis failed")

        finally:
            # Reset UI
            self.progress["value"] = 100

    def _log(self, message):
        """
        Add a message to the status log

        Args:
            message: Message to add
        """
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _clear_log(self):
        """Clear the status log"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
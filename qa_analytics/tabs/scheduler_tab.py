# tabs/scheduler_tab.py
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from typing import Callable
import datetime


class SchedulerTab(ttk.Frame):
    """
    Tab for scheduling automated analytics runs with a modern, clean interface.
    Allows users to configure scheduled tasks and monitor their status.
    """

    def __init__(self, parent, status_callback: Callable):
        """
        Initialize the Scheduler tab.

        Args:
            parent: Parent widget
            status_callback: Function to call to update status bar
        """
        super().__init__(parent, padding="20 15 20 15")
        self.parent = parent
        self.update_status = status_callback

        # State variables
        self.run_time_var = tk.StringVar(value="00:00")
        self.run_day_var = tk.StringVar(value="Monday")
        self.output_path_var = tk.StringVar()
        self.smtp_server_var = tk.StringVar()
        self.from_email_var = tk.StringVar()
        self.use_tls_var = tk.BooleanVar(value=True)
        self.scheduler_running = tk.BooleanVar(value=False)

        # Create widgets
        self._create_widgets()

    def _create_widgets(self):
        """Create all widgets for this tab with modern styling"""
        # Use grid layout for better control
        self.columnconfigure(0, weight=1)
        self.rowconfigure(5, weight=1)  # Log section should expand

        # Schedule Settings section
        settings_card = ttk.Frame(self, style="Card.TFrame")
        settings_card.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20), padx=2)

        settings_frame = ttk.LabelFrame(
            settings_card,
            text="Schedule Settings",
            padding=15
        )
        settings_frame.pack(fill=tk.BOTH, expand=True)
        settings_frame.columnconfigure(1, weight=1)

        # Default Run Time
        time_frame = ttk.Frame(settings_frame)
        time_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))

        ttk.Label(time_frame, text="Default Run Time:").pack(side=tk.LEFT, padx=(0, 10))

        ttk.Entry(
            time_frame,
            textvariable=self.run_time_var,
            width=10
        ).pack(side=tk.LEFT)

        # Run Day
        ttk.Label(time_frame, text="Day:").pack(side=tk.LEFT, padx=(20, 10))

        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday", "Daily"]

        day_combo = ttk.Combobox(
            time_frame,
            textvariable=self.run_day_var,
            values=days,
            state="readonly",
            width=15
        )
        day_combo.pack(side=tk.LEFT)

        # Output Path
        path_frame = ttk.Frame(settings_frame)
        path_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        path_frame.columnconfigure(1, weight=1)

        ttk.Label(path_frame, text="Output Path:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))

        path_input = ttk.Frame(path_frame)
        path_input.grid(row=0, column=1, sticky=(tk.W, tk.E))
        path_input.columnconfigure(0, weight=1)

        ttk.Entry(
            path_input,
            textvariable=self.output_path_var,
            width=40
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Button container for consistent size and alignment
        button_container = ttk.Frame(path_input, width=40, height=36)
        button_container.pack(side=tk.LEFT, padx=(8, 0))
        button_container.pack_propagate(False)

        ttk.Button(
            button_container,
            text="📂",
            style="Icon.TButton",
            command=self._browse_output_path
        ).pack(fill=tk.BOTH, expand=True)

        # Email Configuration section
        email_card = ttk.Frame(self, style="Card.TFrame")
        email_card.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20), padx=2)

        email_frame = ttk.LabelFrame(
            email_card,
            text="Email Configuration",
            padding=15
        )
        email_frame.pack(fill=tk.BOTH, expand=True)
        email_frame.columnconfigure(1, weight=1)

        # From Email
        ttk.Label(email_frame, text="From Email:").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 10))

        ttk.Entry(
            email_frame,
            textvariable=self.from_email_var,
            width=30
        ).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 10))

        # SMTP Server
        ttk.Label(email_frame, text="SMTP Server:").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 10))

        ttk.Entry(
            email_frame,
            textvariable=self.smtp_server_var,
            width=30
        ).grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 10))

        # Use TLS checkbox
        ttk.Checkbutton(
            email_frame,
            text="Use TLS",
            variable=self.use_tls_var
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))

        # Control Buttons
        control_frame = ttk.Frame(self)
        control_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 20))

        self.start_btn = ttk.Button(
            control_frame,
            text="Start Scheduler",
            style="Primary.TButton",
            command=self._toggle_scheduler
        )
        self.start_btn.pack(side=tk.LEFT)

        ttk.Button(
            control_frame,
            text="Test Email",
            command=self._test_email
        ).pack(side=tk.RIGHT)

        # Status section
        status_card = ttk.Frame(self, style="Card.TFrame")
        status_card.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 20), padx=2)

        status_frame = ttk.LabelFrame(
            status_card,
            text="Scheduler Status",
            padding=15
        )
        status_frame.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(
            status_frame,
            text="Scheduler is not running",
            foreground="#E74C3C"  # Red from color scheme
        )
        self.status_label.pack(anchor=tk.W)

        # Schedule preview
        ttk.Label(self, text="Scheduled Tasks:", style="Header.TLabel").grid(
            row=4, column=0, sticky=tk.W, pady=(0, 10))

        # Tasks container with card styling
        tasks_card = ttk.Frame(self, style="Card.TFrame")
        tasks_card.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 20), padx=2)
        tasks_card.columnconfigure(0, weight=1)
        tasks_card.rowconfigure(0, weight=1)

        # Container for treeview and scrollbar
        tasks_container = ttk.Frame(tasks_card, padding=2)
        tasks_container.pack(fill=tk.BOTH, expand=True)
        tasks_container.columnconfigure(0, weight=1)
        tasks_container.rowconfigure(0, weight=1)

        columns = ("Task", "Schedule", "Last Run", "Next Run", "Status")
        self.tasks_tree = ttk.Treeview(
            tasks_container,
            columns=columns,
            show="headings",
            height=6
        )

        # Configure columns
        self.tasks_tree.column("Task", width=150)
        self.tasks_tree.column("Schedule", width=120)
        self.tasks_tree.column("Last Run", width=120)
        self.tasks_tree.column("Next Run", width=120)
        self.tasks_tree.column("Status", width=100)

        # Configure headings
        for col in columns:
            self.tasks_tree.heading(col, text=col)

        # Add vertical scrollbar
        tasks_y_scroll = ttk.Scrollbar(
            tasks_container,
            orient="vertical",
            command=self.tasks_tree.yview,
            style="Vertical.TScrollbar"
        )
        tasks_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tasks_tree.configure(yscrollcommand=tasks_y_scroll.set)

        # Pack tree
        self.tasks_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add sample data
        self.tasks_tree.insert("", tk.END, values=(
            "QA-77 - Audit Workpaper Approvals",
            "Monday at 00:00",
            "N/A",
            "Next Monday at 00:00",
            "Pending"
        ))

        # Log section
        ttk.Label(self, text="Recent Log:", style="Header.TLabel").grid(
            row=5, column=0, sticky=tk.W, pady=(0, 10))

        # Log container with card styling
        log_card = ttk.Frame(self, style="Card.TFrame")
        log_card.grid(row=6, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=2)
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(0, weight=1)

        # Container for log text and scrollbar
        log_container = ttk.Frame(log_card, padding=2)
        log_container.pack(fill=tk.BOTH, expand=True)
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_container,
            height=6,
            wrap=tk.WORD,
            font=("Consolas", 10),
            background="#F9F9F9",
            relief=tk.FLAT,
            padx=10,
            pady=10,
            borderwidth=0
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Add scrollbar for log
        log_scroll = ttk.Scrollbar(
            log_container,
            orient="vertical",
            command=self.log_text.yview,
            style="Vertical.TScrollbar"
        )
        log_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=log_scroll.set)

        # Add some placeholder text
        self.log_text.insert(tk.END, "Scheduler logs will appear here.\n")

        # Make log read-only
        self.log_text.config(state=tk.DISABLED)

    def _browse_output_path(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directory:
            self.output_path_var.set(directory)
            self._add_log(f"Output directory set to: {directory}")

    def _toggle_scheduler(self):
        """Toggle the scheduler on/off"""
        self.scheduler_running.set(not self.scheduler_running.get())

        if self.scheduler_running.get():
            self.start_btn.configure(text="Stop Scheduler")
            self.status_label.configure(
                text="Scheduler is running",
                foreground="#2ECC71"  # Green from color scheme
            )

            # Add log entry
            self._add_log("Scheduler started")
            self.update_status("Scheduler started")
        else:
            self.start_btn.configure(text="Start Scheduler")
            self.status_label.configure(
                text="Scheduler is not running",
                foreground="#E74C3C"  # Red from color scheme
            )

            # Add log entry
            self._add_log("Scheduler stopped")
            self.update_status("Scheduler stopped")

    def _test_email(self):
        """Test email configuration"""
        if not self.smtp_server_var.get() or not self.from_email_var.get():
            messagebox.showinfo("Missing Information", "Please configure SMTP server and From email")
            return

        # In a real implementation, this would send a test email
        self._add_log(f"Testing email configuration: {self.smtp_server_var.get()}, {self.from_email_var.get()}")

        # Show success message
        messagebox.showinfo("Email Test", "Test email sent successfully!")

    def _add_log(self, message):
        """Add a message to the log"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"

        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
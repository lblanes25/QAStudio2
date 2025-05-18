# enhanced_qa_analytics_app.py
import os
import tkinter as tk
from tkinter import ttk, messagebox
import logging
from typing import Dict, Optional, Tuple

# Import managers and utilities
from qa_analytics.core.config_manager import ConfigManager
from qa_analytics.core.data_source_manager import DataSourceManager
from qa_analytics.core.reference_data_manager import ReferenceDataManager
from qa_analytics.core.enhanced_data_processor import EnhancedDataProcessor
from qa_analytics.reporting.enhanced_report_generator import EnhancedReportGenerator
from qa_analytics.templates.template_manager import TemplateManager
from qa_analytics.utils.modern_theme_manager import ModernThemeManager

# Import tabs
from qa_analytics.tabs.config_wizard_tab import ConfigWizardTab
from qa_analytics.tabs.run_analytics_tab import RunAnalyticsTab
from qa_analytics.tabs.data_sources_tab import DataSourcesTab
from qa_analytics.tabs.reference_data_tab import ReferenceDataTab
from qa_analytics.tabs.testing_tab import TestingTab
from qa_analytics.tabs.scheduler_tab import SchedulerTab

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("qa_analytics.log"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger("qa_analytics")


class EnhancedQAAnalyticsApp:
    """
    Enhanced QA Analytics Application with modern UI and extended capabilities.

    This application provides a comprehensive interface for managing and running
    QA analytics with a clean, modern design and improved user experience.
    """

    def __init__(self, root: tk.Tk):
        """
        Initialize the QA Analytics application.

        Args:
            root: Root tkinter window
        """
        self.root = root
        self.root.title("Enhanced QA Analytics")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)

        # Initialize managers
        self.config_manager = ConfigManager()
        self.data_source_manager = DataSourceManager()
        self.reference_data_manager = ReferenceDataManager()
        self.template_manager = TemplateManager()

        # Apply modern theme
        self.theme_manager = ModernThemeManager(root)
        self.theme_manager.apply_theme()

        # Create UI
        self._create_ui()

        # Log application start
        logger.info("Enhanced QA Analytics application started")

    def _create_ui(self):
        """Create the main application UI"""
        # Configure main grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        # Create header with app title and status bar
        self._create_header()

        # Create main content area
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=5)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)

        # Create tabs
        self._create_tabs()

        # Create footer (version, copyright)
        self._create_footer()

    def _create_header(self):
        """Create application header with title and status bar"""
        header_frame = ttk.Frame(self.root, style="Card.TFrame")
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)
        header_frame.columnconfigure(1, weight=1)  # Status bar expands

        # Application logo and title (left side)
        title_frame = ttk.Frame(header_frame)
        title_frame.grid(row=0, column=0, sticky=tk.W, padx=15, pady=10)

        # App title with larger font
        app_title = ttk.Label(
            title_frame,
            text="Enhanced QA Analytics",
            font=("Segoe UI", 16, "bold"),
            foreground=self.theme_manager.colors['primary']
        )
        app_title.pack(side=tk.LEFT)

        # Version
        version_label = ttk.Label(
            title_frame,
            text="v2.0",
            style="Small.TLabel"
        )
        version_label.pack(side=tk.LEFT, padx=(10, 0), pady=(5, 0))

        # Status bar (right side, expands to fill space)
        status_frame = ttk.Frame(header_frame)
        status_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=15, pady=10)
        status_frame.columnconfigure(0, weight=1)

        # Status message (left-aligned in status frame)
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            style="Small.TLabel"
        )
        self.status_label.grid(row=0, column=0, sticky=tk.W)

    def _create_tabs(self):
        """Create tabbed interface for main content"""
        # Create notebook
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Create tabs
        self.run_tab = RunAnalyticsTab(
            self.notebook,
            self._update_status
        )
        self.notebook.add(self.run_tab, text="Run Analytics")

        self.config_tab = ConfigWizardTab(
            self.notebook,
            self.config_manager,
            self.template_manager,
            self._on_config_saved
        )
        self.notebook.add(self.config_tab, text="Configuration Wizard")

        self.config_tab.debug_initialization()

        self.data_sources_tab = DataSourcesTab(
            self.notebook,
            self._update_status
        )
        self.notebook.add(self.data_sources_tab, text="Data Sources")

        self.reference_tab = ReferenceDataTab(
            self.notebook,
            self._update_status
        )
        self.notebook.add(self.reference_tab, text="Reference Data")

        self.testing_tab = TestingTab(
            self.notebook,
            self._update_status
        )
        self.testing_tab.data_processor_class = EnhancedDataProcessor
        self.testing_tab.report_generator_class = EnhancedReportGenerator
        self.notebook.add(self.testing_tab, text="Testing")

        self.scheduler_tab = SchedulerTab(
            self.notebook,
            self._update_status
        )
        self.notebook.add(self.scheduler_tab, text="Scheduler")

        # Bind tab change event
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _create_footer(self):
        """Create application footer"""
        footer_frame = ttk.Frame(self.root)
        footer_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=10, pady=(0, 10))
        footer_frame.columnconfigure(1, weight=1)

        # Copyright
        copyright_label = ttk.Label(
            footer_frame,
            text="© 2025 QA Analytics Team",
            style="Small.TLabel"
        )
        copyright_label.grid(row=0, column=0, sticky=tk.W, padx=10)

        # Help link
        help_link = ttk.Label(
            footer_frame,
            text="Help & Documentation",
            foreground=self.theme_manager.colors['secondary'],
            cursor="hand2",
            style="Small.TLabel"
        )
        help_link.grid(row=0, column=2, sticky=tk.E, padx=10)
        help_link.bind("<Button-1>", lambda e: self._show_help())

    def _update_status(self, message: str):
        """
        Update the status bar message

        Args:
            message: Status message to display
        """
        self.status_var.set(message)
        logger.info(f"Status: {message}")

    def _on_tab_changed(self, event):
        """
        Handle tab change event

        Args:
            event: Tab change event
        """
        # Get the selected tab name
        tab_id = self.notebook.select()
        tab_name = self.notebook.tab(tab_id, "text")

        # Update status
        self._update_status(f"Switched to {tab_name} tab")

    def _on_config_saved(self):
        """Handle event when a configuration is saved"""
        # Refresh configurations
        self.config_manager.load_all_configs()

        # Update available configs in Run Analytics tab
        if hasattr(self.run_tab, "reload_analytics"):
            self.run_tab.reload_analytics()

    def _show_help(self):
        """Show help information"""
        help_text = """
        Enhanced QA Analytics Help

        Run Analytics Tab:
        - Select a QA-ID from the dropdown
        - Choose a data file to process
        - Set the output directory
        - Click 'Run Analysis' to process the data

        Configuration Wizard:
        - Create new or edit existing analytics configurations
        - Step through template selection, basic settings, and validations
        - Review and save your configuration

        Data Sources Tab:
        - Manage data source definitions
        - Update and refresh data source registry

        Reference Data Tab:
        - Manage reference data files
        - Update and check freshness status

        Testing Tab:
        - Test analytics configurations with sample data
        - Generate reports and validate results

        Scheduler Tab:
        - Configure automated analytics runs
        - Set schedules and email notifications
        """

        # Create help dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("QA Analytics Help")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()

        # Apply theme to dialog
        dialog.configure(background=self.theme_manager.colors['bg'])

        # Create content
        content_frame = ttk.Frame(dialog, padding=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(1, weight=1)

        # Title
        ttk.Label(
            content_frame,
            text="QA Analytics Help",
            style="Header.TLabel"
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 15))

        # Help content in scrollable text widget
        text_frame = ttk.Frame(content_frame, style="Card.TFrame")
        text_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)

        # Text widget with scrollbar
        text_container = ttk.Frame(text_frame, padding=2)
        text_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_container.columnconfigure(0, weight=1)
        text_container.rowconfigure(0, weight=1)

        help_text_widget = tk.Text(
            text_container,
            wrap=tk.WORD,
            background=self.theme_manager.colors['light_bg'],
            relief=tk.FLAT,
            padx=15,
            pady=15,
            borderwidth=0
        )
        help_text_widget.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        help_text_widget.insert(tk.END, help_text)
        help_text_widget.config(state=tk.DISABLED)

        # Scrollbar
        scrollbar = ttk.Scrollbar(
            text_container,
            orient="vertical",
            command=help_text_widget.yview,
            style="Vertical.TScrollbar"
        )
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        help_text_widget.config(yscrollcommand=scrollbar.set)

        # Close button
        ttk.Button(
            content_frame,
            text="Close",
            command=dialog.destroy
        ).grid(row=2, column=0, sticky=tk.E, pady=(15, 0))


def main():
    """Main entry point for the application"""
    root = tk.Tk()
    app = EnhancedQAAnalyticsApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
# tabs/data_sources_tab.py
import tkinter as tk
from tkinter import ttk, messagebox
import datetime
from typing import Callable


class DataSourcesTab(ttk.Frame):
    """
    Tab for managing data sources registered in the system.
    Provides a clean, modern interface for viewing, refreshing and examining
    data source details.
    """

    def __init__(self, parent, status_callback: Callable):
        """
        Initialize the Data Sources tab.

        Args:
            parent: Parent widget
            status_callback: Function to call to update status bar
        """
        super().__init__(parent, padding="20 15 20 15")
        self.parent = parent
        self.update_status = status_callback

        # Create widgets
        self._create_widgets()

        # Load data sources
        self._populate_data_source_tree()

    def _create_widgets(self):
        """Create all widgets for this tab"""
        # Use grid layout with proper expansion
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)  # Tree section expands
        self.rowconfigure(1, weight=0)  # Button section fixed height

        # Create header and description
        header_frame = ttk.Frame(self)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))

        ttk.Label(
            header_frame,
            text="Data Source Registry",
            style="Header.TLabel"
        ).pack(side=tk.LEFT)

        description = ttk.Label(
            header_frame,
            text="Manage data sources configured for analytics processing",
            style="Small.TLabel"
        )
        description.pack(side=tk.LEFT, padx=(15, 0))

        # Create a card-like container for the tree
        tree_container = ttk.Frame(self, style="Card.TFrame")
        tree_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        # Create inner frame with padding
        tree_frame = ttk.Frame(tree_container, padding=2)
        tree_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        # Create treeview for data sources
        columns = ("Name", "Type", "Owner", "Version", "Last Updated", "Analytics")
        self.source_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            height=15
        )

        # Configure columns - more consistent widths
        self.source_tree.column("Name", width=180, anchor=tk.W)
        self.source_tree.column("Type", width=80, anchor=tk.W)
        self.source_tree.column("Owner", width=180, anchor=tk.W)
        self.source_tree.column("Version", width=80, anchor=tk.W)
        self.source_tree.column("Last Updated", width=150, anchor=tk.W)
        self.source_tree.column("Analytics", width=80, anchor=tk.CENTER)

        # Configure headings
        for col in columns:
            self.source_tree.heading(col, text=col)

        # Add scrollbars - vertical and horizontal
        y_scrollbar = ttk.Scrollbar(
            tree_frame,
            orient=tk.VERTICAL,
            command=self.source_tree.yview,
            style="Vertical.TScrollbar"
        )

        x_scrollbar = ttk.Scrollbar(
            tree_frame,
            orient=tk.HORIZONTAL,
            command=self.source_tree.xview,
            style="Horizontal.TScrollbar"
        )

        self.source_tree.configure(
            yscrollcommand=y_scrollbar.set,
            xscrollcommand=x_scrollbar.set
        )

        # Position tree and scrollbars
        self.source_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        y_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        x_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # Button frame with action buttons
        button_frame = ttk.Frame(self)
        button_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 0))

        ttk.Button(
            button_frame,
            text="Refresh Registry",
            command=self._refresh_registry
        ).pack(side=tk.LEFT)

        ttk.Button(
            button_frame,
            text="Add New...",
            command=self._add_new_source
        ).pack(side=tk.LEFT, padx=(10, 0))

        ttk.Button(
            button_frame,
            text="View Details",
            style="Primary.TButton",
            command=self._view_details
        ).pack(side=tk.RIGHT)

    def _populate_data_source_tree(self):
        """Populate the data source tree with sample data"""
        # Clear existing items
        for item in self.source_tree.get_children():
            self.source_tree.delete(item)

        # Sample data - would be replaced with actual data in real implementation
        sample_sources = [
            ("audit_workpaper_approvals", "report", "Quality Assurance Team", "1.0", "2025-05-01", 1),
            ("third_party_risk", "report", "Risk Management", "1.1", "2025-04-15", 1),
            ("audit_planning_approvals", "report", "QA Team", "1.0", "2025-05-01", 3),
            ("risk_assessment_validation", "report", "Risk Management Team", "1.1", "2025-04-15", 3),
            ("audit_workpapers_2025q2", "report", "QA Analytics", "1.0", "2025-05-15", 1)
        ]

        # Add to tree with alternate row coloring
        for i, source in enumerate(sample_sources):
            item_id = self.source_tree.insert("", tk.END, values=source)

            # Apply alternating row colors for better readability
            if i % 2 == 1:
                self.source_tree.item(item_id, tags=("odd",))

        # Configure tag for alternating rows
        self.source_tree.tag_configure("odd", background="#F5F9FC")

    def _refresh_registry(self):
        """Refresh the data source registry"""
        # In a real implementation, this would reload from the data source registry
        self._populate_data_source_tree()
        self.update_status("Data source registry refreshed")

    def _add_new_source(self):
        """Add a new data source"""
        # This would open a dialog to add a new data source
        messagebox.showinfo(
            "Add Data Source",
            "This feature would allow adding a new data source to the registry."
        )

    def _view_details(self):
        """View details for the selected data source"""
        selection = self.source_tree.selection()
        if not selection:
            messagebox.showinfo("No Selection", "Please select a data source first")
            return

        # Get data for the selected item
        item = self.source_tree.item(selection[0])
        values = item['values']

        if not values:
            return

        # Extract data source name
        source_name = values[0]

        # Create modern dialog to show details
        details_dialog = tk.Toplevel(self)
        details_dialog.title(f"Data Source Details: {source_name}")
        details_dialog.geometry("650x550")
        details_dialog.transient(self)  # Set to be on top of the parent window
        details_dialog.grab_set()  # Modal dialog

        # Apply appropriate padding
        dialog_frame = ttk.Frame(details_dialog, padding="20 15 20 15")
        dialog_frame.pack(fill=tk.BOTH, expand=True)
        dialog_frame.columnconfigure(0, weight=1)
        dialog_frame.rowconfigure(2, weight=1)  # Details section should expand

        # Dialog header
        ttk.Label(
            dialog_frame,
            text=f"Data Source: {source_name}",
            style="Header.TLabel"
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 10))

        # Basic information section in a card-like container
        info_frame = ttk.Frame(dialog_frame, style="Card.TFrame")
        info_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))

        # Two-column grid for basic info
        info_grid = ttk.Frame(info_frame, padding=15)
        info_grid.pack(fill=tk.X)

        # Info grid with labels and values - 2x3 grid
        # Row 1
        ttk.Label(info_grid, text="Type:", width=12, style="Subheader.TLabel").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 8))
        ttk.Label(info_grid, text=values[1]).grid(
            row=0, column=1, sticky=tk.W, padx=(0, 20), pady=(0, 8))

        ttk.Label(info_grid, text="Version:", width=12, style="Subheader.TLabel").grid(
            row=0, column=2, sticky=tk.W, padx=(0, 10), pady=(0, 8))
        ttk.Label(info_grid, text=values[3]).grid(
            row=0, column=3, sticky=tk.W, pady=(0, 8))

        # Row 2
        ttk.Label(info_grid, text="Owner:", width=12, style="Subheader.TLabel").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 8))
        ttk.Label(info_grid, text=values[2]).grid(
            row=1, column=1, sticky=tk.W, padx=(0, 20), pady=(0, 8))

        ttk.Label(info_grid, text="Last Updated:", width=12, style="Subheader.TLabel").grid(
            row=1, column=2, sticky=tk.W, padx=(0, 10), pady=(0, 8))
        ttk.Label(info_grid, text=values[4]).grid(
            row=1, column=3, sticky=tk.W, pady=(0, 8))

        # Row 3
        ttk.Label(info_grid, text="Analytics:", width=12, style="Subheader.TLabel").grid(
            row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 0))
        ttk.Label(info_grid, text=str(values[5])).grid(
            row=2, column=1, sticky=tk.W, padx=(0, 20), pady=(0, 0))

        # Detailed configuration section
        config_frame = ttk.LabelFrame(dialog_frame, text="Configuration Details", padding=10)
        config_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        config_frame.columnconfigure(0, weight=1)
        config_frame.rowconfigure(0, weight=1)

        # Create text widget container
        text_container = ttk.Frame(config_frame)
        text_container.pack(fill=tk.BOTH, expand=True)
        text_container.columnconfigure(0, weight=1)
        text_container.rowconfigure(0, weight=1)

        # Create text widget for detailed configuration
        details_text = tk.Text(
            text_container,
            wrap=tk.WORD,
            font=("Consolas", 10),
            background="#F9F9F9",
            relief=tk.FLAT,
            padx=10,
            pady=10,
            borderwidth=0
        )
        details_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Add scrollbar for details text
        details_scroll = ttk.Scrollbar(
            text_container,
            orient=tk.VERTICAL,
            command=details_text.yview,
            style="Vertical.TScrollbar"
        )
        details_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        details_text.config(yscrollcommand=details_scroll.set)

        # Populate with sample configuration details
        if source_name == "audit_workpaper_approvals":
            config_text = """# Data Source Configuration

type: "report"
description: "Audit workpaper approvals tracking data"
version: "1.0"
owner: "Quality Assurance Team"
refresh_frequency: "Weekly"
last_updated: "2025-05-01"
file_type: "xlsx"
file_pattern: "Workpaper_Approvals_{YYYY}{MM}*.xlsx"

# Key Columns
key_columns:
  - Audit TW ID

# Validation Rules
validation_rules:
  - type: row_count_min
    threshold: 10
    description: "Should have at least 10 records"
  - type: required_columns
    columns:
      - Audit TW ID
      - TW submitter
      - TL approver
      - AL approver
      - Submit Date
      - TL Approval Date
      - AL Approval Date
    description: "Critical columns that must be present"

# Column Mappings
columns_mapping:
  - source: "Audit TW ID"
    target: "Audit TW ID"
    data_type: "string"
  - source: "TW submitter"
    target: "TW submitter"
    data_type: "string"
  - source: "TL approver"
    target: "TL approver"
    data_type: "string"
  - source: "AL approver"
    target: "AL approver"
    data_type: "string"
  - source: "Submit Date"
    target: "Submit Date"
    data_type: "date"
  - source: "TL Approval Date"
    target: "TL Approval Date"
    data_type: "date"
  - source: "AL Approval Date"
    target: "AL Approval Date"
    data_type: "date"

# Associated Analytics
associated_analytics:
  - QA-77: Audit Test Workpaper Approvals"""
        else:
            config_text = "# Data Source Configuration\n\nDetailed configuration information would be shown here."

        details_text.insert(tk.END, config_text)

        # Make text read-only
        details_text.config(state=tk.DISABLED)

        # Action buttons
        button_frame = ttk.Frame(dialog_frame)
        button_frame.grid(row=3, column=0, sticky=tk.E, pady=(0, 0))

        ttk.Button(
            button_frame,
            text="Export Configuration",
            command=lambda: self._export_config(source_name)
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(
            button_frame,
            text="Close",
            style="Primary.TButton",
            command=details_dialog.destroy
        ).pack(side=tk.LEFT)

    def _export_config(self, source_name):
        """Export the configuration for a data source"""
        # This would actually save the configuration to a file
        messagebox.showinfo(
            "Export Configuration",
            f"Configuration for '{source_name}' would be exported to a file."
        )
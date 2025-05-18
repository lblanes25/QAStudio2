# tabs/reference_data_tab.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime
from typing import Callable
import os


class ReferenceDataTab(ttk.Frame):
    """
    Tab for managing reference data files with a modern, clean interface.
    Allows users to view, update, and monitor reference data freshness.
    """

    def __init__(self, parent, status_callback: Callable):
        """
        Initialize the Reference Data tab.

        Args:
            parent: Parent widget
            status_callback: Function to call to update status bar
        """
        super().__init__(parent, padding="20 15 20 15")
        self.parent = parent
        self.update_status = status_callback

        # Create widgets with modern styling
        self._create_widgets()

        # Load reference data
        self._populate_reference_tree()

    def _create_widgets(self):
        """Create all widgets for this tab with modern styling"""
        # Use grid layout for better control
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)  # Tree should expand
        self.rowconfigure(1, weight=0)  # Button row

        # Tab title and description section
        header_frame = ttk.Frame(self)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))

        ttk.Label(
            header_frame,
            text="Reference Data Management",
            style="Header.TLabel"
        ).pack(side=tk.LEFT)

        # Create a card-like container for the treeview
        tree_card = ttk.Frame(self, style="Card.TFrame")
        tree_card.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=2, pady=2)
        tree_card.columnconfigure(0, weight=1)
        tree_card.rowconfigure(0, weight=1)

        # Create treeview container with proper padding
        tree_container = ttk.Frame(tree_card, padding=2)
        tree_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        # Create columns with modern sizing
        columns = ("Name", "Format", "Version", "Last Modified", "Rows", "Freshness")
        self.ref_tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show="headings",
            height=15,
            selectmode="browse"
        )

        # Configure columns
        self.ref_tree.column("Name", width=180, anchor=tk.W)
        self.ref_tree.column("Format", width=100, anchor=tk.W)
        self.ref_tree.column("Version", width=100, anchor=tk.W)
        self.ref_tree.column("Last Modified", width=180, anchor=tk.W)
        self.ref_tree.column("Rows", width=70, anchor=tk.CENTER)
        self.ref_tree.column("Freshness", width=120, anchor=tk.W)

        # Configure headings
        for col in columns:
            self.ref_tree.heading(col, text=col)

        # Add vertical scrollbar with modern styling
        y_scrollbar = ttk.Scrollbar(
            tree_container,
            orient="vertical",
            command=self.ref_tree.yview,
            style="Vertical.TScrollbar"
        )
        self.ref_tree.configure(yscrollcommand=y_scrollbar.set)

        # Add horizontal scrollbar with modern styling
        x_scrollbar = ttk.Scrollbar(
            tree_container,
            orient="horizontal",
            command=self.ref_tree.xview,
            style="Horizontal.TScrollbar"
        )
        self.ref_tree.configure(xscrollcommand=x_scrollbar.set)

        # Pack tree and scrollbars
        self.ref_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        y_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        x_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # Define tag colors for different statuses
        self.ref_tree.tag_configure("fresh", background="#E8F8F5")  # Light green
        self.ref_tree.tag_configure("stale", background="#FEF5E7")  # Light orange
        self.ref_tree.tag_configure("not_loaded", background="#F4F6F7")  # Light gray

        # Button section for actions
        button_frame = ttk.Frame(self)
        button_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(20, 0))

        # Left-aligned buttons
        left_buttons = ttk.Frame(button_frame)
        left_buttons.pack(side=tk.LEFT)

        ttk.Button(
            left_buttons,
            text="Refresh Status",
            command=self._refresh_status
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(
            left_buttons,
            text="Update Reference File",
            style="Primary.TButton",
            command=self._update_file
        ).pack(side=tk.LEFT)

        # Right-aligned buttons
        ttk.Button(
            button_frame,
            text="View Update History",
            command=self._view_history
        ).pack(side=tk.RIGHT)

        # Info section explaining reference data
        info_frame = ttk.Frame(self, style="Card.TFrame")
        info_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(20, 0), padx=2)

        info_label = ttk.Label(
            info_frame,
            text=(
                "Reference data files are used for validation lookups and data enrichment. "
                "Files marked as stale (orange) may need to be updated. "
                "Select a reference data file and click 'Update Reference File' to update it."
            ),
            style="Info.TLabel",
            wraplength=700,
            justify=tk.LEFT
        )
        info_label.pack(padx=15, pady=15, fill=tk.X)

    def _populate_reference_tree(self):
        """Populate the reference data tree with sample data"""
        # Clear existing items
        for item in self.ref_tree.get_children():
            self.ref_tree.delete(item)

        # Sample data - would be replaced with actual data in real implementation
        sample_data = [
            ("HR_Titles", "dictionary", "2025-Q2", "2025-04-15 10:30", 250, "✓ Fresh", "fresh"),
            ("Risk_Categories", "dataframe", "2025-01", "2025-01-20 14:22", 15, "⚠ Stale", "stale"),
            ("Control_Standards", "dataframe", "2025-05", "2025-05-01 09:15", 75, "✓ Fresh", "fresh"),
            ("Audit_Leaders", "dictionary", "1.0", "Not loaded", "-", "Not loaded", "not_loaded")
        ]

        # Add to tree with appropriate tags for color coding
        for data in sample_data:
            self.ref_tree.insert("", tk.END, values=data[:-1], tags=(data[-1],))

    def _refresh_status(self):
        """Refresh reference data status"""
        # In a real implementation, this would check the actual status of reference data files
        self._populate_reference_tree()
        self.update_status("Reference data status refreshed")

    def _update_file(self):
        """Update a reference data file with modern file dialog"""
        selection = self.ref_tree.selection()
        if not selection:
            messagebox.showinfo("No Selection", "Please select a reference data entry first")
            return

        # Get data for the selected item
        item = self.ref_tree.item(selection[0])
        values = item['values']

        if not values:
            return

        # Extract reference data name
        ref_name = values[0]

        # Ask for file with modern dialog
        filename = filedialog.askopenfilename(
            title=f"Select New File for Reference Data '{ref_name}'",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("CSV Files", "*.csv"),
                ("All Files", "*.*")
            ],
            initialdir=os.path.join(os.getcwd(), "ref_data")
        )

        if not filename:
            return

        # Confirm update with modern dialog
        confirm = messagebox.askyesno(
            "Confirm Update",
            f"Are you sure you want to update reference data '{ref_name}' with file:\n{filename}?",
            icon="question"
        )

        if confirm:
            # In a real implementation, this would actually update the reference data file
            messagebox.showinfo(
                "Success",
                f"Reference data '{ref_name}' updated successfully",
                icon="info"
            )

            # Refresh the tree to show updated status
            self._refresh_status()

    def _view_history(self):
        """View update history for reference data with modern dialog"""
        # Create dialog to show history
        history_dialog = tk.Toplevel(self)
        history_dialog.title("Reference Data Update History")
        history_dialog.geometry("700x500")
        history_dialog.transient(self)  # Set to be on top of the parent window
        history_dialog.grab_set()  # Modal dialog

        # Set dialog in the center of the parent window
        x = self.winfo_rootx() + (self.winfo_width() // 2) - (700 // 2)
        y = self.winfo_rooty() + (self.winfo_height() // 2) - (500 // 2)
        history_dialog.geometry(f"+{x}+{y}")

        # Add content to dialog
        content_frame = ttk.Frame(history_dialog, padding=20)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Dialog title
        ttk.Label(
            content_frame,
            text="Reference Data Update History",
            style="Header.TLabel"
        ).pack(anchor=tk.W, pady=(0, 15))

        # History content in a card-like container
        history_card = ttk.Frame(content_frame, style="Card.TFrame")
        history_card.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # Container for text and scrollbar
        history_container = ttk.Frame(history_card, padding=2)
        history_container.pack(fill=tk.BOTH, expand=True)
        history_container.columnconfigure(0, weight=1)
        history_container.rowconfigure(0, weight=1)

        # Create text widget with monospace font and subtle background
        history_text = tk.Text(
            history_container,
            wrap=tk.WORD,
            font=('Consolas', 10),
            background='#F9F9F9',
            relief=tk.FLAT,
            padx=15,
            pady=15,
            borderwidth=0
        )
        history_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Add scrollbar with modern styling
        history_scroll = ttk.Scrollbar(
            history_container,
            orient="vertical",
            command=history_text.yview,
            style="Vertical.TScrollbar"
        )
        history_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        history_text.configure(yscrollcommand=history_scroll.set)

        # Populate with sample history with improved formatting
        history_text.insert(tk.END, "Time: 2025-05-15 14:32:45\n", "header")
        history_text.insert(tk.END, "User: admin\n")
        history_text.insert(tk.END, "Action: update_reference\n")
        history_text.insert(tk.END, "Reference Data: Control_Standards\n")
        history_text.insert(tk.END, "Previous Version: 2025-04 (Modified: 2025-04-01)\n")
        history_text.insert(tk.END, "New Version: 2025-05 (Modified: 2025-05-01)\n")
        history_text.insert(tk.END, "\n" + "─" * 50 + "\n\n")  # Unicode separator

        history_text.insert(tk.END, "Time: 2025-04-15 10:30:22\n", "header")
        history_text.insert(tk.END, "User: hr_admin\n")
        history_text.insert(tk.END, "Action: update_reference\n")
        history_text.insert(tk.END, "Reference Data: HR_Titles\n")
        history_text.insert(tk.END, "Previous Version: 2025-Q1 (Modified: 2025-01-15)\n")
        history_text.insert(tk.END, "New Version: 2025-Q2 (Modified: 2025-04-15)\n")
        history_text.insert(tk.END, "\n" + "─" * 50 + "\n\n")  # Unicode separator

        history_text.insert(tk.END, "Time: 2025-01-20 14:22:10\n", "header")
        history_text.insert(tk.END, "User: risk_admin\n")
        history_text.insert(tk.END, "Action: update_reference\n")
        history_text.insert(tk.END, "Reference Data: Risk_Categories\n")
        history_text.insert(tk.END, "Previous Version: 2024-10 (Modified: 2024-10-10)\n")
        history_text.insert(tk.END, "New Version: 2025-01 (Modified: 2025-01-20)\n")

        # Configure tags for styling the text
        history_text.tag_configure("header", font=('Consolas', 10, 'bold'))

        # Make text read-only
        history_text.config(state=tk.DISABLED)

        # Add close button with modern styling
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(fill=tk.X, pady=(15, 0))

        ttk.Button(
            button_frame,
            text="Close",
            command=history_dialog.destroy
        ).pack(side=tk.RIGHT)
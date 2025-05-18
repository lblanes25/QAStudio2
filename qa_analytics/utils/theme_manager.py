# utils/theme_manager.py
import tkinter as tk
from tkinter import ttk, font
import ttkthemes
import os


class ThemeManager:
    """
    Manages the application theme and styling to create a modern,
    minimalistic UI appearance.
    """

    def __init__(self, root):
        """
        Initialize the theme manager.

        Args:
            root: The root tkinter window
        """
        self.root = root

    def apply_theme(self):
        """Apply custom styling to create a modern UI"""
        # Configure root window background
        self.root.configure(background='white')

        # Create a modern style
        self.style = ttk.Style()

        # Try to use a clean font
        available_fonts = font.families()
        preferred_fonts = ['Inter', 'Helvetica Neue', 'Segoe UI', 'SF UI Text', 'Arial']

        # Find the first available preferred font
        ui_font = next((f for f in preferred_fonts if f in available_fonts), None)
        if not ui_font:
            # Default to system font if preferred fonts are not available
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
        self.fresh_color = '#e6ffe6'  # Light green
        self.stale_color = '#fff0e6'  # Light orange
        self.not_loaded_color = '#f0f0f0'  # Light gray

        # Configure widget styles

        # TFrame - regular frames
        self.style.configure('TFrame', background=bg_color)

        # TLabel - text labels
        self.style.configure('TLabel', background=bg_color, font=normal_font)
        self.style.configure('Header.TLabel', font=header_font)
        self.style.configure('Small.TLabel', font=small_font)

        # TButton - buttons with rounded corners
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

        # Configure specific tag colors for treeview
        # (used in reference data and other tables for status indicators)
        # self.style.map('Treeview', tags={
        #    'fresh': {'background': self.fresh_color},
        #    'stale': {'background': self.stale_color},
        #    'not_loaded': {'background': self.not_loaded_color},
        # })
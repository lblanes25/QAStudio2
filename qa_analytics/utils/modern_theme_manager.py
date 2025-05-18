# utils/modern_theme_manager.py (Updated with visual fixes)
import tkinter as tk
from tkinter import ttk, font
from PIL import Image, ImageTk, ImageDraw


class ModernThemeManager:
    """
    Manages a modern, minimalistic UI theme for the QA Analytics application.
    Provides consistent styling across all components with clean white backgrounds,
    rounded buttons, and a professional appearance.
    """

    def __init__(self, root):
        """
        Initialize the theme manager.

        Args:
            root: The root tkinter window
        """
        self.root = root

        # Define color scheme with softer tones and better contrast
        self.colors = {
            'bg': '#FFFFFF',  # White background
            'primary': '#2C3E50',  # Softer blue-black for primary elements
            'primary_hover': '#34495E',  # Slightly lighter shade for hover states
            'secondary': '#3498DB',  # Accent blue for secondary elements
            'light_bg': '#F9F9F9',  # Very light gray for content areas
            'border': '#E0E0E0',  # Light gray for borders
            'tab_active': '#F5F9FC',  # Very light blue for active tab
            'tab_hover': '#F0F4F8',  # Slightly darker for hover state
            'success': '#2ECC71',  # Green for success indicators
            'warning': '#F39C12',  # Orange for warnings
            'error': '#E74C3C',  # Red for errors
            'text': '#2C3E50',  # Main text color
            'text_secondary': '#7F8C8D',  # Secondary text color
            'fresh': '#E8F8F5',  # Light green for fresh data
            'stale': '#FEF5E7',  # Light orange for stale data
            'not_loaded': '#F4F6F7'  # Light gray for not loaded data
        }

        # Font configuration
        self.setup_fonts()

    def setup_fonts(self):
        """Configure fonts based on system availability"""
        available_fonts = font.families()
        preferred_fonts = ['Segoe UI', 'Inter', 'Helvetica Neue', 'Arial', 'SF Pro Display']

        # Find the first available preferred font
        self.ui_font = next((f for f in preferred_fonts if f in available_fonts), None)
        if not self.ui_font:
            # Default to system font if preferred fonts are not available
            self.ui_font = "TkDefaultFont"

        # Find first available monospace font
        mono_fonts = ['Consolas', 'Courier New', 'Courier', 'Monaco']
        self.mono_font = next((f for f in mono_fonts if f in available_fonts), 'Courier')

        # Define font configurations
        self.fonts = {
            'header': (self.ui_font, 11, 'bold'),
            'subheader': (self.ui_font, 10, 'bold'),
            'normal': (self.ui_font, 10),
            'small': (self.ui_font, 9),
            'code': (self.mono_font, 10),
            'button': (self.ui_font, 10),
            'button_primary': (self.ui_font, 10, 'bold')
        }

    def apply_theme(self):
        """Apply the modern theme to all ttk widgets"""
        # Configure root window
        self.root.configure(background=self.colors['bg'])

        # Create style object
        self.style = ttk.Style(self.root)

        # Use a clean base theme
        self.style.theme_use('clam')

        # Configure base styles
        self._configure_frames()
        self._configure_labels()
        self._configure_buttons()
        self._configure_entries()
        self._configure_comboboxes()
        self._configure_notebook()
        self._configure_treeview()
        self._configure_scrollbars()
        self._configure_progressbar()
        self._configure_checkbutton()
        self._configure_radiobutton()

    def _configure_frames(self):
        """Configure frame styles"""
        # Basic frame
        self.style.configure('TFrame',
                             background=self.colors['bg'])

        # Card frame (with subtle border and shadow effect)
        self.style.configure('Card.TFrame',
                             background=self.colors['bg'],
                             relief='solid',
                             borderwidth=1,
                             bordercolor=self.colors['border'])

        # Section frame (with a heading background)
        self.style.configure('Section.TFrame',
                             background=self.colors['bg'],
                             padding=15)

        # LabelFrame with modern styling
        self.style.configure('TLabelframe',
                             background=self.colors['bg'],
                             padding=10,
                             borderwidth=1,
                             relief='solid',
                             bordercolor=self.colors['border'])

        self.style.configure('TLabelframe.Label',
                             background=self.colors['bg'],
                             foreground=self.colors['text'],
                             font=self.fonts['subheader'],
                             padding=(5, 0))

    def _configure_labels(self):
        """Configure label styles"""
        # Regular label
        self.style.configure('TLabel',
                             background=self.colors['bg'],
                             foreground=self.colors['text'],
                             font=self.fonts['normal'])

        # Header label
        self.style.configure('Header.TLabel',
                             background=self.colors['bg'],
                             foreground=self.colors['text'],
                             font=self.fonts['header'])

        # Subheader label
        self.style.configure('Subheader.TLabel',
                             background=self.colors['bg'],
                             foreground=self.colors['text'],
                             font=self.fonts['subheader'])

        # Small label (for secondary text)
        self.style.configure('Small.TLabel',
                             background=self.colors['bg'],
                             foreground=self.colors['text_secondary'],
                             font=self.fonts['small'])

        # Info label (with different background)
        self.style.configure('Info.TLabel',
                             background=self.colors['light_bg'],
                             foreground=self.colors['text'],
                             font=self.fonts['normal'],
                             padding=10)

        # Status labels with clear color indicators
        self.style.configure('Success.TLabel',
                             foreground=self.colors['success'],
                             background=self.colors['bg'],
                             font=self.fonts['normal'])

        self.style.configure('Warning.TLabel',
                             foreground=self.colors['warning'],
                             background=self.colors['bg'],
                             font=self.fonts['normal'])

        self.style.configure('Error.TLabel',
                             foreground=self.colors['error'],
                             background=self.colors['bg'],
                             font=self.fonts['normal'])

    def _configure_buttons(self):
        """Configure button styles with rounded corners"""
        # Regular button with subtle styling
        self.style.configure('TButton',
                             font=self.fonts['button'],
                             background=self.colors['bg'],
                             foreground=self.colors['text'],
                             borderwidth=1,
                             relief="flat",
                             padding=(12, 6))

        self.style.map('TButton',
                       background=[('active', self.colors['light_bg']),
                                   ('disabled', self.colors['light_bg'])],
                       foreground=[('disabled', self.colors['text_secondary'])])

        # Primary button (accent color with white text)
        self.style.configure('Primary.TButton',
                             font=self.fonts['button_primary'],
                             foreground='white',
                             background=self.colors['primary'],
                             borderwidth=0,
                             relief="flat",
                             padding=(15, 8))

        self.style.map('Primary.TButton',
                       foreground=[('active', 'white'),
                                   ('disabled', self.colors['text_secondary'])],
                       background=[('active', self.colors['primary_hover']),
                                   ('disabled', self.colors['light_bg'])])

        # Secondary button
        self.style.configure('Secondary.TButton',
                             font=self.fonts['button'],
                             foreground='white',
                             background=self.colors['secondary'],
                             borderwidth=0,
                             relief="flat",
                             padding=(12, 6))

        self.style.map('Secondary.TButton',
                       foreground=[('active', 'white'),
                                   ('disabled', self.colors['text_secondary'])])

        # Icon button (for file browser buttons)
        self.style.configure('Icon.TButton',
                             font=self.fonts['button'],
                             background=self.colors['light_bg'],
                             foreground=self.colors['text'],
                             borderwidth=1,
                             relief="flat",
                             padding=(6, 6))

        self.style.map('Icon.TButton',
                       background=[('active', self.colors['border']),
                                   ('disabled', self.colors['light_bg'])])

    def _configure_entries(self):
        """Configure entry styles"""
        # Regular entry with better padding and border
        self.style.configure('TEntry',
                             font=self.fonts['normal'],
                             foreground=self.colors['text'],
                             fieldbackground=self.colors['bg'],
                             bordercolor=self.colors['border'],
                             borderwidth=1,
                             padding=8)

        self.style.map('TEntry',
                       fieldbackground=[('disabled', self.colors['light_bg'])],
                       foreground=[('disabled', self.colors['text_secondary'])])

    def _configure_comboboxes(self):
        """Configure combobox styles"""
        # Regular combobox with consistent styling
        self.style.configure('TCombobox',
                             font=self.fonts['normal'],
                             foreground=self.colors['text'],
                             background=self.colors['bg'],
                             fieldbackground=self.colors['bg'],
                             arrowsize=15,
                             padding=8)

        self.style.map('TCombobox',
                       fieldbackground=[('readonly', self.colors['bg']),
                                        ('disabled', self.colors['light_bg'])],
                       foreground=[('readonly', self.colors['text']),
                                   ('disabled', self.colors['text_secondary'])])

    def _configure_notebook(self):
        """Configure notebook (tabbed interface) styles"""
        # Notebook (tab container) with shadow effect
        self.style.configure('TNotebook',
                             background=self.colors['bg'],
                             borderwidth=0)

        # Create a clean, flat tab design
        self.style.configure('TNotebook.Tab',
                             font=self.fonts['normal'],
                             background=self.colors['light_bg'],
                             foreground=self.colors['text'],
                             borderwidth=0,
                             relief="flat",
                             padding=(20, 8))

        # Use underline effect for selected tab
        self.style.map('TNotebook.Tab',
                       background=[('selected', self.colors['tab_active']),
                                   ('active', self.colors['tab_hover'])],
                       foreground=[('selected', self.colors['primary']),
                                   ('active', self.colors['text'])],
                       font=[('selected', self.fonts['subheader'])])

    def _configure_treeview(self):
        """Configure treeview (table) styles"""
        # Modern treeview styling
        self.style.configure('Treeview',
                             font=self.fonts['normal'],
                             foreground=self.colors['text'],
                             background=self.colors['bg'],
                             fieldbackground=self.colors['bg'],
                             rowheight=30,
                             borderwidth=1)

        # Clean column headers
        self.style.configure('Treeview.Heading',
                             font=self.fonts['subheader'],
                             foreground=self.colors['text'],
                             background=self.colors['light_bg'],
                             relief='flat',
                             padding=5)

        self.style.map('Treeview.Heading',
                       background=[('active', self.colors['tab_hover'])])

        # Remove the focus dotted border
        self.style.layout("Treeview", [
            ('Treeview.treearea', {'sticky': 'nswe'})
        ])

    def _configure_scrollbars(self):
        """Configure scrollbar styles"""
        # Thin, modern scrollbars
        # Vertical scrollbar
        self.style.configure('Vertical.TScrollbar',
                             background=self.colors['light_bg'],
                             arrowsize=8,
                             borderwidth=0,
                             relief='flat',
                             width=8)

        self.style.map('Vertical.TScrollbar',
                       background=[('active', self.colors['border']),
                                   ('!active', self.colors['light_bg'])])

        # Horizontal scrollbar
        self.style.configure('Horizontal.TScrollbar',
                             background=self.colors['light_bg'],
                             arrowsize=8,
                             borderwidth=0,
                             relief='flat',
                             width=8)

        self.style.map('Horizontal.TScrollbar',
                       background=[('active', self.colors['border']),
                                   ('!active', self.colors['light_bg'])])

    def _configure_progressbar(self):
        """Configure progressbar styles"""
        # Thin, modern progressbar
        self.style.configure('TProgressbar',
                             background=self.colors['primary'],
                             troughcolor=self.colors['light_bg'],
                             borderwidth=0,
                             thickness=8)

        # Success progressbar (green)
        self.style.configure('Success.TProgressbar',
                             background=self.colors['success'],
                             troughcolor=self.colors['light_bg'],
                             borderwidth=0,
                             thickness=8)

    def _configure_checkbutton(self):
        """Configure checkbutton styles"""
        self.style.configure('TCheckbutton',
                             background=self.colors['bg'],
                             foreground=self.colors['text'],
                             font=self.fonts['normal'])

    def _configure_radiobutton(self):
        """Configure radiobutton styles"""
        self.style.configure('TRadiobutton',
                             background=self.colors['bg'],
                             foreground=self.colors['text'],
                             font=self.fonts['normal'])

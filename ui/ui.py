import tkinter as tk
from ttkbootstrap import Style, ttk
from functools import partial

class ApplicationUI:
    def __init__(self, master, style):
        self.master = master
        self.style = style
        self.font_size = 20
        self._init_dpi()
        self._init_window()
        
        # Initialize tabs lazily
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill='both', expand=True)
        
        # Schedule tab creation for after window is shown
        self.master.after(100, self._init_tabs)

    def _init_dpi(self):
        """Initialize DPI settings"""
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass  # Silently fail on non-Windows platforms

    def _init_window(self):
        """Initialize main window settings"""
        self.master.title("Excel Preset Loader")
        self.master.geometry("800x600")

    def _init_tabs(self):
        """Lazy initialization of tabs"""
        # Import tab classes only when needed
        from .settings_ui import SettingsTab
        from .pruefung_ui import PruefungTab
        
        self.pruefung_tab = PruefungTab(self.notebook, self.style)
        self.settings_tab = SettingsTab(self.notebook, self.style)

        self.notebook.add(self.pruefung_tab.frame, text='Excel Füllung')
        self.notebook.add(self.settings_tab.frame, text='Einstellungen')

        if hasattr(self.settings_tab, 'font_size_scale'):
            self.settings_tab.font_size_scale.config(
                command=self.change_font_size
            )

        self.change_font_size(None)
        self.adjust_window_size()

    def change_font_size(self, event):
        """Change font size with lazy style reapplication"""
        if not hasattr(self.settings_tab, 'font_size_scale'):
            return

        self.font_size = int(float(self.settings_tab.font_size_scale.get()))
        
        # Lazy import of styles
        from .styles import reapply_styles
        reapply_styles(self.style, self.font_size)
        
        # Update widgets only if they exist
        if hasattr(self, 'pruefung_tab'):
            self.pruefung_tab.update_widgets(self.font_size)
        if hasattr(self, 'settings_tab'):
            self.settings_tab.update_widgets(self.font_size)
        
        self.adjust_window_size()

    def adjust_window_size(self):
        """Adjust window size efficiently"""
        self.master.update_idletasks()
        self.master.geometry(
            f"{self.master.winfo_reqwidth()}x{self.master.winfo_reqheight()}"
        )
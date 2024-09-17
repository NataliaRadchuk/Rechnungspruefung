import tkinter as tk
from ttkbootstrap import Style, ttk
from .settings_ui import SettingsTab
from .pruefung_ui import PruefungTab
from .styles import apply_styles, reapply_styles

class ApplicationUI:
    def __init__(self, master, style, logger):
        self.master = master
        self.style = style
        self.font_size = 20  # Default font size
        self.logger = logger

        self.set_dpi_awareness()

        self.master.title("Excel Preset Loader")
        self.master.geometry("800x600")

        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill='both', expand=True)

        # Initialisieren Sie die Tabs mit dem Logger
        self.pruefung_tab = PruefungTab(self.notebook, style, self.logger)
        self.settings_tab = SettingsTab(self.notebook, style, self.logger)  # Fügen Sie den Logger hier hinzu

        self.notebook.add(self.pruefung_tab.frame, text='Excel Füllung')
        self.notebook.add(self.settings_tab.frame, text='Einstellungen')

        if hasattr(self.settings_tab, 'font_size_scale'):
            self.settings_tab.font_size_scale.config(command=self.change_font_size)
        else:
            self.logger.info("font_size_scale nicht in settings_tab gefunden")

        self.change_font_size(None)
        self.adjust_window_size()

    def set_dpi_awareness(self):
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception as e:
            self.logger.info(f"Could not set DPI awareness: {e}")

    def change_font_size(self, event):
        if hasattr(self.settings_tab, 'font_size_scale'):
            self.font_size = int(float(self.settings_tab.font_size_scale.get()))
            reapply_styles(self.style, self.font_size)
            self.pruefung_tab.update_widgets(self.font_size)
            self.settings_tab.update_widgets(self.font_size)
            self.adjust_window_size()
        else:
            self.logger.info("font_size_scale nicht in settings_tab gefunden")

    def adjust_window_size(self):
        self.master.update_idletasks()
        width = self.master.winfo_reqwidth()
        height = self.master.winfo_reqheight()
        self.master.geometry(f"{width}x{height}")
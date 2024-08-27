import tkinter as tk
from ttkbootstrap import Style, ttk
from .settings_ui import SettingsTab
from .pruefung_ui import PruefungTab
from .styles import apply_styles, reapply_styles

class ApplicationUI:
    def __init__(self, master, style):
        self.master = master
        self.style = style
        self.font_size = 20  # Default font size

        self.set_dpi_awareness()

        self.master.title("Excel Preset Loader")

        # Set a reasonable default window size
        self.master.geometry("800x600")

        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill='both', expand=True)

        self.pruefung_tab = PruefungTab(self.notebook, style)
        self.settings_tab = SettingsTab(self.notebook, style)

        self.notebook.add(self.pruefung_tab.frame, text='Excel Füllung')
        self.notebook.add(self.settings_tab.frame, text='Einstellungen')

        # Stellen Sie sicher, dass settings_tab initialisiert ist, bevor Sie es verwenden
        if hasattr(self.settings_tab, 'font_size_scale'):
            self.settings_tab.font_size_scale.config(command=self.change_font_size)
        else:
            print("Warnung: font_size_scale nicht in settings_tab gefunden")

        self.change_font_size(None)
        self.adjust_window_size()

    def set_dpi_awareness(self):
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception as e:
            print(f"Could not set DPI awareness: {e}")

    def change_font_size(self, event):
        if hasattr(self.settings_tab, 'font_size_scale'):
            self.font_size = int(float(self.settings_tab.font_size_scale.get()))
            reapply_styles(self.style, self.font_size)
            self.pruefung_tab.update_widgets(self.font_size)
            self.settings_tab.update_widgets(self.font_size)
            self.adjust_window_size()
        else:
            print("Warnung: font_size_scale nicht in settings_tab gefunden")

    def adjust_window_size(self):
        self.master.update_idletasks()
        width = self.master.winfo_reqwidth()
        height = self.master.winfo_reqheight()
        self.master.geometry(f"{width}x{height}")
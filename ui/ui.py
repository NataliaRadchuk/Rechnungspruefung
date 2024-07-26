import tkinter as tk
from ttkbootstrap import Style, ttk
from .merge_columns_ui import MergeColumnsTab
from .settings_ui import SettingsTab
from .fill_preset_ui import FillPresetTab
from .name_list_ui import NameListTab
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

        self.fill_preset_tab = FillPresetTab(self.notebook, style)
        self.name_list_tab = NameListTab(self.notebook, style)
        self.merge_tab = MergeColumnsTab(self.notebook, style)
        self.settings_tab = SettingsTab(self.notebook, style)

        self.notebook.add(self.fill_preset_tab.frame, text='Fill Preset')
        self.notebook.add(self.name_list_tab.frame, text='Name list')
        self.notebook.add(self.merge_tab.frame, text='Merge Columns')
        self.notebook.add(self.settings_tab.frame, text='Settings')

        self.settings_tab.font_size_scale.config(command=self.change_font_size)  # Bind change_font_size to scale

        self.change_font_size(None)

    def set_dpi_awareness(self):
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception as e:
            print(f"Could not set DPI awareness: {e}")

    def change_font_size(self, event):
        self.font_size = int(float(self.settings_tab.font_size_scale.get()))
        reapply_styles(self.style, self.font_size)
        self.fill_preset_tab.update_widgets(self.font_size)  # Update widgets to apply new font size
        self.merge_tab.create_widgets()  # Recreate widgets to apply new font size
        self.settings_tab.update_widgets(self.font_size)  # Update widgets to apply new font size

import tkinter as tk
from ttkbootstrap import ttk
from .styles import reapply_styles  # Import the reapply_styles function

class SettingsTab:
    def __init__(self, notebook, style, logger):
        self.style = style
        self.font_size = 14  # Default font size
        self.frame = ttk.Frame(notebook, padding="10")
        notebook.add(self.frame, text="Settings")
        self.logger = logger
        self.create_widgets()

    def create_widgets(self):
        # Configure column weight for resizing
        self.frame.columnconfigure(1, weight=1)

        self.font_large = f"Helvetica {self.font_size}"
        
        # Theme selection
        ttk.Label(self.frame, text="Farbschema:", font=self.font_large).grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        self.themes = self.style.theme_names()
        self.theme_combobox = ttk.Combobox(self.frame, values=self.themes, state="readonly", font=self.font_large)
        self.theme_combobox.set(self.style.theme_use())
        self.theme_combobox.grid(row=0, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        self.theme_combobox.bind("<<ComboboxSelected>>", self.change_theme)
        self.theme_combobox.bind('<Button-1>', self.adjust_combobox_font)

        # Font size adjustment
        ttk.Label(self.frame, text="Textgröße:", font=self.font_large).grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        self.font_size_scale = ttk.Scale(self.frame, from_=10, to=24, orient="horizontal", command=self.change_font_size)
        self.font_size_scale.set(self.font_size)
        self.font_size_scale.grid(row=2, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))

    def change_theme(self, event):
        selected_theme = self.theme_combobox.get()
        self.style.theme_use(selected_theme)
        # Reapply the styles after changing the theme
        reapply_styles(self.style, self.font_size)
        # Re-adjust the combobox font after theme change
        self.adjust_combobox_font(None)

    def change_font_size(self, event):
        self.font_size = int(float(self.font_size_scale.get()))
        reapply_styles(self.style, self.font_size)
        self.update_widgets(self.font_size)  # Update widgets to apply new font size

    def update_widgets(self, font_size):
        self.font_size = font_size
        self.font_large = f"Helvetica {self.font_size}"
        for widget in self.frame.winfo_children():
            if isinstance(widget, ttk.Label):
                widget.config(font=self.font_large)
            elif isinstance(widget, ttk.Combobox):
                widget.config(font=self.font_large)
                self.adjust_combobox_font(None)

    def adjust_combobox_font(self, event):
        # Adjust the listbox font size
        self._set_listbox_font(self.theme_combobox)

    def _set_listbox_font(self, combobox):
        combobox.update()  # Ensure the widget is updated
        try:
            listbox_id = combobox.tk.call("ttk::combobox::PopdownWindow", combobox) + ".f.l"
            combobox.tk.call(listbox_id, "configure", "-font", f"Helvetica {self.font_size}")
        except tk.TclError:
            pass

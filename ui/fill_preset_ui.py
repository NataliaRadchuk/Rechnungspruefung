import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import ttk
from file_handler import FileHandler
import os

class FillPresetTab:
    def __init__(self, notebook, style):
        self.style = style
        self.font_size = 14
        self.frame = ttk.Frame(notebook, padding="10")
        notebook.add(self.frame, text="Fill Preset")
        self.create_widgets()
        self.file_ops = FileHandler()

    def create_widgets(self):
        self.frame.columnconfigure(1, weight=1)
        self.font_large = f"Helvetica {self.font_size}"

        ttk.Label(self.frame, text="Input File:", font=self.font_large).grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        self.input_entry = ttk.Entry(self.frame, font=self.font_large, width=50)
        self.input_entry.grid(row=0, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        self.input_button = ttk.Button(self.frame, text="Browse", command=self.select_input_file, style="TButton")
        self.input_button.grid(row=0, column=2, padx=10, pady=10)

        ttk.Label(self.frame, text="Output Directory:", font=self.font_large).grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        self.output_dir_entry = ttk.Entry(self.frame, font=self.font_large, width=50)
        self.output_dir_entry.grid(row=1, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        self.output_button = ttk.Button(self.frame, text="Browse", command=self.select_output_dir, style="TButton")
        self.output_button.grid(row=1, column=2, padx=10, pady=10)

        ttk.Label(self.frame, text="Output Filename:", font=self.font_large).grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        self.output_filename_entry = ttk.Entry(self.frame, font=self.font_large, width=50)
        self.output_filename_entry.grid(row=2, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        ttk.Label(self.frame, text=".xlsm", font=self.font_large).grid(row=2, column=2, padx=10, pady=10, sticky=tk.W)

        self.process_button = ttk.Button(self.frame, text="Process", command=self.start_process, style="TButton")
        self.process_button.grid(row=3, column=0, columnspan=3, pady=20)

        self.status_label = ttk.Label(self.frame, text="", font=self.font_large)
        self.status_label.grid(row=4, column=0, columnspan=3, pady=10)

    def select_input_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel and CSV files", "*.xlsx;*.xls;*.xlsm;*.csv")])
        self.input_entry.delete(0, tk.END)
        self.input_entry.insert(0, file_path)

    def select_output_dir(self):
        dir_path = filedialog.askdirectory()
        self.output_dir_entry.delete(0, tk.END)
        self.output_dir_entry.insert(0, dir_path)

    def start_process(self):
        input_file = self.input_entry.get()
        output_dir = self.output_dir_entry.get()
        output_filename = self.output_filename_entry.get()

        if not output_filename:
            messagebox.showerror("Input Error", "Please enter an output filename.")
            return

        if not output_filename.endswith('.xlsm'):
            output_filename += '.xlsm'

        current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        preset_file = os.path.join(current_dir, "templates", "template.xlsm")

        if not input_file or not output_dir or not output_filename:
            messagebox.showerror("Input Error", "Please select all required files, output directory, and provide an output filename.")
            return

        if not os.path.exists(preset_file):
            messagebox.showerror("File Error", f"Template file not found: {preset_file}")
            return

        try:
            trimmed_df = self.file_ops.copy_and_trim_file(input_file)
            self.file_ops.append_trimmed_to_existing(trimmed_df, preset_file, output_dir, output_filename)
            self.status_label.config(text="Process completed successfully.", foreground="green")
        except Exception as e:
            messagebox.showerror("Processing Error", str(e))
            self.status_label.config(text="An error occurred during processing.", foreground="red")

    def update_widgets(self, font_size):
        self.font_size = font_size
        self.font_large = f"Helvetica {self.font_size}"
        print(f"Updating widgets to font size: {self.font_size}")
        for widget in self.frame.winfo_children():
            print(f"Updating {widget} to font size {self.font_size}")
            if isinstance(widget, ttk.Label):
                widget.config(font=self.font_large)
            elif isinstance(widget, ttk.Entry):
                widget.config(font=self.font_large)
            elif isinstance(widget, ttk.Button):
                style_name = widget.cget("style") or "TButton"
                self.style.configure(style_name, font=self.font_large)
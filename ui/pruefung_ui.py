import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from ttkbootstrap import ttk
from pruefung_handler import PruefungHandler
import os
import threading
import logging
from .gui_log_handler import GUILogHandler

class PruefungTab:
    def __init__(self, notebook, style):
        self.style = style
        self.font_size = 14
        self.frame = ttk.Frame(notebook, padding="10")
        notebook.add(self.frame, text="Prüfung")
        self.create_widgets()
        self.file_ops = PruefungHandler()
        self.setup_logger()

    def setup_logger(self):
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        
        # Entfernen Sie alle bestehenden Handler
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        
        # Fügen Sie den benutzerdefinierten GUI-Handler hinzu
        gui_handler = GUILogHandler(self.update_log)
        gui_handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        gui_handler.setFormatter(formatter)
        logger.addHandler(gui_handler)

    def create_widgets(self):
        self.frame.columnconfigure(1, weight=1)
        self.font_large = f"Helvetica {self.font_size}"

        # Input File
        ttk.Label(self.frame, text="Report:", font=self.font_large).grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        self.input_entry = ttk.Entry(self.frame, font=self.font_large, width=50)
        self.input_entry.grid(row=0, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        self.input_button = ttk.Button(self.frame, text="Suchen", command=self.select_input_file)
        self.input_button.grid(row=0, column=2, padx=10, pady=10)

        # Namensliste File
        ttk.Label(self.frame, text="Namensliste:", font=self.font_large).grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        self.namensliste_entry = ttk.Entry(self.frame, font=self.font_large, width=50)
        self.namensliste_entry.grid(row=1, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        self.namensliste_button = ttk.Button(self.frame, text="Suchen", command=self.select_namensliste_file)
        self.namensliste_button.grid(row=1, column=2, padx=10, pady=10)

        # Checkbox
        self.checkbox_var = tk.BooleanVar()
        self.checkbox = ttk.Checkbutton(self.frame, text="Namensliste NICHT aus Tableau", variable=self.checkbox_var)
        self.checkbox.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky=tk.W)

        # Output Directory
        ttk.Label(self.frame, text="Speicherort:", font=self.font_large).grid(row=3, column=0, padx=10, pady=10, sticky=tk.W)
        self.output_dir_entry = ttk.Entry(self.frame, font=self.font_large, width=50)
        self.output_dir_entry.grid(row=3, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        self.output_button = ttk.Button(self.frame, text="Suchen", command=self.select_output_dir)
        self.output_button.grid(row=3, column=2, padx=10, pady=10)

        # Output Filename
        ttk.Label(self.frame, text="Speichername:", font=self.font_large).grid(row=4, column=0, padx=10, pady=10, sticky=tk.W)
        self.output_filename_entry = ttk.Entry(self.frame, font=self.font_large, width=50)
        self.output_filename_entry.grid(row=4, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        ttk.Label(self.frame, text=".xlsm", font=self.font_large).grid(row=4, column=2, padx=10, pady=10, sticky=tk.W)

        # Start Button
        self.process_button = ttk.Button(self.frame, text="Start", command=self.start_process)
        self.process_button.grid(row=5, column=0, columnspan=3, pady=20)

        # Aufgabe 2: Log-Textfeld
        self.log_text = scrolledtext.ScrolledText(self.frame, height=10, width=70, font=self.font_large)
        self.log_text.grid(row=6, column=0, columnspan=3, pady=10)
        self.log_text.config(state=tk.DISABLED)  # Macht das Textfeld schreibgeschützt

    def select_input_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel and CSV files", "*.xlsx;*.xls;*.xlsm;*.csv")])
        self.input_entry.delete(0, tk.END)
        self.input_entry.insert(0, file_path)

    def select_namensliste_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.namensliste_entry.delete(0, tk.END)
        self.namensliste_entry.insert(0, file_path)

    def select_output_dir(self):
        dir_path = filedialog.askdirectory()
        self.output_dir_entry.delete(0, tk.END)
        self.output_dir_entry.insert(0, dir_path)

    def start_process(self):
        # Aufgabe 1: Deaktiviere den Start-Button
        self.process_button.config(state=tk.DISABLED)

        # Lösche vorherige Log-Einträge
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        input_file = self.input_entry.get()
        namensliste_file = self.namensliste_entry.get()
        output_dir = self.output_dir_entry.get()
        output_filename = self.output_filename_entry.get()
        option_aktiviert = self.checkbox_var.get()

        if not output_filename:
            messagebox.showerror("Input Error", "Please enter an output filename.")
            self.process_button.config(state=tk.NORMAL)  # Reaktiviere den Button bei Fehler
            return

        if not output_filename.endswith('.xlsm'):
            output_filename += '.xlsm'

        current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        preset_file = os.path.join(current_dir, "templates", "template.xlsm")

        if not input_file or not namensliste_file or not output_dir or not output_filename:
            messagebox.showerror("Input Error", "Please select all required files, output directory, and provide an output filename.")
            self.process_button.config(state=tk.NORMAL)  # Reaktiviere den Button bei Fehler
            return

        if not os.path.exists(preset_file):
            messagebox.showerror("File Error", f"Template file not found: {preset_file}")
            self.process_button.config(state=tk.NORMAL)  # Reaktiviere den Button bei Fehler
            return

        # Starte den Verarbeitungsprozess in einem separaten Thread
        threading.Thread(target=self.process_files, args=(input_file, namensliste_file, option_aktiviert, output_dir, output_filename)).start()

    def process_files(self, input_file, namensliste_file, option_aktiviert, output_dir, output_filename):
        try:
            success, message = self.file_ops.process_files(input_file, namensliste_file, option_aktiviert, output_dir, output_filename)
            if success:
                self.update_log("Process completed successfully.")
            else:
                self.update_log(f"An error occurred: {message}")
        except Exception as e:
            self.update_log(f"An error occurred: {str(e)}")
        finally:
            # Aufgabe 1: Reaktiviere den Start-Button nach Abschluss der Verarbeitung
            self.frame.after(0, lambda: self.process_button.config(state=tk.NORMAL))

    def update_log(self, message):
        # Aufgabe 2: Aktualisiere das Log-Textfeld
        def update():
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        self.frame.after(0, update)

    def update_widgets(self, font_size):
        self.font_size = font_size
        self.font_large = f"Helvetica {self.font_size}"
        for widget in self.frame.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.configure(style='TButton')
            elif isinstance(widget, ttk.Entry) or isinstance(widget, ttk.Label):
                widget.configure(font=self.font_large)
            elif isinstance(widget, ttk.Checkbutton):
                widget.configure(style='TCheckbutton')
            elif isinstance(widget, scrolledtext.ScrolledText):
                widget.configure(font=self.font_large)

        self.style.configure('TButton', font=self.font_large)
        self.style.configure('TCheckbutton', font=self.font_large)
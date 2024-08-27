# import tkinter as tk
# from ttkbootstrap import ttk
# from tkinter import filedialog
# from merge_columns_handler import MergeColumnsHandler


# class MergeColumnsTab:
#     def __init__(self, notebook, style):
#         self.file_ops = MergeColumnsHandler()
#         self.style = style
#         self.frame = ttk.Frame(notebook, padding="10")
#         notebook.add(self.frame, text="Merge Columns")
#         self.create_widgets()

#     def create_widgets(self):
#         ttk.Label(self.frame, text="Input File:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
#         self.input_entry = ttk.Entry(self.frame, width=50)
#         self.input_entry.grid(row=0, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
#         ttk.Button(self.frame, text="Browse", command=self.select_input_file).grid(row=0, column=2, padx=10, pady=10)

#         self.columns_frame = ttk.LabelFrame(self.frame, text="Found Columns", padding="10")
#         self.columns_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky=(tk.W, tk.E))
        
#         self.columns_text = tk.StringVar()
#         ttk.Label(self.columns_frame, textvariable=self.columns_text, wraplength=400).grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=(tk.W, tk.E))

#         ttk.Label(self.frame, text="Output File:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
#         self.output_entry = ttk.Entry(self.frame, width=50)
#         self.output_entry.grid(row=2, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
#         ttk.Button(self.frame, text="Browse", command=self.select_output_file).grid(row=2, column=2, padx=10, pady=10)

#         ttk.Label(self.frame, text="Columns (comma-separated):").grid(row=3, column=0, padx=10, pady=10, sticky=tk.W)
#         self.columns_entry = ttk.Entry(self.frame, width=50)
#         self.columns_entry.grid(row=3, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))

#         ttk.Button(self.frame, text="Merge", command=self.start_merge).grid(row=4, column=0, columnspan=3, pady=20)
    
#     def select_input_file(self):
#         file_path = filedialog.askopenfilename(filetypes=[("Excel and CSV files", "*.xlsx;*.xls;*.csv")])
#         self.input_entry.delete(0, tk.END)
#         self.input_entry.insert(0, file_path)
#         self.display_columns(file_path)

#     def display_columns(self, file_path):
#         columns = self.file_ops.get_columns(file_path)
#         self.columns_text.set(', '.join(columns))

#     def select_output_file(self):
#         file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
#         self.output_entry.delete(0, tk.END)
#         self.output_entry.insert(0, file_path)

#     def start_merge(self):
#         input_file = self.input_entry.get()
#         output_file = self.output_entry.get()
#         columns = self.columns_entry.get().split(',')
#         self.file_ops.merge_columns(input_file, output_file, columns)

#     def start_merge(self):
#         input_file = self.input_entry.get()
#         output_file = self.output_entry.get()
#         self.file_ops.merge_columns(input_file, output_file)
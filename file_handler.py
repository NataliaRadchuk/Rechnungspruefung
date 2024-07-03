import pandas as pd
import os
import logging
from openpyxl import load_workbook

class FileHandler:
    def __init__(self):
        # Set up logging configuration
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def read_file(self, file_path):
        file_extension = os.path.splitext(file_path)[1].lower()
        try:
            if file_extension in ['.xlsx', '.xls', '.xlsm']:
                df = pd.read_excel(file_path)
            elif file_extension == '.csv':
                with open(file_path, 'r', encoding='utf-8') as f:
                    first_line = f.readline()
                delimiter = ',' if first_line.count(',') > first_line.count(';') else ';'
                df = pd.read_csv(file_path, delimiter=delimiter, quoting=1, on_bad_lines='skip')
            else:
                raise ValueError("Unsupported file extension")
            return df
        except Exception as e:
            logging.error(f"Error reading file: {e}")
            return None

    def copy_and_trim_file(self, input_file):
        try:
            df = self.read_file(input_file)
            if df is None:
                raise ValueError("DataFrame is empty or file could not be read.")
            
            # Remove the first column and the first row
            trimmed_df = df.iloc[1:, 1:]
            return trimmed_df
        except Exception as e:
            logging.error(f"An error occurred: {e}")
            return None

    def append_trimmed_to_existing(self, trimmed_df, preset_file, output_dir):
        try:
            if trimmed_df is None:
                raise ValueError("Trimmed DataFrame is empty.")

            # Load the preset .xlsm file and the specific worksheet "Prüf"
            wb = load_workbook(preset_file, keep_vba=True)
            if "Prüf" not in wb.sheetnames:
                raise ValueError('Worksheet "Prüf" not found in the preset file.')

            ws = wb["Prüf"]

            # Find the first empty row in the worksheet
            start_row = next((row for row, cells in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True), start=2) if all(cell is None for cell in cells)), ws.max_row + 1)

            # Append trimmed DataFrame to the existing worksheet at the first column, starting from the first empty row
            for r_idx, row in enumerate(trimmed_df.itertuples(index=False), start=start_row):
                for c_idx, value in enumerate(row):
                    cell = ws.cell(row=r_idx, column=c_idx + 1, value=value)
                    # Ensure the value is saved as a number if it is numeric
                    if pd.api.types.is_numeric_dtype(type(value)):
                        cell.value = float(value)
                        cell.number_format = 'General'  # or other format as needed
                    else:
                        cell.value = value  # Ensure text is copied as is

            # Find the new last row with data
            last_row = ws.max_row

            # Move the summary row down one line for visual separation
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=last_row + 1, column=col_idx, value=ws.cell(row=last_row, column=col_idx).value)
                ws.cell(row=last_row, column=col_idx).value = None

            # Update the moved summary row to remove the first three "Gesamtwert" entries
            ws.cell(row=last_row + 1, column=1).value = ""
            ws.cell(row=last_row + 1, column=2).value = ""
            ws.cell(row=last_row + 1, column=3).value = ""
            ws.cell(row=last_row + 1, column=4).value = "Gesamtwert"

            # Create the output file name
            base_name = os.path.splitext(os.path.basename(preset_file))[0]
            output_file = os.path.join(output_dir, f"{base_name}_filled.xlsm")

            # Save the updated workbook
            wb.save(output_file)
            logging.info(f"Successfully appended trimmed data to {output_file} in worksheet 'Prüf'")
        except Exception as e:
            logging.error(f"An error occurred: {e}")

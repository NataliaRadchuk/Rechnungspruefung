import pandas as pd
import os
import logging
from openpyxl import load_workbook
import chardet  # Import chardet for encoding detection

class FileHandler:
    def __init__(self):
        # Set up logging configuration
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def detect_encoding(self, file_path):
        # Detect the file encoding using chardet
        with open(file_path, 'rb') as f:
            result = chardet.detect(f.read(100000))
        return result['encoding']

    def read_file(self, file_path):
        file_extension = os.path.splitext(file_path)[1].lower()
        try:
            if file_extension in ['.xlsx', '.xls', '.xlsm']:
                # Read Excel files using pandas
                df = pd.read_excel(file_path)
            elif file_extension == '.csv':
                # Detect the file encoding
                encoding = self.detect_encoding(file_path)
                logging.info(f"Detected file encoding: {encoding}")

                try:
                    # Try reading the CSV file with the detected encoding
                    df = pd.read_csv(file_path, delimiter='\t', encoding=encoding, quoting=1, on_bad_lines='skip')
                except UnicodeDecodeError:
                    # If detected encoding fails, try 'ISO-8859-1'
                    encoding = 'ISO-8859-1'
                    logging.info(f"Falling back to encoding: {encoding}")
                    df = pd.read_csv(file_path, delimiter='\t', encoding=encoding, quoting=1, on_bad_lines='skip')
                except Exception as e:
                    logging.error(f"Error reading file with detected encoding: {e}")
                    return None
            else:
                raise ValueError("Unsupported file extension")
            return df
        except Exception as e:
            logging.error(f"Error reading file: {e}")
            return None

    def copy_and_trim_file(self, input_file):
        try:
            df = self.read_file(input_file)
            if df is None or df.empty:
                raise ValueError("DataFrame is empty or file could not be read.")
            
            # Remove the first column and the first row
            logging.info(f"Original DataFrame:\n{df.head()}")  # Debugging output
            trimmed_df = df.iloc[1:, 1:]
            logging.info(f"Trimmed DataFrame:\n{trimmed_df.head()}")  # Debugging output
            return trimmed_df
        except Exception as e:
            logging.error(f"An error occurred: {e}")
            return None

    def append_trimmed_to_existing(self, trimmed_df, preset_file, output_dir):
        try:
            if trimmed_df is None or trimmed_df.empty:
                raise ValueError("Trimmed DataFrame is empty.")

            # Load the preset .xlsm file and the specific worksheet "Prüf"
            wb = load_workbook(preset_file, keep_vba=True)
            if "Prüf" not in wb.sheetnames:
                raise ValueError('Worksheet "Prüf" not found in the preset file.')

            ws = wb["Prüf"]

            # Find the first empty row in the worksheet
            start_row = next((row for row, cells in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True), start=2) if all(cell is None for cell in cells)), ws.max_row + 1)

            # Define the columns that should be treated as numbers
            number_columns = {
                'A': 1,  # "Reservierungsnummer"
                'F': 6,  # "EK FRST (Gesamt B)"
                'G': 7,  # "EK LOGIS (Gesamt B)"
                'H': 8,  # "EK (Gesamt B)"
                'I': 9,  # "Komm. (Gesamt B)"
                'J': 10  # "VK (Gesamt B)"
            }

            # Append trimmed DataFrame to the existing worksheet at the first column, starting from the first empty row
            for r_idx, row in enumerate(trimmed_df.itertuples(index=False), start=start_row):
                for c_idx, value in enumerate(row):
                    # Check if the column should be a number
                    if (c_idx + 1) in number_columns.values() and pd.api.types.is_numeric_dtype(type(value)):
                        cell = ws.cell(row=r_idx, column=c_idx + 1, value=float(value))
                        cell.number_format = 'General'  # or other format as needed
                    else:
                        cell = ws.cell(row=r_idx, column=c_idx + 1, value=value)  # Ensure text is copied as is

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

# Instantiate and use the FileHandler class
file_handler = FileHandler()
input_file = 'path_to_input_file.csv'  # Update with actual path
preset_file = 'path_to_preset_file.xlsm'  # Update with actual path
output_dir = 'path_to_output_directory'  # Update with actual path

trimmed_df = file_handler.copy_and_trim_file(input_file)
if trimmed_df is not None:
    file_handler.append_trimmed_to_existing(trimmed_df, preset_file, output_dir)

import pandas as pd
import os
import logging

class MergeColumnsHandler:
    def __init__(self):
        # Set up logging configuration
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def get_columns(self, file_path):
        file_extension = os.path.splitext(file_path)[1].lower()
        try:
            if file_extension in ['.xlsx', '.xls']:
                df = pd.read_excel(file_path)
            elif file_extension == '.csv':
                with open(file_path, 'r', encoding='utf-8') as f:
                    first_line = f.readline()
                delimiter = ',' if first_line.count(',') > first_line.count(';') else ';'
                df = pd.read_csv(file_path, delimiter=delimiter, quoting=1, on_bad_lines='skip')
            else:
                raise ValueError("Unsupported file extension")
            return df.columns.tolist()
        except Exception as e:
            logging.error(f"Error reading file: {e}")
            return [f"Error reading file: {e}"]

    def merge_columns(self, input_file, output_file, columns):
        file_extension = os.path.splitext(input_file)[1].lower()
        try:
            if file_extension in ['.xlsx', '.xls']:
                df = pd.read_excel(input_file)
            elif file_extension == '.csv':
                with open(input_file, 'r', encoding='utf-8') as f:
                    first_line = f.readline()
                delimiter = ',' if first_line.count(',') > first_line.count(';') else ';'
                df = pd.read_csv(input_file, delimiter=delimiter, quoting=1, on_bad_lines='skip')
            else:
                raise ValueError("Unsupported file extension")

            columns = [col.strip() for col in columns]
            actual_columns = [col.strip() for col in df.columns]
            missing_columns = [col for col in columns if col not in actual_columns]
            if missing_columns:
                raise ValueError(f"Columns not found in the input file: {', '.join(missing_columns)}")

            merged_df = df[columns]
            merged_df.to_excel(output_file, index=False, engine='openpyxl')
            logging.info(f"Successfully merged columns and saved to {output_file}")
        except Exception as e:
            logging.error(f"An error occurred: {e}")


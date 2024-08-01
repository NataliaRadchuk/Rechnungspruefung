import pandas as pd
import os
import logging
import chardet
import locale
from file_handler import FileHandler 

class NameListHandler:
    def __init__(self):
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        locale.setlocale(locale.LC_NUMERIC, 'de_DE.UTF-8')  # Setze das deutsche Zahlenformat
        self.file_ops = FileHandler()

    def prepare_table(self, name_list_file):
        pandas_file = self.file_ops.copy_and_trim_file(name_list_file)
        logging.info(f"We are here and everythinig seems to work ")
        logging.info(f"Trimmed DataFrame:\n{pandas_file}")
        # TODO Logik
        return pandas_file
    
    def append_into_preset(self, trimmed_df, preset_file, output_dir, output_filename):
        logging.info(f"in append to preset")
        return None
    
name_list_handler = NameListHandler()
input_file = 'path_to_input_file.csv'
preset_file = 'path_to_preset_file.xlsm'
output_dir = 'path_to_output_directory'

trimmed_df = name_list_handler.prepare_table(input_file)
if trimmed_df is not None:
    name_list_handler.append_into_preset(trimmed_df, preset_file, output_dir)
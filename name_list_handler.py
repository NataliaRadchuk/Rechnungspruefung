import pandas as pd
import os
import logging
import chardet
from datetime import datetime
import locale
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from file_handler import FileHandler 

class NameListHandler:
    def __init__(self):
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        locale.setlocale(locale.LC_NUMERIC, 'de_DE.UTF-8')  # Setze das deutsche Zahlenformat
        self.file_ops = FileHandler()

    def prepare_table(self, name_list_file):
        pandas_file = self.file_ops.copy_and_trim_file(name_list_file)
    
        if pandas_file is None or pandas_file.empty:
            logging.error("Failed to load or process the file, or the file is empty.")
            return None

        logging.info("File loaded and trimmed successfully.")
        logging.info(f"Trimmed DataFrame:\n{pandas_file}")

        # Aufgabe 1: Letzte Spalte löschen, wenn sie nur leere Werte enthält
        if pandas_file.iloc[:, -1].astype(str).str.strip().eq('').all():
            pandas_file = pandas_file.iloc[:, :-1]

        # Aufgabe 2: Spalte K (Roomnights) löschen, falls vorhanden
        if pandas_file.shape[1] >= 11:
            pandas_file = pandas_file.iloc[:, :10]

        # Aufgabe 4: Spalte A löschen (falls noch nicht geschehen)
        if pandas_file.shape[1] > 1:
            pandas_file = pandas_file.iloc[:, 1:]

        # Aufgabe 3: Inhalt von Spalte I nach J verschieben, I leer lassen
        if pandas_file.shape[1] >= 9:
            # Füge eine neue leere Spalte am Ende hinzu
            pandas_file['New'] = ''
            # Verschiebe den Inhalt von I nach J
            pandas_file.iloc[:, -1] = pandas_file.iloc[:, -2]
            # Leere Spalte I
            pandas_file.iloc[:, -2] = ''
        
        # Stelle sicher, dass wir genau 10 Spalten haben
        if pandas_file.shape[1] < 10:
            for i in range(pandas_file.shape[1], 10):
                pandas_file[f'New_{i}'] = ''
        elif pandas_file.shape[1] > 10:
            pandas_file = pandas_file.iloc[:, :10]

        # Spalten neu indizieren
        pandas_file.columns = [chr(65 + i) for i in range(10)]

        # Aufgabe 7: Zeilen basierend auf Werten in Spalte H duplizieren
        new_rows = []
        for index, row in pandas_file.iterrows():
            h_value = row['H']
            try:
                h_value = int(float(h_value))  # Versuche, den Wert in eine Ganzzahl umzuwandeln
                if h_value > 0:
                    for i in range(h_value):
                        new_row = row.copy()
                        new_row['H'] = 1
                        new_rows.append(new_row)
                else:
                    new_rows.append(row)  # Behalte die Originalzeile für Werte <= 0
            except ValueError:
                new_rows.append(row)  # Behalte die Originalzeile, wenn H keine Zahl ist

        pandas_file = pd.DataFrame(new_rows, columns=pandas_file.columns)

        # Aufgabe 2: Kopiere und füge den gesamten Inhalt ein, füge "-F" in Spalte F hinzu
        copied_df = pandas_file.copy()
        copied_df['F'] = copied_df['F'].astype(str) + '-F'
        pandas_file = pd.concat([pandas_file, copied_df], ignore_index=True)

        # Alphabetisch sortieren nach Spalte F
        pandas_file = pandas_file.sort_values(by='C')

        logging.info(f"Prepared DataFrame:\n{pandas_file}")
        return pandas_file
    
    def append_into_preset(self, trimmed_df, workbook):
        logging.info("In append to preset")

        # Wähle den Tab "SR Lt. AM"
        sheet = workbook['SR Lt. AM']

        # Füge trimmed_df ab Zeile 12, Spalte A ein
        last_data_row = 11  # Initialisiere mit der Zeile vor den Daten
        for r, row in enumerate(dataframe_to_rows(trimmed_df, index=False, header=False), start=12):
            for c, value in enumerate(row, start=1):
                cell = sheet.cell(row=r, column=c)
                if isinstance(value, (int, float)):
                    cell.value = value
                else:
                    cell.value = str(value)
            # Füge die Formel in Spalte I ein
            sheet.cell(row=r, column=9).value = f'=E{r}-D{r}'
            last_data_row = r  # Aktualisiere die letzte Datenzeile

        # Konvertiere Spalte D und E in Datumsformat
        def convert_to_date(cell):
            if cell.value:
                try:
                    date_value = datetime.strptime(str(cell.value), "%d.%m.%Y")
                    cell.value = date_value
                    cell.number_format = 'DD.MM.YYYY'
                except ValueError:
                    logging.warning(f"Konnte Datum in Zelle {cell.coordinate} nicht konvertieren: {cell.value}")

        for col in [4, 5]:  # 4 für Spalte D, 5 für Spalte E
            for row in range(12, last_data_row + 1):
                convert_to_date(sheet.cell(row=row, column=col))
        
        # Aufgabe 3: Einträge aus Spalte H in die innere Tabelle einfügen
        unique_entries = sorted(set(trimmed_df['F'].astype(str)))
        
        for i, entry in enumerate(unique_entries):
            if i < 6:  # Für die ersten 6 Einträge
                sheet.cell(row=i+3, column=11).value = entry
        
        # Suche nach "Gesamt" in Spalte K ab Zeile 10
        for row in range(10, sheet.max_row + 1):
            if sheet.cell(row=row, column=11).value == "Gesamt":
                # Passe die Formel an, um nur den Bereich mit Daten zu berücksichtigen
                sheet.cell(row=row, column=12).value = f'=MIN(D12:D{last_data_row})-2'
                sheet.cell(row=row, column=12).number_format = 'DD.MM.YYYY'
                break
        
        # Färbe Zellen von Spalte L bis AD ein, wenn sie einen sichtbaren Wert haben
        # fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")  # Hellgrün
        # for row in range(12, last_data_row + 1):
        #     for col in range(12, 30):  # L = 12, AD = 30
        #         cell = sheet.cell(row=row, column=col)
                
        #         # Verwende den angezeigten Wert statt des tatsächlichen Zellwerts
        #         displayed_value = cell.value

        #         # Überprüfe, ob der angezeigte Wert leer oder "Falsch" ist
        #         if displayed_value is not None and displayed_value != "" and displayed_value != "Falsch":
        #             cell.fill = fill
        #         else:
        #             cell.fill = PatternFill(fill_type=None)
        return workbook

    def process_template(self, template_filled, input_namelist, checkset):
        try:
            prepared_table = self.prepare_table(input_namelist)
            if prepared_table is None:
                raise ValueError("Failed to prepare name list table.")
            
            fully_filled_template = self.append_into_preset(prepared_table, template_filled)
            if fully_filled_template is None:
                raise ValueError("Failed to append prepared table into template.")
            
            return fully_filled_template
        except Exception as e:
            logging.error(f"Error in process_template: {e}")
            return None

# Verwendung der NameListHandler-Klasse
if __name__ == "__main__":
    name_list_handler = NameListHandler()
    input_file = 'path_to_input_file.csv'
    template_filled = openpyxl.load_workbook('path_to_template_filled.xlsm', keep_vba=True)
    
    fully_filled_template = name_list_handler.process_template(template_filled, input_file, True)
    if fully_filled_template:
        print("Template processed successfully.")
    else:
        print("Failed to process template.")
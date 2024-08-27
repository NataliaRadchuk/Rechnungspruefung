from file_handler import FileHandler
from name_list_handler import NameListHandler
import os
import logging

class PruefungHandler:
    def __init__(self):
        self.file_handler = FileHandler()
        self.name_list_handler = NameListHandler()
        self.logger = logging.getLogger(__name__)

    def process_files(self, input_report, input_namelist, checkset, output_dir, output_filename):
        try:
            # Schritt 1: Verarbeite den Input-Report mit FileHandler
            logging.info(f"Schritt 1 start")
            template_filled = self.file_handler.process_report(input_report)
            logging.info(f"Schritt 1 Ende und Schritt 2 Anfang")

            # Schritt 2: Verarbeite das gefüllte Template mit NameListHandler
            fully_filled_template = self.name_list_handler.process_template(template_filled, input_namelist, checkset)
            logging.info(f"Schritt 2 zuende und nur noch Speichern")
            # Schritt 3: Speichere das vollständig gefüllte Template
            self.save_template(fully_filled_template, output_dir, output_filename)

            return True, "Verarbeitung erfolgreich abgeschlossen."
        except Exception as e:
            return False, f"Fehler bei der Verarbeitung: {str(e)}"

    def save_template(self, template, output_dir, output_filename):
        # Stelle sicher, dass das Ausgabeverzeichnis existiert
        os.makedirs(output_dir, exist_ok=True)

        # Erstelle den vollständigen Dateipfad
        full_path = os.path.join(output_dir, output_filename)

        # Speichere das Template
        template.save(full_path)
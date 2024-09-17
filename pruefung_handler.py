from file_handler import FileHandler
from name_list_handler import NameListHandler
import os
#import logging


class PruefungHandler:
    def __init__(self, logger):
        self.file_handler = FileHandler(logger)
        self.name_list_handler = NameListHandler(logger)
        #self.logger = logging.getLogger(__name__)
        self.logger = logger


    def process_files(self, input_report, input_namelist, checkset, output_dir, output_filename):
        try:
            # Schritt 1: Verarbeite den Input-Report mit FileHandler
            self.logger.info("Schritt 1 start")
            template_filled = self.file_handler.process_report(input_report)
            self.logger.info("Schritt 1 Ende")

            if not checkset:
                # Schritt 2: Verarbeite das gefüllte Template mit NameListHandler
                self.logger.info("Schritt 2 Anfang")
                fully_filled_template = self.name_list_handler.process_template(template_filled, input_namelist, checkset)
                self.logger.info("Schritt 2 Ende")
            else:
                fully_filled_template = template_filled

            # Schritt 3: Speichere das Template
            self.logger.info("Speichern des Templates")
            self.save_template(fully_filled_template, output_dir, output_filename)
            
            return True, "Verarbeitung erfolgreich abgeschlossen."
        
        except Exception as e:
            self.logger.info(f"Fehler bei der Verarbeitung: {str(e)}", exc_info=True)
            return False, f"Fehler bei der Verarbeitung: {str(e)}"

    def save_template(self, template, output_dir, output_filename):
        # Stelle sicher, dass das Ausgabeverzeichnis existiert
        os.makedirs(output_dir, exist_ok=True)

        # Erstelle den vollständigen Dateipfad
        full_path = os.path.join(output_dir, output_filename)

        # Speichere das Template
        template.save(full_path)
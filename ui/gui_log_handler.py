import logging
from datetime import datetime

class GUILogHandler(logging.Handler):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def emit(self, record):
        try:
            log_entry = self.format(record)
            self.callback(log_entry)
        except Exception:
            self.handleError(record)
    
    def format(self, record):
        # Anpassen des Formats wie gewünscht
        timestamp = datetime.fromtimestamp(record.created).strftime("%H:%M")
        return f"{timestamp} - {record.getMessage()}"
    
def setup_logger(callback):
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    
    # Entfernen aller bestehenden Handler
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # Hinzufügen des neuen GUILogHandler
    gui_handler = GUILogHandler(callback)
    gui_handler.setLevel(logging.INFO)
    logger.addHandler(gui_handler)
    
    return logger
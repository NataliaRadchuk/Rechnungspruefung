import logging

class GUILogHandler(logging.Handler):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def emit(self, record):
        log_entry = self.format(record)
        self.callback(log_entry)
    
    def setup_logger(self):
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        
        gui_handler = GUILogHandler(self.log_queue.put)
        gui_handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', 
                                    datefmt='%d-%m %H:%M')
        gui_handler.setFormatter(formatter)
        logger.addHandler(gui_handler)
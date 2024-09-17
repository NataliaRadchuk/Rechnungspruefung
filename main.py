from ui.ui import ApplicationUI
from ttkbootstrap import Style
from ui.styles import apply_styles
from ui.gui_log_handler import setup_logger

def log_to_gui(log_entry):
    # Diese Funktion wird später verwendet, um Log-Einträge an die GUI zu senden
    print(log_entry)  # Vorerst geben wir es einfach in der Konsole aus

def main():
    style = Style(theme='united')  # Default theme
    root = style.master
    apply_styles(style)

    # Logger einrichten
    logger = setup_logger(log_to_gui)

    app = ApplicationUI(root, style, logger)
    root.mainloop()

    
if __name__ == "__main__":
    main()
from ui.ui import ApplicationUI
from ttkbootstrap import Style
from ui.styles import apply_styles

def main():
    style = Style(theme='cosmo')  # Default theme
    root = style.master
    apply_styles(style)
    app = ApplicationUI(root, style)
    root.mainloop()

if __name__ == "__main__":
    main()
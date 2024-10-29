import tkinter as tk
from ttkbootstrap import Style
import sys
from functools import partial

def lazy_import():
    """Lazy import of UI components only when needed"""
    from ui.ui import ApplicationUI
    from ui.styles import apply_styles
    return ApplicationUI, apply_styles

def main():
    # Create the root window immediately for faster perceived startup
    root = tk.Tk()
    root.withdraw()  # Hide window while loading
    
    # Show a simple loading window
    loading = tk.Toplevel(root)
    loading.title("Loading...")
    loading.geometry("200x50")
    loading.transient(root)
    loading_label = tk.Label(loading, text="Loading application...")
    loading_label.pack(padx=20, pady=10)
    
    # Schedule the actual application loading
    root.after(100, partial(initialize_app, root, loading))
    root.mainloop()

def initialize_app(root, loading):
    try:
        # Lazy import of heavy components
        ApplicationUI, apply_styles = lazy_import()
        
        # Initialize style after imports to prevent double initialization
        style = Style(theme='united')
        style.master = root
        
        # Apply styles
        apply_styles(style)
        
        # Create main application
        app = ApplicationUI(root, style)
        
        # Show main window and destroy loading screen
        root.deiconify()
        loading.destroy()
        
    except Exception as e:
        import traceback
        print(f"Error during initialization: {e}")
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
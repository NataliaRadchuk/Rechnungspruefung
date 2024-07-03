from ttkbootstrap import Style

def apply_styles(style, font_size=20):
    font_large = f"Helvetica {font_size}"
    button_font = (f"Helvetica {font_size}", "bold")
    combobox_font = f"Helvetica {font_size}"
    combobox_list_font = f"Helvetica {font_size}"

    # Configure button style
    style.configure('TButton', font=button_font, padding=10)
    
    # Configure tab style
    style.configure('TNotebook.Tab', font=font_large, padding=[20, 10])
    
    # Configure combobox style
    style.configure('TCombobox', font=combobox_font)
    
    # Configure combobox list items style
    style.configure('TCombobox.LIST', font=combobox_list_font)

    return font_large, button_font, combobox_font

def reapply_styles(style, font_size=14):
    apply_styles(style, font_size)

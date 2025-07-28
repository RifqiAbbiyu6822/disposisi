import tkinter as tk

def create_menu_bar(parent, save_to_pdf_callback, clear_form_callback, quit_callback, show_shortcuts_callback, show_about_callback):
    menubar = tk.Menu(parent, font=("Segoe UI", 10))
    parent.config(menu=menubar)
    file_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10))
    menubar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Simpan ke PDF", command=save_to_pdf_callback)
    file_menu.add_separator()
    file_menu.add_command(label="Keluar", command=quit_callback)
    help_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10))
    menubar.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="Shortcuts", command=show_shortcuts_callback)
    help_menu.add_command(label="About", command=show_about_callback)
    return menubar 
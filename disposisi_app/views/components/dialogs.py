import tkinter as tk
from tkinter import messagebox

def show_shortcuts(parent):
    shortcuts_text = """
Keyboard Shortcuts:

Ctrl + S          : Simpan ke PDF
Ctrl + Shift + C  : Clear Form
Ctrl + Tab        : Next Tab
Ctrl + Shift + Tab: Previous Tab

Mouse Gestures:
- Double click untuk zoom
- Scroll untuk navigasi
- Drag untuk gesture navigation

Navigation:
- Tab untuk berpindah field
- Enter untuk konfirmasi
- Escape untuk cancel
    """
    help_window = tk.Toplevel(parent)
    help_window.title("Keyboard Shortcuts")
    help_window.geometry("500x400")
    help_window.resizable(False, False)
    try:
        help_window.iconbitmap('JapekELEVATED.ico')
    except:
        pass
    text_widget = tk.Text(help_window, wrap="word", font=("Segoe UI", 11), bg=getattr(parent, 'canvas_bg', '#f8fafc'), fg=getattr(parent, 'text_fg', '#1e293b'), relief="flat", padx=20, pady=20)
    text_widget.pack(fill="both", expand=True)
    text_widget.insert("1.0", shortcuts_text)
    text_widget.config(state="disabled")

def show_about(parent):
    about_text = """Aplikasi Formulir Disposisi

Version: 2.0
Fitur:
• Windows shortcuts support
• Touchpad gesture support
• Enhanced user interface
• Keyboard navigation
• PDF export functionality
• Google Sheets integration

Dikembangkan dengan Python dan Tkinter"""
    messagebox.showinfo("About", about_text, parent=parent) 
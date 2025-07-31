import tkinter as tk
from tkinter import ttk, messagebox
import os

def create_button_frame(parent, callbacks):
    """
    Creates the main action button frame with modern styling.
    """
    # Main container with card style
    button_container = tk.Frame(parent, bg="#FFFFFF", relief="flat")
    button_container.grid(row=2, column=0, sticky="ew", pady=(20, 0))
    
    # Add subtle border
    border_frame = tk.Frame(button_container, bg="#E2E8F0", height=1)
    border_frame.pack(fill="x", side="top")
    
    # Inner frame with padding
    button_frame = ttk.Frame(button_container, padding=(20, 15), style="Card.TFrame")
    button_frame.pack(fill="x")
    button_frame.columnconfigure(0, weight=1)

    # --- Quick Actions Section ---
    quick_actions_label = ttk.Label(button_frame, text="Quick Actions", 
                                   style="Caption.TLabel")
    quick_actions_label.grid(row=0, column=0, sticky="w", pady=(0, 10))
    
    action_frame = ttk.Frame(button_frame, style="Card.TFrame")
    action_frame.grid(row=1, column=0, sticky="ew", pady=(0, 15))
    action_frame.columnconfigure(0, weight=1)
    
    # Modern Clear Form Button with icon
    def on_clear_form():
        func = callbacks.get("clear_form")
        if callable(func):
            func()
        else:
            messagebox.showwarning("Fungsi Tidak Ditemukan", 
                                 "Fungsi untuk membersihkan form belum diimplementasikan.")
    
    clear_btn_frame = tk.Frame(action_frame, bg="#FFFFFF")
    clear_btn_frame.grid(row=0, column=0, sticky="w")
    
    btn_bersihkan = ttk.Button(clear_btn_frame, text="🔄  Reset Form", 
                              command=on_clear_form, style="Ghost.TButton")
    btn_bersihkan.pack(side="left")
    
    # Add hover effect
    def on_enter(e):
        btn_bersihkan.configure(style="Secondary.TButton")
    def on_leave(e):
        btn_bersihkan.configure(style="Ghost.TButton")
    
    btn_bersihkan.bind("<Enter>", on_enter)
    btn_bersihkan.bind("<Leave>", on_leave)

    # --- Main Actions Section ---
    main_actions_label = ttk.Label(button_frame, text="Complete & Submit", 
                                  style="Caption.TLabel")
    main_actions_label.grid(row=2, column=0, sticky="w", pady=(0, 10))
    
    selesai_frame = ttk.Frame(button_frame, style="Card.TFrame")
    selesai_frame.grid(row=3, column=0, sticky="ew")
    selesai_frame.columnconfigure(0, weight=1)

    def show_finish_dialog():
        from .finish_dialog import FinishDialog
        disposisi_labels = callbacks.get("get_disposisi_labels", lambda: [])()
        dialog_callbacks = {
            "save_pdf": callbacks.get("save_pdf"),
            "save_sheet": callbacks.get("save_sheet"),
            "send_email": callbacks.get("send_email")
        }
        dialog = FinishDialog(parent, disposisi_labels, dialog_callbacks)
        parent.wait_window(dialog)

    # Main submit button with modern design
    submit_container = tk.Frame(selesai_frame, bg="#FFFFFF")
    submit_container.grid(row=0, column=0, sticky="ew")
    
    # Button with gradient effect (simulated)
    btn_selesai = ttk.Button(submit_container, 
                            text="✅  Selesai & Lanjutkan", 
                            command=show_finish_dialog, 
                            style="Primary.TButton")
    btn_selesai.pack(fill="x")
    
    # Helper text
    helper_text = ttk.Label(selesai_frame, 
                           text="Simpan formulir dan pilih tindakan selanjutnya",
                           style="Caption.TLabel")
    helper_text.grid(row=1, column=0, sticky="w", pady=(5, 0))
    
    # Add shadow effect to main button - using valid color
    shadow = tk.Frame(submit_container, bg="#E5E7EB", height=2)
    shadow.pack(fill="x", side="bottom", pady=(2, 0))

    return button_frame
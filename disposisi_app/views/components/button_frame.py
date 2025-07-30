# disposisi_app/views/components/button_frame.py - FIXED VERSION
import tkinter as tk
from tkinter import ttk, messagebox
import os

def create_button_frame(parent, callbacks):
    """
    Creates the main action button frame with improved layout and positioning.
    """
    # Main container with proper layout management
    button_container = tk.Frame(parent, bg="#FFFFFF", relief="flat")
    button_container.grid(row=2, column=0, sticky="ew", pady=(15, 0))  # Reduced top padding
    
    # Add subtle border
    border_frame = tk.Frame(button_container, bg="#E2E8F0", height=1)
    border_frame.pack(fill="x", side="top")
    
    # Inner frame with reduced padding
    button_frame = ttk.Frame(button_container, padding=(15, 10), style="Card.TFrame")  # Reduced padding
    button_frame.pack(fill="x")
    button_frame.columnconfigure(0, weight=1)
    button_frame.columnconfigure(1, weight=0)  # Right section doesn't expand

    # --- Left Section: Quick Actions ---
    left_section = ttk.Frame(button_frame, style="Card.TFrame")
    left_section.grid(row=0, column=0, sticky="w", padx=(0, 20))

    # Quick actions header - smaller
    quick_actions_label = ttk.Label(left_section, text="Quick Actions", 
                                   style="Caption.TLabel")
    quick_actions_label.grid(row=0, column=0, sticky="w", pady=(0, 8))  # Reduced padding
    
    action_frame = ttk.Frame(left_section, style="Card.TFrame")
    action_frame.grid(row=1, column=0, sticky="w")
    
    # Modern Clear Form Button with icon - compact
    def on_clear_form():
        func = callbacks.get("clear_form")
        if callable(func):
            func()
        else:
            messagebox.showwarning("Fungsi Tidak Ditemukan", 
                                 "Fungsi untuk membersihkan form belum diimplementasikan.")
    
    btn_bersihkan = ttk.Button(action_frame, text="🔄 Reset Form", 
                              command=on_clear_form, style="Ghost.TButton")
    btn_bersihkan.pack(side="left", padx=(0, 10))
    
    # Add hover effect
    def on_enter(e):
        btn_bersihkan.configure(style="Secondary.TButton")
    def on_leave(e):
        btn_bersihkan.configure(style="Ghost.TButton")
    
    btn_bersihkan.bind("<Enter>", on_enter)
    btn_bersihkan.bind("<Leave>", on_leave)

    # --- Right Section: Main Actions ---
    right_section = ttk.Frame(button_frame, style="Card.TFrame")
    right_section.grid(row=0, column=1, sticky="e")

    # Main actions header - smaller
    main_actions_label = ttk.Label(right_section, text="Complete & Submit", 
                                  style="Caption.TLabel")
    main_actions_label.grid(row=0, column=0, sticky="e", pady=(0, 8))  # Reduced padding
    
    # Action buttons container
    actions_container = ttk.Frame(right_section, style="Card.TFrame")
    actions_container.grid(row=1, column=0, sticky="e")

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

    # Main submit button with compact design
    btn_selesai = ttk.Button(actions_container, 
                            text="✅ Selesai & Lanjutkan", 
                            command=show_finish_dialog, 
                            style="Primary.TButton")
    btn_selesai.pack(side="right")
    
    # Helper text - smaller and more compact
    helper_frame = ttk.Frame(right_section, style="Card.TFrame")
    helper_frame.grid(row=2, column=0, sticky="e", pady=(3, 0))  # Reduced padding
    
    helper_text = ttk.Label(helper_frame, 
                           text="Simpan formulir dan pilih tindakan selanjutnya",
                           style="Caption.TLabel")
    helper_text.pack(side="right")

    return button_frame
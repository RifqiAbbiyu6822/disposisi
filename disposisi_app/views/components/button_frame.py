import tkinter as tk
from tkinter import ttk, messagebox
import os

def create_button_frame(parent, callbacks):
    """
    Creates the main action button frame with form controls.
    """
    button_frame = ttk.Frame(parent, padding="10")
    button_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    button_frame.columnconfigure(0, weight=1)

    # --- Form Action Button ---
    action_frame = ttk.Frame(button_frame)
    action_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
    action_frame.columnconfigure(0, weight=1)
    
    # "Clear Form" Button
    def on_clear_form():
        func = callbacks.get("clear_form")
        if callable(func):
            func()
        else:
            messagebox.showwarning("Fungsi Tidak Ditemukan", "Fungsi untuk membersihkan form belum diimplementasikan.")
    
    btn_bersihkan = ttk.Button(action_frame, text="🔄 Bersihkan Form", command=on_clear_form, style="Secondary.TButton")
    btn_bersihkan.grid(row=0, column=0, padx=(5, 0), sticky="ew")


    # --- Selesai (Finish) Button Section ---
    selesai_frame = ttk.Frame(button_frame)
    selesai_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
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

    btn_selesai = ttk.Button(selesai_frame, text="✅ Selesai & Lanjutkan", 
                             command=show_finish_dialog, style="Success.TButton")
    btn_selesai.grid(row=0, column=0, sticky="ew")

    return button_frame
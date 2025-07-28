# disposisi_app/views/components/button_frame.py

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


    # --- Preview & Export Section (Initially hidden) ---
    preview_frame = ttk.LabelFrame(button_frame, text="Pratinjau & Opsi Ekspor", padding="10")
    preview_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    preview_frame.columnconfigure(0, weight=1)
    preview_frame.grid_remove()  # Hide it until "Selesai" is clicked
    parent.preview_frame = preview_frame  # Make it accessible from the parent

    # PDF preview area (to be filled by callback)
    parent.pdf_preview_label = ttk.Label(preview_frame, text="[Pratinjau PDF akan muncul di sini]", font=("Segoe UI", 9), foreground="#64748b")
    parent.pdf_preview_label.grid(row=0, column=0, sticky="ew", pady=(0, 8))

    # Export options
    export_options_frame = ttk.Frame(preview_frame)
    export_options_frame.grid(row=1, column=0, sticky="ew")
    export_options_frame.columnconfigure(0, weight=1)

    parent.chk_export_pdf = tk.BooleanVar(value=True)
    parent.chk_upload_sheet = tk.BooleanVar(value=False)
    parent.chk_send_email = tk.BooleanVar(value=False)
    chk_pdf = ttk.Checkbutton(export_options_frame, text="Export PDF", variable=parent.chk_export_pdf)
    chk_pdf.grid(row=0, column=0, sticky="w", padx=4)
    chk_sheet = ttk.Checkbutton(export_options_frame, text="Upload ke Sheet", variable=parent.chk_upload_sheet)
    chk_sheet.grid(row=0, column=1, sticky="w", padx=4)
    chk_email = ttk.Checkbutton(export_options_frame, text="Kirim Email", variable=parent.chk_send_email)
    chk_email.grid(row=0, column=2, sticky="w", padx=4)

    # Sub-checkbox posisi untuk email
    parent.email_posisi_vars = {}
    posisi = getattr(parent, 'posisi_options', ["Direktur Utama", "Direktur Keuangan", "Direktur Teknik", "GM Keuangan & Administrasi", "GM Operasional & Pemeliharaan", "Manager"])
    sub_frame = ttk.Frame(export_options_frame)
    sub_frame.grid(row=1, column=2, sticky="w", padx=4)
    for i, pos in enumerate(posisi):
        var = tk.BooleanVar(value=False)
        parent.email_posisi_vars[pos] = var
        chk = ttk.Checkbutton(sub_frame, text=pos, variable=var)
        chk.grid(row=0, column=i, sticky="w", padx=2)
    sub_frame.grid_remove()  # Hide initially
    parent.sub_email_frame = sub_frame

    def on_email_toggle():
        if parent.chk_send_email.get():
            sub_frame.grid()
        else:
            sub_frame.grid_remove()
    parent.chk_send_email.trace_add('write', lambda *args: on_email_toggle())

    # Action buttons for export
    action_frame = ttk.Frame(preview_frame)
    action_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
    action_frame.columnconfigure((0, 1), weight=1)
    btn_export_pdf = ttk.Button(action_frame, text="Export PDF", command=callbacks.get("export_pdf"), style="Primary.TButton")
    btn_export_pdf.grid(row=0, column=0, padx=(0, 4), sticky="ew")
    btn_upload_sheet = ttk.Button(action_frame, text="Upload ke Sheet", command=callbacks.get("upload_sheet"), style="Primary.TButton")
    btn_upload_sheet.grid(row=0, column=1, padx=(4, 0), sticky="ew")
    
    return button_frame
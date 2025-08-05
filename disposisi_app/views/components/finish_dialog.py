# Implementasi baru untuk disposisi_app/views/components/finish_dialog.py

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import traceback

# Import email sender dengan modifikasi untuk mengambil dari spreadsheet
from email_sender.send_email import EmailSender
from email_sender.template_handler import render_email_template

class FinishDialog(tk.Toplevel):
    def __init__(self, parent, disposisi_labels, callbacks):
        super().__init__(parent)
        self.transient(parent)
        self.parent = parent
        self.disposisi_labels = disposisi_labels
        self.callbacks = callbacks
        self.title("Finalisasi Disposisi")
        self.geometry("500x400")
        self.grab_set()
        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=(20, 20))
        main_frame.pack(fill="both", expand=True)
        title_label = ttk.Label(main_frame, text="Pilih Tindakan Selanjutnya", font=("Segoe UI", 16, "bold"))
        title_label.pack(pady=(0, 20))

        actions_frame = ttk.LabelFrame(main_frame, text="Opsi", padding=15)
        actions_frame.pack(fill="x", pady=(0, 15))

        self.save_pdf_var = tk.BooleanVar(value=True)
        self.save_sheet_var = tk.BooleanVar(value=True)
        self.send_email_var = tk.BooleanVar(value=False)

        ttk.Checkbutton(actions_frame, text="📄 Simpan sebagai PDF", variable=self.save_pdf_var).pack(anchor="w")
        ttk.Checkbutton(actions_frame, text="📊 Unggah ke Google Sheets", variable=self.save_sheet_var).pack(anchor="w")
        ttk.Checkbutton(actions_frame, text="📧 Kirim Disposisi via Email", variable=self.send_email_var, command=self._toggle_email_frame).pack(anchor="w", pady=(10,0))
        
        self.email_frame = ttk.LabelFrame(main_frame, text="Pilih Penerima Email", padding=15)
        self.email_vars = {}
        if not self.disposisi_labels:
            ttk.Label(self.email_frame, text="Tidak ada posisi yang dipilih.").pack()
        else:
            # Mapping untuk singkatan manager
            abbreviation_map = {
                "Manager Pemeliharaan": "pml",
                "Manager Operasional": "ops",
                "Manager Administrasi": "adm",
                "Manager Keuangan": "keu"
            }
            
            for label in self.disposisi_labels:
                var = tk.BooleanVar(value=True)
                # Gunakan singkatan untuk display
                display_label = abbreviation_map.get(label, label)
                if display_label in ["pml", "ops", "adm", "keu"]:
                    display_label = f"Manager {display_label}"
                self.email_vars[label] = var  # Keep original label for email lookup
                ttk.Checkbutton(self.email_frame, text=display_label, variable=var).pack(anchor="w")
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        ttk.Button(button_frame, text="Batal", command=self.destroy).grid(row=0, column=0, padx=5, sticky="ew")
        ttk.Button(button_frame, text="✅ Proses", command=self._process).grid(row=0, column=1, padx=5, sticky="ew")

    def _toggle_email_frame(self):
        if self.send_email_var.get():
            self.email_frame.pack(fill="x", pady=(0, 15))
        else:
            self.email_frame.pack_forget()

    def _process(self):
        try:
            # Proses simpan PDF jika dipilih
            if self.save_pdf_var.get() and callable(self.callbacks.get("save_pdf")):
                self.callbacks["save_pdf"]()
            
            # Proses simpan ke Sheet jika dipilih
            if self.save_sheet_var.get() and callable(self.callbacks.get("save_sheet")):
                self.callbacks["save_sheet"]()
            
            # Proses kirim email jika dipilih
            if self.send_email_var.get():
                selected_positions = [label for label, var in self.email_vars.items() if var.get()]
                if not selected_positions:
                    messagebox.showwarning("Peringatan", "Pilih setidaknya satu penerima email.", parent=self)
                    return
                # Call the main callback from the parent app
                if callable(self.callbacks.get("send_email")):
                    self.callbacks["send_email"](selected_positions)
            
            self.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan: {e}", parent=self)
            traceback.print_exc()
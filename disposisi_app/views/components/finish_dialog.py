# main_app.py

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import traceback

# Import from the email_sender package
from email_sender.send_email import EmailSender
from email_sender.template_handler import render_email_template

# --- The Finish Dialog (from your code) ---
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
            for label in self.disposisi_labels:
                var = tk.BooleanVar(value=True)
                self.email_vars[label] = var
                ttk.Checkbutton(self.email_frame, text=label, variable=var).pack(anchor="w")
        
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
            if self.send_email_var.get():
                selected_positions = [label for label, var in self.email_vars.items() if var.get()]
                if not selected_positions:
                    messagebox.showwarning("Peringatan", "Pilih setidaknya satu penerima email.", parent=self)
                    return
                # Call the main callback from the parent app
                if callable(self.callbacks.get("send_email")):
                    self.callbacks["send_email"](selected_positions)
            
            # Placeholder for other actions
            if self.save_pdf_var.get():
                print("Action: Save PDF")
            if self.save_sheet_var.get():
                print("Action: Save to Sheet")

            self.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan: {e}", parent=self)
            traceback.print_exc()

# --- Main Application ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistem Disposisi Surat")
        self.geometry("600x500")

        # --- Main Frame ---
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill="both", expand=True)

        # --- Example Form Fields ---
        ttk.Label(main_frame, text="Formulir Disposisi", font=("Segoe UI", 18, "bold")).pack(pady=10)
        
        self.nomor_surat = self._create_entry(main_frame, "Nomor Surat:", "DIS/2025/07/28-001")
        self.nama_pengirim = self._create_entry(main_frame, "Nama Pengirim:", "PT Mitra Sejati")
        self.perihal = self._create_entry(main_frame, "Perihal:", "Penawaran Kerjasama Proyek")
        
        ttk.Label(main_frame, text="Instruksi:").pack(anchor="w", pady=(10,0))
        self.instruksi = tk.Text(main_frame, height=5, width=50)
        self.instruksi.pack(fill="x", expand=True)
        self.instruksi.insert("1.0", "1. Mohon dipelajari dan berikan tanggapan.\n2. Jadwalkan pertemuan.")
        
        # --- The positions that can receive a disposition ---
        self.disposisi_options = ["Direktur Utama", "Direktur Keuangan", "GM Operasional & Pemeliharaan"]

        ttk.Button(main_frame, text="Selesai & Lanjutkan", command=self.open_finish_dialog).pack(pady=20, ipadx=10, ipady=5)

    def _create_entry(self, parent, label, default_text=""):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=5)
        ttk.Label(frame, text=label, width=15).pack(side="left")
        entry = ttk.Entry(frame)
        entry.pack(side="left", fill="x", expand=True)
        entry.insert(0, default_text)
        return entry

    def open_finish_dialog(self):
        callbacks = {
            "send_email": self._process_and_send_email,
            # Add other callbacks if needed
        }
        FinishDialog(self, self.disposisi_options, callbacks)

    def _process_and_send_email(self, selected_positions):
        """
        The main callback function that prepares data and sends the email.
        """
        # 1. Fetch recipient emails from the spreadsheet
        email_sender = EmailSender()
        recipient_emails = []
        failed_lookups = []
        
        for position in selected_positions:
            email, msg = email_sender.get_recipient_email(position)
            if email:
                recipient_emails.append(email)
            else:
                failed_lookups.append(f"{position} ({msg})")

        if failed_lookups:
            messagebox.showwarning(
                "Email Tidak Ditemukan",
                "Tidak dapat menemukan alamat email untuk posisi berikut:\n- " + "\n- ".join(failed_lookups),
                parent=self
            )

        if not recipient_emails:
            messagebox.showerror("Gagal", "Tidak ada alamat email yang valid untuk dikirimi.", parent=self)
            return

        # 2. Prepare data for the email template
        template_data = {
            'nomor_surat': self.nomor_surat.get(),
            'nama_pengirim': self.nama_pengirim.get(),
            'perihal': self.perihal.get(),
            'instruksi': self.instruksi.get("1.0", "end-1c"),
            'klasifikasi': ["PENTING", "SEGERA"], # Example data
            'tanggal': datetime.now().strftime('%d %B %Y'),
        }

        # 3. Render the HTML content
        html_content = render_email_template(template_data)
        subject = f"Disposisi Surat: {self.perihal.get()}"

        # 4. Send the email
        # Use set to avoid sending duplicate emails if positions share an email address
        unique_recipients = list(set(recipient_emails))
        success, message = email_sender.send_disposisi_email(unique_recipients, subject, html_content)
        
        # 5. Show final result
        if success:
            messagebox.showinfo("Sukses", message, parent=self)
        else:
            messagebox.showerror("Gagal Mengirim Email", message, parent=self)

if __name__ == "__main__":
    app = App()
    app.mainloop()
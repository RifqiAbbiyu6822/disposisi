import tkinter as tk
from tkinter import ttk, messagebox
from email_sender.send_email import EmailSender
import traceback

class FinishDialog(tk.Toplevel):
    """
    A dialog window that appears after clicking "Selesai & Lanjutkan".
    It provides final options for saving, uploading, and emailing the disposition.
    """
    def __init__(self, parent, disposisi_labels, callbacks):
        super().__init__(parent)
        self.transient(parent)
        self.parent = parent
        self.disposisi_labels = disposisi_labels  # List of positions to email
        self.callbacks = callbacks

        self.title("Finalisasi Disposisi")
        self.geometry("500x400")
        self.configure(bg="#f8fafc")

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grab_set()

        self._create_widgets()
        # self.wait_window(self) # This was the line causing the error, now removed.

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=(20, 20))
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(0, weight=1)

        # --- Title ---
        title_label = ttk.Label(main_frame, text="Pilih Tindakan Selanjutnya", font=("Segoe UI", 16, "bold"), anchor="center")
        title_label.grid(row=0, column=0, pady=(0, 20), sticky="ew")

        # --- Action Checkboxes ---
        actions_frame = ttk.LabelFrame(main_frame, text="Opsi", padding=15)
        actions_frame.grid(row=1, column=0, pady=(0, 15), sticky="ew")
        actions_frame.columnconfigure(0, weight=1)

        self.save_pdf_var = tk.BooleanVar(value=True)
        self.save_sheet_var = tk.BooleanVar(value=True)
        self.send_email_var = tk.BooleanVar(value=False)

        ttk.Checkbutton(actions_frame, text="📄 Simpan sebagai PDF", variable=self.save_pdf_var).pack(anchor="w", pady=3)
        ttk.Checkbutton(actions_frame, text="📊 Unggah ke Google Sheets", variable=self.save_sheet_var).pack(anchor="w", pady=3)
        
        email_check = ttk.Checkbutton(actions_frame, text="📧 Kirim Disposisi via Email", variable=self.send_email_var, command=self._toggle_email_frame)
        email_check.pack(anchor="w", pady=(10, 3))

        # --- Email Recipients Frame (Initially Hidden) ---
        self.email_frame = ttk.LabelFrame(main_frame, text="Pilih Penerima Email", padding=15)
        self.email_vars = {}

        if not self.disposisi_labels:
            ttk.Label(self.email_frame, text="Tidak ada posisi yang dipilih di form disposisi.",).pack() #Removed style
        else:
            for label in self.disposisi_labels:
                var = tk.BooleanVar(value=True)
                self.email_vars[label] = var
                ttk.Checkbutton(self.email_frame, text=label, variable=var).pack(anchor="w")

        # --- Action Buttons ---
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, pady=(20, 0), sticky="ew")
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)

        ttk.Button(button_frame, text="Batal", command=self.destroy, style="Secondary.TButton").grid(row=0, column=0, padx=(0, 5), sticky="ew")
        ttk.Button(button_frame, text="✅ Proses", command=self._process, style="Success.TButton").grid(row=0, column=1, padx=(5, 0), sticky="ew")

    def _toggle_email_frame(self):
        """Show or hide the email recipient selection frame."""
        if self.send_email_var.get():
            self.email_frame.grid(row=2, column=0, pady=(0, 15), sticky="ew")
        else:
            self.email_frame.grid_forget()

    def _process(self):
        """Process the selected final actions."""
        try:
            if self.save_pdf_var.get() and callable(self.callbacks.get("save_pdf")):
                self.callbacks["save_pdf"]()
            
            if self.save_sheet_var.get() and callable(self.callbacks.get("save_sheet")):
                self.callbacks["save_sheet"]()

            if self.send_email_var.get():
                selected_positions = [label for label, var in self.email_vars.items() if var.get()]
                if not selected_positions:
                    messagebox.showwarning("Peringatan", "Tidak ada penerima email yang dipilih.", parent=self)
                else:
                    self._send_emails_to_positions(selected_positions)

            self.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan saat memproses: {e}", parent=self)
            traceback.print_exc()

    def _send_emails_to_positions(self, positions):
        """Fetch emails and call the main email function."""
        email_sender = EmailSender()
        recipient_emails = []
        failed_lookups = []

        for pos in positions:
            email, msg = email_sender.get_recipient_email(pos)
            if email:
                recipient_emails.append(email)
            else:
                failed_lookups.append(pos)
        
        if failed_lookups:
            messagebox.showwarning("Email Tidak Ditemukan", 
                                 "Tidak dapat menemukan alamat email untuk posisi berikut:\n- " + "\n- ".join(failed_lookups),
                                 parent=self)

        if recipient_emails and callable(self.callbacks.get("send_email")):
            self.callbacks["send_email"](list(set(recipient_emails))) # Use set to avoid duplicates
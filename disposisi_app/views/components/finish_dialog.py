import tkinter as tk
from tkinter import ttk, messagebox
from .email_manager import EmailManager
import os
from dotenv import load_dotenv

# Load environment variables at module level
load_dotenv()

class FinishDialog(tk.Toplevel):
    """Dialog yang muncul saat tombol Selesai & Lanjutkan ditekan"""
    
    def __init__(self, parent, disposisi_labels, callbacks):
        super().__init__(parent)
        
        self.parent = parent
        self.disposisi_labels = disposisi_labels
        self.callbacks = callbacks
        
        # Initialize email manager with environment variables
        sender_email = os.getenv('SENDER_EMAIL')
        sender_password = os.getenv('SENDER_PASSWORD')
        
        if not sender_email or not sender_password:
            messagebox.showerror(
                "Konfigurasi Email Error",
                "Email sender belum dikonfigurasi. Pastikan file .env berisi SENDER_EMAIL dan SENDER_PASSWORD."
            )
        
        self.email_manager = EmailManager()
        
        self.title("Pilih Opsi Penyimpanan & Pengiriman")
        self.configure(bg="#f8fafc")
        self.resizable(False, False)
        
        # Modern styling
        self.style = ttk.Style(self)
        self._setup_styles()
        
        # Create main container with padding
        main_frame = ttk.Frame(self, style="Dialog.TFrame", padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # Title and description
        title = ttk.Label(main_frame, 
                         text="✅ Selesai",
                         style="DialogTitle.TLabel")
        title.pack(anchor="w", pady=(0, 10))
        
        description = ttk.Label(main_frame,
                              text="Pilih opsi untuk menyimpan dan mengirim disposisi:",
                              style="DialogDesc.TLabel")
        description.pack(anchor="w", pady=(0, 20))
        
        # Checkboxes frame
        checkbox_frame = ttk.LabelFrame(main_frame, text="Opsi", padding="10")
        checkbox_frame.pack(fill="x", pady=(0, 20))
        
        # Save options
        self.save_pdf_var = tk.BooleanVar(value=True)
        self.save_sheet_var = tk.BooleanVar(value=True)
        self.send_email_var = tk.BooleanVar(value=True)
        
        save_pdf_cb = ttk.Checkbutton(checkbox_frame, 
                                     text="💾 Simpan sebagai PDF",
                                     variable=self.save_pdf_var,
                                     style="Dialog.TCheckbutton")
        save_pdf_cb.pack(anchor="w", pady=2)
        
        save_sheet_cb = ttk.Checkbutton(checkbox_frame,
                                       text="📊 Simpan ke Google Sheet",
                                       variable=self.save_sheet_var,
                                       style="Dialog.TCheckbutton")
        save_sheet_cb.pack(anchor="w", pady=2)
        
        send_email_cb = ttk.Checkbutton(checkbox_frame,
                                       text="📧 Kirim Email",
                                       variable=self.send_email_var,
                                       style="Dialog.TCheckbutton")
        send_email_cb.pack(anchor="w", pady=2)

        # Email recipient options (only shown when send_email is checked)
        self.email_frame = ttk.LabelFrame(main_frame, text="Pilih Penerima Email", padding="10")
        
        # Create checkboxes for each disposisi label
        self.email_vars = {}
        for label in self.disposisi_labels:
            if label != "Lainnya":  # Skip "Lainnya" option
                var = tk.BooleanVar(value=True)
                self.email_vars[label] = var
                cb = ttk.Checkbutton(self.email_frame, 
                                   text=label,
                                   variable=var,
                                   style="Dialog.TCheckbutton")
                cb.pack(anchor="w", pady=2)
        
        # Only show email options when send_email is checked
        def toggle_email_frame(*args):
            if self.send_email_var.get():
                self.email_frame.pack(fill="x", pady=(0, 20))
            else:
                self.email_frame.pack_forget()
        
        self.send_email_var.trace_add('write', toggle_email_frame)
        toggle_email_frame()  # Initial state
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        
        cancel_btn = ttk.Button(button_frame,
                              text="Batal",
                              style="Secondary.TButton",
                              command=self.destroy)
        cancel_btn.pack(side="left", padx=5)
        
        process_btn = ttk.Button(button_frame,
                               text="✨ Proses",
                               style="Primary.TButton",
                               command=self._process)
        process_btn.pack(side="right", padx=5)
        
        # Center dialog on parent window
        self.transient(parent)
        self.grab_set()
        
    def _setup_styles(self):
        """Setup custom styles for dialog"""
        self.style.configure("Dialog.TFrame",
                           background="#f8fafc")
        
        self.style.configure("DialogTitle.TLabel",
                           font=("Segoe UI", 16, "bold"),
                           background="#f8fafc")
        
        self.style.configure("DialogDesc.TLabel",
                           font=("Segoe UI", 10),
                           background="#f8fafc")
        
        self.style.configure("Dialog.TCheckbutton",
                           background="#f8fafc")
        
        # Button styles
        self.style.configure("Primary.TButton",
                           font=("Segoe UI", 10),
                           padding=10)
        
        self.style.configure("Secondary.TButton",
                           font=("Segoe UI", 10),
                           padding=10)
        
    def _process(self):
        """Process selected options"""
        selected_positions = [
            label for label, var in self.email_vars.items()
            if var.get()
        ]
        
        # Get form data for email
        form_data = self.callbacks.get("get_form_data", lambda: {})()
        
        if self.save_pdf_var.get() and callable(self.callbacks.get("save_pdf")):
            self.callbacks["save_pdf"]()
            
        if self.save_sheet_var.get() and callable(self.callbacks.get("save_sheet")):
            self.callbacks["save_sheet"]()
            
        if (self.send_email_var.get() and selected_positions):
            # Send emails using EmailManager
            self.email_manager.send_disposisi_emails(selected_positions, form_data)
        
        self.destroy()
        
        save_pdf_cb = ttk.Checkbutton(checkbox_frame, 
                                     text="💾 Simpan sebagai PDF",
                                     variable=self.save_pdf_var,
                                     style="Dialog.TCheckbutton")
        save_pdf_cb.pack(anchor="w", pady=2)
        
        save_sheet_cb = ttk.Checkbutton(checkbox_frame,
                                       text="📊 Upload ke Google Sheets",
                                       variable=self.save_sheet_var,
                                       style="Dialog.TCheckbutton")
        save_sheet_cb.pack(anchor="w", pady=2)
        
        send_email_cb = ttk.Checkbutton(checkbox_frame,
                                       text="📧 Kirim Email ke Penerima Disposisi",
                                       variable=self.send_email_var,
                                       style="Dialog.TCheckbutton")
        send_email_cb.pack(anchor="w", pady=2)
        
        # Email recipients
        email_frame = ttk.LabelFrame(main_frame, text="Penerima Email", padding="10")
        email_frame.pack(fill="x", pady=(0, 20))
        
        self.email_vars = {}
        for label in self.disposisi_labels:
            var = tk.BooleanVar(value=True)
            self.email_vars[label] = var
            cb = ttk.Checkbutton(email_frame,
                                text=label,
                                variable=var,
                                style="Dialog.TCheckbutton")
            cb.pack(anchor="w", pady=2)
        
        if not self.disposisi_labels:
            ttk.Label(email_frame,
                     text="Belum ada penerima disposisi dipilih",
                     style="DialogDesc.TLabel").pack(pady=5)
        
        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x")
        btn_frame.columnconfigure(1, weight=1)
        
        cancel_btn = ttk.Button(btn_frame,
                               text="Batal",
                               style="Secondary.TButton",
                               command=self.destroy)
        cancel_btn.grid(row=0, column=0, padx=(0, 5))
        
        process_btn = ttk.Button(btn_frame,
                                text="✅ Proses",
                                style="Success.TButton",
                                command=self._process)
        process_btn.grid(row=0, column=1, sticky="e")
        
        # Center dialog on parent window
        self.transient(parent)
        self.grab_set()
        self._center_dialog()
        
    def _setup_styles(self):
        """Setup custom styles for dialog"""
        self.style.configure("Dialog.TFrame",
                           background="#f8fafc")
        
        self.style.configure("DialogTitle.TLabel",
                           font=("Segoe UI", 16, "bold"),
                           background="#f8fafc",
                           foreground="#1f2937")
        
        self.style.configure("DialogDesc.TLabel",
                           font=("Segoe UI", 10),
                           background="#f8fafc",
                           foreground="#6b7280")
        
        self.style.configure("Dialog.TCheckbutton",
                           font=("Segoe UI", 10),
                           background="#f8fafc",
                           foreground="#1f2937")
        
    def _center_dialog(self):
        """Center the dialog on parent window"""
        self.update_idletasks()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        
        width = self.winfo_width()
        height = self.winfo_height()
        
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        
        self.geometry(f"+{x}+{y}")
        
    def _process(self):
        """Process selected options"""
        selected_emails = [
            label for label, var in self.email_vars.items()
            if var.get()
        ]
        
        if self.save_pdf_var.get() and callable(self.callbacks.get("save_pdf")):
            self.callbacks["save_pdf"]()
            
        if self.save_sheet_var.get() and callable(self.callbacks.get("save_sheet")):
            self.callbacks["save_sheet"]()
            
        if (self.send_email_var.get() and callable(self.callbacks.get("send_email")) 
            and selected_emails):
            self.callbacks["send_email"](selected_emails)
        
        self.destroy()

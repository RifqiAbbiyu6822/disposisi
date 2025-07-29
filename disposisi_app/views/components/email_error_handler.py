# disposisi_app/views/components/email_error_handler.py

import tkinter as tk
from tkinter import ttk, messagebox
import os
import json
import traceback
from pathlib import Path

class EmailConfigDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Konfigurasi Email")
        self.geometry("500x350")
        self.transient(parent)
        self.grab_set()
        
        # Pastikan dialog muncul di tengah parent window
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
        
        self._create_widgets()
        self._load_config()
    
    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill="both", expand=True)
        
        # Judul
        title_label = ttk.Label(main_frame, 
                              text="Konfigurasi SMTP Email", 
                              font=("Segoe UI", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Form frame
        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill="both", expand=True)
        
        # SMTP Server
        ttk.Label(form_frame, text="SMTP Server:").grid(row=0, column=0, sticky="w", pady=5)
        self.smtp_server_var = tk.StringVar(value="smtp.gmail.com")
        ttk.Entry(form_frame, textvariable=self.smtp_server_var, width=30).grid(row=0, column=1, sticky="ew", pady=5)
        
        # SMTP Port
        ttk.Label(form_frame, text="SMTP Port:").grid(row=1, column=0, sticky="w", pady=5)
        self.smtp_port_var = tk.StringVar(value="587")
        ttk.Entry(form_frame, textvariable=self.smtp_port_var, width=30).grid(row=1, column=1, sticky="ew", pady=5)
        
        # Use TLS option
        self.use_tls_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(form_frame, text="Gunakan TLS", variable=self.use_tls_var).grid(row=2, column=0, columnspan=2, sticky="w", pady=5)
        
        # Email
        ttk.Label(form_frame, text="Email:").grid(row=3, column=0, sticky="w", pady=5)
        self.email_var = tk.StringVar(value="")
        ttk.Entry(form_frame, textvariable=self.email_var, width=30).grid(row=3, column=1, sticky="ew", pady=5)
        
        # Password
        ttk.Label(form_frame, text="Password:").grid(row=4, column=0, sticky="w", pady=5)
        self.password_var = tk.StringVar(value="")
        password_entry = ttk.Entry(form_frame, textvariable=self.password_var, width=30, show="*")
        password_entry.grid(row=4, column=1, sticky="ew", pady=5)
        
        # Show/hide password
        self.show_password_var = tk.BooleanVar(value=False)
        
        def toggle_password_visibility():
            if self.show_password_var.get():
                password_entry.config(show="")
            else:
                password_entry.config(show="*")
                
        ttk.Checkbutton(form_frame, text="Tampilkan password", 
                      variable=self.show_password_var,
                      command=toggle_password_visibility).grid(row=5, column=0, columnspan=2, sticky="w", pady=5)
        
        # Bantuan untuk Google App Password
        info_label = ttk.Label(form_frame, 
                             text="Catatan: Untuk Gmail, gunakan App Password.\nLihat: https://support.google.com/accounts/answer/185833",
                             foreground="#1D4ED8",
                             font=("Segoe UI", 9))
        info_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=5)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        
        # Simpan button
        save_btn = ttk.Button(button_frame, text="Simpan Konfigurasi", command=self._save_config)
        save_btn.pack(side="left", padx=(0, 10))
        
        # Test button
        test_btn = ttk.Button(button_frame, text="Test Koneksi", command=self._test_connection)
        test_btn.pack(side="left")
        
        # Batal button
        cancel_btn = ttk.Button(button_frame, text="Batal", command=self.destroy)
        cancel_btn.pack(side="right")
        
        # Configure grid layout
        form_frame.columnconfigure(1, weight=1)
    
    def _load_config(self):
        """Load configuration from config.py or .env file"""
        try:
            # Coba load dari config.py
            from email_sender import config
            self.smtp_server_var.set(getattr(config, 'EMAIL_HOST', 'smtp.gmail.com'))
            self.smtp_port_var.set(str(getattr(config, 'EMAIL_PORT', '587')))
            self.use_tls_var.set(getattr(config, 'EMAIL_USE_TLS', True))
            self.email_var.set(getattr(config, 'EMAIL_HOST_USER', ''))
            self.password_var.set(getattr(config, 'EMAIL_HOST_PASSWORD', ''))
        except ImportError:
            # Coba load dari .env file
            try:
                import dotenv
                dotenv.load_dotenv()
                self.smtp_server_var.set(os.environ.get('EMAIL_HOST', 'smtp.gmail.com'))
                self.smtp_port_var.set(os.environ.get('EMAIL_PORT', '587'))
                self.use_tls_var.set(os.environ.get('EMAIL_USE_TLS', 'True').lower() == 'true')
                self.email_var.set(os.environ.get('EMAIL_HOST_USER', ''))
                self.password_var.set(os.environ.get('EMAIL_HOST_PASSWORD', ''))
            except Exception as e:
                print(f"Warning: Failed to load configuration from .env: {str(e)}")
                
            # Coba load dari email_sender/.env juga
            try:
                env_path = Path(__file__).parent.parent.parent.parent / 'email_sender' / '.env'
                if env_path.exists():
                    dotenv.load_dotenv(env_path)
                    if not self.email_var.get():
                        self.email_var.set(os.environ.get('SENDER_EMAIL', ''))
                    if not self.password_var.get():
                        self.password_var.set(os.environ.get('SENDER_PASSWORD', ''))
            except Exception as e:
                print(f"Warning: Failed to load configuration from email_sender/.env: {str(e)}")
    
    def _save_config(self):
        """Save configuration to both config.py and .env file"""
        try:
            # Get values
            smtp_server = self.smtp_server_var.get()
            smtp_port = self.smtp_port_var.get()
            use_tls = self.use_tls_var.get()
            email = self.email_var.get()
            password = self.password_var.get()
            
            # Validasi input
            if not smtp_server or not smtp_port or not email or not password:
                messagebox.showerror("Error", "Semua field harus diisi!", parent=self)
                return
            
            # Coba simpan ke config.py
            config_path = Path(__file__).parent.parent.parent.parent / 'email_sender' / 'config.py'
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    config_content = f.read()
                
                # Update nilai dalam file config.py
                import re
                config_content = re.sub(r"EMAIL_HOST\s*=\s*['\"](.*)['\"]", 
                                      f"EMAIL_HOST = '{smtp_server}'", config_content)
                config_content = re.sub(r"EMAIL_PORT\s*=\s*(\d+)", 
                                      f"EMAIL_PORT = {smtp_port}", config_content)
                config_content = re.sub(r"EMAIL_USE_TLS\s*=\s*(True|False)", 
                                      f"EMAIL_USE_TLS = {use_tls}", config_content)
                config_content = re.sub(r"EMAIL_HOST_USER\s*=\s*['\"](.*)['\"]", 
                                      f"EMAIL_HOST_USER = '{email}'", config_content)
                config_content = re.sub(r"EMAIL_HOST_PASSWORD\s*=\s*['\"](.*)['\"]", 
                                      f"EMAIL_HOST_PASSWORD = '{password}'", config_content)
                
                # Tulis kembali file
                with open(config_path, 'w', encoding='utf-8') as f:
                    f.write(config_content)
                
                print(f"Configuration saved to {config_path}")
            
            # Simpan ke .env file juga
            env_path = Path(__file__).parent.parent.parent.parent / '.env'
            env_content = f"""EMAIL_HOST={smtp_server}
EMAIL_PORT={smtp_port}
EMAIL_USE_TLS={'True' if use_tls else 'False'}
EMAIL_HOST_USER={email}
EMAIL_HOST_PASSWORD={password}
"""
            with open(env_path, 'w', encoding='utf-8') as f:
                f.write(env_content)
            
            print(f"Configuration saved to {env_path}")
            
            # Simpan ke email_sender/.env juga
            env_path = Path(__file__).parent.parent.parent.parent / 'email_sender' / '.env'
            env_content = f"""SENDER_EMAIL={email}
SENDER_PASSWORD={password}
"""
            with open(env_path, 'w', encoding='utf-8') as f:
                f.write(env_content)
            
            print(f"Configuration saved to {env_path}")
            
            messagebox.showinfo("Sukses", "Konfigurasi email berhasil disimpan!", parent=self)
            self.destroy()
            
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"Gagal menyimpan konfigurasi: {str(e)}", parent=self)
    
    def _test_connection(self):
        """Test SMTP connection with the provided credentials"""
        import smtplib
        
        smtp_server = self.smtp_server_var.get()
        smtp_port = int(self.smtp_port_var.get())
        use_tls = self.use_tls_var.get()
        email = self.email_var.get()
        password = self.password_var.get()
        
        if not smtp_server or not smtp_port or not email or not password:
            messagebox.showerror("Error", "Semua field harus diisi!", parent=self)
            return
        
        try:
            if use_tls:
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
            else:
                server = smtplib.SMTP_SSL(smtp_server, smtp_port)
            
            server.login(email, password)
            server.quit()
            
            messagebox.showinfo("Sukses", "Koneksi SMTP berhasil!", parent=self)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"Gagal terhubung ke server SMTP: {str(e)}", parent=self)

def handle_email_error(parent_window, error_message=None):
    """
    Menampilkan dialog konfigurasi email ketika terjadi error email
    
    Args:
        parent_window: Window induk
        error_message: Pesan error (opsional)
    """
    if error_message:
        messagebox.showerror("Email Error", error_message, parent=parent_window)
    
    response = messagebox.askyesno(
        "Konfigurasi Email", 
        "Email belum dikonfigurasi dengan benar. Apakah Anda ingin mengatur konfigurasi email sekarang?",
        parent=parent_window
    )
    
    if response:
        EmailConfigDialog(parent_window)
        return True
    return False

# Contoh penggunaan:
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Test Email Error Handler")
    root.geometry("300x200")
    
    def show_dialog():
        handle_email_error(root, "Email belum dikonfigurasi. Pastikan file config.py atau .env berisi konfigurasi email.")
    
    ttk.Button(root, text="Test Dialog", command=show_dialog).pack(pady=50)
    
    root.mainloop()
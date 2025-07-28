from email_sender.send_email import EmailSender
import tkinter as tk
from tkinter import messagebox

class EmailManager:
    def __init__(self):
        self.email_sender = EmailSender()

    def send_disposisi_emails(self, positions, data):
        """
        Send emails to multiple positions with disposisi data
        
        Args:
            positions (list): List of positions to send emails to
            data (dict): Dictionary containing disposisi data
        """
        results = []
        for position in positions:
            try:
                status, message = self.email_sender.send_disposisi_email(
                    position=position,
                    nama_pengirim=data.get('nama_pengirim', ''),
                    nomor_surat=data.get('nomor_surat', ''),
                    perihal=data.get('perihal', ''),
                    instruksi=data.get('instruksi', '')
                )
                results.append((position, status, message))
            except Exception as e:
                results.append((position, False, f"Error sending to {position}: {str(e)}"))

        # Show summary of email sending results
        success_count = sum(1 for _, status, _ in results if status)
        if success_count == len(positions):
            messagebox.showinfo("Sukses", "Email berhasil dikirim ke semua penerima")
        else:
            error_messages = "\n".join([f"{pos}: {msg}" for pos, status, msg in results if not status])
            messagebox.showerror("Error", f"Beberapa email gagal terkirim:\n{error_messages}")

        return results

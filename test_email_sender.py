# test_email_complete.py
from email_sender.send_email import EmailSender
from email_sender.template_handler import render_email_template
from datetime import datetime

def test_complete_email():
    email_sender = EmailSender()
    
    # Test data
    template_data = {
        'nomor_surat': 'TEST/2025/001',
        'nama_pengirim': 'Test Pengirim',
        'perihal': 'Testing Email Disposisi',
        'tanggal': datetime.now().strftime('%d %B %Y'),
        'klasifikasi': ['PENTING'],
        'instruksi_list': ['Test instruksi 1', 'Test instruksi 2'],
        'tahun': 2025
    }
    
    html_content = render_email_template(template_data)
    
    # Test kirim ke satu email
    success, message = email_sender.send_disposisi_email(
        ['imneon2003@gmail.com'], 
        'Test Disposisi Email', 
        html_content
    )
    
    print(f"Result: {success}")
    print(f"Message: {message}")

if __name__ == "__main__":
    test_complete_email()
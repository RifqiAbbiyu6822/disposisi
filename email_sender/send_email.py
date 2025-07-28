import os
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from pathlib import Path
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from .template_handler import render_email_template
from . import config

# Load environment variables
load_dotenv()

# Constants for Google Sheets - gunakan admin sheet
CREDENTIALS_FILE = 'admin/credentials.json'
SHEET_ID = '1LoAzVPBMJo08uPHR7MdzGEXmaoFXMrTSKS0vl099qrU'  # Admin sheet ID
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

class EmailSender:
    def __init__(self):
        # Try multiple sources for email configuration
        self.sender_email = os.getenv('SENDER_EMAIL') or os.getenv('EMAIL_HOST_USER')
        self.sender_password = os.getenv('SENDER_PASSWORD') or os.getenv('EMAIL_HOST_PASSWORD')
        
        # Fallback to config file
        if not self.sender_email or not self.sender_password:
            try:
                from config import EMAIL_HOST_USER, EMAIL_HOST_PASSWORD
                self.sender_email = EMAIL_HOST_USER
                self.sender_password = EMAIL_HOST_PASSWORD
            except:
                pass
        
        self.sheets_service = self._get_sheets_service()
    
    def _get_sheets_service(self):
        """Initialize Google Sheets service with admin credentials"""
        try:
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
            return build('sheets', 'v4', credentials=creds)
        except Exception as e:
            print(f"Error initializing sheets service: {str(e)}")
            return None
    
    def get_recipient_email(self, position):
        """Fetch email address from admin spreadsheet based on position"""
        try:
            if not self.sheets_service:
                return None, "Google Sheets service not available"

            # Read from Sheet1 which contains position-email mapping
            range_name = 'Sheet1!A2:B'  # Skip header row
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=SHEET_ID,
                range=range_name
            ).execute()

            values = result.get('values', [])
            if not values:
                return None, f"Email tidak ditemukan untuk posisi: {position}"

            # Search for matching position
            for row in values:
                if len(row) >= 2:
                    sheet_position = row[0].strip()
                    sheet_email = row[1].strip()
                    
                    # Case-insensitive comparison
                    if sheet_position.lower() == position.lower():
                        # Basic email validation
                        if '@' in sheet_email and '.' in sheet_email:
                            return sheet_email, f"Email ditemukan: {sheet_email}"
                        else:
                            return None, f"Format email tidak valid untuk {position}: {sheet_email}"

            return None, f"Email tidak ditemukan untuk posisi: {position}"
            
        except Exception as e:
            return None, f"Error membaca data email: {str(e)}"
        
    def send_disposisi_email(self, position, nama_pengirim, nomor_surat, perihal, instruksi):
        """Send disposition email to a specific position"""
        try:
            # Validate email configuration
            if not self.sender_email or not self.sender_password:
                return False, "Konfigurasi email pengirim tidak ditemukan. Periksa file .env atau config.py"

            # Get recipient email from spreadsheet
            recipient_email, message = self.get_recipient_email(position)
            if not recipient_email:
                return False, message

            print(f"Mengirim email ke {position} ({recipient_email})")
            
            # Create message
            msg = MIMEMultipart('alternative')
            msg['Subject'] = f'Disposisi Surat - {nomor_surat}'
            msg['From'] = self.sender_email
            msg['To'] = recipient_email

            # Prepare template data
            template_data = {
                'position': position,
                'nama_pengirim': nama_pengirim,
                'nomor_surat': nomor_surat, 
                'perihal': perihal,
                'instruksi': instruksi,
                'tanggal': datetime.now().strftime('%d %B %Y'),
                'tahun': datetime.now().year
            }

            # Get HTML content from template
            html_content = render_email_template(template_data)

            # Attach HTML content
            part = MIMEText(html_content, 'html')
            msg.attach(part)

            # Send email with proper error handling
            try:
                # Try SSL first (port 465)
                server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
                server.login(self.sender_email, self.sender_password)
            except:
                # Fallback to TLS (port 587)
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(self.sender_email, self.sender_password)

            # Send email
            server.sendmail(
                self.sender_email,
                recipient_email,
                msg.as_string()
            )
            server.quit()
            
            return True, f"Email berhasil dikirim ke {position} ({recipient_email})"
            
        except smtplib.SMTPAuthenticationError:
            return False, "Gagal autentikasi email. Periksa email dan password di file .env"
        except smtplib.SMTPException as e:
            return False, f"Gagal mengirim email (SMTP Error): {str(e)}"
        except Exception as e:
            return False, f"Gagal mengirim email: {str(e)}"
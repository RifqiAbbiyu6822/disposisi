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

# Constants for Google Sheets
CREDENTIALS_FILE = 'admin/credentials.json'
SHEET_ID = '13-EgGz8JYYQ7FLQeCcVyPEXwYFruSGyz8KxPXrJlB7c'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

class EmailSender:
    def __init__(self):
        self.sender_email = os.getenv('SENDER_EMAIL')
        self.sender_password = os.getenv('SENDER_PASSWORD')
        self.sheets_service = self._get_sheets_service()
    
    def _get_sheets_service(self):
        """Initialize Google Sheets service with admin credentials"""
        try:
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
            return build('sheets', 'v4', credentials=creds)
        except Exception as e:
            print(f"Error initializing sheets service: {str(e)}")
            return None
    
    def validate_spreadsheet_structure(self):
        """Validate that the spreadsheet has the correct structure"""
        try:
            result = self.sheets_service.spreadsheets().get(
                spreadsheetId=SHEET_ID
            ).execute()
            
            # Check if 'Penerima' sheet exists
            sheet_exists = any(sheet['properties']['title'] == 'Penerima' 
                             for sheet in result['sheets'])
            if not sheet_exists:
                return False, "Sheet 'Penerima' tidak ditemukan di spreadsheet"
            
            # Check column headers
            range_name = 'Penerima!A1:B1'
            headers = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=SHEET_ID,
                range=range_name
            ).execute()
            
            if not headers.get('values'):
                return False, "Header tidak ditemukan di spreadsheet"
            
            header_row = headers['values'][0]
            if len(header_row) < 2:
                return False, "Struktur kolom tidak sesuai (diperlukan kolom Posisi dan Email)"
                
            return True, "Struktur spreadsheet valid"
        except Exception as e:
            return False, f"Error validasi spreadsheet: {str(e)}"

    def get_recipient_email(self, position):
        """Fetch email address from spreadsheet based on position"""
        try:
            if not self.sheets_service:
                return None, config.EMAIL_CONFIG_ERROR

            # Validate spreadsheet structure first
            is_valid, message = self.validate_spreadsheet_structure()
            if not is_valid:
                return None, message

            # Get all recipient data
            range_name = f'{config.RECIPIENT_SHEET_NAME}!{config.RECIPIENT_RANGE}'
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=SHEET_ID,
                range=range_name
            ).execute()

            values = result.get('values', [])
            if not values:
                return None, config.EMAIL_NOT_FOUND.format(position=position)

            # Search for matching position
            for row in values:
                if len(row) >= 2 and row[config.POSITION_COL].lower().strip() == position.lower().strip():
                    email = row[config.EMAIL_COL].strip()
                    # Basic email validation
                    if '@' in email and '.' in email:
                        return email, config.EMAIL_SUCCESS.format(recipient=email)
                    else:
                        return None, config.EMAIL_VALIDATION_ERROR.format(
                            position=position,
                            email=email
                        )

            return None, config.EMAIL_NOT_FOUND.format(position=position)
        except Exception as e:
            return None, config.EMAIL_SHEET_ERROR.format(error=str(e))
        
    def send_disposisi_email(self, position, nama_pengirim, nomor_surat, perihal, instruksi):
        try:
            # Validate position
            if not position or not position.strip():
                return False, "Posisi tidak boleh kosong"

            # Get recipient email from spreadsheet with validation
            recipient_email, message = self.get_recipient_email(position)
            if not recipient_email:
                return False, message

            print(f"Mengirim email ke {position} ({recipient_email})")
            
            # Create message container
            msg = MIMEMultipart('alternative')
            msg['Subject'] = f'Disposisi Surat - {nomor_surat}'
            msg['From'] = self.sender_email
            msg['To'] = recipient_email

            # Get HTML content
            html_content = get_email_template(
                position=position,
                nama_pengirim=nama_pengirim,
                nomor_surat=nomor_surat,
                perihal=perihal,
                instruksi=instruksi
            )

            # Attach HTML content
            part = MIMEText(html_content, 'html')
            msg.attach(part)

            # Create secure SSL/TLS connection
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.login(self.sender_email, self.sender_password)

            # Send email
            server.sendmail(
                self.sender_email,
                recipient_email,
                msg.as_string()
            )
            server.quit()
            return True, "Email berhasil dikirim"
            
        except Exception as e:
            return False, f"Gagal mengirim email: {str(e)}"



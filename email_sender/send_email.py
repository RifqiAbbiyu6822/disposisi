# email_sender/send_email.py

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from pathlib import Path
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Import from other files in the same package
from . import config
from .template_handler import LOGO_PATH

class EmailSender:
    def __init__(self):
        self.sender_email = config.EMAIL_HOST_USER
        self.sender_password = config.EMAIL_HOST_PASSWORD
        self.sheets_service = self._get_sheets_service()

    def _get_sheets_service(self):
        """Initialize Google Sheets service with admin credentials"""
        try:
            creds = Credentials.from_service_account_file(
                config.ADMIN_CREDENTIALS_FILE,
                scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
            )
            return build('sheets', 'v4', credentials=creds)
        except FileNotFoundError:
            print(f"Error: Credentials file not found at '{config.ADMIN_CREDENTIALS_FILE}'")
            return None
        except Exception as e:
            print(f"Error initializing sheets service: {str(e)}")
            return None

    def get_recipient_email(self, position):
        """Fetch email address from admin spreadsheet based on position"""
        if not self.sheets_service:
            return None, "Google Sheets service not available."
        try:
            range_name = 'Sheet1!A2:B'
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=config.ADMIN_SHEET_ID,
                range=range_name
            ).execute()
            values = result.get('values', [])
            if not values:
                return None, f"No data found in spreadsheet for position: {position}"

            for row in values:
                if len(row) >= 2 and row[0].strip().lower() == position.lower():
                    email = row[1].strip()
                    if '@' in email and '.' in email:
                        return email, f"Email found: {email}"
                    else:
                        return None, f"Invalid email format for {position}: {email}"
            return None, f"Email not found for position: {position}"
        except Exception as e:
            return None, f"Error reading email data from sheet: {str(e)}"

    def send_disposisi_email(self, recipients, subject, html_body):
        """
        Sends a disposition email with an embedded logo to a list of recipients.
        """
        if not self.sender_email or not self.sender_password:
            return False, "Sender email configuration is missing."

        # Use 'related' to embed images in the email body
        msg = MIMEMultipart('related')
        msg['Subject'] = subject
        msg['From'] = self.sender_email
        msg['To'] = ", ".join(recipients)

        # Attach the HTML part
        msg.attach(MIMEText(html_body, 'html'))

        # Attach the logo image if it exists
        if LOGO_PATH.exists():
            try:
                with open(LOGO_PATH, 'rb') as f:
                    logo_data = f.read()
                logo = MIMEImage(logo_data)
                # This Content-ID must match the 'cid:' in the HTML template's <img> tag
                logo.add_header('Content-ID', '<logo>')
                msg.attach(logo)
            except Exception as e:
                print(f"Warning: Could not read or attach logo file. Error: {e}")

        # Send the email using a secure connection
        try:
            with smtplib.SMTP(config.EMAIL_HOST, config.EMAIL_PORT) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.sendmail(self.sender_email, recipients, msg.as_string())
            return True, f"Email sent successfully to: {', '.join(recipients)}"
        except smtplib.SMTPAuthenticationError:
            return False, "Authentication failed. Check your email and App Password in config.py."
        except Exception as e:
            return False, f"Failed to send email: {str(e)}"
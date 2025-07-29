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
        
        # Define the cell mappings for each position
        # Based on the admin sheet structure shown in the screenshot
        self.position_cell_mapping = {
            "Direktur Utama": "B2",
            "Direktur Keuangan": "B3", 
            "Direktur Teknik": "B4",
            "GM Keuangan & Administrasi": "B5",
            "GM Operasional & Pemeliharaan": "B6",
            "Manager": "B7"
        }

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
        
        # Get the cell reference for this position
        cell_ref = self.position_cell_mapping.get(position)
        if not cell_ref:
            return None, f"Position '{position}' not found in mapping"
        
        try:
            # Read the specific cell
            range_name = f'Sheet1!{cell_ref}'
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=config.ADMIN_SHEET_ID,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            if not values or not values[0]:
                return None, f"No email found in cell {cell_ref} for position: {position}"
            
            email = values[0][0].strip()
            
            # Validate email format
            if '@' in email and '.' in email.split('@')[-1]:
                return email, f"Email found: {email}"
            else:
                return None, f"Invalid email format for {position}: {email}"
                
        except Exception as e:
            return None, f"Error reading email data from sheet: {str(e)}"

    def get_all_position_emails(self):
        """Fetch all position emails from the admin sheet"""
        if not self.sheets_service:
            return {}, "Google Sheets service not available."
        
        emails = {}
        errors = []
        
        for position, cell_ref in self.position_cell_mapping.items():
            email, msg = self.get_recipient_email(position)
            if email:
                emails[position] = email
            else:
                errors.append(f"{position}: {msg}")
        
        return emails, errors

    def send_disposisi_email(self, recipients, subject, html_body):
        """
        Sends a disposition email with an embedded logo to a list of recipients.
        
        Args:
            recipients: List of email addresses
            subject: Email subject
            html_body: HTML content of the email
        """
        if not self.sender_email or not self.sender_password:
            return False, "Sender email configuration is missing."

        # Validate recipients
        if not recipients:
            return False, "No recipients provided."
        
        # Remove duplicates and empty emails
        valid_recipients = list(set([r.strip() for r in recipients if r and '@' in r]))
        
        if not valid_recipients:
            return False, "No valid email addresses in recipient list."

        # Use 'related' to embed images in the email body
        msg = MIMEMultipart('related')
        msg['Subject'] = subject
        msg['From'] = self.sender_email
        msg['To'] = ", ".join(valid_recipients)

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
                server.sendmail(self.sender_email, valid_recipients, msg.as_string())
            return True, f"Email sent successfully to: {', '.join(valid_recipients)}"
        except smtplib.SMTPAuthenticationError:
            return False, "Authentication failed. Check your email and App Password in config.py."
        except Exception as e:
            return False, f"Failed to send email: {str(e)}"

    def send_disposisi_to_positions(self, positions, subject, html_body):
        """
        Send disposisi email to specific positions by looking up their emails
        
        Args:
            positions: List of position names (e.g., ["Direktur Utama", "Manager"])
            subject: Email subject
            html_body: HTML content
            
        Returns:
            tuple: (success, message, details)
        """
        if not positions:
            return False, "No positions specified", []
        
        # Fetch emails for the positions
        recipient_emails = []
        failed_lookups = []
        successful_lookups = []
        
        for position in positions:
            email, msg = self.get_recipient_email(position)
            if email:
                recipient_emails.append(email)
                successful_lookups.append(f"{position}: {email}")
            else:
                failed_lookups.append(f"{position} - {msg}")
        
        # Prepare detailed results
        details = {
            'successful_lookups': successful_lookups,
            'failed_lookups': failed_lookups,
            'emails_to_send': recipient_emails
        }
        
        # Check if we have any valid emails
        if not recipient_emails:
            return False, "No valid email addresses found for the selected positions", details
        
        # Send the email
        success, send_msg = self.send_disposisi_email(recipient_emails, subject, html_body)
        
        if success:
            message = f"Email sent to {len(recipient_emails)} recipient(s): {', '.join(recipient_emails)}"
            if failed_lookups:
                message += f"\n\nNote: Could not find emails for: {', '.join(failed_lookups)}"
        else:
            message = send_msg
        
        return success, message, details
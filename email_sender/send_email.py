# email_sender/send_email.py - ENHANCED VERSION with Senior Officer Support

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from pathlib import Path
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os

# Import from other files in the same package
try:
    from . import config
    from .template_handler import LOGO_PATH
except ImportError:
    # Fallback for direct execution
    import config
    from template_handler import LOGO_PATH

class EmailSender:
    def __init__(self):
        self.sender_email = config.EMAIL_HOST_USER
        self.sender_password = config.EMAIL_HOST_PASSWORD
        
        # Initialize admin_sheet_id BEFORE using it
        self.admin_sheet_id = config.ADMIN_SHEET_ID
        
        # Initialize sheets service
        self.sheets_service = self._get_sheets_service()
        
        # ENHANCED: Position to cell mapping including senior officers
        self.position_cell_mapping = {
            # Original positions
            "Direktur Utama": "B2",
            "Direktur Keuangan": "B3", 
            "Direktur Teknik": "B4",
            "GM Keuangan & Administrasi": "B5",
            "GM Operasional & Pemeliharaan": "B6",
            "Manager Pemeliharaan": "B7",
            "Manager Operasional": "B8",
            "Manager Administrasi": "B9",
            "Manager Keuangan": "B10",
            
            # ENHANCED: Senior Officers - Add new rows in the admin sheet
            "Senior Officer Pemeliharaan 1": "B11",
            "Senior Officer Pemeliharaan 2": "B12",
            "Senior Officer Operasional 1": "B13",
            "Senior Officer Operasional 2": "B14",
            "Senior Officer Administrasi 1": "B15",
            "Senior Officer Administrasi 2": "B16",
            "Senior Officer Keuangan 1": "B17",
            "Senior Officer Keuangan 2": "B18"
        }

    def _get_sheets_service(self):
        """Initialize Google Sheets service with admin credentials"""
        try:
            # Use absolute path for credentials
            current_dir = Path(__file__).parent.parent
            credentials_path = current_dir / 'credentials' / 'credentials.json'
            
            if not credentials_path.exists():
                # Try alternative paths
                alt_paths = [
                    current_dir / 'credentials.json',
                    Path('credentials/credentials.json'),
                    Path('credentials.json')
                ]
                
                for alt_path in alt_paths:
                    if alt_path.exists():
                        credentials_path = alt_path
                        break
                else:
                    print(f"Error: No credentials.json found. Tried paths:")
                    print(f"  - {current_dir / 'credentials' / 'credentials.json'}")
                    for alt_path in alt_paths:
                        print(f"  - {alt_path}")
                    return None
            
            print(f"Using credentials from: {credentials_path}")
            
            creds = Credentials.from_service_account_file(
                str(credentials_path),
                scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
            )
            service = build('sheets', 'v4', credentials=creds)
            
            # Test the connection
            try:
                sheet_metadata = service.spreadsheets().get(
                    spreadsheetId=self.admin_sheet_id
                ).execute()
                print(f"✓ Connected to admin sheet: {sheet_metadata.get('properties', {}).get('title', 'Unknown')}")
                return service
            except Exception as e:
                print(f"Error: Cannot access admin sheet {self.admin_sheet_id}: {e}")
                return None
                
        except FileNotFoundError as e:
            print(f"Error: Credentials file not found: {e}")
            return None
        except Exception as e:
            print(f"Error initializing sheets service: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def get_recipient_email(self, position):
        """ENHANCED: Fetch email address from admin spreadsheet based on position (including senior officers)"""
        if not self.sheets_service:
            return None, "Google Sheets service not available. Check credentials/credentials.json file."
        
        # Get the cell reference for this position
        cell_ref = self.position_cell_mapping.get(position)
        if not cell_ref:
            return None, f"Position '{position}' not found in mapping"
        
        try:
            # Read the specific cell from admin sheet
            range_name = f'Sheet1!{cell_ref}'
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=self.admin_sheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            if not values or not values[0] or not values[0][0]:
                return None, f"No email found in cell {cell_ref} for position: {position}"
            
            email = str(values[0][0]).strip()
            
            # Validate email format
            if '@' in email and '.' in email.split('@')[-1] and len(email) > 5:
                print(f"✓ Found email for {position}: {email}")
                return email, f"Email found: {email}"
            else:
                return None, f"Invalid email format for {position}: {email}"
                
        except Exception as e:
            error_msg = f"Error reading email data from admin sheet: {str(e)}"
            print(f"[EmailSender] {error_msg}")
            return None, error_msg

    def get_all_position_emails(self):
        """ENHANCED: Fetch all position emails including senior officers from the admin sheet"""
        if not self.sheets_service:
            return {}, ["Google Sheets service not available"]
        
        emails = {}
        errors = []
        
        for position, cell_ref in self.position_cell_mapping.items():
            email, msg = self.get_recipient_email(position)
            if email:
                emails[position] = email
            else:
                errors.append(f"{position}: {msg}")
        
        return emails, errors

    def send_disposisi_email(self, recipients, subject, html_body, pdf_attachment=None):
        """
        ENHANCED: Sends a disposition email with an embedded logo to a list of recipients (including senior officers).
        
        Args:
            recipients: List of email addresses or single email address
            subject: Email subject
            html_body: HTML content of the email
            pdf_attachment: Path to PDF file to attach (optional)
        """
        if not self.sender_email or not self.sender_password:
            return False, "Sender email configuration is missing. Check config.py file."

        # Ensure recipients is a list
        if isinstance(recipients, str):
            recipients = [recipients]
            
        # Validate recipients
        if not recipients:
            return False, "No recipients provided."
        
        # Remove duplicates and empty emails
        valid_recipients = []
        for r in recipients:
            email = str(r).strip()
            if email and '@' in email and '.' in email.split('@')[-1]:
                if email not in valid_recipients:
                    valid_recipients.append(email)
        
        if not valid_recipients:
            return False, "No valid email addresses in recipient list."

        print(f"Sending email to: {valid_recipients}")

        try:
            # Use 'related' to embed images in the email body
            msg = MIMEMultipart('related')
            msg['Subject'] = subject
            msg['From'] = self.sender_email
            msg['To'] = ", ".join(valid_recipients)

            # Attach the HTML part
            html_part = MIMEText(html_body, 'html', 'utf-8')
            msg.attach(html_part)

            # Attach the logo image if it exists
            if LOGO_PATH.exists():
                try:
                    with open(LOGO_PATH, 'rb') as f:
                        logo_data = f.read()
                    from email.mime.image import MIMEImage
                    logo = MIMEImage(logo_data)
                    # This Content-ID must match the 'cid:' in the HTML template's <img> tag
                    logo.add_header('Content-ID', '<logo>')
                    msg.attach(logo)
                except Exception as e:
                    print(f"Warning: Could not attach logo file: {e}")

            # Attach PDF if provided
            if pdf_attachment and Path(pdf_attachment).exists():
                try:
                    with open(pdf_attachment, 'rb') as f:
                        pdf_data = f.read()
                    
                    pdf_part = MIMEApplication(pdf_data, _subtype='pdf')
                    pdf_filename = Path(pdf_attachment).name
                    pdf_part.add_header('Content-Disposition', 'attachment', filename=pdf_filename)
                    msg.attach(pdf_part)
                    print(f"✓ Attached PDF: {pdf_filename}")
                except Exception as e:
                    print(f"Warning: Could not attach PDF file: {e}")

            # Send the email using SMTP
            print(f"Connecting to SMTP server {config.EMAIL_HOST}:{config.EMAIL_PORT}...")
            
            if config.EMAIL_PORT == 465:
                # Use SSL
                with smtplib.SMTP_SSL(config.EMAIL_HOST, config.EMAIL_PORT) as server:
                    server.login(self.sender_email, self.sender_password)
                    server.sendmail(self.sender_email, valid_recipients, msg.as_string())
            else:
                # Use TLS
                with smtplib.SMTP(config.EMAIL_HOST, config.EMAIL_PORT) as server:
                    server.starttls()
                    server.login(self.sender_email, self.sender_password)
                    server.sendmail(self.sender_email, valid_recipients, msg.as_string())
            
            return True, f"Email sent successfully to: {', '.join(valid_recipients)}"
            
        except smtplib.SMTPAuthenticationError as e:
            return False, f"Authentication failed. Check your email and App Password in config.py. Error: {e}"
        except smtplib.SMTPException as e:
            return False, f"SMTP error occurred: {e}"
        except Exception as e:
            import traceback
            traceback.print_exc()
            return False, f"Failed to send email: {str(e)}"

    def send_disposisi_to_positions(self, positions, subject, html_body, pdf_attachment=None):
        """
        ENHANCED: Send disposisi email to specific positions including senior officers by looking up their emails
        
        Args:
            positions: List of position names (e.g., ["Direktur Utama", "Manager", "Senior Officer Pemeliharaan 1"])
            subject: Email subject
            html_body: HTML content
            pdf_attachment: Path to PDF file to attach (optional)
            
        Returns:
            tuple: (success, message, details)
        """
        if not positions:
            return False, "No positions specified", []
        
        # Fetch emails for the positions
        recipient_emails = []
        failed_lookups = []
        successful_lookups = []
        
        print(f"[DEBUG] Looking up emails for positions: {positions}")
        
        for position in positions:
            print(f"[DEBUG] Looking up email for: {position}")
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
        success, send_msg = self.send_disposisi_email(
            recipient_emails, 
            subject, 
            html_body, 
            pdf_attachment
        )
        
        if success:
            message = f"Email sent to {len(recipient_emails)} recipient(s): {', '.join(recipient_emails)}"
            if failed_lookups:
                # Use abbreviation for manager in message
                abbreviation_map = {
                    "Manager Pemeliharaan": "pml",
                    "Manager Operasional": "ops",
                    "Manager Administrasi": "adm",
                    "Manager Keuangan": "keu"
                }
                
                # Convert failed lookups to abbreviation for display
                display_failed_positions = []
                for f in failed_lookups:
                    position = f.split(' - ')[0]
                    for full_name, abbrev in abbreviation_map.items():
                        if full_name in position:
                            position = position.replace(full_name, f"Manager {abbrev}")
                            break
                    display_failed_positions.append(position)
                
                message += f"\n\nNote: Could not find emails for: {', '.join(display_failed_positions)}"
        else:
            message = send_msg
        
        return success, message, details

    def test_connection(self):
        """ENHANCED: Test the email connection and admin sheet access including senior officers"""
        results = {
            'smtp_connection': False,
            'sheets_connection': False,
            'admin_sheet_access': False,
            'email_data': {},
            'errors': []
        }
        
        # Test SMTP connection
        try:
            if config.EMAIL_PORT == 465:
                with smtplib.SMTP_SSL(config.EMAIL_HOST, config.EMAIL_PORT) as server:
                    server.login(self.sender_email, self.sender_password)
            else:
                with smtplib.SMTP(config.EMAIL_HOST, config.EMAIL_PORT) as server:
                    server.starttls()
                    server.login(self.sender_email, self.sender_password)
            results['smtp_connection'] = True
            print("✓ SMTP connection successful")
        except Exception as e:
            results['errors'].append(f"SMTP connection failed: {e}")
            print(f"✗ SMTP connection failed: {e}")
        
        # Test Google Sheets connection
        if self.sheets_service:
            results['sheets_connection'] = True
            print("✓ Google Sheets service initialized")
            
            # Test admin sheet access
            try:
                emails, errors = self.get_all_position_emails()
                if emails:
                    results['admin_sheet_access'] = True
                    results['email_data'] = emails
                    print(f"✓ Found {len(emails)} email addresses in admin sheet (including senior officers)")
                if errors:
                    results['errors'].extend(errors)
            except Exception as e:
                results['errors'].append(f"Admin sheet access failed: {e}")
                print(f"✗ Admin sheet access failed: {e}")
        else:
            results['errors'].append("Google Sheets service not initialized")
            print("✗ Google Sheets service not initialized")
        
        return results


# Test function for the enhanced system
def test_enhanced_email_system():
    """Test the enhanced email system with senior officer support"""
    print("Testing Enhanced Email System with Senior Officer Support")
    print("=" * 60)
    
    email_sender = EmailSender()
    
    # Test connection
    results = email_sender.test_connection()
    
    print("\nConnection Test Summary:")
    print(f"{'SMTP Connection:':<25} {'✓ OK' if results['smtp_connection'] else '✗ FAILED'}")
    print(f"{'Google Sheets:':<25} {'✓ OK' if results['sheets_connection'] else '✗ FAILED'}")
    print(f"{'Admin Sheet Access:':<25} {'✓ OK' if results['admin_sheet_access'] else '✗ FAILED'}")
    
    if results['email_data']:
        print(f"\nFound emails for {len(results['email_data'])} positions:")
        
        # Group by category
        managers = []
        senior_officers = []
        others = []
        
        for position, email in results['email_data'].items():
            if position.startswith("Senior Officer"):
                senior_officers.append((position, email))
            elif position.startswith("Manager"):
                managers.append((position, email))
            else:
                others.append((position, email))
        
        # Display by category
        if others:
            print("\n  Directors & GM:")
            for position, email in others:
                print(f"    • {position}: {email}")
        
        if managers:
            print("\n  Managers:")
            for position, email in managers:
                print(f"    • {position}: {email}")
        
        if senior_officers:
            print("\n  Senior Officers:")
            for position, email in senior_officers:
                print(f"    • {position}: {email}")
    
    if results['errors']:
        print(f"\nErrors encountered:")
        for error in results['errors']:
            print(f"  • {error}")
    
    # Test senior officer email lookup
    print(f"\nTesting Senior Officer Email Lookup:")
    print("-" * 40)
    
    test_senior_positions = [
        "Senior Officer Pemeliharaan 1",
        "Senior Officer Operasional 2",
        "Senior Officer Keuangan 1"
    ]
    
    for position in test_senior_positions:
        email, msg = email_sender.get_recipient_email(position)
        if email:
            print(f"✓ {position}: {email}")
        else:
            print(f"✗ {position}: {msg}")
    
    return results


if __name__ == "__main__":
    test_enhanced_email_system()
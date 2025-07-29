# email_sender/send_email.py - FIXED VERSION

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
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
        
        # FIX: Use the admin sheet ID from config
        self.admin_sheet_id = config.ADMIN_SHEET_ID
        
        # Position to cell mapping for reading emails from admin sheet
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
            # FIX: Use absolute path for credentials
            import os
            current_dir = Path(__file__).parent.parent
            credentials_path = current_dir / 'admin' / 'credentials.json'
            
            if not credentials_path.exists():
                # Try alternative paths
                alt_paths = [
                    current_dir / 'credentials.json',
                    Path('admin/credentials.json'),
                    Path('credentials.json')
                ]
                
                for alt_path in alt_paths:
                    if alt_path.exists():
                        credentials_path = alt_path
                        break
                else:
                    print(f"Error: No credentials.json found. Tried paths:")
                    print(f"  - {current_dir / 'admin' / 'credentials.json'}")
                    for alt_path in alt_paths:
                        print(f"  - {alt_path}")
                    return None
            
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
            return None

    def get_recipient_email(self, position):
        """Fetch email address from admin spreadsheet based on position"""
        if not self.sheets_service:
            return None, "Google Sheets service not available. Check credentials.json file."
        
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
                return email, f"Email found: {email}"
            else:
                return None, f"Invalid email format for {position}: {email}"
                
        except Exception as e:
            error_msg = f"Error reading email data from admin sheet: {str(e)}"
            print(f"[EmailSender] {error_msg}")
            return None, error_msg

    def get_all_position_emails(self):
        """Fetch all position emails from the admin sheet"""
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
        Sends a disposition email with an embedded logo to a list of recipients.
        
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
                except Exception as e:
                    print(f"Warning: Could not attach PDF file: {e}")

            # Send the email using SMTP
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
            return False, f"Failed to send email: {str(e)}"

    def send_disposisi_to_positions(self, positions, subject, html_body, pdf_attachment=None):
        """
        Send disposisi email to specific positions by looking up their emails
        
        Args:
            positions: List of position names (e.g., ["Direktur Utama", "Manager"])
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
        success, send_msg = self.send_disposisi_email(
            recipient_emails, 
            subject, 
            html_body, 
            pdf_attachment
        )
        
        if success:
            message = f"Email sent to {len(recipient_emails)} recipient(s): {', '.join(recipient_emails)}"
            if failed_lookups:
                message += f"\n\nNote: Could not find emails for: {', '.join([f.split(' - ')[0] for f in failed_lookups])}"
        else:
            message = send_msg
        
        return success, message, details

    def test_connection(self):
        """Test the email connection and admin sheet access"""
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
        except Exception as e:
            results['errors'].append(f"SMTP connection failed: {e}")
        
        # Test Google Sheets connection
        if self.sheets_service:
            results['sheets_connection'] = True
            
            # Test admin sheet access
            try:
                emails, errors = self.get_all_position_emails()
                if emails:
                    results['admin_sheet_access'] = True
                    results['email_data'] = emails
                if errors:
                    results['errors'].extend(errors)
            except Exception as e:
                results['errors'].append(f"Admin sheet access failed: {e}")
        else:
            results['errors'].append("Google Sheets service not initialized")
        
        return results


# Fixed export_utils.py function for sending email
def send_email_with_disposisi_fixed(self, selected_positions):
    """
    FIXED VERSION: Mengirimkan email disposisi ke posisi yang dipilih 
    dengan mengambil alamat email dari admin spreadsheet.
    """
    from tkinter import messagebox
    import tempfile
    import os
    import traceback
    from email_sender.send_email import EmailSender
    from email_sender.template_handler import render_email_template
    from pdf_output import save_form_to_pdf, merge_pdfs
    from datetime import datetime
    from disposisi_app.views.components.email_error_handler import handle_email_error

    # Update status
    self.update_status("Menyiapkan pengiriman email...")
    
    # Collect form data
    from disposisi_app.views.components.export_utils import collect_form_data_safely
    data = collect_form_data_safely(self)
    
    if not data.get("no_surat", "").strip():
        messagebox.showerror("Validasi Error", "No. Surat tidak boleh kosong untuk mengirim email.")
        return

    # Generate PDF for attachment
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf_path = temp_pdf.name
        
        save_form_to_pdf(temp_pdf_path, data)
        
        final_pdf_path = temp_pdf_path
        if hasattr(self, 'pdf_attachments') and self.pdf_attachments:
            merged_pdf_path = os.path.join(tempfile.gettempdir(), f"merged_{os.path.basename(temp_pdf_path)}")
            pdf_list = [temp_pdf_path] + self.pdf_attachments
            merge_pdfs(pdf_list, merged_pdf_path)
            final_pdf_path = merged_pdf_path

    except Exception as e:
        messagebox.showerror("PDF Generation Error", f"Gagal membuat PDF untuk email: {e}")
        traceback.print_exc()
        return
    
    # Initialize EmailSender
    try:
        email_sender = EmailSender()
        
        # Test connection first
        if not email_sender.sheets_service:
            error_msg = ("Google Sheets tidak dapat diakses.\n\n"
                        "Pastikan:\n"
                        "1. File admin/credentials.json tersedia\n"
                        "2. Sheet admin dengan email sudah dibuat\n"
                        "3. Koneksi internet tersedia")
            handle_email_error(self, error_msg)
            return
            
    except Exception as e:
        error_msg = f"Gagal menginisialisasi email sender: {str(e)}"
        handle_email_error(self, error_msg)
        return
    
    # Get emails for selected positions
    recipient_emails = []
    failed_lookups = []
    successful_lookups = []
    
    self.update_status("Mencari alamat email...")
    
    for position in selected_positions:
        email, msg = email_sender.get_recipient_email(position)
        if email:
            recipient_emails.append(email)
            successful_lookups.append(f"{position}: {email}")
        else:
            failed_lookups.append(f"{position} ({msg})")
    
    # Show warnings for failed lookups
    if failed_lookups:
        warning_msg = ("Tidak dapat menemukan alamat email untuk posisi berikut:\n" + 
                      "\n".join([f"• {lookup}" for lookup in failed_lookups]))
        
        if len(failed_lookups) == len(selected_positions):
            # All positions failed - show error dialog
            full_error = (warning_msg + 
                         "\n\nPastikan admin sudah mengisi email untuk posisi-posisi tersebut "
                         "di spreadsheet admin.")
            handle_email_error(self, full_error)
            return
        else:
            # Some positions failed - show warning but continue
            messagebox.showwarning("Email Tidak Lengkap", warning_msg, parent=self)
    
    if not recipient_emails:
        messagebox.showerror("Gagal", "Tidak ada alamat email yang valid untuk dikirimi.", parent=self)
        return
    
    # Prepare email template data
    template_data = {
        'nomor_surat': data.get('no_surat', 'N/A'),
        'nama_pengirim': data.get('asal_surat', 'N/A'),
        'perihal': data.get('perihal', 'N/A'),
        'tanggal': datetime.now().strftime('%d %B %Y'),
        'klasifikasi': [],
        'instruksi_list': [],
        'tahun': datetime.now().year
    }
    
    # Add classification
    if data.get('rahasia', 0):
        template_data['klasifikasi'].append("RAHASIA")
    if data.get('penting', 0):
        template_data['klasifikasi'].append("PENTING")
    if data.get('segera', 0):
        template_data['klasifikasi'].append("SEGERA")
    
    # Add instructions from checkboxes
    instruksi_mapping = [
        ("ketahui_file", "Ketahui & File"),
        ("proses_selesai", "Proses Selesai"),
        ("teliti_pendapat", "Teliti & Pendapat"),
        ("buatkan_resume", "Buatkan Resume"),
        ("edarkan", "Edarkan"),
        ("sesuai_disposisi", "Sesuai Disposisi"),
        ("bicarakan_saya", "Bicarakan dengan Saya")
    ]
    
    for key, label in instruksi_mapping:
        if data.get(key, 0):
            template_data['instruksi_list'].append(label)
    
    # Add detailed instructions from table
    if 'isi_instruksi' in data and data['isi_instruksi']:
        for instr in data['isi_instruksi']:
            if instr.get('instruksi', '').strip():
                posisi = instr.get('posisi', '')
                instruksi_text = instr.get('instruksi', '')
                tanggal = instr.get('tanggal', '')
                
                instr_line = f"{posisi}: {instruksi_text}"
                if tanggal:
                    instr_line += f" (Tanggal: {tanggal})"
                template_data['instruksi_list'].append(instr_line)
    
    # Add additional instructions
    if data.get("bicarakan_dengan", "").strip():
        template_data['instruksi_list'].append(f"Bicarakan dengan: {data['bicarakan_dengan']}")
    
    if data.get("teruskan_kepada", "").strip():
        template_data['instruksi_list'].append(f"Teruskan kepada: {data['teruskan_kepada']}")
    
    if data.get("harap_selesai_tgl", "").strip():
        template_data['instruksi_list'].append(f"Harap diselesaikan tanggal: {data['harap_selesai_tgl']}")
    
    # Render email template
    html_content = render_email_template(template_data)
    subject = f"Disposisi Surat: {data.get('perihal', 'N/A')}"
    
    # Send email
    self.update_status(f"Mengirim email ke {len(recipient_emails)} penerima...")
    
    success, message = email_sender.send_disposisi_email(
        recipient_emails, 
        subject, 
        html_content,
        pdf_attachment=final_pdf_path
    )
    
    # Show results
    if success:
        self.update_status("Email berhasil dikirim!")
        success_msg = f"Email berhasil dikirim ke:\n{chr(10).join([f'• {email}' for email in recipient_emails])}"
        
        if successful_lookups:
            success_msg += f"\n\nDetail penerima:\n{chr(10).join([f'• {lookup}' for lookup in successful_lookups])}"
        
        messagebox.showinfo("Email Terkirim", success_msg, parent=self)
    else:
        self.update_status("Gagal mengirim email.")
        handle_email_error(self, f"Gagal mengirim email: {message}")
    
    # Clean up temporary files
    try:
        if 'temp_pdf_path' in locals() and os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if 'final_pdf_path' in locals() and final_pdf_path != temp_pdf_path and os.path.exists(final_pdf_path):
            os.remove(final_pdf_path)
    except Exception as e:
        print(f"Warning: Could not remove temporary files: {e}")


# Test script to verify the fix
def test_email_system():
    """Test the fixed email system"""
    print("Testing Fixed Email System")
    print("=" * 50)
    
    email_sender = EmailSender()
    
    # Test connection
    results = email_sender.test_connection()
    
    print("Connection Test Results:")
    print(f"✓ SMTP Connection: {'OK' if results['smtp_connection'] else 'FAILED'}")
    print(f"✓ Google Sheets: {'OK' if results['sheets_connection'] else 'FAILED'}")
    print(f"✓ Admin Sheet Access: {'OK' if results['admin_sheet_access'] else 'FAILED'}")
    
    if results['email_data']:
        print(f"\nFound emails for {len(results['email_data'])} positions:")
        for position, email in results['email_data'].items():
            print(f"  • {position}: {email}")
    
    if results['errors']:
        print(f"\nErrors encountered:")
        for error in results['errors']:
            print(f"  • {error}")
    
    return results

if __name__ == "__main__":
    test_email_system()
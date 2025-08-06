
# Script to test reading emails from the admin Google Sheet

import sys
sys.path.append('.')  # Add current directory to path

from email_sender.send_email import EmailSender

def test_email_reading():
    """Test the email reading functionality"""
    print("Testing Email Reading from Admin Google Sheet")
    print("=" * 50)
    
    # Initialize EmailSender
    email_sender = EmailSender()
    
    if not email_sender.sheets_service:
        print("ERROR: Could not initialize Google Sheets service!")
        print("Please check:")
        print("1. credentials.json file exists in credentials/ directory")
        print("2. The file has proper permissions")
        return
    
    print("✓ Google Sheets service initialized successfully")
    print()
    
    # Test individual position lookups
    print("Testing individual position lookups:")
    print("-" * 30)
    
    test_positions = [
        "Direktur Utama",
        "Direktur Keuangan", 
        "Direktur Teknik",
        "GM Keuangan & Administrasi",
        "GM Operasional & Pemeliharaan",
        "Manager"
    ]
    
    for position in test_positions:
        email, msg = email_sender.get_recipient_email(position)
        if email:
            print(f"✓ {position}: {email}")
        else:
            print(f"✗ {position}: {msg}")
    
    print()
    
    # Test getting all emails at once
    print("Testing batch email retrieval:")
    print("-" * 30)
    
    all_emails, errors = email_sender.get_all_position_emails()
    
    if all_emails:
        print(f"✓ Successfully retrieved {len(all_emails)} email(s):")
        for position, email in all_emails.items():
            print(f"  - {position}: {email}")
    
    if errors:
        print(f"\n✗ Errors encountered:")
        for error in errors:
            print(f"  - {error}")
    
    print()
    
    # Test sending to multiple positions
    print("Testing position-based email sending (dry run):")
    print("-" * 30)
    
    test_positions_to_send = ["Direktur Utama", "Manager"]
    print(f"Attempting to prepare email for: {', '.join(test_positions_to_send)}")
    
    # Just test the lookup, don't actually send
    recipient_emails = []
    for position in test_positions_to_send:
        email, _ = email_sender.get_recipient_email(position)
        if email:
            recipient_emails.append(email)
    
    if recipient_emails:
        print(f"✓ Would send to: {', '.join(recipient_emails)}")
    else:
        print("✗ No valid emails found for selected positions")

if __name__ == "__main__":
    test_email_reading()
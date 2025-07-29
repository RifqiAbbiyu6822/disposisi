# email_sender/config.py

# --- Email Configuration ---
EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
EMAIL_USE_TLS = True

# IMPORTANT: Use the Google Account email and the 16-character App Password.
# How to get an App Password: https://support.google.com/accounts/answer/185833
EMAIL_HOST_USER = 'magangjcc@gmail.com'
EMAIL_HOST_PASSWORD = 'hnfv azrs snou aays' 

# --- Google Sheets Configuration for Admin Panel ---
# Used by EmailSender to find recipient emails.
ADMIN_SHEET_ID = '1LoAzVPBMJo08uPHR7MdzGEXmaoFXMrTSKS0vl099qrU'
ADMIN_CREDENTIALS_FILE = 'admin/credentials.json'
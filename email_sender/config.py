import os

# Email configuration - support multiple sources
EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
EMAIL_USE_TLS = True
EMAIL_USE_SSL = False

EMAIL_HOST_USER = 'magangjcc@gmail.com'
EMAIL_HOST_PASSWORD = 'cygh bhpt fqnb yexb'
# If still empty, try to get from email_sender config
if not EMAIL_HOST_USER or not EMAIL_HOST_PASSWORD:
    try:
        from email_sender.config import SENDER_EMAIL, SENDER_PASSWORD
        EMAIL_HOST_USER = EMAIL_HOST_USER or SENDER_EMAIL
        EMAIL_HOST_PASSWORD = EMAIL_HOST_PASSWORD or SENDER_PASSWORD
    except:
        pass

# Admin sheet configuration for email recipients
ADMIN_SHEET_ID = '1LoAzVPBMJo08uPHR7MdzGEXmaoFXMrTSKS0vl099qrU'
ADMIN_CREDENTIALS_FILE = 'admin/credentials.json'
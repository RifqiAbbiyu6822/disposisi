import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Email configuration - support multiple sources
EMAIL_HOST = os.getenv('EMAIL_HOST', 'smtp.gmail.com')
EMAIL_PORT = int(os.getenv('EMAIL_PORT', '587'))
EMAIL_USE_TLS = os.getenv('EMAIL_USE_TLS', 'True').lower() == 'true'
EMAIL_USE_SSL = os.getenv('EMAIL_USE_SSL', 'False').lower() == 'true'

# Email credentials - try multiple sources
EMAIL_HOST_USER = os.getenv('EMAIL_HOST_USER') or os.getenv('SENDER_EMAIL', '')
EMAIL_HOST_PASSWORD = os.getenv('EMAIL_HOST_PASSWORD') or os.getenv('SENDER_PASSWORD', '')

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
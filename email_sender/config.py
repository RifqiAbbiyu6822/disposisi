# email_sender/config.py

# --- Email Configuration ---
EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
EMAIL_USE_TLS = True

# IMPORTANT: Gunakan email Google dan App Password 16 karakter
# Cara mendapatkan App Password: https://support.google.com/accounts/answer/185833
EMAIL_HOST_USER = 'magangjcc@gmail.com'
EMAIL_HOST_PASSWORD = 'rbsc nsmd lshz wmdm'  # Menggunakan password dari .env

# --- Google Sheets Configuration for Admin Panel ---
# Menggunakan spreadsheet yang sama dengan yang ditampilkan dalam admin panel
ADMIN_SHEET_ID = '1LoAzVPBMJo08uPHR7MdzGEXmaoFXMrTSKS0vl099qrU'
ADMIN_CREDENTIALS_FILE = 'credentials/credentials.json'

# --- Fallback Configuration ---
# Pengaturan alternatif jika konfigurasi utama gagal
FALLBACK_EMAIL_HOST = 'smtp.gmail.com'
FALLBACK_EMAIL_PORT = 465
FALLBACK_EMAIL_USE_SSL = True
FALLBACK_EMAIL_HOST_USER = 'magangjcc@gmail.com'
FALLBACK_EMAIL_HOST_PASSWORD = 'rbsc nsmd lshz wmdm'  # Menggunakan password dari .env
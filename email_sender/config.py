"""
Konfigurasi untuk pengiriman email
"""

# Email server settings
EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 465
EMAIL_USE_SSL = True
SENDER_EMAIL=magangjcc@gmail.com
SENDER_PASSWORD=cygh bhpt fqnb yexb

# Template messages
EMAIL_SUCCESS = "Email berhasil dikirim ke {recipient}"
EMAIL_FAILED = "Gagal mengirim email ke {recipient}: {error}"
EMAIL_CONFIG_ERROR = "Konfigurasi email tidak valid. Periksa file .env"
EMAIL_VALIDATION_ERROR = "Format email tidak valid untuk {position}: {email}"
EMAIL_NOT_FOUND = "Email tidak ditemukan untuk posisi: {position}"
EMAIL_SHEET_ERROR = "Error membaca data dari spreadsheet: {error}"

# Spreadsheet configuration
RECIPIENT_SHEET_NAME = "Penerima"
RECIPIENT_RANGE = "A2:B"  # Skip header row
HEADER_RANGE = "A1:B1"

# Column indices in spreadsheet
POSITION_COL = 0
EMAIL_COL = 1

# Email template settings
DEFAULT_SUBJECT = "Disposisi Surat - {nomor_surat}"

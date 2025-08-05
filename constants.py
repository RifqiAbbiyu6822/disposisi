# Constants untuk aplikasi formulir disposisi
import logging
logging.basicConfig(level=logging.WARNING)

# Field labels untuk UI
FIELD_LABELS = {
    "id": "ID",
    "no_agenda": "No. Agenda",
    "no_surat": "No. Surat",
    "tgl_surat": "Tgl. Surat",
    "perihal": "Perihal",
    "asal_surat": "Asal Surat",
    "ditujukan": "Ditujukan",
    "rahasia": "Rahasia",
    "penting": "Penting",
    "segera": "Segera",
    "kode_klasifikasi": "Kode Klasifikasi",
    "indeks": "Indeks",
    "dir_utama": "Direktur Utama",
    "dir_keu": "Direktur Keuangan",
    "dir_teknik": "Direktur Teknik",
    "gm_keu": "GM Keuangan & Administrasi",
    "gm_ops": "GM Operasional & Pemeliharaan",
    "manager_pemeliharaan": "Manager Pemeliharaan",
    "manager_operasional": "Manager Operasional",
    "manager_administrasi": "Manager Administrasi",
    "manager_keuangan": "Manager Keuangan",
    "ketahui_file": "Ketahui & File",
    "proses_selesai": "Proses Selesai",
    "teliti_pendapat": "Teliti & Pendapat",
    "buatkan_resume": "Buatkan Resume",
    "edarkan": "Edarkan",
    "sesuai_disposisi": "Sesuai Disposisi",
    "bicarakan_saya": "Bicarakan dengan Saya",
    "bicarakan_dengan": "Bicarakan dengan",
    "teruskan_kepada": "Teruskan kepada",
    "harap_selesai_tgl": "Harap Selesai Tanggal",
    "isi_instruksi": "Isi Instruksi",
    "created_at": "Waktu Simpan",
    "version_label": "Versi"
}

# Semua field yang akan ditampilkan
ALL_FIELDS = [
    "id", "no_agenda", "no_surat", "tgl_surat", "perihal", "asal_surat",
    "ditujukan", "rahasia", "penting", "segera", "kode_klasifikasi",
    "indeks", "dir_utama", "dir_keu", "dir_teknik", "gm_keu", "gm_ops", 
    "manager_pemeliharaan", "manager_operasional", "manager_administrasi", "manager_keuangan",
    "ketahui_file", "proses_selesai", "teliti_pendapat", "buatkan_resume",
    "edarkan", "sesuai_disposisi", "bicarakan_saya", "bicarakan_dengan",
    "teruskan_kepada", "harap_selesai_tgl", "isi_instruksi", "created_at"
]

# Field yang berupa checkbox
CHECKBOX_FIELDS = [
    "rahasia", "penting", "segera", "dir_utama", "dir_keu", "dir_teknik", "gm_keu", "gm_ops", 
    "manager_pemeliharaan", "manager_operasional", "manager_administrasi", "manager_keuangan",
    "ketahui_file", "proses_selesai", "teliti_pendapat", "buatkan_resume", "edarkan", "sesuai_disposisi", "bicarakan_saya"
]

# Keyboard shortcuts untuk aplikasi
KEYBOARD_SHORTCUTS = {
    # File Operations
    "save_pdf": {"key": "Ctrl+S", "description": "Simpan ke PDF"},
    "export_excel": {"key": "Ctrl+E", "description": "Export ke Excel"},
    "exit_app": {"key": "Alt+F4", "description": "Keluar aplikasi"},

    # Navigation
    "next_widget": {"key": "Tab", "description": "Next widget"},
    "prev_widget": {"key": "Shift+Tab", "description": "Previous widget"},
    "form_tab": {"key": "Ctrl+1", "description": "Switch to Form tab"},
    "history_tab": {"key": "Ctrl+2", "description": "Switch to History tab"},

    # View
    "zoom_in": {"key": "Ctrl++", "description": "Zoom in"},
    "zoom_out": {"key": "Ctrl+-", "description": "Zoom out"},
    "reset_zoom": {"key": "Ctrl+0", "description": "Reset zoom"},
    "refresh": {"key": "F5", "description": "Refresh data"},
    "fullscreen": {"key": "F12", "description": "Toggle fullscreen"},

    # Help
    "show_help": {"key": "F1", "description": "Show keyboard shortcuts"},

    # History View Specific
    "focus_search": {"key": "Ctrl+F", "description": "Focus search field"},
    "delete_selected": {"key": "Delete", "description": "Delete selected items"},
    "invert_selection": {"key": "Ctrl+I", "description": "Invert selection"},
    "find_next": {"key": "F3", "description": "Find next"},
    "find_previous": {"key": "Shift+F3", "description": "Find previous"},
    "export_excel_history": {"key": "F7", "description": "Export to Excel"},

    # Navigation in History
    "navigate_up": {"key": "Up Arrow", "description": "Navigate up"},
    "navigate_down": {"key": "Down Arrow", "description": "Navigate down"},
    "page_up": {"key": "Page Up", "description": "Page up"},
    "page_down": {"key": "Page Down", "description": "Page down"},
    "go_to_top": {"key": "Home", "description": "Go to top"},
    "go_to_bottom": {"key": "End", "description": "Go to bottom"},

    # Windows Specific
    "emergency_exit": {"key": "Ctrl+Alt+Delete", "description": "Emergency exit"},
    "task_manager": {"key": "Ctrl+Shift+Escape", "description": "Task manager"}
}

# Touchpad gestures
TOUCHPAD_GESTURES = {
    "two_finger_scroll": "Vertical scrolling",
    "shift_scroll": "Horizontal scrolling",
    "double_tap": "Zoom in",
    "ctrl_drag": "Pinch gesture (zoom)",
    "drag_up": "Scroll up",
    "drag_down": "Scroll down"
}

# Help text untuk aplikasi
HELP_TEXT = """
Aplikasi Formulir Disposisi - Keyboard Shortcuts & Gestures

FILE OPERATIONS:
• Ctrl+S: Simpan ke PDF
• Ctrl+E: Export ke Excel
• Alt+F4: Keluar aplikasi

NAVIGATION:
• Tab: Next widget
• Shift+Tab: Previous widget
• Enter: Next widget
• Shift+Enter: Previous widget
• Ctrl+1: Switch to Form tab
• Ctrl+2: Switch to History tab

VIEW:
• Ctrl++: Zoom in
• Ctrl+-: Zoom out
• Ctrl+0: Reset zoom
• F5: Refresh data
• F12: Toggle fullscreen

HISTORY VIEW:
• Ctrl+F: Focus search field
• Delete: Delete selected items
• Ctrl+I: Invert selection
• F3: Find next
• Shift+F3: Find previous
• F7: Export to Excel
• Up/Down Arrow: Navigate rows
• Page Up/Down: Page navigation
• Home/End: Go to top/bottom

TOUCHPAD GESTURES:
• Two-finger scroll: Vertical scrolling
• Shift+scroll: Horizontal scrolling
• Double-tap: Zoom in
• Ctrl+drag: Pinch gesture (zoom)
• Drag up/down: Scroll up/down

WINDOWS SPECIFIC:
• Ctrl+Alt+Delete: Emergency exit
• Ctrl+Shift+Escape: Task manager

Tips:
- Use Tab/Shift+Tab to navigate between form fields
- Use arrow keys to navigate in history table
- Double-click items in history to open
- Use Ctrl+A to select all items in history
- Use Delete key to remove selected items
"""

# Status messages
STATUS_MESSAGES = {
    "ready": "Ready",
    "saving": "Saving...",
    "saved": "Saved successfully",
    "exporting": "Exporting...",
    "exported": "Exported successfully",
    "deleting": "Deleting...",
    "deleted": "Deleted successfully",
    "refreshing": "Refreshing...",
    "refreshed": "Refreshed successfully",
    "error": "Error occurred",
    "no_data": "No data available",
    "searching": "Searching...",
    "found": "Found results",
    "not_found": "No results found"
} 
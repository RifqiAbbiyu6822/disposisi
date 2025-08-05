import traceback
import logging
from google_sheets_connect import get_sheets_service, SHEET_ID, update_row_in_sheet

logging.basicConfig(level=logging.WARNING)

ENHANCED_HEADER = [
    "No. Agenda", "No. Surat", "Tgl. Surat", "Perihal", "Asal Surat", "Ditujukan", 
    "Klasifikasi", "Disposisi kepada", "Untuk Di :", "Selesai Tgl.", "Kode Klasifikasi", 
    "Tgl. Penerimaan", "Indeks", "Bicarakan dengan", "Teruskan kepada", "Harap Selesai Tanggal",
    "Direktur Utama Instruksi", "Direktur Utama Tanggal", "Direktur Keuangan Instruksi", "Direktur Keuangan Tanggal",
    "Direktur Teknik Instruksi", "Direktur Teknik Tanggal", "GM Keuangan & Administrasi Instruksi", "GM Keuangan & Administrasi Tanggal",
    "GM Operasional & Pemeliharaan Instruksi", "GM Operasional & Pemeliharaan Tanggal", 
    "Manager Pemeliharaan Instruksi", "Manager Pemeliharaan Tanggal",
    "Manager Operasional Instruksi", "Manager Operasional Tanggal",
    "Manager Administrasi Instruksi", "Manager Administrasi Tanggal",
    "Manager Keuangan Instruksi", "Manager Keuangan Tanggal"
]

def normalize_date_format(date_str):
    """Normalize date format with robust null handling"""
    if not date_str or str(date_str).strip() == "" or str(date_str).lower() in ['none', 'null']:
        return ""
    
    try:
        from datetime import datetime
        date_str = str(date_str).strip()
        
        # Common date formats to try
        formats = ["%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d.%m.%Y"]
        
        for fmt in formats:
            try:
                date_obj = datetime.strptime(date_str, fmt)
                return date_obj.strftime("%d-%m-%Y")
            except ValueError:
                continue
        
        # If no format matches, return as is (but cleaned)
        return date_str
        
    except Exception as e:
        print(f"[normalize_date_format] Warning: Could not parse date '{date_str}': {e}")
        return str(date_str) if date_str is not None else ""

def safe_get_value(data, key, default=""):
    """Safely get value from data dict with null handling"""
    try:
        value = data.get(key, default)
        if value is None:
            return default
        return str(value).strip() if value != default else default
    except:
        return default

# Ambil satu entry log dari Google Sheets berdasarkan No. Surat (bukan No. Agenda)
def get_log_entry_by_no_surat(no_surat):
    service = get_sheets_service()
    range_name = 'Sheet1!A6:AB'
    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=range_name
    ).execute()
    values = result.get('values', [])
    for row in values:
        row = row + ["" for _ in range(len(ENHANCED_HEADER) - len(row))]
        if row[1] == no_surat:
            data = {ENHANCED_HEADER[i]: row[i] for i in range(len(ENHANCED_HEADER))}
            # Konversi instruksi jabatan ke list of dict
            instruksi_jabatan_map = [
                ("Direktur Utama Instruksi", "Direktur Utama", "Direktur Utama Tanggal"),
                ("Direktur Keuangan Instruksi", "Direktur Keuangan", "Direktur Keuangan Tanggal"),
                ("Direktur Teknik Instruksi", "Direktur Teknik", "Direktur Teknik Tanggal"),
                ("GM Keuangan & Administrasi Instruksi", "GM Keuangan & Administrasi", "GM Keuangan & Administrasi Tanggal"),
                ("GM Operasional & Pemeliharaan Instruksi", "GM Operasional & Pemeliharaan", "GM Operasional & Pemeliharaan Tanggal"),
                ("Manager Pemeliharaan Instruksi", "Manager Pemeliharaan", "Manager Pemeliharaan Tanggal"),
                ("Manager Operasional Instruksi", "Manager Operasional", "Manager Operasional Tanggal"),
                ("Manager Administrasi Instruksi", "Manager Administrasi", "Manager Administrasi Tanggal"),
                ("Manager Keuangan Instruksi", "Manager Keuangan", "Manager Keuangan Tanggal")
            ]
            instruksi_from_log = []
            for instr_col, posisi_label, tgl_col in instruksi_jabatan_map:
                instruksi_val = data.get(instr_col, "").strip()
                tgl_val = data.get(tgl_col, "").strip()
                if instruksi_val:
                    instruksi_from_log.append({
                        "posisi": posisi_label,
                        "instruksi": instruksi_val,
                        "tanggal": tgl_val
                    })
            data["isi_instruksi"] = instruksi_from_log
            return data
    return None

# Overwrite baris di Google Sheets sesuai No. Surat dengan data_baru
def update_log_entry(data_lama, data_baru):
    try:
        service = get_sheets_service()
        range_name = 'Sheet1!A6:AB'
        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=range_name
        ).execute()
        values = result.get('values', []) or []
        idx = -1
        no_surat_lama = safe_get_value(data_lama, "No. Surat")
        
        # Find the row to update
        for i, row in enumerate(values):
            row_full = row + ["" for _ in range(len(ENHANCED_HEADER) - len(row))]
            if str(row_full[1]).strip() == no_surat_lama:
                idx = i
                break
        
        if idx == -1:
            raise Exception(f"Data dengan No. Surat '{no_surat_lama}' tidak ditemukan di sheet, tidak bisa update.")

        # ROBUST MERGE: Gabungkan data_lama dan data_baru agar field yang tidak diubah tetap ada
        merged_data = dict(data_lama) if data_lama else {}
        
        # Update with new data, handling both snake_case and Title Case keys
        for k, v in data_baru.items():
            if v is not None:  # Allow empty strings but not None
                merged_data[k] = v
                
                # Handle field name mappings
                field_mappings = {
                    "no_agenda": "No. Agenda",
                    "no_surat": "No. Surat", 
                    "tgl_surat": "Tgl. Surat",
                    "perihal": "Perihal",
                    "asal_surat": "Asal Surat",
                    "ditujukan": "Ditujukan",
                    "kode_klasifikasi": "Kode Klasifikasi",
                    "tgl_terima": "Tgl. Penerimaan",
                    "indeks": "Indeks",
                    "bicarakan_dengan": "Bicarakan dengan",
                    "teruskan_kepada": "Teruskan kepada",
                    "harap_selesai_tgl": "Harap Selesai Tanggal"
                }
                
                # If snake_case key exists, also set the Title Case version
                if k in field_mappings:
                    merged_data[field_mappings[k]] = v

        # Sinkronisasi kolom tanggal penyelesaian agar konsisten
        harap_selesai = merged_data.get("Harap Selesai Tanggal", merged_data.get("harap_selesai_tgl", ""))
        if harap_selesai:
            merged_data["Selesai Tgl."] = harap_selesai

        # FIXED: Valid position labels only
        pejabat_labels = [
            "Direktur Utama", "Direktur Keuangan", "Direktur Teknik",
            "GM Keuangan & Administrasi", "GM Operasional & Pemeliharaan", 
            "Manager Pemeliharaan", "Manager Operasional", "Manager Administrasi", "Manager Keuangan"
        ]
        
        # Initialize instruction map with valid positions only
        instruksi_map = {label: {"instruksi": "", "tanggal": ""} for label in pejabat_labels}
        
        # Process instruction data
        isi_instruksi = merged_data.get("isi_instruksi", [])
        if isi_instruksi and isinstance(isi_instruksi, list):
            for instruksi_item in isi_instruksi:
                if isinstance(instruksi_item, dict):
                    posisi = safe_get_value(instruksi_item, "posisi")
                    if posisi in instruksi_map:  # Only process valid positions
                        instruksi_map[posisi]["instruksi"] = safe_get_value(instruksi_item, "instruksi")
                        instruksi_map[posisi]["tanggal"] = normalize_date_format(instruksi_item.get("tanggal", ""))

        # Build row data with robust null handling
        row_data = []
        
        for col in ENHANCED_HEADER:
            try:
                if col.endswith(" Instruksi"):
                    # Extract position name by removing " Instruksi" suffix
                    label = col.replace(" Instruksi", "")
                    if label in instruksi_map:
                        row_data.append(safe_get_value(instruksi_map[label], "instruksi"))
                    else:
                        print(f"[WARNING] Unknown position for instruction: {label}")
                        row_data.append("")
                        
                elif col.endswith(" Tanggal"):
                    # Extract position name by removing " Tanggal" suffix  
                    label = col.replace(" Tanggal", "")
                    if label in instruksi_map:
                        row_data.append(safe_get_value(instruksi_map[label], "tanggal"))
                    else:
                        print(f"[WARNING] Unknown position for date: {label}")
                        row_data.append("")
                        
                elif col == "Klasifikasi":
                    # Handle classification with null safety
                    if merged_data.get("rahasia", 0):
                        row_data.append("RAHASIA")
                    elif merged_data.get("penting", 0):
                        row_data.append("PENTING") 
                    elif merged_data.get("segera", 0):
                        row_data.append("SEGERA")
                    else:
                        row_data.append("")
                        
                elif col == "Disposisi kepada":
                    disposisi_labels = []
                    if merged_data.get("dir_utama", 0): disposisi_labels.append("Direktur Utama")
                    if merged_data.get("dir_keu", 0): disposisi_labels.append("Direktur Keuangan")
                    if merged_data.get("dir_teknik", 0): disposisi_labels.append("Direktur Teknik")
                    if merged_data.get("gm_keu", 0): disposisi_labels.append("GM Keuangan & Administrasi")
                    if merged_data.get("gm_ops", 0): disposisi_labels.append("GM Operasional & Pemeliharaan")
                    if merged_data.get("manager_pemeliharaan", 0): disposisi_labels.append("Manager Pemeliharaan")
                    if merged_data.get("manager_operasional", 0): disposisi_labels.append("Manager Operasional")
                    if merged_data.get("manager_administrasi", 0): disposisi_labels.append("Manager Administrasi")
                    if merged_data.get("manager_keuangan", 0): disposisi_labels.append("Manager Keuangan")
                    merged_data[col] = ", ".join(disposisi_labels)
                    
                elif col == "Untuk Di :":
                    # Handle action checkboxes
                    untuk_di_labels = []
                    if merged_data.get("ketahui_file", 0): untuk_di_labels.append("Ketahui & File")
                    if merged_data.get("proses_selesai", 0): untuk_di_labels.append("Proses Selesai")
                    if merged_data.get("teliti_pendapat", 0): untuk_di_labels.append("Teliti & Pendapat")
                    if merged_data.get("buatkan_resume", 0): untuk_di_labels.append("Buatkan Resume")
                    if merged_data.get("edarkan", 0): untuk_di_labels.append("Edarkan")
                    if merged_data.get("sesuai_disposisi", 0): untuk_di_labels.append("Sesuai Disposisi")
                    if merged_data.get("bicarakan_saya", 0): untuk_di_labels.append("Bicarakan dengan Saya")
                    row_data.append(", ".join(untuk_di_labels))
                    
                elif col in ["Tgl. Surat", "Tgl. Penerimaan", "Harap Selesai Tanggal", "Selesai Tgl."]:
                    # Handle date fields with normalization
                    date_value = safe_get_value(merged_data, col)
                    row_data.append(normalize_date_format(date_value))
                    
                else:
                    # Handle regular fields with null safety
                    row_data.append(safe_get_value(merged_data, col))
                    
            except Exception as e:
                print(f"[WARNING] Error processing column {col}: {e}")
                row_data.append("")  # Add empty value to maintain column alignment

        print(f"[update_log_entry] Updating row {idx+6} with {len(row_data)} columns")
        
        # Ensure we have exactly 28 columns
        while len(row_data) < 28:
            row_data.append("")
        row_data = row_data[:28]  # Trim if too many
        
        # Update the row
        update_row_in_sheet(row_data, idx+1)  # idx+1 karena row_number mulai dari 1 = baris ke-6
        
        print(f"[update_log_entry] Successfully updated row for No. Surat: {no_surat_lama}")
        return True
        
    except Exception as e:
        print(f"[update_log_entry] Error: {e}")
        traceback.print_exc()
        raise e  # Re-raise to show error to user

def build_complete_data(data_lama, data_baru, header):
    """Build complete data with robust null handling"""
    result = {}
    for col in header:
        # Prioritize new data, fallback to old data, then empty string
        if col in data_baru and data_baru[col] is not None:
            result[col] = data_baru[col]
        elif col in data_lama and data_lama[col] is not None:
            result[col] = data_lama[col]
        else:
            result[col] = ""
    return result
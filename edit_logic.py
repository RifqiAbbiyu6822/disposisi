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

# Field mapping for consistent data conversion
FIELD_MAPPING = {
    "Tgl. Surat": "tgl_surat",
    "Tgl. Penerimaan": "tgl_terima", 
    "Kode Klasifikasi": "kode_klasifikasi",
    "Indeks": "indeks",
    "Harap Selesai Tanggal": "harap_selesai_tgl",
    "Selesai Tgl.": "selesai_tgl",
    "Bicarakan dengan": "bicarakan_dengan",
    "Teruskan kepada": "teruskan_kepada",
    "No. Agenda": "no_agenda",
    "No. Surat": "no_surat",
    "Perihal": "perihal",
    "Asal Surat": "asal_surat",
    "Ditujukan": "ditujukan"
}

# Position mapping for instruction fields
POSITION_MAPPING = {
    "Direktur Utama": "dir_utama",
    "Direktur Keuangan": "dir_keu", 
    "Direktur Teknik": "dir_teknik",
    "GM Keuangan & Administrasi": "gm_keu",
    "GM Operasional & Pemeliharaan": "gm_ops",
    "Manager Pemeliharaan": "manager_pemeliharaan",
    "Manager Operasional": "manager_operasional", 
    "Manager Administrasi": "manager_administrasi",
    "Manager Keuangan": "manager_keuangan"
}

# Reverse mapping for converting back to sheet format
REVERSE_FIELD_MAPPING = {v: k for k, v in FIELD_MAPPING.items()}

def normalize_date_format(date_str):
    """Normalize date format with robust null handling"""
    # Handle null/empty values
    if date_str is None:
        return ""
    
    date_str = str(date_str).strip()
    if not date_str or date_str.lower() in ['none', 'null', 'nan', '']:
        return ""
    
    try:
        from datetime import datetime
        
        # Common date formats to try
        formats = ["%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d.%m.%Y", "%m/%d/%Y"]
        
        for fmt in formats:
            try:
                date_obj = datetime.strptime(date_str, fmt)
                return date_obj.strftime("%d-%m-%Y")
            except ValueError:
                continue
        
        # If no format matches, return as is (but cleaned)
        print(f"[normalize_date_format] Warning: Could not parse date '{date_str}', returning as-is")
        return date_str
        
    except Exception as e:
        print(f"[normalize_date_format] Warning: Could not parse date '{date_str}': {e}")
        return ""

def safe_get_value(data, key, default=""):
    """Safely get value from data dict with null handling"""
    try:
        if key not in data:
            return default
        
        value = data[key]
        if value is None:
            return default
        
        # Convert to string and handle empty/null-like values
        value_str = str(value).strip()
        if not value_str or value_str.lower() in ['none', 'null', 'nan', '']:
            return default
        
        return value_str
    except Exception as e:
        print(f"[safe_get_value] Error getting {key}: {e}")
        return default

# Ambil satu entry log dari Google Sheets berdasarkan No. Surat (bukan No. Agenda)
def get_log_entry_by_no_surat(no_surat):
    service = get_sheets_service()
    range_name = 'Sheet1!A6:AH'  # Changed from AB to AH for 34 columns
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
        range_name = 'Sheet1!A6:AH'  # Changed from AB to AH for 34 columns
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
                
                # Use consistent field mapping
                if k in REVERSE_FIELD_MAPPING:
                    merged_data[REVERSE_FIELD_MAPPING[k]] = v

        # Sinkronisasi kolom tanggal penyelesaian agar konsisten
        harap_selesai = merged_data.get("Harap Selesai Tanggal", merged_data.get("harap_selesai_tgl", ""))
        if harap_selesai:
            merged_data["Selesai Tgl."] = harap_selesai

        # FIXED: Valid position labels only - match with ENHANCED_HEADER
        pejabat_labels = [
            "Direktur Utama", "Direktur Keuangan", "Direktur Teknik",
            "GM Keuangan & Administrasi", "GM Operasional & Pemeliharaan", 
            "Manager Pemeliharaan", "Manager Operasional", "Manager Administrasi", "Manager Keuangan"
        ]
        
        # Initialize instruction map with valid positions only
        instruksi_map = {label: {"instruksi": "", "tanggal": ""} for label in pejabat_labels}
        
        # Process instruction data from isi_instruksi
        isi_instruksi = merged_data.get("isi_instruksi", [])
        if isi_instruksi and isinstance(isi_instruksi, list):
            for instruksi_item in isi_instruksi:
                if isinstance(instruksi_item, dict):
                    posisi = safe_get_value(instruksi_item, "posisi")
                    if posisi in instruksi_map:  # Only process valid positions
                        instruksi_map[posisi]["instruksi"] = safe_get_value(instruksi_item, "instruksi")
                        instruksi_map[posisi]["tanggal"] = normalize_date_format(instruksi_item.get("tanggal", ""))
        
        # Also process individual instruction fields if they exist in merged_data
        for label in pejabat_labels:
            instruksi_key = f"{label} Instruksi"
            tanggal_key = f"{label} Tanggal"
            
            if instruksi_key in merged_data and merged_data[instruksi_key]:
                instruksi_map[label]["instruksi"] = safe_get_value(merged_data, instruksi_key)
            if tanggal_key in merged_data and merged_data[tanggal_key]:
                instruksi_map[label]["tanggal"] = normalize_date_format(merged_data[tanggal_key])

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
                    # Special handling for "Harap Selesai Tanggal" field
                    if col == "Harap Selesai Tanggal":
                        value = safe_get_value(merged_data, "harap_selesai_tgl")
                        value = normalize_date_format(value)
                        row_data.append(value)
                    elif label in instruksi_map:
                        row_data.append(safe_get_value(instruksi_map[label], "tanggal"))
                    else:
                        print(f"[WARNING] Unknown position for date: {label}")
                        row_data.append("")
                        
                elif col in ["No. Agenda", "No. Surat", "Tgl. Surat", "Perihal", "Asal Surat", "Ditujukan", 
                           "Kode Klasifikasi", "Tgl. Penerimaan", "Indeks", "Bicarakan dengan", 
                           "Teruskan kepada", "Selesai Tgl."]:
                    # Handle basic fields with null safety
                    value = safe_get_value(merged_data, col)
                    if col in ["Tgl. Surat", "Tgl. Penerimaan", "Selesai Tgl."]:
                        value = normalize_date_format(value)
                        print(f"[update_log_entry] Date field {col}: '{value}'")
                    row_data.append(value)
                        
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
                    # Handle disposisi checkboxes with abbreviation for managers
                    disposisi_labels = []
                    manager_labels = []
                    
                    if merged_data.get("dir_utama", 0): disposisi_labels.append("Direktur Utama")
                    if merged_data.get("dir_keu", 0): disposisi_labels.append("Direktur Keuangan")
                    if merged_data.get("dir_teknik", 0): disposisi_labels.append("Direktur Teknik")
                    if merged_data.get("gm_keu", 0): disposisi_labels.append("GM Keuangan & Administrasi")
                    if merged_data.get("gm_ops", 0): disposisi_labels.append("GM Operasional & Pemeliharaan")
                    
                    # Collect managers separately for abbreviation
                    if merged_data.get("manager_pemeliharaan", 0): manager_labels.append("pml")
                    if merged_data.get("manager_operasional", 0): manager_labels.append("ops")
                    if merged_data.get("manager_administrasi", 0): manager_labels.append("adm")
                    if merged_data.get("manager_keuangan", 0): manager_labels.append("keu")
                    
                    # Add abbreviated managers
                    if manager_labels:
                        disposisi_labels.append(f"Manager {', '.join(manager_labels)}")
                    
                    row_data.append(", ".join(disposisi_labels))
                    
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
                    
                else:
                    # Handle any remaining fields with null safety
                    # Check if we have the field in merged_data, otherwise use empty string
                    if col in merged_data:
                        row_data.append(safe_get_value(merged_data, col))
                    else:
                        row_data.append("")
                    
            except Exception as e:
                print(f"[WARNING] Error processing column {col}: {e}")
                row_data.append("")  # Add empty value to maintain column alignment

        print(f"[update_log_entry] Updating row {idx+6} with {len(row_data)} columns")
        print(f"[update_log_entry] Expected columns: {len(ENHANCED_HEADER)}")
        
        # Ensure we have exactly 34 columns for Google Sheets (A-AH)
        while len(row_data) < 34:
            row_data.append("")
        row_data = row_data[:34]  # Trim if too many
        
        print(f"[update_log_entry] Final row data length: {len(row_data)} (should be 34)")
        print(f"[update_log_entry] Row data preview: {row_data[:5]}...")  # Show first 5 columns
        
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
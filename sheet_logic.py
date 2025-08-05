import traceback
import logging
from tkinter import messagebox
from google_sheets_connect import append_row_to_sheet, write_multilayer_header, get_sheets_service, SHEET_ID, update_row_in_sheet
from datetime import datetime

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

# Reverse mapping for converting back to sheet format
REVERSE_FIELD_MAPPING = {v: k for k, v in FIELD_MAPPING.items()}

def safe_get_value(data, key, default=""):
    """Safely get value from data dict with comprehensive null handling"""
    try:
        value = data.get(key, default)
        
        # Handle None, empty string, and null-like values
        if value is None:
            return default
        
        # Convert to string and handle various null representations
        str_value = str(value).strip()
        if str_value.lower() in ['none', 'null', 'nan', '']:
            return default
            
        return str_value
    except Exception as e:
        print(f"[safe_get_value] Error getting {key}: {e}")
        return default

def normalize_date_format(date_str, target_format="%d-%m-%Y"):
    """Convert date string to consistent format with robust null handling"""
    if not date_str:
        return ""
    
    # Handle None and null-like values
    date_str = safe_get_value({"date": date_str}, "date", "")
    if not date_str:
        return ""
    
    # Common date formats to try
    formats = ["%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d", "%d.%m.%Y"]
    
    for fmt in formats:
        try:
            date_obj = datetime.strptime(date_str, fmt)
            return date_obj.strftime(target_format)
        except ValueError:
            continue
    
    # If no format matches, return the original string (cleaned)
    print(f"[normalize_date_format] Warning: Could not parse date '{date_str}', returning as-is")
    return date_str

def safe_get_widget_value(widget, widget_name="unknown"):
    """Safely get value from any widget type with enhanced null handling"""
    try:
        if widget is None:
            print(f"[WARNING] Widget {widget_name} is None")
            return ""
        
        # For tkcalendar DateEntry
        if hasattr(widget, 'get_date'):
            try:
                date_obj = widget.get_date()
                if date_obj:
                    return date_obj.strftime("%d-%m-%Y")
                else:
                    return ""
            except Exception as e:
                print(f"[WARNING] Error getting date from {widget_name}: {e}")
                # Fallback to regular get() if available
                return safe_get_value({"val": widget.get() if hasattr(widget, 'get') else ""}, "val", "")
        
        # For regular Entry widgets
        elif hasattr(widget, 'get'):
            # Check if get method requires arguments
            import inspect
            sig = inspect.signature(widget.get)
            if len(sig.parameters) == 0:
                return safe_get_value({"val": widget.get()}, "val", "")
            else:
                # For widgets like Listbox that need selection
                try:
                    selection = widget.curselection()
                    if selection:
                        return safe_get_value({"val": widget.get(selection[0])}, "val", "")
                    else:
                        return ""
                except:
                    return ""
        
        # For Text widgets
        elif hasattr(widget, 'get') and hasattr(widget, 'insert'):
            text_content = widget.get("1.0", "end").strip()
            return safe_get_value({"val": text_content}, "val", "")
        
        else:
            print(f"[WARNING] Unknown widget type for {widget_name}: {type(widget)}")
            return ""
            
    except Exception as e:
        print(f"[ERROR] Error getting value from {widget_name}: {e}")
        traceback.print_exc()
        return ""

def get_disposisi_labels(self):
    """Get selected disposisi labels with null handling"""
    mapping = [
        ("dir_utama", "Direktur Utama"),
        ("dir_keu", "Direktur Keuangan"),
        ("dir_teknik", "Direktur Teknik"),
        ("gm_keu", "GM Keuangan & Administrasi"),
        ("gm_ops", "GM Operasional & Pemeliharaan"),
        ("manager_pemeliharaan", "Manager Pemeliharaan"),
        ("manager_operasional", "Manager Operasional"),
        ("manager_administrasi", "Manager Administrasi"),
        ("manager_keuangan", "Manager Keuangan"),
    ]
    labels = []
    for var, label in mapping:
        try:
            if hasattr(self, 'vars') and var in self.vars and self.vars[var].get():
                labels.append(label)
        except Exception as e:
            print(f"[WARNING] Error getting disposisi value for {var}: {e}")
    return labels

def get_disposisi_labels_with_abbreviation(self):
    """Get selected disposisi labels with abbreviation for managers"""
    mapping = [
        ("dir_utama", "Direktur Utama"),
        ("dir_keu", "Direktur Keuangan"),
        ("dir_teknik", "Direktur Teknik"),
        ("gm_keu", "GM Keuangan & Administrasi"),
        ("gm_ops", "GM Operasional & Pemeliharaan"),
        ("manager_pemeliharaan", "Manager Pemeliharaan"),
        ("manager_operasional", "Manager Operasional"),
        ("manager_administrasi", "Manager Administrasi"),
        ("manager_keuangan", "Manager Keuangan"),
    ]
    
    labels = []
    manager_labels = []
    
    for var, label in mapping:
        try:
            if hasattr(self, 'vars') and var in self.vars and self.vars[var].get():
                # Pisahkan manager dari label lainnya
                if label.startswith("Manager"):
                    manager_labels.append(label)
                else:
                    labels.append(label)
        except Exception as e:
            print(f"[WARNING] Error getting disposisi value for {var}: {e}")
    
    # Gabungkan semua manager menjadi satu dengan singkatan pendek
    if manager_labels:
        manager_abbreviations = []
        for manager in manager_labels:
            if "Pemeliharaan" in manager:
                manager_abbreviations.append("pml")
            elif "Operasional" in manager:
                manager_abbreviations.append("ops")
            elif "Administrasi" in manager:
                manager_abbreviations.append("adm")
            elif "Keuangan" in manager:
                manager_abbreviations.append("keu")
        
        if manager_abbreviations:
            labels.append(f"Manager {', '.join(manager_abbreviations)}")
    
    return labels

def get_untuk_di_labels(self, data):
    """Get selected untuk di labels with null handling"""
    mapping = [
        ("ketahui_file", "Ketahui & File"),
        ("proses_selesai", "Proses Selesai"),
        ("teliti_pendapat", "Teliti & Pendapat"),
        ("buatkan_resume", "Buatkan Resume"),
        ("edarkan", "Edarkan"),
        ("sesuai_disposisi", "Sesuai Disposisi"),
        ("bicarakan_saya", "Bicarakan dengan Saya")
    ]
    labels = []
    for var, label in mapping:
        try:
            if data.get(var, 0):
                labels.append(label)
        except Exception as e:
            print(f"[WARNING] Error getting untuk_di value for {var}: {e}")
    return ", ".join(labels)

def prepare_row_data(self, data):
    """Prepare row data for Google Sheets with robust null handling"""
    try:
        # Normalize dates with null handling
        date_fields = ["tgl_surat", "tgl_terima", "harap_selesai_tgl", "selesai_tgl"]
        for field in date_fields:
            if field in data:
                data[field] = normalize_date_format(data[field])
        
        # Get classification - only one can be active
        klasifikasi = ""
        if data.get("rahasia", 0):
            klasifikasi = "RAHASIA"
        elif data.get("penting", 0):
            klasifikasi = "PENTING"
        elif data.get("segera", 0):
            klasifikasi = "SEGERA"
        
        # Build basic row data
        row_data = [
            safe_get_value(data, "no_agenda"),
            safe_get_value(data, "no_surat"),
            safe_get_value(data, "tgl_surat"),
            safe_get_value(data, "perihal"),
            safe_get_value(data, "asal_surat"),
            safe_get_value(data, "ditujukan"),
            klasifikasi,
            ", ".join(get_disposisi_labels_with_abbreviation(self)),  # Use abbreviation for output
            get_untuk_di_labels(self, data),
            safe_get_value(data, "selesai_tgl", safe_get_value(data, "harap_selesai_tgl")),  # Use selesai_tgl if available, else harap_selesai_tgl
            safe_get_value(data, "kode_klasifikasi"),
            safe_get_value(data, "tgl_terima"),
            safe_get_value(data, "indeks"),
            safe_get_value(data, "bicarakan_dengan"),
            safe_get_value(data, "teruskan_kepada"),
            safe_get_value(data, "harap_selesai_tgl"),
        ]
        
        # Add instruction columns with null handling
        pejabat_keys = [
            ("dir_utama", "Direktur Utama"),
            ("dir_keu", "Direktur Keuangan"),
            ("dir_teknik", "Direktur Teknik"),
            ("gm_keu", "GM Keuangan & Administrasi"),
            ("gm_ops", "GM Operasional & Pemeliharaan"),
            ("manager_pemeliharaan", "Manager Pemeliharaan"),
            ("manager_operasional", "Manager Operasional"),
            ("manager_administrasi", "Manager Administrasi"),
            ("manager_keuangan", "Manager Keuangan")
        ]
        
        # Initialize instruction map
        instruksi_map = {p[1]: {"instruksi": "", "tanggal": ""} for p in pejabat_keys}
        
        # Process instruction data
        isi_instruksi = data.get("isi_instruksi", [])
        if isi_instruksi and isinstance(isi_instruksi, list):
            for instruksi_item in isi_instruksi:
                if isinstance(instruksi_item, dict):
                    posisi = safe_get_value(instruksi_item, "posisi")
                    if posisi in instruksi_map:
                        instruksi_map[posisi]["instruksi"] = safe_get_value(instruksi_item, "instruksi")
                        instruksi_map[posisi]["tanggal"] = normalize_date_format(instruksi_item.get("tanggal", ""))
        
        # Add instruction and date columns for all positions (including managers separately)
        for _, label in pejabat_keys:
            row_data.append(instruksi_map[label]["instruksi"])
            row_data.append(instruksi_map[label]["tanggal"])
        
        # Ensure we have exactly 34 columns
        while len(row_data) < 34:
            row_data.append("")
        row_data = row_data[:34]  # Trim if too many
        
        return row_data
        
    except Exception as e:
        print(f"[prepare_row_data] Error: {e}")
        traceback.print_exc()
        # Return minimal valid row data to prevent complete failure
        return [""] * 34

def upload_to_sheet(self, call_from_pdf=False, data_override=None):
    """Upload data to Google Sheets with improved error handling and null safety"""
    try:
        if data_override:
            data = data_override
        else:
            # Collect form data with safe widget value extraction
            data = {}
            
            # Get data from vars (usually IntVar, StringVar, etc)
            if hasattr(self, 'vars'):
                for key, var in self.vars.items():
                    try:
                        value = var.get()
                        data[key] = value if value is not None else (0 if 'IntVar' in str(type(var)) else "")
                    except Exception as e:
                        print(f"[WARNING] Error getting value from var {key}: {e}")
                        data[key] = 0 if 'IntVar' in str(type(var)) else ""
            
            # Get data from form input widgets
            if hasattr(self, 'form_input_widgets'):
                # Text widgets (multiline)
                text_fields = ["perihal", "asal_surat", "ditujukan", "bicarakan_dengan", "teruskan_kepada"]
                for field in text_fields:
                    if field in self.form_input_widgets:
                        try:
                            widget = self.form_input_widgets[field]
                            if hasattr(widget, 'get'):
                                if hasattr(widget, 'insert'):  # Text widget
                                    data[field] = widget.get("1.0", "end").strip()
                                else:  # Entry widget
                                    data[field] = safe_get_widget_value(widget, field)
                            else:
                                data[field] = ""
                        except Exception as e:
                            print(f"[WARNING] Error getting value from {field}: {e}")
                            data[field] = ""
                
                # Date widgets
                if "tgl_surat" in self.form_input_widgets:
                    data["tgl_surat"] = safe_get_widget_value(self.form_input_widgets["tgl_surat"], "tgl_surat")
            
            # Handle special date entries with null safety
            if hasattr(self, 'tgl_terima_entry'):
                data["tgl_terima"] = safe_get_widget_value(self.tgl_terima_entry, "tgl_terima_entry")
            else:
                data["tgl_terima"] = ""
            
            if hasattr(self, 'harap_selesai_tgl_entry'):
                data["harap_selesai_tgl"] = safe_get_widget_value(self.harap_selesai_tgl_entry, "harap_selesai_tgl_entry")
            else:
                data["harap_selesai_tgl"] = ""
            
            # Ensure consistent field names with defaults
            required_fields = {
                "kode_klasifikasi": "",
                "indeks": "",
                "no_agenda": "",
                "no_surat": "",
                "perihal": "",
                "asal_surat": "",
                "ditujukan": "",
                "bicarakan_dengan": "",
                "teruskan_kepada": "",
                "tgl_surat": "",
                "rahasia": 0,
                "penting": 0,
                "segera": 0
            }
            
            for field, default_value in required_fields.items():
                if field not in data or data[field] is None:
                    data[field] = default_value
            
            # Collect instructions with null safety
            isi_instruksi = []
            if hasattr(self, "instruksi_table") and hasattr(self.instruksi_table, "get_data"):
                try:
                    isi_instruksi = self.instruksi_table.get_data()
                    # Filter out completely empty instruction rows
                    isi_instruksi = [
                        instr for instr in isi_instruksi 
                        if isinstance(instr, dict) and (
                            safe_get_value(instr, "posisi") or 
                            safe_get_value(instr, "instruksi") or 
                            safe_get_value(instr, "tanggal")
                        )
                    ]
                except Exception as e:
                    print(f"[WARNING] Error getting instruction data: {e}")
                    isi_instruksi = []
            data["isi_instruksi"] = isi_instruksi
        
        # VALIDASI: Hanya No. Surat yang wajib diisi
        no_surat = safe_get_value(data, "no_surat")
        if not no_surat:
            messagebox.showerror("Validasi", "No. Surat wajib diisi!")
            return False
        
        # Cek duplikasi No. Surat dengan error handling
        try:
            from disposisi_app.views.components.validation import is_no_surat_unique
            if not is_no_surat_unique(no_surat, get_sheets_service, SHEET_ID):
                messagebox.showerror("Validasi", f"No. Surat '{no_surat}' sudah ada di database!")
                return False
        except Exception as e:
            print(f"[WARNING] Error checking uniqueness: {e}")
            # Continue anyway if validation fails - better to save than lose data
        
        # Ensure sheet has proper headers
        from googleapiclient.errors import HttpError
        import time
        sheet_name = 'Sheet1'
        
        try:
            service = get_sheets_service()
            result = service.spreadsheets().values().get(
                spreadsheetId=SHEET_ID,
                range=f'{sheet_name}!A1:A4'
            ).execute()
            values = result.get('values', [])
            if len(values) < 4:
                write_multilayer_header(sheet_id=SHEET_ID, sheet_name=sheet_name)
                time.sleep(1)
        except HttpError as e:
            print(f"[WARNING] Error checking headers: {e}")
            try:
                write_multilayer_header(sheet_id=SHEET_ID, sheet_name=sheet_name)
                time.sleep(1)
            except Exception as header_error:
                print(f"[WARNING] Could not write headers: {header_error}")
        
        # Prepare and upload data
        row_data = prepare_row_data(self, data)
        append_row_to_sheet(row_data, range_name=f'{sheet_name}!A6')
        
        if not call_from_pdf:
            messagebox.showinfo("Sukses", "Data berhasil diupload ke Google Sheets.")
        
        return True
            
    except Exception as e:
        traceback.print_exc()
        error_msg = f"Gagal upload ke Google Sheets: {e}"
        print(f"[ERROR] {error_msg}")
        messagebox.showerror("Google Sheets", error_msg)
        return False

def update_log_entry(data_lama, data_baru):
    """Update log entry in Google Sheets with robust error handling"""
    try:
        service = get_sheets_service()
        range_name = 'Sheet1!A6:AB'
        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=range_name
        ).execute()
        values = result.get('values', []) or []
        
        # Find the row to update
        idx = -1
        no_surat_lama = safe_get_value(data_lama, "No. Surat")
        
        for i, row in enumerate(values):
            row_full = row + ["" for _ in range(len(ENHANCED_HEADER) - len(row))]
            if str(row_full[1]).strip() == no_surat_lama:
                idx = i
                break
        
        if idx == -1:
            raise Exception(f"Data dengan No. Surat '{no_surat_lama}' tidak ditemukan di sheet, tidak bisa update.")
        
        # Convert snake_case fields back to sheet headers with null handling
        converted_data_baru = {}
        for k, v in data_baru.items():
            if k in REVERSE_FIELD_MAPPING:
                converted_data_baru[REVERSE_FIELD_MAPPING[k]] = v
            else:
                converted_data_baru[k] = v
        
        # Merge old and new data with null safety
        merged_data = dict(data_lama) if data_lama else {}
        for k, v in converted_data_baru.items():
            if v is not None:  # Allow empty strings but not None
                merged_data[k] = v
        
        # Synchronize completion dates
        if "Harap Selesai Tanggal" in merged_data:
            merged_data["Selesai Tgl."] = merged_data["Harap Selesai Tanggal"]
        
        # Process instructions with null handling
        pejabat_labels = [
            "Direktur Utama", "Direktur Keuangan", "Direktur Teknik",
            "GM Keuangan & Administrasi", "GM Operasional & Pemeliharaan", "Manager"
        ]
        
        instruksi_map = {label: {"instruksi": "", "tanggal": ""} for label in pejabat_labels}
        isi_instruksi = merged_data.get("isi_instruksi", [])
        
        if isi_instruksi and isinstance(isi_instruksi, list):
            for instruksi_item in isi_instruksi:
                if isinstance(instruksi_item, dict):
                    posisi = safe_get_value(instruksi_item, "posisi")
                    if posisi in instruksi_map:
                        instruksi_map[posisi]["instruksi"] = safe_get_value(instruksi_item, "instruksi")
                        instruksi_map[posisi]["tanggal"] = normalize_date_format(instruksi_item.get("tanggal", ""))
        
        # Build row data with comprehensive error handling
        row_data = []
        missing_keys = []
        
        for col in ENHANCED_HEADER:
            try:
                if col in ["No. Agenda", "No. Surat", "Tgl. Surat", "Perihal", "Asal Surat", "Ditujukan", 
                           "Klasifikasi", "Tgl. Penerimaan", "Indeks", "Bicarakan dengan", 
                           "Teruskan kepada", "Harap Selesai Tanggal", "Selesai Tgl."]:
                    value = safe_get_value(merged_data, col)
                    if col in ["Tgl. Surat", "Tgl. Penerimaan", "Harap Selesai Tanggal", "Selesai Tgl."]:
                        value = normalize_date_format(value)
                    row_data.append(value)
                    
                elif col.endswith("Instruksi"):
                    label = col.replace(" Instruksi", "")
                    if label in instruksi_map:
                        row_data.append(instruksi_map[label]["instruksi"])
                    else:
                        print(f"[WARNING] Unknown position for instruction: {label}")
                        row_data.append("")
                        
                elif col.endswith("Tanggal"):
                    label = col.replace(" Tanggal", "")
                    if label in instruksi_map:
                        row_data.append(instruksi_map[label]["tanggal"])
                    else:
                        print(f"[WARNING] Unknown position for date: {label}")
                        row_data.append("")
                        
                elif col == "Klasifikasi":
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
                    row_data.append(", ".join(disposisi_labels))
                    
                elif col == "Untuk Di :":
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
                    val = safe_get_value(merged_data, col)
                    if val == "" and col not in merged_data:
                        missing_keys.append(col)
                    row_data.append(val)
                    
            except Exception as e:
                print(f"[WARNING] Error processing column {col}: {e}")
                row_data.append("")  # Add empty value to maintain column structure
        
        # Ensure we have exactly 28 columns
        while len(row_data) < 28:
            row_data.append("")
        row_data = row_data[:28]  # Trim if too many
        
        print(f"[update_log_entry] Akan update baris ke-{idx+6} dengan {len(row_data)} kolom")
        if missing_keys:
            print(f"[update_log_entry][INFO] Kolom kosong: {missing_keys}")
        
        # Update the row (idx+1 because row_number starts from 1, which corresponds to row 6 in the sheet)
        update_row_in_sheet(row_data, idx+1)
        
        return True
        
    except Exception as e:
        print(f"[update_log_entry] Error: {e}")
        traceback.print_exc()
        return False

def build_complete_data(data_lama, data_baru, header):
    """Build complete data by merging old and new data with null safety"""
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
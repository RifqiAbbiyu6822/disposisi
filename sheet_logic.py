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
    "GM Operasional & Pemeliharaan Instruksi", "GM Operasional & Pemeliharaan Tanggal", "Manager Instruksi", "Manager Tanggal"
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

def normalize_date_format(date_str, target_format="%d-%m-%Y"):
    """Convert date string to consistent format"""
    if not date_str or date_str.strip() == "":
        return ""
    
    # Common date formats to try
    formats = ["%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%Y"]
    
    for fmt in formats:
        try:
            date_obj = datetime.strptime(date_str.strip(), fmt)
            return date_obj.strftime(target_format)
        except ValueError:
            continue
    
    # If no format matches, return as is
    return date_str.strip()

def safe_get_widget_value(widget, widget_name="unknown"):
    """Safely get value from any widget type"""
    try:
        if widget is None:
            print(f"[WARNING] Widget {widget_name} is None")
            return ""
        
        # For tkcalendar DateEntry
        if hasattr(widget, 'get_date'):
            try:
                date_obj = widget.get_date()
                return date_obj.strftime("%d-%m-%Y")
            except Exception as e:
                print(f"[WARNING] Error getting date from {widget_name}: {e}")
                # Fallback to regular get()
                return str(widget.get()) if hasattr(widget, 'get') else ""
        
        # For regular Entry widgets
        elif hasattr(widget, 'get'):
            # Check if get method requires arguments
            import inspect
            sig = inspect.signature(widget.get)
            if len(sig.parameters) == 0:
                return str(widget.get())
            else:
                # For widgets like Listbox that need selection
                try:
                    selection = widget.curselection()
                    if selection:
                        return str(widget.get(selection[0]))
                    else:
                        return ""
                except:
                    return ""
        
        # For Text widgets
        elif hasattr(widget, 'get') and hasattr(widget, 'insert'):
            return str(widget.get("1.0", "end")).strip()
        
        else:
            print(f"[WARNING] Unknown widget type for {widget_name}: {type(widget)}")
            return ""
            
    except Exception as e:
        print(f"[ERROR] Error getting value from {widget_name}: {e}")
        traceback.print_exc()
        return ""

def get_disposisi_labels(self):
    """Get selected disposisi labels"""
    mapping = [
        ("dir_utama", "Direktur Utama"),
        ("dir_keu", "Direktur Keuangan"),
        ("dir_teknik", "Direktur Teknik"),
        ("gm_keu", "GM Keuangan & Administrasi"),
        ("gm_ops", "GM Operasional & Pemeliharaan"),
        ("manager", "Manager"),
    ]
    labels = []
    for var, label in mapping:
        if hasattr(self, 'vars') and var in self.vars and self.vars[var].get():
            labels.append(label)
    return labels

def get_untuk_di_labels(self, data):
    """Get selected untuk di labels"""
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
        if data.get(var, 0):
            labels.append(label)
    return ", ".join(labels)

def prepare_row_data(self, data):
    """Prepare row data for Google Sheets"""
    # Normalize dates
    date_fields = ["tgl_surat", "tgl_terima", "harap_selesai_tgl", "selesai_tgl"]
    for field in date_fields:
        if field in data:
            data[field] = normalize_date_format(data[field])
    
    row_data = [
        data.get("no_agenda", ""),
        data.get("no_surat", ""),
        data.get("tgl_surat", ""),
        data.get("perihal", ""),
        data.get("asal_surat", ""),
        data.get("ditujukan", ""),
        # Klasifikasi (RAHASIA/PENTING/SEGERA)
        "RAHASIA" if data.get("rahasia", 0) else ("PENTING" if data.get("penting", 0) else ("SEGERA" if data.get("segera", 0) else "")),
        ", ".join(get_disposisi_labels(self)),
        get_untuk_di_labels(self, data),
        data.get("selesai_tgl", data.get("harap_selesai_tgl", "")),  # Use selesai_tgl if available, else harap_selesai_tgl
        data.get("kode_klasifikasi", ""),
        data.get("tgl_terima", ""),
        data.get("indeks", ""),
        data.get("bicarakan_dengan", ""),
        data.get("teruskan_kepada", ""),
        data.get("harap_selesai_tgl", ""),
    ]
    
    # Add instruction columns
    pejabat_keys = [
        ("dir_utama", "Direktur Utama"),
        ("dir_keu", "Direktur Keuangan"),
        ("dir_teknik", "Direktur Teknik"),
        ("gm_keu", "GM Keuangan & Administrasi"),
        ("gm_ops", "GM Operasional & Pemeliharaan"),
        ("manager", "Manager")
    ]
    
    instruksi_map = {p[1]: {"instruksi": "", "tanggal": ""} for p in pejabat_keys}
    for instruksi_item in data.get("isi_instruksi", []):
        posisi = instruksi_item.get("posisi", "").strip()
        for _, label in pejabat_keys:
            if posisi == label:
                instruksi_map[label]["instruksi"] = instruksi_item.get("instruksi", "")
                instruksi_map[label]["tanggal"] = normalize_date_format(instruksi_item.get("tanggal", ""))
    
    for _, label in pejabat_keys:
        row_data.append(instruksi_map[label]["instruksi"])
        row_data.append(instruksi_map[label]["tanggal"])
    
    return row_data

def upload_to_sheet(self, call_from_pdf=False, data_override=None):
    """Upload data to Google Sheets with improved error handling"""
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
                        data[key] = var.get()
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
            
            # Handle special date entries
            if hasattr(self, 'tgl_terima_entry'):
                data["tgl_terima"] = safe_get_widget_value(self.tgl_terima_entry, "tgl_terima_entry")
            
            if hasattr(self, 'harap_selesai_tgl_entry'):
                data["harap_selesai_tgl"] = safe_get_widget_value(self.harap_selesai_tgl_entry, "harap_selesai_tgl_entry")
            
            # Ensure consistent field names
            data["kode_klasifikasi"] = data.get("kode_klasifikasi", "")
            data["indeks"] = data.get("indeks", "")
            
            # Collect instructions
            isi_instruksi = []
            if hasattr(self, "instruksi_table") and hasattr(self.instruksi_table, "get_data"):
                try:
                    isi_instruksi = self.instruksi_table.get_data()
                except Exception as e:
                    print(f"[WARNING] Error getting instruction data: {e}")
                    isi_instruksi = []
            data["isi_instruksi"] = isi_instruksi
        
        # VALIDASI: Hanya No. Surat yang wajib diisi
        no_surat = data.get("no_surat", "").strip()
        if not no_surat:
            messagebox.showerror("Validasi", "No. Surat wajib diisi!")
            return False
        
        # Cek duplikasi No. Surat
        try:
            from disposisi_app.views.components.validation import is_no_surat_unique
            if not is_no_surat_unique(no_surat, get_sheets_service, SHEET_ID):
                messagebox.showerror("Validasi", f"No. Surat '{no_surat}' sudah ada di database!")
                return False
        except Exception as e:
            print(f"[WARNING] Error checking uniqueness: {e}")
            # Continue anyway if validation fails
        
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
        except HttpError:
            write_multilayer_header(sheet_id=SHEET_ID, sheet_name=sheet_name)
            time.sleep(1)
        
        # Prepare and upload data
        row_data = prepare_row_data(self, data)
        append_row_to_sheet(row_data, range_name=f'{sheet_name}!A6')
        
        if not call_from_pdf:
            messagebox.showinfo("Sukses", "Data berhasil diupload ke Google Sheets.")
        
        return True
            
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Google Sheets", f"Gagal upload ke Google Sheets: {e}")
        return False

def update_log_entry(data_lama, data_baru):
    """Update log entry in Google Sheets"""
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
        no_surat_lama = str(data_lama.get("No. Surat", "")).strip()
        
        for i, row in enumerate(values):
            row_full = row + ["" for _ in range(len(ENHANCED_HEADER) - len(row))]
            if str(row_full[1]).strip() == no_surat_lama:
                idx = i
                break
        
        if idx == -1:
            raise Exception(f"Data dengan No. Surat '{no_surat_lama}' tidak ditemukan di sheet, tidak bisa update.")
        
        # Convert snake_case fields back to sheet headers
        converted_data_baru = {}
        for k, v in data_baru.items():
            if k in REVERSE_FIELD_MAPPING:
                converted_data_baru[REVERSE_FIELD_MAPPING[k]] = v
            else:
                converted_data_baru[k] = v
        
        # Merge old and new data
        merged_data = dict(data_lama)
        merged_data.update(converted_data_baru)
        
        # Synchronize completion dates
        if "Harap Selesai Tanggal" in merged_data:
            merged_data["Selesai Tgl."] = merged_data["Harap Selesai Tanggal"]
        
        # Process instructions
        pejabat_labels = [
            "Direktur Utama", "Direktur Keuangan", "Direktur Teknik",
            "GM Keuangan & Administrasi", "GM Operasional & Pemeliharaan", "Manager"
        ]
        
        instruksi_map = {label: {"instruksi": "", "tanggal": ""} for label in pejabat_labels}
        for instruksi_item in merged_data.get("isi_instruksi", []):
            posisi = instruksi_item.get("posisi", "").strip()
            if posisi in instruksi_map:
                instruksi_map[posisi]["instruksi"] = instruksi_item.get("instruksi", "")
                instruksi_map[posisi]["tanggal"] = normalize_date_format(instruksi_item.get("tanggal", ""))
        
        # Build row data
        row_data = []
        missing_keys = []
        
        for col in ENHANCED_HEADER:
            if col in ["No. Agenda", "No. Surat", "Tgl. Surat", "Perihal", "Asal Surat", "Ditujukan", 
                       "Kode Klasifikasi", "Tgl. Penerimaan", "Indeks", "Bicarakan dengan", 
                       "Teruskan kepada", "Harap Selesai Tanggal", "Selesai Tgl."]:
                value = merged_data.get(col, "")
                if col in ["Tgl. Surat", "Tgl. Penerimaan", "Harap Selesai Tanggal", "Selesai Tgl."]:
                    value = normalize_date_format(value)
                row_data.append(value)
                
            elif col.endswith("Instruksi"):
                label = col.replace(" Instruksi", "")
                if label in instruksi_map:
                    row_data.append(instruksi_map[label]["instruksi"])
                else:
                    row_data.append("")
                    
            elif col.endswith("Tanggal"):
                label = col.replace(" Tanggal", "")
                if label in instruksi_map:
                    row_data.append(instruksi_map[label]["tanggal"])
                else:
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
                if merged_data.get("manager", 0): disposisi_labels.append("Manager")
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
                val = merged_data.get(col, "")
                if val == "" and col not in merged_data:
                    missing_keys.append(col)
                row_data.append(val)
        
        print(f"[update_log_entry] Akan update baris ke-{idx+6} dengan data: {row_data}")
        if missing_keys:
            print(f"[update_log_entry][WARNING] Kolom berikut tidak ditemukan di data_baru: {missing_keys}")
        
        # Update the row (idx+1 because row_number starts from 1, which corresponds to row 6 in the sheet)
        update_row_in_sheet(row_data, idx+1)
        
        return True
        
    except Exception as e:
        print(f"[update_log_entry] Error: {e}")
        traceback.print_exc()
        return False

def build_complete_data(data_lama, data_baru, header):
    """Build complete data by merging old and new data"""
    result = {}
    for col in header:
        if col in data_baru and data_baru[col] not in [None, ""]:
            result[col] = data_baru[col]
        elif col in data_lama:
            result[col] = data_lama[col]
        else:
            result[col] = ""
    return result
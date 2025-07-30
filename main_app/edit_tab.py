import tkinter as tk
from tkinter import ttk, Text, messagebox
from tkcalendar import DateEntry
from logic.instruksi_table import InstruksiTable
from edit_logic import update_log_entry
from disposisi_app.views.components.form_sections import (
    populate_frame_kiri, populate_frame_kanan, populate_frame_disposisi, populate_frame_instruksi
)
from disposisi_app.views.components.export_utils import collect_form_data_safely
from datetime import datetime

def bind_mousewheel_recursive(widget, func):
    widget.bind_all("<MouseWheel>", func)
    widget.bind_all("<Shift-MouseWheel>", func)
    widget.bind_all("<Button-4>", func)
    widget.bind_all("<Button-5>", func)
    for child in widget.winfo_children():
        bind_mousewheel_recursive(child, func)

class EditTab(ttk.Frame):
    def __init__(self, parent, data_log, on_save_callback=None):
        super().__init__(parent)
        self.parent = parent
        self.data_log = data_log
        self.on_save_callback = on_save_callback
        self.vars = {}
        self.form_input_widgets = {}
        self.pdf_attachments = []
        
        # Initialize variables and create widgets
        self._init_variables()
        self._create_widgets()
        self._fill_form_from_log(data_log)
        
        # Configure grid weights for proper resizing
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def _init_variables(self):
        # Initialize all variables (same as original)
        keys = [
            "no_agenda", "no_surat", "perihal", "asal_surat", "ditujukan",
            "rahasia", "penting", "segera", "kode_klasifikasi", "indeks",
            "dir_utama", "dir_keu", "dir_teknik", "gm_keu", "gm_ops", "manager",
            "ketahui_file", "proses_selesai", "teliti_pendapat", "buatkan_resume",
            "edarkan", "sesuai_disposisi", "bicarakan_saya", "bicarakan_dengan",
            "teruskan_kepada", "harap_selesai_tgl"
        ]
        
        for key in keys:
            if key in ["rahasia", "penting", "segera", "dir_utama", "dir_keu", "dir_teknik", 
                      "gm_keu", "gm_ops", "manager", "ketahui_file", "proses_selesai", 
                      "teliti_pendapat", "buatkan_resume", "edarkan", "sesuai_disposisi", 
                      "bicarakan_saya"]:
                self.vars[key] = tk.IntVar()
            else:
                self.vars[key] = tk.StringVar()

    def _create_widgets(self):
        # Main container with scrollbars
        root = self.winfo_toplevel()
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, bg=getattr(root, 'canvas_bg', '#f8fafc'))
        self.v_scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.h_scroll = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        
        # Grid layout for scrollable area
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll.grid(row=1, column=0, sticky="ew")
        
        # Main frame inside canvas
        self.main_frame = ttk.Frame(self.canvas, padding="15", style="TFrame")
        self._form_main_frame = self.main_frame  # Add reference for consistency with main app
        self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        
        # Configure resizing
        self.main_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # Create the three main sections
        self._create_top_frame()
        self._create_middle_frame()
        self._create_button_frame()
        
        # Mouse wheel bindings
        bind_mousewheel_recursive(self.main_frame, self._on_mousewheel)

    def _create_top_frame(self):
        # Top frame with two columns
        self.top_frame = ttk.Frame(self.main_frame, style="TFrame", padding=10)
        self.top_frame.grid(row=0, column=0, sticky="ew", pady=5)
        self.top_frame.columnconfigure(0, weight=1)
        self.top_frame.columnconfigure(1, weight=1)
        
        # Left frame - Detail Surat
        self.frame_kiri = ttk.LabelFrame(self.top_frame, text="Detail Surat", padding="15", style="TLabelframe")
        self.frame_kiri.grid(row=0, column=0, sticky="nsew", padx=5)
        self.form_input_widgets.update(populate_frame_kiri(self.frame_kiri, self.vars))
        
        # Right frame - Klasifikasi
        self.frame_kanan = ttk.LabelFrame(self.top_frame, text="Klasifikasi", padding="15", style="TLabelframe")
        self.frame_kanan.grid(row=0, column=1, sticky="nsew", padx=5)
        # FIX: Store the returned widgets properly
        kanan_widgets = populate_frame_kanan(self.frame_kanan, self.vars)
        if isinstance(kanan_widgets, dict):
            self.form_input_widgets.update(kanan_widgets)
        # Store the tgl_terima_entry separately if it's returned directly
        elif kanan_widgets:
            self.tgl_terima_entry = kanan_widgets
            self.form_input_widgets["tgl_terima"] = kanan_widgets

    def _create_middle_frame(self):
        # Middle frame with three columns
        self.middle_frame = ttk.Frame(self.main_frame, style="TFrame", padding=10)
        self.middle_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        
        # Configure grid weights
        self.middle_frame.columnconfigure(0, weight=1)
        self.middle_frame.columnconfigure(1, weight=1)
        self.middle_frame.columnconfigure(2, weight=2)
        self.middle_frame.rowconfigure(0, weight=1)
        
        # Left column - Disposisi Kepada
        self.frame_disposisi = ttk.LabelFrame(self.middle_frame, text="Disposisi Kepada", padding="15", style="TLabelframe")
        self.frame_disposisi.grid(row=0, column=0, sticky="nsew", padx=5)
        populate_frame_disposisi(self.frame_disposisi, self.vars)
        
        # Middle column - Untuk di
        self.frame_instruksi = ttk.LabelFrame(self.middle_frame, text="Untuk di", padding="15", style="TLabelframe")
        self.frame_instruksi.grid(row=0, column=1, sticky="nsew", padx=5)
        instruksi_widgets = populate_frame_instruksi(self.frame_instruksi, self.vars)
        self.form_input_widgets.update(instruksi_widgets)
        # Store harap_selesai_tgl_entry if it exists
        if "harap_selesai_tgl" in instruksi_widgets:
            self.harap_selesai_tgl_entry = instruksi_widgets["harap_selesai_tgl"]
        
        # Right column - Isi Instruksi/Informasi
        self.frame_info = ttk.LabelFrame(self.middle_frame, text="Isi Instruksi / Informasi", padding="15", style="TLabelframe")
        self.frame_info.grid(row=0, column=2, sticky="nsew", padx=5)
        
        # Instruction table inside frame_info
        self.posisi_options = [
            "Direktur Utama", "Direktur Keuangan", "Direktur Teknik",
            "GM Keuangan & Administrasi", "GM Operasional & Pemeliharaan", "Manager"
        ]
        self.frame_instruksi_table = ttk.Frame(self.frame_info)
        self.frame_instruksi_table.grid(row=0, column=0, sticky="nsew")
        self.instruksi_table = InstruksiTable(self.frame_instruksi_table, self.posisi_options, use_grid=True)
        
        # Tombol tambah/hapus baris
        self.frame_instruksi_btn = ttk.Frame(self.frame_info)
        self.frame_instruksi_btn.grid(row=1, column=0, sticky="ew", pady=(5,0))
        self.btn_tambah_baris = ttk.Button(self.frame_instruksi_btn, text="+ Tambah Baris", command=self._on_tambah_baris, style="Secondary.TButton")
        self.btn_tambah_baris.pack(side="left", padx=2)
        self.btn_hapus_baris = ttk.Button(self.frame_instruksi_btn, text="🗑 Hapus Baris Terpilih", command=self._on_hapus_baris, style="Secondary.TButton")
        self.btn_hapus_baris.pack(side="left", padx=2)
        self.btn_kosongkan_baris = ttk.Button(self.frame_instruksi_btn, text="⏹ Kosongkan Baris", command=self._on_kosongkan_baris, style="Secondary.TButton")
        self.btn_kosongkan_baris.pack(side="left", padx=2)

    def _create_button_frame(self):
        # Create edit-specific buttons
        btn_frame = ttk.Frame(self.main_frame)
        btn_frame.grid(row=2, column=0, sticky="ew", pady=10)
        
        # Save button
        btn_save = ttk.Button(btn_frame, text="💾 Simpan Perubahan", 
                             command=self._on_save, style="Success.TButton")
        btn_save.pack(side="left", padx=5)
        
        # Export PDF button
        btn_pdf = ttk.Button(btn_frame, text="📄 Export ke PDF", 
                            command=self._on_export_pdf, style="Primary.TButton")
        btn_pdf.pack(side="left", padx=5)
        
        # Send Email button - use simplified approach
        btn_email = ttk.Button(btn_frame, text="📧 Kirim Email", 
                              command=self._on_send_email_simple, style="Primary.TButton")
        btn_email.pack(side="left", padx=5)
        
        # Cancel button
        btn_cancel = ttk.Button(btn_frame, text="❌ Batal", 
                               command=self._on_cancel, style="Secondary.TButton")
        btn_cancel.pack(side="right", padx=5)

    def get_disposisi_labels(self):
        """Get selected disposisi labels - compatible with main app"""
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
            if var in self.vars and self.vars[var].get():
                labels.append(label)
        return labels

    def _on_send_email_simple(self):
        """Simplified email sending without complex dialogs"""
        try:
            disposisi_labels = self.get_disposisi_labels()
            if not disposisi_labels:
                messagebox.showwarning("Peringatan", "Pilih minimal satu disposisi kepada untuk mengirim email.")
                return
            
            # Simple confirmation dialog
            recipients_str = ", ".join(disposisi_labels)
            confirm = messagebox.askyesno(
                "Konfirmasi Kirim Email", 
                f"Kirim email disposisi kepada:\n{recipients_str}\n\nLanjutkan?"
            )
            
            if not confirm:
                return
            
            # Send email directly
            self._send_email_to_positions(disposisi_labels)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Gagal mengirim email: {e}")

    def _send_email_to_positions(self, positions):
        """Send email to specific positions"""
        try:
            # Import email sender
            from email_sender.send_email import EmailSender
            from email_sender.template_handler import render_email_template
            from datetime import datetime
            
            # Collect current form data
            data = collect_form_data_safely(self)
            
            # Prepare email data
            klasifikasi = []
            if data.get("rahasia", 0):
                klasifikasi.append("RAHASIA")
            if data.get("penting", 0):
                klasifikasi.append("PENTING") 
            if data.get("segera", 0):
                klasifikasi.append("SEGERA")
            
            # Get instruksi text
            instruksi_list = []
            
            # Add untuk di items
            untuk_di_items = []
            if data.get("ketahui_file", 0):
                untuk_di_items.append("Ketahui & File")
            if data.get("proses_selesai", 0):
                untuk_di_items.append("Proses Selesai")
            if data.get("teliti_pendapat", 0):
                untuk_di_items.append("Teliti & Pendapat")
            if data.get("buatkan_resume", 0):
                untuk_di_items.append("Buatkan Resume")
            if data.get("edarkan", 0):
                untuk_di_items.append("Edarkan")
            if data.get("sesuai_disposisi", 0):
                untuk_di_items.append("Sesuai Disposisi")
            if data.get("bicarakan_saya", 0):
                untuk_di_items.append("Bicarakan dengan Saya")
            
            if untuk_di_items:
                instruksi_list.append(f"Untuk di: {', '.join(untuk_di_items)}")
            
            # Add specific instructions from the table
            if data.get("isi_instruksi"):
                for instr in data["isi_instruksi"]:
                    if instr.get("instruksi", "").strip():
                        posisi = instr.get("posisi", "")
                        instruksi_text = instr.get("instruksi", "")
                        tanggal = instr.get("tanggal", "")
                        
                        instr_line = f"{posisi}: {instruksi_text}"
                        if tanggal:
                            instr_line += f" (Tanggal: {tanggal})"
                        instruksi_list.append(instr_line)
            
            # Add bicarakan dengan if exists
            if data.get("bicarakan_dengan", "").strip():
                instruksi_list.append(f"Bicarakan dengan: {data['bicarakan_dengan']}")
            
            # Add teruskan kepada if exists
            if data.get("teruskan_kepada", "").strip():
                instruksi_list.append(f"Teruskan kepada: {data['teruskan_kepada']}")
            
            # Add deadline if exists
            if data.get("harap_selesai_tgl", "").strip():
                instruksi_list.append(f"Harap diselesaikan tanggal: {data['harap_selesai_tgl']}")
            
            template_data = {
                'nomor_surat': data.get("no_surat", ""),
                'nama_pengirim': data.get("asal_surat", ""),
                'perihal': data.get("perihal", ""),
                'tanggal': datetime.now().strftime('%d %B %Y'),
                'klasifikasi': klasifikasi,
                'instruksi_list': instruksi_list,
                'tahun': datetime.now().year
            }
            
            # Render HTML content
            html_content = render_email_template(template_data)
            subject = f"Disposisi Surat: {data.get('perihal', 'N/A')}"
            
            # Initialize email sender and send
            email_sender = EmailSender()
            success, message, details = email_sender.send_disposisi_to_positions(
                positions, 
                subject, 
                html_content
            )
            
            # Show results
            if success:
                success_msg = f"Email berhasil dikirim!\n\n{message}"
                if details.get('failed_lookups'):
                    success_msg += "\n\nCatatan: Beberapa posisi tidak memiliki email yang valid di database."
                # Use root window as parent to avoid the messagebox error
                messagebox.showinfo("Email Sent", success_msg, parent=self.winfo_toplevel())
            else:
                error_msg = f"Gagal mengirim email:\n{message}"
                if details.get('failed_lookups'):
                    error_msg += "\n\nPosisi tanpa email:\n" + "\n".join(details['failed_lookups'])
                messagebox.showerror("Email Error", error_msg, parent=self.winfo_toplevel())
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Email Error", f"Terjadi kesalahan: {e}", parent=self.winfo_toplevel())

    def update_status(self, message):
        """Update status message - for compatibility"""
        print(f"[EditTab] Status: {message}")

    def _on_mousewheel(self, event):
        # Handle mouse wheel scrolling
        if event.num == 5 or event.delta == -120:
            self.canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta == 120:
            self.canvas.yview_scroll(-1, "units")
        else:
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_shift_mousewheel(self, event):
        # Handle shift+mouse wheel for horizontal scrolling
        self.canvas.xview_scroll(int(-1*(event.delta/120)), "units")

    def _parse_date_string(self, date_str):
        """Parse date string in various formats and return in dd-mm-yyyy format"""
        if not date_str or date_str.strip() == "":
            return ""
        
        date_str = str(date_str).strip()
        
        # Common date formats to try
        date_formats = [
            "%d-%m-%Y",  # 31-12-2024
            "%d/%m/%Y",  # 31/12/2024
            "%Y-%m-%d",  # 2024-12-31
            "%Y/%m/%d",  # 2024/12/31
            "%d-%m-%y",  # 31-12-24
            "%d/%m/%y",  # 31/12/24
            "%d %B %Y",  # 31 December 2024
            "%d %b %Y",  # 31 Dec 2024
            "%B %d, %Y", # December 31, 2024
            "%b %d, %Y", # Dec 31, 2024
        ]
        
        for fmt in date_formats:
            try:
                dt = datetime.strptime(date_str, fmt)
                return dt.strftime("%d-%m-%Y")
            except ValueError:
                continue
        
        # If no format matches, return the original string
        print(f"[EditTab] Warning: Could not parse date '{date_str}'")
        return date_str

    def _fill_form_from_log(self, data):
        print("[EditTab] Mengisi form dari data log:", data)
        # Mapping dari header Google Sheets ke key pythonic
        key_map = {
            "No. Agenda": "no_agenda",
            "No. Surat": "no_surat",
            "Tgl. Surat": "tgl_surat",
            "Perihal": "perihal",
            "Asal Surat": "asal_surat",
            "Ditujukan": "ditujukan",
            "Kode Klasifikasi": "kode_klasifikasi",
            "Tgl. Penerimaan": "tgl_terima",
            "Indeks": "indeks",
            "Bicarakan dengan": "bicarakan_dengan",
            "Teruskan kepada": "teruskan_kepada",
            "Harap Selesai Tanggal": "harap_selesai_tgl",
        }
        # Klasifikasi (RAHASIA/PENTING/SEGERA)
        klasifikasi_val = data.get("Klasifikasi", "").upper() if "Klasifikasi" in data else data.get("klasifikasi", "").upper()
        data["rahasia"] = 1 if "RAHASIA" in klasifikasi_val else 0
        data["penting"] = 1 if "PENTING" in klasifikasi_val else 0
        data["segera"] = 1 if "SEGERA" in klasifikasi_val else 0
        # Checkbox disposisi
        disposisi_map = {
            "Direktur Utama": "dir_utama",
            "Direktur Keuangan": "dir_keu",
            "Direktur Teknik": "dir_teknik",
            "GM Keuangan & Administrasi": "gm_keu",
            "GM Operasional & Pemeliharaan": "gm_ops",
            "Manager": "manager"
        }
        for label, key in disposisi_map.items():
            data[key] = 1 if label in data.get("Disposisi kepada", "") else 0
        # Checkbox untuk di
        untuk_di_map = {
            "Ketahui & File": "ketahui_file",
            "Proses Selesai": "proses_selesai",
            "Teliti & Pendapat": "teliti_pendapat",
            "Buatkan Resume": "buatkan_resume",
            "Edarkan": "edarkan",
            "Sesuai Disposisi": "sesuai_disposisi",
            "Bicarakan dengan Saya": "bicarakan_saya"
        }
        untuk_di_val = data.get("Untuk Di :", "")
        for label, key in untuk_di_map.items():
            data[key] = 1 if label in untuk_di_val else 0
        # Mapping field lain
        for sheet_key, py_key in key_map.items():
            if sheet_key in data:
                data[py_key] = data[sheet_key]
        # Instruksi jabatan
        if "isi_instruksi" not in data or not data["isi_instruksi"]:
            instruksi_jabatan_map = [
                ("Direktur Utama Instruksi", "Direktur Utama", "Direktur Utama Tanggal"),
                ("Direktur Keuangan Instruksi", "Direktur Keuangan", "Direktur Keuangan Tanggal"),
                ("Direktur Teknik Instruksi", "Direktur Teknik", "Direktur Teknik Tanggal"),
                ("GM Keuangan & Administrasi Instruksi", "GM Keuangan & Administrasi", "GM Keuangan & Administrasi Tanggal"),
                ("GM Operasional & Pemeliharaan Instruksi", "GM Operasional & Pemeliharaan", "GM Operasional & Pemeliharaan Tanggal"),
                ("Manager Instruksi", "Manager", "Manager Tanggal")
            ]
            instruksi_from_log = []
            for instr_col, posisi_label, tgl_col in instruksi_jabatan_map:
                instruksi_val = data.get(instr_col, "").strip()
                tgl_val = data.get(tgl_col, "").strip()
                if instruksi_val or tgl_val:
                    instruksi_from_log.append({
                        "posisi": posisi_label,
                        "instruksi": instruksi_val,
                        "tanggal": tgl_val
                    })
            data["isi_instruksi"] = instruksi_from_log
        
        # Isi variabel dan widget dari data log
        for key, var in self.vars.items():
            val = data.get(key, "")
            if isinstance(var, tk.IntVar):
                try:
                    var.set(int(val) if val not in (None, "") else 0)
                except Exception:
                    var.set(0)
            else:
                var.set(val if val is not None else "")
        
        # Widget khusus - Text widgets
        for key in ["perihal", "asal_surat", "ditujukan", "bicarakan_dengan", "teruskan_kepada"]:
            if key in self.form_input_widgets:
                widget = self.form_input_widgets[key]
                widget.delete("1.0", tk.END)
                widget.insert("1.0", data.get(key, ""))
        
        # FIX: Improved date handling for DateEntry widgets - ALLOW NULL VALUES
        date_fields = [
            ("tgl_surat", "tgl_surat"),
            ("tgl_terima", "tgl_terima"), 
            ("harap_selesai_tgl", "harap_selesai_tgl")
        ]
        
        for widget_key, data_key in date_fields:
            # Look for widget in multiple locations
            widget = None
            if widget_key in self.form_input_widgets:
                widget = self.form_input_widgets[widget_key]
            elif hasattr(self, f'{widget_key}_entry'):
                widget = getattr(self, f'{widget_key}_entry')
            
            if widget:
                date_value = data.get(data_key, "") or data.get(widget_key.replace("_", " ").title(), "")
                
                try:
                    # Clear widget first
                    widget.delete(0, tk.END)
                    
                    # Only set date if value exists and is not empty
                    if date_value and str(date_value).strip():
                        # Parse the date string to ensure it's in the correct format
                        parsed_date = self._parse_date_string(date_value)
                        if parsed_date and parsed_date != date_value:
                            # Convert to datetime object for DateEntry
                            dt = datetime.strptime(parsed_date, "%d-%m-%Y")
                            widget.set_date(dt)
                        else:
                            # Insert as text if parsing fails but value exists
                            widget.insert(0, str(date_value))
                    # If empty, leave widget empty (null values allowed)
                    
                except Exception as e:
                    print(f"[EditTab] Error setting date for {widget_key}: {e}")
                    # Clear the widget on error
                    try:
                        widget.delete(0, tk.END)
                    except:
                        pass
        
        # Instruksi table
        if "isi_instruksi" in data:
            self.instruksi_table.data = data["isi_instruksi"]
            self.instruksi_table.render_table()
        print("[EditTab] Form berhasil diisi dari data log.")

    def _on_save(self):
        import threading
        def do_save():
            print("[EditTab] Mulai proses simpan edit.")
            
            # Collect form data safely using the improved function
            from disposisi_app.views.components.export_utils import collect_form_data_safely
            try:
                data_baru = collect_form_data_safely(self)
            except Exception as e:
                print(f"[EditTab][ERROR] Error collecting form data: {e}")
                messagebox.showerror("Error", f"Gagal mengumpulkan data form: {e}")
                return
            
            # Enhanced validation with null handling
            def safe_get_value(data, key, default=""):
                """Safely get value from data dict"""
                try:
                    value = data.get(key, default)
                    if value is None:
                        return default
                    return str(value).strip() if value != default else default
                except:
                    return default
            
            # FIX: Allow null values for all fields except no_surat
            # Convert None values to appropriate defaults
            nullable_fields = [
                "no_agenda", "perihal", "asal_surat", "ditujukan",
                "kode_klasifikasi", "indeks", "bicarakan_dengan", 
                "teruskan_kepada", "tgl_surat", "tgl_terima", "harap_selesai_tgl"
            ]
            
            # Ensure null safety for all fields
            for field in nullable_fields:
                if field in data_baru:
                    if data_baru[field] is None:
                        data_baru[field] = ""  # Convert None to empty string
                    else:
                        data_baru[field] = str(data_baru[field]).strip()
            
            print("[EditTab] Data baru yang akan diupdate:", {k: v for k, v in data_baru.items() if k != "isi_instruksi"})
            if "isi_instruksi" in data_baru:
                print("[EditTab] Instruksi data:", data_baru["isi_instruksi"])

            # VALIDASI: Hanya No. Surat yang wajib diisi
            no_surat_baru = safe_get_value(data_baru, "no_surat")
            if not no_surat_baru:
                print(f"[EditTab][ERROR] Field 'No. Surat' kosong!")
                messagebox.showerror("Error", "Field 'No. Surat' tidak boleh kosong!")
                return
            
            # Cek duplikasi No. Surat (kecuali data yang sedang diedit)
            try:
                from google_sheets_connect import get_sheets_service, SHEET_ID
                service = get_sheets_service()
                range_name = 'Sheet1!A6:B'
                result = service.spreadsheets().values().get(
                    spreadsheetId=SHEET_ID,
                    range=range_name
                ).execute()
                values = result.get('values', [])
                no_surat_lama = safe_get_value(self.data_log, "No. Surat")
                
                for row in values:
                    if len(row) > 1:
                        no_surat = str(row[1]).strip()
                        if no_surat and no_surat == no_surat_baru and no_surat != no_surat_lama:
                            messagebox.showerror("Error", f"No. Surat '{no_surat_baru}' sudah ada di data lain!")
                            return
            except Exception as e:
                print(f"[EditTab][WARNING] Tidak bisa cek duplikasi No. Surat: {e}")
                # Continue even if duplicate check fails to avoid data loss

            try:
                from disposisi_app.views.components.loading_screen import LoadingScreen
                loading = LoadingScreen(self)
                
                # Progress simulation
                for i in range(1, 101):
                    import time; time.sleep(0.01)
                    loading.update_progress(i)
                
                print("[EditTab] Memanggil update_log_entry...")
                
                # Use the fixed update_log_entry function
                from edit_logic import update_log_entry
                success = update_log_entry(self.data_log, data_baru)
                
                if success:
                    print("[EditTab] Update berhasil.")
                    messagebox.showinfo("Sukses", "Data berhasil diupdate.")
                    
                    # Call callback if available
                    if self.on_save_callback:
                        print("[EditTab] Memanggil on_save_callback...")
                        self.on_save_callback()
                else:
                    print("[EditTab] Update gagal.")
                    messagebox.showerror("Error", "Gagal update data ke Google Sheets.")
                
                loading.destroy()
                
            except Exception as e:
                import traceback; traceback.print_exc()
                print(f"[EditTab][ERROR] Gagal update data: {e}")
                messagebox.showerror("Error", f"Gagal update data: {e}")
                if 'loading' in locals():
                    loading.destroy()
                    
        # Run in separate thread to avoid blocking UI
        threading.Thread(target=do_save, daemon=True).start()
        
    def _on_cancel(self):
        # Tutup tab edit (akan di-handle oleh FormApp)
        parent = self.master
        # Pastikan parent adalah Notebook sebelum memanggil forget
        from tkinter import ttk
        if isinstance(parent, ttk.Notebook):
            parent.forget(self) 

    def _on_export_pdf(self):
        import threading
        from tkinter import filedialog
        from pdf_output import save_form_to_pdf, merge_pdfs
        import tempfile, os, traceback
        def do_export():
            try:
                filepath = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF Documents", "*.pdf"), ("All Files", "*.*")],
                    title="Export Edit ke PDF"
                )
                if not filepath:
                    return
                # Ambil data dari form menggunakan collect_form_data_safely
                data_baru = collect_form_data_safely(self)
                
                # Simpan PDF ke file sementara
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    temp_pdf_path = temp_pdf.name
                save_form_to_pdf(temp_pdf_path, data_baru)
                # Gabungkan PDF disposisi + lampiran jika ada
                pdf_list = [temp_pdf_path] + list(self.pdf_attachments)
                merge_pdfs(pdf_list, filepath)
                os.remove(temp_pdf_path)
                # Upload ke Google Sheets juga
                try:
                    update_log_entry(self.data_log, data_baru)
                except Exception as e:
                    import traceback; traceback.print_exc()
                    messagebox.showerror("Google Sheets", f"Gagal upload ke Google Sheets: {e}")
                messagebox.showinfo("Sukses", f"Data edit berhasil diekspor ke PDF:\n{filepath}")
            except Exception as e:
                traceback.print_exc()
                messagebox.showerror("Export PDF", f"Gagal ekspor ke PDF: {e}")
        threading.Thread(target=do_export).start() 

    def _on_tambah_baris(self):
        self.instruksi_table.add_row()

    def _on_hapus_baris(self):
        self.instruksi_table.remove_selected_rows() 

    def _on_kosongkan_baris(self):
        self.instruksi_table.kosongkan_semua_baris()
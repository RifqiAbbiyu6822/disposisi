import tkinter as tk
from tkinter import ttk, Text, messagebox
from tkcalendar import DateEntry
from logic.instruksi_table import InstruksiTable
from sheet_logic import update_log_entry
from disposisi_app.views.components.form_sections import (
    populate_frame_kiri, populate_frame_kanan, populate_frame_disposisi, populate_frame_instruksi
)

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
        # FIX: Update form_input_widgets with widgets returned from populate_frame_kanan
        self.form_input_widgets.update(populate_frame_kanan(self.frame_kanan, self.vars))

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
        self.form_input_widgets.update(populate_frame_instruksi(self.frame_instruksi, self.vars))
        
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
        # Use the new unified button frame from button_frame.py
        from disposisi_app.views.components.button_frame import create_button_frame
        def on_selesai():
            # Show preview and export/email options
            if hasattr(self, 'preview_frame'):
                self.preview_frame.grid()
        def export_pdf():
            self._on_export_pdf()
        def upload_sheet():
            self._on_save()
        def on_remove_pdf(idx):
            if 0 <= idx < len(self.pdf_attachments):
                del self.pdf_attachments[idx]
                # Re-render button frame to update attachments
                self.button_frame.destroy()
                self._create_button_frame()
        callbacks = {
            "on_selesai": on_selesai,
            "export_pdf": export_pdf,
            "upload_sheet": upload_sheet,
        }
        self.button_frame = create_button_frame(
            self.main_frame,
            callbacks,
            self.pdf_attachments,
            on_remove_pdf
        )

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
        # Widget khusus
        for key in ["perihal", "asal_surat", "ditujukan", "bicarakan_dengan", "teruskan_kepada"]:
            if key in self.form_input_widgets:
                widget = self.form_input_widgets[key]
                widget.delete("1.0", tk.END)
                widget.insert("1.0", data.get(key, ""))
        
        # FIX: Use proper set_date() method for DateEntry widgets
        for key in ["tgl_surat", "tgl_terima", "harap_selesai_tgl"]:
            if key in self.form_input_widgets:
                widget = self.form_input_widgets[key]
                date_value = data.get(key, "")
                if date_value:
                    try:
                        # Try to set the date using set_date method
                        widget.set_date(date_value)
                    except Exception as e:
                        print(f"[EditTab] Warning: Could not set date for {key}: {e}")
                        # Fallback: clear the date if setting fails
                        widget.set_date("")
                else:
                    # Clear the date if no value
                    widget.set_date("")
        
        # Instruksi table
        if "isi_instruksi" in data:
            self.instruksi_table.data = data["isi_instruksi"]
            self.instruksi_table.render_table()
        print("[EditTab] Form berhasil diisi dari data log.")

    def _on_save(self):
        import threading
        def do_save():
            print("[EditTab] Mulai proses simpan edit.")
            # Ambil data dari form
            data_baru = {key: var.get() for key, var in self.vars.items()}
            for key in ["perihal", "asal_surat", "ditujukan", "bicarakan_dengan", "teruskan_kepada"]:
                if key in self.form_input_widgets:
                    data_baru[key] = self.form_input_widgets[key].get("1.0", tk.END).strip()
            
            # FIX: Remove fallback logic that prevented date clearing
            for key in ["tgl_surat", "tgl_terima", "harap_selesai_tgl"]:
                if key in self.form_input_widgets:
                    # Simply get the current value from the widget
                    data_baru[key] = self.form_input_widgets[key].get()
            
            data_baru["isi_instruksi"] = self.instruksi_table.get_data()
            print("[EditTab] Data baru yang akan diupdate:", data_baru)

            # Validasi hanya No. Surat yang wajib diisi dan harus unik
            no_surat_baru = data_baru.get("no_surat", "").strip()
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
                no_surat_lama = str(self.data_log.get("No. Surat", "")).strip()
                for row in values:
                    if len(row) > 1:
                        no_surat = str(row[1]).strip()
                        if no_surat and no_surat == no_surat_baru and no_surat != no_surat_lama:
                            messagebox.showerror("Error", f"No. Surat '{no_surat_baru}' sudah ada di data lain!")
                            return
            except Exception as e:
                print(f"[EditTab][WARNING] Tidak bisa cek duplikasi No. Surat: {e}")
                # Jika gagal cek, tetap lanjutkan (opsional, bisa diubah)

            try:
                from coba import LoadingScreen
                loading = LoadingScreen(self)
                for i in range(1, 101):
                    import time; time.sleep(0.01)
                    loading.update_progress(i)
                print("[EditTab] Memanggil update_log_entry...")
                update_log_entry(self.data_log, data_baru)
                print("[EditTab] Update berhasil.")
                messagebox.showinfo("Sukses", "Data berhasil diupdate.")
                if self.on_save_callback:
                    print("[EditTab] Memanggil on_save_callback...")
                    self.on_save_callback()
                loading.destroy()
            except Exception as e:
                import traceback; traceback.print_exc()
                print(f"[EditTab][ERROR] Gagal update data: {e}")
                messagebox.showerror("Error", f"Gagal update data: {e}")
                if 'loading' in locals():
                    loading.destroy()
        threading.Thread(target=do_save).start()

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
                # Ambil data dari form (seperti _on_save)
                data_baru = {key: var.get() for key, var in self.vars.items()}
                for key in ["perihal", "asal_surat", "ditujukan", "bicarakan_dengan", "teruskan_kepada"]:
                    if key in self.form_input_widgets:
                        data_baru[key] = self.form_input_widgets[key].get("1.0", tk.END).strip()
                for key in ["tgl_surat", "tgl_terima", "harap_selesai_tgl"]:
                    if key in self.form_input_widgets:
                        data_baru[key] = self.form_input_widgets[key].get()
                data_baru["isi_instruksi"] = self.instruksi_table.get_data()
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
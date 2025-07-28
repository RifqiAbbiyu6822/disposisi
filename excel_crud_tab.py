import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import logging
from google_sheets_connect import get_sheets_service, SHEET_ID
from constants import FIELD_LABELS, ALL_FIELDS
import csv
import os
import time
import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from disposisi_app.views.components.loading_screen import LoadingScreen
from disposisi_app.views.components.status_utils import update_status
from disposisi_app.views.components.tooltip_utils import attach_tooltip, add_tooltips
from disposisi_app.views.components.dialogs import show_shortcuts, show_about
from disposisi_app.views.components.export_utils import save_to_pdf, save_to_sheet
from disposisi_app.views.components.form_utils import clear_form
from disposisi_app.views.components.constants import POSISI_OPTIONS, TOOLTIP_LABELS
from disposisi_app.views.components.styles import setup_styles

logging.basicConfig(level=logging.WARNING)

SHEET_NAME = 'Sheet1'

# Enhanced header that matches the actual Google Sheets structure
# This header corresponds to the 28 columns (A-AB) in the Google Sheets
# Structure matches the write_multilayer_header function in google_sheets_connect.py
ENHANCED_HEADER = [
    "No. Agenda", "No. Surat", "Tgl. Surat", "Perihal", "Asal Surat", "Ditujukan", 
    "Klasifikasi", "Disposisi kepada", "Untuk Di :", "Selesai Tgl.", "Kode Klasifikasi", 
    "Tgl. Penerimaan", "Indeks", "Bicarakan dengan", "Teruskan kepada", "Harap Selesai Tanggal",
    "Direktur Utama Instruksi", "Direktur Utama Tanggal", "Direktur Keuangan Instruksi", "Direktur Keuangan Tanggal",
    "Direktur Teknik Instruksi", "Direktur Teknik Tanggal", "GM Keuangan & Administrasi Instruksi", "GM Keuangan & Administrasi Tanggal",
    "GM Operasional & Pemeliharaan Instruksi", "GM Operasional & Pemeliharaan Tanggal", "Manager Instruksi", "Manager Tanggal"
]

# Searchable columns (without position instruction columns)
SEARCHABLE_LABELS = [
    "No. Agenda", "No. Surat", "Tgl. Surat", "Perihal", "Asal Surat", "Ditujukan",
    "Klasifikasi", "Disposisi kepada", "Untuk Di :", "Selesai Tgl.", "Kode Klasifikasi",
    "Tgl. Penerimaan", "Indeks", "Bicarakan dengan", "Teruskan kepada", "Harap Selesai Tanggal"
]

class LogTab(ttk.Frame):
    _log_cache = None
    _last_cache_time = 0
    _cache_ttl = 60  # detik, cache berlaku 1 menit
    PAGE_SIZE = 20  # Jumlah data per halaman
    """
    Enhanced Log Tab with improved table styling and grid lines.
    Matches the Google Sheets structure with 28 columns (A-AB).
    Features:
    - Clear grid lines with improved styling
    - Alternating row colors for better readability
    - Tooltips for long text and column information
    - Status bar with record count and column info
    - Responsive column widths based on content type
    """
    def __init__(self, parent, on_edit_log=None):
        super().__init__(parent)
        setup_styles(self)  # Terapkan style global
        self.on_edit_log = on_edit_log
        self.pack(fill="both", expand=True)
        self.data = []  # Data log
        self.filtered_data = []
        self.tooltip = None  # Untuk tooltip tombol
        self.current_page = 1
        self.total_pages = 1
        self.page_size = LogTab.PAGE_SIZE
        self.create_widgets()
        self.setup_shortcuts()
        # self.load_sheet_data()  # Hapus auto-load di awal, load hanya saat tab aktif

    def setup_shortcuts(self):
        self.bind_all("<Control-f>", lambda e: self.focus_search())
        self.search_entry.bind("<Return>", lambda e: self.do_search())
        self.tree.bind("<Delete>", lambda e: self.delete_selected())
        self.tree.bind("<Return>", lambda e: self.edit_selected())
        self.tree.bind("<Up>", self.on_arrow_up)
        self.tree.bind("<Down>", self.on_arrow_down)
        self.tree.bind("<Shift-Up>", self.on_shift_arrow_up)
        self.tree.bind("<Shift-Down>", self.on_shift_arrow_down)

    def on_arrow_up(self, event):
        selected = self.tree.selection()
        items = self.tree.get_children()
        if not items:
            return
        if not selected:
            self.tree.selection_set(items[0])
            self.tree.focus(items[0])
            self.tree.see(items[0])
            return
        idx = list(items).index(selected[0]) if selected[0] in items else 0
        if idx > 0:
            new_item = items[idx-1]
            self.tree.selection_set(new_item)
            self.tree.focus(new_item)
            self.tree.see(new_item)
        return "break"

    def on_arrow_down(self, event):
        selected = self.tree.selection()
        items = self.tree.get_children()
        if not items:
            return
        if not selected:
            self.tree.selection_set(items[0])
            self.tree.focus(items[0])
            self.tree.see(items[0])
            return
        idx = list(items).index(selected[0]) if selected[0] in items else 0
        if idx < len(items)-1:
            new_item = items[idx+1]
            self.tree.selection_set(new_item)
            self.tree.focus(new_item)
            self.tree.see(new_item)
        return "break"

    def on_shift_arrow_up(self, event):
        selected = self.tree.selection()
        items = self.tree.get_children()
        if not items:
            return
        if not selected:
            self.tree.selection_set(items[0])
            self.tree.focus(items[0])
            self.tree.see(items[0])
            return "break"
        idxs = sorted([list(items).index(s) for s in selected if s in items])
        if not idxs:
            return "break"
        min_idx = idxs[0]
        max_idx = idxs[-1]
        focus_item = self.tree.focus() or selected[-1]
        focus_idx = list(items).index(focus_item) if focus_item in items else max_idx
        # Jika blok lebih dari satu dan fokus di bawah, kurangi dari bawah
        if len(selected) > 1 and focus_idx == max_idx:
            # Kurangi blok dari bawah
            new_selection = list(selected)
            new_selection.remove(items[max_idx])
            self.tree.selection_set(new_selection)
            self.tree.focus(items[max_idx-1])
            self.tree.see(items[max_idx-1])
        elif min_idx > 0:
            # Perluas blok ke atas
            new_item = items[min_idx-1]
            new_selection = list(selected) + [new_item]
            new_selection = list(dict.fromkeys(new_selection))
            self.tree.selection_set(new_selection)
            self.tree.focus(new_item)
            self.tree.see(new_item)
        return "break"

    def on_shift_arrow_down(self, event):
        selected = self.tree.selection()
        items = self.tree.get_children()
        if not items:
            return
        if not selected:
            self.tree.selection_set(items[0])
            self.tree.focus(items[0])
            self.tree.see(items[0])
            return "break"
        idxs = sorted([list(items).index(s) for s in selected if s in items])
        if not idxs:
            return "break"
        min_idx = idxs[0]
        max_idx = idxs[-1]
        focus_item = self.tree.focus() or selected[-1]
        focus_idx = list(items).index(focus_item) if focus_item in items else max_idx
        # Jika blok lebih dari satu dan fokus di atas, kurangi dari atas
        if len(selected) > 1 and focus_idx == min_idx:
            # Kurangi blok dari atas
            new_selection = list(selected)
            new_selection.remove(items[min_idx])
            self.tree.selection_set(new_selection)
            self.tree.focus(items[min_idx+1])
            self.tree.see(items[min_idx+1])
        elif max_idx < len(items)-1:
            # Perluas blok ke bawah
            new_item = items[max_idx+1]
            new_selection = list(selected) + [new_item]
            new_selection = list(dict.fromkeys(new_selection))
            self.tree.selection_set(new_selection)
            self.tree.focus(new_item)
            self.tree.see(new_item)
        return "break"

    def focus_search(self):
        self.search_entry.focus_set()

    def create_widgets(self):
        # --- FRAME ATAS: SEARCH, FILTER, TAHUN/BULAN ---
        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", pady=(10, 0), padx=10)

        # Search group
        search_group = ttk.LabelFrame(top_frame, text="Pencarian", padding=(8, 4))
        search_group.pack(side="left", padx=(0, 12))
        ttk.Label(search_group, text="Kolom:").pack(side="left", padx=(0, 4))
        self.search_col = ttk.Combobox(search_group, values=SEARCHABLE_LABELS, state="readonly", width=18)
        self.search_col.current(0)
        self.search_col.pack(side="left", padx=(0, 4))
        self.search_entry = ttk.Entry(search_group, width=20)
        self.search_entry.pack(side="left", padx=(0, 4))
        ttk.Button(search_group, text="Cari", command=self.do_search, style="Primary.TButton").pack(side="left")

        # Filter group
        filter_group = ttk.LabelFrame(top_frame, text="Filter", padding=(8, 4))
        filter_group.pack(side="left", padx=(0, 12))
        ttk.Label(filter_group, text="Kolom:").pack(side="left", padx=(0, 4))
        self.filter_col = ttk.Combobox(filter_group, values=SEARCHABLE_LABELS, state="readonly", width=18)
        self.filter_col.current(0)
        self.filter_col.pack(side="left", padx=(0, 4))
        self.filter_val = ttk.Combobox(filter_group, values=[], state="readonly", width=18)
        self.filter_val.pack(side="left", padx=(0, 4))
        ttk.Button(filter_group, text="Terapkan", command=self.do_filter, style="Primary.TButton").pack(side="left", padx=(0, 4))
        ttk.Button(filter_group, text="Clear", command=self.clear_filter, style="Secondary.TButton").pack(side="left")

        # Tahun & Bulan group
        date_group = ttk.LabelFrame(top_frame, text="Tahun & Bulan", padding=(8, 4))
        date_group.pack(side="left")
        ttk.Label(date_group, text="Tahun:").pack(side="left", padx=(0, 4))
        self.year_var = tk.StringVar()
        self.year_combo = ttk.Combobox(date_group, textvariable=self.year_var, state="readonly", width=6, values=[])
        self.year_combo.pack(side="left", padx=(0, 4))
        ttk.Label(date_group, text="Bulan:").pack(side="left", padx=(0, 4))
        self.month_var = tk.StringVar()
        self.month_combo = ttk.Combobox(date_group, textvariable=self.month_var, state="readonly", width=10, values=[])
        self.month_combo.pack(side="left", padx=(0, 4))
        self.year_combo.bind('<<ComboboxSelected>>', self.apply_year_month_filter)
        self.month_combo.bind('<<ComboboxSelected>>', self.apply_year_month_filter)

        # --- STATUS BAR ---
        status_frame = ttk.Frame(self)
        status_frame.pack(fill="x", padx=10, pady=(2, 12))  # Tambah padding bawah agar header tabel tidak tertutup
        self.status_label = ttk.Label(status_frame, text="Ready", foreground="gray")
        self.status_label.pack(side="left", padx=(0, 16))
        self.record_count_label = ttk.Label(status_frame, text="0 records", foreground="blue", font=("Arial", 8))
        self.record_count_label.pack(side="left", padx=(0, 16))
        self.column_info_label = ttk.Label(status_frame, text="28 columns", foreground="green", font=("Arial", 8))
        self.column_info_label.pack(side="left", padx=(0, 16))
        self.structure_info = ttk.Label(status_frame, text="Matches Google Sheets", foreground="purple", font=("Arial", 7))
        self.structure_info.pack(side="left")

        # --- FRAME TABEL ---
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=16, pady=6)
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview",
                       background="#f4f6fa",
                       foreground="#1a2233",
                       rowheight=30,
                       fieldbackground="#f4f6fa",
                       borderwidth=0,
                       relief="flat",
                       font=("Segoe UI", 11))
        style.configure("Treeview.Heading",
                       background="#2563eb",
                       foreground="#fff",
                       relief="flat",
                       borderwidth=0,
                       font=('Segoe UI', 12, 'bold'))
        style.map("Treeview.Heading",
                  background=[('active', '#0D234E'), ('!active', '#2563eb')],
                  foreground=[('active', '#fff'), ('!active', '#fff')])
        style.map("Treeview",
                 background=[('selected', '#2563eb')],
                 foreground=[('selected', '#fff')])
        self.tree = ttk.Treeview(table_frame, columns=ENHANCED_HEADER, show="headings", selectmode="extended", height=18)
        self.tree['show'] = 'headings'  # Pastikan heading selalu aktif
        v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        for col in ENHANCED_HEADER:
            display_name = self.get_column_display_name(col)
            self.tree.heading(col, text=str(display_name) if display_name is not None else "")
            # Lebar kolom lebih besar agar header tidak terpotong
            self.tree.column(col, width=180, anchor="center", minwidth=120, stretch=True)

        # --- FRAME BAWAH: TOMBOL AKSI & PAGING ---
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill="x", pady=5, padx=10)

        # Kiri: tombol aksi
        btn_frame = ttk.Frame(bottom_frame)
        btn_frame.pack(side="left", anchor="w")
        btn_edit = ttk.Button(btn_frame, text="✏️ Edit", command=self.edit_selected, style="Primary.TButton")
        btn_edit.pack(side="left", padx=4)
        btn_hapus = ttk.Button(btn_frame, text="🗑️ Hapus", command=self.delete_selected, style="Primary.TButton")
        btn_hapus.pack(side="left", padx=4)
        btn_refresh = ttk.Button(btn_frame, text="🔄 Refresh", command=lambda: self.refresh_log_data(force_refresh=True), style="Secondary.TButton")
        btn_refresh.pack(side="left", padx=4)
        btn_export = ttk.Button(btn_frame, text="⬇️ Export Excel", command=self.export_excel, style="Primary.TButton")
        btn_export.pack(side="left", padx=4)
        attach_tooltip(btn_edit, "Edit data terpilih (Enter)")
        attach_tooltip(btn_hapus, "Hapus data terpilih (Delete)")
        attach_tooltip(btn_refresh, "Refresh data log")
        attach_tooltip(btn_export, "Export data ke Excel")

        # Kanan: tombol paging
        paging_frame = ttk.Frame(bottom_frame)
        paging_frame.pack(side="right", anchor="e")
        self.btn_prev = ttk.Button(paging_frame, text="⏮️ Sebelumnya", command=self.go_prev_page, style="Secondary.TButton")
        self.btn_prev.pack(side="left", padx=4)
        self.paging_info_label = ttk.Label(paging_frame, text="Halaman 1 dari 1", font=("Arial", 9))
        self.paging_info_label.pack(side="left", padx=8)
        self.btn_next = ttk.Button(paging_frame, text="Berikutnya ⏭️", command=self.go_next_page, style="Secondary.TButton")
        self.btn_next.pack(side="left", padx=4)
        self.filter_col.bind('<<ComboboxSelected>>', self.update_filter_values)
        self.apply_table_styling()

    def apply_table_styling(self):
        """
        Apply additional styling to make grid lines clearer and improve readability.
        Only set alternating row colors here.
        """
        self.tree.tag_configure('oddrow', background='#f4f6fa')
        self.tree.tag_configure('evenrow', background='#ffffff')
        self.add_custom_borders()
    
    def add_custom_borders(self):
        """Add custom border styling to make grid lines more visible."""
        # This function can be extended to add more custom border styling
        # For now, we rely on the ttk.Style configuration above
        pass
    
    def on_click(self, event):
        """Handle click events on the treeview."""
        # Clear any existing tooltip
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None
    
    def on_motion(self, event):
        """Handle mouse motion events for tooltip."""
        # Get the item under the cursor
        item = self.tree.identify_row(event.y)
        if item:
            # Get the column under the cursor
            column = self.tree.identify_column(event.x)
            if column:
                # Get the value at this position
                values = self.tree.item(item, "values")
                col_index = int(column[1]) - 1  # Convert #1, #2, etc. to 0, 1, etc.
                if 0 <= col_index < len(ENHANCED_HEADER):
                    col_name = ENHANCED_HEADER[col_index]
                    value = values[col_index] if col_index < len(values) else ""
                    
                    # Show tooltip for long values or column info
                    if value and len(value) > 30:
                        tooltip_text = f"{col_name}:\n{value}"
                        self.show_tooltip(event, tooltip_text)
                    elif col_name in ["Perihal", "Asal Surat", "Bicarakan dengan", "Teruskan kepada"]:
                        # Show column info for important columns
                        col_info = self.get_column_info(col_name)
                        tooltip_text = f"{col_name}:\n{col_info}"
                        self.show_tooltip(event, tooltip_text)
    
    def show_tooltip(self, event, text):
        """Show a tooltip with the full text."""
        # Only show tooltip if not already showing for this cell
        if self.tooltip and getattr(self, '_last_tooltip_text', None) == text:
            return
        if self.tooltip:
            self.tooltip.destroy()
        self._last_tooltip_text = text
        self.tooltip = tk.Toplevel()
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
        frame = tk.Frame(self.tooltip, background="#ffffe0", relief=tk.SOLID, borderwidth=2)
        frame.pack(padx=2, pady=2)
        label = tk.Label(frame, text=text, justify=tk.LEFT,
                        background="#ffffe0", relief=tk.FLAT,
                        font=("Arial", 9), wraplength=300)
        label.pack(padx=5, pady=5)
        self.tooltip.after(4000, lambda: self.destroy_tooltip())
    
    def destroy_tooltip(self):
        """Destroy the tooltip if it exists."""
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

    def update_record_count(self):
        """Update the record count display."""
        total_records = len(self.data)
        filtered_records = len(self.filtered_data)
        
        if total_records == filtered_records:
            self.record_count_label.config(text=f"{total_records} records")
        else:
            self.record_count_label.config(text=f"{filtered_records}/{total_records} records")

    def clear_filter(self):
        """Clear all filters and show all data."""
        self.filtered_data = self.data.copy()
        self.refresh_table()
        self.update_filter_values()
        self.update_record_count()
        update_status(self.status_label, "Filter cleared")

    def load_sheet_data(self, force_refresh=False):
        import threading
        from disposisi_app.views.components.loading_screen import LoadingScreen
        def do_load(loading):
            try:
                update_status(self.status_label, "Loading data...")
                use_cache = (not force_refresh and LogTab._log_cache is not None and (time.time() - LogTab._last_cache_time) < LogTab._cache_ttl)
                retries = 3
                for attempt in range(retries):
                    try:
                        if use_cache:
                            values = LogTab._log_cache or []
                        else:
                            service = get_sheets_service()
                            range_name = f'{SHEET_NAME}!A6:AB'  # A to AB = 28 columns
                            result = service.spreadsheets().values().get(
                                spreadsheetId=SHEET_ID,
                                range=range_name
                            ).execute()
                            values = result.get('values', []) or []
                            # Filter baris yang benar-benar kosong (semua kolom kosong)
                            values = [row for row in values if any(cell.strip() for cell in row)]
                            LogTab._log_cache = values
                            LogTab._last_cache_time = time.time()
                        break
                    except Exception as e:
                        if attempt == retries - 1:
                            raise
                        time.sleep(1)
                MAX_ROWS = 500
                if len(values) > MAX_ROWS:
                    values = values[:MAX_ROWS]
                    truncated = True
                else:
                    truncated = False
                self.data = []
                tahun_set = set()
                bulan_set = set()
                for row in values:
                    row = row + ["" for _ in range(len(ENHANCED_HEADER) - len(row))]
                    data_row = {}
                    tgl_surat_val = None
                    for i, col in enumerate(ENHANCED_HEADER):
                        if i < len(row):
                            val = row[i]
                            # Konversi serial Excel ke tanggal string jika perlu
                            def excel_serial_to_date(val):
                                try:
                                    import datetime
                                    val_float = float(val)
                                    if val_float > 30000 and val_float < 90000 and str(int(val_float)) == str(val):
                                        # Excel 1900 date system
                                        base = datetime.datetime(1899, 12, 30)
                                        return (base + datetime.timedelta(days=int(val_float))).strftime('%d-%m-%Y')
                                except:
                                    pass
                                return val
                            if col in ["Tgl. Surat", "Selesai Tgl.", "Tgl. Penerimaan", "Harap Selesai Tanggal", "Direktur Utama Tanggal", "Direktur Keuangan Tanggal", "Direktur Teknik Tanggal", "GM Keuangan & Administrasi Tanggal", "GM Operasional & Pemeliharaan Tanggal", "Manager Tanggal"]:
                                val = excel_serial_to_date(val)
                            if col == "Tgl. Surat":
                                tgl_surat_val = val
                            if col == "Klasifikasi":
                                klasifikasi = val.strip().upper()
                                if klasifikasi in ["RAHASIA", "PENTING", "SEGERA"]:
                                    data_row[col] = klasifikasi
                                else:
                                    data_row[col] = ""
                            elif col == "Disposisi kepada":
                                disposisi = val.strip()
                                if disposisi:
                                    data_row[col] = disposisi
                                else:
                                    data_row[col] = ""
                            elif col == "Untuk Di :":
                                untuk_di = val.strip()
                                if untuk_di:
                                    data_row[col] = untuk_di
                                else:
                                    data_row[col] = ""
                            else:
                                data_row[col] = val if i < len(row) else ""
                        else:
                            data_row[col] = ""
                    # Tambahkan parsing tahun dan bulan dari Tgl. Surat
                    if tgl_surat_val:
                        tgl_split = tgl_surat_val.replace('/', '-').replace('.', '-').split('-')
                        if len(tgl_split) == 3:
                            if len(tgl_split[2]) == 4:
                                tahun_set.add(tgl_split[2])
                                bulan_set.add(tgl_split[1])
                            else:
                                tahun_set.add(tgl_split[0])
                                bulan_set.add(tgl_split[1])
                    self.data.append(data_row)
                self.filtered_data = self.data.copy()
                self.refresh_table()
                self.update_filter_values()
                self.update_record_count()
                tahun_list = sorted(list(tahun_set))
                bulan_list = sorted(list(bulan_set), key=lambda x: int(x) if x.isdigit() else 99)
                self.year_combo['values'] = [''] + tahun_list
                self.month_combo['values'] = [''] + bulan_list
                if truncated:
                    update_status(self.status_label, f"Loaded {len(self.data)} records (truncated)")
                    messagebox.showwarning("Load Log", f"Data log terlalu besar, hanya {MAX_ROWS} baris pertama yang dimuat.")
                else:
                    update_status(self.status_label, f"Loaded {len(self.data)} records")
                    messagebox.showinfo("Load Log", f"Berhasil memuat {len(self.data)} data dari Google Sheets.")
            except Exception as e:
                import traceback; traceback.print_exc()
                update_status(self.status_label, "Error loading data")
                messagebox.showerror("Load Log", f"Gagal memuat data: {e}")
            finally:
                loading.after(0, loading.destroy)
        loading = LoadingScreen(self)
        threading.Thread(target=lambda: do_load(loading)).start()

    def refresh_log_data(self, force_refresh=False):
        """Refresh hanya data log, tanpa me-refresh sheet lain."""
        self.load_sheet_data(force_refresh=force_refresh)

    def refresh_table(self):
        """
        Refresh the table display with enhanced styling.
        Features:
        - Alternating row colors for better visual separation
        - Proper data mapping to all 28 columns
        - Automatic column width adjustment
        """
        for row in self.tree.get_children():
            self.tree.delete(row)
        
        # Paging: hanya tampilkan data sesuai halaman
        start_idx = (self.current_page - 1) * self.page_size
        end_idx = start_idx + self.page_size
        page_data = self.filtered_data[start_idx:end_idx]
        for i, item in enumerate(page_data):
            values = []
            for col in ENHANCED_HEADER:
                val = item.get(col, "")
                values.append(str(val) if val is not None else "")
            
            # Apply alternating row colors for better readability
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            self.tree.insert("", "end", values=values, tags=(tag,))
        
        # Update column widths to ensure proper display
        self.tree.update_idletasks()
        self.update_paging_info()
    
    def get_column_display_name(self, col_name):
        """Get a shorter display name for column headers."""
        display_names = {
            "No. Agenda": "No. Agenda",
            "No. Surat": "No. Surat", 
            "Tgl. Surat": "Tgl. Surat",
            "Perihal": "Perihal",
            "Asal Surat": "Asal Surat",
            "Ditujukan": "Ditujukan",
            "Klasifikasi": "Klasifikasi",
            "Disposisi kepada": "Disposisi",
            "Untuk Di :": "Untuk Di",
            "Selesai Tgl.": "Selesai Tgl",
            "Kode Klasifikasi": "Kode",
            "Tgl. Penerimaan": "Tgl. Terima",
            "Indeks": "Indeks",
            "Bicarakan dengan": "Bicarakan",
            "Teruskan kepada": "Teruskan",
            "Harap Selesai Tanggal": "Harap Selesai",
            "Direktur Utama Instruksi": "Dir Utama Inst",
            "Direktur Utama Tanggal": "Dir Utama Tgl",
            "Direktur Keuangan Instruksi": "Dir Keu Inst",
            "Direktur Keuangan Tanggal": "Dir Keu Tgl",
            "Direktur Teknik Instruksi": "Dir Tek Inst",
            "Direktur Teknik Tanggal": "Dir Tek Tgl",
            "GM Keuangan & Administrasi Instruksi": "GM Keu Inst",
            "GM Keuangan & Administrasi Tanggal": "GM Keu Tgl",
            "GM Operasional & Pemeliharaan Instruksi": "GM Ops Inst",
            "GM Operasional & Pemeliharaan Tanggal": "GM Ops Tgl",
            "Manager Instruksi": "Manager Inst",
            "Manager Tanggal": "Manager Tgl"
        }
        return display_names.get(col_name, col_name)
    
    def get_column_info(self, col_name):
        """Get detailed information about a column."""
        column_info = {
            "No. Agenda": "Nomor urut agenda surat masuk",
            "No. Surat": "Nomor surat dari pengirim",
            "Tgl. Surat": "Tanggal surat dari pengirim",
            "Perihal": "Isi perihal surat",
            "Asal Surat": "Pengirim surat",
            "Ditujukan": "Tujuan surat",
            "Klasifikasi": "RAHASIA/PENTING/SEGERA",
            "Disposisi kepada": "Jabatan yang diberi disposisi",
            "Untuk Di :": "Tindakan yang diminta",
            "Selesai Tgl.": "Tanggal target selesai",
            "Kode Klasifikasi": "Kode klasifikasi surat",
            "Tgl. Penerimaan": "Tanggal surat diterima",
            "Indeks": "Indeks arsip",
            "Bicarakan dengan": "Orang yang perlu diajak bicara",
            "Teruskan kepada": "Orang yang perlu meneruskan",
            "Harap Selesai Tanggal": "Deadline penyelesaian"
        }
        return column_info.get(col_name, "Kolom instruksi dan tanggal untuk jabatan tertentu")

    def do_search(self):
        col = self.search_col.get()
        key = self.search_entry.get().lower()
        if not key:
            self.filtered_data = self.data.copy()
        else:
            # Pencarian hanya pada kolom label, bukan instruksi
            self.filtered_data = [d for d in self.data if key in str(d.get(col, "")).lower()]
        self.current_page = 1
        self.refresh_table()
        self.update_filter_values()
        self.update_record_count()
        update_status(self.status_label, f"Found {len(self.filtered_data)} matching records")

    def update_filter_values(self, event=None):
        col = self.filter_col.get()
        # Daftar pilihan tetap untuk kolom checkbox
        pilihan_checkbox = {
            "Disposisi kepada": [
                "Direktur Utama",
                "Direktur Keuangan",
                "Direktur Teknik",
                "GM Keuangan & Administrasi",
                "GM Operasional & Pemeliharaan",
                "Manager"
            ],
            "Klasifikasi": ["RAHASIA", "PENTING", "SEGERA"],
            "Untuk Di :": [
                "Ketahui & File",
                "Proses Selesai",
                "Teliti & Pendapat",
                "Buatkan Resume",
                "Edarkan",
                "Sesuai Disposisi",
                "Bicarakan dengan Saya"
            ]
        }
        if col in pilihan_checkbox:
            unique_vals = pilihan_checkbox[col]
        else:
            unique_vals = sorted({str(d.get(col, "")) for d in self.filtered_data})
        self.filter_val['values'] = unique_vals
        if unique_vals:
            self.filter_val.current(0)
        else:
            self.filter_val.set("")

    def do_filter(self):
        col = self.filter_col.get()
        val = self.filter_val.get()
        if not val:
            self.filtered_data = self.data.copy()
        else:
            # Ubah: Filter data yang MENGANDUNG nilai (bukan persis sama)
            self.filtered_data = [d for d in self.data if val in str(d.get(col, ""))]
        self.current_page = 1
        self.refresh_table()
        self.update_record_count()
        update_status(self.status_label, f"Filtered to {len(self.filtered_data)} records")

    def apply_year_month_filter(self, event=None):
        tahun = self.year_var.get()
        bulan = self.month_var.get()
        if not tahun and not bulan:
            self.filtered_data = self.data.copy()
        else:
            filtered = []
            for row in self.data:
                tgl = row.get("Tgl. Surat", "")
                if tgl:
                    tgl_split = tgl.replace('/', '-').replace('.', '-').split('-')
                    if len(tgl_split) == 3:
                        if len(tgl_split[2]) == 4:
                            t_tahun = tgl_split[2]
                            t_bulan = tgl_split[1]
                        else:
                            t_tahun = tgl_split[0]
                            t_bulan = tgl_split[1]
                        if (not tahun or t_tahun == tahun) and (not bulan or t_bulan == bulan):
                            filtered.append(row)
            self.filtered_data = filtered
        self.current_page = 1
        self.refresh_table()
        self.update_record_count()

    def on_double_click(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0], "values")
        if self.on_edit_log:
            data = {col: values[i] for i, col in enumerate(ENHANCED_HEADER)}
            self.on_edit_log(data)

    def edit_selected(self):
        selected = self.tree.selection()
        if not selected or len(selected) == 0:
            messagebox.showwarning("Edit Log", "Pilih data yang ingin diedit.")
            return
        if len(selected) > 1:
            messagebox.showwarning("Edit Log", "Pilih salah satu data saja untuk edit.")
            return
        values = self.tree.item(selected[0], "values")
        data = {col: values[i] for i, col in enumerate(ENHANCED_HEADER)}
        if self.on_edit_log:
            self.on_edit_log(data)

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Hapus Log", "Pilih data yang ingin dihapus.")
            return
        # Confirm deletion
        if not messagebox.askyesno("Confirm Delete", f"Yakin ingin menghapus {len(selected)} data terpilih?"):
            return
        # Ambil semua index baris yang dipilih (urutan dari bawah agar tidak bergeser saat hapus)
        idxs = sorted([self.tree.index(item) for item in selected], reverse=True)
        try:
            update_status(self.status_label, "Deleting record(s)...", self)
            service = get_sheets_service()
            sheet_id = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()['sheets'][0]['properties']['sheetId']
            requests = []
            for idx in idxs:
                sheet_row = idx + 6
                requests.append({
                    'deleteDimension': {
                        'range': {
                            'sheetId': sheet_id,
                            'dimension': 'ROWS',
                            'startIndex': sheet_row - 1,
                            'endIndex': sheet_row
                        }
                    }
                })
            service.spreadsheets().batchUpdate(
                spreadsheetId=SHEET_ID,
                body={'requests': requests}
            ).execute()
            update_status(self.status_label, "Record(s) deleted successfully", self)
            messagebox.showinfo("Hapus Log", f"{len(selected)} data berhasil dihapus dari Google Sheets.")
            self.refresh_log_data(force_refresh=True)  # Selalu refresh tanpa cache
        except Exception as e:
            import traceback; traceback.print_exc()
            update_status(self.status_label, "Error deleting record(s)", self)
            messagebox.showerror("Hapus Log", f"Gagal menghapus data: {e}")

    def export_excel(self):
        """Export filtered data to Excel with multi-layer header and filter info (centered)."""
        from tkinter import filedialog
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            title="Export Log ke Excel"
        )
        if not filepath or not isinstance(filepath, str):
            return
        from disposisi_app.views.components.export_utils import export_excel_advanced
        export_excel_advanced(self, filepath)

    def is_cache_expired(self):
        """Return True if the cache is expired, False otherwise."""
        return (time.time() - LogTab._last_cache_time) >= LogTab._cache_ttl 

    def update_paging_info(self):
        total_data = len(self.filtered_data)
        self.total_pages = max(1, (total_data + self.page_size - 1) // self.page_size)
        self.paging_info_label.config(text=f"Halaman {self.current_page} dari {self.total_pages}")
        self.btn_prev.config(state="normal" if self.current_page > 1 else "disabled")
        self.btn_next.config(state="normal" if self.current_page < self.total_pages else "disabled")
        # Tambahkan efek hover pada tombol paging
        for btn in [self.btn_prev, self.btn_next]:
            btn.bind("<Enter>", lambda e, b=btn: b.configure(cursor="hand2", style="Accent.TButton"))
            btn.bind("<Leave>", lambda e, b=btn: b.configure(cursor="", style="Secondary.TButton")) 

    def go_prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.refresh_table()

    def go_next_page(self):
        if self.current_page < self.total_pages:
            self.current_page += 1
            self.refresh_table()
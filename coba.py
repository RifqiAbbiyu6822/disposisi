import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Text
import logging
import traceback
import platform
import sys
import os
from google_sheets_connect import append_row_to_sheet, create_new_sheet, write_multilayer_header, append_rows_to_sheet, get_sheets_service, SHEET_ID
import smtplib
import tempfile
from email.message import EmailMessage
from sheet_logic import upload_to_sheet, get_untuk_di_labels, get_disposisi_labels
from logic.instruksi_table import InstruksiTable
import time
from tkcalendar import DateEntry
from main_app.edit_tab import EditTab
from pdf_output import save_form_to_pdf, merge_pdfs

try:
    from email_sender.config import EMAIL_HOST, EMAIL_PORT, EMAIL_HOST_USER, EMAIL_HOST_PASSWORD
except ImportError:
    EMAIL_HOST, EMAIL_PORT, EMAIL_HOST_USER, EMAIL_HOST_PASSWORD = (None, None, None, None)

# Import komponen UI modular
from disposisi_app.views.components.loading_screen import LoadingScreen
from disposisi_app.views.components.header import create_header
from disposisi_app.views.components.status_bar import create_status_bar
from disposisi_app.views.components.menu_bar import create_menu_bar
from disposisi_app.views.components.form_sections import (
    populate_frame_kiri, populate_frame_kanan, populate_frame_disposisi, populate_frame_instruksi
)
from disposisi_app.views.components.button_frame import create_button_frame
from disposisi_app.views.components.window_utils import center_window, setup_windowed_fullscreen
from disposisi_app.views.components.status_utils import update_status
from disposisi_app.views.components.validation import is_no_surat_unique
from disposisi_app.views.components.shortcuts import setup_shortcuts
from disposisi_app.views.components.gesture_handlers import setup_touchpad_gestures
from disposisi_app.views.components.dialogs import show_shortcuts, show_about
from disposisi_app.views.components.tooltip_utils import attach_tooltip, add_tooltips
from disposisi_app.views.components.export_utils import save_to_pdf, save_to_sheet
from disposisi_app.views.components.form_utils import clear_form
from disposisi_app.views.components.constants import POSISI_OPTIONS, TOOLTIP_LABELS
from disposisi_app.views.components.styles import setup_styles  # Impor style global

class FormApp(tk.Tk):
    """
    Main class aplikasi formulir disposisi dengan dukungan shortcut Windows dan gesture touchpad.
    """
    def __init__(self):
        super().__init__()
        
        # Enhanced window configuration
        self.title("📋 Aplikasi Persuratan Disposisi")
        self.geometry("1400x1000")
        self.minsize(1200, 800)
        
        # Initialize form variables
        self.form_vars = {}
        
        # Modern color scheme
        self.configure(bg="#f8fafc")
        
        # Setup styles AFTER tkinter initialization
        setup_styles(self)
        
        # Initialize variables
        self.logged_in = False
        self.sheet = None
        self.edit_mode = False
        self.pdf_attachments = []
        self.is_windows = platform.system() == "Windows"
        self.touchpad_gesture_active = False
        self.last_mouse_position = (0, 0)
        self.gesture_start_y = 0
        self.pinch_start_distance = 0
        self.log_data_dirty = False
        
        # Initialize all form variables
        self.init_variables()
        
        # Create header
        self.create_header()
        
        # Create menu bar
        self.create_menu_bar()
        
        # Create main content tabs
        self.create_tabs()
        
        # Create status bar
        self.create_status_bar()
        
        # Setup shortcuts and gestures
        self.setup_shortcuts()
        self.setup_touchpad_gestures()
        
        # Center window
        self.center_window()

    def center_window(self):
        center_window(self)

    def setup_windowed_fullscreen(self):
        setup_windowed_fullscreen(self)

    def create_header(self):
        """Create a compact, minimalist header"""
        # Header frame dengan height yang sangat minimal
        header_frame = tk.Frame(self, bg="#3b82f6", height=45)  # Reduced from 80 to 45
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)
        
        # Content dengan padding minimal
        header_content = tk.Frame(header_frame, bg="#3b82f6")
        header_content.pack(fill="both", expand=True, padx=12, pady=6)  # Minimal padding
        
        # Single row layout
        content_row = tk.Frame(header_content, bg="#3b82f6")
        content_row.pack(fill="x", expand=True)
        
        # Left side - compact title only
        left_frame = tk.Frame(content_row, bg="#3b82f6")
        left_frame.pack(side="left", anchor="w")
        
        # Single line title - no logo, no subtitle
        self.title_label = tk.Label(
            left_frame, 
            text="📋 Sistem Disposisi", 
            font=("Segoe UI", 13, "bold"),  # Smaller font
            bg="#3b82f6", 
            fg="white"
        )
        self.title_label.pack(anchor="w")
        
        # Right side - minimal info only
        right_frame = tk.Frame(content_row, bg="#3b82f6")
        right_frame.pack(side="right", anchor="e")
        
        # Compact info container
        info_container = tk.Frame(right_frame, bg="#3b82f6")
        info_container.pack(anchor="e")
        
        # Version only
        version_label = tk.Label(
            info_container, 
            text="v2.0", 
            font=("Segoe UI", 9, "bold"), 
            bg="#3b82f6", 
            fg="#bfdbfe"
        )
        version_label.pack(side="right", padx=(0, 0))
        
        # Simple status dot
        status_dot = tk.Label(
            info_container, 
            text="●", 
            font=("Segoe UI", 8), 
            bg="#3b82f6", 
            fg="#10b981"
        )
        status_dot.pack(side="right", padx=(0, 6))

    def create_menu_bar(self):
        """Create a compact menu bar"""
        menubar = tk.Menu(self, font=("Segoe UI", 9), bg="#ffffff", fg="#1f2937", 
                          activebackground="#3b82f6", activeforeground="white")
        self.config(menu=menubar)
        
        # File menu - lebih sedikit item
        file_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 9), 
                            bg="#ffffff", fg="#1f2937",
                            activebackground="#3b82f6", activeforeground="white")
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="💾 Simpan PDF", command=self.save_to_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="🚪 Keluar", command=self.quit)
        
        # Edit menu - minimal
        edit_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 9),
                            bg="#ffffff", fg="#1f2937",
                            activebackground="#3b82f6", activeforeground="white")
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="🗑️ Reset", command=self.clear_form)
        
        # Help menu - minimal
        help_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 9),
                            bg="#ffffff", fg="#1f2937",
                            activebackground="#3b82f6", activeforeground="white")
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="⌨️ Shortcuts", command=self.show_shortcuts)

    def create_status_bar(self):
        """Create a compact status bar"""
        status_frame = tk.Frame(self, bg="#f8fafc", height=28, relief="solid", borderwidth=1)  # Reduced height
        status_frame.pack(side="bottom", fill="x")
        status_frame.pack_propagate(False)
        
        left_status = tk.Frame(status_frame, bg="#f8fafc")
        left_status.pack(side="left", fill="both", expand=True)
        
        self.status_message = tk.Label(left_status, 
                                       text="✓ Ready", 
                                       font=("Segoe UI", 9),  # Smaller font
                                       bg="#f8fafc", 
                                       fg="#10b981")
        self.status_message.pack(side="left", padx=15, pady=4)  # Reduced padding
        
        right_status = tk.Frame(status_frame, bg="#f8fafc")
        right_status.pack(side="right")
        
        # Compact version only
        version_label = tk.Label(right_status, 
                                 text="v2.0", 
                                 font=("Segoe UI", 9, "bold"), 
                                 bg="#f8fafc", 
                                 fg="#3b82f6")
        version_label.pack(side="right", padx=12, pady=4)

    def setup_shortcuts(self):
        setup_shortcuts(
            self,
            self.save_to_pdf,
            self.clear_form,
            self.select_next_tab,
            self.select_prev_tab
        )

    def select_next_tab(self):
        if hasattr(self, 'notebook'):
            current = self.notebook.index(self.notebook.select())
            total = self.notebook.index('end')
            self.notebook.select((current + 1) % total)

    def select_prev_tab(self):
        if hasattr(self, 'notebook'):
            current = self.notebook.index(self.notebook.select())
            total = self.notebook.index('end')
            self.notebook.select((current - 1) % total)

    def setup_touchpad_gestures(self):
        if self.is_windows:
            setup_touchpad_gestures(
                self,
                self.on_mouse_down,
                self.on_mouse_up,
                self.on_mouse_drag,
                self.on_double_click,
                self.on_mousewheel,
                self.on_horizontal_scroll,
                self.on_pinch_gesture_start,
                self.on_pinch_gesture
            )

    def on_mouse_down(self, event):
        self.last_mouse_position = (event.x, event.y)
        self.gesture_start_y = event.y
        self.touchpad_gesture_active = True

    def on_mouse_up(self, event):
        self.touchpad_gesture_active = False

    def on_mouse_drag(self, event):
        if not self.touchpad_gesture_active:
            return
            
        delta_y = event.y - self.gesture_start_y
        
        if abs(delta_y) > 50:
            if delta_y > 0:
                self.scroll_down()
            else:
                self.scroll_up()
            self.gesture_start_y = event.y

    def on_double_click(self, event):
        pass

    def on_mousewheel(self, event):
        pass

    def on_horizontal_scroll(self, event):
        pass

    def on_pinch_gesture_start(self, event):
        self.pinch_start_distance = 0

    def on_pinch_gesture(self, event):
        pass

    def scroll_up(self):
        pass

    def scroll_down(self):
        pass

    def refresh_data(self):
        pass

    def toggle_fullscreen(self):
        self.attributes('-fullscreen', not self.attributes('-fullscreen'))
        self.update_status("Fullscreen toggled")

    def export_excel(self):
        messagebox.showwarning("Export", "No data available to export")

    def show_shortcuts(self):
        show_shortcuts(self)

    def show_about(self):
        show_about(self)

    def update_status(self, message):
        if hasattr(self, 'status_message'):
            update_status(self.status_message, message, self)

    def create_tabs(self):
        notebook_frame = ttk.Frame(self, style="TFrame", padding=(10, 5, 10, 5))
        notebook_frame.pack(fill="both", expand=True, padx=0, pady=0)
        
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        notebook_frame.rowconfigure(0, weight=1)
        notebook_frame.columnconfigure(0, weight=1)
        
        self.notebook = ttk.Notebook(notebook_frame, style="TNotebook")
        self.notebook.grid(row=0, column=0, sticky="nsew")
        
        self.form_frame = ttk.Frame(self.notebook, style="Surface.TFrame")
        self.form_frame.rowconfigure(0, weight=1)
        self.form_frame.columnconfigure(0, weight=1)
        self.notebook.add(self.form_frame, text="📝 Formulir Disposisi")
        
        self.form_input_widgets = self.create_widgets(self.form_frame)
        
        from excel_crud_tab import LogTab
        self.log_frame = LogTab(self.notebook, on_edit_log=self.open_edit_tab)
        self.notebook.add(self.log_frame, text="📊 Data & Log")
        
        def on_tab_changed(event):
            tab_id = event.widget.index("current")
            if tab_id == 1:
                if self.log_data_dirty or self.log_frame.is_cache_expired():
                    self.log_frame.refresh_log_data(force_refresh=True)
                    self.log_data_dirty = False
        
        self.notebook.bind("<<NotebookTabChanged>>", on_tab_changed)

    def open_edit_tab(self, data_log):
        for i in range(self.notebook.index('end')):
            if self.notebook.tab(i, 'text') == '✏️ Edit Data':
                self.notebook.forget(i)
                break
        edit_tab = EditTab(self.notebook, data_log, on_save_callback=self._on_edit_saved)
        self.notebook.add(edit_tab, text="✏️ Edit Data")
        self.notebook.select(edit_tab)
        self.edit_tab = edit_tab

    def _on_edit_saved(self):
        if hasattr(self, 'notebook') and hasattr(self, 'log_frame'):
            self.notebook.select(self.log_frame)
            self.log_frame.refresh_log_data(force_refresh=True)
        for i in range(self.notebook.index('end')):
            if self.notebook.tab(i, 'text') == '✏️ Edit Data':
                self.notebook.forget(i)
                break

    def init_variables(self):
        self.vars = {
            "no_agenda": tk.StringVar(), "no_surat": tk.StringVar(),
            "perihal": tk.StringVar(),
            "asal_surat": tk.StringVar(),
            "ditujukan": tk.StringVar(),
            "rahasia": tk.IntVar(), "penting": tk.IntVar(), "segera": tk.IntVar(),
            "kode_klasifikasi": tk.StringVar(), "indeks": tk.StringVar(),
            "dir_utama": tk.IntVar(), "dir_keu": tk.IntVar(), "dir_teknik": tk.IntVar(),
            "gm_keu": tk.IntVar(), "gm_ops": tk.IntVar(), 
            "manager_pemeliharaan": tk.IntVar(), "manager_operasional": tk.IntVar(),
            "manager_administrasi": tk.IntVar(), "manager_keuangan": tk.IntVar(),
            "ketahui_file": tk.IntVar(), "proses_selesai": tk.IntVar(),
            "teliti_pendapat": tk.IntVar(), "buatkan_resume": tk.IntVar(),
            "edarkan": tk.IntVar(), "sesuai_disposisi": tk.IntVar(),
            "bicarakan_saya": tk.IntVar(), "bicarakan_dengan": tk.StringVar(),
            "teruskan_kepada": tk.StringVar(),
        }
        self.instruction_vars = {}
        self.posisi_labels = [f"Posisi{i+1}" for i in range(5)]
        for row in range(5):
            self.instruction_vars[f"instruksi_{row}"] = tk.StringVar()
            self.instruction_vars[f"posisi_{row}"] = tk.StringVar()
            self.instruction_vars[f"tanggal_{row}"] = tk.StringVar()

    def create_widgets(self, parent):
        input_widgets = {}
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        
        canvas = tk.Canvas(parent, borderwidth=0, highlightthickness=0, bg="#f8fafc")
        
        v_scroll = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(parent, orient="horizontal", command=canvas.xview)
        
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        
        main_frame = ttk.Frame(canvas, padding=(20, 20, 20, 20), style="Surface.TFrame")
        self._form_main_frame = main_frame
        
        canvas.create_window((0, 0), window=main_frame, anchor="nw")
        
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        main_frame.bind("<Configure>", on_frame_configure)
        
        def _on_mousewheel(event):
            if event.num == 5 or event.delta == -120:
                canvas.yview_scroll(1, "units")
            elif event.num == 4 or event.delta == 120:
                canvas.yview_scroll(-1, "units")
            else:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _on_shift_mousewheel(event):
            canvas.xview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Shift-MouseWheel>", _on_shift_mousewheel)
        canvas.bind_all("<Button-4>", _on_mousewheel)
        canvas.bind_all("<Button-5>", _on_mousewheel)
        
        input_widgets.update(self.create_top_frame(main_frame))
        input_widgets.update(self.create_middle_frame(main_frame))
        
        attachment_frame = ttk.LabelFrame(main_frame, text="📎 Lampiran PDF", padding=(20, 15, 20, 20), style="TLabelframe")
        attachment_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 20))
        
        listbox_frame = ttk.Frame(attachment_frame)
        listbox_frame.pack(fill="both", expand=True)
        
        self.attachment_listbox = tk.Listbox(listbox_frame, height=4, selectmode=tk.SINGLE,
                                              font=("Segoe UI", 9),
                                              bg="#ffffff", fg="#1f2937",
                                              selectbackground="#3b82f6",
                                              relief="solid", borderwidth=1)
        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=self.attachment_listbox.yview)
        self.attachment_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.attachment_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        attachment_buttons = ttk.Frame(attachment_frame)
        attachment_buttons.pack(fill="x", pady=(10, 0))
        
        add_btn = ttk.Button(attachment_buttons, text="➕ Tambah PDF", 
                             command=self.add_pdf_attachment, style="Secondary.TButton")
        add_btn.pack(side="left", padx=(0, 5))
        
        remove_btn = ttk.Button(attachment_buttons, text="🗑️ Hapus PDF", 
                                  command=self.remove_pdf_attachment, style="Secondary.TButton")
        remove_btn.pack(side="left")
        
        self.create_button_frame(main_frame)
        self.add_tooltips(input_widgets)
        
        return input_widgets

    def add_pdf_attachment(self):
        """Add PDF attachment to the form"""
        from tkinter import filedialog
        import os
        
        files = filedialog.askopenfilenames(
            title="Pilih PDF lampiran",
            filetypes=[("PDF Documents", "*.pdf"), ("All Files", "*.*")]
        )
        
        for file_path in files:
            if file_path not in self.pdf_attachments:
                self.pdf_attachments.append(file_path)
                filename = os.path.basename(file_path)
                display_name = f"📄 {filename}"
                self.attachment_listbox.insert(tk.END, display_name)
                print(f"[DEBUG] Added PDF attachment: {filename}")
        
        if files:
            self.update_status(f"✓ {len(files)} PDF ditambahkan")

    def remove_pdf_attachment(self):
        """Remove selected PDF attachment from the form"""
        import os
        
        selected = self.attachment_listbox.curselection()
        if selected:
            idx = selected[0]
            if idx < len(self.pdf_attachments):
                filename = os.path.basename(self.pdf_attachments[idx])
                self.attachment_listbox.delete(idx)
                del self.pdf_attachments[idx]
                self.update_status(f"🗑️ {filename} dihapus")
                print(f"[DEBUG] Removed PDF attachment: {filename}")
            else:
                print(f"[WARNING] Index {idx} out of range for pdf_attachments")
        else:
            messagebox.showwarning("Hapus Lampiran", "Pilih file yang akan dihapus terlebih dahulu.")

    def refresh_pdf_attachments(self, parent):
        """Refresh PDF attachments display"""
        # Clear current listbox
        if hasattr(self, 'attachment_listbox'):
            self.attachment_listbox.delete(0, tk.END)
            
            # Repopulate with current attachments
            for pdf_path in self.pdf_attachments:
                filename = os.path.basename(pdf_path)
                display_name = f"📄 {filename}"
                self.attachment_listbox.insert(tk.END, display_name)

    def add_tooltips(self, input_widgets):
        add_tooltips(input_widgets, TOOLTIP_LABELS)

    def attach_tooltip(self, widget, text):
        attach_tooltip(widget, text)

    def create_top_frame(self, parent):
        input_widgets = {}
        
        top_frame = ttk.Frame(parent, style="TFrame")
        top_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 20))
        top_frame.columnconfigure(0, weight=1)
        top_frame.columnconfigure(1, weight=1)
        top_frame.rowconfigure(0, weight=1)
        
        frame_kiri = ttk.LabelFrame(top_frame, text="📄 Detail Surat", 
                                    padding=(20, 15, 20, 20), style="TLabelframe")
        frame_kiri.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        frame_kiri.columnconfigure(1, weight=1)
        
        input_widgets.update(populate_frame_kiri(frame_kiri, self.vars))
        
        frame_kanan = ttk.LabelFrame(top_frame, text="🏷️ Klasifikasi & Status", 
                                     padding=(20, 15, 20, 20), style="TLabelframe")
        frame_kanan.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        frame_kanan.columnconfigure(0, weight=1)
        
        self.tgl_terima_entry = populate_frame_kanan(frame_kanan, self.vars)
        
        return input_widgets

    def create_middle_frame(self, parent):
        input_widgets = {}
        
        middle_frame = ttk.Frame(parent, style="TFrame")
        middle_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 20))
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)
        
        # Fixed: More balanced grid weights
        middle_frame.columnconfigure(0, weight=1)  # Disposisi Kepada
        middle_frame.columnconfigure(1, weight=1)  # Untuk di
        middle_frame.columnconfigure(2, weight=1)  # Isi Instruksi - Fixed: Changed from 2 to 1
        middle_frame.rowconfigure(0, weight=1)
        
        frame_disposisi = ttk.LabelFrame(middle_frame, text="👥 Disposisi Kepada", 
                                         padding=(20, 15, 20, 20), style="TLabelframe")
        frame_disposisi.grid(row=0, column=0, sticky="nsew", padx=(0, 10))  # Fixed: Increased padding
        frame_disposisi.columnconfigure(0, weight=1)  # Fixed: Added column weight
        
        populate_frame_disposisi(frame_disposisi, self.vars)
        
        frame_instruksi = ttk.LabelFrame(middle_frame, text="📋 Untuk di", 
                                         padding=(20, 15, 20, 20), style="TLabelframe")
        frame_instruksi.grid(row=0, column=1, sticky="nsew", padx=(10, 10))  # Fixed: Increased padding
        frame_instruksi.columnconfigure(0, weight=1)  # Fixed: Added column weight
        frame_instruksi.columnconfigure(1, weight=1)  # Fixed: Added column weight
        
        input_widgets.update(populate_frame_instruksi(frame_instruksi, self.vars))
        
        self.harap_selesai_tgl_entry = input_widgets.get("harap_selesai_tgl_entry")
        
        frame_info = ttk.LabelFrame(middle_frame, text="📝 Isi Instruksi / Informasi", 
                                    padding=(20, 15, 20, 20), style="TLabelframe")
        frame_info.grid(row=0, column=2, sticky="nsew", padx=(10, 0))  # Fixed: Increased padding
        frame_info.columnconfigure(0, weight=1)  # Fixed: Added column weight
        frame_info.rowconfigure(0, weight=1)  # Fixed: Added row weight
        
        frame_instruksi_table = ttk.Frame(frame_info, style="Surface.TFrame")
        frame_instruksi_table.grid(row=0, column=0, sticky="nsew")
        frame_instruksi_table.columnconfigure(0, weight=1)  # Fixed: Added column weight
        frame_instruksi_table.rowconfigure(0, weight=1)  # Fixed: Added row weight
        
        self.posisi_options = POSISI_OPTIONS
        self.instruksi_table = InstruksiTable(frame_instruksi_table, self.posisi_options, use_grid=True)
        
        btn_frame = ttk.Frame(frame_info, style="TFrame")
        btn_frame.grid(row=1, column=0, pady=(15, 0), sticky="ew")
        
        self.tambah_baris_btn = ttk.Button(btn_frame, text="➕ Tambah Baris", 
                                           command=self.instruksi_table.add_row, 
                                           style="Secondary.TButton")
        self.tambah_baris_btn.pack(side="left", padx=(0, 8))
        
        self.hapus_baris_btn = ttk.Button(btn_frame, text="➖ Hapus Baris", 
                                         command=self.instruksi_table.remove_selected_rows, 
                                         style="Danger.TButton")
        self.hapus_baris_btn.pack(side="left", padx=(0, 8))
        
        self.kosongkan_baris_btn = ttk.Button(btn_frame, text="🗑️ Kosongkan", 
                                              command=self.instruksi_table.kosongkan_semua_baris, 
                                              style="Secondary.TButton")
        self.kosongkan_baris_btn.pack(side="left")
        
        return input_widgets

    def create_button_frame(self, parent):
        from disposisi_app.views.components.button_frame import create_button_frame

        callbacks = {
            "save_pdf": self.save_to_pdf,
            "save_sheet": self.save_to_sheet,
            "send_email": self.send_email_with_disposisi,
            "get_disposisi_labels": self.get_disposisi_labels,
            "clear_form": self.clear_form
        }
        self._button_frame = create_button_frame(
            parent,
            callbacks
        )
        return self._button_frame

    def get_disposisi_labels(self):
        """Get selected disposisi labels"""
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
            if var in self.vars and self.vars[var].get():
                labels.append(label)
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
            if var in self.vars and self.vars[var].get():
                # Pisahkan manager dari label lainnya
                if label.startswith("Manager"):
                    manager_labels.append(label)
                else:
                    labels.append(label)
        
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

    def send_email_with_disposisi(self, selected_positions):
        """
        FIXED VERSION: Mengirimkan email disposisi ke posisi yang dipilih 
        dengan mengambil alamat email yang BENAR dari spreadsheet admin.
        """
        import os
        import tempfile
        import traceback
        from tkinter import messagebox
        from datetime import datetime
        
        # Import modul email yang sudah diperbaiki
        try:
            from email_sender.send_email import EmailSender
            from email_sender.template_handler import render_email_template
            from pdf_output import save_form_to_pdf, merge_pdfs
        except ImportError as e:
            messagebox.showerror("Import Error", f"Gagal import modul email: {e}")
            return

        print(f"[DEBUG] Selected positions: {selected_positions}")
        self.update_status("Menyiapkan pengiriman email...")
        
        # Collect form data
        data = collect_form_data_safely(self)
        if not data.get("no_surat", "").strip():
            messagebox.showerror("Validation Error", "No. Surat tidak boleh kosong untuk mengirim email.")
            return

        # Generate PDF attachment
        temp_pdf_path = None
        final_pdf_path = None
        
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf_path = temp_pdf.name
            
            save_form_to_pdf(temp_pdf_path, data)
            
            final_pdf_path = temp_pdf_path
            if hasattr(self, 'pdf_attachments') and self.pdf_attachments:
                merged_pdf_path = os.path.join(tempfile.gettempdir(), f"merged_{os.path.basename(temp_pdf_path)}")
                pdf_list = [temp_pdf_path] + self.pdf_attachments
                merge_pdfs(pdf_list, merged_pdf_path)
                final_pdf_path = merged_pdf_path

        except Exception as e:
            messagebox.showerror("PDF Generation Error", f"Gagal membuat PDF untuk email: {e}")
            traceback.print_exc()
            return
        
        # Initialize EmailSender
        try:
            email_sender = EmailSender()
            
            if not email_sender.sheets_service:
                error_msg = ("Google Sheets tidak dapat diakses.\n\n"
                            "Pastikan:\n"
                            "1. File admin/credentials.json tersedia\n"
                            "2. Sheet admin dengan email sudah dibuat\n"
                            "3. ID sheet admin benar di config.py")
                messagebox.showerror("Email Configuration Error", error_msg)
                return
                
        except Exception as e:
            error_msg = f"Gagal menginisialisasi email sender: {str(e)}"
            messagebox.showerror("Email Sender Error", error_msg)
            traceback.print_exc()
            return
        
        # PERBAIKAN UTAMA: Ambil email yang benar dari spreadsheet
        self.update_status("Mencari alamat email dari database admin...")
        
        recipient_emails = []
        failed_lookups = []
        successful_lookups = []
        
        print(f"[DEBUG] Looking up emails for positions: {selected_positions}")
        
        for position in selected_positions:
            print(f"[DEBUG] Looking up email for: {position}")
            email, msg = email_sender.get_recipient_email(position)
            print(f"[DEBUG] Result for {position}: email={email}, msg={msg}")
            
            if email and email != position:
                # Validasi format email
                if '@' in email and '.' in email.split('@')[-1]:
                    recipient_emails.append(email)
                    successful_lookups.append(f"{position}: {email}")
                    print(f"[DEBUG] ✓ Valid email found: {position} -> {email}")
                else:
                    failed_lookups.append(f"{position} (invalid email format: {email})")
                    print(f"[DEBUG] ✗ Invalid email format: {position} -> {email}")
            else:
                failed_lookups.append(f"{position} ({msg})")
                print(f"[DEBUG] ✗ No email found: {position} -> {msg}")
        
        print(f"[DEBUG] Final recipient emails: {recipient_emails}")
        print(f"[DEBUG] Failed lookups: {failed_lookups}")
        
        # Show warnings for failed lookups
        if failed_lookups:
            # Gunakan singkatan untuk manager dalam warning message
            abbreviation_map = {
                "Manager Pemeliharaan": "pml",
                "Manager Operasional": "ops",
                "Manager Administrasi": "adm",
                "Manager Keuangan": "keu"
            }
            
            # Konversi failed lookups ke singkatan untuk display
            display_failed_lookups = []
            for lookup in failed_lookups:
                for full_name, abbrev in abbreviation_map.items():
                    if full_name in lookup:
                        lookup = lookup.replace(full_name, f"Manager {abbrev}")
                        break
                display_failed_lookups.append(lookup)
            
            warning_msg = ("Tidak dapat menemukan alamat email yang valid untuk posisi berikut:\n" + 
                          "\n".join([f"• {lookup}" for lookup in display_failed_lookups]))
            
            if len(failed_lookups) == len(selected_positions):
                # All positions failed
                full_error = (warning_msg + 
                             "\n\nPastikan admin sudah mengisi email yang valid untuk posisi-posisi tersebut "
                             "di spreadsheet admin.\n\n"
                             "Format email harus: user@domain.com")
                messagebox.showerror("No Valid Emails Found", full_error)
                return
            else:
                # Some positions failed - show warning but continue
                result = messagebox.askyesno(
                    "Some Emails Missing", 
                    warning_msg + "\n\nLanjutkan mengirim ke email yang valid?", 
                    parent=self
                )
                if not result:
                    return
        
        if not recipient_emails:
            messagebox.showerror("No Recipients", "Tidak ada alamat email yang valid untuk dikirimi.")
            return
        
        # Prepare email template data
        template_data = {
            'nomor_surat': data.get('no_surat', 'N/A'),
            'nama_pengirim': data.get('asal_surat', 'N/A'),
            'perihal': data.get('perihal', 'N/A'),
            'tanggal': datetime.now().strftime('%d %B %Y'),
            'klasifikasi': [],
            'instruksi_list': [],
            'tahun': datetime.now().year
        }
        
        # Tambahkan informasi disposisi kepada dengan format abbreviation
        try:
            disposisi_labels = self.get_disposisi_labels_with_abbreviation()
            template_data['disposisi_kepada'] = disposisi_labels
        except Exception as e:
            print(f"[WARNING] Error getting disposisi labels: {e}")
            template_data['disposisi_kepada'] = []
        
        # Add classification
        if data.get('rahasia', 0):
            template_data['klasifikasi'].append("RAHASIA")
        if data.get('penting', 0):
            template_data['klasifikasi'].append("PENTING")
        if data.get('segera', 0):
            template_data['klasifikasi'].append("SEGERA")
        
        # Add instructions from checkboxes
        instruksi_mapping = [
            ("ketahui_file", "Ketahui & File"),
            ("proses_selesai", "Proses Selesai"),
            ("teliti_pendapat", "Teliti & Pendapat"),
            ("buatkan_resume", "Buatkan Resume"),
            ("edarkan", "Edarkan"),
            ("sesuai_disposisi", "Sesuai Disposisi"),
            ("bicarakan_saya", "Bicarakan dengan Saya")
        ]
        
        for key, label in instruksi_mapping:
            if data.get(key, 0):
                template_data['instruksi_list'].append(label)
        
        # Add detailed instructions from table
        if 'isi_instruksi' in data and data['isi_instruksi']:
            for instr in data['isi_instruksi']:
                if instr.get('instruksi', '').strip():
                    posisi = instr.get('posisi', '')
                    instruksi_text = instr.get('instruksi', '')
                    tanggal = instr.get('tanggal', '')
                    
                    # Gunakan singkatan untuk manager dalam instruksi
                    abbreviation_map = {
                        "Manager Pemeliharaan": "pml",
                        "Manager Operasional": "ops",
                        "Manager Administrasi": "adm",
                        "Manager Keuangan": "keu"
                    }
                    
                    # Konversi posisi ke singkatan jika ada
                    display_posisi = abbreviation_map.get(posisi, posisi)
                    if display_posisi in ["pml", "ops", "adm", "keu"]:
                        display_posisi = f"Manager {display_posisi}"
                    
                    instr_line = f"{display_posisi}: {instruksi_text}"
                    if tanggal:
                        instr_line += f" (Tanggal: {tanggal})"
                    template_data['instruksi_list'].append(instr_line)
        
        # Add additional instructions
        if data.get("bicarakan_dengan", "").strip():
            template_data['instruksi_list'].append(f"Bicarakan dengan: {data['bicarakan_dengan']}")
        
        if data.get("teruskan_kepada", "").strip():
            template_data['instruksi_list'].append(f"Teruskan kepada: {data['teruskan_kepada']}")
        
        if data.get("harap_selesai_tgl", "").strip():
            template_data['instruksi_list'].append(f"Harap diselesaikan tanggal: {data['harap_selesai_tgl']}")
        
        # Render email template and send
        try:
            html_content = render_email_template(template_data)
            subject = f"Disposisi Surat: {data.get('perihal', 'N/A')}"
            
            self.update_status(f"Mengirim email ke {len(recipient_emails)} penerima...")
            print(f"[DEBUG] Sending email to: {recipient_emails}")
            print(f"[DEBUG] Subject: {subject}")
            print(f"[DEBUG] PDF attachment: {final_pdf_path}")
            
            # PERBAIKAN: Gunakan method yang benar dengan PDF attachment
            success, message = email_sender.send_disposisi_email(
                recipient_emails, 
                subject, 
                html_content,
                pdf_attachment=final_pdf_path
            )
            
            print(f"[DEBUG] Email send result: success={success}, message={message}")
            
            # Show results
            if success:
                self.update_status("Email berhasil dikirim!")
                success_msg = f"Email berhasil dikirim ke:\n{chr(10).join([f'• {email}' for email in recipient_emails])}"
                
                if successful_lookups:
                    # Gunakan singkatan untuk manager dalam successful lookups
                    abbreviation_map = {
                        "Manager Pemeliharaan": "pml",
                        "Manager Operasional": "ops",
                        "Manager Administrasi": "adm",
                        "Manager Keuangan": "keu"
                    }
                    
                    # Konversi successful lookups ke singkatan untuk display
                    display_successful_lookups = []
                    for lookup in successful_lookups:
                        for full_name, abbrev in abbreviation_map.items():
                            if full_name in lookup:
                                lookup = lookup.replace(full_name, f"Manager {abbrev}")
                                break
                        display_successful_lookups.append(lookup)
                    
                    success_msg += f"\n\nDetail penerima:\n{chr(10).join([f'• {lookup}' for lookup in display_successful_lookups])}"
                
                messagebox.showinfo("Email Sent Successfully", success_msg, parent=self)
            else:
                self.update_status("Gagal mengirim email.")
                messagebox.showerror("Email Send Failed", f"Gagal mengirim email:\n{message}", parent=self)
                
        except Exception as e:
            self.update_status("Error saat mengirim email.")
            messagebox.showerror("Email Error", f"Terjadi kesalahan saat mengirim email: {e}")
            traceback.print_exc()
        
        finally:
            # Clean up temporary files
            try:
                if temp_pdf_path and os.path.exists(temp_pdf_path):
                    os.remove(temp_pdf_path)
                    print(f"[DEBUG] Cleaned up temp PDF: {temp_pdf_path}")
                if final_pdf_path and final_pdf_path != temp_pdf_path and os.path.exists(final_pdf_path):
                    os.remove(final_pdf_path)
                    print(f"[DEBUG] Cleaned up final PDF: {final_pdf_path}")
            except Exception as e:
                print(f"[WARNING] Could not remove temporary files: {e}")

    def save_to_pdf(self):
        save_to_pdf(self)

    def save_to_sheet(self):
        save_to_sheet(self)

    def clear_form(self):
        clear_form(self)

    def _global_on_mousewheel(self, event):
        pass

    def _global_on_shift_mousewheel(self, event):
        pass


# TAMBAHAN: Fungsi untuk debug dan test sistem email
def test_admin_sheet_connection():
    """
    Test function untuk memverifikasi koneksi ke admin sheet
    Jalankan ini untuk debug masalah email
    """
    try:
        from email_sender.send_email import EmailSender
        
        print("=== TESTING ADMIN SHEET CONNECTION ===")
        
        email_sender = EmailSender()
        
        if not email_sender.sheets_service:
            print("❌ ERROR: Google Sheets service not initialized")
            print("Check:")
            print("1. admin/credentials.json file exists")
            print("2. Credentials have proper permissions")
            print("3. Internet connection available")
            return False
        
        print("✅ Google Sheets service initialized")
        print(f"Admin Sheet ID: {email_sender.admin_sheet_id}")
        
        # Test reading each position
        print("\n=== TESTING EMAIL LOOKUPS ===")
        all_positions = [
            "Direktur Utama",
            "Direktur Keuangan", 
            "Direktur Teknik",
            "GM Keuangan & Administrasi",
            "GM Operasional & Pemeliharaan",
            "Manager Pemeliharaan",
            "Manager Operasional",
            "Manager Administrasi",
            "Manager Keuangan"
        ]
        
        valid_emails = 0
        for position in all_positions:
            email, msg = email_sender.get_recipient_email(position)
            if email and '@' in email:
                print(f"✅ {position}: {email}")
                valid_emails += 1
            else:
                print(f"❌ {position}: {msg}")
        
        print(f"\n=== SUMMARY ===")
        print(f"Valid emails found: {valid_emails}/{len(all_positions)}")
        
        if valid_emails > 0:
            print("✅ System ready for sending emails")
            return True
        else:
            print("❌ No valid emails found - check admin spreadsheet")
            return False
            
    except Exception as e:
        print(f"❌ ERROR during test: {e}")
        import traceback
        traceback.print_exc()
        return False


def collect_form_data_safely(self):
    """Safely collect all form data with proper error handling"""
    try:
        # Start with vars data (IntVar, StringVar, etc.)
        data = {}
        if hasattr(self, 'vars') and self.vars:
            for key, var in self.vars.items():
                try:
                    data[key] = var.get()
                except Exception as e:
                    print(f"[WARNING] Error getting value from var {key}: {e}")
                    data[key] = 0 if 'IntVar' in str(type(var)) else ""
        
        # Get data from form input widgets
        if hasattr(self, 'form_input_widgets') and self.form_input_widgets:
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
                                data[field] = str(widget.get())
                        else:
                            data[field] = ""
                    except Exception as e:
                        print(f"[WARNING] Error getting value from {field}: {e}")
                        data[field] = ""
            
            # Date widget for tgl_surat
            if "tgl_surat" in self.form_input_widgets:
                try:
                    widget = self.form_input_widgets["tgl_surat"]
                    if hasattr(widget, 'get_date'):
                        date_obj = widget.get_date()
                        data["tgl_surat"] = date_obj.strftime("%d-%m-%Y")
                    elif hasattr(widget, 'get'):
                        data["tgl_surat"] = str(widget.get())
                    else:
                        data["tgl_surat"] = ""
                except Exception as e:
                    print(f"[WARNING] Error getting tgl_surat: {e}")
                    data["tgl_surat"] = ""
        
        # Handle special date entries
        if hasattr(self, 'tgl_terima_entry'):
            try:
                if hasattr(self.tgl_terima_entry, 'get_date'):
                    date_obj = self.tgl_terima_entry.get_date()
                    data["tgl_terima"] = date_obj.strftime("%d-%m-%Y")
                elif hasattr(self.tgl_terima_entry, 'get'):
                    data["tgl_terima"] = str(self.tgl_terima_entry.get())
                else:
                    data["tgl_terima"] = ""
            except Exception as e:
                print(f"[WARNING] Error getting tgl_terima: {e}")
                data["tgl_terima"] = ""
        else:
            data["tgl_terima"] = ""
        
        if hasattr(self, 'harap_selesai_tgl_entry'):
            try:
                if hasattr(self.harap_selesai_tgl_entry, 'get_date'):
                    date_obj = self.harap_selesai_tgl_entry.get_date()
                    data["harap_selesai_tgl"] = date_obj.strftime("%d-%m-%Y")
                elif hasattr(self.harap_selesai_tgl_entry, 'get'):
                    data["harap_selesai_tgl"] = str(self.harap_selesai_tgl_entry.get())
                else:
                    data["harap_selesai_tgl"] = ""
            except Exception as e:
                print(f"[WARNING] Error getting harap_selesai_tgl: {e}")
                data["harap_selesai_tgl"] = ""
        else:
            data["harap_selesai_tgl"] = ""
        
        # Ensure required fields have default values
        required_defaults = {
            "indeks": "",
            "rahasia": 0,
            "penting": 0,
            "segera": 0,
            "kode_klasifikasi": "",
            "no_agenda": "",
            "no_surat": "",
            "perihal": "",
            "asal_surat": "",
            "ditujukan": ""
        }
        
        for field, default_value in required_defaults.items():
            if field not in data or data[field] is None:
                data[field] = default_value
        
        # Collect instructions safely
        if hasattr(self, "instruksi_table") and hasattr(self.instruksi_table, "get_data"):
            try:
                data["isi_instruksi"] = self.instruksi_table.get_data()
            except Exception as e:
                print(f"[WARNING] Error getting instruction data: {e}")
                data["isi_instruksi"] = []
        else:
            data["isi_instruksi"] = []
        
        return data
        
    except Exception as e:
        print(f"[ERROR] Error collecting form data: {e}")
        import traceback
        traceback.print_exc()
        return {}


# Main application entry point
def main():
    """Main function to run the Email Manager Application"""
    try:
        app = FormApp()
        app.mainloop()
    except Exception as e:
        print(f"Error starting application: {str(e)}")
        messagebox.showerror("Error", f"Terjadi kesalahan saat memulai aplikasi:\n{str(e)}")


if __name__ == "__main__":
    main()
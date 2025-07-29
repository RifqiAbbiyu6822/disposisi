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
from disposisi_app.views.components.export_utils import collect_form_data_safely

try:
    from config import EMAIL_HOST, EMAIL_PORT, EMAIL_HOST_USER, EMAIL_HOST_PASSWORD
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
        
        # Enhanced styling setup
        self.setup_enhanced_styles()
        
        try:
            self.iconbitmap('JapekELEVATED.ico')
        except:
            pass
        
        self.setup_windowed_fullscreen()
        self.is_windows = platform.system() == "Windows"
        setup_styles(self)  # Terapkan style global
        
        self.edit_mode = False
        self.current_focus = None
        self.last_mouse_position = (0, 0)
        self.touchpad_gesture_active = False
        self.gesture_start_y = 0
        self.tooltip = None  # Untuk tooltip
        
        self.init_variables()
        self.posisi_options = [
            "Direktur Utama",
            "Direktur Keuangan",
            "Direktur Teknik",
            "GM Keuangan & Administrasi",
            "GM Operasional & Pemeliharaan",
            "Manager"
        ]
        
        # Enhanced header with modern styling
        self.header = create_header(self)
        self.log_data_dirty = False
        
        # Inisialisasi pdf_attachments sebelum create_widgets
        self.pdf_attachments = []
        
        self.create_tabs()
        self.setup_shortcuts()
        self.setup_touchpad_gestures()
        
        # Enhanced menu bar
        create_menu_bar(
            self,
            save_to_pdf_callback=self.save_to_pdf,
            clear_form_callback=self.clear_form,
            quit_callback=self.quit,
            show_shortcuts_callback=self.show_shortcuts,
            show_about_callback=self.show_about
        )
        
        # Enhanced status bar
        self.status_frame, self.status_message = create_status_bar(self)
        
        self.center_window()
        self.bind_all("<MouseWheel>", self._global_on_mousewheel)
        self.bind_all("<Shift-MouseWheel>", self._global_on_shift_mousewheel)
        self.bind_all("<Button-4>", self._global_on_mousewheel)
        self.bind_all("<Button-5>", self._global_on_mousewheel)

    def setup_enhanced_styles(self):
        """Setup enhanced modern styling for the application"""
        self.style = ttk.Style()
        
        # Modern color palette
        colors = {
            'primary': '#3b82f6',      # Blue
            'primary_dark': '#1d4ed8', # Darker blue
            'secondary': '#6b7280',    # Gray
            'success': '#10b981',      # Green
            'warning': '#f59e0b',      # Amber
            'danger': '#ef4444',       # Red
            'background': '#f8fafc',   # Light gray background
            'surface': '#ffffff',      # White surface
            'text': '#1f2937',         # Dark gray text
            'text_secondary': '#6b7280', # Secondary text
            'border': '#e5e7eb',       # Light border
            'hover': '#f3f4f6'         # Hover background
        }
        
        self.style.configure('TFrame', background=colors['background'], relief='flat')
        self.style.configure('Surface.TFrame', background=colors['surface'], relief='solid', borderwidth=1)
        self.style.configure('Card.TFrame', background=colors['surface'], relief='solid', borderwidth=1)
        
        self.style.configure('TLabelframe', 
                             background=colors['surface'],
                             borderwidth=2,
                             relief='solid',
                             focuscolor='none')
        self.style.configure('TLabelframe.Label', 
                             background=colors['surface'],
                             foreground=colors['primary'],
                             font=('Segoe UI', 11, 'bold'),
                             padding=(10, 5))
        
        self.style.configure('TButton',
                             background=colors['primary'],
                             foreground='white',
                             font=('Segoe UI', 10, 'bold'),
                             borderwidth=0,
                             focuscolor='none',
                             padding=(15, 8))
        
        self.style.configure('Secondary.TButton',
                             background=colors['secondary'],
                             foreground='white',
                             font=('Segoe UI', 9),
                             borderwidth=0,
                             focuscolor='none',
                             padding=(12, 6))
        
        self.style.configure('Success.TButton',
                             background=colors['success'],
                             foreground='white',
                             font=('Segoe UI', 10, 'bold'),
                             borderwidth=0,
                             focuscolor='none',
                             padding=(15, 8))
        
        self.style.configure('Danger.TButton',
                             background=colors['danger'],
                             foreground='white',
                             font=('Segoe UI', 10),
                             borderwidth=0,
                             focuscolor='none',
                             padding=(12, 6))
        
        self.style.configure('TEntry',
                             fieldbackground=colors['surface'],
                             borderwidth=2,
                             insertcolor=colors['primary'],
                             relief='solid',
                             padding=(8, 6),
                             font=('Segoe UI', 10))
        
        self.style.configure('TCombobox',
                             fieldbackground=colors['surface'],
                             borderwidth=2,
                             relief='solid',
                             padding=(8, 6),
                             font=('Segoe UI', 10))
        
        self.style.configure('TNotebook', 
                             background=colors['background'],
                             borderwidth=0)
        self.style.configure('TNotebook.Tab',
                             background=colors['surface'],
                             foreground=colors['text'],
                             font=('Segoe UI', 11, 'bold'),
                             padding=(20, 12),
                             borderwidth=1,
                             relief='solid')
        
        self.style.configure('TLabel',
                             background=colors['background'],
                             foreground=colors['text'],
                             font=('Segoe UI', 10))
        
        self.style.configure('Heading.TLabel',
                             background=colors['background'],
                             foreground=colors['primary'],
                             font=('Segoe UI', 14, 'bold'))
        
        self.style.configure('TCheckbutton',
                             background=colors['surface'],
                             foreground=colors['text'],
                             font=('Segoe UI', 10),
                             focuscolor='none')
        
        self.style.configure('TText',
                             background=colors['surface'],
                             borderwidth=2,
                             relief='solid',
                             padding=8,
                             font=('Segoe UI', 10))

    def center_window(self):
        center_window(self)

    def setup_windowed_fullscreen(self):
        setup_windowed_fullscreen(self)

    def create_header(self):
        header_frame = ttk.Frame(self, style="Surface.TFrame", padding=(0, 0, 0, 0))
        header_frame.pack(fill="x", padx=0, pady=0)
        
        gradient_frame = tk.Frame(header_frame, height=80, bg="#3b82f6")
        gradient_frame.pack(fill="x")
        gradient_frame.pack_propagate(False)
        
        header_content = tk.Frame(gradient_frame, bg="#3b82f6")
        header_content.pack(fill="both", expand=True, padx=20, pady=15)
        
        left_frame = tk.Frame(header_content, bg="#3b82f6")
        left_frame.pack(side="left", fill="both", expand=True)
        
        self.logo_label = tk.Label(left_frame, 
                                   text="📋 DISPOSISI", 
                                   font=("Segoe UI", 16, "bold"), 
                                   bg="#3b82f6", 
                                   fg="white")
        self.logo_label.pack(anchor="w")
        
        self.title_label = tk.Label(left_frame, 
                                      text="Sistem Pembuatan dan Pelaporan Disposisi", 
                                      font=("Segoe UI", 22, "bold"), 
                                      bg="#3b82f6", 
                                      fg="white")
        self.title_label.pack(anchor="w", pady=(2, 0))
        
        self.subtitle_label = tk.Label(left_frame, 
                                       text="Pelaporan dan Pembuatan Disposisi Modern", 
                                       font=("Segoe UI", 12), 
                                       bg="#3b82f6", 
                                       fg="#bfdbfe")
        self.subtitle_label.pack(anchor="w", pady=(5, 0))
        
        right_frame = tk.Frame(header_content, bg="#3b82f6")
        right_frame.pack(side="right")
        
        version_label = tk.Label(right_frame, 
                                 text="Version 2.0", 
                                 font=("Segoe UI", 11, "bold"), 
                                 bg="#3b82f6", 
                                 fg="#bfdbfe")
        version_label.pack(anchor="e")
        
        status_label = tk.Label(right_frame, 
                                text="✓ Online", 
                                font=("Segoe UI", 10), 
                                bg="#3b82f6", 
                                fg="#10b981")
        status_label.pack(anchor="e", pady=(5, 0))

    def create_menu_bar(self):
        menubar = tk.Menu(self, font=("Segoe UI", 10), bg="#ffffff", fg="#1f2937", 
                          activebackground="#3b82f6", activeforeground="white")
        self.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10), 
                            bg="#ffffff", fg="#1f2937",
                            activebackground="#3b82f6", activeforeground="white")
        menubar.add_cascade(label="📁 File", menu=file_menu)
        file_menu.add_command(label="💾 Simpan ke PDF", command=self.save_to_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="🚪 Keluar", command=self.quit)
        
        edit_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10),
                            bg="#ffffff", fg="#1f2937",
                            activebackground="#3b82f6", activeforeground="white")
        menubar.add_cascade(label="✏️ Edit", menu=edit_menu)
        edit_menu.add_command(label="🗑️ Bersihkan Form", command=self.clear_form)
        
        help_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10),
                            bg="#ffffff", fg="#1f2937",
                            activebackground="#3b82f6", activeforeground="white")
        menubar.add_cascade(label="❓ Help", menu=help_menu)
        help_menu.add_command(label="⌨️ Shortcuts", command=self.show_shortcuts)
        help_menu.add_command(label="ℹ️ About", command=self.show_about)

    def create_status_bar(self):
        status_frame = tk.Frame(self, bg="#f8fafc", height=35, relief="solid", borderwidth=1)
        status_frame.pack(side="bottom", fill="x")
        status_frame.pack_propagate(False)
        
        left_status = tk.Frame(status_frame, bg="#f8fafc")
        left_status.pack(side="left", fill="both", expand=True)
        
        self.status_message = tk.Label(left_status, 
                                       text="✓ Ready", 
                                       font=("Segoe UI", 10), 
                                       bg="#f8fafc", 
                                       fg="#10b981")
        self.status_message.pack(side="left", padx=20, pady=8)
        
        right_status = tk.Frame(status_frame, bg="#f8fafc")
        right_status.pack(side="right")
        
        system_label = tk.Label(right_status, 
                                text=f"🖥️ {platform.system()}", 
                                font=("Segoe UI", 9), 
                                bg="#f8fafc", 
                                fg="#6b7280")
        system_label.pack(side="right", padx=10, pady=8)
        
        separator_frame = tk.Frame(status_frame, bg="#e5e7eb", width=1)
        separator_frame.pack(side="right", fill="y", padx=10)
        
        version_label = tk.Label(right_status, 
                                 text="v2.0", 
                                 font=("Segoe UI", 10, "bold"), 
                                 bg="#f8fafc", 
                                 fg="#3b82f6")
        version_label.pack(side="right", padx=15, pady=8)

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
            "gm_keu": tk.IntVar(), "gm_ops": tk.IntVar(), "manager": tk.IntVar(),
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
        
        middle_frame.columnconfigure(0, weight=1)
        middle_frame.columnconfigure(1, weight=1)
        middle_frame.columnconfigure(2, weight=2)
        middle_frame.rowconfigure(0, weight=1)
        
        frame_disposisi = ttk.LabelFrame(middle_frame, text="👥 Disposisi Kepada", 
                                         padding=(20, 15, 20, 20), style="TLabelframe")
        frame_disposisi.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        frame_disposisi.columnconfigure(0, weight=1)
        
        populate_frame_disposisi(frame_disposisi, self.vars)
        
        frame_instruksi = ttk.LabelFrame(middle_frame, text="📋 Untuk di", 
                                         padding=(20, 15, 20, 20), style="TLabelframe")
        frame_instruksi.grid(row=0, column=1, sticky="nsew", padx=(8, 8))
        frame_instruksi.columnconfigure(0, weight=1)
        
        input_widgets.update(populate_frame_instruksi(frame_instruksi, self.vars))
        
        self.harap_selesai_tgl_entry = input_widgets.get("harap_selesai_tgl_entry")
        
        frame_info = ttk.LabelFrame(middle_frame, text="📝 Isi Instruksi / Informasi", 
                                    padding=(20, 15, 20, 20), style="TLabelframe")
        frame_info.grid(row=0, column=2, sticky="nsew", padx=(8, 0))
        frame_info.columnconfigure(0, weight=1)
        frame_info.rowconfigure(0, weight=1)
        
        frame_instruksi_table = ttk.Frame(frame_info, style="Surface.TFrame")
        frame_instruksi_table.grid(row=0, column=0, sticky="nsew")
        frame_instruksi_table.columnconfigure(0, weight=1)
        frame_instruksi_table.rowconfigure(0, weight=1)
        
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
            "send_email": lambda recipients: send_email_with_disposisi(self, recipients),
            "get_disposisi_labels": self.get_disposisi_labels,
            "clear_form": self.clear_form
        }
        self._button_frame = create_button_frame(
            parent,
            callbacks
        )
        return self._button_frame

    def send_email_with_disposisi(self, recipients):
        """
        Generates the disposition PDF, attaches it, and sends it to the specified recipients.
        """
        try:
            from config import EMAIL_HOST, EMAIL_PORT, EMAIL_HOST_USER, EMAIL_HOST_PASSWORD
            if not all([EMAIL_HOST, EMAIL_PORT, EMAIL_HOST_USER, EMAIL_HOST_PASSWORD]):
                raise ImportError("Email configuration incomplete")
        except ImportError:
            import os
            EMAIL_HOST = os.getenv('EMAIL_HOST', 'smtp.gmail.com')
            EMAIL_PORT = int(os.getenv('EMAIL_PORT', '587'))
            EMAIL_HOST_USER = os.getenv('EMAIL_HOST_USER', '')
            EMAIL_HOST_PASSWORD = os.getenv('EMAIL_HOST_PASSWORD', '')
            
            if not EMAIL_HOST_USER or not EMAIL_HOST_PASSWORD:
                try:
                    from email_sender.config import SENDER_EMAIL, SENDER_PASSWORD
                    EMAIL_HOST_USER = SENDER_EMAIL
                    EMAIL_HOST_PASSWORD = SENDER_PASSWORD
                except:
                    messagebox.showerror("Email Error", 
                                         "Email belum dikonfigurasi. Pastikan file config.py atau .env berisi konfigurasi email.")
                    return

        self.update_status("Preparing email...")
        
        data = collect_form_data_safely(self)
        if not data.get("no_surat", "").strip():
            messagebox.showerror("Validation Error", "No. Surat tidak boleh kosong untuk mengirim email.")
            return

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
            
        try:
            msg = EmailMessage()
            msg['Subject'] = f'Disposisi Surat: {data.get("perihal", "N/A")}'
            msg['From'] = EMAIL_HOST_USER
            msg['To'] = ", ".join(recipients)
            
            body = f"""Yth. Bapak/Ibu,

Berikut terlampir lembar disposisi untuk surat dengan detail:
• Perihal: {data.get('perihal', 'N/A')}
• Nomor Surat: {data.get('no_surat', 'N/A')}
• Asal Surat: {data.get('asal_surat', 'N/A')}

Mohon untuk dapat ditindaklanjuti sesuai dengan disposisi terlampir.

Terima kasih.

--
Sistem Disposisi Otomatis
PT Jasamarga Jalanlayang Cikampek"""
            
            msg.set_content(body)

            with open(final_pdf_path, 'rb') as f:
                file_data = f.read()
                file_name = f"Disposisi_{data.get('no_surat', 'surat').replace('/', '_')}.pdf"
                msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

            self.update_status(f"Mengirim email ke {', '.join(recipients)}...")
            
            if EMAIL_PORT == 465:
                import smtplib
                with smtplib.SMTP_SSL(EMAIL_HOST, EMAIL_PORT) as server:
                    server.login(EMAIL_HOST_USER, EMAIL_HOST_PASSWORD)
                    server.send_message(msg)
            else:
                import smtplib
                with smtplib.SMTP(EMAIL_HOST, EMAIL_PORT) as server:
                    server.starttls()
                    server.login(EMAIL_HOST_USER, EMAIL_HOST_PASSWORD)
                    server.send_message(msg)
            
            self.update_status("Email berhasil dikirim!")
            messagebox.showinfo("Email Sent", f"Email berhasil dikirim ke: {', '.join(recipients)}")

        except Exception as e:
            self.update_status("Gagal mengirim email.")
            messagebox.showerror("Email Error", f"Gagal mengirim email: {e}")
            traceback.print_exc()
            
        finally:
            if 'temp_pdf_path' in locals() and os.path.exists(temp_pdf_path):
                try:
                    os.remove(temp_pdf_path)
                except:
                    pass
            if 'final_pdf_path' in locals() and final_pdf_path != temp_pdf_path and os.path.exists(final_pdf_path):
                try:
                    os.remove(final_pdf_path)
                except:
                    pass

    def refresh_pdf_attachments(self, parent):
        if hasattr(self, '_button_frame') and self._button_frame:
            self._button_frame.destroy()
        self.create_button_frame(parent)

    def add_pdf_attachment(self):
        files = filedialog.askopenfilenames(
            title="Pilih PDF lampiran",
            filetypes=[("PDF Documents", "*.pdf"), ("All Files", "*.*")]
        )
        for f in files:
            if f not in self.pdf_attachments:
                self.pdf_attachments.append(f)
                filename = os.path.basename(f)
                display_name = f"📄 {filename}"
                self.attachment_listbox.insert(tk.END, display_name)
        
        if files:
            self.update_status(f"✓ {len(files)} PDF ditambahkan")

    def remove_pdf_attachment(self):
        selected = self.attachment_listbox.curselection()
        if selected:
            idx = selected[0]
            filename = os.path.basename(self.pdf_attachments[idx])
            self.attachment_listbox.delete(idx)
            del self.pdf_attachments[idx]
            self.update_status(f"🗑️ {filename} dihapus")
        else:
            messagebox.showwarning("Hapus Lampiran", "Pilih file yang akan dihapus terlebih dahulu.")

    def get_disposisi_labels(self):
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

    def save_to_pdf(self):
        save_to_pdf(self)

    def save_to_sheet(self):
        save_to_sheet(self)

    def load_entry_to_form_and_switch(self, entry):
        import threading
        loading = LoadingScreen(self)
        def do_load():
            try:
                no_surat = entry.get("No. Surat", "")
                data = get_log_entry_by_no_surat(str(no_surat))
                if not data:
                    messagebox.showerror("Edit Log", "Data tidak ditemukan di Google Sheets.")
                    return
                self.notebook.select(1)
                self.update()
            finally:
                loading.destroy()
        threading.Thread(target=do_load).start()

    def clear_form(self):
        clear_form(self)

    def upload_to_sheet(self, call_from_pdf=False, data_override=None):
        try:
            return upload_to_sheet(self, call_from_pdf=call_from_pdf, data_override=data_override)
        finally:
            pass

    def get_untuk_di_labels(self, data):
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

    def _global_on_mousewheel(self, event):
        try:
            current_tab = self.notebook.index(self.notebook.select())
        except Exception:
            return
        if current_tab == 0:
            try:
                canvas = self.form_frame.winfo_children()[0]
                if isinstance(canvas, tk.Canvas):
                    if event.num == 5 or event.delta == -120:
                        canvas.yview_scroll(1, "units")
                    elif event.num == 4 or event.delta == 120:
                        canvas.yview_scroll(-1, "units")
                    else:
                        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except Exception:
                pass
        elif current_tab == 1:
            try:
                tree = self.log_frame.tree
                if hasattr(tree, 'yview_scroll'):
                    if event.num == 5 or event.delta == -120:
                        tree.yview_scroll(1, "units")
                    elif event.num == 4 or event.delta == 120:
                        tree.yview_scroll(-1, "units")
                    else:
                        tree.yview_scroll(int(-1*(event.delta/120)), "units")
            except Exception:
                pass
        elif current_tab == 2:
            try:
                edit_tab = self.notebook.nametowidget(self.notebook.select())
                children = edit_tab.winfo_children()
                if children and isinstance(children[0], tk.Canvas):
                    canvas = children[0]
                    if event.num == 5 or event.delta == -120:
                        canvas.yview_scroll(1, "units")
                    elif event.num == 4 or event.delta == 120:
                        canvas.yview_scroll(-1, "units")
                    else:
                        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except Exception:
                pass
    
    def _global_on_shift_mousewheel(self, event):
        try:
            current_tab = self.notebook.index(self.notebook.select())
        except Exception:
            return
        if current_tab == 0:
            try:
                canvas = self.form_frame.winfo_children()[0]
                if isinstance(canvas, tk.Canvas):
                    canvas.xview_scroll(int(-1*(event.delta/120)), "units")
            except Exception:
                pass
        elif current_tab == 1:
            try:
                tree = self.log_frame.tree
                if hasattr(tree, 'xview_scroll'):
                    tree.xview_scroll(int(-1*(event.delta/120)), "units")
            except Exception:
                pass
        elif current_tab == 2:
            try:
                edit_tab = self.notebook.nametowidget(self.notebook.select())
                children = edit_tab.winfo_children()
                if children and isinstance(children[0], tk.Canvas):
                    canvas = children[0]
                    canvas.xview_scroll(int(-1*(event.delta/120)), "units")
            except Exception:
                pass

if __name__ == "__main__":
    app = FormApp()
    app.mainloop()
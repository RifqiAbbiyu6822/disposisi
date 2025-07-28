import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Text
import logging
import traceback
import platform
import sys
import os
from google_sheets_connect import append_row_to_sheet, create_new_sheet, write_multilayer_header, append_rows_to_sheet, get_sheets_service, SHEET_ID
from sheet_logic import upload_to_sheet, get_untuk_di_labels, get_disposisi_labels, get_log_entry_by_no_surat
from logic.instruksi_table import InstruksiTable
import time
from tkcalendar import DateEntry
from main_app.edit_tab import EditTab
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

# --- LoadingScreen class ---
# (Sudah dipindah ke loading_screen.py)

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
        
        # Enhanced frame styles
        self.style.configure('TFrame', background=colors['background'], relief='flat')
        self.style.configure('Surface.TFrame', background=colors['surface'], relief='solid', borderwidth=1)
        self.style.configure('Card.TFrame', background=colors['surface'], relief='solid', borderwidth=1)
        
        # Enhanced LabelFrame styles with modern borders
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
        
        # Enhanced button styles with modern gradients and hover effects
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
        
        # Enhanced entry styles with modern borders
        self.style.configure('TEntry',
                           fieldbackground=colors['surface'],
                           borderwidth=2,
                           insertcolor=colors['primary'],
                           relief='solid',
                           padding=(8, 6),
                           font=('Segoe UI', 10))
        
        # Enhanced combobox styles
        self.style.configure('TCombobox',
                           fieldbackground=colors['surface'],
                           borderwidth=2,
                           relief='solid',
                           padding=(8, 6),
                           font=('Segoe UI', 10))
        
        # Enhanced notebook styles with modern tabs
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
        
        # Enhanced label styles
        self.style.configure('TLabel',
                           background=colors['background'],
                           foreground=colors['text'],
                           font=('Segoe UI', 10))
        
        self.style.configure('Heading.TLabel',
                           background=colors['background'],
                           foreground=colors['primary'],
                           font=('Segoe UI', 14, 'bold'))
        
        # Enhanced checkbutton styles
        self.style.configure('TCheckbutton',
                           background=colors['surface'],
                           foreground=colors['text'],
                           font=('Segoe UI', 10),
                           focuscolor='none')
        
        # Enhanced text widget styling
        self.style.configure('TText',
                           background=colors['surface'],
                           borderwidth=2,
                           relief='solid',
                           padding=8,
                           font=('Segoe UI', 10))

    def center_window(self):
        center_window(self)

    def setup_windowed_fullscreen(self):
        """Membuat window fullscreen tapi tetap windowed (ada titlebar dan tombol window)."""
        setup_windowed_fullscreen(self)

    def create_header(self):
        """Enhanced header with modern gradient background and better typography"""
        header_frame = ttk.Frame(self, style="Surface.TFrame", padding=(0, 0, 0, 0))
        header_frame.pack(fill="x", padx=0, pady=0)
        
        # Add subtle gradient effect using multiple frames
        gradient_frame = tk.Frame(header_frame, height=80, bg="#3b82f6")
        gradient_frame.pack(fill="x")
        gradient_frame.pack_propagate(False)
        
        header_content = tk.Frame(gradient_frame, bg="#3b82f6")
        header_content.pack(fill="both", expand=True, padx=20, pady=15)
        
        # Left side - Icon and main title
        left_frame = tk.Frame(header_content, bg="#3b82f6")
        left_frame.pack(side="left", fill="both", expand=True)
        
        # Enhanced logo with modern styling
        self.logo_label = tk.Label(left_frame, 
                                  text="📋 DISPOSISI", 
                                  font=("Segoe UI", 16, "bold"), 
                                  bg="#3b82f6", 
                                  fg="white")
        self.logo_label.pack(anchor="w")
        
        # Enhanced title with better hierarchy
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
        
        # Right side - Version and status info
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
        """Enhanced menu bar with modern styling and better organization"""
        menubar = tk.Menu(self, font=("Segoe UI", 10), bg="#ffffff", fg="#1f2937", 
                         activebackground="#3b82f6", activeforeground="white")
        self.config(menu=menubar)
        
        # File menu with icons (using unicode symbols)
        file_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10), 
                           bg="#ffffff", fg="#1f2937",
                           activebackground="#3b82f6", activeforeground="white")
        menubar.add_cascade(label="📁 File", menu=file_menu)
        file_menu.add_command(label="💾 Simpan ke PDF", command=self.save_to_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="🚪 Keluar", command=self.quit)
        
        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10),
                           bg="#ffffff", fg="#1f2937",
                           activebackground="#3b82f6", activeforeground="white")
        menubar.add_cascade(label="✏️ Edit", menu=edit_menu)
        edit_menu.add_command(label="🗑️ Bersihkan Form", command=self.clear_form)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10),
                           bg="#ffffff", fg="#1f2937",
                           activebackground="#3b82f6", activeforeground="white")
        menubar.add_cascade(label="❓ Help", menu=help_menu)
        help_menu.add_command(label="⌨️ Shortcuts", command=self.show_shortcuts)
        help_menu.add_command(label="ℹ️ About", command=self.show_about)

    def create_status_bar(self):
        """Enhanced status bar with modern styling and better visual hierarchy"""
        status_frame = tk.Frame(self, bg="#f8fafc", height=35, relief="solid", borderwidth=1)
        status_frame.pack(side="bottom", fill="x")
        status_frame.pack_propagate(False)
        
        # Left side - Status message with icon
        left_status = tk.Frame(status_frame, bg="#f8fafc")
        left_status.pack(side="left", fill="both", expand=True)
        
        self.status_message = tk.Label(left_status, 
                                      text="✓ Ready", 
                                      font=("Segoe UI", 10), 
                                      bg="#f8fafc", 
                                      fg="#10b981")
        self.status_message.pack(side="left", padx=20, pady=8)
        
        # Right side - Version and system info
        right_status = tk.Frame(status_frame, bg="#f8fafc")
        right_status.pack(side="right")
        
        system_label = tk.Label(right_status, 
                               text=f"🖥️ {platform.system()}", 
                               font=("Segoe UI", 9), 
                               bg="#f8fafc", 
                               fg="#6b7280")
        system_label.pack(side="right", padx=10, pady=8)
        
        # Separator
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
        """Handle mouse down event untuk gesture detection."""
        self.last_mouse_position = (event.x, event.y)
        self.gesture_start_y = event.y
        self.touchpad_gesture_active = True

    def on_mouse_up(self, event):
        """Handle mouse up event."""
        self.touchpad_gesture_active = False

    def on_mouse_drag(self, event):
        """Handle mouse drag untuk gesture detection."""
        if not self.touchpad_gesture_active:
            return
            
        delta_y = event.y - self.gesture_start_y
        
        # Vertical scroll gesture
        if abs(delta_y) > 50:
            if delta_y > 0:
                self.scroll_down()
            else:
                self.scroll_up()
            self.gesture_start_y = event.y

    def on_double_click(self, event):
        """Handle double click untuk zoom."""
        pass

    def on_mousewheel(self, event):
        """Handle mousewheel untuk scrolling."""
        pass

    def on_horizontal_scroll(self, event):
        """Handle horizontal scroll."""
        pass

    def on_pinch_gesture_start(self, event):
        """Start pinch gesture detection."""
        self.pinch_start_distance = 0

    def on_pinch_gesture(self, event):
        """Handle pinch gesture untuk zoom."""
        # Simulate pinch gesture with Ctrl+drag
        pass

    def scroll_up(self):
        """Scroll up gesture."""
        pass

    def scroll_down(self):
        """Scroll down gesture."""
        pass

    def refresh_data(self):
        """Refresh data in history view."""
        # Hapus fungsi ini karena tidak ada riwayat
        pass

    def toggle_fullscreen(self):
        """Toggle fullscreen mode."""
        self.attributes('-fullscreen', not self.attributes('-fullscreen'))
        self.update_status("Fullscreen toggled")

    def export_excel(self):
        """Export data to Excel."""
        # Hapus fungsi ini karena tidak ada riwayat
        messagebox.showwarning("Export", "No data available to export")

    def show_shortcuts(self):
        show_shortcuts(self)

    def show_about(self):
        show_about(self)

    def update_status(self, message):
        if hasattr(self, 'status_message'):
            update_status(self.status_message, message, self)

    def create_tabs(self):
        """Enhanced tabs with modern styling and better organization"""
        notebook_frame = ttk.Frame(self, style="TFrame", padding=(10, 5, 10, 5))
        notebook_frame.pack(fill="both", expand=True, padx=0, pady=0)
        
        # Enhanced grid configuration
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        notebook_frame.rowconfigure(0, weight=1)
        notebook_frame.columnconfigure(0, weight=1)
        
        # Enhanced notebook with modern tab styling
        self.notebook = ttk.Notebook(notebook_frame, style="TNotebook")
        self.notebook.grid(row=0, column=0, sticky="nsew")
        
        # Enhanced form frame with card styling
        self.form_frame = ttk.Frame(self.notebook, style="Surface.TFrame")
        self.form_frame.rowconfigure(0, weight=1)
        self.form_frame.columnconfigure(0, weight=1)
        self.notebook.add(self.form_frame, text="📝 Formulir Disposisi")
        
        self.form_input_widgets = self.create_widgets(self.form_frame)
        
        # Enhanced log tab
        from excel_crud_tab import LogTab
        self.log_frame = LogTab(self.notebook, on_edit_log=self.open_edit_tab)
        self.notebook.add(self.log_frame, text="📊 Data & Log")
        
        def on_tab_changed(event):
            tab_id = event.widget.index("current")
            if tab_id == 1:  # Tab Log
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
        """Inisialisasi variabel data UI."""
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
        # Instruction table: 5 rows × 3 columns (Instruksi, Posisi, Tanggal)
        self.instruction_vars = {}
        self.posisi_labels = [f"Posisi{i+1}" for i in range(5)]  # Bisa diubah nanti
        for row in range(5):
            self.instruction_vars[f"instruksi_{row}"] = tk.StringVar()
            self.instruction_vars[f"posisi_{row}"] = tk.StringVar()
            self.instruction_vars[f"tanggal_{row}"] = tk.StringVar()

    def create_widgets(self, parent):
        """Enhanced widget creation with modern card-based layout"""
        input_widgets = {}
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        
        # Enhanced canvas with modern scrolling
        canvas = tk.Canvas(parent, borderwidth=0, highlightthickness=0, bg="#f8fafc")
        
        # Modern scrollbars with enhanced styling
        v_scroll = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(parent, orient="horizontal", command=canvas.xview)
        
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        # Enhanced grid configuration
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        
        # Enhanced main frame with modern card styling and better spacing
        main_frame = ttk.Frame(canvas, padding=(20, 20, 20, 20), style="Surface.TFrame")
        self._form_main_frame = main_frame
        
        form_window = canvas.create_window((0, 0), window=main_frame, anchor="nw")
        
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        main_frame.bind("<Configure>", on_frame_configure)
        
        # Enhanced scrolling functions
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
        
        # Enhanced form sections with better spacing
        input_widgets.update(self.create_top_frame(main_frame))
        input_widgets.update(self.create_middle_frame(main_frame))
        
        # Pastikan attachment_listbox sudah dibuat sebelum create_button_frame
        if not hasattr(self, "attachment_listbox"):
            self.attachment_listbox = tk.Listbox(main_frame, height=2, selectmode=tk.SINGLE, 
                                               font=("Segoe UI", 9),
                                               bg="#ffffff", fg="#1f2937",
                                               selectbackground="#3b82f6",
                                               relief="solid", borderwidth=1)
        
        self.create_button_frame(main_frame)
        self.add_tooltips(input_widgets)
        
        return input_widgets

    def add_tooltips(self, input_widgets):
        add_tooltips(input_widgets, TOOLTIP_LABELS)

    def attach_tooltip(self, widget, text):
        attach_tooltip(widget, text)

    def create_top_frame(self, parent):
        """Enhanced top frame with modern card-based sections"""
        input_widgets = {}
        
        # Enhanced top frame with better spacing
        top_frame = ttk.Frame(parent, style="TFrame")
        top_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 20))
        top_frame.columnconfigure(0, weight=1)
        top_frame.columnconfigure(1, weight=1)
        top_frame.rowconfigure(0, weight=1)
        
        # Enhanced left frame with modern card styling
        frame_kiri = ttk.LabelFrame(top_frame, text="📄 Detail Surat", 
                                   padding=(20, 15, 20, 20), style="TLabelframe")
        frame_kiri.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        frame_kiri.columnconfigure(1, weight=1)
        
        input_widgets.update(populate_frame_kiri(frame_kiri, self.vars))
        
        # Enhanced right frame with modern card styling
        frame_kanan = ttk.LabelFrame(top_frame, text="🏷️ Klasifikasi & Status", 
                                    padding=(20, 15, 20, 20), style="TLabelframe")
        frame_kanan.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        frame_kanan.columnconfigure(0, weight=1)
        
        self.tgl_terima_entry = populate_frame_kanan(frame_kanan, self.vars)
        
        return input_widgets

    def create_middle_frame(self, parent):
        """Enhanced middle frame with modern three-column card layout"""
        input_widgets = {}
        
        # Enhanced middle frame with better spacing and modern layout
        middle_frame = ttk.Frame(parent, style="TFrame")
        middle_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 20))
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)
        
        # Enhanced grid configuration for responsive design
        middle_frame.columnconfigure(0, weight=1)
        middle_frame.columnconfigure(1, weight=1)
        middle_frame.columnconfigure(2, weight=2)
        middle_frame.rowconfigure(0, weight=1)
        
        # Enhanced disposisi frame with modern card styling
        frame_disposisi = ttk.LabelFrame(middle_frame, text="👥 Disposisi Kepada", 
                                        padding=(20, 15, 20, 20), style="TLabelframe")
        frame_disposisi.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        frame_disposisi.columnconfigure(0, weight=1)
        
        populate_frame_disposisi(frame_disposisi, self.vars)
        
        # Enhanced instruksi frame with modern card styling
        frame_instruksi = ttk.LabelFrame(middle_frame, text="📋 Untuk di", 
                                        padding=(20, 15, 20, 20), style="TLabelframe")
        frame_instruksi.grid(row=0, column=1, sticky="nsew", padx=(8, 8))
        frame_instruksi.columnconfigure(0, weight=1)
        
        input_widgets.update(populate_frame_instruksi(frame_instruksi, self.vars))
        
        # Simpan harap_selesai_tgl_entry ke atribut agar bisa diakses
        self.harap_selesai_tgl_entry = input_widgets.get("harap_selesai_tgl_entry")
        
        # Enhanced info frame with modern card styling
        frame_info = ttk.LabelFrame(middle_frame, text="📝 Isi Instruksi / Informasi", 
                                   padding=(20, 15, 20, 20), style="TLabelframe")
        frame_info.grid(row=0, column=2, sticky="nsew", padx=(8, 0))
        frame_info.columnconfigure(0, weight=1)
        frame_info.rowconfigure(0, weight=1)
        
        # Enhanced instruction table frame
        frame_instruksi_table = ttk.Frame(frame_info, style="Surface.TFrame")
        frame_instruksi_table.grid(row=0, column=0, sticky="nsew")
        frame_instruksi_table.columnconfigure(0, weight=1)
        frame_instruksi_table.rowconfigure(0, weight=1)
        
        self.instruksi_table = InstruksiTable(frame_instruksi_table, self.posisi_options, use_grid=True)
        
        # Enhanced button frame with modern styling
        btn_frame = ttk.Frame(frame_info, style="TFrame")
        btn_frame.grid(row=1, column=0, pady=(15, 0), sticky="ew")
        
        # Enhanced buttons with modern styling and icons
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
        # Use the new unified button frame from button_frame.py
        from disposisi_app.views.components.button_frame import create_button_frame
        def on_selesai():
            # Show preview and export/email options
            if hasattr(self, 'preview_frame'):
                self.preview_frame.grid()
        def export_pdf():
            self.save_to_pdf()
        def upload_sheet():
            self.save_to_sheet()
        def on_remove_pdf(idx):
            if 0 <= idx < len(self.pdf_attachments):
                del self.pdf_attachments[idx]
                self.create_button_frame(parent)
        callbacks = {
            "on_selesai": on_selesai,
            "export_pdf": export_pdf,
            "upload_sheet": upload_sheet,
        }
        self._button_frame = create_button_frame(
            parent,
            callbacks,
            self.pdf_attachments,
            on_remove_pdf
        )
        return self._button_frame

    def refresh_pdf_attachments(self, parent):
        # Hapus frame lama dan render ulang
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
                # Enhanced display in listbox
                filename = os.path.basename(f)
                display_name = f"📄 {filename}"
                self.attachment_listbox.insert(tk.END, display_name)
        
        # Update status
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
        # Ambil label dari checkbox disposisi yang dicentang
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
                # Pindah ke tab Log saja, tidak ada tab Edit
                self.notebook.select(1)
                self.update()  # Paksa render widget LogTab
                # Tidak perlu fill_form ke EditTab
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
        # Gabungkan label "Untuk Di :" yang dicentang, termasuk Edarkan, Sesuai Disposisi, dst, dipisah koma
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
        # Deteksi tab aktif
        try:
            current_tab = self.notebook.index(self.notebook.select())
        except Exception:
            return
        if current_tab == 0:
            # Tab Form
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
            # Tab Log
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
            # Tab Edit (jika ada)
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
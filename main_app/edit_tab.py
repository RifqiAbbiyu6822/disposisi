# main_app/edit_tab.py - FIXED VERSION with better layout
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
import os

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
            "dir_utama", "dir_keu", "dir_teknik", "gm_keu", "gm_ops", 
            "manager_pemeliharaan", "manager_operasional", "manager_administrasi", "manager_keuangan",
            "ketahui_file", "proses_selesai", "teliti_pendapat", "buatkan_resume",
            "edarkan", "sesuai_disposisi", "bicarakan_saya", "bicarakan_dengan",
            "teruskan_kepada", "harap_selesai_tgl"
        ]
        
        for key in keys:
            if key in ["rahasia", "penting", "segera", "dir_utama", "dir_keu", "dir_teknik", 
                      "gm_keu", "gm_ops", "manager_pemeliharaan", "manager_operasional", 
                      "manager_administrasi", "manager_keuangan", "ketahui_file", "proses_selesai", 
                      "teliti_pendapat", "buatkan_resume", "edarkan", "sesuai_disposisi", 
                      "bicarakan_saya"]:
                self.vars[key] = tk.IntVar()
            else:
                self.vars[key] = tk.StringVar()

    def _create_widgets(self):
        # FIXED: Better layout structure with proper spacing
        root = self.winfo_toplevel()
        
        # Main scrollable container
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, bg='#f8fafc')
        self.v_scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.h_scroll = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        
        # Grid layout for scrollable area
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll.grid(row=1, column=0, sticky="ew")
        
        # FIXED: Proper main frame with better spacing
        self.main_frame = ttk.Frame(self.canvas, padding=(10, 10, 10, 10), style="TFrame")  # Reduced padding
        self._form_main_frame = self.main_frame
        self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        
        # Configure resizing
        self.main_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # FIXED: Create compact header
        self._create_compact_header()
        
        # Create the three main sections with better spacing
        self._create_top_frame()
        self._create_middle_frame()
        self._create_attachment_frame()  # NEW: Add attachment frame
        self._create_action_bar()  # Changed from button frame to action bar
        
        # Mouse wheel bindings
        bind_mousewheel_recursive(self.main_frame, self._on_mousewheel)

    def _create_compact_header(self):
        """FIXED: Create a compact header for edit mode"""
        header_frame = ttk.Frame(self.main_frame, style="TFrame")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        # Header content
        header_content = tk.Frame(header_frame, bg="#3B82F6", height=40)  # Compact height
        header_content.pack(fill="x")
        header_content.pack_propagate(False)
        
        # Title with icon
        title_frame = tk.Frame(header_content, bg="#3B82F6")
        title_frame.pack(expand=True, fill="both")
        
        title_label = tk.Label(
            title_frame,
            text="✏️ Edit Mode - Formulir Disposisi",
            font=("Segoe UI", 14, "bold"),  # Smaller font
            bg="#3B82F6",
            fg="white"
        )
        title_label.pack(expand=True, pady=8)  # Reduced padding

    def _create_top_frame(self):
        # FIXED: More compact top frame
        self.top_frame = ttk.Frame(self.main_frame, style="TFrame")
        self.top_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))  # Reduced padding
        self.top_frame.columnconfigure(0, weight=1)
        self.top_frame.columnconfigure(1, weight=1)
        
        # Left frame - Detail Surat with reduced padding
        self.frame_kiri = ttk.LabelFrame(self.top_frame, text="Detail Surat", padding=(10, 8, 10, 8), style="TLabelframe")  # Reduced padding
        self.frame_kiri.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        self.form_input_widgets.update(populate_frame_kiri(self.frame_kiri, self.vars))
        
        # Right frame - Klasifikasi with reduced padding
        self.frame_kanan = ttk.LabelFrame(self.top_frame, text="Klasifikasi", padding=(10, 8, 10, 8), style="TLabelframe")  # Reduced padding
        self.frame_kanan.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        kanan_widgets = populate_frame_kanan(self.frame_kanan, self.vars)
        if isinstance(kanan_widgets, dict):
            self.form_input_widgets.update(kanan_widgets)
        elif kanan_widgets:
            self.tgl_terima_entry = kanan_widgets
            self.form_input_widgets["tgl_terima"] = kanan_widgets

    def _create_middle_frame(self):
        # FIXED: More compact middle frame with better spacing
        self.middle_frame = ttk.Frame(self.main_frame, style="TFrame")
        self.middle_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 10))  # Reduced padding
        
        # Configure grid weights - Fixed: More balanced weights
        self.middle_frame.columnconfigure(0, weight=1)  # Disposisi Kepada
        self.middle_frame.columnconfigure(1, weight=1)  # Untuk di
        self.middle_frame.columnconfigure(2, weight=1)  # Isi Instruksi - Fixed: Changed from 2 to 1
        self.middle_frame.rowconfigure(0, weight=1)
        
        # Left column - Disposisi Kepada with proper padding
        self.frame_disposisi = ttk.LabelFrame(self.middle_frame, text="Disposisi Kepada", padding=(10, 8, 10, 8), style="TLabelframe")
        self.frame_disposisi.grid(row=0, column=0, sticky="nsew", padx=(0, 8))  # Fixed: Increased padding
        self.frame_disposisi.grid_columnconfigure(0, weight=1)  # Fixed: Added column weight
        populate_frame_disposisi(self.frame_disposisi, self.vars)
        
        # Middle column - Untuk di with proper padding
        self.frame_instruksi = ttk.LabelFrame(self.middle_frame, text="Untuk di", padding=(10, 8, 10, 8), style="TLabelframe")
        self.frame_instruksi.grid(row=0, column=1, sticky="nsew", padx=(8, 8))  # Fixed: Increased padding
        self.frame_instruksi.grid_columnconfigure(0, weight=1)  # Fixed: Added column weight
        self.frame_instruksi.grid_columnconfigure(1, weight=1)  # Fixed: Added column weight
        instruksi_widgets = populate_frame_instruksi(self.frame_instruksi, self.vars)
        self.form_input_widgets.update(instruksi_widgets)
        if "harap_selesai_tgl" in instruksi_widgets:
            self.harap_selesai_tgl_entry = instruksi_widgets["harap_selesai_tgl"]
        
        # Right column - Isi Instruksi/Informasi with proper padding
        self.frame_info = ttk.LabelFrame(self.middle_frame, text="Isi Instruksi / Informasi", padding=(10, 8, 10, 8), style="TLabelframe")
        self.frame_info.grid(row=0, column=2, sticky="nsew", padx=(8, 0))  # Fixed: Increased padding
        self.frame_info.grid_columnconfigure(0, weight=1)  # Fixed: Added column weight
        self.frame_info.grid_rowconfigure(0, weight=1)  # Fixed: Added row weight
        
        # Instruction table inside frame_info
        self.posisi_options = [
            "Direktur Utama", "Direktur Keuangan", "Direktur Teknik",
            "GM Keuangan & Administrasi", "GM Operasional & Pemeliharaan",
            "Manager Pemeliharaan", "Manager Operasional", "Manager Administrasi", "Manager Keuangan"
        ]
        self.frame_instruksi_table = ttk.Frame(self.frame_info)
        self.frame_instruksi_table.grid(row=0, column=0, sticky="nsew")
        self.frame_instruksi_table.grid_columnconfigure(0, weight=1)  # Fixed: Added column weight
        self.frame_instruksi_table.grid_rowconfigure(0, weight=1)  # Fixed: Added row weight
        self.instruksi_table = InstruksiTable(self.frame_instruksi_table, self.posisi_options, use_grid=True)
        
        # FIXED: Compact button row for instruction table
        self.frame_instruksi_btn = ttk.Frame(self.frame_info)
        self.frame_instruksi_btn.grid(row=1, column=0, sticky="ew", pady=(5,0))
        
        # Simplified buttons - only essential ones
        self.btn_tambah_baris = ttk.Button(self.frame_instruksi_btn, text="+ Tambah", command=self._on_tambah_baris, style="Secondary.TButton")
        self.btn_tambah_baris.pack(side="left", padx=(0, 5))
        
        self.btn_hapus_baris = ttk.Button(self.frame_instruksi_btn, text="🗑 Hapus", command=self._on_hapus_baris, style="Secondary.TButton")
        self.btn_hapus_baris.pack(side="left")

    def _create_attachment_frame(self):
        """SEDERHANAKAN: Attachment frame yang lebih minimalis"""
        self.attachment_frame = ttk.LabelFrame(self.main_frame, text="📎 Lampiran PDF", padding=(10, 8, 10, 8), style="TLabelframe")
        self.attachment_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        self.attachment_frame.columnconfigure(0, weight=1)
        
        # Attachment controls frame - simplified
        controls_frame = ttk.Frame(self.attachment_frame)
        controls_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        
        # Add attachment button only
        self.btn_add_attachment = ttk.Button(controls_frame, text="📎 Tambah PDF", command=self._add_pdf_attachment, style="Secondary.TButton")
        self.btn_add_attachment.pack(side="left")
        
        # Attachment listbox with scrollbar
        listbox_frame = ttk.Frame(self.attachment_frame)
        listbox_frame.grid(row=1, column=0, sticky="ew")
        listbox_frame.columnconfigure(0, weight=1)
        
        # Create listbox and scrollbar
        self.attachment_listbox = tk.Listbox(listbox_frame, height=3, selectmode=tk.SINGLE, font=("Segoe UI", 9))
        self.attachment_scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=self.attachment_listbox.yview)
        self.attachment_listbox.configure(yscrollcommand=self.attachment_scrollbar.set)
        
        self.attachment_listbox.grid(row=0, column=0, sticky="ew")
        self.attachment_scrollbar.grid(row=0, column=1, sticky="ns")

    def _add_pdf_attachment(self):
        """Add PDF attachment to the form"""
        from tkinter import filedialog
        
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

    def _remove_pdf_attachment(self):
        """Remove selected PDF attachment from the form"""
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

    def _create_action_bar(self):
        """SEDERHANAKAN: Action bar dengan hanya tombol Selesai dan Batal"""
        # Main action bar container
        action_container = tk.Frame(self.main_frame, bg="#FFFFFF", relief="flat")
        action_container.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        
        # Add subtle border
        border_frame = tk.Frame(action_container, bg="#E2E8F0", height=1)
        border_frame.pack(fill="x", side="top")
        
        # Inner frame with proper padding
        action_frame = ttk.Frame(action_container, padding=(15, 10), style="Card.TFrame")
        action_frame.pack(fill="x")
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=0)

        # --- Left Section: Cancel Button ---
        left_section = ttk.Frame(action_frame, style="Card.TFrame")
        left_section.grid(row=0, column=0, sticky="w", padx=(0, 20))

        # Cancel button
        btn_cancel = ttk.Button(left_section, text="❌ Batal", 
                               command=self._on_cancel, style="Secondary.TButton")
        btn_cancel.pack(side="left")

        # --- Right Section: Finish Button ---
        right_section = ttk.Frame(action_frame, style="Card.TFrame")
        right_section.grid(row=0, column=1, sticky="e")
        
        # Finish button (akan menyimpan dan menutup tab)
        btn_finish = ttk.Button(right_section, text="✅ Selesai", 
                               command=self._on_finish, style="Primary.TButton")
        btn_finish.pack(side="top")
        
        # Helper text - compact
        helper_text = ttk.Label(right_section, 
                               text="Simpan dan pilih opsi finish",
                               style="Caption.TLabel")
        helper_text.pack(side="top", pady=(3, 0))

    def get_disposisi_labels(self):
        """Get selected disposisi labels - compatible with main app"""
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

    def _on_send_email_simple(self):
        """Simplified email sending without complex dialogs"""
        try:
            disposisi_labels = self.get_disposisi_labels_with_abbreviation()  # Use abbreviation for display
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
            
            # Send email directly - use full names for email lookup
            full_disposisi_labels = self.get_disposisi_labels()  # Use full names for email lookup
            self._send_email_to_positions(full_disposisi_labels)
            
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
            
            # Tambahkan informasi disposisi kepada
            try:
                disposisi_labels = self.get_disposisi_labels_with_abbreviation()
            except Exception as e:
                print(f"[WARNING] Error getting disposisi labels: {e}")
                disposisi_labels = []
            
            template_data = {
                'nomor_surat': data.get("no_surat", ""),
                'nama_pengirim': data.get("asal_surat", ""),
                'perihal': data.get("perihal", ""),
                'tanggal': datetime.now().strftime('%d %B %Y'),
                'klasifikasi': klasifikasi,
                'instruksi_list': instruksi_list,
                'disposisi_kepada': disposisi_labels,
                'tahun': datetime.now().year
            }
            
            # Render HTML content
            html_content = render_email_template(template_data)
            subject = f"Disposisi Surat: {data.get('perihal', 'N/A')}"
            
            # Initialize email sender and send
            email_sender = EmailSender()
            
            # Prepare PDF attachment if available
            pdf_attachment_path = None
            if hasattr(self, 'pdf_attachments') and self.pdf_attachments:
                try:
                    import tempfile
                    # Create temporary merged PDF with attachments
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                        temp_pdf_path = temp_pdf.name
                    
                    # Create main form PDF
                    from pdf_output import save_form_to_pdf, merge_pdfs
                    save_form_to_pdf(temp_pdf_path, data)
                    
                    # Merge with attachments
                    pdf_list = [temp_pdf_path] + list(self.pdf_attachments)
                    merged_pdf_path = os.path.join(tempfile.gettempdir(), f"merged_disposisi_{os.getpid()}.pdf")
                    merge_pdfs(pdf_list, merged_pdf_path)
                    
                    pdf_attachment_path = merged_pdf_path
                    
                except Exception as e:
                    print(f"[WARNING] Failed to create PDF attachment: {e}")
                    pdf_attachment_path = None
            
            success, message, details = email_sender.send_disposisi_to_positions(
                positions, 
                subject, 
                html_content,
                pdf_attachment_path
            )
            
            # Show results
            if success:
                success_msg = f"Email berhasil dikirim!\n\n{message}"
                if details.get('failed_lookups'):
                    success_msg += "\n\nCatatan: Beberapa posisi tidak memiliki email yang valid di database."
                messagebox.showinfo("Email Sent", success_msg, parent=self.winfo_toplevel())
            else:
                error_msg = f"Gagal mengirim email:\n{message}"
                if details.get('failed_lookups'):
                    # Gunakan singkatan untuk manager dalam error message
                    abbreviation_map = {
                        "Manager Pemeliharaan": "pml",
                        "Manager Operasional": "ops",
                        "Manager Administrasi": "adm",
                        "Manager Keuangan": "keu"
                    }
                    
                    # Konversi failed lookups ke singkatan untuk display
                    display_failed_lookups = []
                    for lookup in details['failed_lookups']:
                        for full_name, abbrev in abbreviation_map.items():
                            if full_name in lookup:
                                lookup = lookup.replace(full_name, f"Manager {abbrev}")
                                break
                        display_failed_lookups.append(lookup)
                    
                    error_msg += "\n\nPosisi tanpa email:\n" + "\n".join(display_failed_lookups)
                messagebox.showerror("Email Error", error_msg, parent=self.winfo_toplevel())
            
            # Clean up temporary files
            try:
                if 'pdf_attachment_path' in locals() and pdf_attachment_path and os.path.exists(pdf_attachment_path):
                    os.remove(pdf_attachment_path)
                    print(f"[DEBUG] Cleaned up temp PDF: {pdf_attachment_path}")
                if 'temp_pdf_path' in locals() and temp_pdf_path and os.path.exists(temp_pdf_path):
                    os.remove(temp_pdf_path)
                    print(f"[DEBUG] Cleaned up temp PDF: {temp_pdf_path}")
            except Exception as e:
                print(f"[WARNING] Could not remove temporary files: {e}")
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Email Error", f"Terjadi kesalahan: {e}", parent=self.winfo_toplevel())
            
            # Clean up temporary files on error too
            try:
                if 'pdf_attachment_path' in locals() and pdf_attachment_path and os.path.exists(pdf_attachment_path):
                    os.remove(pdf_attachment_path)
                if 'temp_pdf_path' in locals() and temp_pdf_path and os.path.exists(temp_pdf_path):
                    os.remove(temp_pdf_path)
            except Exception as cleanup_error:
                print(f"[WARNING] Could not remove temporary files: {cleanup_error}")

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
            "Manager Pemeliharaan": "manager_pemeliharaan",
            "Manager Operasional": "manager_operasional",
            "Manager Administrasi": "manager_administrasi",
            "Manager Keuangan": "manager_keuangan"
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
                ("Manager Pemeliharaan Instruksi", "Manager Pemeliharaan", "Manager Pemeliharaan Tanggal"),
                ("Manager Operasional Instruksi", "Manager Operasional", "Manager Operasional Tanggal"),
                ("Manager Administrasi Instruksi", "Manager Administrasi", "Manager Administrasi Tanggal"),
                ("Manager Keuangan Instruksi", "Manager Keuangan", "Manager Keuangan Tanggal")
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
        
        # FIXED: Improved date handling for DateEntry widgets - ALLOW NULL VALUES
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
                # Get date value from multiple possible sources
                date_value = data.get(data_key, "")
                if not date_value:
                    # Try alternative key names
                    alt_keys = [
                        widget_key.replace("_", " ").title(),
                        widget_key.upper(),
                        data_key.replace("_", " ").title()
                    ]
                    for alt_key in alt_keys:
                        if alt_key in data and data[alt_key]:
                            date_value = data[alt_key]
                            break
                
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
        
        # Load PDF attachments if available in data log
        if "pdf_attachments" in data and data["pdf_attachments"]:
            try:
                self.pdf_attachments = data["pdf_attachments"]
                self.refresh_pdf_attachments(self.main_frame)
                print(f"[EditTab] Loaded {len(self.pdf_attachments)} PDF attachments from data log.")
            except Exception as e:
                print(f"[EditTab] Warning: Could not load PDF attachments: {e}")
        
        print("[EditTab] Form berhasil diisi dari data log.")

    def _on_save(self):
        import threading
        from disposisi_app.views.components.loading_screen import loading_manager, LoadingMessageBox
        
        def do_save():
            try:
                # Show loading screen
                loading_manager.show_loading(self, "Saving Changes...", True)
                
                print("[EditTab] Mulai proses simpan edit.")
                
                # Collect form data safely using the improved function
                from disposisi_app.views.components.export_utils import collect_form_data_safely
                try:
                    data_baru = collect_form_data_safely(self)
                    
                    # ENHANCED: Ensure date fields are properly collected for EditTab
                    # Handle DateEntry widgets specifically
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
                            try:
                                # Try to get date from DateEntry widget
                                if hasattr(widget, 'get_date'):
                                    date_obj = widget.get_date()
                                    if date_obj:
                                        data_baru[data_key] = date_obj.strftime("%d-%m-%Y")
                                    else:
                                        data_baru[data_key] = ""
                                elif hasattr(widget, 'get'):
                                    # Fallback to regular get() method
                                    value = widget.get()
                                    if value and str(value).strip():
                                        data_baru[data_key] = str(value).strip()
                                    else:
                                        data_baru[data_key] = ""
                                else:
                                    data_baru[data_key] = ""
                            except Exception as e:
                                print(f"[EditTab] Error getting date from {widget_key}: {e}")
                                data_baru[data_key] = ""
                        else:
                            # If widget not found, keep existing value or set empty
                            if data_key not in data_baru:
                                data_baru[data_key] = ""
                    
                    # Add PDF attachments to data
                    if hasattr(self, 'pdf_attachments') and self.pdf_attachments:
                        data_baru["pdf_attachments"] = self.pdf_attachments
                        print(f"[EditTab] Added {len(self.pdf_attachments)} PDF attachments to save data.")
                    
                    print(f"[EditTab] Collected data keys: {list(data_baru.keys())}")
                    print(f"[EditTab] Date values - tgl_surat: {data_baru.get('tgl_surat', 'NOT_FOUND')}")
                    print(f"[EditTab] Date values - tgl_terima: {data_baru.get('tgl_terima', 'NOT_FOUND')}")
                    print(f"[EditTab] Date values - harap_selesai_tgl: {data_baru.get('harap_selesai_tgl', 'NOT_FOUND')}")
                    
                except Exception as e:
                    print(f"[EditTab][ERROR] Error collecting form data: {e}")
                    LoadingMessageBox.showerror("Error", f"Gagal mengumpulkan data form: {e}", parent=self)
                    loading_manager.hide_loading()
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
                    LoadingMessageBox.showerror("Error", "Field 'No. Surat' tidak boleh kosong!", parent=self)
                    loading_manager.hide_loading()
                    return
                
                # Cek duplikasi No. Surat (kecuali data yang sedang diedit)
                try:
                    from google_sheets_connect import get_sheets_service, SHEET_ID
                    service = get_sheets_service()
                    range_name = 'Sheet1!A6:AH'  # Changed to include all columns
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
                                LoadingMessageBox.showerror("Error", f"No. Surat '{no_surat_baru}' sudah ada di data lain!", parent=self)
                                loading_manager.hide_loading()
                                return
                except Exception as e:
                    print(f"[EditTab][WARNING] Tidak bisa cek duplikasi No. Surat: {e}")
                    # Continue even if duplicate check fails to avoid data loss

                # Progress simulation
                for i in range(1, 101):
                    import time; time.sleep(0.01)
                    loading_manager.update_progress(i)
                
                print("[EditTab] Memanggil update_log_entry...")
                
                # Use the fixed update_log_entry function
                from edit_logic import update_log_entry
                success = update_log_entry(self.data_log, data_baru)
                
                if success:
                    print("[EditTab] Update berhasil.")
                    LoadingMessageBox.showinfo("Sukses", "Data berhasil diupdate.", parent=self)
                    
                    # Call callback if available
                    if self.on_save_callback:
                        print("[EditTab] Memanggil on_save_callback...")
                        self.on_save_callback()
                else:
                    print("[EditTab] Update gagal.")
                    LoadingMessageBox.showerror("Error", "Gagal update data ke Google Sheets.", parent=self)
                
            except Exception as e:
                import traceback; traceback.print_exc()
                print(f"[EditTab][ERROR] Gagal update data: {e}")
                LoadingMessageBox.showerror("Error", f"Gagal update data: {e}", parent=self)
            finally:
                loading_manager.hide_loading()
                    
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
        from disposisi_app.views.components.loading_screen import loading_manager, LoadingMessageBox
        
        def do_export():
            try:
                # Show loading screen
                loading_manager.show_loading(self, "Exporting to PDF...", True)
                
                filepath = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF Documents", "*.pdf"), ("All Files", "*.*")],
                    title="Export Edit ke PDF"
                )
                if not filepath:
                    loading_manager.hide_loading()
                    return
                
                # Progress simulation
                for i in range(1, 101):
                    import time; time.sleep(0.01)
                    loading_manager.update_progress(i)
                
                # Ambil data dari form menggunakan collect_form_data_safely
                data_baru = collect_form_data_safely(self)
                
                # Simpan PDF ke file sementara
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    temp_pdf_path = temp_pdf.name
                save_form_to_pdf(temp_pdf_path, data_baru)
                
                # Gabungkan PDF disposisi + lampiran jika ada
                pdf_list = [temp_pdf_path]
                if hasattr(self, 'pdf_attachments') and self.pdf_attachments:
                    pdf_list.extend(list(self.pdf_attachments))
                
                merge_pdfs(pdf_list, filepath)
                os.remove(temp_pdf_path)
                
                # Upload ke Google Sheets juga
                try:
                    update_log_entry(self.data_log, data_baru)
                except Exception as e:
                    import traceback; traceback.print_exc()
                    LoadingMessageBox.showerror("Google Sheets", f"Gagal upload ke Google Sheets: {e}", parent=self)
                LoadingMessageBox.showinfo("Sukses", f"Data edit berhasil diekspor ke PDF:\n{filepath}", parent=self)
            except Exception as e:
                traceback.print_exc()
                LoadingMessageBox.showerror("Export PDF", f"Gagal ekspor ke PDF: {e}", parent=self)
            finally:
                loading_manager.hide_loading()
        threading.Thread(target=do_export, daemon=True).start() 

    def _on_tambah_baris(self):
        self.instruksi_table.add_row()

    def _on_hapus_baris(self):
        self.instruksi_table.remove_selected_rows() 

    def _on_kosongkan_baris(self):
        self.instruksi_table.kosongkan_semua_baris()

    def _on_finish_simple(self):
        """SEDERHANAKAN: Simpan perubahan dan tutup tab edit"""
        import threading
        from disposisi_app.views.components.loading_screen import loading_manager, LoadingMessageBox
        
        def do_finish():
            try:
                # Show loading screen
                loading_manager.show_loading(self, "Menyimpan perubahan...", True)
                
                # Progress simulation
                for i in range(1, 101):
                    import time; time.sleep(0.01)
                    loading_manager.update_progress(i)
                
                # Ambil data dari form
                data_baru = collect_form_data_safely(self)
                
                # Update ke Google Sheets
                from edit_logic import update_log_entry
                success = update_log_entry(self.data_log, data_baru)
                
                if success:
                    # Call callback untuk refresh data dan tutup tab
                    if self.on_save_callback:
                        self.on_save_callback()
                    
                    LoadingMessageBox.showinfo("Sukses", "Perubahan berhasil disimpan dan tab ditutup.", parent=self)
                else:
                    LoadingMessageBox.showerror("Error", "Gagal menyimpan perubahan ke Google Sheets.", parent=self)
                
            except Exception as e:
                import traceback
                traceback.print_exc()
                LoadingMessageBox.showerror("Error", f"Gagal menyimpan perubahan: {e}", parent=self)
            finally:
                loading_manager.hide_loading()
        
        # Run in separate thread to avoid blocking UI
        threading.Thread(target=do_finish, daemon=True).start()

    def _on_finish(self):
        """Show finish dialog for completing the edit process"""
        try:
            from disposisi_app.views.components.finish_dialog import FinishDialog
            
            # Get disposisi labels
            disposisi_labels = self.get_disposisi_labels()
            
            # Create callbacks for finish dialog
            dialog_callbacks = {
                "save_pdf": self._on_export_pdf,
                "save_sheet": self._on_save,
                "send_email": self._send_email_to_positions
            }
            
            # Create and show finish dialog
            dialog = FinishDialog(self, disposisi_labels, dialog_callbacks)
            self.wait_window(dialog)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Gagal membuka dialog selesai: {e}")
def safe_get_widget_value(widget, widget_name="unknown"):
    """Safely get value from any widget type - reusable function"""
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
        import traceback
        traceback.print_exc()
        return ""

def safe_get_text_widget_value(widget, widget_name="unknown"):
    """Safely get value from Text widgets"""
    try:
        if widget is None:
            print(f"[WARNING] Text widget {widget_name} is None")
            return ""
        
        if hasattr(widget, 'get'):
            return widget.get("1.0", "end").strip()
        else:
            print(f"[WARNING] Text widget {widget_name} doesn't have get method")
            return ""
            
    except Exception as e:
        print(f"[ERROR] Error getting text value from {widget_name}: {e}")
        import traceback
        traceback.print_exc()
        return ""

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
                    data[field] = safe_get_text_widget_value(
                        self.form_input_widgets[field], 
                        field
                    )
                else:
                    data[field] = ""
            
            # Date widget for tgl_surat
            if "tgl_surat" in self.form_input_widgets:
                data["tgl_surat"] = safe_get_widget_value(
                    self.form_input_widgets["tgl_surat"], 
                    "tgl_surat"
                )
            else:
                data["tgl_surat"] = ""
        
        # Handle special date entries
        if hasattr(self, 'tgl_terima_entry'):
            data["tgl_terima"] = safe_get_widget_value(
                self.tgl_terima_entry, 
                "tgl_terima_entry"
            )
        else:
            data["tgl_terima"] = ""
        
        if hasattr(self, 'harap_selesai_tgl_entry'):
            data["harap_selesai_tgl"] = safe_get_widget_value(
                self.harap_selesai_tgl_entry, 
                "harap_selesai_tgl_entry"
            )
        else:
            data["harap_selesai_tgl"] = ""
        
        # Ensure required fields have default values
        data["indeks"] = data.get("indeks", "")
        data["rahasia"] = data.get("rahasia", 0)
        data["penting"] = data.get("penting", 0)
        data["segera"] = data.get("segera", 0)
        
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

def save_to_pdf(self):
    import threading
    from tkinter import filedialog, messagebox
    import time, traceback, os, tempfile
    from pdf_output import save_form_to_pdf, merge_pdfs
    from sheet_logic import upload_to_sheet
    from disposisi_app.views.components.loading_screen import LoadingScreen
    
    def do_save():
        loading = None
        try:
            loading = LoadingScreen(self)
            for i in range(1, 101):
                time.sleep(0.02)
                loading.update_progress(i)
            
            filepath = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Documents", "*.pdf"), ("All Files", "*.*")]
            )
            if not filepath:
                if loading and loading.winfo_exists():
                    loading.destroy()
                return
            
            # Safely collect form data
            data = collect_form_data_safely(self)
            
            # Validate required fields
            if not str(data.get("no_surat", "")).strip():
                messagebox.showerror("Validasi", "No. Surat tidak boleh kosong!")
                if loading and loading.winfo_exists():
                    loading.destroy()
                return
            
            # Check uniqueness
            try:
                from disposisi_app.views.components.validation import is_no_surat_unique
                from google_sheets_connect import get_sheets_service, SHEET_ID
                if not is_no_surat_unique(data.get("no_surat", ""), get_sheets_service, SHEET_ID):
                    messagebox.showerror("Validasi", "No. Surat sudah ada di sheet, harus unik!")
                    if loading and loading.winfo_exists():
                        loading.destroy()
                    return
            except Exception as e:
                print(f"[WARNING] Error checking uniqueness: {e}")
                # Continue anyway if validation fails
            
            # Create PDF
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf_path = temp_pdf.name
            
            save_form_to_pdf(temp_pdf_path, data)
            
            # Merge with attachments if they exist
            pdf_list = [temp_pdf_path]
            if hasattr(self, 'pdf_attachments') and self.pdf_attachments:
                pdf_list.extend(list(self.pdf_attachments))
            
            merge_pdfs(pdf_list, filepath)
            
            # Clean up temp file
            try:
                os.remove(temp_pdf_path)
            except Exception as e:
                print(f"[WARNING] Error removing temp file: {e}")
            
            # Upload to Google Sheets
            try:
                upload_to_sheet(self, call_from_pdf=True, data_override=data)
            except Exception as e:
                traceback.print_exc()
                messagebox.showerror("Google Sheets", f"Gagal upload ke Google Sheets: {e}")
            
            # Clean up and show success
            if loading and loading.winfo_exists():
                loading.destroy()
            messagebox.showinfo("Sukses", f"Formulir disposisi dan lampiran telah disimpan di:\n{filepath}")
            
            # Reset form state
            self.edit_mode = False
            if hasattr(self, 'pdf_attachments'):
                self.pdf_attachments = []
            if hasattr(self, 'refresh_pdf_attachments'):
                try:
                    self.refresh_pdf_attachments(self._form_main_frame)
                except Exception as e:
                    print(f"[WARNING] Error refreshing PDF attachments: {e}")
            
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"Terjadi error: {e}")
            if loading and loading.winfo_exists():
                loading.destroy()
    
    threading.Thread(target=do_save, daemon=True).start()

def save_to_sheet(self):
    import threading
    from tkinter import messagebox
    import time, traceback
    from sheet_logic import upload_to_sheet
    from disposisi_app.views.components.loading_screen import LoadingScreen
    
    def do_save():
        loading = None
        try:
            loading = LoadingScreen(self)
            for i in range(1, 101):
                time.sleep(0.01)
                loading.update_progress(i)
            
            # Safely collect form data
            data = collect_form_data_safely(self)
            
            # Validate required fields
            if not str(data.get("no_surat", "")).strip():
                messagebox.showerror("Validasi", "No. Surat tidak boleh kosong!")
                if loading and loading.winfo_exists():
                    loading.destroy()
                return
            
            # Check uniqueness
            try:
                from disposisi_app.views.components.validation import is_no_surat_unique
                from google_sheets_connect import get_sheets_service, SHEET_ID
                if not is_no_surat_unique(data.get("no_surat", ""), get_sheets_service, SHEET_ID):
                    messagebox.showerror("Validasi", "No. Surat sudah ada di sheet, harus unik!")
                    if loading and loading.winfo_exists():
                        loading.destroy()
                    return
            except Exception as e:
                print(f"[WARNING] Error checking uniqueness: {e}")
                messagebox.showerror("Error", f"Error validasi: {e}")
                if loading and loading.winfo_exists():
                    loading.destroy()
                return
            
            # Upload to sheet
            upload_to_sheet(self, call_from_pdf=False)
            
            if loading and loading.winfo_exists():
                loading.destroy()
                
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"Terjadi error: {e}")
            if loading and loading.winfo_exists():
                loading.destroy()
    
    threading.Thread(target=do_save, daemon=True).start()

def safe_get_filter_value(widget, widget_name="unknown", default=""):
    """Safely get value from filter widgets"""
    try:
        if widget is None:
            return default
        
        if hasattr(widget, 'get'):
            return str(widget.get())
        else:
            return default
            
    except Exception as e:
        print(f"[WARNING] Error getting filter value from {widget_name}: {e}")
        return default

def export_excel_advanced(self, filepath):
    import openpyxl
    from openpyxl.styles import Alignment, Font, Border, Side
    from openpyxl.utils import get_column_letter
    from tkinter import messagebox
    
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        if ws is None:
            ws = wb.worksheets[0]
        if ws is None:
            messagebox.showerror("Export Excel", "Worksheet tidak dapat dibuat.")
            return
        
        ws.title = "Log Disposisi"
        
        # Header setup
        ws.merge_cells('A1:AB1')
        ws['A1'] = 'LAPORAN DISPOSISI'
        ws['A1'].font = Font(bold=True, size=14)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        import datetime
        today_str = datetime.datetime.now().strftime('%d-%m-%Y')
        ws.merge_cells('A2:AB2')
        ws['A2'] = f'terakhir di update: {today_str}'
        ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
        
        # Filter information - safely get values
        filter_info = []
        
        disposisi_val = safe_get_filter_value(
            getattr(self, 'filter_val', None), 
            'filter_val'
        )
        if disposisi_val:
            filter_info.append(f"Disposisi ke: {disposisi_val}")
        
        bulan = safe_get_filter_value(
            getattr(self, 'month_var', None), 
            'month_var'
        )
        if bulan:
            filter_info.append(f"Bulan: {bulan}")
        
        tahun = safe_get_filter_value(
            getattr(self, 'year_var', None), 
            'year_var'
        )
        if tahun:
            filter_info.append(f"Tahun: {tahun}")
        
        search_col = safe_get_filter_value(
            getattr(self, 'search_col', None), 
            'search_col'
        )
        search_val = safe_get_filter_value(
            getattr(self, 'search_entry', None), 
            'search_entry'
        )
        if search_col and search_val:
            filter_info.append(f"Pencarian: {search_col} = {search_val}")
        
        ws.merge_cells('A3:AB3')
        ws['A3'] = ' | '.join(filter_info) if filter_info else 'Semua Data'
        ws['A3'].alignment = Alignment(horizontal='center', vertical='center')
        
        # Headers
        header_row3 = [
            "No. Agenda", "No. Surat", "Tgl. Surat", "Perihal", "Asal Surat", "Ditujukan", 
            "Klasifikasi", "Disposisi kepada", "Untuk Di :", "Selesai Tgl.", "Kode Klasifikasi", 
            "Tgl. Penerimaan", "Indeks", "Bicarakan dengan", "Teruskan kepada", "Harap Selesai Tanggal",
            "Direktur Utama", "", "Direktur Keuangan", "", "Direktur Teknik", "", 
            "GM Keuangan & Administrasi", "", "GM Operasional & Pemeliharaan", "", "Manager", ""
        ]
        ws.append(header_row3)
        
        header_row4 = ["" for _ in range(28)]
        positions = [16, 18, 20, 22, 24, 26]  # Instruction columns
        for pos in positions:
            header_row4[pos] = "Instruksi"
            header_row4[pos + 1] = "Tanggal"
        ws.append(header_row4)
        
        # Define enhanced header for data extraction
        ENHANCED_HEADER = [
            "No. Agenda", "No. Surat", "Tgl. Surat", "Perihal", "Asal Surat", "Ditujukan", 
            "Klasifikasi", "Disposisi kepada", "Untuk Di :", "Selesai Tgl.", "Kode Klasifikasi", 
            "Tgl. Penerimaan", "Indeks", "Bicarakan dengan", "Teruskan kepada", "Harap Selesai Tanggal",
            "Direktur Utama Instruksi", "Direktur Utama Tanggal", "Direktur Keuangan Instruksi", "Direktur Keuangan Tanggal",
            "Direktur Teknik Instruksi", "Direktur Teknik Tanggal", "GM Keuangan & Administrasi Instruksi", "GM Keuangan & Administrasi Tanggal",
            "GM Operasional & Pemeliharaan Instruksi", "GM Operasional & Pemeliharaan Tanggal", "Manager Instruksi", "Manager Tanggal"
        ]
        
        # Add data rows safely
        if hasattr(self, 'filtered_data') and self.filtered_data:
            for row in self.filtered_data:
                try:
                    row_data = []
                    for col in ENHANCED_HEADER:
                        cell_value = row.get(col, "")
                        # Handle None values
                        if cell_value is None:
                            cell_value = ""
                        row_data.append(str(cell_value))
                    ws.append(row_data)
                except Exception as e:
                    print(f"[WARNING] Error adding row to Excel: {e}")
                    # Add empty row to maintain structure
                    ws.append(["" for _ in range(len(ENHANCED_HEADER))])
        else:
            # Add a placeholder row if no data
            ws.append(["Tidak ada data" for _ in range(len(ENHANCED_HEADER))])
        
        # Apply formatting
        thin = Side(border_style="thin", color="000000")
        try:
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=28):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
                    cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        except Exception as e:
            print(f"[WARNING] Error applying cell formatting: {e}")
        
        # Adjust column widths
        try:
            for i, col in enumerate(ENHANCED_HEADER, 1):
                try:
                    max_length = 10  # Default width
                    for r in range(1, ws.max_row + 1):
                        cell_value = ws.cell(row=r, column=i).value
                        if cell_value is not None:
                            length = len(str(cell_value))
                            if length > max_length:
                                max_length = length
                    
                    ws.column_dimensions[get_column_letter(i)].width = min(max_length + 2, 40)
                except Exception as e:
                    print(f"[WARNING] Error adjusting column {i} width: {e}")
                    ws.column_dimensions[get_column_letter(i)].width = 15
        except Exception as e:
            print(f"[WARNING] Error in column width adjustment: {e}")
        
        # Save file
        wb.save(filepath)
        messagebox.showinfo("Export Excel", f"Data berhasil diekspor ke {filepath}")
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("Export Excel", f"Gagal ekspor: {e}")
def send_email_with_disposisi(self, recipients):
    """
    Generates the disposition PDF, attaches it, and sends it to the specified recipients.
    """
    # Import the fixed EmailSender
    try:
        from email_sender.send_email import EmailSender
        from email_sender.template_handler import render_email_template
    except ImportError:
        messagebox.showerror("Import Error", 
                             "Email sender module not found. Please ensure email_sender package is properly installed.")
        return

    self.update_status("Preparing email...")
    
    # Collect form data
    from disposisi_app.views.components.export_utils import collect_form_data_safely
    data = collect_form_data_safely(self)
    
    if not data.get("no_surat", "").strip():
        messagebox.showerror("Validation Error", "No. Surat tidak boleh kosong untuk mengirim email.")
        return

    # Initialize email sender
    email_sender = EmailSender()
    
    # Create PDF attachment
    try:
        import tempfile
        import os
        from pdf_output import save_form_to_pdf, merge_pdfs
        
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
        import traceback
        traceback.print_exc()
        return
    
    # Prepare email data
    from datetime import datetime
    
    # Get klasifikasi
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
    
    # Send email using the position-based method
    self.update_status(f"Mengirim email ke {len(recipients)} penerima...")
    
    success, message, details = email_sender.send_disposisi_to_positions(
        recipients, 
        subject, 
        html_content
    )
    
    self.update_status("Email berhasil dikirim!" if success else "Gagal mengirim email.")
    
    # Clean up temporary files
    try:
        if 'temp_pdf_path' in locals() and os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if 'final_pdf_path' in locals() and final_pdf_path != temp_pdf_path and os.path.exists(final_pdf_path):
            os.remove(final_pdf_path)
    except:
        pass
    
    # Show detailed results
    if success:
        success_msg = f"Email berhasil dikirim!\n\n{message}"
        if details.get('failed_lookups'):
            success_msg += "\n\nCatatan: Beberapa posisi tidak memiliki email yang valid di database."
        messagebox.showinfo("Email Sent", success_msg)
    else:
        error_msg = f"Gagal mengirim email:\n{message}"
        if details.get('failed_lookups'):
            error_msg += "\n\nPosisi tanpa email:\n" + "\n".join(details['failed_lookups'])
        messagebox.showerror("Email Error", error_msg)
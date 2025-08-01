import tkinter as tk
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
import logging
from typing import Dict, Any
import PyPDF2

logging.basicConfig(
    filename='app.log',
    level=logging.ERROR,
    format='%(asctime)s %(levelname)s:%(name)s:%(message)s'
)
import traceback

def convert_to_boolean(value):
    """Convert various input types to boolean consistently."""
    if isinstance(value, bool):
        return value
    elif isinstance(value, int):
        return value != 0
    elif isinstance(value, str):
        return value.lower() in ('true', '1', 'yes', 'y', 'on')
    else:
        try:
            str_val = str(value).lower()
            return str_val in ('true', '1', 'yes', 'y', 'on') and str_val != '0'
        except:
            return False

def save_form_to_pdf(filepath: str, data: Dict[str, Any]) -> None:
    try:
        # FIX: Check and handle file permissions before creating PDF
        import os
        import tempfile
        
        # Try to create a temporary file first to test permissions
        try:
            temp_dir = os.path.dirname(filepath)
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir, exist_ok=True)
            
            # Test write permissions by creating a temporary file
            temp_file = os.path.join(temp_dir, f"temp_{os.getpid()}.tmp")
            with open(temp_file, 'w') as f:
                f.write("test")
            os.remove(temp_file)
            
        except (PermissionError, OSError) as e:
            # If permission denied, use user's Documents folder as fallback
            import os.path
            fallback_dir = os.path.expanduser("~/Documents")
            filename = os.path.basename(filepath)
            filepath = os.path.join(fallback_dir, filename)
        
        # Check if file already exists and is open
        if os.path.exists(filepath):
            try:
                # Try to open file in append mode to check if it's locked
                with open(filepath, 'a'):
                    pass
            except PermissionError:
                # File is open, create with timestamp suffix
                import time
                timestamp = int(time.time())
                name, ext = os.path.splitext(filepath)
                filepath = f"{name}_{timestamp}{ext}"
        
        # FIX: Add default values for missing required fields
        default_values = {
            'tgl_terima': '',
            'kode_klasifikasi': '',
            'indeks': '',
            'harap_selesai_tgl': '',
            'rahasia': 0,
            'penting': 0,
            'segera': 0,
            'no_agenda': '',
            'no_surat': '',
            'tgl_surat': '',
            'perihal': '',
            'asal_surat': '',
            'ditujukan': ''
        }
        
        # Apply default values for missing fields
        for field, default_value in default_values.items():
            if field not in data or data[field] is None:
                data[field] = default_value
        
        c = canvas.Canvas(filepath, pagesize=A4)
        width, height = A4
        VERTICAL_SPACING = 0.7 * cm  
        LINE_HEIGHT = 0.6 * cm      
        CHECKBOX_SIZE = 0.4 * cm    
        MARGIN_LEFT = 1.0 * cm    
        MARGIN_RIGHT = 1.0 * cm     
        FONT_SIZE = 11  # Changed from variable to 11
        
        def draw_checkbox(x, y, text, checked, bold=False):
            box_size = CHECKBOX_SIZE
            box_y = y - box_size/2
            c.rect(x, box_y, box_size, box_size)
            
            # Use the helper function for consistent boolean conversion
            is_checked = convert_to_boolean(checked)
            
            # Draw checkmark if checked
            if is_checked:
                c.setFont("Helvetica-Bold", 11)
                c.drawString(x + box_size/2 - c.stringWidth("✓", "Helvetica-Bold", 11)/2,
                             box_y + box_size/2 - 11/2.5, "✓")
            
            c.setFont("Helvetica-Bold" if bold else "Helvetica", 11)
            label_y = y - 11/2.5/2
            c.drawString(x + box_size + 0.3*cm, label_y, text)
        
        def draw_field_flex(x, y, label, value, colon_x, value_width=8*cm, font_size=11):
            c.setFont("Helvetica", font_size)
            c.drawString(x, y, label)
            c.drawString(colon_x, y, ":")
            value_x = colon_x + 0.3*cm
            max_width = value_width
            lines = wrap_text(str(value), max_width, font_size=font_size) if value else []
            for i, line in enumerate(lines):
                c.drawString(value_x, y - (i * 0.6*cm), line)
            return y - (len(lines) if lines else 1) * 0.6*cm - 0.2*cm
        
        def draw_field_vertical_flex(x, y, label, value, colon_x, value_width=8*cm, font_size=11):
            c.setFont("Helvetica", font_size)
            c.drawString(x, y, label)
            c.drawString(colon_x, y, ":")
            value_x = x
            value_y = y - 0.7*cm
            max_width = value_width
            lines = wrap_text(str(value), max_width, font_size=font_size) if value else []
            for i, line in enumerate(lines):
                c.drawString(value_x, value_y - (i * 0.6*cm), line)
            return value_y - (len(lines) if lines else 1) * 0.6*cm - 0.2*cm
        
        def format_tanggal(val):
            import datetime
            if not val:
                return ''
            try:
                if len(val) > 10:
                    dt = datetime.datetime.strptime(val[:19], '%Y-%m-%d %H:%M:%S')
                    return dt.strftime('%d-%m-%Y')
                else:
                    dt = datetime.datetime.strptime(val, '%Y-%m-%d')
                    return dt.strftime('%d-%m-%Y')
            except Exception:
                return val
        
        def wrap_text(text, max_width, font_size=11):
            c.setFont("Helvetica", font_size)
            lines = []
            current_line = ""
            words = text.split(' ')
            for word in words:
                test_line = current_line + word + ' '
                if c.stringWidth(test_line) <= max_width:
                    current_line = test_line
                else:
                    if current_line:
                        lines.append(current_line.strip())
                    current_line = word + ' '
            if current_line:
                lines.append(current_line.strip())
            return lines
        
        # Header - DISPOSISI and Logo
        x_disposisi = MARGIN_LEFT + 12*cm
        y_disposisi = height - 2.5*cm
        c.setFont("Helvetica-Bold", 24)
        c.drawString(x_disposisi, y_disposisi, "DISPOSISI")
        
        # Try to add logo
        try:
            kop_path = "kop.jpeg"
            kop_width_px = 170
            kop_height_px = 100
            px_to_cm = 2.54 / 96
            kop_width_cm = kop_width_px * px_to_cm
            kop_height_cm = kop_height_px * px_to_cm
            kop_width_pt = kop_width_cm * cm
            kop_height_pt = kop_height_cm * cm
            kop_x = MARGIN_LEFT
            disposisi_center_y = y_disposisi + 0.5*cm
            kop_y = disposisi_center_y - kop_height_pt/2 + kop_height_pt
            from reportlab.lib.utils import ImageReader
            c.drawImage(kop_path, kop_x, kop_y - kop_height_pt, kop_width_pt, kop_height_pt, preserveAspectRatio=False, mask='auto')
        except Exception as e:
            pass
        
        # Right side classification checkboxes
        y_klasifikasi_start = y_disposisi - 1.5*cm
        x_klasifikasi = x_disposisi
        draw_checkbox(x_klasifikasi, y_klasifikasi_start, "RAHASIA", data.get("rahasia", 0), bold=True)
        draw_checkbox(x_klasifikasi, y_klasifikasi_start - LINE_HEIGHT, "PENTING", data.get("penting", 0), bold=True)
        draw_checkbox(x_klasifikasi, y_klasifikasi_start - 2*LINE_HEIGHT, "SEGERA", data.get("segera", 0), bold=True)
        
        # Right side classification fields
        klasifikasi_labels = ["Tanggal Penerimaan", "Kode / Klasifikasi", "Indeks"]
        max_label_width = max([c.stringWidth(label + " ", "Helvetica", 11) for label in klasifikasi_labels])
        colon_x_klasifikasi = x_klasifikasi + max_label_width + 0.2*cm
        y_field_klasifikasi = y_klasifikasi_start - 3.2*LINE_HEIGHT
        klasifikasi_width = 8*cm
        y_field_klasifikasi = draw_field_vertical_flex(x_klasifikasi, y_field_klasifikasi, "Tanggal Penerimaan", format_tanggal(data.get("tgl_terima", "")), colon_x_klasifikasi, klasifikasi_width)
        y_field_klasifikasi = draw_field_vertical_flex(x_klasifikasi, y_field_klasifikasi, "Kode / Klasifikasi", data.get("kode_klasifikasi", ""), colon_x_klasifikasi, klasifikasi_width)
        y_field_klasifikasi = draw_field_vertical_flex(x_klasifikasi, y_field_klasifikasi, "Indeks", data.get("indeks", ""), colon_x_klasifikasi, klasifikasi_width)
        
        # Left side document details
        y_checkbox_segera_top = y_klasifikasi_start - 2*LINE_HEIGHT + 0.2*cm  
        y_start = y_checkbox_segera_top
        x_left = MARGIN_LEFT
        label_no_agenda = "No. Agenda"
        colon_x_left = x_left + c.stringWidth(label_no_agenda + " ", "Helvetica", 11)
        y_detil = y_checkbox_segera_top
        y_detil = draw_field_flex(x_left, y_detil, label_no_agenda, data.get("no_agenda", ""), colon_x_left)
        y_detil = draw_field_flex(x_left, y_detil, "No. Surat", data.get("no_surat", ""), colon_x_left)
        y_detil = draw_field_flex(x_left, y_detil, "Tgl. Surat", format_tanggal(data.get("tgl_surat", "")), colon_x_left)
        y_detil = draw_field_flex(x_left, y_detil, "Perihal", data.get("perihal", ""), colon_x_left)
        y_detil = draw_field_flex(x_left, y_detil, "Asal Surat", data.get("asal_surat", ""), colon_x_left)
        y_detil = draw_field_flex(x_left, y_detil, "Ditujukan", data.get("ditujukan", ""), colon_x_left)
        
        # Disposisi Kepada section
        y_middle = y_detil - 0.32*cm
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_left, y_middle, "Disposisi Kepada :")
        y_disp = y_middle - VERTICAL_SPACING
        baris_y = [y_disp]
        for i in range(1, 4):
            baris_y.append(baris_y[-1] - LINE_HEIGHT)
        
        # Direktur Utama
        draw_checkbox(x_left, baris_y[0], "", data.get("dir_utama", 0))
        c.setFont("Helvetica", 11)
        c.drawString(x_left + CHECKBOX_SIZE + 0.3*cm, baris_y[0] - 11/2.5/2, "Direktur Utama")
        
        # FIXED: Direktur Keuangan / Direktur Teknik with proper strikethrough
        keu = data.get("dir_keu", 0) 
        teknik = data.get("dir_teknik", 0)
        
        # Convert to boolean properly
        keu_checked = convert_to_boolean(keu)
        teknik_checked = convert_to_boolean(teknik)
        
        # Main checkbox is checked if either is selected
        keu_teknik_checked = keu_checked or teknik_checked
        draw_checkbox(x_left, baris_y[1], "", keu_teknik_checked)
        
        label_keu = "Direktur Keuangan"
        label_teknik = "Direktur Teknik"
        label_gabungan = f"{label_keu} / {label_teknik}"
        x_label = x_left + CHECKBOX_SIZE + 0.3*cm
        c.setFont("Helvetica", 11)
        c.drawString(x_label, baris_y[1] - 11/2.5/2, label_gabungan)
        
        # Calculate positions for strikethrough
        width_keu = c.stringWidth(label_keu, "Helvetica", 11)
        width_slash = c.stringWidth(" / ", "Helvetica", 11)
        width_teknik = c.stringWidth(label_teknik, "Helvetica", 11)
        y_strike = baris_y[1] - 11/2.5/2 + 0.2*cm
        
        # Apply strikethrough logic - only strike through the unselected option
        if keu_checked and not teknik_checked:
            # Strike through "Direktur Teknik"
            c.saveState()
            c.setLineWidth(1)
            x_teknik = x_label + width_keu + width_slash
            c.line(x_teknik, y_strike, x_teknik + width_teknik, y_strike)
            c.restoreState()
        elif teknik_checked and not keu_checked:
            # Strike through "Direktur Keuangan"
            c.saveState()
            c.setLineWidth(1)
            x_keu = x_label
            c.line(x_keu, y_strike, x_keu + width_keu, y_strike)
            c.restoreState()
        
        # FIXED: GM Keuangan & Administrasi / GM Operasional & Pemeliharaan with proper strikethrough
        gm_keu = data.get("gm_keu", 0)
        gm_ops = data.get("gm_ops", 0)
        
        # Convert to boolean properly
        gm_keu_checked = convert_to_boolean(gm_keu)
        gm_ops_checked = convert_to_boolean(gm_ops)
        
        # Main checkbox is checked if either is selected
        gm_gabungan_checked = gm_keu_checked or gm_ops_checked
        draw_checkbox(x_left, baris_y[2], "", gm_gabungan_checked)
        
        label_gm_keu = "GM Keuangan & Administrasi"
        label_gm_ops = "GM Operasional & Pemeliharaan"
        label_gm_gabungan = f"{label_gm_keu} / {label_gm_ops}"
        x_label_gm = x_left + CHECKBOX_SIZE + 0.3*cm
        c.setFont("Helvetica", 11)
        c.drawString(x_label_gm, baris_y[2] - 11/2.5/2, label_gm_gabungan)
        
        # Calculate positions for strikethrough
        width_gm_keu = c.stringWidth(label_gm_keu, "Helvetica", 11)
        width_gm_slash = c.stringWidth(" / ", "Helvetica", 11)
        width_gm_ops = c.stringWidth(label_gm_ops, "Helvetica", 11)
        y_strike_gm = baris_y[2] - 11/2.5/2 + 0.2*cm
        
        # Apply strikethrough logic - only strike through the unselected option
        if gm_keu_checked and not gm_ops_checked:
            # Strike through "GM Operasional & Pemeliharaan"
            c.saveState()
            c.setLineWidth(1)
            x_gm_ops = x_label_gm + width_gm_keu + width_gm_slash
            c.line(x_gm_ops, y_strike_gm, x_gm_ops + width_gm_ops, y_strike_gm)
            c.restoreState()
        elif gm_ops_checked and not gm_keu_checked:
            # Strike through "GM Keuangan & Administrasi"
            c.saveState()
            c.setLineWidth(1)
            x_gm_keu = x_label_gm
            c.line(x_gm_keu, y_strike_gm, x_gm_keu + width_gm_keu, y_strike_gm)
            c.restoreState()
        
        # Manager
        draw_checkbox(x_left, baris_y[3], "", data.get("manager", 0))
        c.setFont("Helvetica", 11)
        c.drawString(x_left + CHECKBOX_SIZE + 0.3*cm, baris_y[3] - 11/2.5/2, "Manager")
        
        # Untuk di section
        y_untuk = y_disp - 4.0*LINE_HEIGHT - 0.2*cm  # Reduced spacing before "Untuk di" section
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_left, y_untuk, "Untuk di :")
        c.setFont("Helvetica", 11)
        y_untuk_item = y_untuk - VERTICAL_SPACING  # Changed to use same spacing as "Disposisi Kepada"
        draw_checkbox(x_left, y_untuk_item, "Ketahui & File", data.get("ketahui_file", 0))
        draw_checkbox(x_left, y_untuk_item - LINE_HEIGHT, "Proses Selesai", data.get("proses_selesai", 0))
        draw_checkbox(x_left, y_untuk_item - 2*LINE_HEIGHT, "Teliti & Pendapat", data.get("teliti_pendapat", 0))
        draw_checkbox(x_left, y_untuk_item - 3*LINE_HEIGHT, "Buatkan Resume", data.get("buatkan_resume", 0))
        draw_checkbox(x_left, y_untuk_item - 4*LINE_HEIGHT, "Edarkan", data.get("edarkan", 0))
        draw_checkbox(x_left, y_untuk_item - 5*LINE_HEIGHT, "Sesuai Disposisi", data.get("sesuai_disposisi", 0))
        draw_checkbox(x_left, y_untuk_item - 6*LINE_HEIGHT, "Bicarakan dengan Saya", data.get("bicarakan_saya", 0))
        
        # Additional fields
        y_field_tambahan = y_untuk_item - 7*LINE_HEIGHT - 0.32*cm
        
        # Bicarakan dengan field
        bicarakan_label = "Bicarakan dengan :"
        bicarakan_isi = str(data.get('bicarakan_dengan', ''))
        bicarakan_checked = bool(bicarakan_isi.strip())
        draw_checkbox(x_left, y_field_tambahan, bicarakan_label, bicarakan_checked)
        
        # Position text below label with 12pt (0.42cm) spacing
        value_x_bicarakan = x_left + CHECKBOX_SIZE + 0.5*cm  # Indented from checkbox
        value_y_bicarakan = y_field_tambahan - 0.6*cm  # 12pt spacing below label
        # Limit to 1 line maximum
        value_width_bicarakan = 5*cm  # Reduced width for single line
        bicarakan_lines = wrap_text(bicarakan_isi, value_width_bicarakan, font_size=11) if bicarakan_isi else []
        bicarakan_lines = bicarakan_lines[:1]  # Limit to 1 line maximum
        
        # Draw all lines with proper indentation
        for i, line in enumerate(bicarakan_lines):
            c.drawString(value_x_bicarakan, value_y_bicarakan - (i * 0.4*cm), line)
        
        # Spacing between sections
        line_height = 0.4*cm
        lines_count = max(1, len(bicarakan_lines))
        EXTRA_SPACING = 0.2 * cm  # Spacing between sections
        y_field_tambahan2 = value_y_bicarakan - (lines_count * line_height) - EXTRA_SPACING
        
        # Teruskan Kepada field
        teruskan_label = "Teruskan Kepada :"
        teruskan_isi = str(data.get('teruskan_kepada', ''))
        teruskan_checked = bool(teruskan_isi.strip())
        draw_checkbox(x_left, y_field_tambahan2, teruskan_label, teruskan_checked)
        
        # Position text below label with 12pt (0.42cm) spacing
        value_x_teruskan = x_left + CHECKBOX_SIZE + 0.5*cm  # Indented from checkbox
        value_y_teruskan = y_field_tambahan2 - 0.6*cm  # 12pt spacing below label
        # Limit to 1 line maximum
        value_width_teruskan = 5*cm  # Reduced width for single line
        teruskan_lines = wrap_text(teruskan_isi, value_width_teruskan, font_size=11) if teruskan_isi else []
        teruskan_lines = teruskan_lines[:1]  # Limit to 1 line maximum
        
        # Draw all lines with proper indentation
        for i, line in enumerate(teruskan_lines):
            c.drawString(value_x_teruskan, value_y_teruskan - (i * 0.4*cm), line)
        
        # Spacing after teruskan section
        lines_count = max(1, len(teruskan_lines))
        y_field_tambahan3 = value_y_teruskan - (lines_count * 0.4*cm) - 0.3*cm
        
        # Deadline section
        y_deadline_label = y_field_tambahan3 - 0.32*cm
        c.setFont("Helvetica-Bold", 11)
        deadline_label = "Harap diselesaikan Tanggal :"
        c.drawString(x_left, y_deadline_label, deadline_label)
        
        # Calculate colon position for deadline label
        deadline_colon_x = x_left + c.stringWidth(deadline_label, "Helvetica-Bold", 11)
        c.drawString(deadline_colon_x, y_deadline_label, "")  # No colon drawn separately since it's in the label
        
        y_deadline_box = y_deadline_label - 0.32*cm
        deadline_box_height = 0.8*cm
        # Make box width match the colon position
        deadline_box_width = deadline_colon_x - x_left + 0.3*cm
        c.rect(x_left, y_deadline_box - deadline_box_height, deadline_box_width, deadline_box_height)
        if data.get('harap_selesai_tgl', ''):
            c.setFont("Helvetica", 11)
            c.drawString(x_left + 0.3*cm, y_deadline_box - 0.6*cm, format_tanggal(data.get('harap_selesai_tgl', '')))
        
        # Instruction section
        x_instruksi = deadline_colon_x + 0.8*cm  # Add spacing from deadline area
        y_checkbox_manager_disp = y_disp - 3*LINE_HEIGHT
        y_label_instruksi = y_checkbox_manager_disp
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_instruksi, y_label_instruksi, "Isi Instruksi / Informasi")
        
        table_x = x_instruksi
        margin_bawah = 1.0*cm
        n_rows = 5
        # Calculate table width ensuring equal margins - use remaining space after left content
        available_width = width - MARGIN_LEFT - MARGIN_RIGHT
        left_content_width = x_instruksi - MARGIN_LEFT
        instruksi_table_width = available_width - left_content_width
        
        # Get instruction data
        isi_instruksi = data.get("isi_instruksi", [])
        
        # Ensure isi_instruksi is a list of dictionaries
        if not isinstance(isi_instruksi, list):
            isi_instruksi = []
        
        n_rows = max(1, len(isi_instruksi))
        # Kolom tanggal kembali ke 20%, dan wrapping jika terlalu panjang
        col_widths = [instruksi_table_width*0.20, instruksi_table_width*0.55, instruksi_table_width*0.25]
        instruksi_lines_per_row = []
        INSTRUKSI_FONT_SIZE = 11  # Changed to 11
        
        for row in range(n_rows):
            instruksi = ""
            if len(isi_instruksi) > row and isinstance(isi_instruksi[row], dict):
                instruksi = str(isi_instruksi[row].get("instruksi", ""))
            instruksi_lines = wrap_text(instruksi, col_widths[1] - 0.3*cm, font_size=INSTRUKSI_FONT_SIZE) if instruksi else []
            instruksi_lines_per_row.append(max(1, len(instruksi_lines)))
        
        # Definisikan table_y_top dan table_y_bottom sebelum digunakan
        table_y_top = y_label_instruksi - 0.4*cm  # Reduced spacing between label and table
        table_y_bottom = margin_bawah
        min_row_height = (table_y_top - table_y_bottom) / n_rows
        row_heights = []
        for lines in instruksi_lines_per_row:
            row_heights.append(max(min_row_height, lines * 0.6*cm + 0.3*cm))
        
        total_height = sum(row_heights)
        if total_height > (table_y_top - table_y_bottom):
            table_y = table_y_top + (total_height - (table_y_top - table_y_bottom))
        else:
            table_y = table_y_top
        
        # Jika semua instruksi kosong, gambar satu kotak kosong besar tanpa kolom
        if not isi_instruksi or all((not d.get('posisi') and not d.get('instruksi') and not d.get('tanggal')) for d in isi_instruksi if isinstance(d, dict)):
            table_y_top = y_label_instruksi - 0.4*cm
            table_y_bottom = margin_bawah
            c.rect(table_x, table_y_bottom, instruksi_table_width, table_y_top - table_y_bottom, stroke=1, fill=0)
            # Tidak perlu gambar kolom atau isi apapun
            c.save()
            return
        
        # Draw table header
        c.line(table_x, table_y, table_x + sum(col_widths), table_y)
        y_row = table_y
        row_bottoms = [y_row]
        
        # Draw table rows
        for row in range(n_rows):
            x_cursor = table_x
            posisi = ""
            instruksi = tanggal = ""
            
            # Get instruction data for this row
            if len(isi_instruksi) > row and isinstance(isi_instruksi[row], dict):
                posisi = str(isi_instruksi[row].get("posisi", ""))
                instruksi = str(isi_instruksi[row].get("instruksi", ""))
                tanggal = str(isi_instruksi[row].get("tanggal", ""))
            
            posisi_lines = wrap_text(posisi, col_widths[0] - 0.3*cm, font_size=INSTRUKSI_FONT_SIZE) if posisi else [""]
            instruksi_lines = wrap_text(instruksi, col_widths[1] - 0.3*cm, font_size=INSTRUKSI_FONT_SIZE) if instruksi else [""]
            tanggal_lines = wrap_text(tanggal, col_widths[2] - 0.3*cm, font_size=INSTRUKSI_FONT_SIZE) if tanggal else [""]
            
            row_height = row_heights[row]
            y_row -= row_height
            row_bottoms.append(y_row)
            
            # CENTER posisi
            total_lines = len(posisi_lines)
            text_block_height = total_lines * (0.6 * cm)
            y_start_posisi = (y_row + row_height) - ((row_height - text_block_height) / 2) - (0.4 * cm)
            for i, line in enumerate(posisi_lines):
                c.setFont("Helvetica", INSTRUKSI_FONT_SIZE)
                text_width = c.stringWidth(line, "Helvetica", INSTRUKSI_FONT_SIZE)
                cell_width = col_widths[0]
                x_center = x_cursor + (cell_width - text_width) / 2
                y_line = y_start_posisi - (i * 0.6 * cm)
                if line.strip() != "":
                    c.drawString(x_center, y_line, line)
            x_cursor += col_widths[0]
            
            # CENTER isi instruksi
            total_lines = len(instruksi_lines)
            text_block_height = total_lines * (0.6 * cm)
            y_start_instruksi = (y_row + row_height) - ((row_height - text_block_height) / 2) - (0.4 * cm)
            for i, line in enumerate(instruksi_lines):
                c.setFont("Helvetica", INSTRUKSI_FONT_SIZE)
                text_width = c.stringWidth(line, "Helvetica", INSTRUKSI_FONT_SIZE)
                cell_width = col_widths[1]
                x_center = x_cursor + (cell_width - text_width) / 2
                y_line = y_start_instruksi - i*0.6*cm
                if line.strip() != "":
                    c.drawString(x_center, y_line, line)
            x_cursor += col_widths[1]
            
            # CENTER tanggal instruksi
            total_lines = len(tanggal_lines)
            text_block_height = total_lines * (0.6 * cm)
            y_start_tanggal = (y_row + row_height) - ((row_height - text_block_height) / 2) - (0.4 * cm)
            for i, line in enumerate(tanggal_lines):
                c.setFont("Helvetica", INSTRUKSI_FONT_SIZE)
                text_width = c.stringWidth(line, "Helvetica", INSTRUKSI_FONT_SIZE)
                cell_width = col_widths[2]
                x_center = x_cursor + (cell_width - text_width) / 2
                y_line = y_start_tanggal - i*0.6*cm
                if line.strip() != "":
                    c.drawString(x_center, y_line, line)
        
        # Draw table borders
        for i in range(len(col_widths)+1):
            x = table_x + sum(col_widths[:i])
            c.line(x, table_y, x, y_row)
        for y in row_bottoms:
            c.line(table_x, y, table_x + sum(col_widths), y)
        
        c.save()
        
    except Exception as e:
        logging.error(f"save_form_to_pdf: {e}", exc_info=True)
        try:
            from tkinter import messagebox
            messagebox.showerror("Export PDF", f"Gagal ekspor ke PDF: {e}")
        except Exception:
            pass

def merge_pdfs(pdf_files, output_path):
    """Gabungkan beberapa file PDF menjadi satu file output_path."""
    merger = PyPDF2.PdfMerger()
    for pdf in pdf_files:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()
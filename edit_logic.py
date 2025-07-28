import traceback
import logging
from google_sheets_connect import get_sheets_service, SHEET_ID, update_row_in_sheet

logging.basicConfig(level=logging.WARNING)

ENHANCED_HEADER = [
    "No. Agenda", "No. Surat", "Tgl. Surat", "Perihal", "Asal Surat", "Ditujukan", 
    "Klasifikasi", "Disposisi kepada", "Untuk Di :", "Selesai Tgl.", "Kode Klasifikasi", 
    "Tgl. Penerimaan", "Indeks", "Bicarakan dengan", "Teruskan kepada", "Harap Selesai Tanggal",
    "Direktur Utama Instruksi", "Direktur Utama Tanggal", "Direktur Keuangan Instruksi", "Direktur Keuangan Tanggal",
    "Direktur Teknik Instruksi", "Direktur Teknik Tanggal", "GM Keuangan & Administrasi Instruksi", "GM Keuangan & Administrasi Tanggal",
    "GM Operasional & Pemeliharaan Instruksi", "GM Operasional & Pemeliharaan Tanggal", "Manager Instruksi", "Manager Tanggal"
]

# Ambil satu entry log dari Google Sheets berdasarkan No. Surat (bukan No. Agenda)
def get_log_entry_by_no_surat(no_surat):
    service = get_sheets_service()
    range_name = 'Sheet1!A6:AB'
    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=range_name
    ).execute()
    values = result.get('values', [])
    for row in values:
        row = row + ["" for _ in range(len(ENHANCED_HEADER) - len(row))]
        if row[1] == no_surat:
            data = {ENHANCED_HEADER[i]: row[i] for i in range(len(ENHANCED_HEADER))}
            # Konversi instruksi jabatan ke list of dict
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
                if instruksi_val:
                    instruksi_from_log.append({
                        "posisi": posisi_label,
                        "instruksi": instruksi_val,
                        "tanggal": tgl_val
                    })
            data["isi_instruksi"] = instruksi_from_log
            return data
    return None

# Overwrite baris di Google Sheets sesuai No. Surat dengan data_baru
def update_log_entry(data_lama, data_baru):
    service = get_sheets_service()
    range_name = 'Sheet1!A6:AB'
    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=range_name
    ).execute()
    values = result.get('values', []) or []
    idx = -1
    no_surat_lama = str(data_lama.get("No. Surat", "")).strip()
    for i, row in enumerate(values):
        row_full = row + ["" for _ in range(len(ENHANCED_HEADER) - len(row))]
        if str(row_full[1]).strip() == no_surat_lama:
            idx = i
            break
    if idx == -1:
        raise Exception(f"Data dengan No. Surat '{no_surat_lama}' tidak ditemukan di sheet, tidak bisa update.")

    # Gabungkan data_lama dan data_baru agar field yang tidak diubah tetap ada
    merged_data = dict(data_lama)
    for k, v in data_baru.items():
        # Izinkan update dengan string kosong untuk bisa mengosongkan field
        if v is not None:
            merged_data[k] = v

    # Sinkronisasi kolom tanggal penyelesaian agar konsisten
    if "Harap Selesai Tanggal" in merged_data:
        merged_data["Selesai Tgl."] = merged_data.get("Harap Selesai Tanggal", "")

    pejabat_labels = [
        "Direktur Utama", "Direktur Keuangan", "Direktur Teknik",
        "GM Keuangan & Administrasi", "GM Operasional & Pemeliharaan", "Manager"
    ]
    instruksi_map = {label: {"instruksi": "", "tanggal": ""} for label in pejabat_labels}
    for instruksi_item in merged_data.get("isi_instruksi", []):
        posisi = instruksi_item.get("posisi", "").strip()
        if posisi in instruksi_map:
            instruksi_map[posisi]["instruksi"] = instruksi_item.get("instruksi", "")
            instruksi_map[posisi]["tanggal"] = instruksi_item.get("tanggal", "")
    row_data = []
    for col in ENHANCED_HEADER:
        if col.endswith("Instruksi"):
            label = col.replace(" Instruksi", "")
            row_data.append(instruksi_map[label]["instruksi"])
        elif col.endswith("Tanggal"):
            label = col.replace(" Tanggal", "")
            row_data.append(instruksi_map[label]["tanggal"])
        else:
            row_data.append(merged_data.get(col, ""))
    update_row_in_sheet(row_data, idx+1)  # idx+1 karena row_number mulai dari 1 = baris ke-6

def build_complete_data(data_lama, data_baru, header):
    result = {}
    for col in header:
        if col in data_baru and data_baru[col] not in [None, ""]:
            result[col] = data_baru[col]
        elif col in data_lama:
            result[col] = data_lama[col]
        else:
            result[col] = ""
    return result
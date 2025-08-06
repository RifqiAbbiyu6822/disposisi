import os
import json
import logging
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

logging.basicConfig(level=logging.WARNING)
# Path ke credentials.json
CREDENTIALS_FILE = 'credentials/credentials.json'
# Sheet ID yang diberikan
SHEET_ID = '13-EgGz8JYYQ7FLQeCcVyPEXwYFruSGyz8KxPXrJlB7c'

# Scope untuk Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_sheets_service():
    """
    Membuat koneksi ke Google Sheets API dan mengembalikan service object.
    """
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    return service

def append_row_to_sheet(row_data, sheet_id=None, range_name='Sheet1!A1'):
    """
    Tambahkan satu baris data ke Google Sheet.
    row_data: list of values (1D)
    sheet_id: opsional, gunakan SHEET_ID default jika None
    range_name: range tujuan (default Sheet1!A1, akan append di bawah)
    """
    service = get_sheets_service()
    if sheet_id is None:
        sheet_id = SHEET_ID
    body = {'values': [row_data]}
    result = service.spreadsheets().values().append(
        spreadsheetId=sheet_id,
        range=range_name,
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body=body
    ).execute()
    # Ambil baris terakhir yang diisi dari response
    updated_range = result.get('updates', {}).get('updatedRange')
    if updated_range:
        # updated_range format: 'Sheet1!A12:AB12'
        import re
        m = re.match(r'[^!]+!A(\d+):', updated_range)
        if m:
            row_idx = int(m.group(1)) - 1  # 0-based index
            sheet_id_num = service.spreadsheets().get(spreadsheetId=sheet_id).execute()['sheets'][0]['properties']['sheetId']
            requests = [
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': sheet_id_num,
                            'startRowIndex': row_idx,
                            'endRowIndex': row_idx + 1,
                            'startColumnIndex': 0,
                            'endColumnIndex': 34
                        },
                        'cell': {
                            'userEnteredFormat': {
                                'wrapStrategy': 'WRAP',
                                'horizontalAlignment': 'LEFT',
                                'verticalAlignment': 'TOP'
                            }
                        },
                        'fields': 'userEnteredFormat.wrapStrategy,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment'
                    }
                }
            ]
            service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id,
                body={'requests': requests}
            ).execute()
    return result

def append_rows_to_sheet(rows_data, sheet_id=None, range_name='Sheet1!A1'):
    """
    Tambahkan banyak baris data ke Google Sheet sekaligus.
    rows_data: list of list of values (2D)
    sheet_id: opsional, gunakan SHEET_ID default jika None
    range_name: range tujuan (default Sheet1!A1, akan append di bawah)
    """
    service = get_sheets_service()
    if sheet_id is None:
        sheet_id = SHEET_ID
    body = {'values': rows_data}
    result = service.spreadsheets().values().append(
        spreadsheetId=sheet_id,
        range=range_name,
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body=body
    ).execute()
    return result

def create_new_sheet(title, sheet_id=None):
    """
    Membuat sheet baru di spreadsheet dengan nama title.
    sheet_id: opsional, gunakan SHEET_ID default jika None
    Return: sheetId dari sheet baru
    """
    service = get_sheets_service()
    if sheet_id is None:
        sheet_id = SHEET_ID
    requests = [{
        'addSheet': {
            'properties': {
                'title': title
            }
        }
    }]
    body = {'requests': requests}
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body=body
    ).execute()
    return response['replies'][0]['addSheet']['properties']['sheetId']

def write_multilayer_header(sheet_id=None, sheet_name='Sheet1'):
    """
    Menulis header multi-layer (4 baris pertama) sesuai format laporan disposisi ke Google Sheets.
    Akan menimpa 4 baris pertama sheet.
    """
    service = get_sheets_service()
    if sheet_id is None:
        sheet_id = SHEET_ID
    # Baris 1: Label utama (akan di-merge)
    row1 = ["LAPORAN DISPOSISI"] + ["" for _ in range(33)]
    # Baris 2: Tanggal update (akan di-merge)
    import datetime
    today_str = datetime.datetime.now().strftime('%d-%m-%Y')
    row2 = [f"terakhir di update: {today_str}"] + ["" for _ in range(33)]
    # Baris 3: Header utama (label field)
    row3 = [
        "No. Agenda", "No. Surat", "Tgl. Surat", "Perihal", "Asal Surat", "Ditujukan", "Klasifikasi", "Disposisi kepada", "Untuk Di :", "Selesai Tgl.", "Kode Klasifikasi", "Tgl. Penerimaan", "Indeks", "Bicarakan dengan", "Teruskan kepada", "Harap Selesai Tanggal",
        "Direktur Utama", "", "Direktur Keuangan", "", "Direktur Teknik", "", "GM Keuangan & Administrasi", "", "GM Operasional & Pemeliharaan", "", 
        "Manager Pemeliharaan", "", "Manager Operasional", "", "Manager Administrasi", "", "Manager Keuangan", ""
    ]
    # Baris 4: Sub-header jabatan
    row4 = ["" for _ in range(34)]
    row4[16] = "Instruksi"; row4[17] = "Tanggal"
    row4[18] = "Instruksi"; row4[19] = "Tanggal"
    row4[20] = "Instruksi"; row4[21] = "Tanggal"
    row4[22] = "Instruksi"; row4[23] = "Tanggal"
    row4[24] = "Instruksi"; row4[25] = "Tanggal"
    row4[26] = "Instruksi"; row4[27] = "Tanggal"
    row4[28] = "Instruksi"; row4[29] = "Tanggal"
    row4[30] = "Instruksi"; row4[31] = "Tanggal"
    row4[32] = "Instruksi"; row4[33] = "Tanggal"
    # Baris 5: Kosong
    row5 = ["" for _ in range(34)]
    body = {
        'values': [row1, row2, row3, row4, row5]
    }
    range_name = f'{sheet_name}!A1:AH5'
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()
    sheet_id_num = service.spreadsheets().get(spreadsheetId=sheet_id).execute()['sheets'][0]['properties']['sheetId']
    requests = [
        # Merge A1:AH1 for main label
        {
            'mergeCells': {
                'range': {
                    'sheetId': sheet_id_num,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 34
                },
                'mergeType': 'MERGE_ALL'
            }
        },
        # Merge A2:AH2 for update date
        {
            'mergeCells': {
                'range': {
                    'sheetId': sheet_id_num,
                    'startRowIndex': 1,
                    'endRowIndex': 2,
                    'startColumnIndex': 0,
                    'endColumnIndex': 34
                },
                'mergeType': 'MERGE_ALL'
            }
        },
        # Set value for A1 (center, font besar)
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id_num,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 34
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': 'CENTER',
                        'verticalAlignment': 'MIDDLE',
                        'textFormat': {'fontSize': 16, 'bold': True}
                    }
                },
                'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat)'
            }
        },
        # Set value for A2 (center, font normal)
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id_num,
                    'startRowIndex': 1,
                    'endRowIndex': 2,
                    'startColumnIndex': 0,
                    'endColumnIndex': 34
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': 'CENTER',
                        'verticalAlignment': 'MIDDLE',
                        'textFormat': {'fontSize': 11, 'bold': False}
                    }
                },
                'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat)'
            }
        },
        # Center all header cells (A1:AH5)
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id_num,
                    'startRowIndex': 0,
                    'endRowIndex': 5,
                    'startColumnIndex': 0,
                    'endColumnIndex': 34
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': 'CENTER',
                        'verticalAlignment': 'MIDDLE'
                    }
                },
                'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment)'
            }
        },
        # Tambahkan border all untuk header (A1:AH5)
        {
            'updateBorders': {
                'range': {
                    'sheetId': sheet_id_num,
                    'startRowIndex': 0,
                    'endRowIndex': 5,
                    'startColumnIndex': 0,
                    'endColumnIndex': 34
                },
                'top':    {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                'bottom': {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                'left':   {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                'right':  {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                'innerHorizontal': {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                'innerVertical':   {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}}
            }
        }
    ]
    service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={'requests': requests}
    ).execute()
    # Set auto wrap untuk seluruh kolom data (A6:AB)
    requests_wrap = [
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id_num,
                    'startRowIndex': 5,  
                    'endRowIndex': 1000, 
                    'startColumnIndex': 0,
                    'endColumnIndex': 34
                },
                'cell': {
                    'userEnteredFormat': {
                        'wrapStrategy': 'WRAP',
                        'horizontalAlignment': 'LEFT',
                        'verticalAlignment': 'TOP'
                    }
                },
                'fields': 'userEnteredFormat.wrapStrategy,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment'
            }
        }
    ]
    service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={'requests': requests_wrap}
    ).execute()

def update_row_in_sheet(row_data, row_number, sheet_id=None, sheet_name='Sheet1'):
    """
    Overwrite baris ke-row_number (mulai dari 1 = baris A6) dengan row_data (list, 28 kolom).
    row_number: 1 berarti baris ke-6 di sheet (A6), 2 = A7, dst.
    """
    service = get_sheets_service()
    if sheet_id is None:
        sheet_id = SHEET_ID
    range_name = f'{sheet_name}!A{row_number+5}:AB{row_number+5}'
    body = {'values': [row_data]}
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()

if __name__ == "__main__":
    # Contoh penggunaan: cek koneksi dan print judul spreadsheet
    service = get_sheets_service()
    sheet = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    print("Connected to sheet:", sheet.get('properties', {}).get('title')) 
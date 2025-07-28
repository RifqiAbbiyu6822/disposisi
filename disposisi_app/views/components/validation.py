def is_no_surat_unique(no_surat, get_sheets_service, SHEET_ID):
    service = get_sheets_service()
    range_name = 'Sheet1!B6:B'  # Kolom B = No. Surat
    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=range_name
    ).execute()
    values = result.get('values', []) or []
    no_surat_list = [str(row[0]).strip() for row in values if row and len(row) > 0]
    return str(no_surat).strip() not in no_surat_list 
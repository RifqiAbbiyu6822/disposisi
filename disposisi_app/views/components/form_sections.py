from tkinter import ttk, Text
from tkcalendar import DateEntry

def populate_frame_kiri(frame, vars):
    input_widgets = {}
    labels = ["No. Agenda:", "No. Surat:", "Tgl. Surat:", "Perihal:", "Asal Surat:", "Ditujukan:"]
    vars_keys = ["no_agenda", "no_surat", "tgl_surat", "perihal", "asal_surat", "ditujukan"]
    for i, (label, key) in enumerate(zip(labels, vars_keys)):
        ttk.Label(frame, text=label, style="TLabel").grid(row=i, column=0, sticky="w", pady=1, padx=(0,5))
        if key == "tgl_surat":
            input_widgets[key] = DateEntry(frame, width=25, date_pattern="dd-mm-yyyy", font=("Segoe UI", 10))
            input_widgets[key].grid(row=i, column=1, sticky="ew", pady=1)
            input_widgets[key].delete(0, 'end')
            # Tombol emoji clear di kanan field
            def clear_date(entry=input_widgets[key]):
                entry.delete(0, 'end')
            btn_clear = ttk.Button(frame, text="🗑️", width=2, command=lambda entry=input_widgets[key]: clear_date(entry))
            btn_clear.grid(row=i, column=2, padx=(5,0), sticky="e")
        elif key in ["perihal", "asal_surat", "ditujukan"]:
            input_widgets[key] = Text(frame, height=2, width=30, wrap="word", font=("Segoe UI", 10), borderwidth=1, relief="solid", highlightthickness=0)
            input_widgets[key].grid(row=i, column=1, sticky="ew", pady=1)
        else:
            width = 35 if key in ["no_agenda", "no_surat"] else 30
            input_widgets[key] = ttk.Entry(frame, textvariable=vars[key], width=width, style="TEntry")
            input_widgets[key].grid(row=i, column=1, sticky="ew", pady=1)
    return input_widgets

def populate_frame_kanan(frame, vars):
    input_widgets = {}
    def handle_klasifikasi_change(selected):
        for key in ["rahasia", "penting", "segera"]:
            if key == selected:
                vars[key].set(1)
            else:
                vars[key].set(0)
    frame.columnconfigure(0, weight=1)
    frame.columnconfigure(1, weight=1)
    checkbox_frame = ttk.Frame(frame)
    checkbox_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
    checkbox_frame.columnconfigure(0, weight=1)
    checkbox_frame.columnconfigure(1, weight=1)
    klasifikasi_items = [
        ("RAHASIA", "rahasia"),
        ("PENTING", "penting"),
        ("SEGERA", "segera")
    ]
    for i, (text, key) in enumerate(klasifikasi_items):
        row = i // 2
        col = i % 2
        ttk.Checkbutton(checkbox_frame, text=text, variable=vars[key], style="TCheckbutton", command=lambda k=key: handle_klasifikasi_change(k)).grid(
            row=row, column=col, sticky="w", pady=1, padx=(0, 10)
        )
    # Tanggal Penerimaan
    ttk.Label(frame, text="Tanggal Penerimaan:", style="TLabel").grid(row=1, column=0, sticky="w", pady=(5, 0))
    tgl_terima_entry = DateEntry(frame, width=20, date_pattern="dd-mm-yyyy", font=("Segoe UI", 10))
    tgl_terima_entry.grid(row=1, column=1, sticky="ew", pady=1)
    tgl_terima_entry.delete(0, 'end')
    input_widgets["tgl_terima"] = tgl_terima_entry
    def clear_tgl_terima():
        tgl_terima_entry.delete(0, 'end')
    btn_clear_tgl_terima = ttk.Button(frame, text="🗑️", width=2, command=clear_tgl_terima)
    btn_clear_tgl_terima.grid(row=1, column=2, padx=(5,0), sticky="e")
    # Kode / Klasifikasi
    ttk.Label(frame, text="Kode / Klasifikasi:", style="TLabel").grid(row=2, column=0, sticky="w", pady=(5, 0))
    ttk.Entry(frame, textvariable=vars["kode_klasifikasi"], style="TEntry", width=35).grid(row=2, column=1, columnspan=2, sticky="ew", pady=1)
    # Indeks
    ttk.Label(frame, text="Indeks:", style="TLabel").grid(row=3, column=0, sticky="w", pady=(5, 0))
    ttk.Entry(frame, textvariable=vars["indeks"], style="TEntry", width=35).grid(row=3, column=1, columnspan=2, sticky="ew", pady=1)
    return input_widgets

def populate_frame_disposisi(frame, vars):
    for key in ["dir_utama", "dir_keu", "dir_teknik", "gm_keu", "gm_ops", "manager"]:
        if key not in vars:
            from tkinter import IntVar
            vars[key] = IntVar()
    frame.columnconfigure(0, weight=1)
    frame.columnconfigure(1, weight=1)
    disposisi_items = [
        ("Direktur Utama", "dir_utama"),
        ("Direktur Keuangan", "dir_keu"),
        ("Direktur Teknik", "dir_teknik"),
        ("GM Keuangan & Administrasi", "gm_keu"),
        ("GM Operasional & Pemeliharaan", "gm_ops"),
        ("Manager", "manager")
    ]
    for i, (text, key) in enumerate(disposisi_items):
        row = i // 2
        col = i % 2
        ttk.Checkbutton(frame, text=text, variable=vars[key], style="TCheckbutton").grid(
            row=row, column=col, sticky="w", pady=1, padx=(0, 10)
        )

def populate_frame_instruksi(frame, vars):
    input_widgets = {}
    frame.columnconfigure(0, weight=1)
    frame.columnconfigure(1, weight=1)
    checks = ["Ketahui & File", "Proses Selesai", "Teliti & Pendapat", "Buatkan Resume", "Edarkan", "Sesuai Disposisi", "Bicarakan dengan Saya"]
    vars_keys = ["ketahui_file", "proses_selesai", "teliti_pendapat", "buatkan_resume", "edarkan", "sesuai_disposisi", "bicarakan_saya"]
    for i, (label, key) in enumerate(zip(checks, vars_keys)):
        row = i // 2
        col = i % 2
        ttk.Checkbutton(frame, text=label, variable=vars[key], style="TCheckbutton").grid(
            row=row, column=col, sticky="w", pady=1, padx=(0, 10)
        )
    row_offset = (len(checks) + 1) // 2
    # Harap Selesai Tanggal
    ttk.Label(frame, text="Harap Selesai Tgl:", style="TLabel").grid(row=row_offset, column=0, sticky="w", pady=(8,0))
    harap_selesai_entry = DateEntry(frame, width=20, date_pattern="dd-mm-yyyy", font=("Segoe UI", 10))
    harap_selesai_entry.grid(row=row_offset, column=1, sticky="ew", pady=1)
    harap_selesai_entry.delete(0, 'end')
    input_widgets["harap_selesai_tgl"] = harap_selesai_entry
    def clear_harap_selesai():
        harap_selesai_entry.delete(0, 'end')
    btn_clear_harap_selesai = ttk.Button(frame, text="🗑️", width=2, command=clear_harap_selesai)
    btn_clear_harap_selesai.grid(row=row_offset, column=2, padx=(5,0), sticky="e")
    ttk.Label(frame, text="Bicarakan dengan:", style="TLabel").grid(row=row_offset+1, column=0, columnspan=2, sticky="w", pady=(8,0))
    bicarakan_dengan = Text(frame, height=2, width=25, wrap="word", font=("Segoe UI", 10), borderwidth=1, relief="solid", highlightthickness=0)
    bicarakan_dengan.grid(row=row_offset+2, column=0, columnspan=2, sticky="ew", pady=1)
    input_widgets["bicarakan_dengan"] = bicarakan_dengan
    ttk.Label(frame, text="Teruskan Kepada:", style="TLabel").grid(row=row_offset+3, column=0, columnspan=2, sticky="w", pady=(5,0))
    teruskan_kepada = Text(frame, height=2, width=25, wrap="word", font=("Segoe UI", 10), borderwidth=1, relief="solid", highlightthickness=0)
    teruskan_kepada.grid(row=row_offset+4, column=0, columnspan=2, sticky="ew", pady=1)
    input_widgets["teruskan_kepada"] = teruskan_kepada
    return input_widgets
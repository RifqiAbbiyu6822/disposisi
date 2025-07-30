# disposisi_app/views/components/form_sections.py - FIXED VERSION with better layout
from tkinter import ttk, Text
from tkcalendar import DateEntry
import tkinter as tk

def create_modern_input_frame(parent, label_text, compact=False):
    """Create a modern input frame with label - now with compact option"""
    frame = tk.Frame(parent, bg="#FFFFFF")
    padding_y = (0, 8) if compact else (0, 12)
    frame.pack(fill="x", pady=padding_y)
    
    # Label with modern styling - smaller for compact mode
    font_size = 9 if compact else 10
    label = tk.Label(frame, text=label_text, 
                    font=("Inter", font_size) if "Inter" in tk.font.families() else ("Segoe UI", font_size),
                    bg="#FFFFFF", fg="#475569")
    label_pady = (0, 4) if compact else (0, 6)
    label.pack(anchor="w", pady=label_pady)
    
    return frame

def populate_frame_kiri(frame, vars, compact=False):
    """FIXED: More compact version of left frame"""
    input_widgets = {}
    
    # Modern card-like inner frame
    inner_frame = tk.Frame(frame, bg="#FFFFFF")
    inner_frame.pack(fill="both", expand=True)
    
    fields = [
        ("No. Agenda", "no_agenda", "text"),
        ("No. Surat", "no_surat", "text"),
        ("Tgl. Surat", "tgl_surat", "date"),
        ("Perihal", "perihal", "textarea"),
        ("Asal Surat", "asal_surat", "textarea"),
        ("Ditujukan", "ditujukan", "textarea")
    ]
    
    for label, key, field_type in fields:
        container = create_modern_input_frame(inner_frame, label, compact)
        
        if field_type == "date":
            # Compact date picker
            date_frame = tk.Frame(container, bg="#FFFFFF")
            date_frame.pack(fill="x")
            
            font_size = 10 if compact else 11
            input_widgets[key] = DateEntry(date_frame, width=25 if compact else 30, 
                                         date_pattern="dd-mm-yyyy", 
                                         font=("Inter", font_size) if "Inter" in tk.font.families() else ("Segoe UI", font_size),
                                         background='#3B82F6',
                                         foreground='white',
                                         borderwidth=2,
                                         headersbackground='#1E293B',
                                         headersforeground='white',
                                         selectbackground='#2563EB',
                                         selectforeground='white',
                                         normalbackground='white',
                                         normalforeground='#0F172A',
                                         weekendbackground='#F1F5F9',
                                         weekendforeground='#64748B')
            input_widgets[key].pack(side="left", fill="x", expand=True)
            input_widgets[key].delete(0, 'end')
            
            # Compact clear button
            clear_btn = tk.Button(date_frame, text="✕", 
                                font=("Arial", 9 if compact else 10, "bold"),
                                bg="#F1F5F9", fg="#64748B",
                                bd=0, padx=8 if compact else 12, pady=4 if compact else 6,
                                cursor="hand2",
                                command=lambda w=input_widgets[key]: w.delete(0, 'end'))
            clear_btn.pack(side="right", padx=(6 if compact else 8, 0))
            
            # Hover effect
            def on_enter(e, btn=clear_btn):
                btn.config(bg="#E2E8F0", fg="#475569")
            def on_leave(e, btn=clear_btn):
                btn.config(bg="#F1F5F9", fg="#64748B")
            
            clear_btn.bind("<Enter>", on_enter)
            clear_btn.bind("<Leave>", on_leave)
            
        elif field_type == "textarea":
            # Compact text area with border
            text_frame = tk.Frame(container, bg="#E2E8F0", bd=1)
            text_frame.pack(fill="x")
            
            height = 1.5 if compact else 2
            font_size = 10 if compact else 11
            padding_x = 8 if compact else 12
            padding_y = 6 if compact else 10
            
            input_widgets[key] = Text(text_frame, height=height, wrap="word", 
                                    font=("Inter", font_size) if "Inter" in tk.font.families() else ("Segoe UI", font_size),
                                    bg="#FFFFFF", fg="#0F172A",
                                    bd=0, padx=padding_x, pady=padding_y,
                                    insertbackground="#3B82F6")
            input_widgets[key].pack(fill="x", padx=1, pady=1)
            
            # Focus effect
            def on_focus_in(e, frame=text_frame):
                frame.config(bg="#3B82F6")
            def on_focus_out(e, frame=text_frame):
                frame.config(bg="#E2E8F0")
            
            input_widgets[key].bind("<FocusIn>", on_focus_in)
            input_widgets[key].bind("<FocusOut>", on_focus_out)
            
        else:  # text input
            font_size = 10 if compact else 11
            input_widgets[key] = ttk.Entry(container, textvariable=vars[key], 
                                         font=("Inter", font_size) if "Inter" in tk.font.families() else ("Segoe UI", font_size),
                                         style="TEntry")
            input_widgets[key].pack(fill="x")
    
    return input_widgets

def populate_frame_kanan(frame, vars, compact=False):
    """FIXED: More compact version of right frame"""
    input_widgets = {}
    
    # Modern classification section with reduced spacing
    klasifikasi_frame = tk.Frame(frame, bg="#FFFFFF")
    padding_y = (0, 15) if compact else (0, 20)
    klasifikasi_frame.pack(fill="x", pady=padding_y)
    
    font_size = 10 if compact else 11
    klasifikasi_label = tk.Label(klasifikasi_frame, text="Klasifikasi Dokumen",
                               font=("Inter", font_size, "600") if "Inter" in tk.font.families() else ("Segoe UI", font_size, "bold"),
                               bg="#FFFFFF", fg="#0F172A")
    label_pady = (0, 8) if compact else (0, 10)
    klasifikasi_label.pack(anchor="w", pady=label_pady)
    
    # Compact radio-style checkboxes
    klasifikasi_items = [
        ("RAHASIA", "rahasia", "#EF4444"),
        ("PENTING", "penting", "#F59E0B"),
        ("SEGERA", "segera", "#3B82F6")
    ]
    
    for text, key, color in klasifikasi_items:
        checkbox_frame = tk.Frame(klasifikasi_frame, bg="#FFFFFF")
        checkbox_pady = 3 if compact else 4
        checkbox_frame.pack(anchor="w", pady=checkbox_pady)
        
        # Custom checkbox with colored indicator - smaller for compact
        indicator_font_size = 12 if compact else 14
        indicator = tk.Label(checkbox_frame, text="○", font=("Arial", indicator_font_size),
                           bg="#FFFFFF", fg="#E2E8F0")
        indicator_padx = (0, 6) if compact else (0, 8)
        indicator.pack(side="left", padx=indicator_padx)
        
        label_font_size = 9 if compact else 10
        label = tk.Label(checkbox_frame, text=text,
                        font=("Inter", label_font_size) if "Inter" in tk.font.families() else ("Segoe UI", label_font_size),
                        bg="#FFFFFF", fg="#475569")
        label.pack(side="left")
        
        def toggle_check(e, var=vars[key], ind=indicator, clr=color):
            current = var.get()
            var.set(1 - current)
            if var.get():
                ind.config(text="●", fg=clr)
                # Uncheck others
                for t, k, _ in klasifikasi_items:
                    if k != key:
                        vars[k].set(0)
            else:
                ind.config(text="○", fg="#E2E8F0")
        
        checkbox_frame.bind("<Button-1>", toggle_check)
        label.bind("<Button-1>", toggle_check)
        indicator.bind("<Button-1>", toggle_check)
        
        # Update indicator based on var
        def update_indicator(*args, var=vars[key], ind=indicator, clr=color):
            if var.get():
                ind.config(text="●", fg=clr)
            else:
                ind.config(text="○", fg="#E2E8F0")
        
        vars[key].trace("w", update_indicator)
        update_indicator()
    
    # Date and classification fields - more compact
    fields = [
        ("Tanggal Penerimaan", "tgl_terima", "date"),
        ("Kode / Klasifikasi", "kode_klasifikasi", "text"),
        ("Indeks", "indeks", "text")
    ]
    
    for label, key, field_type in fields:
        container = create_modern_input_frame(frame, label, compact)
        
        if field_type == "date":
            date_frame = tk.Frame(container, bg="#FFFFFF")
            date_frame.pack(fill="x")
            
            width = 20 if compact else 25
            font_size = 10 if compact else 11
            
            tgl_terima_entry = DateEntry(date_frame, width=width,
                                       date_pattern="dd-mm-yyyy",
                                       font=("Inter", font_size) if "Inter" in tk.font.families() else ("Segoe UI", font_size),
                                       background='#3B82F6',
                                       foreground='white',
                                       borderwidth=2)
            tgl_terima_entry.pack(side="left", fill="x", expand=True)
            tgl_terima_entry.delete(0, 'end')
            input_widgets[key] = tgl_terima_entry
            
            # Compact clear button
            clear_btn = tk.Button(date_frame, text="✕",
                                font=("Arial", 9 if compact else 10, "bold"),
                                bg="#F1F5F9", fg="#64748B",
                                bd=0, padx=8 if compact else 12, pady=4 if compact else 6,
                                cursor="hand2",
                                command=lambda: tgl_terima_entry.delete(0, 'end'))
            clear_btn.pack(side="right", padx=(6 if compact else 8, 0))
        else:
            font_size = 10 if compact else 11
            entry = ttk.Entry(container, textvariable=vars[key],
                            font=("Inter", font_size) if "Inter" in tk.font.families() else ("Segoe UI", font_size),
                            style="TEntry")
            entry.pack(fill="x")
    
    # Return the tgl_terima_entry for compatibility
    return input_widgets.get("tgl_terima") or input_widgets

def populate_frame_disposisi(frame, vars, compact=False):
    """FIXED: More compact disposisi frame"""
    # Modern checkbox list with hover effects - more compact
    disposisi_items = [
        ("Direktur Utama", "dir_utama"),
        ("Direktur Keuangan", "dir_keu"),
        ("Direktur Teknik", "dir_teknik"),
        ("GM Keuangan & Administrasi", "gm_keu"),
        ("GM Operasional & Pemeliharaan", "gm_ops"),
        ("Manager", "manager")
    ]
    
    for text, key in disposisi_items:
        # Compact checkbox container
        check_container = tk.Frame(frame, bg="#FFFFFF", relief="flat")
        pady = 2 if compact else 3
        check_container.pack(fill="x", pady=pady)
        
        cb = ttk.Checkbutton(check_container, text=text, variable=vars[key],
                           style="TCheckbutton")
        padx = 6 if compact else 8
        pady = 4 if compact else 8
        cb.pack(anchor="w", padx=padx, pady=pady)
        
        # Hover effect
        def on_enter(e, container=check_container):
            container.config(bg="#F8FAFC")
        def on_leave(e, container=check_container):
            container.config(bg="#FFFFFF")
        
        check_container.bind("<Enter>", on_enter)
        check_container.bind("<Leave>", on_leave)

def populate_frame_instruksi(frame, vars, compact=False):
    """FIXED: More compact instruction frame"""
    input_widgets = {}
    
    # Compact checkbox grid
    checks = [
        "Ketahui & File", "Proses Selesai", "Teliti & Pendapat",
        "Buatkan Resume", "Edarkan", "Sesuai Disposisi", "Bicarakan dengan Saya"
    ]
    vars_keys = [
        "ketahui_file", "proses_selesai", "teliti_pendapat",
        "buatkan_resume", "edarkan", "sesuai_disposisi", "bicarakan_saya"
    ]
    
    # Create a more compact grid layout
    for i, (label, key) in enumerate(zip(checks, vars_keys)):
        row = i // 2
        col = i % 2
        
        check_frame = tk.Frame(frame, bg="#FFFFFF")
        padx = 2 if compact else 4
        pady = 2 if compact else 3
        check_frame.grid(row=row, column=col, sticky="ew", padx=padx, pady=pady)
        
        cb = ttk.Checkbutton(check_frame, text=label, variable=vars[key],
                           style="TCheckbutton")
        cb_padx = 4 if compact else 8
        cb_pady = 4 if compact else 8
        cb.pack(anchor="w", padx=cb_padx, pady=cb_pady)
    
    # Additional fields section - more compact spacing
    row_offset = (len(checks) + 1) // 2 + 1
    
    # Deadline date - compact
    deadline_container = create_modern_input_frame(frame, "Harap Selesai Tanggal", compact)
    deadline_pady = (10, 0) if compact else (15, 0)
    deadline_container.grid(row=row_offset, column=0, columnspan=2, sticky="ew", pady=deadline_pady)
    
    date_frame = tk.Frame(deadline_container, bg="#FFFFFF")
    date_frame.pack(fill="x")
    
    width = 20 if compact else 25
    font_size = 10 if compact else 11
    
    harap_selesai_entry = DateEntry(date_frame, width=width,
                                  date_pattern="dd-mm-yyyy",
                                  font=("Inter", font_size) if "Inter" in tk.font.families() else ("Segoe UI", font_size),
                                  background='#3B82F6',
                                  foreground='white')
    harap_selesai_entry.pack(side="left", fill="x", expand=True)
    harap_selesai_entry.delete(0, 'end')
    input_widgets["harap_selesai_tgl"] = harap_selesai_entry
    
    # Compact text areas
    text_fields = [
        ("Bicarakan dengan", "bicarakan_dengan"),
        ("Teruskan Kepada", "teruskan_kepada")
    ]
    
    for idx, (label, key) in enumerate(text_fields):
        container = create_modern_input_frame(frame, label, compact)
        container_pady = (8, 0) if compact else (10, 0)
        container.grid(row=row_offset + idx + 1, column=0, columnspan=2, 
                      sticky="ew", pady=container_pady)
        
        text_frame = tk.Frame(container, bg="#E2E8F0", bd=1)
        text_frame.pack(fill="x")
        
        # More compact text widget
        height = 1.5 if compact else 2
        font_size = 10 if compact else 11
        padx = 8 if compact else 12
        pady = 6 if compact else 10
        
        text_widget = Text(text_frame, height=height, wrap="word",
                          font=("Inter", font_size) if "Inter" in tk.font.families() else ("Segoe UI", font_size),
                          bg="#FFFFFF", fg="#0F172A",
                          bd=0, padx=padx, pady=pady,
                          insertbackground="#3B82F6")
        text_widget.pack(fill="x", padx=1, pady=1)
        input_widgets[key] = text_widget
        
        # Focus effect
        def on_focus_in(e, frame=text_frame):
            frame.config(bg="#3B82F6")
        def on_focus_out(e, frame=text_frame):
            frame.config(bg="#E2E8F0")
        
        text_widget.bind("<FocusIn>", on_focus_in)
        text_widget.bind("<FocusOut>", on_focus_out)
    
    return input_widgets
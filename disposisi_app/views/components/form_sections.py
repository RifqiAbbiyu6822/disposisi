from tkinter import ttk, Text, font
from tkcalendar import DateEntry
import tkinter as tk

def create_modern_input_frame(parent, label_text):
    """Create a modern input frame with label using grid"""
    frame = tk.Frame(parent, bg="#FFFFFF")
    # Use grid instead of pack to be consistent
    row_count = len([child for child in parent.winfo_children() if isinstance(child, (tk.Frame, ttk.Frame))])
    frame.grid(row=row_count, column=0, sticky="ew", pady=(0, 12))
    parent.grid_rowconfigure(row_count, weight=0)
    parent.grid_columnconfigure(0, weight=1)
    
    # Label with modern styling
    label = tk.Label(frame, text=label_text, 
                    font=("Inter", 10) if "Inter" in tk.font.families() else ("Segoe UI", 10),
                    bg="#FFFFFF", fg="#475569")
    label.grid(row=0, column=0, sticky="w", pady=(0, 6))
    
    frame.grid_columnconfigure(0, weight=1)
    
    return frame

def populate_frame_kiri(frame, vars):
    input_widgets = {}
    
    # Modern card-like inner frame
    inner_frame = tk.Frame(frame, bg="#FFFFFF")
    inner_frame.grid(row=0, column=0, sticky="nsew")
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)
    
    fields = [
        ("No. Agenda", "no_agenda", "text"),
        ("No. Surat", "no_surat", "text"),
        ("Tgl. Surat", "tgl_surat", "date"),
        ("Perihal", "perihal", "textarea"),
        ("Asal Surat", "asal_surat", "textarea"),
        ("Ditujukan", "ditujukan", "textarea")
    ]
    
    for row_idx, (label, key, field_type) in enumerate(fields):
        # Create label
        label_widget = tk.Label(inner_frame, text=label,
                               font=("Inter", 10) if "Inter" in tk.font.families() else ("Segoe UI", 10),
                               bg="#FFFFFF", fg="#475569")
        label_widget.grid(row=row_idx*2, column=0, sticky="w", pady=(0, 6))
        
        if field_type == "date":
            # Modern date picker
            date_frame = tk.Frame(inner_frame, bg="#FFFFFF")
            date_frame.grid(row=row_idx*2+1, column=0, sticky="ew", pady=(0, 12))
            
            input_widgets[key] = DateEntry(date_frame, width=30, 
                                         date_pattern="dd-mm-yyyy", 
                                         font=("Inter", 11) if "Inter" in tk.font.families() else ("Segoe UI", 11),
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
            input_widgets[key].grid(row=0, column=0, sticky="ew")
            input_widgets[key].delete(0, 'end')
            
            # Modern clear button
            clear_btn = tk.Button(date_frame, text="✕", 
                                font=("Arial", 10, "bold"),
                                bg="#F1F5F9", fg="#64748B",
                                bd=0, padx=12, pady=6,
                                cursor="hand2",
                                command=lambda w=input_widgets[key]: w.delete(0, 'end'))
            clear_btn.grid(row=0, column=1, padx=(8, 0))
            
            date_frame.grid_columnconfigure(0, weight=1)
            
            # Hover effect
            def on_enter(e, btn=clear_btn):
                btn.config(bg="#E2E8F0", fg="#475569")
            def on_leave(e, btn=clear_btn):
                btn.config(bg="#F1F5F9", fg="#64748B")
            
            clear_btn.bind("<Enter>", on_enter)
            clear_btn.bind("<Leave>", on_leave)
            
        elif field_type == "textarea":
            # Modern text area with border
            text_frame = tk.Frame(inner_frame, bg="#E2E8F0", bd=1)
            text_frame.grid(row=row_idx*2+1, column=0, sticky="ew", pady=(0, 12))
            
            input_widgets[key] = Text(text_frame, height=2, wrap="word", 
                                    font=("Inter", 11) if "Inter" in tk.font.families() else ("Segoe UI", 11),
                                    bg="#FFFFFF", fg="#0F172A",
                                    bd=0, padx=12, pady=10,
                                    insertbackground="#3B82F6")
            input_widgets[key].grid(row=0, column=0, sticky="ew")
            text_frame.grid_columnconfigure(0, weight=1)
            
            # Focus effect
            def on_focus_in(e, frame=text_frame):
                frame.config(bg="#3B82F6")
            def on_focus_out(e, frame=text_frame):
                frame.config(bg="#E2E8F0")
            
            input_widgets[key].bind("<FocusIn>", on_focus_in)
            input_widgets[key].bind("<FocusOut>", on_focus_out)
            
        else:  # text input
            input_widgets[key] = ttk.Entry(inner_frame, textvariable=vars[key], 
                                         font=("Inter", 11) if "Inter" in tk.font.families() else ("Segoe UI", 11),
                                         style="TEntry")
            input_widgets[key].grid(row=row_idx*2+1, column=0, sticky="ew", pady=(0, 12))
    
    # Configure grid weights
    inner_frame.grid_columnconfigure(0, weight=1)
    for i in range(len(fields)*2):
        inner_frame.grid_rowconfigure(i, weight=0)
    
    return input_widgets

def populate_frame_kanan(frame, vars):
    input_widgets = {}
    
    # Modern classification section
    klasifikasi_frame = tk.Frame(frame, bg="#FFFFFF")
    klasifikasi_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
    frame.grid_rowconfigure(0, weight=0)
    frame.grid_columnconfigure(0, weight=1)
    
    klasifikasi_label = tk.Label(klasifikasi_frame, text="Klasifikasi Dokumen",
                               font=("Inter", 11, "600") if "Inter" in tk.font.families() else ("Segoe UI", 11, "bold"),
                               bg="#FFFFFF", fg="#0F172A")
    klasifikasi_label.grid(row=0, column=0, sticky="w", pady=(0, 10))
    
    # Modern radio-style checkboxes
    klasifikasi_items = [
        ("RAHASIA", "rahasia", "#EF4444"),
        ("PENTING", "penting", "#F59E0B"),
        ("SEGERA", "segera", "#3B82F6")
    ]
    
    for idx, (text, key, color) in enumerate(klasifikasi_items):
        checkbox_frame = tk.Frame(klasifikasi_frame, bg="#FFFFFF")
        checkbox_frame.grid(row=idx+1, column=0, sticky="w", pady=4)
        
        # Custom checkbox with colored indicator
        indicator = tk.Label(checkbox_frame, text="○", font=("Arial", 14),
                           bg="#FFFFFF", fg="#E2E8F0")
        indicator.grid(row=0, column=0, padx=(0, 8))
        
        label = tk.Label(checkbox_frame, text=text,
                        font=("Inter", 10) if "Inter" in tk.font.families() else ("Segoe UI", 10),
                        bg="#FFFFFF", fg="#475569")
        label.grid(row=0, column=1)
        
        def toggle_check(e, var=vars[key], ind=indicator, clr=color, k=key):
            current = var.get()
            var.set(1 - current)
            if var.get():
                ind.config(text="●", fg=clr)
                # Uncheck others
                for t, other_key, _ in klasifikasi_items:
                    if other_key != k:
                        vars[other_key].set(0)
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
    
    klasifikasi_frame.grid_columnconfigure(0, weight=1)
    
    # Date and classification fields - Fixed: Proper row calculation
    fields = [
        ("Tanggal Penerimaan", "tgl_terima", "date"),
        ("Kode / Klasifikasi", "kode_klasifikasi", "text"),
        ("Indeks", "indeks", "text")
    ]
    
    tgl_terima_entry = None
    
    # Fixed: Calculate proper row offset after klasifikasi section
    row_offset = len(klasifikasi_items) + 1  # +1 for the label
    
    for idx, (label, key, field_type) in enumerate(fields):
        current_row = row_offset + (idx * 2)  # Fixed: Proper row calculation
        
        # Create label
        label_widget = tk.Label(frame, text=label,
                               font=("Inter", 10) if "Inter" in tk.font.families() else ("Segoe UI", 10),
                               bg="#FFFFFF", fg="#475569")
        label_widget.grid(row=current_row, column=0, sticky="w", pady=(10, 6))
        frame.grid_rowconfigure(current_row, weight=0)  # Fixed: Set row weight
        
        if field_type == "date":
            date_frame = tk.Frame(frame, bg="#FFFFFF")
            date_frame.grid(row=current_row + 1, column=0, sticky="ew", pady=(0, 10))  # Fixed: Proper row and padding
            frame.grid_rowconfigure(current_row + 1, weight=0)  # Fixed: Set row weight
            
            tgl_terima_entry = DateEntry(date_frame, width=25,
                                       date_pattern="dd-mm-yyyy",
                                       font=("Inter", 11) if "Inter" in tk.font.families() else ("Segoe UI", 11),
                                       background='#3B82F6',
                                       foreground='white',
                                       borderwidth=2)
            tgl_terima_entry.grid(row=0, column=0, sticky="ew")
            tgl_terima_entry.delete(0, 'end')
            input_widgets[key] = tgl_terima_entry
            
            # Clear button
            clear_btn = tk.Button(date_frame, text="✕",
                                font=("Arial", 10, "bold"),
                                bg="#F1F5F9", fg="#64748B",
                                bd=0, padx=12, pady=6,
                                cursor="hand2",
                                command=lambda: tgl_terima_entry.delete(0, 'end'))
            clear_btn.grid(row=0, column=1, padx=(8, 0))
            
            date_frame.grid_columnconfigure(0, weight=1)
        else:
            entry = ttk.Entry(frame, textvariable=vars[key],
                            font=("Inter", 11) if "Inter" in tk.font.families() else ("Segoe UI", 11),
                            style="TEntry")
            entry.grid(row=current_row + 1, column=0, sticky="ew", pady=(0, 10))  # Fixed: Proper row and padding
            frame.grid_rowconfigure(current_row + 1, weight=0)  # Fixed: Set row weight
    
    return tgl_terima_entry

def populate_frame_disposisi(frame, vars):
    # Modern checkbox list with hover effects
    disposisi_items = [
        ("Direktur Utama", "dir_utama"),
        ("Direktur Keuangan", "dir_keu"),
        ("Direktur Teknik", "dir_teknik"),
        ("GM Keuangan & Administrasi", "gm_keu"),
        ("GM Operasional & Pemeliharaan", "gm_ops"),
        ("Manager Pemeliharaan", "manager_pemeliharaan"),
        ("Manager Operasional", "manager_operasional"),
        ("Manager Administrasi", "manager_administrasi"),
        ("Manager Keuangan", "manager_keuangan")
    ]
    
    for idx, (text, key) in enumerate(disposisi_items):
        # Modern checkbox container
        check_container = tk.Frame(frame, bg="#FFFFFF", relief="flat")
        check_container.grid(row=idx, column=0, sticky="ew", pady=3)
        frame.grid_rowconfigure(idx, weight=0)
        
        cb = ttk.Checkbutton(check_container, text=text, variable=vars[key],
                           style="TCheckbutton")
        cb.grid(row=0, column=0, sticky="w", padx=8, pady=8)
        
        check_container.grid_columnconfigure(0, weight=1)
        
        # Hover effect
        def on_enter(e, container=check_container):
            container.config(bg="#F8FAFC")
        def on_leave(e, container=check_container):
            container.config(bg="#FFFFFF")
        
        check_container.bind("<Enter>", on_enter)
        check_container.bind("<Leave>", on_leave)
    
    frame.grid_columnconfigure(0, weight=1)

def populate_frame_instruksi(frame, vars):
    input_widgets = {}
    
    # Configure frame grid weights properly
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_columnconfigure(1, weight=1)
    
    # Modern checkbox grid
    checks = [
        "Ketahui & File", "Proses Selesai", "Teliti & Pendapat",
        "Buatkan Resume", "Edarkan", "Sesuai Disposisi", "Bicarakan dengan Saya"
    ]
    vars_keys = [
        "ketahui_file", "proses_selesai", "teliti_pendapat",
        "buatkan_resume", "edarkan", "sesuai_disposisi", "bicarakan_saya"
    ]
    
    # Create a grid layout with proper spacing
    for i, (label, key) in enumerate(zip(checks, vars_keys)):
        row = i // 2
        col = i % 2
        
        check_frame = tk.Frame(frame, bg="#FFFFFF")
        check_frame.grid(row=row, column=col, sticky="ew", padx=4, pady=3)
        
        cb = ttk.Checkbutton(check_frame, text=label, variable=vars[key],
                           style="TCheckbutton")
        cb.grid(row=0, column=0, sticky="w", padx=8, pady=8)
        
        check_frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(row, weight=0)  # Fixed: Set row weight to 0
    
    # Additional fields section - calculate proper row offset
    row_offset = (len(checks) + 1) // 2  # Fixed: Removed +1 to prevent extra spacing
    
    # Deadline date
    deadline_label = tk.Label(frame, text="Harap Selesai Tanggal",
                             font=("Inter", 10) if "Inter" in tk.font.families() else ("Segoe UI", 10),
                             bg="#FFFFFF", fg="#475569")
    deadline_label.grid(row=row_offset, column=0, columnspan=2, sticky="w", pady=(15, 6))
    frame.grid_rowconfigure(row_offset, weight=0)  # Fixed: Set row weight
    
    date_frame = tk.Frame(frame, bg="#FFFFFF")
    date_frame.grid(row=row_offset+1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
    frame.grid_rowconfigure(row_offset+1, weight=0)  # Fixed: Set row weight
    
    harap_selesai_entry = DateEntry(date_frame, width=25,
                                  date_pattern="dd-mm-yyyy",
                                  font=("Inter", 11) if "Inter" in tk.font.families() else ("Segoe UI", 11),
                                  background='#3B82F6',
                                  foreground='white')
    harap_selesai_entry.grid(row=0, column=0, sticky="ew")
    harap_selesai_entry.delete(0, 'end')
    input_widgets["harap_selesai_tgl"] = harap_selesai_entry
    input_widgets["harap_selesai_tgl_entry"] = harap_selesai_entry  # Add both keys for compatibility
    
    date_frame.grid_columnconfigure(0, weight=1)
    
    # Modern text areas with proper spacing
    text_fields = [
        ("Bicarakan dengan", "bicarakan_dengan"),
        ("Teruskan Kepada", "teruskan_kepada")
    ]
    
    for idx, (label, key) in enumerate(text_fields):
        current_row = row_offset + 2 + (idx * 2)  # Fixed: Better row calculation
        
        label_widget = tk.Label(frame, text=label,
                               font=("Inter", 10) if "Inter" in tk.font.families() else ("Segoe UI", 10),
                               bg="#FFFFFF", fg="#475569")
        label_widget.grid(row=current_row, column=0, columnspan=2, 
                         sticky="w", pady=(10, 6))
        frame.grid_rowconfigure(current_row, weight=0)  # Fixed: Set row weight
        
        text_frame = tk.Frame(frame, bg="#E2E8F0", bd=1)
        text_frame.grid(row=current_row + 1, column=0, columnspan=2, 
                       sticky="ew", pady=(0, 10))
        frame.grid_rowconfigure(current_row + 1, weight=0)  # Fixed: Set row weight
        
        text_widget = Text(text_frame, height=2, wrap="word",
                          font=("Inter", 11) if "Inter" in tk.font.families() else ("Segoe UI", 11),
                          bg="#FFFFFF", fg="#0F172A",
                          bd=0, padx=12, pady=10,
                          insertbackground="#3B82F6")
        text_widget.grid(row=0, column=0, sticky="ew")
        text_frame.grid_columnconfigure(0, weight=1)
        input_widgets[key] = text_widget
        
        # Focus effect
        def on_focus_in(e, frame=text_frame):
            frame.config(bg="#3B82F6")
        def on_focus_out(e, frame=text_frame):
            frame.config(bg="#E2E8F0")
        
        text_widget.bind("<FocusIn>", on_focus_in)
        text_widget.bind("<FocusOut>", on_focus_out)
    
    return input_widgets
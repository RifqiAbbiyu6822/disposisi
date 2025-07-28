import tkinter as tk
from tkinter import ttk

def create_button_frame(parent, callbacks, pdf_attachments, on_remove_pdf):
    button_frame = ttk.Frame(parent, padding="10")
    button_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    button_frame.columnconfigure(0, weight=1)

    # Compact PDF attachments section
    if pdf_attachments or True:  # Always show the section
        listbox_frame = ttk.LabelFrame(button_frame, text="📎 Lampiran PDF", padding="8")
        listbox_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        listbox_frame.columnconfigure(0, weight=1)
        if pdf_attachments:
            for i, f in enumerate(pdf_attachments):
                file_row = ttk.Frame(listbox_frame)
                file_row.grid(row=i, column=0, sticky="ew", pady=1)
                file_row.columnconfigure(1, weight=1)
                ttk.Label(file_row, text="📄", font=("Segoe UI", 9)).grid(row=0, column=0, padx=(0, 4))
                ttk.Label(file_row, text=f, font=("Segoe UI", 8), foreground=getattr(parent, 'text_fg', '#34495e')).grid(row=0, column=1, sticky="w")
                remove_btn = ttk.Button(file_row, text="×", width=2, command=lambda idx=i: on_remove_pdf(idx))
                remove_btn.grid(row=0, column=2, padx=(4, 0))
        else:
            ttk.Label(listbox_frame, text="Belum ada file terlampir", font=("Segoe UI", 8, "italic"), foreground=getattr(parent, 'muted_color', '#95a5a6')).grid(row=0, column=0, pady=4)

        # Add buttons for managing attachments and clearing form
        action_attach_frame = ttk.Frame(listbox_frame)
        action_attach_frame.grid(row=99, column=0, sticky="ew", pady=(8, 0))
        action_attach_frame.columnconfigure((0, 1, 2), weight=1)

        def on_add_pdf_attachment():
            func = callbacks.get("add_pdf_attachment")
            if callable(func):
                func()
            else:
                from tkinter import messagebox
                messagebox.showwarning("Tambah Lampiran", "Fungsi tambah lampiran belum diimplementasikan di parent.")
        btn_tambah = ttk.Button(action_attach_frame, text="➕ Tambah Lampiran", command=on_add_pdf_attachment, style="Secondary.TButton")
        btn_tambah.grid(row=0, column=0, padx=2, sticky="ew")

        def remove_selected_attachment():
            # Try to find a Listbox in the parent or listbox_frame
            listbox = None
            for child in listbox_frame.winfo_children():
                if isinstance(child, tk.Listbox):
                    listbox = child
                    break
            if listbox:
                selected = listbox.curselection()
                if selected:
                    idx = selected[0]
                    on_remove_pdf(idx)
                else:
                    from tkinter import messagebox
                    messagebox.showwarning("Hapus Lampiran", "Pilih file yang akan dihapus terlebih dahulu.")
            else:
                from tkinter import messagebox
                messagebox.showwarning("Hapus Lampiran", "Listbox lampiran tidak ditemukan.")

        btn_hapus = ttk.Button(action_attach_frame, text="🗑 Hapus Lampiran", command=remove_selected_attachment, style="Danger.TButton")
        btn_hapus.grid(row=0, column=1, padx=2, sticky="ew")

        def on_clear_form():
            func = callbacks.get("clear_form")
            if callable(func):
                func()
            else:
                from tkinter import messagebox
                messagebox.showwarning("Bersihkan Form", "Fungsi bersihkan form belum diimplementasikan di parent.")
        btn_bersihkan = ttk.Button(action_attach_frame, text="🔄 Bersihkan Form", command=on_clear_form, style="Secondary.TButton")
        btn_bersihkan.grid(row=0, column=2, padx=2, sticky="ew")

    # Selesai button section
    selesai_frame = ttk.Frame(button_frame)
    selesai_frame.grid(row=1, column=0, sticky="ew")
    selesai_frame.columnconfigure(0, weight=1)

    # Selesai button
    btn_selesai = ttk.Button(selesai_frame, text="✅ Selesai", command=callbacks.get("on_selesai"), width=22, style="Success.TButton")
    btn_selesai.grid(row=0, column=0, pady=8, sticky="ew")

    # Preview & options (hidden until selesai pressed)
    preview_frame = ttk.LabelFrame(button_frame, text="Preview & Export", padding="10")
    preview_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    preview_frame.columnconfigure(0, weight=1)
    preview_frame.grid_remove()  # Hide initially
    parent.preview_frame = preview_frame  # For access in callbacks

    # PDF preview area (to be filled by callback)
    parent.pdf_preview_label = ttk.Label(preview_frame, text="[Preview PDF akan muncul di sini]", font=("Segoe UI", 9), foreground="#64748b")
    parent.pdf_preview_label.grid(row=0, column=0, sticky="ew", pady=(0, 8))

    # Export options
    export_options_frame = ttk.Frame(preview_frame)
    export_options_frame.grid(row=1, column=0, sticky="ew")
    export_options_frame.columnconfigure(0, weight=1)

    parent.chk_export_pdf = tk.BooleanVar(value=True)
    parent.chk_upload_sheet = tk.BooleanVar(value=False)
    parent.chk_send_email = tk.BooleanVar(value=False)
    chk_pdf = ttk.Checkbutton(export_options_frame, text="Export PDF", variable=parent.chk_export_pdf)
    chk_pdf.grid(row=0, column=0, sticky="w", padx=4)
    chk_sheet = ttk.Checkbutton(export_options_frame, text="Upload ke Sheet", variable=parent.chk_upload_sheet)
    chk_sheet.grid(row=0, column=1, sticky="w", padx=4)
    chk_email = ttk.Checkbutton(export_options_frame, text="Kirim Email", variable=parent.chk_send_email)
    chk_email.grid(row=0, column=2, sticky="w", padx=4)

    # Sub-checkbox posisi untuk email
    parent.email_posisi_vars = {}
    posisi = getattr(parent, 'posisi_options', ["Direktur Utama", "Direktur Keuangan", "Direktur Teknik", "GM Keuangan & Administrasi", "GM Operasional & Pemeliharaan", "Manager"])
    sub_frame = ttk.Frame(export_options_frame)
    sub_frame.grid(row=1, column=2, sticky="w", padx=4)
    for i, pos in enumerate(posisi):
        var = tk.BooleanVar(value=False)
        parent.email_posisi_vars[pos] = var
        chk = ttk.Checkbutton(sub_frame, text=pos, variable=var)
        chk.grid(row=0, column=i, sticky="w", padx=2)
    sub_frame.grid_remove()  # Hide initially
    parent.sub_email_frame = sub_frame

    def on_email_toggle():
        if parent.chk_send_email.get():
            sub_frame.grid()
        else:
            sub_frame.grid_remove()
    parent.chk_send_email.trace_add('write', lambda *args: on_email_toggle())

    # Action buttons for export
    action_frame = ttk.Frame(preview_frame)
    action_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
    action_frame.columnconfigure((0, 1), weight=1)
    btn_export_pdf = ttk.Button(action_frame, text="Export PDF", command=callbacks.get("export_pdf"), style="Primary.TButton")
    btn_export_pdf.grid(row=0, column=0, padx=(0, 4), sticky="ew")
    btn_upload_sheet = ttk.Button(action_frame, text="Upload ke Sheet", command=callbacks.get("upload_sheet"), style="Primary.TButton")
    btn_upload_sheet.grid(row=0, column=1, padx=(4, 0), sticky="ew")

    # Show preview_frame when selesai pressed
    def show_preview():
        preview_frame.grid()
        # Optionally, update preview label with PDF preview
        parent.pdf_preview_label.config(text="[Preview PDF akan muncul di sini]")
    callbacks["show_preview"] = show_preview
    btn_selesai.config(command=show_preview)

    return button_frame
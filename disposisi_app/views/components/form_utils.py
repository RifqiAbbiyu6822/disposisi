def clear_form(self):
    try:
        for key, var in self.vars.items():
            if hasattr(var, 'set'):
                if isinstance(var.get(), str):
                    var.set("")
                else:
                    var.set(0)
        for key in ["perihal", "asal_surat", "ditujukan", "bicarakan_dengan", "teruskan_kepada"]:
            if key in self.form_input_widgets:
                self.form_input_widgets[key].delete("1.0", "end")
        # Perbaikan akses field tanggal
        for date_key in ["tgl_surat", "tgl_terima", "harap_selesai_tgl"]:
            if date_key in self.form_input_widgets:
                self.form_input_widgets[date_key].delete(0, "end")
        if hasattr(self, "instruksi_table"):
            self.instruksi_table.clear()
        self.edit_mode = False
        self.pdf_attachments = []
        if hasattr(self, 'refresh_pdf_attachments'):
            self.refresh_pdf_attachments(self._form_main_frame)
    except Exception as e:
        import traceback
        traceback.print_exc()
        from tkinter import messagebox
        messagebox.showerror("Error", f"Gagal mengosongkan form: {e}")

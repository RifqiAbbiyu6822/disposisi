import tkinter as tk

class DisposisiForm:
    def __init__(self, parent):
        self.parent = parent
        # Form variables dictionary to store all form fields
        self.form_vars = {}
        
    def get_form_data(self):
        """
        Get all form data needed for email
        Returns dictionary with all necessary data for email
        """
        data = {
            'nomor_surat': self.form_vars.get('no_surat', tk.StringVar()).get(),
            'nama_pengirim': self.form_vars.get('asal_surat', tk.StringVar()).get(),
            'perihal': self.form_vars.get('perihal', tk.StringVar()).get(),
            'instruksi': self.get_instruksi_text()
        }
        return data
        
    def get_instruksi_text(self):
        """
        Combine all instructions and disposisi selections into one text
        """
        instruksi_parts = []
        
        # Add klasifikasi if selected
        klasifikasi = []
        if self.form_vars.get('rahasia', tk.BooleanVar()).get():
            klasifikasi.append("RAHASIA")
        if self.form_vars.get('penting', tk.BooleanVar()).get():
            klasifikasi.append("PENTING")
        if self.form_vars.get('segera', tk.BooleanVar()).get():
            klasifikasi.append("SEGERA")
        
        if klasifikasi:
            instruksi_parts.append("Klasifikasi: " + ", ".join(klasifikasi))
        
        # Add selected disposisi options
        disposisi = []
        for key, var in self.form_vars.items():
            if key.startswith('disposisi_') and var.get():
                disposisi.append(key.replace('disposisi_', ''))
        
        if disposisi:
            instruksi_parts.append("Disposisi: " + ", ".join(disposisi))
        
        # Add custom instruksi text if any
        instruksi_text = self.form_vars.get('instruksi', tk.StringVar()).get().strip()
        if instruksi_text:
            instruksi_parts.append("Catatan: " + instruksi_text)
            
        return "\n".join(instruksi_parts)

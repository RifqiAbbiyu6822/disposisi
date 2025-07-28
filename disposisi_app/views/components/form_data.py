def get_form_data(self):
    """Get all form data needed for email"""
    data = {
        'nomor_surat': self.form_vars.get('no_surat', '').get(),
        'nama_pengirim': self.form_vars.get('asal_surat', '').get(),
        'perihal': self.form_vars.get('perihal', '').get(),
        'instruksi': self.get_instruksi_text()  # Implement this method to get combined instruksi
    }
    return data

def get_instruksi_text(self):
    """Get combined instruksi text from all selected options"""
    instruksi_list = []
    
    # Add selected disposisi labels
    for key, var in self.form_vars.items():
        if key.startswith('disposisi_') and var.get():
            instruksi_list.append(key.replace('disposisi_', ''))
    
    # Add instruksi text if any
    instruksi_text = self.form_vars.get('instruksi', '').get().strip()
    if instruksi_text:
        instruksi_list.append(instruksi_text)
    
    return "\n".join(instruksi_list)

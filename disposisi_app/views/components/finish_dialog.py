def _process(self):
    """Process selected options"""
    try:
        # Get selected positions for email
        selected_positions = [
            label for label, var in self.email_vars.items()
            if var.get()
        ]
        
        # Save PDF if selected
        if self.save_pdf_var.get() and callable(self.callbacks.get("save_pdf")):
            self.callbacks["save_pdf"]()
        
        # Save to Sheet if selected  
        if self.save_sheet_var.get() and callable(self.callbacks.get("save_sheet")):
            self.callbacks["save_sheet"]()
        
        # Send emails if selected
        if self.send_email_var.get() and selected_positions:
            from email_sender.send_email import EmailSender
            email_sender = EmailSender()
            
            # Get form data
            form_data = {}
            if hasattr(self.parent, 'vars'):
                # Get basic form data
                form_data['nomor_surat'] = self.parent.vars.get('no_surat', tk.StringVar()).get()
                form_data['nama_pengirim'] = self.parent.vars.get('asal_surat', tk.StringVar()).get() 
                form_data['perihal'] = self.parent.vars.get('perihal', tk.StringVar()).get()
                
                # Get instruksi from instruksi table
                instruksi_text = ""
                if hasattr(self.parent, 'instruksi_table'):
                    instruksi_data = self.parent.instruksi_table.get_data()
                    for item in instruksi_data:
                        if item.get('instruksi'):
                            instruksi_text += f"{item['posisi']}: {item['instruksi']}\n"
                
                form_data['instruksi'] = instruksi_text or "Mohon ditindaklanjuti sesuai disposisi"
            
            # Send emails to selected positions
            success_count = 0
            failed_positions = []
            
            for position in selected_positions:
                try:
                    status, message = email_sender.send_disposisi_email(
                        position=position,
                        nama_pengirim=form_data.get('nama_pengirim', ''),
                        nomor_surat=form_data.get('nomor_surat', ''),
                        perihal=form_data.get('perihal', ''),
                        instruksi=form_data.get('instruksi', '')
                    )
                    
                    if status:
                        success_count += 1
                    else:
                        failed_positions.append(f"{position}: {message}")
                        
                except Exception as e:
                    failed_positions.append(f"{position}: {str(e)}")
            
            # Show results
            if success_count == len(selected_positions):
                messagebox.showinfo("Sukses", 
                    f"Email berhasil dikirim ke {success_count} penerima")
            elif success_count > 0:
                error_msg = "\n".join(failed_positions)
                messagebox.showwarning("Sebagian Berhasil",
                    f"Email berhasil dikirim ke {success_count} penerima.\n\n"
                    f"Gagal mengirim ke:\n{error_msg}")
            else:
                error_msg = "\n".join(failed_positions)
                messagebox.showerror("Gagal", 
                    f"Gagal mengirim email:\n{error_msg}")
        
        self.destroy()
        
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")
        import traceback
        traceback.print_exc()
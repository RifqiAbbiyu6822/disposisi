#!/usr/bin/env python3
"""
Script to update all messagebox calls to use LoadingMessageBox with icon support
"""

import os
import re

def update_messagebox_calls():
    """Replace all messagebox calls with LoadingMessageBox calls"""
    
    # Files to update
    files_to_update = [
        'disposisi_app/views/components/email_manager.py',
        'disposisi_app/views/components/email_error_handler.py',
        'disposisi_app/views/components/finish_dialog.py',
        'disposisi_app/views/components/form_utils.py',
        'disposisi_app/views/components/dialogs.py',
        'disposisi_app/views/components/button_frame.py',
        'admin/main.py',
        'coba.py',
        'pdf_output.py'
    ]
    
    for file_path in files_to_update:
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            continue
            
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            original_content = content
            
            # Replace messagebox imports
            content = re.sub(
                r'from tkinter import messagebox',
                'from disposisi_app.views.components.loading_screen import LoadingMessageBox',
                content
            )
            content = re.sub(
                r'import messagebox',
                'from disposisi_app.views.components.loading_screen import LoadingMessageBox',
                content
            )
            
            # Replace messagebox calls with parent parameter
            content = re.sub(
                r'messagebox\.showinfo\(([^)]+)\)',
                r'LoadingMessageBox.showinfo(\1, parent=self)',
                content
            )
            content = re.sub(
                r'messagebox\.showwarning\(([^)]+)\)',
                r'LoadingMessageBox.showwarning(\1, parent=self)',
                content
            )
            content = re.sub(
                r'messagebox\.showerror\(([^)]+)\)',
                r'LoadingMessageBox.showerror(\1, parent=self)',
                content
            )
            content = re.sub(
                r'messagebox\.askyesno\(([^)]+)\)',
                r'LoadingMessageBox.askyesno(\1, parent=self)',
                content
            )
            content = re.sub(
                r'messagebox\.askokcancel\(([^)]+)\)',
                r'LoadingMessageBox.askokcancel(\1, parent=self)',
                content
            )
            
            # Also handle cases where parent is already specified
            content = re.sub(
                r'messagebox\.showinfo\(([^)]+), parent=([^)]+)\)',
                r'LoadingMessageBox.showinfo(\1, parent=\2)',
                content
            )
            content = re.sub(
                r'messagebox\.showwarning\(([^)]+), parent=([^)]+)\)',
                r'LoadingMessageBox.showwarning(\1, parent=\2)',
                content
            )
            content = re.sub(
                r'messagebox\.showerror\(([^)]+), parent=([^)]+)\)',
                r'LoadingMessageBox.showerror(\1, parent=\2)',
                content
            )
            content = re.sub(
                r'messagebox\.askyesno\(([^)]+), parent=([^)]+)\)',
                r'LoadingMessageBox.askyesno(\1, parent=\2)',
                content
            )
            content = re.sub(
                r'messagebox\.askokcancel\(([^)]+), parent=([^)]+)\)',
                r'LoadingMessageBox.askokcancel(\1, parent=\2)',
                content
            )
            
            if content != original_content:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                print(f"Updated: {file_path}")
            else:
                print(f"No changes needed: {file_path}")
                
        except Exception as e:
            print(f"Error updating {file_path}: {e}")

if __name__ == "__main__":
    print("Updating messagebox calls to use LoadingMessageBox...")
    update_messagebox_calls()
    print("Done!") 
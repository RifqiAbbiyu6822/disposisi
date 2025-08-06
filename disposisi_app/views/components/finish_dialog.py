import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import traceback

# Import email sender dengan modifikasi untuk mengambil dari spreadsheet
from email_sender.send_email import EmailSender
from email_sender.template_handler import render_email_template

class FinishDialog(tk.Toplevel):
    def __init__(self, parent, disposisi_labels, callbacks):
        super().__init__(parent)
        self.transient(parent)
        self.parent = parent
        self.disposisi_labels = disposisi_labels
        self.callbacks = callbacks
        self.title("Finalisasi Disposisi")
        self.geometry("600x600")  # Increased height for senior officers
        self.grab_set()
        
        # ENHANCED: Track senior officer selections
        self.senior_officer_vars = {}
        self.senior_officer_frames = {}
        
        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=(20, 20))
        main_frame.pack(fill="both", expand=True)
        title_label = ttk.Label(main_frame, text="Pilih Tindakan Selanjutnya", font=("Segoe UI", 16, "bold"))
        title_label.pack(pady=(0, 20))

        actions_frame = ttk.LabelFrame(main_frame, text="Opsi", padding=15)
        actions_frame.pack(fill="x", pady=(0, 15))

        self.save_pdf_var = tk.BooleanVar(value=True)
        self.save_sheet_var = tk.BooleanVar(value=True)
        self.send_email_var = tk.BooleanVar(value=False)

        ttk.Checkbutton(actions_frame, text="📄 Simpan sebagai PDF", variable=self.save_pdf_var).pack(anchor="w")
        ttk.Checkbutton(actions_frame, text="📊 Unggah ke Google Sheets", variable=self.save_sheet_var).pack(anchor="w")
        ttk.Checkbutton(actions_frame, text="📧 Kirim Disposisi via Email", variable=self.send_email_var, command=self._toggle_email_frame).pack(anchor="w", pady=(10,0))
        
        # ENHANCED: Email recipients frame with scrollable area
        self.email_container = ttk.Frame(main_frame)
        
        # Create scrollable frame for email recipients
        self.email_canvas = tk.Canvas(self.email_container, height=300)
        self.email_scrollbar = ttk.Scrollbar(self.email_container, orient="vertical", command=self.email_canvas.yview)
        self.email_scrollable_frame = ttk.Frame(self.email_canvas)
        
        self.email_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.email_canvas.configure(scrollregion=self.email_canvas.bbox("all"))
        )
        
        self.email_canvas.create_window((0, 0), window=self.email_scrollable_frame, anchor="nw")
        self.email_canvas.configure(yscrollcommand=self.email_scrollbar.set)
        
        self.email_canvas.pack(side="left", fill="both", expand=True)
        self.email_scrollbar.pack(side="right", fill="y")
        
        # Create email recipients frame
        self._create_email_recipients_frame()
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        ttk.Button(button_frame, text="Batal", command=self.destroy).grid(row=0, column=0, padx=5, sticky="ew")
        ttk.Button(button_frame, text="✅ Proses", command=self._process).grid(row=0, column=1, padx=5, sticky="ew")

    def _create_email_recipients_frame(self):
        """Create email recipients selection with senior officer support"""
        # Clear existing widgets
        for widget in self.email_scrollable_frame.winfo_children():
            widget.destroy()
        
        self.email_vars = {}
        self.senior_officer_vars = {}
        self.senior_officer_frames = {}
        
        if not self.disposisi_labels:
            ttk.Label(self.email_scrollable_frame, text="Tidak ada posisi yang dipilih.").pack()
            return
        
        # Create header
        header_frame = ttk.LabelFrame(self.email_scrollable_frame, text="Pilih Penerima Email", padding=15)
        header_frame.pack(fill="x", pady=(0, 10))
        
        # Mapping untuk singkatan manager
        abbreviation_map = {
            "Manager Pemeliharaan": "pml",
            "Manager Operasional": "ops", 
            "Manager Administrasi": "adm",
            "Manager Keuangan": "keu"
        }
        
        # Senior officer mapping
        senior_officer_map = {
            "Manager Pemeliharaan": [
                "Senior Officer Pemeliharaan 1",
                "Senior Officer Pemeliharaan 2"
            ],
            "Manager Operasional": [
                "Senior Officer Operasional 1", 
                "Senior Officer Operasional 2"
            ],
            "Manager Administrasi": [
                "Senior Officer Administrasi 1",
                "Senior Officer Administrasi 2"
            ],
            "Manager Keuangan": [
                "Senior Officer Keuangan 1",
                "Senior Officer Keuangan 2"
            ]
        }
        
        row_counter = 0
        
        for label in self.disposisi_labels:
            # Main position frame
            position_frame = ttk.Frame(header_frame)
            position_frame.grid(row=row_counter, column=0, sticky="ew", pady=5)
            header_frame.grid_rowconfigure(row_counter, weight=0)
            
            var = tk.BooleanVar(value=True)
            
            # Use abbreviation for display
            display_label = abbreviation_map.get(label, label)
            if display_label in ["pml", "ops", "adm", "keu"]:
                display_label = f"Manager {display_label}"
            
            self.email_vars[label] = var  # Keep original label for email lookup
            
            # Main checkbox for the position
            main_cb = ttk.Checkbutton(position_frame, text=display_label, variable=var)
            main_cb.grid(row=0, column=0, sticky="w")
            
            row_counter += 1
            
            # ENHANCED: Add senior officers if this is a manager position
            if label in senior_officer_map:
                # Create collapsible frame for senior officers
                senior_frame = ttk.LabelFrame(header_frame, text=f"Senior Officers - {display_label}", padding=10)
                senior_frame.grid(row=row_counter, column=0, sticky="ew", padx=(20, 0), pady=(0, 10))
                header_frame.grid_rowconfigure(row_counter, weight=0)
                
                # Track senior officer frame
                self.senior_officer_frames[label] = senior_frame
                self.senior_officer_vars[label] = {}
                
                # Add senior officer checkboxes
                for i, senior_officer in enumerate(senior_officer_map[label]):
                    senior_var = tk.BooleanVar(value=False)  # Default unchecked
                    self.senior_officer_vars[label][senior_officer] = senior_var
                    
                    senior_cb = ttk.Checkbutton(
                        senior_frame, 
                        text=senior_officer, 
                        variable=senior_var
                    )
                    senior_cb.grid(row=i, column=0, sticky="w", padx=(10, 0))
                
                # Add "Select All" functionality for senior officers
                def create_select_all_handler(manager_label):
                    def select_all_seniors():
                        all_selected = all(var.get() for var in self.senior_officer_vars[manager_label].values())
                        new_state = not all_selected
                        for var in self.senior_officer_vars[manager_label].values():
                            var.set(new_state)
                    return select_all_seniors
                
                select_all_btn = ttk.Button(
                    senior_frame, 
                    text="Toggle All", 
                    command=create_select_all_handler(label),
                    style="Secondary.TButton"
                )
                select_all_btn.grid(row=len(senior_officer_map[label]), column=0, sticky="w", padx=(10, 0), pady=(5, 0))
                
                row_counter += 1
                
                # Initially hide senior officer frame
                senior_frame.grid_remove()
                
                # Bind main checkbox to show/hide senior officers
                def create_toggle_handler(manager_label, frame):
                    def toggle_senior_officers():
                        if self.email_vars[manager_label].get():
                            frame.grid()
                        else:
                            frame.grid_remove()
                            # Uncheck all senior officers when manager is unchecked
                            for senior_var in self.senior_officer_vars[manager_label].values():
                                senior_var.set(False)
                    return toggle_senior_officers
                
                var.trace('w', lambda *args, label=label, frame=senior_frame: create_toggle_handler(label, frame)())
        
        header_frame.grid_columnconfigure(0, weight=1)

    def _toggle_email_frame(self):
        if self.send_email_var.get():
            self.email_container.pack(fill="both", expand=True, pady=(0, 15))
        else:
            self.email_container.pack_forget()

    def _get_selected_recipients(self):
        """Get all selected email recipients including senior officers"""
        selected_positions = []
        
        # Get main positions
        for label, var in self.email_vars.items():
            if var.get():
                selected_positions.append(label)
        
        # Get selected senior officers
        for manager_label, senior_vars in self.senior_officer_vars.items():
            if self.email_vars[manager_label].get():  # Only if manager is selected
                for senior_officer, var in senior_vars.items():
                    if var.get():
                        selected_positions.append(senior_officer)
        
        return selected_positions

    def _process(self):
        try:
            # Process save PDF if selected
            if self.save_pdf_var.get() and callable(self.callbacks.get("save_pdf")):
                self.callbacks["save_pdf"]()
            
            # Process save to Sheet if selected
            if self.save_sheet_var.get() and callable(self.callbacks.get("save_sheet")):
                self.callbacks["save_sheet"]()
            
            # Process send email if selected
            if self.send_email_var.get():
                selected_positions = self._get_selected_recipients()
                if not selected_positions:
                    messagebox.showwarning("Peringatan", "Pilih setidaknya satu penerima email.", parent=self)
                    return
                
                # Show confirmation with all selected recipients
                abbreviation_map = {
                    "Manager Pemeliharaan": "Manager pml",
                    "Manager Operasional": "Manager ops",
                    "Manager Administrasi": "Manager adm", 
                    "Manager Keuangan": "Manager keu"
                }
                
                display_recipients = []
                for pos in selected_positions:
                    display_name = abbreviation_map.get(pos, pos)
                    display_recipients.append(display_name)
                
                confirm_msg = f"Kirim email ke {len(selected_positions)} penerima:\n\n"
                confirm_msg += "\n".join([f"• {name}" for name in display_recipients])
                confirm_msg += "\n\nLanjutkan pengiriman?"
                
                if not messagebox.askyesno("Konfirmasi Email", confirm_msg, parent=self):
                    return
                
                # Call the main callback from the parent app
                if callable(self.callbacks.get("send_email")):
                    self.callbacks["send_email"](selected_positions)
            
            self.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan: {e}", parent=self)
            traceback.print_exc()
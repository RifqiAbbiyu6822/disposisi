import tkinter as tk
from tkinter import ttk
from datetime import datetime
import traceback
import threading

# Import email sender dengan modifikasi untuk mengambil dari spreadsheet
from email_sender.send_email import EmailSender
from email_sender.template_handler import render_email_template
# Import loading screen
from .loading_screen import LoadingManager, LoadingMessageBox

class FinishDialog(tk.Toplevel):
    def __init__(self, parent, disposisi_labels, callbacks):
        super().__init__(parent)
        self.transient(parent)
        self.parent = parent
        self.disposisi_labels = disposisi_labels
        self.callbacks = callbacks
        self.title("Finalisasi Disposisi")
        self.geometry("650x700")  # Slightly wider for better proportions
        self.grab_set()
        
        # Initialize loading manager
        self.loading_manager = LoadingManager()
        
        # Center the window
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (650 // 2)
        y = (self.winfo_screenheight() // 2) - (700 // 2)
        self.geometry(f"650x700+{x}+{y}")
        
        # Configure window style
        self.configure(bg='#f0f0f0')
        
        # ENHANCED: Track senior officer selections
        self.senior_officer_vars = {}
        self.senior_officer_frames = {}
        
        self._create_widgets()

    def _create_widgets(self):
        # Main container with better padding
        main_container = ttk.Frame(self, style="Card.TFrame")
        main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        main_frame = ttk.Frame(main_container, padding=(25, 20))
        main_frame.pack(fill="both", expand=True)
        
        # Title with icon-like styling
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 25))
        
        title_label = ttk.Label(
            title_frame, 
            text="✓ Finalisasi Disposisi", 
            font=("Segoe UI", 18, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack(side="left")
        
        # Separator line
        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=(0, 20))

        # Actions frame with improved styling
        actions_frame = ttk.LabelFrame(
            main_frame, 
            text=" Pilih Tindakan ", 
            padding=20,
            style="Card.TLabelframe"
        )
        actions_frame.pack(fill="x", pady=(0, 20))

        self.save_pdf_var = tk.BooleanVar(value=True)
        self.save_sheet_var = tk.BooleanVar(value=True)
        self.send_email_var = tk.BooleanVar(value=False)

        # Action checkboxes with better spacing and icons
        action_items = [
            ("📄 Simpan sebagai PDF", self.save_pdf_var, None),
            ("📊 Unggah ke Google Sheets", self.save_sheet_var, None),
            ("📧 Kirim Disposisi via Email", self.send_email_var, self._toggle_email_frame)
        ]
        
        for i, (text, var, command) in enumerate(action_items):
            cb_frame = ttk.Frame(actions_frame)
            cb_frame.pack(fill="x", pady=(0, 8) if i < len(action_items)-1 else 0)
            
            cb = ttk.Checkbutton(
                cb_frame, 
                text=text, 
                variable=var,
                command=command,
                style="Action.TCheckbutton"
            )
            cb.pack(side="left", padx=(5, 0))
            
            # Add separator after second item
            if i == 1:
                ttk.Separator(actions_frame, orient="horizontal").pack(fill="x", pady=(12, 12))
        
        # ENHANCED: Email recipients frame with scrollable area
        self.email_container = ttk.LabelFrame(
            main_frame,
            text=" Penerima Email ",
            padding=(15, 10),
            style="Email.TLabelframe"
        )
        
        # Inner container for better organization
        email_inner = ttk.Frame(self.email_container)
        email_inner.pack(fill="both", expand=True)
        
        # Create scrollable frame for email recipients
        self.email_canvas = tk.Canvas(
            email_inner, 
            height=280,
            bg='white',
            highlightthickness=1,
            highlightbackground='#ddd'
        )
        self.email_scrollbar = ttk.Scrollbar(
            email_inner, 
            orient="vertical", 
            command=self.email_canvas.yview
        )
        self.email_scrollable_frame = ttk.Frame(self.email_canvas, style="White.TFrame")
        
        self.email_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.email_canvas.configure(scrollregion=self.email_canvas.bbox("all"))
        )
        
        self.email_canvas.create_window((0, 0), window=self.email_scrollable_frame, anchor="nw")
        self.email_canvas.configure(yscrollcommand=self.email_scrollbar.set)
        
        # Enable mouse wheel scrolling
        def _on_mousewheel(event):
            try:
                if hasattr(self, 'email_canvas') and self.email_canvas.winfo_exists():
                    self.email_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except tk.TclError:
                # Widget sudah dihancurkan, unbind event
                try:
                    self.unbind_all("<MouseWheel>")
                except:
                    pass
        
        self.email_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        self.email_canvas.pack(side="left", fill="both", expand=True)
        self.email_scrollbar.pack(side="right", fill="y", padx=(2, 0))
        
        # Create email recipients frame
        self._create_email_recipients_frame()
        
        # Separator before buttons
        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=(20, 15))
        
        # Button frame with improved styling
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x")
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        
        cancel_btn = ttk.Button(
            button_frame, 
            text="✕ Batal", 
            command=self.destroy,
            style="Cancel.TButton"
        )
        cancel_btn.grid(row=0, column=0, padx=(0, 10), sticky="ew", ipady=5)
        
        process_btn = ttk.Button(
            button_frame, 
            text="✓ Proses", 
            command=self._process,
            style="Process.TButton"
        )
        process_btn.grid(row=0, column=1, padx=(10, 0), sticky="ew", ipady=5)

    def destroy(self):
        """Clean up resources before destroying the dialog"""
        try:
            # Unbind mouse wheel events
            self.unbind_all("<MouseWheel>")
        except:
            pass
        
        # Hide loading screen if it's showing
        try:
            if hasattr(self, 'loading_manager'):
                self.loading_manager.hide_loading()
        except:
            pass
        
        super().destroy()

    def _create_email_recipients_frame(self):
        """Create email recipients selection with senior officer support"""
        # Clear existing widgets
        for widget in self.email_scrollable_frame.winfo_children():
            widget.destroy()
        
        self.email_vars = {}
        self.senior_officer_vars = {}
        self.senior_officer_frames = {}
        self.dropdown_buttons = {}  # Track dropdown buttons for state management
        
        if not self.disposisi_labels:
            no_data_frame = ttk.Frame(self.email_scrollable_frame)
            no_data_frame.pack(expand=True, fill="both", pady=50)
            ttk.Label(
                no_data_frame, 
                text="Tidak ada posisi yang dipilih.",
                font=("Segoe UI", 10, "italic"),
                foreground="#999"
            ).pack()
            return
        
        # Create container with padding
        container = ttk.Frame(self.email_scrollable_frame, padding=10)
        container.pack(fill="both", expand=True)
        
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
        
        for idx, label in enumerate(self.disposisi_labels):
            # Manager container frame with border
            manager_container = ttk.Frame(container, relief="solid", borderwidth=1)
            manager_container.pack(fill="x", pady=(0, 10) if idx < len(self.disposisi_labels)-1 else 0)
            
            # Main position frame with full width layout
            position_frame = ttk.Frame(manager_container, padding=(12, 10))
            position_frame.pack(fill="x")
            position_frame.columnconfigure(1, weight=1)  # Make middle column expand
            
            var = tk.BooleanVar(value=True)
            
            # Use abbreviation for display
            display_label = abbreviation_map.get(label, label)
            if display_label in ["pml", "ops", "adm", "keu"]:
                display_label = f"Manager {display_label.upper()}"
            
            self.email_vars[label] = var  # Keep original label for email lookup
            
            # Try to get name from admin sheet
            try:
                from email_sender.send_email import EmailSender
                email_sender = EmailSender()
                name, _ = email_sender.get_recipient_name(label)
                if name:
                    display_text = f"📋 {display_label} - {name}"
                else:
                    display_text = f"📋 {display_label}"
            except:
                display_text = f"📋 {display_label}"
            
            # Main checkbox on the left
            main_cb = ttk.Checkbutton(
                position_frame, 
                text=display_text, 
                variable=var,
                style="Manager.TCheckbutton"
            )
            main_cb.grid(row=0, column=0, sticky="w")
            
            # ENHANCED: Add senior officers if this is a manager position
            if label in senior_officer_map:
                # Empty space in middle
                ttk.Frame(position_frame).grid(row=0, column=1, sticky="ew")
                
                # Dropdown button on the right
                dropdown_btn = ttk.Button(
                    position_frame,
                    text="▼ Tim Senior Officers",
                    style="Dropdown.TButton",
                    width=20
                )
                dropdown_btn.grid(row=0, column=2, sticky="e", padx=(5, 0))
                self.dropdown_buttons[label] = dropdown_btn
                
                # Create collapsible frame for senior officers
                senior_container = ttk.Frame(manager_container)
                senior_container.pack(fill="x")
                
                senior_frame = ttk.Frame(
                    senior_container,
                    relief="ridge",
                    borderwidth=1
                )
                senior_frame.pack(fill="x", padx=(0, 0), pady=(0, 0))
                
                # Inner frame with padding
                senior_inner = ttk.Frame(senior_frame, padding=(40, 15, 15, 15))
                senior_inner.pack(fill="x")
                
                # Track senior officer frame
                self.senior_officer_frames[label] = senior_frame
                self.senior_officer_vars[label] = {}
                
                # Title for senior officers section
                title_frame = ttk.Frame(senior_inner)
                title_frame.pack(fill="x", pady=(0, 10))
                ttk.Label(
                    title_frame,
                    text="Tim Senior Officers",
                    font=("Segoe UI", 10, "bold"),
                    foreground="#4a90e2"
                ).pack(anchor="w")
                
                # Create a grid for senior officers
                officers_grid = ttk.Frame(senior_inner)
                officers_grid.pack(fill="x")
                
                # Add senior officer checkboxes
                for i, senior_officer in enumerate(senior_officer_map[label]):
                    senior_var = tk.BooleanVar(value=False)  # Default unchecked
                    self.senior_officer_vars[label][senior_officer] = senior_var
                    
                    officer_frame = ttk.Frame(officers_grid)
                    officer_frame.pack(fill="x", pady=3)
                    
                    # Try to get name for senior officer
                    try:
                        from email_sender.send_email import EmailSender
                        email_sender = EmailSender()
                        name, _ = email_sender.get_recipient_name(senior_officer)
                        if name:
                            display_text = f"👤 {senior_officer} - {name}"
                        else:
                            display_text = f"👤 {senior_officer}"
                    except:
                        display_text = f"👤 {senior_officer}"
                    
                    senior_cb = ttk.Checkbutton(
                        officer_frame, 
                        text=display_text, 
                        variable=senior_var,
                        style="Senior.TCheckbutton"
                    )
                    senior_cb.pack(anchor="w")
                
                # Add "Toggle All" functionality with better styling
                def create_select_all_handler(manager_label):
                    def select_all_seniors():
                        all_selected = all(var.get() for var in self.senior_officer_vars[manager_label].values())
                        new_state = not all_selected
                        for var in self.senior_officer_vars[manager_label].values():
                            var.set(new_state)
                    return select_all_seniors
                
                # Button container with separator
                ttk.Separator(senior_inner, orient="horizontal").pack(fill="x", pady=(10, 10))
                
                btn_frame = ttk.Frame(senior_inner)
                btn_frame.pack(fill="x")
                
                select_all_btn = ttk.Button(
                    btn_frame, 
                    text="☑ Toggle All", 
                    command=create_select_all_handler(label),
                    style="Toggle.TButton"
                )
                select_all_btn.pack(anchor="w")
                
                # Initially hide senior officer frame
                senior_frame.pack_forget()
                
                # Create dropdown toggle handler
                def create_dropdown_handler(manager_label, frame, button):
                    is_expanded = [False]  # Use list to maintain state in closure
                    
                    def toggle_dropdown():
                        if not self.email_vars[manager_label].get():
                            # If manager is unchecked, don't allow dropdown
                            return
                        
                        is_expanded[0] = not is_expanded[0]
                        if is_expanded[0]:
                            frame.pack(fill="x", padx=(0, 0), pady=(0, 0))
                            button.configure(text="▲ Tim Senior Officers")
                        else:
                            frame.pack_forget()
                            button.configure(text="▼ Tim Senior Officers")
                    
                    return toggle_dropdown, is_expanded
                
                # Bind dropdown button
                toggle_func, expanded_state = create_dropdown_handler(label, senior_frame, dropdown_btn)
                dropdown_btn.configure(command=toggle_func)
                
                # Bind main checkbox to enable/disable dropdown and reset state
                def create_checkbox_handler(manager_label, frame, button, expanded_state):
                    def handle_checkbox_change():
                        if self.email_vars[manager_label].get():
                            button.configure(state="normal")
                        else:
                            button.configure(state="disabled")
                            frame.pack_forget()
                            button.configure(text="▼ Tim Senior Officers")
                            expanded_state[0] = False
                            # Uncheck all senior officers when manager is unchecked
                            for senior_var in self.senior_officer_vars[manager_label].values():
                                senior_var.set(False)
                    return handle_checkbox_change
                
                var.trace('w', lambda *args, label=label, frame=senior_frame, btn=dropdown_btn, exp=expanded_state: 
                         create_checkbox_handler(label, frame, btn, exp)())

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
            # Show loading screen
            loading = self.loading_manager.show_loading(self, "Memproses Disposisi...", show_progress=True)
            
            # Process operations in thread to avoid blocking UI
            def process_operations():
                try:
                    self.current_progress = 0
                    total_operations = 0
                    
                    # Count total operations
                    if self.save_pdf_var.get():
                        total_operations += 1
                    if self.save_sheet_var.get():
                        total_operations += 1
                    if self.send_email_var.get():
                        total_operations += 1
                    
                    if total_operations == 0:
                        self.after(0, self.loading_manager.hide_loading)
                        self.after(100, self.destroy)
                        return
                    
                    # Process save PDF if selected
                    if self.save_pdf_var.get() and callable(self.callbacks.get("save_pdf")):
                        self.after(0, lambda: self.loading_manager.update_progress(self.current_progress, "Menyimpan PDF..."))
                        self.callbacks["save_pdf"]()
                        self.current_progress += (100 // total_operations)
                    
                    # Process save to Sheet if selected
                    if self.save_sheet_var.get() and callable(self.callbacks.get("save_sheet")):
                        self.after(0, lambda: self.loading_manager.update_progress(self.current_progress, "Mengunggah ke Google Sheets..."))
                        self.callbacks["save_sheet"]()
                        self.current_progress += (100 // total_operations)
                    
                    # Process send email if selected
                    if self.send_email_var.get():
                        selected_positions = self._get_selected_recipients()
                        if not selected_positions:
                            self.after(0, self.loading_manager.hide_loading)
                            self.after(100, lambda: LoadingMessageBox.showwarning("Peringatan", "Pilih setidaknya satu penerima email.", parent=self))
                            return
                        
                        # Show confirmation with all selected recipients
                        abbreviation_map = {
                            "Manager Pemeliharaan": "Manager pml",
                            "Manager Operasional": "Manager ops",
                            "Manager Administrasi": "Manager adm", 
                            "Manager Keuangan": "Manager keu"
                        }
                        
                        # Get names from email sender
                        from email_sender.send_email import EmailSender
                        email_sender = EmailSender()
                        
                        display_recipients = []
                        for pos in selected_positions:
                            # Try to get name from admin sheet
                            name, _ = email_sender.get_recipient_name(pos)
                            if name:
                                display_name = f"{name} ({abbreviation_map.get(pos, pos)})"
                            else:
                                display_name = abbreviation_map.get(pos, pos)
                            display_recipients.append(display_name)
                        
                        confirm_msg = f"Kirim email ke {len(selected_positions)} penerima:\n\n"
                        confirm_msg += "\n".join([f"• {name}" for name in display_recipients])
                        confirm_msg += "\n\nLanjutkan pengiriman?"
                        
                        # Hide loading temporarily for confirmation
                        self.after(0, self.loading_manager.hide_loading)
                        
                        # Use after to schedule confirmation dialog
                        def show_confirmation():
                            if not LoadingMessageBox.askyesno("Konfirmasi Email", confirm_msg, parent=self):
                                self.after(100, self.destroy)
                                return
                            
                            # Show loading again for email sending
                            self.after(0, lambda: self.loading_manager.show_loading(self, "Mengirim Email...", show_progress=True))
                            
                            # Call the main callback from the parent app
                            if callable(self.callbacks.get("send_email")):
                                self.after(0, lambda: self.loading_manager.update_progress(self.current_progress, "Mengirim email..."))
                                self.callbacks["send_email"](selected_positions)
                                self.current_progress += (100 // total_operations)
                            
                            # Complete
                            self.after(0, lambda: self.loading_manager.update_progress(100, "Selesai!"))
                            
                            # Hide loading and close dialog
                            self.after(1000, self.loading_manager.hide_loading)
                            self.after(1100, self.destroy)
                        
                        self.after(100, show_confirmation)
                        return
                    
                    # Complete
                    self.after(0, lambda: self.loading_manager.update_progress(100, "Selesai!"))
                    
                    # Hide loading and close dialog
                    self.after(1000, self.loading_manager.hide_loading)
                    self.after(1100, self.destroy)
                    
                except Exception as e:
                    self.after(0, self.loading_manager.hide_loading)
                    self.after(100, lambda: LoadingMessageBox.showerror("Error", f"Terjadi kesalahan: {e}", parent=self))
                    traceback.print_exc()
            
            # Start processing in thread
            threading.Thread(target=process_operations, daemon=True).start()

        except Exception as e:
            self.loading_manager.hide_loading()
            LoadingMessageBox.showerror("Error", f"Terjadi kesalahan: {e}", parent=self)
            traceback.print_exc()
import gspread
import pandas as pd
import hashlib
from datetime import datetime
import json
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter import font as tkFont

class EmailManagerApp:
    def __init__(self):
        """Initialize the Enhanced Email Manager Application with Senior Officer support"""
        self.root = tk.Tk()
        self.root.title("Email Manager Dashboard - Enhanced")
        self.root.geometry("1400x900")  # Increased size for senior officers
        self.root.configure(bg='#f0f0f0')
        self.root.minsize(1000, 700)
        
        # Initialize variables
        self.logged_in = False
        self.sheet = None
        self.admin_password_hash = self.hash_password("admin123")  # Default password
        
        # ENHANCED: Define positions including senior officers
        self.positions = [
            "Direktur Utama",
            "Direktur Keuangan", 
            "Direktur Teknik",
            "GM Keuangan & Administrasi",
            "GM Operasional & Pemeliharaan",
            "Manager Pemeliharaan",
            "Manager Operasional",
            "Manager Administrasi",
            "Manager Keuangan",
            # ENHANCED: Senior Officers
            "Senior Officer Pemeliharaan 1",
            "Senior Officer Pemeliharaan 2",
            "Senior Officer Operasional 1",
            "Senior Officer Operasional 2",
            "Senior Officer Administrasi 1",
            "Senior Officer Administrasi 2",
            "Senior Officer Keuangan 1",
            "Senior Officer Keuangan 2"
        ]
        
        # Setup UI
        self.setup_styles()
        self.create_main_container()
        
        # Show login initially
        self.show_login()
    
    def setup_styles(self):
        """Setup custom styles for the application"""
        self.colors = {
            'primary': '#2c3e50',
            'secondary': '#3498db',
            'success': '#27ae60',
            'danger': '#e74c3c',
            'warning': '#f39c12',
            'light': '#ecf0f1',
            'dark': '#34495e',
            'white': '#ffffff',
            'senior_officer': '#9b59b6'  # ENHANCED: Special color for senior officers
        }
        self.style = ttk.Style()        
        self.style.theme_use('clam')
        
        # Configure styles (keeping existing styles)
        self.style.configure('Title.TLabel', 
                           font=('Arial', 24, 'bold'), 
                           background='#f0f0f0', 
                           foreground=self.colors['primary'])
        
        self.style.configure('Heading.TLabel', 
                           font=('Arial', 14, 'bold'), 
                           background='#f0f0f0', 
                           foreground=self.colors['dark'])
        
        # ENHANCED: Senior Officer specific style
        self.style.configure('SeniorOfficer.TLabel',
                           font=('Arial', 12, 'bold'),
                           background='#f0f0f0',
                           foreground=self.colors['senior_officer'])
        
        self.style.configure('Info.TLabel', 
                           font=('Arial', 11), 
                           background='#f0f0f0', 
                           foreground=self.colors['dark'])
        
        self.style.configure('Primary.TButton', 
                           font=('Arial', 10, 'bold'), 
                           padding=(20, 10))
        
        self.style.configure('Secondary.TButton', 
                           font=('Arial', 9), 
                           padding=(15, 8))
        
        # Configure Treeview styles
        self.style.configure("Treeview", 
                           background="white",
                           foreground="black",
                           rowheight=30,
                           fieldbackground="white")
        
        self.style.configure("Treeview.Heading", 
                           font=('Arial', 11, 'bold'),
                           background=self.colors['light'],
                           foreground=self.colors['dark'])
    
    def create_main_container(self):
        """Create the main application container"""
        # Header
        self.header_frame = tk.Frame(self.root, bg=self.colors['primary'], height=80)
        self.header_frame.pack(fill='x', side='top')
        self.header_frame.pack_propagate(False)
        
        # Navigation
        self.nav_frame = tk.Frame(self.root, bg=self.colors['dark'], height=50)
        self.nav_frame.pack(fill='x', side='top')
        self.nav_frame.pack_propagate(False)
        
        # Main content area
        self.content_frame = tk.Frame(self.root, bg='#f0f0f0')
        self.content_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
    def create_header(self, title="Email Manager Dashboard - Enhanced"):
        """Create header with enhanced title"""
        for widget in self.header_frame.winfo_children():
            widget.destroy()
            
        # Left side - Title
        left_frame = tk.Frame(self.header_frame, bg=self.colors['primary'])
        left_frame.pack(side='left', fill='y')
        
        title_label = tk.Label(left_frame, 
                              text=title,
                              font=('Arial', 20, 'bold'),
                              bg=self.colors['primary'],
                              fg='white')
        title_label.pack(side='left', padx=20, pady=20)
        
        # Right side - Enhanced user actions
        if self.logged_in:
            right_frame = tk.Frame(self.header_frame, bg=self.colors['primary'])
            right_frame.pack(side='right', fill='y')
            
            # Enhanced status indicator
            status_label = tk.Label(right_frame,
                                  text="● Connected (Enhanced)",
                                  font=('Arial', 10),
                                  bg=self.colors['primary'],
                                  fg=self.colors['success'])
            status_label.pack(side='right', padx=(0, 10), pady=25)
            
            # Position count indicator
            total_positions = len(self.positions)
            count_label = tk.Label(right_frame,
                                 text=f"📊 {total_positions} Positions",
                                 font=('Arial', 10),
                                 bg=self.colors['primary'],
                                 fg='white')
            count_label.pack(side='right', padx=(0, 15), pady=25)
            
            # Logout button
            logout_btn = tk.Button(right_frame,
                                 text="🚪 Logout",
                                 font=('Arial', 10, 'bold'),
                                 bg=self.colors['danger'],
                                 fg='white',
                                 border=0,
                                 padx=20,
                                 pady=8,
                                 cursor='hand2',
                                 relief='flat',
                                 command=self.logout)
            logout_btn.pack(side='right', padx=20, pady=20)
    
    def create_navigation(self):
        """Create enhanced navigation bar"""
        for widget in self.nav_frame.winfo_children():
            widget.destroy()
            
        if not self.logged_in:
            return
            
        nav_buttons = [
            ("📊 Dashboard", self.show_dashboard, 'dashboard'),
            ("📧 Kelola Email", self.show_email_management, 'email'),
            ("👨‍💼 Senior Officers", self.show_senior_officers, 'senior'),  # ENHANCED
            ("🔒 Ganti Password", self.show_change_password, 'password')
        ]
        
        self.nav_buttons = {}
        for text, command, key in nav_buttons:
            btn = tk.Button(self.nav_frame,
                          text=text,
                          font=('Arial', 10, 'bold'),
                          bg=self.colors['dark'],
                          fg='white',
                          border=0,
                          padx=20,
                          pady=10,
                          cursor='hand2',
                          relief='flat',
                          command=command)
            btn.pack(side='left', padx=2)
            self.nav_buttons[key] = btn
            
            # Add hover effects
            def make_hover_handler(button):
                def on_enter(e):
                    button.config(bg=self.colors['secondary'])
                def on_leave(e):
                    button.config(bg=self.colors['dark'])
                return on_enter, on_leave
            
            enter_handler, leave_handler = make_hover_handler(btn)
            btn.bind("<Enter>", enter_handler)
            btn.bind("<Leave>", leave_handler)
    
    def set_active_nav(self, active_key):
        """Set active navigation button"""
        if hasattr(self, 'nav_buttons'):
            for key, btn in self.nav_buttons.items():
                if key == active_key:
                    btn.config(bg=self.colors['secondary'])
                else:
                    btn.config(bg=self.colors['dark'])
            
    def clear_content(self):
        """Clear current content frame"""
        for widget in self.content_frame.winfo_children():
            widget.destroy()
            
    def show_login(self):
        """Show login screen"""
        self.create_header("Login - Email Manager")
        self.clear_content()
        
        # Hide navigation
        self.nav_frame.pack_forget()
        
        # Login container with shadow effect
        login_outer = tk.Frame(self.content_frame, bg='#f0f0f0')
        login_outer.pack(expand=True)
        
        # Shadow frame
        shadow_frame = tk.Frame(login_outer, bg='#d0d0d0', height=302, width=402)
        shadow_frame.pack(padx=(3, 0), pady=(3, 0))
        shadow_frame.pack_propagate(False)
        
        # Main login container
        login_container = tk.Frame(login_outer, bg='white', height=300, width=400, relief='solid', borderwidth=1)
        login_container.place(in_=shadow_frame, x=-3, y=-3)
        login_container.pack_propagate(False)
        
        # Login form
        form_frame = tk.Frame(login_container, bg='white')
        form_frame.pack(expand=True, fill='both', padx=40, pady=40)
        
        # Logo/Icon
        tk.Label(form_frame, 
                text="🔐",
                font=('Arial', 48),
                bg='white',
                fg=self.colors['primary']).pack(pady=(0, 10))
        
        tk.Label(form_frame, 
                text="Admin Login",
                font=('Arial', 18, 'bold'),
                bg='white',
                fg=self.colors['primary']).pack(pady=(0, 30))
        
        # Password field with improved styling
        tk.Label(form_frame,
                text="Password:",
                font=('Arial', 12, 'bold'),
                bg='white',
                fg=self.colors['dark']).pack(anchor='w', pady=(0, 5))
        
        password_frame = tk.Frame(form_frame, bg='white')
        password_frame.pack(fill='x', pady=(0, 20))
        
        self.password_entry = tk.Entry(password_frame,
                                     font=('Arial', 12),
                                     show="*",
                                     relief='solid',
                                     borderwidth=2,
                                     highlightthickness=0,
                                     bd=0)
        self.password_entry.pack(fill='x', ipady=8)
        self.password_entry.bind('<Return>', lambda e: self.try_login())
        self.password_entry.bind('<FocusIn>', lambda e: self.password_entry.config(relief='solid', borderwidth=2))
        self.password_entry.bind('<FocusOut>', lambda e: self.password_entry.config(relief='solid', borderwidth=1))
        
        # Login button with improved styling
        login_btn = tk.Button(form_frame,
                            text="🔓 Login",
                            font=('Arial', 12, 'bold'),
                            bg=self.colors['secondary'],
                            fg='white',
                            border=0,
                            relief='flat',
                            padx=30,
                            pady=12,
                            cursor='hand2',
                            command=self.try_login)
        login_btn.pack(pady=(10, 0))
        
        # Hover effect for login button
        def on_enter(e):
            login_btn.config(bg='#2980b9')
        def on_leave(e):
            login_btn.config(bg=self.colors['secondary'])
        
        login_btn.bind("<Enter>", on_enter)
        login_btn.bind("<Leave>", on_leave)
        
        # Focus on password entry
        self.password_entry.focus()
        
    def try_login(self):
        """Attempt to login"""
        password = self.password_entry.get()
        if self.verify_password(password, self.admin_password_hash):
            self.logged_in = True
            self.sheet = self.init_google_sheets()
            if self.sheet is None:
                messagebox.showerror("Error", "Tidak dapat terhubung ke Google Sheets.\nPastikan credentials/credentials.json tersedia dan valid.")
                return
            self.show_dashboard()
        else:
            messagebox.showerror("Login Gagal", "Password yang Anda masukkan salah!")
            self.password_entry.delete(0, 'end')
            self.password_entry.focus()
            
    def logout(self):
        """Logout user"""
        if messagebox.askyesno("Konfirmasi Logout", "Apakah Anda yakin ingin keluar?"):
            self.logged_in = False
            self.sheet = None
            self.show_login()
        
    def show_dashboard(self):
        """ENHANCED: Show main dashboard with senior officer statistics"""
        self.create_header("Dashboard Overview - Enhanced")
        self.nav_frame.pack(fill='x', side='top', after=self.header_frame)
        self.create_navigation()
        self.set_active_nav('dashboard')
        self.clear_content()
        
        # Dashboard content
        df = self.read_data(self.sheet)
        
        # Stats cards container
        stats_container = tk.Frame(self.content_frame, bg='#f0f0f0')
        stats_container.pack(fill='x', pady=(0, 20))
        
        if df.empty:
            # Initialize data if empty - ENHANCED with senior officers
            df = pd.DataFrame({
                'posisi': self.positions,
                'email': [''] * len(self.positions),
                'last_updated': [''] * len(self.positions)
            })
            self.write_data(self.sheet, df)
        
        total_positions = len(df)
        filled_emails = df['email'].apply(lambda x: x != '' and pd.notna(x)).sum() if 'email' in df.columns else 0
        empty_emails = total_positions - filled_emails
        completion_rate = (filled_emails / total_positions * 100) if total_positions > 0 else 0
        
        # ENHANCED: Separate statistics for different position types
        manager_positions = len([pos for pos in self.positions if pos.startswith("Manager")])
        senior_officer_positions = len([pos for pos in self.positions if pos.startswith("Senior Officer")])
        director_gm_positions = total_positions - manager_positions - senior_officer_positions
        
        # Create stat cards with enhanced design
        self.create_stat_card(stats_container, "Total Posisi", str(total_positions), self.colors['primary'], "👥", 0)
        self.create_stat_card(stats_container, "Email Tersedia", str(filled_emails), self.colors['success'], "✅", 1)
        self.create_stat_card(stats_container, "Email Kosong", str(empty_emails), self.colors['warning'], "❗", 2)
        self.create_stat_card(stats_container, "Tingkat Kelengkapan", f"{completion_rate:.1f}%", self.colors['secondary'], "📊", 3)
        
        # ENHANCED: Additional stats row for position breakdown
        stats_container2 = tk.Frame(self.content_frame, bg='#f0f0f0')
        stats_container2.pack(fill='x', pady=(0, 20))
        
        self.create_stat_card(stats_container2, "Directors & GM", str(director_gm_positions), self.colors['primary'], "🏢", 0)
        self.create_stat_card(stats_container2, "Managers", str(manager_positions), self.colors['secondary'], "👔", 1)
        self.create_stat_card(stats_container2, "Senior Officers", str(senior_officer_positions), self.colors['senior_officer'], "👨‍💼", 2)
        self.create_stat_card(stats_container2, "Avg per Manager", "2", self.colors['dark'], "📈", 3)
        
        # ENHANCED: Data table with grouping and improved design
        self._create_enhanced_table(df)
    
    def _create_enhanced_table(self, df):
        """ENHANCED: Create data table with position grouping"""
        table_frame = tk.Frame(self.content_frame, bg='white', relief='solid', borderwidth=1)
        table_frame.pack(fill='both', expand=True, pady=(20, 0))
        
        # Table header with enhanced functionality
        header_frame = tk.Frame(table_frame, bg='white')
        header_frame.pack(fill='x', padx=20, pady=15)
        
        tk.Label(header_frame,
                text="📋 Data Email Posisi (Enhanced)",
                font=('Arial', 16, 'bold'),
                bg='white',
                fg=self.colors['primary']).pack(side='left')
        
        # Enhanced search functionality
        search_frame = tk.Frame(header_frame, bg='white')
        search_frame.pack(side='right')
        
        # Filter dropdown
        tk.Label(search_frame, text="Filter:", font=('Arial', 10), bg='white').pack(side='left', padx=(0, 5))
        self.filter_var = tk.StringVar()
        filter_combo = ttk.Combobox(search_frame, textvariable=self.filter_var, width=15,
                                   values=["Semua", "Directors & GM", "Managers", "Senior Officers", "Email Kosong", "Email Tersedia"])
        filter_combo.set("Semua")
        filter_combo.pack(side='left', padx=(0, 10))
        filter_combo.bind('<<ComboboxSelected>>', lambda e: self._filter_table(df))
        
        tk.Label(search_frame, text="🔍 Cari:", font=('Arial', 10), bg='white').pack(side='left', padx=(0, 5))
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, font=('Arial', 10), width=20)
        search_entry.pack(side='left')
        search_var.trace('w', lambda *args: self._search_table(df, search_var.get()))
        
        # Notebook for tabbed view
        notebook = ttk.Notebook(table_frame)
        notebook.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # All positions tab
        all_tab = ttk.Frame(notebook)
        notebook.add(all_tab, text="Semua Posisi")
        self._create_position_tree(all_tab, df, self.positions)
        
        # Directors & GM tab
        directors_tab = ttk.Frame(notebook)
        notebook.add(directors_tab, text="Directors & GM")
        director_positions = [pos for pos in self.positions if not pos.startswith("Manager") and not pos.startswith("Senior Officer")]
        self._create_position_tree(directors_tab, df, director_positions)
        
        # Managers tab
        managers_tab = ttk.Frame(notebook)
        notebook.add(managers_tab, text="Managers")
        manager_positions = [pos for pos in self.positions if pos.startswith("Manager")]
        self._create_position_tree(managers_tab, df, manager_positions)
        
        # Senior Officers tab
        senior_tab = ttk.Frame(notebook)
        notebook.add(senior_tab, text="Senior Officers")
        senior_positions = [pos for pos in self.positions if pos.startswith("Senior Officer")]
        self._create_position_tree(senior_tab, df, senior_positions, highlight_senior=True)
    
    def _create_position_tree(self, parent, df, positions, highlight_senior=False):
        """Create treeview for specific position group"""
        tree_frame = tk.Frame(parent, bg='white')
        tree_frame.pack(fill='both', expand=True)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient='vertical')
        h_scrollbar = ttk.Scrollbar(tree_frame, orient='horizontal')
        
        tree = ttk.Treeview(tree_frame, 
                           columns=('posisi', 'email', 'last_updated', 'status'),
                           show="headings",
                           yscrollcommand=v_scrollbar.set,
                           xscrollcommand=h_scrollbar.set)
        
        v_scrollbar.config(command=tree.yview)
        h_scrollbar.config(command=tree.xview)
        
        # Configure columns
        tree.heading('posisi', text='POSISI')
        tree.heading('email', text='EMAIL')
        tree.heading('last_updated', text='TERAKHIR DIUPDATE')
        tree.heading('status', text='STATUS')
        
        tree.column('posisi', width=250, anchor='w')
        tree.column('email', width=300, anchor='w')
        tree.column('last_updated', width=150, anchor='center')
        tree.column('status', width=120, anchor='center')
        
        # Insert data for specified positions
        for i, position in enumerate(positions):
            # Find data for this position
            position_data = df[df['posisi'] == position] if not df.empty else None
            
            if position_data is not None and not position_data.empty:
                row_data = position_data.iloc[0]
                email_val = row_data['email'] if pd.notna(row_data['email']) and row_data['email'] != '' else '-'
                updated_val = row_data['last_updated'] if pd.notna(row_data['last_updated']) and row_data['last_updated'] != '' else '-'
            else:
                email_val = '-'
                updated_val = '-'
            
            status = "✅ Lengkap" if email_val != '-' else "❌ Kosong"
            
            # ENHANCED: Different colors for senior officers
            if highlight_senior and position.startswith("Senior Officer"):
                tags = ('senior_officer',)
            elif email_val != '-':
                tags = ('filled',)
            else:
                tags = ('empty',)
            
            if i % 2 == 0:
                tags = tags + ('evenrow',)
            else:
                tags = tags + ('oddrow',)
            
            tree.insert("", "end", values=(
                position,
                email_val,
                updated_val,
                status
            ), tags=tags)
        
        # Configure row colors
        tree.tag_configure('evenrow', background='#f8f9fa')
        tree.tag_configure('oddrow', background='white')
        tree.tag_configure('filled', foreground='#27ae60')
        tree.tag_configure('empty', foreground='#e74c3c')
        tree.tag_configure('senior_officer', foreground='#9b59b6', font=('Arial', 10, 'bold'))  # ENHANCED
        
        # Bind double click
        tree.bind('<Double-1>', lambda e: self._on_tree_double_click(e, tree))
        
        # Pack scrollbars and treeview
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')
        tree.pack(fill='both', expand=True)
    
    def _on_tree_double_click(self, event, tree):
        """Handle double click on tree item for editing"""
        selection = tree.selection()
        if selection:
            item = tree.item(selection[0])
            position = item['values'][0]
            self.show_edit_email_form(position)
    
    def _filter_table(self, df):
        """Filter table based on selected criteria"""
        # Implementation for filtering - this would update the tree views
        filter_value = self.filter_var.get()
        print(f"Filtering by: {filter_value}")
        # Add filtering logic here
    
    def _search_table(self, df, search_term):
        """Search table based on search term"""
        # Implementation for searching - this would update the tree views
        print(f"Searching for: {search_term}")
        # Add search logic here

    def create_stat_card(self, parent, title, value, color, icon, column):
        """Create a statistics card with improved design"""
        # Shadow frame
        shadow_frame = tk.Frame(parent, bg='#d0d0d0', height=102, width=202)
        shadow_frame.grid(row=0, column=column, padx=(10, 8), pady=(3, 0), sticky='ew')
        parent.grid_columnconfigure(column, weight=1)
        
        # Main card
        card = tk.Frame(shadow_frame, bg='white', relief='solid', borderwidth=1, height=100, width=200)
        card.place(x=-3, y=-3)
        card.pack_propagate(False)
        
        # Color bar
        color_bar = tk.Frame(card, bg=color, height=4)
        color_bar.pack(fill='x')
        
        # Content frame
        content_frame = tk.Frame(card, bg='white')
        content_frame.pack(fill='both', expand=True, padx=15, pady=12)
        
        # Icon and value container
        top_frame = tk.Frame(content_frame, bg='white')
        top_frame.pack(fill='x')
        
        # Icon
        tk.Label(top_frame,
                text=icon,
                font=('Arial', 20),
                bg='white',
                fg=color).pack(side='right')
        
        # Value
        tk.Label(top_frame,
                text=value,
                font=('Arial', 22, 'bold'),
                bg='white',
                fg=color).pack(side='left', anchor='w')
        
        # Title
        tk.Label(content_frame,
                text=title,
                font=('Arial', 11),
                bg='white',
                fg=self.colors['dark']).pack(anchor='w', pady=(5, 0))
        
    def show_email_management(self):
        """Show email management interface"""
        self.create_header("Kelola Email")
        self.create_navigation()
        self.set_active_nav('email')
        self.clear_content()
        
        # Main container with improved styling
        main_container = tk.Frame(self.content_frame, bg='white', relief='solid', borderwidth=1)
        main_container.pack(fill='both', expand=True)
        
        # Header with enhanced design
        header_frame = tk.Frame(main_container, bg='white')
        header_frame.pack(fill='x', padx=20, pady=20)
        
        # Title with icon
        title_frame = tk.Frame(header_frame, bg='white')
        title_frame.pack(side='left')
        
        tk.Label(title_frame,
                text="📧 Kelola Email Posisi",
                font=('Arial', 18, 'bold'),
                bg='white',
                fg=self.colors['primary']).pack()
        
        # Action buttons with improved styling
        btn_frame = tk.Frame(header_frame, bg='white')
        btn_frame.pack(side='right')
        
        add_btn = tk.Button(btn_frame,
                text="➕ Tambah Email",
                font=('Arial', 10, 'bold'),
                bg=self.colors['success'],
                fg='white',
                border=0,
                relief='flat',
                padx=15,
                pady=8,
                cursor='hand2',
                command=self.show_add_email_form)
        add_btn.pack(side='left', padx=5)
        
        refresh_btn = tk.Button(btn_frame,
                text="🔄 Refresh",
                font=('Arial', 10, 'bold'),
                bg=self.colors['secondary'],
                fg='white',
                border=0,
                relief='flat',
                padx=15,
                pady=8,
                cursor='hand2',
                command=self.refresh_email_table)
        refresh_btn.pack(side='left', padx=5)
        
        # Add hover effects
        def add_hover_effect(button, normal_color, hover_color):
            def on_enter(e):
                button.config(bg=hover_color)
            def on_leave(e):
                button.config(bg=normal_color)
            button.bind("<Enter>", on_enter)
            button.bind("<Leave>", on_leave)
        
        add_hover_effect(add_btn, self.colors['success'], '#229954')
        add_hover_effect(refresh_btn, self.colors['secondary'], '#2980b9')
        
        # Table container
        table_container = tk.Frame(main_container, bg='white')
        table_container.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # Create email table
        self.create_email_table(table_container)
        
    def create_email_table(self, parent):
        """Create email management table"""
        # Gunakan singkatan untuk manager dalam display
        abbreviation_map = {
            "Manager Pemeliharaan": "pml",
            "Manager Operasional": "ops",
            "Manager Administrasi": "adm",
            "Manager Keuangan": "keu"
        }
        
        # Clear existing widgets
        for widget in parent.winfo_children():
            widget.destroy()
            
        df = self.read_data(self.sheet)
        if df.empty:
            df = pd.DataFrame({
                'posisi': self.positions,
                'email': [''] * len(self.positions),
                'last_updated': [''] * len(self.positions)
            })
            self.write_data(self.sheet, df)
        
        # Search functionality
        search_frame = tk.Frame(parent, bg='white')
        search_frame.pack(fill='x', pady=(0, 10))
        
        tk.Label(search_frame, text="🔍 Cari Posisi:", font=('Arial', 10, 'bold'), bg='white').pack(side='left')
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, font=('Arial', 10), width=25)
        search_entry.pack(side='left', padx=(10, 0))
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(parent, orient='vertical')
        h_scrollbar = ttk.Scrollbar(parent, orient='horizontal')
        
        self.email_tree = ttk.Treeview(parent,
                                     columns=('posisi', 'email', 'last_updated', 'status'),
                                     show="headings",
                                     yscrollcommand=v_scrollbar.set,
                                     xscrollcommand=h_scrollbar.set)
        
        v_scrollbar.config(command=self.email_tree.yview)
        h_scrollbar.config(command=self.email_tree.xview)
        
        # Configure columns
        self.email_tree.heading('posisi', text='POSISI')
        self.email_tree.heading('email', text='EMAIL')
        self.email_tree.heading('last_updated', text='TERAKHIR DIUPDATE')
        self.email_tree.heading('status', text='STATUS')
        
        self.email_tree.column('posisi', width=200, anchor='w')
        self.email_tree.column('email', width=250, anchor='w')
        self.email_tree.column('last_updated', width=150, anchor='center')
        self.email_tree.column('status', width=100, anchor='center')
        
        # Insert data with status indicators
        for i, (_, row) in enumerate(df.iterrows()):
            email_val = row['email'] if pd.notna(row['email']) and row['email'] != '' else '-'
            updated_val = row['last_updated'] if pd.notna(row['last_updated']) and row['last_updated'] != '' else '-'
            status = "✅ Lengkap" if email_val != '-' else "❌ Kosong"
            
            # Gunakan singkatan untuk manager dalam display
            posisi = row['posisi']
            display_posisi = abbreviation_map.get(posisi, posisi)
            if display_posisi in ["pml", "ops", "adm", "keu"]:
                display_posisi = f"Manager {display_posisi}"
            
            tags = ('filled',) if email_val != '-' else ('empty',)
            if i % 2 == 0:
                tags = tags + ('evenrow',)
            else:
                tags = tags + ('oddrow',)
            
            self.email_tree.insert("", "end", values=(
                display_posisi,
                email_val,
                updated_val,
                status
            ), tags=tags)
        
        # Configure row colors and status colors
        self.email_tree.tag_configure('evenrow', background='#f8f9fa')
        self.email_tree.tag_configure('oddrow', background='white')
        self.email_tree.tag_configure('filled', foreground='#27ae60')
        self.email_tree.tag_configure('empty', foreground='#e74c3c')
        
        # Bind double click
        self.email_tree.bind('<Double-1>', self.on_email_row_double_click)
        
        # Pack scrollbars and treeview
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')
        self.email_tree.pack(fill='both', expand=True)
        
        # Search functionality
        def filter_email_data(*args):
            search_term = search_var.get().lower()
            for item in self.email_tree.get_children():
                self.email_tree.delete(item)
            
            for i, (_, row) in enumerate(df.iterrows()):
                # Gunakan singkatan untuk manager dalam search
                posisi = row['posisi']
                display_posisi = abbreviation_map.get(posisi, posisi)
                if display_posisi in ["pml", "ops", "adm", "keu"]:
                    display_posisi = f"Manager {display_posisi}"
                
                if search_term in display_posisi.lower():
                    email_val = row['email'] if pd.notna(row['email']) and row['email'] != '' else '-'
                    updated_val = row['last_updated'] if pd.notna(row['last_updated']) and row['last_updated'] != '' else '-'
                    status = "✅ Lengkap" if email_val != '-' else "❌ Kosong"
                    
                    tags = ('filled',) if email_val != '-' else ('empty',)
                    if i % 2 == 0:
                        tags = tags + ('evenrow',)
                    else:
                        tags = tags + ('oddrow',)
                    
                    self.email_tree.insert("", "end", values=(
                        display_posisi,
                        email_val,
                        updated_val,
                        status
                    ), tags=tags)
        
        search_var.trace('w', filter_email_data)
        
    def on_email_row_double_click(self, event):
        """Handle double click on email table row"""
        selection = self.email_tree.selection()
        if selection:
            item = self.email_tree.item(selection[0])
            display_posisi = item['values'][0]
            
            # Konversi dari singkatan ke nama lengkap untuk edit
            reverse_abbreviation_map = {
                "Manager pml": "Manager Pemeliharaan",
                "Manager ops": "Manager Operasional",
                "Manager adm": "Manager Administrasi",
                "Manager keu": "Manager Keuangan"
            }
            
            full_posisi = reverse_abbreviation_map.get(display_posisi, display_posisi)
            self.show_edit_email_form(full_posisi)
            
    def show_add_email_form(self):
        """Show add email form"""
        self.show_email_form()
        
    def show_edit_email_form(self, posisi):
        """Show edit email form for specific position"""
        self.show_email_form(posisi)
        
    def show_email_form(self, edit_posisi=None):
        """Show email form (add/edit) with enhanced UI"""
        # Gunakan singkatan untuk manager dalam display
        abbreviation_map = {
            "Manager Pemeliharaan": "pml",
            "Manager Operasional": "ops",
            "Manager Administrasi": "adm",
            "Manager Keuangan": "keu"
        }
        
        form_window = tk.Toplevel(self.root)
        if edit_posisi:
            # Gunakan singkatan untuk display dalam window title
            display_edit_posisi = abbreviation_map.get(edit_posisi, edit_posisi)
            if display_edit_posisi in ["pml", "ops", "adm", "keu"]:
                display_edit_posisi = f"Manager {display_edit_posisi}"
            form_window.title(f"Edit Email - {display_edit_posisi}")
        else:
            form_window.title("Tambah Email")
        form_window.geometry("450x350")
        form_window.configure(bg='white')
        form_window.transient(self.root)
        form_window.grab_set()
        form_window.resizable(False, False)
        
        # Center the window
        form_window.update_idletasks()
        x = (form_window.winfo_screenwidth() // 2) - (450 // 2)
        y = (form_window.winfo_screenheight() // 2) - (350 // 2)
        form_window.geometry(f"450x350+{x}+{y}")
        
        # Header frame with gradient-like effect
        header_frame = tk.Frame(form_window, bg=self.colors['primary'], height=60)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        # Header title
        if edit_posisi:
            # Gunakan singkatan untuk display dalam title
            display_edit_posisi = abbreviation_map.get(edit_posisi, edit_posisi)
            if display_edit_posisi in ["pml", "ops", "adm", "keu"]:
                display_edit_posisi = f"Manager {display_edit_posisi}"
            title_text = f"Edit Email - {display_edit_posisi}"
        else:
            title_text = "Tambah Email Baru"
            
        header_title = tk.Label(header_frame,
                               text="📝 " + title_text,
                               font=('Arial', 14, 'bold'),
                               bg=self.colors['primary'],
                               fg='white')
        header_title.pack(expand=True)
        
        # Form content
        form_frame = tk.Frame(form_window, bg='white')
        form_frame.pack(fill='both', expand=True, padx=30, pady=30)
        
        # Get current email if editing
        if edit_posisi:
            df = self.read_data(self.sheet)
            current_email = df[df['posisi'] == edit_posisi]['email'].iloc[0] if not df.empty and edit_posisi in df['posisi'].values else ''
        else:
            current_email = ''
        
        # Position selection
        tk.Label(form_frame, 
                text="Posisi:",
                font=('Arial', 12, 'bold'),
                bg='white',
                fg=self.colors['dark']).pack(anchor='w', pady=(0, 5))
        
        position_var = tk.StringVar()
        
        # Buat list untuk display dengan singkatan
        display_positions = []
        for pos in self.positions:
            abbrev = abbreviation_map.get(pos, pos)
            if abbrev in ["pml", "ops", "adm", "keu"]:
                abbrev = f"Manager {abbrev}"
            display_positions.append(abbrev)
        
        position_combo = ttk.Combobox(form_frame, 
                                    textvariable=position_var, 
                                    values=display_positions, 
                                    font=('Arial', 11),
                                    width=40,
                                    state='readonly' if edit_posisi else 'normal')
        position_combo.pack(pady=(0, 20), fill='x', ipady=5)
        
        if edit_posisi:
            # Konversi edit_posisi ke singkatan untuk display
            display_edit_posisi = abbreviation_map.get(edit_posisi, edit_posisi)
            if display_edit_posisi in ["pml", "ops", "adm", "keu"]:
                display_edit_posisi = f"Manager {display_edit_posisi}"
            position_var.set(display_edit_posisi)
        
        # Email input with validation
        tk.Label(form_frame, 
                text="Email:",
                font=('Arial', 12, 'bold'),
                bg='white',
                fg=self.colors['dark']).pack(anchor='w', pady=(0, 5))
        
        email_frame = tk.Frame(form_frame, bg='white')
        email_frame.pack(fill='x', pady=(0, 20))
        
        email_entry = tk.Entry(email_frame, 
                             font=('Arial', 11),
                             relief='solid',
                             borderwidth=1)
        email_entry.pack(fill='x', ipady=8)
        email_entry.insert(0, current_email if pd.notna(current_email) and current_email != '' else '')
        
        # Email validation label
        validation_label = tk.Label(form_frame,
                                   text="",
                                   font=('Arial', 9),
                                   bg='white')
        validation_label.pack(anchor='w')
        
        # Email validation function
        def validate_email(*args):
            email = email_entry.get().strip()
            if email:
                if '@' in email and '.' in email.split('@')[-1]:
                    validation_label.config(text="✅ Format email valid", fg=self.colors['success'])
                    return True
                else:
                    validation_label.config(text="❌ Format email tidak valid", fg=self.colors['danger'])
                    return False
            else:
                validation_label.config(text="")
                return False
        
        # Bind validation to email entry
        email_entry.bind('<KeyRelease>', validate_email)
        
        # Button frame
        btn_frame = tk.Frame(form_frame, bg='white')
        btn_frame.pack(fill='x', pady=(20, 0))
        
        # Save function
        def save_email():
            posisi = position_var.get().strip()
            email = email_entry.get().strip()
            
            # Konversi dari singkatan ke nama lengkap untuk penyimpanan
            reverse_abbreviation_map = {
                "Manager pml": "Manager Pemeliharaan",
                "Manager ops": "Manager Operasional",
                "Manager adm": "Manager Administrasi",
                "Manager keu": "Manager Keuangan"
            }
            
            # Konversi posisi dari display ke nama lengkap
            full_posisi = reverse_abbreviation_map.get(posisi, posisi)
            
            if not full_posisi:
                messagebox.showerror("Error", "Pilih posisi terlebih dahulu!")
                return
            if not email:
                messagebox.showerror("Error", "Email tidak boleh kosong!")
                return
            if not validate_email():
                messagebox.showerror("Error", "Format email tidak valid!")
                return
                
            df = self.read_data(self.sheet)
            if df.empty:
                df = pd.DataFrame({
                    'posisi': self.positions,
                    'email': [''] * len(self.positions),
                    'last_updated': [''] * len(self.positions)
                })
            
            # Check if position already has email (for add mode)
            if not edit_posisi and full_posisi in df['posisi'].values:
                existing_email = df[df['posisi'] == full_posisi]['email'].iloc[0]
                if pd.notna(existing_email) and existing_email != '':
                    # Gunakan singkatan untuk display dalam konfirmasi
                    display_posisi = abbreviation_map.get(full_posisi, full_posisi)
                    if display_posisi in ["pml", "ops", "adm", "keu"]:
                        display_posisi = f"Manager {display_posisi}"
                    if not messagebox.askyesno("Konfirmasi", f"Posisi {display_posisi} sudah memiliki email.\nApakah Anda ingin menggantinya?"):
                        return
            
            # Update or add
            if full_posisi in df['posisi'].values:
                df.loc[df['posisi'] == full_posisi, 'email'] = email
                df.loc[df['posisi'] == full_posisi, 'last_updated'] = datetime.now().strftime("%d/%m/%Y %H:%M")
            else:
                new_row = pd.DataFrame({
                    'posisi': [full_posisi],
                    'email': [email],
                    'last_updated': [datetime.now().strftime("%d/%m/%Y %H:%M")]
                })
                df = pd.concat([df, new_row], ignore_index=True)
            
            if self.write_data(self.sheet, df):
                # Gunakan singkatan untuk display dalam pesan sukses
                display_posisi = abbreviation_map.get(full_posisi, full_posisi)
                if display_posisi in ["pml", "ops", "adm", "keu"]:
                    display_posisi = f"Manager {display_posisi}"
                messagebox.showinfo("Sukses", f"Email untuk posisi '{display_posisi}' berhasil disimpan!")
                form_window.destroy()
                self.refresh_email_table()
            else:
                messagebox.showerror("Error", "Gagal menyimpan data ke Google Sheets!")
        
        # Delete function
        def delete_email():
            if edit_posisi:
                # Gunakan singkatan untuk display dalam konfirmasi
                display_edit_posisi = abbreviation_map.get(edit_posisi, edit_posisi)
                if display_edit_posisi in ["pml", "ops", "adm", "keu"]:
                    display_edit_posisi = f"Manager {display_edit_posisi}"
                if messagebox.askyesno("Konfirmasi Hapus", 
                                     f"Apakah Anda yakin ingin menghapus email untuk posisi '{display_edit_posisi}'?"):
                    df = self.read_data(self.sheet)
                    if not df.empty and edit_posisi in df['posisi'].values:
                        df.loc[df['posisi'] == edit_posisi, 'email'] = ''
                        df.loc[df['posisi'] == edit_posisi, 'last_updated'] = datetime.now().strftime("%d/%m/%Y %H:%M")
                        
                        if self.write_data(self.sheet, df):
                            # Gunakan singkatan untuk display dalam pesan sukses
                            display_edit_posisi = abbreviation_map.get(edit_posisi, edit_posisi)
                            if display_edit_posisi in ["pml", "ops", "adm", "keu"]:
                                display_edit_posisi = f"Manager {display_edit_posisi}"
                            messagebox.showinfo("Sukses", f"Email untuk posisi '{display_edit_posisi}' berhasil dihapus!")
                            form_window.destroy()
                            self.refresh_email_table()
                        else:
                            messagebox.showerror("Error", "Gagal menghapus data!")
        
        # Save button
        save_btn = tk.Button(btn_frame,
                text="💾 Simpan",
                font=('Arial', 11, 'bold'),
                bg=self.colors['success'],
                fg='white',
                border=0,
                relief='flat',
                padx=25,
                pady=10,
                cursor='hand2',
                command=save_email)
        save_btn.pack(side='left', padx=(0, 10))
        
        # Delete button (only for edit mode)
        if edit_posisi:
            delete_btn = tk.Button(btn_frame,
                    text="🗑️ Hapus",
                    font=('Arial', 11, 'bold'),
                    bg=self.colors['danger'],
                    fg='white',
                    border=0,
                    relief='flat',
                    padx=25,
                    pady=10,
                    cursor='hand2',
                    command=delete_email)
            delete_btn.pack(side='left', padx=(0, 10))
        
        # Cancel button
        cancel_btn = tk.Button(btn_frame,
                text="❌ Batal",
                font=('Arial', 11),
                bg=self.colors['dark'],
                fg='white',
                border=0,
                relief='flat',
                padx=25,
                pady=10,
                cursor='hand2',
                command=form_window.destroy)
        cancel_btn.pack(side='right')
        
        # Add hover effects to buttons
        def add_button_hover(button, normal_color, hover_color):
            def on_enter(e):
                button.config(bg=hover_color)
            def on_leave(e):
                button.config(bg=normal_color)
            button.bind("<Enter>", on_enter)
            button.bind("<Leave>", on_leave)
        
        add_button_hover(save_btn, self.colors['success'], '#229954')
        if edit_posisi:
            add_button_hover(delete_btn, self.colors['danger'], '#c0392b')
        add_button_hover(cancel_btn, self.colors['dark'], '#2c3e50')
        
        # Focus on appropriate field
        if edit_posisi:
            email_entry.focus()
            email_entry.select_range(0, 'end')
        else:
            position_combo.focus()
        
    def refresh_email_table(self):
        """Refresh email management table"""
        if hasattr(self, 'email_tree'):
            parent = self.email_tree.master
            self.create_email_table(parent)
        
    def show_senior_officers(self):
        """ENHANCED: Show dedicated senior officers management interface"""
        self.create_header("Senior Officers Management")
        self.create_navigation()
        self.set_active_nav('senior')
        self.clear_content()
        
        # Senior Officers specific interface
        main_container = tk.Frame(self.content_frame, bg='white', relief='solid', borderwidth=1)
        main_container.pack(fill='both', expand=True)
        
        # Header
        header_frame = tk.Frame(main_container, bg='white')
        header_frame.pack(fill='x', padx=20, pady=20)
        
        title_frame = tk.Frame(header_frame, bg='white')
        title_frame.pack(side='left')
        
        tk.Label(title_frame,
                text="👨‍💼 Senior Officers Email Management",
                font=('Arial', 18, 'bold'),
                bg='white',
                fg=self.colors['senior_officer']).pack()
        
        # Info panel
        info_frame = tk.Frame(main_container, bg='#f8f9fa', padx=20, pady=15)
        info_frame.pack(fill='x', padx=20, pady=(0, 20))
        
        info_text = """
        Senior Officers are automatically shown when their respective Manager is selected for email disposition.
        Each Manager has 2 associated Senior Officers:
        • Manager Pemeliharaan → Senior Officer Pemeliharaan 1 & 2
        • Manager Operasional → Senior Officer Operasional 1 & 2
        • Manager Administrasi → Senior Officer Administrasi 1 & 2
        • Manager Keuangan → Senior Officer Keuangan 1 & 2
        """
        
        tk.Label(info_frame,
                text=info_text,
                font=('Arial', 11),
                bg='#f8f9fa',
                fg=self.colors['dark'],
                justify='left').pack(anchor='w')
        
        # Senior Officers table
        self._create_senior_officers_table(main_container)
    
    def _create_senior_officers_table(self, parent):
        """Create dedicated table for senior officers"""
        table_container = tk.Frame(parent, bg='white')
        table_container.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        df = self.read_data(self.sheet)
        senior_positions = [pos for pos in self.positions if pos.startswith("Senior Officer")]
        
        # Create grouped table
        groups = {
            'Pemeliharaan': [pos for pos in senior_positions if 'Pemeliharaan' in pos],
            'Operasional': [pos for pos in senior_positions if 'Operasional' in pos],
            'Administrasi': [pos for pos in senior_positions if 'Administrasi' in pos],
            'Keuangan': [pos for pos in senior_positions if 'Keuangan' in pos]
        }
        
        for group_name, group_positions in groups.items():
            # Group header
            group_frame = tk.LabelFrame(table_container, 
                                       text=f"Senior Officers {group_name}",
                                       font=('Arial', 12, 'bold'),
                                       fg=self.colors['senior_officer'],
                                       padx=10, pady=10)
            group_frame.pack(fill='x', pady=(0, 15))
            
            for position in group_positions:
                pos_frame = tk.Frame(group_frame, bg='white')
                pos_frame.pack(fill='x', pady=5)
                
                # Position name
                tk.Label(pos_frame,
                        text=position,
                        font=('Arial', 11, 'bold'),
                        bg='white',
                        fg=self.colors['senior_officer'],
                        width=35).pack(side='left')
                
                # Email display/edit
                if not df.empty and position in df['posisi'].values:
                    email_val = df[df['posisi'] == position]['email'].iloc[0]
                    if pd.isna(email_val) or email_val == '':
                        email_val = 'Belum diisi'
                        email_color = self.colors['danger']
                    else:
                        email_color = self.colors['success']
                else:
                    email_val = 'Belum diisi'
                    email_color = self.colors['danger']
                
                tk.Label(pos_frame,
                        text=email_val,
                        font=('Arial', 11),
                        bg='white',
                        fg=email_color,
                        width=40).pack(side='left', padx=(10, 0))
                
                # Edit button
                edit_btn = tk.Button(pos_frame,
                                   text="✏️ Edit",
                                   font=('Arial', 9),
                                   bg=self.colors['secondary'],
                                   fg='white',
                                   border=0,
                                   padx=15,
                                   pady=5,
                                   cursor='hand2',
                                   command=lambda p=position: self.show_edit_email_form(p))
                edit_btn.pack(side='right')

    def show_change_password(self):
        """Show change password interface with enhanced UI"""
        self.create_header("Ganti Password")
        self.create_navigation()
        self.set_active_nav('password')
        self.clear_content()
        
        # Password form container with shadow effect
        form_outer = tk.Frame(self.content_frame, bg='#f0f0f0')
        form_outer.pack(expand=True)
        
        # Shadow frame
        shadow_frame = tk.Frame(form_outer, bg='#d0d0d0', height=402, width=502)
        shadow_frame.pack()
        shadow_frame.pack_propagate(False)
        
        # Main container
        form_container = tk.Frame(shadow_frame, bg='white', relief='solid', borderwidth=1, height=400, width=500)
        form_container.place(x=-3, y=-3)
        form_container.pack_propagate(False)
        
        # Header
        header_frame = tk.Frame(form_container, bg=self.colors['primary'], height=60)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame,
                text="🔒 Ganti Password Admin",
                font=('Arial', 16, 'bold'),
                bg=self.colors['primary'],
                fg='white').pack(expand=True)
        
        # Form frame
        form_frame = tk.Frame(form_container, bg='white')
        form_frame.pack(fill='both', expand=True, padx=40, pady=30)
        
        # Current password
        tk.Label(form_frame, 
                text="Password Saat Ini:",
                font=('Arial', 12, 'bold'),
                bg='white',
                fg=self.colors['dark']).pack(anchor='w', pady=(0, 5))
        
        current_entry = tk.Entry(form_frame, 
                               font=('Arial', 12), 
                               show="*",
                               relief='solid',
                               borderwidth=1)
        current_entry.pack(pady=(0, 20), fill='x', ipady=8)
        
        # New password
        tk.Label(form_frame, 
                text="Password Baru:",
                font=('Arial', 12, 'bold'),
                bg='white',
                fg=self.colors['dark']).pack(anchor='w', pady=(0, 5))
        
        new_entry = tk.Entry(form_frame, 
                           font=('Arial', 12), 
                           show="*",
                           relief='solid',
                           borderwidth=1)
        new_entry.pack(pady=(0, 20), fill='x', ipady=8)
        
        # Confirm password
        tk.Label(form_frame, 
                text="Konfirmasi Password Baru:",
                font=('Arial', 12, 'bold'),
                bg='white',
                fg=self.colors['dark']).pack(anchor='w', pady=(0, 5))
        
        confirm_entry = tk.Entry(form_frame, 
                               font=('Arial', 12), 
                               show="*",
                               relief='solid',
                               borderwidth=1)
        confirm_entry.pack(pady=(0, 20), fill='x', ipady=8)
        
        # Password strength indicator
        strength_label = tk.Label(form_frame,
                                text="",
                                font=('Arial', 10),
                                bg='white')
        strength_label.pack(anchor='w', pady=(0, 10))
        
        # Password strength validation
        def check_password_strength(*args):
            password = new_entry.get()
            if len(password) == 0:
                strength_label.config(text="", fg='black')
            elif len(password) < 6:
                strength_label.config(text="❌ Password terlalu pendek (minimal 6 karakter)", fg=self.colors['danger'])
            elif len(password) < 8:
                strength_label.config(text="⚠️ Password cukup kuat", fg=self.colors['warning'])
            else:
                strength_label.config(text="✅ Password kuat", fg=self.colors['success'])
        
        new_entry.bind('<KeyRelease>', check_password_strength)
        
        # Submit function
        def submit_password():
            current_password = current_entry.get()
            new_password = new_entry.get()
            confirm_password = confirm_entry.get()
            
            if not current_password:
                messagebox.showerror("Error", "Masukkan password saat ini!")
                current_entry.focus()
                return
                
            if not self.verify_password(current_password, self.admin_password_hash):
                messagebox.showerror("Error", "Password saat ini salah!")
                current_entry.delete(0, 'end')
                current_entry.focus()
                return
                
            if len(new_password) < 6:
                messagebox.showerror("Error", "Password baru minimal 6 karakter!")
                new_entry.focus()
                return
                
            if new_password != confirm_password:
                messagebox.showerror("Error", "Konfirmasi password tidak cocok!")
                confirm_entry.delete(0, 'end')
                confirm_entry.focus()
                return
            
            # Update password hash
            self.admin_password_hash = self.hash_password(new_password)
            messagebox.showinfo("Sukses", 
                              "Password berhasil diubah!\n\n" +
                              "Catatan: Perubahan password hanya berlaku selama sesi aplikasi ini berjalan. " +
                              "Setelah aplikasi ditutup, password akan kembali ke default.")
            
            # Clear all fields
            current_entry.delete(0, 'end')
            new_entry.delete(0, 'end')
            confirm_entry.delete(0, 'end')
            strength_label.config(text="")
        
        # Button frame
        btn_frame = tk.Frame(form_frame, bg='white')
        btn_frame.pack(fill='x', pady=(10, 0))
        
        # Submit button
        submit_btn = tk.Button(btn_frame,
                text="🔄 Ubah Password",
                font=('Arial', 12, 'bold'),
                bg=self.colors['secondary'],
                fg='white',
                border=0,
                relief='flat',
                padx=30,
                pady=12,
                cursor='hand2',
                command=submit_password)
        submit_btn.pack()
        
        # Add hover effect
        def on_enter(e):
            submit_btn.config(bg='#2980b9')
        def on_leave(e):
            submit_btn.config(bg=self.colors['secondary'])
        
        submit_btn.bind("<Enter>", on_enter)
        submit_btn.bind("<Leave>", on_leave)
        
        # Focus on first field
        current_entry.focus()
        
        # Bind Enter key to submit
        def on_enter_key(event):
            submit_password()
        
        current_entry.bind('<Return>', lambda e: new_entry.focus())
        new_entry.bind('<Return>', lambda e: confirm_entry.focus())
        confirm_entry.bind('<Return>', on_enter_key)
        
    # Utility functions
    def hash_password(self, password):
        """Hash password using SHA-256"""
        return hashlib.sha256(password.encode()).hexdigest()
    
    def verify_password(self, password, hashed):
        """Verify password against hash"""
        return self.hash_password(password) == hashed
    
    def init_google_sheets(self):
        """Initialize Google Sheets connection"""
        try:
            gc = gspread.service_account(filename='credentials/credentials.json')
            sheet_url = "https://docs.google.com/spreadsheets/d/1LoAzVPBMJo08uPHR7MdzGEXmaoFXMrTSKS0vl099qrU/edit?gid=0#gid=0"
            sheet = gc.open_by_url(sheet_url).sheet1
            return sheet
        except FileNotFoundError:
            print("Error: credentials/credentials.json not found")
            return None
        except Exception as e:
            print(f"Error connecting to Google Sheets: {str(e)}")
            return None
    
    def read_data(self, sheet):
        """Read data from Google Sheets"""
        try:
            if sheet is None:
                return pd.DataFrame()
            data = sheet.get_all_records()
            return pd.DataFrame(data)
        except Exception as e:
            print(f"Error reading data: {str(e)}")
            return pd.DataFrame()
    
    def write_data(self, sheet, df):
        """Write data to Google Sheets"""
        try:
            if sheet is None:
                return False
            sheet.clear()
            # Convert DataFrame to list format for Google Sheets
            data_to_write = [df.columns.values.tolist()] + df.values.tolist()
            sheet.update(data_to_write)
            return True
        except Exception as e:
            print(f"Error writing data: {str(e)}")
            return False
    
    def run(self):
        """Run the application"""
        try:
            # Set window icon if available
            # self.root.iconbitmap('icon.ico')  # Uncomment if you have an icon file
            pass
        except:
            pass
        
        # Handle window closing
        def on_closing():
            if messagebox.askokcancel("Keluar", "Apakah Anda yakin ingin keluar dari aplikasi?"):
                self.root.destroy()
        
        self.root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # Start the application
        self.root.mainloop()

# Test function for the enhanced system
def test_enhanced_email_system():
    """Test the enhanced email system with senior officer support"""
    print("Testing Enhanced Email System with Senior Officer Support")
    print("=" * 60)
    
    email_sender = EmailManagerApp()
    
    # Test connection
    results = email_sender.test_connection()
    
    print("\nConnection Test Summary:")
    print(f"{'SMTP Connection:':<25} {'✓ OK' if results['smtp_connection'] else '✗ FAILED'}")
    print(f"{'Google Sheets:':<25} {'✓ OK' if results['sheets_connection'] else '✗ FAILED'}")
    print(f"{'Admin Sheet Access:':<25} {'✓ OK' if results['admin_sheet_access'] else '✗ FAILED'}")
    
    if results['email_data']:
        print(f"\nFound emails for {len(results['email_data'])} positions:")
        
        # Group by category
        managers = []
        senior_officers = []
        others = []
        
        for position, email in results['email_data'].items():
            if position.startswith("Senior Officer"):
                senior_officers.append((position, email))
            elif position.startswith("Manager"):
                managers.append((position, email))
            else:
                others.append((position, email))
        
        # Display by category
        if others:
            print("\n  Directors & GM:")
            for position, email in others:
                print(f"    • {position}: {email}")
        
        if managers:
            print("\n  Managers:")
            for position, email in managers:
                print(f"    • {position}: {email}")
        
        if senior_officers:
            print("\n  Senior Officers:")
            for position, email in senior_officers:
                print(f"    • {position}: {email}")
    
    if results['errors']:
        print(f"\nErrors encountered:")
        for error in results['errors']:
            print(f"  • {error}")
    
    # Test senior officer email lookup
    print(f"\nTesting Senior Officer Email Lookup:")
    print("-" * 40)
    
    test_senior_positions = [
        "Senior Officer Pemeliharaan 1",
        "Senior Officer Operasional 2",
        "Senior Officer Keuangan 1"
    ]
    
    for position in test_senior_positions:
        email, msg = email_sender.get_recipient_email(position)
        if email:
            print(f"✓ {position}: {email}")
        else:
            print(f"✗ {position}: {msg}")
    
    return results


if __name__ == "__main__":
    test_enhanced_email_system()
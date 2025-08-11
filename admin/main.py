import gspread
import pandas as pd
from datetime import datetime
import json
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter import font as tkFont

class EmailManagerApp:
    def __init__(self):
        """Initialize the Minimalist Email Manager Application"""
        self.root = tk.Tk()
        self.root.title("Email Manager")
        self.root.geometry("1400x900")
        self.root.configure(bg='#ffffff')
        self.root.minsize(1200, 700)
        
        # Initialize variables
        self.logged_in = False
        self.sheet = None
        
        # Position data
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
            "Senior Officer Pemeliharaan 1",
            "Senior Officer Pemeliharaan 2",
            "Senior Officer Operasional 1",
            "Senior Officer Operasional 2",
            "Senior Officer Administrasi 1",
            "Senior Officer Administrasi 2",
            "Senior Officer Keuangan 1",
            "Senior Officer Keuangan 2"
        ]
        
        # Position categorization
        self.position_categories = {
            'Directors & GM': [
                "Direktur Utama",
                "Direktur Keuangan", 
                "Direktur Teknik",
                "GM Keuangan & Administrasi",
                "GM Operasional & Pemeliharaan"
            ],
            'Managers': [
                "Manager Pemeliharaan",
                "Manager Operasional",
                "Manager Administrasi",
                "Manager Keuangan"
            ],
            'Senior Officers': [
                "Senior Officer Pemeliharaan 1",
                "Senior Officer Pemeliharaan 2",
                "Senior Officer Operasional 1",
                "Senior Officer Operasional 2",
                "Senior Officer Administrasi 1",
                "Senior Officer Administrasi 2",
                "Senior Officer Keuangan 1",
                "Senior Officer Keuangan 2"
            ]
        }
        
        # Setup minimalist styles
        self.setup_minimalist_styles()
        self.create_main_layout()
        
        # Show login initially
        self.show_login()
    
    def setup_minimalist_styles(self):
        """Setup minimalist color scheme and styles"""
        self.colors = {
            # Minimalist color palette
            'primary': '#1a1a1a',
            'secondary': '#4a5568',
            'accent': '#667eea',
            'success': '#48bb78',
            'warning': '#ed8936',
            'danger': '#f56565',
            'background': '#ffffff',
            'surface': '#f7fafc',
            'border': '#e2e8f0',
            'text_primary': '#2d3748',
            'text_secondary': '#718096',
            'text_muted': '#a0aec0',
            'white': '#ffffff',
            'overlay': 'rgba(0,0,0,0.05)'
        }
        
        # Setup ttk styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Minimalist button styles
        self.style.configure('Minimal.TButton',
                           font=('Inter', 10),
                           borderwidth=0,
                           focuscolor='none',
                           relief='flat')
        
        self.style.configure('Primary.TButton',
                           font=('Inter', 10, 'bold'),
                           background=self.colors['accent'],
                           foreground='white',
                           borderwidth=0,
                           focuscolor='none')
        
        # Minimalist combobox
        self.style.configure('Minimal.TCombobox',
                           font=('Inter', 10),
                           fieldbackground='white',
                           borderwidth=1,
                           relief='solid')
        
        # Clean treeview
        self.style.configure("Clean.Treeview",
                           background="white",
                           foreground=self.colors['text_primary'],
                           rowheight=35,
                           fieldbackground="white",
                           borderwidth=0,
                           font=('Inter', 10))
        
        self.style.configure("Clean.Treeview.Heading",
                           font=('Inter', 10, 'bold'),
                           background=self.colors['surface'],
                           foreground=self.colors['text_primary'],
                           relief='flat',
                           borderwidth=0)
        
        self.style.map("Clean.Treeview",
                      background=[('selected', self.colors['accent'])])
    
    def create_main_layout(self):
        """Create minimalist main layout"""
        # Sidebar for navigation
        self.sidebar = tk.Frame(self.root, bg=self.colors['surface'], width=240)
        self.sidebar.pack(side='left', fill='y')
        self.sidebar.pack_propagate(False)
        
        # Main content area
        self.main_area = tk.Frame(self.root, bg=self.colors['background'])
        self.main_area.pack(side='right', fill='both', expand=True)
        
        # Content container with padding
        self.content_frame = tk.Frame(self.main_area, bg=self.colors['background'])
        self.content_frame.pack(fill='both', expand=True, padx=32, pady=32)
    
    def create_sidebar(self):
        """Create minimalist sidebar navigation"""
        # Clear sidebar
        for widget in self.sidebar.winfo_children():
            widget.destroy()
        
        if not self.logged_in:
            return
        
        # Logo/Title area
        logo_frame = tk.Frame(self.sidebar, bg=self.colors['surface'], height=80)
        logo_frame.pack(fill='x', padx=24, pady=(24, 0))
        logo_frame.pack_propagate(False)
        
        # App title
        title_label = tk.Label(logo_frame,
                              text="Email Manager",
                              font=('Inter', 16, 'bold'),
                              bg=self.colors['surface'],
                              fg=self.colors['primary'])
        title_label.pack(anchor='w', pady=(20, 0))
        
        # Subtitle
        subtitle_label = tk.Label(logo_frame,
                                 text="v2.0",
                                 font=('Inter', 10),
                                 bg=self.colors['surface'],
                                 fg=self.colors['text_secondary'])
        subtitle_label.pack(anchor='w')
        
        # Navigation menu
        nav_frame = tk.Frame(self.sidebar, bg=self.colors['surface'])
        nav_frame.pack(fill='both', expand=True, padx=16, pady=24)
        
        # Navigation items
        nav_items = [
            ("Dashboard", "📊", self.show_dashboard, 'dashboard'),
            ("Email Management", "✉️", self.show_email_management, 'email')
        ]
        
        self.nav_buttons = {}
        for i, (text, icon, command, key) in enumerate(nav_items):
            btn_container = tk.Frame(nav_frame, bg=self.colors['surface'])
            btn_container.pack(fill='x', pady=(0, 8))
            
            btn = tk.Button(btn_container,
                          text=f"  {icon}  {text}",
                          font=('Inter', 11),
                          bg=self.colors['surface'],
                          fg=self.colors['text_primary'],
                          anchor='w',
                          border=0,
                          relief='flat',
                          pady=12,
                          cursor='hand2',
                          command=command)
            btn.pack(fill='x', padx=8)
            
            self.nav_buttons[key] = btn
            
            # Hover effects
            def make_nav_hover(button, key):
                def on_enter(e):
                    if not hasattr(self, 'active_nav') or self.active_nav != key:
                        button.config(bg=self.colors['border'])
                def on_leave(e):
                    if not hasattr(self, 'active_nav') or self.active_nav != key:
                        button.config(bg=self.colors['surface'])
                return on_enter, on_leave
            
            enter_h, leave_h = make_nav_hover(btn, key)
            btn.bind("<Enter>", enter_h)
            btn.bind("<Leave>", leave_h)
        
        # Status section
        status_frame = tk.Frame(self.sidebar, bg=self.colors['surface'])
        status_frame.pack(side='bottom', fill='x', padx=24, pady=24)
        
        # Connection status
        status_indicator = tk.Frame(status_frame, bg=self.colors['surface'])
        status_indicator.pack(fill='x', pady=(0, 16))
        
        tk.Label(status_indicator,
                text="● Connected",
                font=('Inter', 9),
                bg=self.colors['surface'],
                fg=self.colors['success']).pack(anchor='w')
        
        # Logout button
        logout_btn = tk.Button(status_frame,
                             text="Sign Out",
                             font=('Inter', 10),
                             bg=self.colors['surface'],
                             fg=self.colors['text_secondary'],
                             border=0,
                             relief='flat',
                             cursor='hand2',
                             command=self.logout)
        logout_btn.pack(anchor='w')
        
        # Logout hover
        def logout_hover(e):
            logout_btn.config(fg=self.colors['danger'])
        def logout_leave(e):
            logout_btn.config(fg=self.colors['text_secondary'])
        
        logout_btn.bind("<Enter>", logout_hover)
        logout_btn.bind("<Leave>", logout_leave)
    
    def set_active_nav(self, active_key):
        """Set active navigation button"""
        self.active_nav = active_key
        if hasattr(self, 'nav_buttons'):
            for key, btn in self.nav_buttons.items():
                if key == active_key:
                    btn.config(bg=self.colors['accent'], fg='white')
                else:
                    btn.config(bg=self.colors['surface'], fg=self.colors['text_primary'])
    
    def clear_content(self):
        """Clear main content area"""
        for widget in self.content_frame.winfo_children():
            widget.destroy()
    
    def show_login(self):
        """Minimalist login screen"""
        # Hide sidebar
        self.sidebar.pack_forget()
        
        # Center the main area
        self.main_area.pack(fill='both', expand=True)
        self.clear_content()
        
        # Login container
        login_container = tk.Frame(self.content_frame, bg=self.colors['background'])
        login_container.place(relx=0.5, rely=0.5, anchor='center')
        
        # Login card
        card = tk.Frame(login_container, bg='white', relief='solid', borderwidth=1)
        card.pack(padx=40, pady=40)
        
        # Card content
        card_content = tk.Frame(card, bg='white')
        card_content.pack(padx=60, pady=60)
        
        # Title
        title = tk.Label(card_content,
                        text="Welcome",
                        font=('Inter', 24, 'bold'),
                        bg='white',
                        fg=self.colors['primary'])
        title.pack(pady=(0, 8))
        
        subtitle = tk.Label(card_content,
                           text="Sign in to continue",
                           font=('Inter', 12),
                           bg='white',
                           fg=self.colors['text_secondary'])
        subtitle.pack(pady=(0, 32))
        
        # Password field
        password_label = tk.Label(card_content,
                                 text="Password",
                                 font=('Inter', 11),
                                 bg='white',
                                 fg=self.colors['text_primary'])
        password_label.pack(anchor='w', pady=(0, 8))
        
        self.password_entry = tk.Entry(card_content,
                                     font=('Inter', 12),
                                     show="*",
                                     relief='solid',
                                     borderwidth=1,
                                     bg='white',
                                     width=25)
        self.password_entry.pack(pady=(0, 24), ipady=8)
        self.password_entry.bind('<Return>', lambda e: self.try_login())
        
        # Login button
        login_btn = tk.Button(card_content,
                            text="Sign In",
                            font=('Inter', 11, 'bold'),
                            bg=self.colors['accent'],
                            fg='white',
                            border=0,
                            relief='flat',
                            pady=10,
                            cursor='hand2',
                            command=self.try_login)
        login_btn.pack(fill='x', pady=(0, 16))
        
        # Help text
        help_text = tk.Label(card_content,
                           text="Default password: admin123",
                           font=('Inter', 9),
                           bg='white',
                           fg=self.colors['text_muted'])
        help_text.pack()
        
        # Focus
        self.password_entry.focus()
    
    def try_login(self):
        """Handle login attempt"""
        password = self.password_entry.get()
        if password == "admin123":
            self.logged_in = True
            self.sheet = self.init_google_sheets()
            if self.sheet is None:
                messagebox.showerror("Connection Error", 
                                   "Cannot connect to Google Sheets.\n\nPlease check credentials/credentials.json file.")
                return
            self.show_dashboard()
        else:
            messagebox.showerror("Login Failed", "Incorrect password!")
            self.password_entry.delete(0, 'end')
            self.password_entry.focus()
    
    def logout(self):
        """Handle logout"""
        if messagebox.askyesno("Confirm Logout", "Are you sure you want to sign out?"):
            self.logged_in = False
            self.sheet = None
            self.show_login()
    
    def show_dashboard(self):
        """Minimalist dashboard"""
        # Show sidebar
        self.sidebar.pack(side='left', fill='y')
        self.create_sidebar()
        self.set_active_nav('dashboard')
        self.clear_content()
        
        # Page header
        header_frame = tk.Frame(self.content_frame, bg=self.colors['background'])
        header_frame.pack(fill='x', pady=(0, 32))
        
        tk.Label(header_frame,
                text="Dashboard",
                font=('Inter', 28, 'bold'),
                bg=self.colors['background'],
                fg=self.colors['primary']).pack(side='left')
        
        # Stats cards
        self.create_stats_section()
        
        # Data table
        self.create_dashboard_table()
    
    def create_stats_section(self):
        """Create minimalist stats cards with accurate statistics"""
        # Read data
        df = self.read_data(self.sheet)
        
        # Initialize empty dataframe if needed
        if df.empty:
            df = pd.DataFrame({
                'posisi': self.positions,
                'email': [''] * len(self.positions),
                'last_updated': [''] * len(self.positions)
            })
            self.write_data(self.sheet, df)
        
        # Ensure all positions exist in dataframe
        for position in self.positions:
            if position not in df['posisi'].values:
                new_row = pd.DataFrame({
                    'posisi': [position],
                    'email': [''],
                    'last_updated': ['']
                })
                df = pd.concat([df, new_row], ignore_index=True)
        
        # Calculate accurate statistics
        total_positions = len(self.positions)  # Use predefined positions count
        
        # Count configured emails (not empty and not null)
        if 'email' in df.columns:
            configured_count = 0
            for position in self.positions:
                position_data = df[df['posisi'] == position]
                if not position_data.empty:
                    email_val = position_data.iloc[0]['email']
                    if pd.notna(email_val) and email_val.strip() != '':
                        configured_count += 1
        else:
            configured_count = 0
        
        pending_count = total_positions - configured_count
        completion_rate = (configured_count / total_positions * 100) if total_positions > 0 else 0
        
        # Calculate category-wise statistics
        category_stats = {}
        for category, positions_list in self.position_categories.items():
            total_in_category = len(positions_list)
            configured_in_category = 0
            
            for position in positions_list:
                position_data = df[df['posisi'] == position]
                if not position_data.empty:
                    email_val = position_data.iloc[0]['email']
                    if pd.notna(email_val) and email_val.strip() != '':
                        configured_in_category += 1
            
            category_stats[category] = {
                'total': total_in_category,
                'configured': configured_in_category,
                'rate': (configured_in_category / total_in_category * 100) if total_in_category > 0 else 0
            }
        
        # Stats container
        stats_container = tk.Frame(self.content_frame, bg=self.colors['background'])
        stats_container.pack(fill='x', pady=(0, 32))
        
        # Main stats cards
        stats_data = [
            ("Total Positions", str(total_positions), self.colors['primary']),
            ("Configured", str(configured_count), self.colors['success']),
            ("Pending", str(pending_count), self.colors['warning']),
            ("Completion", f"{completion_rate:.1f}%", self.colors['accent'])
        ]
        
        # Create main stats row
        main_stats_frame = tk.Frame(stats_container, bg=self.colors['background'])
        main_stats_frame.pack(fill='x', pady=(0, 16))
        
        for i, (label, value, color) in enumerate(stats_data):
            card = tk.Frame(main_stats_frame, bg='white', relief='solid', borderwidth=1)
            card.grid(row=0, column=i, padx=(0, 16) if i < 3 else 0, sticky='ew')
            main_stats_frame.grid_columnconfigure(i, weight=1)
            
            # Card content
            content = tk.Frame(card, bg='white')
            content.pack(fill='both', expand=True, padx=24, pady=20)
            
            # Value
            tk.Label(content,
                    text=value,
                    font=('Inter', 24, 'bold'),
                    bg='white',
                    fg=color).pack()
            
            # Label
            tk.Label(content,
                    text=label,
                    font=('Inter', 11),
                    bg='white',
                    fg=self.colors['text_secondary']).pack(pady=(4, 0))
        
        # Category breakdown stats
        category_frame = tk.Frame(stats_container, bg=self.colors['background'])
        category_frame.pack(fill='x')
        
        category_cards_data = [
            ("Directors & GM", f"{category_stats['Directors & GM']['configured']}/{category_stats['Directors & GM']['total']}", 
             f"{category_stats['Directors & GM']['rate']:.1f}%", self.colors['primary']),
            ("Managers", f"{category_stats['Managers']['configured']}/{category_stats['Managers']['total']}", 
             f"{category_stats['Managers']['rate']:.1f}%", self.colors['secondary']),
            ("Senior Officers", f"{category_stats['Senior Officers']['configured']}/{category_stats['Senior Officers']['total']}", 
             f"{category_stats['Senior Officers']['rate']:.1f}%", self.colors['accent']),
            ("Average Rate", f"{completion_rate:.1f}%", "Overall", self.colors['text_secondary'])
        ]
        
        for i, (category, ratio, rate, color) in enumerate(category_cards_data):
            card = tk.Frame(category_frame, bg='white', relief='solid', borderwidth=1)
            card.grid(row=0, column=i, padx=(0, 16) if i < 3 else 0, sticky='ew')
            category_frame.grid_columnconfigure(i, weight=1)
            
            # Card content
            content = tk.Frame(card, bg='white')
            content.pack(fill='both', expand=True, padx=20, pady=16)
            
            # Category name
            tk.Label(content,
                    text=category,
                    font=('Inter', 10, 'bold'),
                    bg='white',
                    fg=color).pack()
            
            # Ratio/Value
            tk.Label(content,
                    text=ratio,
                    font=('Inter', 16, 'bold'),
                    bg='white',
                    fg=self.colors['text_primary']).pack(pady=(2, 0))
            
            # Rate
            tk.Label(content,
                    text=rate,
                    font=('Inter', 9),
                    bg='white',
                    fg=self.colors['text_secondary']).pack()
    
    def create_dashboard_table(self):
        """Create minimalist dashboard table with accurate data"""
        # Table container
        table_container = tk.Frame(self.content_frame, bg='white', relief='solid', borderwidth=1)
        table_container.pack(fill='both', expand=True)
        
        # Table header
        table_header = tk.Frame(table_container, bg='white')
        table_header.pack(fill='x', padx=24, pady=16)
        
        tk.Label(table_header,
                text="Position Overview",
                font=('Inter', 16, 'bold'),
                bg='white',
                fg=self.colors['primary']).pack(side='left')
        
        # Add button
        add_btn = tk.Button(table_header,
                          text="+ Add Email",
                          font=('Inter', 10, 'bold'),
                          bg=self.colors['accent'],
                          fg='white',
                          border=0,
                          relief='flat',
                          padx=16,
                          pady=8,
                          cursor='hand2',
                          command=lambda: self.show_email_form())
        add_btn.pack(side='right')
        
        # Separator
        separator = tk.Frame(table_container, bg=self.colors['border'], height=1)
        separator.pack(fill='x', padx=24)
        
        # Treeview container
        tree_container = tk.Frame(table_container, bg='white')
        tree_container.pack(fill='both', expand=True, padx=24, pady=16)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_container, orient='vertical')
        v_scrollbar.pack(side='right', fill='y')
        
        # Treeview
        self.dashboard_tree = ttk.Treeview(tree_container,
                                         columns=('position', 'category', 'email', 'status'),
                                         show="headings",
                                         style="Clean.Treeview",
                                         yscrollcommand=v_scrollbar.set)
        
        v_scrollbar.config(command=self.dashboard_tree.yview)
        
        # Configure columns
        self.dashboard_tree.heading('position', text='Position')
        self.dashboard_tree.heading('category', text='Category')
        self.dashboard_tree.heading('email', text='Email')
        self.dashboard_tree.heading('status', text='Status')
        
        self.dashboard_tree.column('position', width=200, anchor='w')
        self.dashboard_tree.column('category', width=120, anchor='center')
        self.dashboard_tree.column('email', width=250, anchor='w')
        self.dashboard_tree.column('status', width=100, anchor='center')
        
        # Get accurate data
        df = self.read_data(self.sheet)
        
        # Ensure all positions exist in dataframe
        if df.empty:
            df = pd.DataFrame({
                'posisi': self.positions,
                'email': [''] * len(self.positions),
                'last_updated': [''] * len(self.positions)
            })
            self.write_data(self.sheet, df)
        
        # Add any missing positions
        for position in self.positions:
            if position not in df['posisi'].values:
                new_row = pd.DataFrame({
                    'posisi': [position],
                    'email': [''],
                    'last_updated': ['']
                })
                df = pd.concat([df, new_row], ignore_index=True)
        
        # Populate data accurately
        for position in self.positions:
            # Determine category
            if position in self.position_categories['Directors & GM']:
                category = "Director/GM"
            elif position in self.position_categories['Managers']:
                category = "Manager"
            elif position in self.position_categories['Senior Officers']:
                category = "Senior Officer"
            else:
                category = "Other"
            
            # Get email data accurately
            email_val = ''
            position_data = df[df['posisi'] == position]
            if not position_data.empty:
                email_raw = position_data.iloc[0]['email']
                if pd.notna(email_raw) and str(email_raw).strip() != '':
                    email_val = str(email_raw).strip()
            
            # Determine status and display
            if email_val:
                status = "✓ Configured"
                display_email = email_val
                tags = ('configured',)
            else:
                status = "○ Pending"
                display_email = "Not configured"
                tags = ('pending',)
            
            self.dashboard_tree.insert("", "end", values=(position, category, display_email, status), tags=tags)
        
        # Configure row styling
        self.dashboard_tree.tag_configure('configured', foreground=self.colors['success'])
        self.dashboard_tree.tag_configure('pending', foreground=self.colors['warning'])
        
        # Double click to edit
        self.dashboard_tree.bind('<Double-1>', self.on_dashboard_double_click)
        
        self.dashboard_tree.pack(fill='both', expand=True)
    
    def on_dashboard_double_click(self, event):
        """Handle dashboard table double click"""
        selection = self.dashboard_tree.selection()
        if selection:
            item = self.dashboard_tree.item(selection[0])
            position = item['values'][0]
            self.show_email_form(position)
    
    def show_email_management(self):
        """Minimalist email management"""
        self.create_sidebar()
        self.set_active_nav('email')
        self.clear_content()
        
        # Page header
        header_frame = tk.Frame(self.content_frame, bg=self.colors['background'])
        header_frame.pack(fill='x', pady=(0, 32))
        
        tk.Label(header_frame,
                text="Email Management",
                font=('Inter', 28, 'bold'),
                bg=self.colors['background'],
                fg=self.colors['primary']).pack(side='left')
        
        # Action buttons
        btn_frame = tk.Frame(header_frame, bg=self.colors['background'])
        btn_frame.pack(side='right')
        
        add_btn = tk.Button(btn_frame,
                          text="+ Add",
                          font=('Inter', 10, 'bold'),
                          bg=self.colors['accent'],
                          fg='white',
                          border=0,
                          relief='flat',
                          padx=16,
                          pady=8,
                          cursor='hand2',
                          command=lambda: self.show_email_form())
        add_btn.pack(side='left', padx=(0, 8))
        
        bulk_btn = tk.Button(btn_frame,
                           text="Bulk Import",
                           font=('Inter', 10),
                           bg=self.colors['surface'],
                           fg=self.colors['text_primary'],
                           border=0,
                           relief='flat',
                           padx=16,
                           pady=8,
                           cursor='hand2',
                           command=self.show_bulk_import)
        bulk_btn.pack(side='left')
        
        # Email table
        self.create_email_table()
    
    def create_email_table(self):
        """Create minimalist email table with accurate data"""
        # Table container
        table_container = tk.Frame(self.content_frame, bg='white', relief='solid', borderwidth=1)
        table_container.pack(fill='both', expand=True)
        
        # Search bar
        search_frame = tk.Frame(table_container, bg='white')
        search_frame.pack(fill='x', padx=24, pady=16)
        
        tk.Label(search_frame,
                text="Search:",
                font=('Inter', 11),
                bg='white',
                fg=self.colors['text_primary']).pack(side='left', padx=(0, 8))
        
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame,
                               textvariable=self.search_var,
                               font=('Inter', 10),
                               relief='solid',
                               borderwidth=1,
                               width=30)
        search_entry.pack(side='left')
        
        # Filter dropdown
        tk.Label(search_frame,
                text="Filter:",
                font=('Inter', 11),
                bg='white',
                fg=self.colors['text_primary']).pack(side='left', padx=(24, 8))
        
        self.filter_var = tk.StringVar()
        filter_combo = ttk.Combobox(search_frame,
                                   textvariable=self.filter_var,
                                   values=["All", "Configured", "Pending", "Directors & GM", "Managers", "Senior Officers"],
                                   font=('Inter', 10),
                                   style='Minimal.TCombobox',
                                   state='readonly',
                                   width=15)
        filter_combo.set("All")
        filter_combo.pack(side='left')
        
        # Bind filter events
        self.search_var.trace('w', self.filter_email_table)
        filter_combo.bind('<<ComboboxSelected>>', self.filter_email_table)
        
        # Separator
        separator = tk.Frame(table_container, bg=self.colors['border'], height=1)
        separator.pack(fill='x', padx=24)
        
        # Treeview container
        tree_container = tk.Frame(table_container, bg='white')
        tree_container.pack(fill='both', expand=True, padx=24, pady=16)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_container, orient='vertical')
        v_scrollbar.pack(side='right', fill='y')
        
        # Treeview
        self.email_tree = ttk.Treeview(tree_container,
                                     columns=('position', 'category', 'email', 'updated', 'actions'),
                                     show="headings",
                                     style="Clean.Treeview",
                                     yscrollcommand=v_scrollbar.set)
        
        v_scrollbar.config(command=self.email_tree.yview)
        
        # Configure columns
        self.email_tree.heading('position', text='Position')
        self.email_tree.heading('category', text='Category')
        self.email_tree.heading('email', text='Email')
        self.email_tree.heading('updated', text='Last Updated')
        self.email_tree.heading('actions', text='Actions')
        
        self.email_tree.column('position', width=180, anchor='w')
        self.email_tree.column('category', width=100, anchor='center')
        self.email_tree.column('email', width=220, anchor='w')
        self.email_tree.column('updated', width=120, anchor='center')
        self.email_tree.column('actions', width=80, anchor='center')
        
        # Store all data for filtering
        self.all_email_data = []
        
        # Get accurate data
        df = self.read_data(self.sheet)
        
        # Ensure all positions exist
        if df.empty:
            df = pd.DataFrame({
                'posisi': self.positions,
                'email': [''] * len(self.positions),
                'last_updated': [''] * len(self.positions)
            })
            self.write_data(self.sheet, df)
        
        # Add missing positions
        for position in self.positions:
            if position not in df['posisi'].values:
                new_row = pd.DataFrame({
                    'posisi': [position],
                    'email': [''],
                    'last_updated': ['']
                })
                df = pd.concat([df, new_row], ignore_index=True)
        
        # Populate data accurately
        for position in self.positions:
            # Determine category
            if position in self.position_categories['Directors & GM']:
                category = "Director/GM"
            elif position in self.position_categories['Managers']:
                category = "Manager"
            elif position in self.position_categories['Senior Officers']:
                category = "Senior Officer"
            else:
                category = "Other"
            
            # Get data accurately
            email_val = ''
            updated_val = '-'
            position_data = df[df['posisi'] == position]
            if not position_data.empty:
                email_raw = position_data.iloc[0]['email']
                updated_raw = position_data.iloc[0]['last_updated']
                
                if pd.notna(email_raw) and str(email_raw).strip() != '':
                    email_val = str(email_raw).strip()
                
                if pd.notna(updated_raw) and str(updated_raw).strip() != '':
                    updated_val = str(updated_raw).strip()
            
            display_email = email_val if email_val else "Not configured"
            
            # Store data for filtering
            row_data = {
                'position': position,
                'category': category,
                'email': display_email,
                'updated': updated_val,
                'actions': 'Edit',
                'has_email': bool(email_val)
            }
            self.all_email_data.append(row_data)
        
        # Initial population
        self.populate_email_tree()
        
        # Double click to edit
        self.email_tree.bind('<Double-1>', self.on_email_double_click)
        
        self.email_tree.pack(fill='both', expand=True)
    
    def populate_email_tree(self, filtered_data=None):
        """Populate email tree with data"""
        # Clear existing items
        for item in self.email_tree.get_children():
            self.email_tree.delete(item)
        
        # Use filtered data or all data
        data_to_show = filtered_data if filtered_data is not None else self.all_email_data
        
        for row_data in data_to_show:
            tags = ('configured',) if row_data['has_email'] else ('pending',)
            
            self.email_tree.insert("", "end", values=(
                row_data['position'],
                row_data['category'],
                row_data['email'],
                row_data['updated'],
                row_data['actions']
            ), tags=tags)
        
        # Configure styling
        self.email_tree.tag_configure('configured', foreground=self.colors['success'])
        self.email_tree.tag_configure('pending', foreground=self.colors['warning'])
    
    def filter_email_table(self, *args):
        """Filter email table based on search and filter criteria"""
        search_term = self.search_var.get().lower()
        filter_term = self.filter_var.get()
        
        filtered_data = []
        
        for row_data in self.all_email_data:
            # Apply search filter
            if search_term:
                if (search_term not in row_data['position'].lower() and 
                    search_term not in row_data['email'].lower()):
                    continue
            
            # Apply category/status filter
            if filter_term and filter_term != "All":
                if filter_term == "Configured" and not row_data['has_email']:
                    continue
                elif filter_term == "Pending" and row_data['has_email']:
                    continue
                elif filter_term in ["Directors & GM", "Managers", "Senior Officers"]:
                    if filter_term == "Directors & GM" and row_data['category'] != "Director/GM":
                        continue
                    elif filter_term == "Managers" and row_data['category'] != "Manager":
                        continue
                    elif filter_term == "Senior Officers" and row_data['category'] != "Senior Officer":
                        continue
            
            filtered_data.append(row_data)
        
        self.populate_email_tree(filtered_data)
        
        # Double click to edit
        self.email_tree.bind('<Double-1>', self.on_email_double_click)
        
        self.email_tree.pack(fill='both', expand=True)
    
    def on_email_double_click(self, event):
        """Handle email table double click"""
        selection = self.email_tree.selection()
        if selection:
            item = self.email_tree.item(selection[0])
            position = item['values'][0]
            self.show_email_form(position)
    
    def show_email_form(self, edit_position=None):
        """Minimalist email form"""
        # Create modal window
        form_window = tk.Toplevel(self.root)
        form_window.title("Email Configuration")
        form_window.geometry("450x350")
        form_window.configure(bg='white')
        form_window.transient(self.root)
        form_window.grab_set()
        form_window.resizable(False, False)
        
        # Center window
        form_window.update_idletasks()
        x = (form_window.winfo_screenwidth() // 2) - (450 // 2)
        y = (form_window.winfo_screenheight() // 2) - (350 // 2)
        form_window.geometry(f"450x350+{x}+{y}")
        
        # Form content
        content = tk.Frame(form_window, bg='white')
        content.pack(fill='both', expand=True, padx=40, pady=40)
        
        # Title
        title_text = "Edit Email" if edit_position else "Add Email"
        title = tk.Label(content,
                        text=title_text,
                        font=('Inter', 20, 'bold'),
                        bg='white',
                        fg=self.colors['primary'])
        title.pack(pady=(0, 24))
        
        # Get current data if editing
        current_email = ''
        if edit_position:
            df = self.read_data(self.sheet)
            if not df.empty and edit_position in df['posisi'].values:
                current_email = df[df['posisi'] == edit_position]['email'].iloc[0]
                if pd.isna(current_email):
                    current_email = ''
        
        # Position field
        tk.Label(content,
                text="Position",
                font=('Inter', 11),
                bg='white',
                fg=self.colors['text_primary']).pack(anchor='w', pady=(0, 8))
        
        position_var = tk.StringVar()
        position_combo = ttk.Combobox(content,
                                    textvariable=position_var,
                                    values=self.positions,
                                    font=('Inter', 10),
                                    style='Minimal.TCombobox',
                                    state='readonly' if edit_position else 'normal')
        position_combo.pack(fill='x', pady=(0, 16), ipady=8)
        
        if edit_position:
            position_var.set(edit_position)
        
        # Email field
        tk.Label(content,
                text="Email Address",
                font=('Inter', 11),
                bg='white',
                fg=self.colors['text_primary']).pack(anchor='w', pady=(0, 8))
        
        email_entry = tk.Entry(content,
                             font=('Inter', 11),
                             relief='solid',
                             borderwidth=1,
                             bg='white')
        email_entry.pack(fill='x', pady=(0, 24), ipady=8)
        email_entry.insert(0, current_email)
        
        # Buttons
        btn_frame = tk.Frame(content, bg='white')
        btn_frame.pack(fill='x')
        
        # Save function
        def save_email():
            position = position_var.get().strip()
            email = email_entry.get().strip()
            
            if not position:
                messagebox.showerror("Error", "Please select a position!")
                return
            if not email:
                messagebox.showerror("Error", "Please enter an email address!")
                return
            if '@' not in email or '.' not in email.split('@')[-1]:
                messagebox.showerror("Error", "Please enter a valid email!")
                return
            
            # Update data
            df = self.read_data(self.sheet)
            if df.empty:
                df = pd.DataFrame({
                    'posisi': self.positions,
                    'email': [''] * len(self.positions),
                    'last_updated': [''] * len(self.positions)
                })
            
            timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
            if position in df['posisi'].values:
                df.loc[df['posisi'] == position, 'email'] = email
                df.loc[df['posisi'] == position, 'last_updated'] = timestamp
            else:
                new_row = pd.DataFrame({
                    'posisi': [position],
                    'email': [email],
                    'last_updated': [timestamp]
                })
                df = pd.concat([df, new_row], ignore_index=True)
            
            if self.write_data(self.sheet, df):
                messagebox.showinfo("Success", f"Email saved for '{position}'!")
                form_window.destroy()
                # Refresh current view
                if hasattr(self, 'active_nav'):
                    if self.active_nav == 'dashboard':
                        self.show_dashboard()
                    elif self.active_nav == 'email':
                        self.show_email_management()
            else:
                messagebox.showerror("Error", "Failed to save to Google Sheets!")
        
        # Delete function (only for edit mode)
        def delete_email():
            if edit_position:
                if messagebox.askyesno("Confirm Delete", f"Delete email for '{edit_position}'?"):
                    df = self.read_data(self.sheet)
                    if not df.empty and edit_position in df['posisi'].values:
                        df.loc[df['posisi'] == edit_position, 'email'] = ''
                        df.loc[df['posisi'] == edit_position, 'last_updated'] = datetime.now().strftime("%d/%m/%Y %H:%M")
                        
                        if self.write_data(self.sheet, df):
                            messagebox.showinfo("Success", f"Email deleted for '{edit_position}'!")
                            form_window.destroy()
                            # Refresh current view
                            if hasattr(self, 'active_nav'):
                                if self.active_nav == 'dashboard':
                                    self.show_dashboard()
                                elif self.active_nav == 'email':
                                    self.show_email_management()
                        else:
                            messagebox.showerror("Error", "Failed to delete!")
        
        # Save button
        save_btn = tk.Button(btn_frame,
                           text="Save",
                           font=('Inter', 11, 'bold'),
                           bg=self.colors['accent'],
                           fg='white',
                           border=0,
                           relief='flat',
                           padx=20,
                           pady=10,
                           cursor='hand2',
                           command=save_email)
        save_btn.pack(side='left', padx=(0, 8))
        
        # Delete button (only for edit mode)
        if edit_position:
            delete_btn = tk.Button(btn_frame,
                                 text="Delete",
                                 font=('Inter', 11),
                                 bg=self.colors['danger'],
                                 fg='white',
                                 border=0,
                                 relief='flat',
                                 padx=20,
                                 pady=10,
                                 cursor='hand2',
                                 command=delete_email)
            delete_btn.pack(side='left', padx=(0, 8))
        
        # Cancel button
        cancel_btn = tk.Button(btn_frame,
                             text="Cancel",
                             font=('Inter', 11),
                             bg=self.colors['surface'],
                             fg=self.colors['text_primary'],
                             border=0,
                             relief='flat',
                             padx=20,
                             pady=10,
                             cursor='hand2',
                             command=form_window.destroy)
        cancel_btn.pack(side='right')
        
        # Focus management
        if edit_position:
            email_entry.focus()
            email_entry.select_range(0, 'end')
        else:
            position_combo.focus()
        
        # Enter key binding
        email_entry.bind('<Return>', lambda e: save_email())
    
    def show_bulk_import(self):
        """Minimalist bulk import interface"""
        bulk_window = tk.Toplevel(self.root)
        bulk_window.title("Bulk Import")
        bulk_window.geometry("500x400")
        bulk_window.configure(bg='white')
        bulk_window.transient(self.root)
        bulk_window.grab_set()
        
        # Center window
        bulk_window.update_idletasks()
        x = (bulk_window.winfo_screenwidth() // 2) - (500 // 2)
        y = (bulk_window.winfo_screenheight() // 2) - (400 // 2)
        bulk_window.geometry(f"500x400+{x}+{y}")
        
        # Content
        content = tk.Frame(bulk_window, bg='white')
        content.pack(fill='both', expand=True, padx=40, pady=40)
        
        # Title
        title = tk.Label(content,
                        text="Bulk Import",
                        font=('Inter', 20, 'bold'),
                        bg='white',
                        fg=self.colors['primary'])
        title.pack(pady=(0, 16))
        
        # Instructions
        instructions = tk.Label(content,
                              text="Format: Position | Email\nOne entry per line\nExample: Manager Pemeliharaan | manager@company.com",
                              font=('Inter', 10),
                              bg='white',
                              fg=self.colors['text_secondary'],
                              justify='left')
        instructions.pack(anchor='w', pady=(0, 16))
        
        # Text area
        text_frame = tk.Frame(content, bg='white')
        text_frame.pack(fill='both', expand=True, pady=(0, 16))
        
        text_scroll = tk.Scrollbar(text_frame)
        text_scroll.pack(side='right', fill='y')
        
        bulk_text = tk.Text(text_frame,
                           font=('Consolas', 10),
                           yscrollcommand=text_scroll.set,
                           relief='solid',
                           borderwidth=1,
                           bg='white')
        bulk_text.pack(fill='both', expand=True)
        text_scroll.config(command=bulk_text.yview)
        
        # Buttons
        btn_frame = tk.Frame(content, bg='white')
        btn_frame.pack(fill='x')
        
        def process_import():
            content_text = bulk_text.get("1.0", "end-1c").strip()
            if not content_text:
                messagebox.showerror("Error", "Please enter data to import!")
                return
            
            lines = content_text.split('\n')
            successful = 0
            errors = []
            
            df = self.read_data(self.sheet)
            if df.empty:
                df = pd.DataFrame({
                    'posisi': self.positions,
                    'email': [''] * len(self.positions),
                    'last_updated': [''] * len(self.positions)
                })
            
            timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
            
            for i, line in enumerate(lines, 1):
                line = line.strip()
                if not line:
                    continue
                
                if '|' not in line:
                    errors.append(f"Line {i}: Invalid format (missing |)")
                    continue
                
                parts = line.split('|', 1)
                if len(parts) != 2:
                    errors.append(f"Line {i}: Invalid format")
                    continue
                
                position = parts[0].strip()
                email = parts[1].strip()
                
                if position not in self.positions:
                    errors.append(f"Line {i}: Position '{position}' not found")
                    continue
                
                if '@' not in email or '.' not in email.split('@')[-1]:
                    errors.append(f"Line {i}: Invalid email '{email}'")
                    continue
                
                # Update dataframe
                if position in df['posisi'].values:
                    df.loc[df['posisi'] == position, 'email'] = email
                    df.loc[df['posisi'] == position, 'last_updated'] = timestamp
                    successful += 1
            
            if successful > 0:
                if self.write_data(self.sheet, df):
                    result_msg = f"Import completed!\n\nSuccessful: {successful}\nErrors: {len(errors)}"
                    if errors:
                        result_msg += f"\n\nFirst few errors:\n" + "\n".join(errors[:3])
                    messagebox.showinfo("Import Results", result_msg)
                    if not errors:
                        bulk_window.destroy()
                        # Refresh current view
                        self.refresh_current_view()
                else:
                    messagebox.showerror("Error", "Failed to save to Google Sheets!")
            else:
                messagebox.showerror("Error", "No valid emails imported!")
        
        import_btn = tk.Button(btn_frame,
                             text="Import",
                             font=('Inter', 11, 'bold'),
                             bg=self.colors['accent'],
                             fg='white',
                             border=0,
                             relief='flat',
                             padx=20,
                             pady=10,
                             cursor='hand2',
                             command=process_import)
        import_btn.pack(side='left', padx=(0, 8))
        
        clear_btn = tk.Button(btn_frame,
                            text="Clear",
                            font=('Inter', 11),
                            bg=self.colors['surface'],
                            fg=self.colors['text_primary'],
                            border=0,
                            relief='flat',
                            padx=20,
                            pady=10,
                            cursor='hand2',
                            command=lambda: bulk_text.delete("1.0", "end"))
        clear_btn.pack(side='left', padx=(0, 8))
        
        close_btn = tk.Button(btn_frame,
                            text="Close",
                            font=('Inter', 11),
                            bg=self.colors['surface'],
                            fg=self.colors['text_primary'],
                            border=0,
                            relief='flat',
                            padx=20,
                            pady=10,
                            cursor='hand2',
                            command=bulk_window.destroy)
        close_btn.pack(side='right')
        
        bulk_text.focus()
    
    def init_google_sheets(self):
        """Initialize Google Sheets connection"""
        try:
            gc = gspread.service_account(filename='credentials/credentials.json')
            sheet_url = "https://docs.google.com/spreadsheets/d/1LoAzVPBMJo08uPHR7MdzGEXmaoFXMrTSKS0vl099qrU/edit?gid=0#gid=0"
            sheet = gc.open_by_url(sheet_url).sheet1
            return sheet
        except FileNotFoundError:
            print("Error: credentials/credentials.json file not found")
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
            data_to_write = [df.columns.values.tolist()] + df.values.tolist()
            sheet.update(data_to_write)
            return True
        except Exception as e:
            print(f"Error writing data: {str(e)}")
            return False
    
    def run(self):
        """Run the application"""
        def on_closing():
            if self.logged_in:
                if messagebox.askokcancel("Confirm Exit", "Are you sure you want to exit?"):
                    self.root.destroy()
            else:
                self.root.destroy()
        
        self.root.protocol("WM_DELETE_WINDOW", on_closing)
        self.root.mainloop()


# Main execution
if __name__ == "__main__":
    print("🚀 Starting Minimalist Email Manager v2.0")
    print("=" * 50)
    
    try:
        app = EmailManagerApp()
        app.run()
    except KeyboardInterrupt:
        print("\n\n❌ Application stopped by user")
    except Exception as e:
        print(f"\n\n💥 Fatal error: {str(e)}")
        print("Please check your configuration and try again.")
    finally:
        print("\n👋 Thank you for using Email Manager v2.0!")
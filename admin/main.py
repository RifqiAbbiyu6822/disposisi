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
        """Initialize the Enhanced Email Manager Application with complete position support"""
        self.root = tk.Tk()
        self.root.title("Email Manager Dashboard - Complete Edition")
        self.root.geometry("1600x1000")  # Increased size for more positions
        self.root.configure(bg='#f8fafc')
        self.root.minsize(1200, 800)
        
        # Initialize variables
        self.logged_in = False
        self.sheet = None
        self.admin_password_hash = self.hash_password("admin123")  # Default password
        
        # ENHANCED: Complete positions with Senior Officers
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
            # Senior Officers
            "Senior Officer Pemeliharaan 1",
            "Senior Officer Pemeliharaan 2",
            "Senior Officer Operasional 1",
            "Senior Officer Operasional 2",
            "Senior Officer Administrasi 1",
            "Senior Officer Administrasi 2",
            "Senior Officer Keuangan 1",
            "Senior Officer Keuangan 2"
        ]
        
        # Position grouping for better organization
        self.position_groups = {
            "Directors & GM": [
                "Direktur Utama",
                "Direktur Keuangan", 
                "Direktur Teknik",
                "GM Keuangan & Administrasi",
                "GM Operasional & Pemeliharaan"
            ],
            "Managers": [
                "Manager Pemeliharaan",
                "Manager Operasional",
                "Manager Administrasi",
                "Manager Keuangan"
            ],
            "Senior Officers": [
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
        
        # Abbreviation mapping for display
        self.abbreviation_map = {
            "Manager Pemeliharaan": "Manager pml",
            "Manager Operasional": "Manager ops",
            "Manager Administrasi": "Manager adm",
            "Manager Keuangan": "Manager keu"
        }
        
        # Setup modern UI
        self.setup_modern_styles()
        self.create_main_container()
        
        # Show login initially
        self.show_login()
    
    def setup_modern_styles(self):
        """Setup modern dark/light theme styles"""
        self.colors = {
            'primary': '#1e293b',
            'secondary': '#475569',
            'accent': '#3b82f6',
            'accent_hover': '#2563eb',
            'success': '#10b981',
            'success_hover': '#059669',
            'warning': '#f59e0b',
            'danger': '#ef4444',
            'danger_hover': '#dc2626',
            'info': '#06b6d4',
            'light': '#f8fafc',
            'white': '#ffffff',
            'gray_50': '#f9fafb',
            'gray_100': '#f3f4f6',
            'gray_200': '#e5e7eb',
            'gray_300': '#d1d5db',
            'gray_400': '#9ca3af',
            'gray_500': '#6b7280',
            'gray_600': '#4b5563',
            'gray_700': '#374151',
            'gray_800': '#1f2937',
            'gray_900': '#111827',
            'blue_50': '#eff6ff',
            'blue_100': '#dbeafe',
            'blue_500': '#3b82f6',
            'blue_600': '#2563eb',
            'green_50': '#f0fdf4',
            'green_100': '#dcfce7',
            'green_500': '#22c55e',
            'purple_50': '#faf5ff',
            'purple_100': '#f3e8ff',
            'purple_500': '#a855f7',
            'purple_600': '#9333ea'
        }
        
        self.style = ttk.Style()        
        self.style.theme_use('clam')
        
        # Configure modern styles
        self.style.configure('Title.TLabel', 
                           font=('Inter', 24, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 24, 'bold'), 
                           background=self.colors['light'], 
                           foreground=self.colors['primary'])
        
        self.style.configure('Heading.TLabel', 
                           font=('Inter', 16, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 16, 'bold'), 
                           background=self.colors['white'], 
                           foreground=self.colors['gray_800'])
        
        self.style.configure('Card.TFrame',
                           background=self.colors['white'],
                           relief='solid',
                           borderwidth=1)
        
        # Enhanced button styles
        self.style.configure('Primary.TButton', 
                           font=('Inter', 11, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 11, 'bold'),
                           foreground='white',
                           background=self.colors['accent'],
                           borderwidth=0,
                           relief='flat',
                           padding=(20, 12))
        
        self.style.map('Primary.TButton',
                      background=[('active', self.colors['accent_hover']),
                                ('pressed', self.colors['accent_hover'])])
        
        self.style.configure('Success.TButton', 
                           font=('Inter', 11, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 11, 'bold'),
                           foreground='white',
                           background=self.colors['success'],
                           borderwidth=0,
                           relief='flat',
                           padding=(20, 12))
        
        self.style.map('Success.TButton',
                      background=[('active', self.colors['success_hover']),
                                ('pressed', self.colors['success_hover'])])
        
        self.style.configure('Danger.TButton', 
                           font=('Inter', 11, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 11, 'bold'),
                           foreground='white',
                           background=self.colors['danger'],
                           borderwidth=0,
                           relief='flat',
                           padding=(20, 12))
        
        self.style.map('Danger.TButton',
                      background=[('active', self.colors['danger_hover']),
                                ('pressed', self.colors['danger_hover'])])
        
        # Enhanced entry styles
        self.style.configure('Modern.TEntry',
                           font=('Inter', 11) if 'Inter' in tk.font.families() else ('Segoe UI', 11),
                           fieldbackground=self.colors['white'],
                           bordercolor=self.colors['gray_300'],
                           lightcolor=self.colors['gray_300'],
                           darkcolor=self.colors['gray_300'],
                           insertcolor=self.colors['accent'],
                           padding=(12, 10),
                           relief='solid',
                           borderwidth=2)
        
        self.style.map('Modern.TEntry',
                      bordercolor=[('focus', self.colors['accent'])],
                      lightcolor=[('focus', self.colors['accent'])],
                      darkcolor=[('focus', self.colors['accent'])])
        
        # Modern treeview
        self.style.configure("Modern.Treeview",
                           background=self.colors['white'],
                           foreground=self.colors['gray_900'],
                           rowheight=40,
                           fieldbackground=self.colors['white'],
                           borderwidth=1,
                           relief='solid',
                           font=('Inter', 10) if 'Inter' in tk.font.families() else ('Segoe UI', 10))
        
        self.style.configure("Modern.Treeview.Heading",
                           background=self.colors['gray_50'],
                           foreground=self.colors['gray_700'],
                           relief='flat',
                           borderwidth=0,
                           font=('Inter', 11, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 11, 'bold'),
                           padding=(12, 8))
        
        self.style.map("Modern.Treeview",
                      background=[('selected', self.colors['accent'])],
                      foreground=[('selected', 'white')])
        
        self.style.map("Modern.Treeview.Heading",
                      background=[('active', self.colors['gray_100'])])
    
    def create_main_container(self):
        """Create the main application container with modern design"""
        # Modern header with gradient effect
        self.header_frame = tk.Frame(self.root, bg=self.colors['primary'], height=100)
        self.header_frame.pack(fill='x', side='top')
        self.header_frame.pack_propagate(False)
        
        # Navigation bar
        self.nav_frame = tk.Frame(self.root, bg=self.colors['gray_800'], height=60)
        self.nav_frame.pack(fill='x', side='top')
        self.nav_frame.pack_propagate(False)
        
        # Main content area with modern styling
        self.content_frame = tk.Frame(self.root, bg=self.colors['light'])
        self.content_frame.pack(fill='both', expand=True, padx=0, pady=0)
        
    def create_header(self, title="Email Manager Dashboard - Complete Edition"):
        """Create modern header with enhanced design"""
        for widget in self.header_frame.winfo_children():
            widget.destroy()
            
        # Header content container
        header_content = tk.Frame(self.header_frame, bg=self.colors['primary'])
        header_content.pack(fill='both', expand=True, padx=30, pady=20)
        
        # Left side - Title and subtitle
        left_frame = tk.Frame(header_content, bg=self.colors['primary'])
        left_frame.pack(side='left', fill='y')
        
        # Main title
        title_label = tk.Label(left_frame, 
                              text="📧 Email Manager",
                              font=('Inter', 24, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 24, 'bold'),
                              bg=self.colors['primary'],
                              fg='white')
        title_label.pack(anchor='w')
        
        # Subtitle
        subtitle_label = tk.Label(left_frame,
                                 text="Complete Edition - All Positions Supported",
                                 font=('Inter', 12) if 'Inter' in tk.font.families() else ('Segoe UI', 12),
                                 bg=self.colors['primary'],
                                 fg=self.colors['gray_300'])
        subtitle_label.pack(anchor='w', pady=(5, 0))
        
        # Right side - Enhanced status and actions
        if self.logged_in:
            right_frame = tk.Frame(header_content, bg=self.colors['primary'])
            right_frame.pack(side='right', fill='y')
            
            # Status indicators container
            status_container = tk.Frame(right_frame, bg=self.colors['primary'])
            status_container.pack(anchor='e')
            
            # Connection status with icon
            status_frame = tk.Frame(status_container, bg=self.colors['primary'])
            status_frame.pack(anchor='e', pady=(0, 5))
            
            tk.Label(status_frame,
                    text="🟢",
                    font=('Arial', 16),
                    bg=self.colors['primary']).pack(side='left')
            
            tk.Label(status_frame,
                    text="Connected & Active",
                    font=('Inter', 11, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 11, 'bold'),
                    bg=self.colors['primary'],
                    fg=self.colors['success']).pack(side='left', padx=(5, 0))
            
            # Position count with enhanced display
            total_positions = len(self.positions)
            positions_frame = tk.Frame(status_container, bg=self.colors['primary'])
            positions_frame.pack(anchor='e', pady=(0, 5))
            
            tk.Label(positions_frame,
                    text="👥",
                    font=('Arial', 14),
                    bg=self.colors['primary']).pack(side='left')
            
            tk.Label(positions_frame,
                    text=f"{total_positions} Total Positions",
                    font=('Inter', 10) if 'Inter' in tk.font.families() else ('Segoe UI', 10),
                    bg=self.colors['primary'],
                    fg='white').pack(side='left', padx=(5, 0))
            
            # Enhanced logout button
            logout_btn = tk.Button(right_frame,
                                 text="🚪 Sign Out",
                                 font=('Inter', 11, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 11, 'bold'),
                                 bg=self.colors['danger'],
                                 fg='white',
                                 border=0,
                                 padx=25,
                                 pady=12,
                                 cursor='hand2',
                                 relief='flat',
                                 command=self.logout)
            logout_btn.pack(anchor='e', pady=(10, 0))
            
            # Hover effects
            def on_logout_enter(e):
                logout_btn.config(bg=self.colors['danger_hover'])
            def on_logout_leave(e):
                logout_btn.config(bg=self.colors['danger'])
            
            logout_btn.bind("<Enter>", on_logout_enter)
            logout_btn.bind("<Leave>", on_logout_leave)
    
    def create_navigation(self):
        """Create enhanced navigation bar with modern design"""
        for widget in self.nav_frame.winfo_children():
            widget.destroy()
            
        if not self.logged_in:
            return
        
        # Navigation content
        nav_content = tk.Frame(self.nav_frame, bg=self.colors['gray_800'])
        nav_content.pack(fill='both', expand=True, padx=30, pady=0)
        
        nav_buttons = [
            ("📊 Dashboard", self.show_dashboard, 'dashboard', self.colors['accent']),
            ("📧 Email Management", self.show_email_management, 'email', self.colors['success']),
            ("👨‍💼 All Positions", self.show_all_positions, 'positions', self.colors['purple_500']),
            ("🏢 By Department", self.show_by_department, 'department', self.colors['info']),
            ("🔒 Security", self.show_change_password, 'password', self.colors['warning'])
        ]
        
        self.nav_buttons = {}
        for text, command, key, color in nav_buttons:
            btn_frame = tk.Frame(nav_content, bg=self.colors['gray_800'])
            btn_frame.pack(side='left', padx=2)
            
            btn = tk.Button(btn_frame,
                          text=text,
                          font=('Inter', 11, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 11, 'bold'),
                          bg=self.colors['gray_800'],
                          fg='white',
                          border=0,
                          padx=25,
                          pady=15,
                          cursor='hand2',
                          relief='flat',
                          command=command)
            btn.pack()
            self.nav_buttons[key] = btn
            
            # Enhanced hover effects with color coding
            def make_hover_handler(button, accent_color):
                def on_enter(e):
                    button.config(bg=accent_color)
                def on_leave(e):
                    if getattr(button, 'active', False):
                        button.config(bg=accent_color)
                    else:
                        button.config(bg=self.colors['gray_800'])
                return on_enter, on_leave
            
            enter_handler, leave_handler = make_hover_handler(btn, color)
            btn.bind("<Enter>", enter_handler)
            btn.bind("<Leave>", leave_handler)
    
    def set_active_nav(self, active_key):
        """Set active navigation button with color coding"""
        if hasattr(self, 'nav_buttons'):
            nav_colors = {
                'dashboard': self.colors['accent'],
                'email': self.colors['success'],
                'positions': self.colors['purple_500'],
                'department': self.colors['info'],
                'password': self.colors['warning']
            }
            
            for key, btn in self.nav_buttons.items():
                if key == active_key:
                    btn.config(bg=nav_colors.get(key, self.colors['accent']))
                    btn.active = True
                else:
                    btn.config(bg=self.colors['gray_800'])
                    btn.active = False
    
    def clear_content(self):
        """Clear current content frame"""
        for widget in self.content_frame.winfo_children():
            widget.destroy()
            
    def show_login(self):
        """Show enhanced login screen"""
        self.create_header("Login - Email Manager Complete Edition")
        self.clear_content()
        
        # Hide navigation
        self.nav_frame.pack_forget()
        
        # Login container with modern design
        login_outer = tk.Frame(self.content_frame, bg=self.colors['light'])
        login_outer.pack(expand=True)
        
        # Card container with shadow effect
        card_container = tk.Frame(login_outer, bg=self.colors['gray_200'])
        card_container.pack()
        
        # Main login card
        login_card = tk.Frame(card_container, bg=self.colors['white'], width=450, height=400)
        login_card.pack(padx=3, pady=3)
        login_card.pack_propagate(False)
        
        # Header section
        header_section = tk.Frame(login_card, bg=self.colors['accent'], height=120)
        header_section.pack(fill='x')
        header_section.pack_propagate(False)
        
        # Login icon and title
        icon_frame = tk.Frame(header_section, bg=self.colors['accent'])
        icon_frame.pack(expand=True)
        
        tk.Label(icon_frame, 
                text="🔐",
                font=('Arial', 48),
                bg=self.colors['accent'],
                fg='white').pack(pady=(20, 10))
        
        tk.Label(icon_frame, 
                text="Admin Access",
                font=('Inter', 18, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 18, 'bold'),
                bg=self.colors['accent'],
                fg='white').pack()
        
        # Form section
        form_section = tk.Frame(login_card, bg=self.colors['white'])
        form_section.pack(fill='both', expand=True, padx=40, pady=30)
        
        # Password field with modern styling
        tk.Label(form_section,
                text="Password",
                font=('Inter', 12, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 12, 'bold'),
                bg=self.colors['white'],
                fg=self.colors['gray_700']).pack(anchor='w', pady=(0, 8))
        
        # Password input with border
        password_container = tk.Frame(form_section, bg=self.colors['gray_300'], height=50)
        password_container.pack(fill='x', pady=(0, 20))
        password_container.pack_propagate(False)
        
        self.password_entry = tk.Entry(password_container,
                                     font=('Inter', 12) if 'Inter' in tk.font.families() else ('Segoe UI', 12),
                                     show="*",
                                     relief='flat',
                                     borderwidth=0,
                                     bg=self.colors['white'],
                                     fg=self.colors['gray_900'])
        self.password_entry.pack(fill='both', padx=2, pady=2)
        self.password_entry.bind('<Return>', lambda e: self.try_login())
        
        # Focus effects for password field
        def on_focus_in(e):
            password_container.config(bg=self.colors['accent'])
        def on_focus_out(e):
            password_container.config(bg=self.colors['gray_300'])
        
        self.password_entry.bind('<FocusIn>', on_focus_in)
        self.password_entry.bind('<FocusOut>', on_focus_out)
        
        # Login button with gradient effect
        login_btn = tk.Button(form_section,
                            text="🔓 Sign In",
                            font=('Inter', 12, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 12, 'bold'),
                            bg=self.colors['accent'],
                            fg='white',
                            border=0,
                            relief='flat',
                            padx=40,
                            pady=15,
                            cursor='hand2',
                            command=self.try_login)
        login_btn.pack(pady=(10, 0))
        
        # Hover effect for login button
        def on_enter(e):
            login_btn.config(bg=self.colors['accent_hover'])
        def on_leave(e):
            login_btn.config(bg=self.colors['accent'])
        
        login_btn.bind("<Enter>", on_enter)
        login_btn.bind("<Leave>", on_leave)
        
        # Info text
        info_label = tk.Label(form_section,
                             text="Complete Email Management System\nwith 18 Position Support",
                             font=('Inter', 9) if 'Inter' in tk.font.families() else ('Segoe UI', 9),
                             bg=self.colors['white'],
                             fg=self.colors['gray_500'],
                             justify='center')
        info_label.pack(pady=(20, 0))
        
        # Focus on password entry
        self.password_entry.focus()
    
    def try_login(self):
        """Try to login with password verification"""
        password = self.password_entry.get()
        if self.verify_password(password, self.admin_password_hash):
            self.logged_in = True
            self.init_google_sheets()
            self.show_dashboard()
        else:
            messagebox.showerror("Login Error", "Password salah!")
    
    def logout(self):
        """Logout and return to login screen"""
        self.logged_in = False
        self.sheet = None
        self.show_login()
        
    def show_dashboard(self):
        """ENHANCED: Show comprehensive dashboard with all position statistics"""
        self.create_header("Complete Dashboard - All Positions Overview")
        self.nav_frame.pack(fill='x', side='top', after=self.header_frame)
        self.create_navigation()
        self.set_active_nav('dashboard')
        self.clear_content()
        
        # Main dashboard container
        dashboard_container = tk.Frame(self.content_frame, bg=self.colors['light'])
        dashboard_container.pack(fill='both', expand=True, padx=30, pady=20)
        
        # Read data from sheet
        df = self.read_data(self.sheet)
        
        if df.empty:
            # Initialize data if empty
            df = pd.DataFrame({
                'posisi': self.positions,
                'email': [''] * len(self.positions),
                'last_updated': [''] * len(self.positions)
            })
            self.write_data(self.sheet, df)
        
        # Calculate statistics
        total_positions = len(df)
        filled_emails = df['email'].apply(lambda x: x != '' and pd.notna(x)).sum() if 'email' in df.columns else 0
        empty_emails = total_positions - filled_emails
        completion_rate = (filled_emails / total_positions * 100) if total_positions > 0 else 0
        
        # Position breakdown
        directors_gm = len([pos for pos in self.positions if pos.startswith(("Direktur", "GM"))])
        managers = len([pos for pos in self.positions if pos.startswith("Manager")])
        senior_officers = len([pos for pos in self.positions if pos.startswith("Senior Officer")])
        
        # Stats cards container
        stats_container = tk.Frame(dashboard_container, bg=self.colors['light'])
        stats_container.pack(fill='x', pady=(0, 30))
        
        # Main statistics row
        self.create_modern_stat_card(stats_container, "Total Positions", str(total_positions), self.colors['primary'], "👥", 0, 0)
        self.create_modern_stat_card(stats_container, "Email Configured", str(filled_emails), self.colors['success'], "✅", 0, 1)
        self.create_modern_stat_card(stats_container, "Pending Setup", str(empty_emails), self.colors['warning'], "⚠️", 0, 2)
        self.create_modern_stat_card(stats_container, "Completion Rate", f"{completion_rate:.1f}%", self.colors['info'], "📊", 0, 3)
        
        # Position breakdown row
        stats_container.grid_rowconfigure(1, weight=0, pady=(20, 0))
        self.create_modern_stat_card(stats_container, "Directors & GM", str(directors_gm), self.colors['accent'], "🏢", 1, 0)
        self.create_modern_stat_card(stats_container, "Managers", str(managers), self.colors['purple_500'], "👔", 1, 1)
        self.create_modern_stat_card(stats_container, "Senior Officers", str(senior_officers), self.colors['purple_600'], "👨‍💼", 1, 2)
        self.create_modern_stat_card(stats_container, "Avg per Dept", "2-3", self.colors['gray_600'], "📈", 1, 3)
        
        # Configure grid weights
        for i in range(4):
            stats_container.grid_columnconfigure(i, weight=1)
        
        # Enhanced data table
        self.create_enhanced_dashboard_table(dashboard_container, df)
    
    def create_modern_stat_card(self, parent, title, value, color, icon, row, col):
        """Create a modern statistics card with enhanced design"""
        # Card container with shadow
        card_shadow = tk.Frame(parent, bg=self.colors['gray_300'])
        card_shadow.grid(row=row, column=col, sticky='ew', padx=15, pady=10)
        
        # Main card
        card = tk.Frame(card_shadow, bg=self.colors['white'], height=120)
        card.pack(fill='both', expand=True, padx=2, pady=2)
        card.pack_propagate(False)
        
        # Color accent bar
        accent_bar = tk.Frame(card, bg=color, height=4)
        accent_bar.pack(fill='x')
        
        # Card content
        content = tk.Frame(card, bg=self.colors['white'])
        content.pack(fill='both', expand=True, padx=20, pady=15)
        
        # Top row - icon and value
        top_row = tk.Frame(content, bg=self.colors['white'])
        top_row.pack(fill='x')
        
        # Icon on the right
        icon_label = tk.Label(top_row,
                             text=icon,
                             font=('Arial', 24),
                             bg=self.colors['white'],
                             fg=color)
        icon_label.pack(side='right')
        
        # Value on the left
        value_label = tk.Label(top_row,
                              text=value,
                              font=('Inter', 28, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 28, 'bold'),
                              bg=self.colors['white'],
                              fg=color)
        value_label.pack(side='left', anchor='w')
        
        # Title
        title_label = tk.Label(content,
                              text=title,
                              font=('Inter', 11) if 'Inter' in tk.font.families() else ('Segoe UI', 11),
                              bg=self.colors['white'],
                              fg=self.colors['gray_600'])
        title_label.pack(anchor='w', pady=(8, 0))
    
    def create_enhanced_dashboard_table(self, parent, df):
        """Create enhanced table with modern design and position grouping"""
        # Table container
        table_container = tk.Frame(parent, bg=self.colors['white'])
        table_container.pack(fill='both', expand=True, pady=(20, 0))
        
        # Table header
        header_frame = tk.Frame(table_container, bg=self.colors['white'])
        header_frame.pack(fill='x', padx=30, pady=20)
        
        tk.Label(header_frame,
                text="📋 Email Configuration Status - All Positions",
                font=('Inter', 18, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 18, 'bold'),
                bg=self.colors['white'],
                fg=self.colors['gray_800']).pack(side='left')
        
        # Quick actions
        actions_frame = tk.Frame(header_frame, bg=self.colors['white'])
        actions_frame.pack(side='right')
        
        refresh_btn = tk.Button(actions_frame,
                               text="🔄 Refresh",
                               font=('Inter', 10, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 10, 'bold'),
                               bg=self.colors['accent'],
                               fg='white',
                               border=0,
                               relief='flat',
                               padx=20,
                               pady=8,
                               cursor='hand2',
                               command=lambda: self.show_dashboard())
        refresh_btn.pack(side='left', padx=(0, 10))
        
        export_btn = tk.Button(actions_frame,
                              text="📤 Export Data",
                              font=('Inter', 10, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 10, 'bold'),
                              bg=self.colors['success'],
                              fg='white',
                              border=0,
                              relief='flat',
                              padx=20,
                              pady=8,
                              cursor='hand2',
                              command=self.export_dashboard_data)
        export_btn.pack(side='left')
        
        # Create modern treeview
        tree_frame = tk.Frame(table_container, bg=self.colors['white'])
        tree_frame.pack(fill='both', expand=True, padx=30, pady=(0, 30))
        
        # Create treeview with modern styling
        columns = ('Position', 'Email', 'Status', 'Last Updated')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', style='Modern.Treeview')
        
        # Configure columns
        tree.heading('Position', text='Position')
        tree.heading('Email', text='Email Address')
        tree.heading('Status', text='Status')
        tree.heading('Last Updated', text='Last Updated')
        
        tree.column('Position', width=300, anchor='w')
        tree.column('Email', width=250, anchor='w')
        tree.column('Status', width=120, anchor='center')
        tree.column('Last Updated', width=150, anchor='center')
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack treeview and scrollbar
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Populate treeview with grouped data
        for group_name, positions in self.position_groups.items():
            # Add group header
            tree.insert('', 'end', values=(f"📁 {group_name}", "", "", ""), tags=('group',))
            
            for position in positions:
                # Find position data
                pos_data = df[df['posisi'] == position]
                email = pos_data['email'].iloc[0] if not pos_data.empty else ""
                last_updated = pos_data['last_updated'].iloc[0] if not pos_data.empty else ""
                
                # Determine status
                if email and email.strip():
                    status = "✅ Configured"
                    status_color = self.colors['success']
                else:
                    status = "⚠️ Pending"
                    status_color = self.colors['warning']
                
                # Insert position row
                item = tree.insert('', 'end', values=(f"  {position}", email, status, last_updated))
                
                # Apply color coding based on position type
                if position.startswith("Senior Officer"):
                    tree.tag_configure(item, background=self.colors['purple_50'])
                elif position.startswith("Manager"):
                    tree.tag_configure(item, background=self.colors['blue_50'])
                elif position.startswith(("Direktur", "GM")):
                    tree.tag_configure(item, background=self.colors['green_50'])
        
        # Configure group header styling
        tree.tag_configure('group', font=('Inter', 12, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 12, 'bold'),
                          background=self.colors['gray_100'])
    
    def export_dashboard_data(self):
        """Export dashboard data to CSV"""
        try:
            df = self.read_data(self.sheet)
            if not df.empty:
                filename = f"email_manager_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                df.to_csv(filename, index=False)
                messagebox.showinfo("Export Success", f"Data berhasil diekspor ke {filename}")
            else:
                messagebox.showwarning("No Data", "Tidak ada data untuk diekspor")
        except Exception as e:
            messagebox.showerror("Export Error", f"Gagal mengekspor data: {str(e)}")
    
    def show_all_positions(self):
        """Show all positions management view"""
        self.create_header("All Positions Management")
        self.create_navigation()
        self.set_active_nav('positions')
        self.clear_content()
        
        # Implementation for all positions view
        container = tk.Frame(self.content_frame, bg=self.colors['light'])
        container.pack(fill='both', expand=True, padx=30, pady=20)
        
        tk.Label(container,
                text="👨‍💼 All Positions Management",
                font=('Inter', 20, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 20, 'bold'),
                bg=self.colors['light'],
                fg=self.colors['gray_800']).pack(pady=(0, 20))
        
        # Show position groups
        for group_name, positions in self.position_groups.items():
            self.create_position_group_card(container, group_name, positions)
    
    def show_by_department(self):
        """Show positions grouped by department"""
        self.create_header("Positions by Department")
        self.create_navigation()
        self.set_active_nav('department')
        self.clear_content()
        
        # Implementation for department view
        container = tk.Frame(self.content_frame, bg=self.colors['light'])
        container.pack(fill='both', expand=True, padx=30, pady=20)
        
        tk.Label(container,
                text="🏢 Department Overview",
                font=('Inter', 20, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 20, 'bold'),
                bg=self.colors['light'],
                fg=self.colors['gray_800']).pack(pady=(0, 20))
        
        # Department statistics
        departments = {
            "Pemeliharaan": ["Manager Pemeliharaan", "Senior Officer Pemeliharaan 1", "Senior Officer Pemeliharaan 2"],
            "Operasional": ["Manager Operasional", "Senior Officer Operasional 1", "Senior Officer Operasional 2"],
            "Administrasi": ["Manager Administrasi", "Senior Officer Administrasi 1", "Senior Officer Administrasi 2"],
            "Keuangan": ["Manager Keuangan", "Senior Officer Keuangan 1", "Senior Officer Keuangan 2"]
        }
        
        for dept_name, positions in departments.items():
            self.create_department_card(container, dept_name, positions)
    
    def create_position_group_card(self, parent, group_name, positions):
        """Create a card for position group"""
        card_frame = tk.Frame(parent, bg=self.colors['white'], relief='solid', borderwidth=1)
        card_frame.pack(fill='x', pady=10, padx=20)
        
        # Header
        header = tk.Frame(card_frame, bg=self.colors['gray_50'], height=50)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(header,
                text=f"📁 {group_name} ({len(positions)} positions)",
                font=('Inter', 14, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 14, 'bold'),
                bg=self.colors['gray_50'],
                fg=self.colors['gray_800']).pack(side='left', padx=20, pady=15)
        
        # Positions list
        for i, position in enumerate(positions):
            pos_frame = tk.Frame(card_frame, bg=self.colors['white'])
            pos_frame.pack(fill='x', padx=20, pady=5)
            
            # Position name with icon
            icon = "👔" if position.startswith("Manager") else "👨‍💼" if position.startswith("Senior") else "🏢"
            tk.Label(pos_frame,
                    text=f"{icon} {position}",
                    font=('Inter', 11) if 'Inter' in tk.font.families() else ('Segoe UI', 11),
                    bg=self.colors['white'],
                    fg=self.colors['gray_700']).pack(side='left')
            
            # Action buttons
            edit_btn = tk.Button(pos_frame,
                               text="✏️ Edit",
                               font=('Inter', 9) if 'Inter' in tk.font.families() else ('Segoe UI', 9),
                               bg=self.colors['accent'],
                               fg='white',
                               border=0,
                               relief='flat',
                               padx=15,
                               pady=5,
                               cursor='hand2',
                               command=lambda p=position: self.show_edit_email_form(p))
            edit_btn.pack(side='right')
    
    def create_department_card(self, parent, dept_name, positions):
        """Create a card for department"""
        card_frame = tk.Frame(parent, bg=self.colors['white'], relief='solid', borderwidth=1)
        card_frame.pack(fill='x', pady=10, padx=20)
        
        # Header with department icon
        header = tk.Frame(card_frame, bg=self.colors['blue_50'], height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(header,
                text=f"🏢 {dept_name} Department",
                font=('Inter', 16, 'bold') if 'Inter' in tk.font.families() else ('Segoe UI', 16, 'bold'),
                bg=self.colors['blue_50'],
                fg=self.colors['blue_600']).pack(side='left', padx=20, pady=20)
        
        tk.Label(header,
                text=f"{len(positions)} Positions",
                font=('Inter', 12) if 'Inter' in tk.font.families() else ('Segoe UI', 12),
                bg=self.colors['blue_50'],
                fg=self.colors['blue_600']).pack(side='right', padx=20, pady=20)
        
        # Department positions
        for position in positions:
            pos_frame = tk.Frame(card_frame, bg=self.colors['white'])
            pos_frame.pack(fill='x', padx=20, pady=8)
            
            # Position with role icon
            role_icon = "👔" if position.startswith("Manager") else "👨‍💼"
            tk.Label(pos_frame,
                    text=f"{role_icon} {position}",
                    font=('Inter', 11) if 'Inter' in tk.font.families() else ('Segoe UI', 11),
                    bg=self.colors['white'],
                    fg=self.colors['gray_700']).pack(side='left')
            
            # Status indicator
            status_btn = tk.Button(pos_frame,
                                 text="📧 Configure",
                                 font=('Inter', 9) if 'Inter' in tk.font.families() else ('Segoe UI', 9),
                                 bg=self.colors['success'],
                                 fg='white',
                                 border=0,
                                 relief='flat',
                                 padx=15,
                                 pady=5,
                                 cursor='hand2',
                                 command=lambda p=position: self.show_edit_email_form(p))
            status_btn.pack(side='right')
    
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
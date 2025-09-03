import gspread
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk
import os
from pathlib import Path

class SimpleAdminApp:
    def __init__(self):
        """Initialize Simple Admin Application"""
        print("🚀 Starting Simple Admin App...")
        
        # Create main window
        self.root = tk.Tk()
        self.root.title("Position Management Admin")
        self.root.geometry("1200x800")
        self.root.configure(bg='white')
        
        # Initialize variables
        self.logged_in = False
        self.sheet = None
        
        # Get credentials path
        self.credentials_path = self.find_credentials_file()
        
        # Position data
        self.positions = [
            "Direktur Utama", "Direktur Keuangan", "Direktur Teknik",
            "GM Keuangan & Administrasi", "GM Operasional & Pemeliharaan",
            "Manager Pemeliharaan", "Manager Operasional", "Manager Administrasi", "Manager Keuangan",
            "Senior Officer Pemeliharaan 1", "Senior Officer Pemeliharaan 2",
            "Senior Officer Operasional 1", "Senior Officer Operasional 2",
            "Senior Officer Administrasi 1", "Senior Officer Administrasi 2",
            "Senior Officer Keuangan 1", "Senior Officer Keuangan 2"
        ]
        
        # Create UI
        self.create_ui()
        
        # Show login first
        self.show_login()
        
        print("✅ App initialized successfully")
    
    def find_credentials_file(self):
        """Find credentials.json file in various possible locations"""
        possible_paths = [
            Path(__file__).parent.parent / 'credentials' / 'credentials.json',  # ../credentials/credentials.json
            Path(__file__).parent.parent / 'credentials.json',  # ../credentials.json
            Path.cwd() / 'credentials' / 'credentials.json',  # credentials/credentials.json from current dir
            Path.cwd() / 'credentials.json',  # credentials.json from current dir
        ]
        
        for path in possible_paths:
            if path.exists():
                print(f"✅ Found credentials at: {path}")
                return str(path)
        
        print("⚠️ No credentials.json found in any expected location")
        return None
    
    def create_ui(self):
        """Create simple UI layout"""
        # Main frame
        self.main_frame = tk.Frame(self.root, bg='white')
        self.main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Title
        title = tk.Label(self.main_frame, text="Position Management Admin", 
                        font=("Arial", 24, "bold"), bg='white', fg='black')
        title.pack(pady=(0, 30))
        
        # Content frame
        self.content_frame = tk.Frame(self.main_frame, bg='white')
        self.content_frame.pack(fill='both', expand=True)
    
    def show_login(self):
        """Show simple login screen"""
        print("🔐 Showing login screen...")
        
        # Clear content
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        # Login frame
        login_frame = tk.Frame(self.content_frame, bg='white')
        login_frame.pack(expand=True)
        
        # Login title
        login_title = tk.Label(login_frame, text="Login Required", 
                             font=("Arial", 18, "bold"), bg='white', fg='black')
        login_title.pack(pady=(0, 20))
        
        # Credentials status
        if self.credentials_path:
            cred_status = tk.Label(login_frame, text=f"✅ Credentials: {os.path.basename(self.credentials_path)}", 
                                 font=("Arial", 10), bg='white', fg='green')
        else:
            cred_status = tk.Label(login_frame, text="⚠️ No credentials found - Offline mode only", 
                                 font=("Arial", 10), bg='white', fg='orange')
        cred_status.pack(pady=(0, 20))
        
        # Password label
        pwd_label = tk.Label(login_frame, text="Password:", 
                           font=("Arial", 12), bg='white', fg='black')
        pwd_label.pack(pady=(0, 10))
        
        # Password entry
        self.password_entry = tk.Entry(login_frame, show="*", font=("Arial", 12), width=30)
        self.password_entry.pack(pady=(0, 20))
        self.password_entry.bind('<Return>', lambda e: self.try_login())
        
        # Login button
        login_btn = tk.Button(login_frame, text="Login", 
                            command=self.try_login, 
                            font=("Arial", 12, "bold"),
                            bg='blue', fg='white', padx=30, pady=10)
        login_btn.pack()
        
        # Help text
        help_text = tk.Label(login_frame, text="Default password: admin123", 
                           font=("Arial", 10), bg='white', fg='gray')
        help_text.pack(pady=(20, 0))
        
        # Focus on password entry
        self.password_entry.focus()
    
    def try_login(self):
        """Handle login"""
        password = self.password_entry.get()
        print(f"🔑 Login attempt with password: {password}")
        
        if password == "admin123":
            print("✅ Login successful")
            self.logged_in = True
            
            # Try to connect to Google Sheets
            try:
                self.sheet = self.connect_google_sheets()
                if self.sheet:
                    print("✅ Connected to Google Sheets")
                else:
                    print("⚠️ Using offline mode")
            except Exception as e:
                print(f"❌ Google Sheets error: {e}")
                self.sheet = None
            
            # Show main interface
            self.show_main_interface()
        else:
            print("❌ Login failed")
            messagebox.showerror("Login Failed", "Incorrect password!")
            self.password_entry.delete(0, 'end')
            self.password_entry.focus()
    
    def connect_google_sheets(self):
        """Connect to Google Sheets"""
        if not self.credentials_path:
            print("❌ No credentials file available")
            return None
            
        try:
            gc = gspread.service_account(filename=self.credentials_path)
            sheet_url = "https://docs.google.com/spreadsheets/d/1LoAzVPBMJo08uPHR7MdzGEXmaoFXMrTSKS0vl099qrU"
            sheet = gc.open_by_url(sheet_url).sheet1
            print(f"✅ Successfully connected to Google Sheets")
            return sheet
        except FileNotFoundError:
            print(f"❌ Credentials file not found: {self.credentials_path}")
            return None
        except Exception as e:
            print(f"❌ Google Sheets connection failed: {e}")
            return None
    
    def show_main_interface(self):
        """Show main interface after login"""
        print("🏠 Showing main interface...")
        
        # Clear content
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        # Create main interface
        self.create_main_interface()
    
    def create_main_interface(self):
        """Create main interface with navigation"""
        # Navigation frame
        nav_frame = tk.Frame(self.content_frame, bg='lightgray')
        nav_frame.pack(fill='x', pady=(0, 20))
        
        # Navigation buttons
        dashboard_btn = tk.Button(nav_frame, text="Dashboard", 
                                command=self.show_dashboard,
                                font=("Arial", 12), bg='white', fg='black',
                                padx=20, pady=10)
        dashboard_btn.pack(side='left', padx=10, pady=10)
        
        positions_btn = tk.Button(nav_frame, text="Manage Positions", 
                                command=self.show_positions,
                                font=("Arial", 12), bg='white', fg='black',
                                padx=20, pady=10)
        positions_btn.pack(side='left', padx=10, pady=10)
        
        logout_btn = tk.Button(nav_frame, text="Logout", 
                             command=self.logout,
                             font=("Arial", 12), bg='red', fg='white',
                             padx=20, pady=10)
        logout_btn.pack(side='right', padx=10, pady=10)
        
        # Show dashboard by default
        self.show_dashboard()
    
    def show_dashboard(self):
        """Show dashboard"""
        print("📊 Showing dashboard...")
        
        # Clear content area
        self.clear_content_area()
        
        # Dashboard title
        title = tk.Label(self.content_frame, text="Dashboard", 
                        font=("Arial", 20, "bold"), bg='white', fg='black')
        title.pack(pady=(0, 20))
        
        # Stats frame
        stats_frame = tk.Frame(self.content_frame, bg='white')
        stats_frame.pack(fill='x', pady=(0, 20))
        
        # Get data
        data = self.get_data()
        
        # Show stats
        total_positions = len(self.positions)
        configured_positions = len([row for row in data if row['nama'] and row['email']])
        pending_positions = total_positions - configured_positions
        
        # Stats labels
        tk.Label(stats_frame, text=f"Total Positions: {total_positions}", 
                font=("Arial", 14), bg='white', fg='black').pack(side='left', padx=20)
        
        tk.Label(stats_frame, text=f"Configured: {configured_positions}", 
                font=("Arial", 14), bg='white', fg='green').pack(side='left', padx=20)
        
        tk.Label(stats_frame, text=f"Pending: {pending_positions}", 
                font=("Arial", 14), bg='white', fg='orange').pack(side='left', padx=20)
        
        # Show data table
        self.show_data_table(data)
    
    def show_positions(self):
        """Show positions management"""
        print("👥 Showing positions management...")
        
        # Clear content area
        self.clear_content_area()
        
        # Title
        title = tk.Label(self.content_frame, text="Manage Positions", 
                        font=("Arial", 20, "bold"), bg='white', fg='black')
        title.pack(pady=(0, 20))
        
        # Get data
        data = self.get_data()
        
        # Show data table with edit capabilities
        self.show_editable_table(data)
    
    def clear_content_area(self):
        """Clear the main content area"""
        # Find and clear the main content area (excluding navigation)
        for widget in self.content_frame.winfo_children():
            if isinstance(widget, tk.Frame) and widget.winfo_children():
                # Check if this is the navigation frame
                first_child = widget.winfo_children()[0]
                if isinstance(first_child, tk.Button) and first_child.cget('text') == 'Dashboard':
                    continue  # Skip navigation frame
            widget.destroy()
    
    def get_data(self):
        """Get data from sheet or return default"""
        if self.sheet:
            try:
                records = self.sheet.get_all_records()
                print(f"✅ Retrieved {len(records)} records from Google Sheets")
                
                # Filter out empty rows and ensure proper structure
                filtered_data = []
                for record in records:
                    if record.get('posisi') and record.get('posisi').strip():
                        # Ensure all required fields exist
                        data_row = {
                            'posisi': record.get('posisi', '').strip(),
                            'nama': record.get('nama', '').strip(),
                            'email': record.get('email', '').strip(),
                            'last_updated': record.get('last_updated', '').strip()
                        }
                        filtered_data.append(data_row)
                
                return filtered_data
            except Exception as e:
                print(f"❌ Error reading from sheet: {e}")
                return self.get_default_data()
        
        # Return default data structure
        return self.get_default_data()
    
    def get_default_data(self):
        """Get default data structure for all positions"""
        return [{'posisi': pos, 'nama': '', 'email': '', 'last_updated': ''} 
                for pos in self.positions]
    
    def show_data_table(self, data):
        """Show data in a simple table"""
        # Table frame
        table_frame = tk.Frame(self.content_frame, bg='white')
        table_frame.pack(fill='both', expand=True)
        
        # Headers
        headers = ['Position', 'Name', 'Email', 'Last Updated']
        for i, header in enumerate(headers):
            label = tk.Label(table_frame, text=header, 
                           font=("Arial", 12, "bold"), 
                           bg='lightblue', fg='black',
                           relief='solid', borderwidth=1)
            label.grid(row=0, column=i, sticky='ew', padx=1, pady=1)
        
        # Data rows
        for i, row in enumerate(data, 1):
            tk.Label(table_frame, text=row.get('posisi', ''), 
                    bg='white', relief='solid', borderwidth=1).grid(row=i, column=0, sticky='ew', padx=1, pady=1)
            tk.Label(table_frame, text=row.get('nama', ''), 
                    bg='white', relief='solid', borderwidth=1).grid(row=i, column=1, sticky='ew', padx=1, pady=1)
            tk.Label(table_frame, text=row.get('email', ''), 
                    bg='white', relief='solid', borderwidth=1).grid(row=i, column=2, sticky='ew', padx=1, pady=1)
            tk.Label(table_frame, text=row.get('last_updated', ''), 
                    bg='white', relief='solid', borderwidth=1).grid(row=i, column=3, sticky='ew', padx=1, pady=1)
        
        # Configure grid weights
        for i in range(4):
            table_frame.grid_columnconfigure(i, weight=1)
    
    def show_editable_table(self, data):
        """Show editable table for positions"""
        # Table frame
        table_frame = tk.Frame(self.content_frame, bg='white')
        table_frame.pack(fill='both', expand=True)
        
        # Headers
        headers = ['Position', 'Name', 'Email', 'Actions']
        for i, header in enumerate(headers):
            label = tk.Label(table_frame, text=header, 
                           font=("Arial", 12, "bold"), 
                           bg='lightblue', fg='black',
                           relief='solid', borderwidth=1)
            label.grid(row=0, column=i, sticky='ew', padx=1, pady=1)
        
        # Data rows with edit buttons
        self.entries = {}
        for i, row in enumerate(data, 1):
            pos = row.get('posisi', '')
            
            # Position (read-only)
            tk.Label(table_frame, text=pos, bg='lightgray', 
                    relief='solid', borderwidth=1).grid(row=i, column=0, sticky='ew', padx=1, pady=1)
            
            # Name entry
            name_entry = tk.Entry(table_frame, font=("Arial", 10))
            name_entry.insert(0, row.get('nama', ''))
            name_entry.grid(row=i, column=1, sticky='ew', padx=1, pady=1)
            self.entries[f'name_{i}'] = name_entry
            
            # Email entry
            email_entry = tk.Entry(table_frame, font=("Arial", 10))
            email_entry.insert(0, row.get('email', ''))
            email_entry.grid(row=i, column=2, sticky='ew', padx=1, pady=1)
            self.entries[f'email_{i}'] = email_entry
            
            # Save button
            save_btn = tk.Button(table_frame, text="Save", 
                               command=lambda row_idx=i, pos=pos: self.save_position(row_idx, pos),
                               bg='green', fg='white', font=("Arial", 9))
            save_btn.grid(row=i, column=3, padx=5, pady=1)
        
        # Configure grid weights
        for i in range(4):
            table_frame.grid_columnconfigure(i, weight=1)
    
    def save_position(self, row_idx, position):
        """Save position data"""
        try:
            name = self.entries[f'name_{row_idx}'].get().strip()
            email = self.entries[f'email_{row_idx}'].get().strip()
            
            # Validate email format if provided
            if email and not self.is_valid_email(email):
                messagebox.showerror("Invalid Email", "Please enter a valid email address!")
                return
            
            # Update data
            data = self.get_data()
            position_found = False
            
            for row in data:
                if row.get('posisi') == position:
                    row['nama'] = name
                    row['email'] = email
                    row['last_updated'] = datetime.now().strftime("%d/%m/%Y %H:%M")
                    position_found = True
                    break
            
            # If position not found in existing data, add it
            if not position_found:
                data.append({
                    'posisi': position,
                    'nama': name,
                    'email': email,
                    'last_updated': datetime.now().strftime("%d/%m/%Y %H:%M")
                })
            
            # Save to sheet if available
            if self.sheet:
                try:
                    # Prepare data for sheet update
                    headers = ['posisi', 'nama', 'email', 'last_updated']
                    sheet_data = [headers]
                    
                    # Add all data rows
                    for row in data:
                        sheet_data.append([row.get(h, '') for h in headers])
                    
                    # Clear and update sheet
                    self.sheet.clear()
                    self.sheet.update(sheet_data)
                    
                    print(f"✅ Saved position: {position} with name: {name}, email: {email}")
                    messagebox.showinfo("Success", f"Position {position} saved successfully!")
                    
                    # Refresh the display
                    self.show_positions()
                    
                except Exception as e:
                    print(f"❌ Error saving to sheet: {e}")
                    messagebox.showerror("Error", f"Failed to save to Google Sheets: {e}")
            else:
                print(f"✅ Saved position locally: {position} with name: {name}, email: {email}")
                messagebox.showinfo("Success", f"Position {position} saved locally!")
                
        except Exception as e:
            print(f"❌ Error saving position: {e}")
            messagebox.showerror("Error", f"Failed to save position: {e}")
    
    def is_valid_email(self, email):
        """Simple email validation"""
        import re
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def logout(self):
        """Logout user"""
        print("🚪 Logging out...")
        self.logged_in = False
        self.sheet = None
        self.show_login()
    
    def run(self):
        """Run the application"""
        print("🚀 Starting application...")
        
        # Force window to front
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.attributes('-topmost', False)
        
        # Start main loop
        self.root.mainloop()
        print("👋 Application closed")

# Main execution
if __name__ == "__main__":
    try:
        app = SimpleAdminApp()
        app.run()
    except Exception as e:
        print(f"💥 Fatal error: {e}")
        import traceback
        traceback.print_exc()
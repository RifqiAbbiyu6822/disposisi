import tkinter as tk
from tkinter import ttk

class LoadingScreen(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self._destroyed = False
        self.title("Processing...")
        self.geometry("450x300")
        self.resizable(False, False)
        self.transient(parent)
        
        # Modern window styling
        self.configure(bg="#FFFFFF")
        self.overrideredirect(True)  # Remove window decorations
        
        try:
            self.iconbitmap('JapekELEVATED.ico')
        except:
            pass

        self.grab_set()
        self.update_idletasks()

        # Center the window with smooth animation
        self.center_window()
        
        # Create modern UI
        self.create_modern_ui()
        
        # Start animations
        self.animate_progress()

    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

    def create_modern_ui(self):
        # Outer frame with border - FIXED: No alpha colors
        outer_frame = tk.Frame(self, bg="#E2E8F0", relief="flat", bd=1)
        outer_frame.pack(fill="both", expand=True, padx=1, pady=1)
        
        # Main container with white background
        main_container = tk.Frame(outer_frame, bg="#FFFFFF")
        main_container.pack(fill="both", expand=True)
        
        # Content area
        content_frame = tk.Frame(main_container, bg="#FFFFFF")
        content_frame.pack(fill="both", expand=True, padx=40, pady=40)
        
        # Modern spinner using Canvas
        self.spinner_canvas = tk.Canvas(content_frame, width=80, height=80, 
                                       bg="#FFFFFF", highlightthickness=0)
        self.spinner_canvas.pack(pady=(0, 30))
        
        # Draw modern circular spinner
        self.spinner_angle = 0
        self.draw_modern_spinner()
        
        # Main status text
        self.status_label = tk.Label(content_frame, 
                                    text="Processing your request...", 
                                    font=("Inter", 14) if "Inter" in tk.font.families() else ("Segoe UI", 14), 
                                    bg="#FFFFFF",
                                    fg="#0F172A")
        self.status_label.pack(pady=(0, 10))
        
        # Progress percentage
        self.percentage_label = tk.Label(content_frame, 
                                       text="0%", 
                                       font=("Inter", 24, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 24, "bold"), 
                                       bg="#FFFFFF",
                                       fg="#3B82F6")
        self.percentage_label.pack(pady=(0, 20))
        
        # Modern progress bar
        progress_container = tk.Frame(content_frame, bg="#F1F5F9", height=6)
        progress_container.pack(fill="x", pady=(0, 15))
        
        self.progress_bar = tk.Frame(progress_container, bg="#3B82F6", height=6)
        self.progress_bar.place(x=0, y=0, relheight=1, width=0)
        
        # Sub-status text
        self.sub_status_label = tk.Label(content_frame, 
                                       text="Please wait...", 
                                       font=("Inter", 10) if "Inter" in tk.font.families() else ("Segoe UI", 10), 
                                       bg="#FFFFFF",
                                       fg="#94A3B8")
        self.sub_status_label.pack()

    def draw_modern_spinner(self):
        """Draw a modern circular spinner - FIXED: No alpha colors"""
        if not self.spinner_canvas.winfo_exists():
            return
            
        self.spinner_canvas.delete("all")
        
        center_x, center_y = 40, 40
        radius = 30
        
        # Background circle - FIXED: Solid color instead of alpha
        for i in range(8):
            start_angle = i * 45
            self.spinner_canvas.create_arc(
                center_x - radius, center_y - radius,
                center_x + radius, center_y + radius,
                start=start_angle, extent=30,
                outline="#E2E8F0", width=4, style="arc"
            )
        
        # Animated arc
        start = self.spinner_angle
        self.spinner_canvas.create_arc(
            center_x - radius, center_y - radius,
            center_x + radius, center_y + radius,
            start=start, extent=60,
            outline="#3B82F6", width=4, style="arc"
        )
        
        # Secondary arc for effect - FIXED: Solid color
        self.spinner_canvas.create_arc(
            center_x - radius, center_y - radius,
            center_x + radius, center_y + radius,
            start=start + 180, extent=60,
            outline="#60A5FA", width=4, style="arc"
        )

    def animate_progress(self):
        """Animate the spinner"""
        if getattr(self, '_destroyed', False):
            return
            
        self.spinner_angle = (self.spinner_angle + 10) % 360
        self.draw_modern_spinner()
        self.after(50, self.animate_progress)

    def update_progress(self, value, status_text=None):
        """Update progress with smooth animation"""
        if getattr(self, '_destroyed', False):
            return
            
        # Update percentage
        self.percentage_label.config(text=f"{value}%")
        
        # Animate progress bar
        target_width = int((value / 100) * self.progress_bar.master.winfo_width())
        self.progress_bar.place(width=target_width)
        
        # Update status text
        if status_text:
            self.sub_status_label.config(text=status_text)
        
        # Change status based on progress
        if value < 30:
            self.status_label.config(text="Initializing...")
        elif value < 60:
            self.status_label.config(text="Processing data...")
        elif value < 90:
            self.status_label.config(text="Finalizing...")
        else:
            self.status_label.config(text="Almost done...")
        
        # Complete state
        if value >= 100:
            self.spinner_canvas.delete("all")
            # Draw checkmark
            center_x, center_y = 40, 40
            self.spinner_canvas.create_oval(
                center_x - 30, center_y - 30,
                center_x + 30, center_y + 30,
                outline="#10B981", width=3
            )
            self.spinner_canvas.create_line(
                center_x - 15, center_y,
                center_x - 5, center_y + 10,
                fill="#10B981", width=3
            )
            self.spinner_canvas.create_line(
                center_x - 5, center_y + 10,
                center_x + 15, center_y - 10,
                fill="#10B981", width=3
            )
            
            self.status_label.config(text="Complete!", fg="#10B981")
            self.percentage_label.config(fg="#10B981")
            self.progress_bar.config(bg="#10B981")
            self.sub_status_label.config(text="Operation completed successfully")
            
            # Auto close after delay
            self.after(1000, self.destroy)
        
        self.update()

    def set_status(self, message):
        """Set custom status message"""
        if not getattr(self, '_destroyed', False):
            self.sub_status_label.config(text=message)
            self.update()

    def destroy(self):
        self._destroyed = True
        super().destroy()
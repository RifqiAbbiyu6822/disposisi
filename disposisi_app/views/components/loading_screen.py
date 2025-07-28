import tkinter as tk
from tkinter import ttk
from .styles import setup_styles  # Tambahkan impor style global

class LoadingScreen(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self._destroyed = False
        self.title("Loading...")
        self.geometry("400x250")
        self.resizable(False, False)
        self.transient(parent)

        # Terapkan style global
        setup_styles(self)
        # Gunakan warna dari root
        self.configure(bg=self.white_color)

        try:
            self.iconbitmap('JapekELEVATED.ico')
        except:
            pass

        self.grab_set()
        self.update_idletasks()

        # Center the window
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

        # Add subtle shadow effect with frame
        self.create_shadow_frame()
        self.create_widgets()

        # Animation variables
        self.dot_count = 0
        self.base_text = "Processing"
        self.animate_text()

    def destroy(self):
        self._destroyed = True
        super().destroy()

    def create_shadow_frame(self):
        """Create a subtle shadow effect"""
        shadow_frame = tk.Frame(self, bg="#e9ecef", relief="flat")
        shadow_frame.place(x=3, y=3, relwidth=1, relheight=1)

        # Main content frame dengan warna putih global
        self.main_container = tk.Frame(self, bg=self.white_color, relief="solid", bd=1)
        self.main_container.place(x=0, y=0, relwidth=1, relheight=1)

    def create_widgets(self):
        # Main content area with padding
        main_frame = ttk.Frame(self.main_container, padding=40)
        main_frame.pack(fill="both", expand=True)

        # Animated loading circle using Canvas
        self.canvas = tk.Canvas(main_frame, width=80, height=80, bg=self.white_color, highlightthickness=0)
        self.canvas.pack(pady=(0, 25))

        # Draw loading circle
        self.circle_angle = 0
        self.draw_loading_circle()
        self.animate_circle()

        # Main loading text with animation
        self.loading_label = ttk.Label(main_frame, text="Processing...", 
                                     font=("Segoe UI", 14, "normal"), 
                                     foreground=self.blue_color)
        self.loading_label.pack(pady=(0, 15))

        # Percentage display with modern typography
        self.percentage = ttk.Label(main_frame, text="0%", 
                                  font=("Segoe UI", 16, "bold"), 
                                  foreground=self.blue_color)
        self.percentage.pack(pady=(0, 15))

        # Subtle status message
        self.status_label = ttk.Label(main_frame, text="Please wait...", 
                                    font=("Segoe UI", 10, "italic"), 
                                    foreground="#adb5bd")
        self.status_label.pack()

    def draw_loading_circle(self):
        """Draw animated loading circle"""
        if not self.canvas.winfo_exists():
            return
        self.canvas.delete("all")

        # Circle parameters
        center_x, center_y = 40, 40
        radius = 25

        # Background circle (light gray)
        self.canvas.create_oval(center_x - radius, center_y - radius,
                              center_x + radius, center_y + radius,
                              outline="#e9ecef", width=4)

        # Animated arc (pakai biru global)
        start_angle = self.circle_angle
        extent_angle = 90  # Length of the arc

        self.canvas.create_arc(center_x - radius, center_y - radius,
                             center_x + radius, center_y + radius,
                             start=start_angle, extent=extent_angle,
                             outline=self.blue_color, width=4, style="arc")

    def animate_circle(self):
        """Animate the loading circle"""
        if getattr(self, '_destroyed', False):
            return
        self.circle_angle = (self.circle_angle + 8) % 360
        self.draw_loading_circle()

        # Continue animation
        self.after(50, self.animate_circle)

    def animate_text(self):
        """Animate the loading text with dots"""
        if getattr(self, '_destroyed', False):
            return
        dots = "." * (self.dot_count % 4)
        spaces = " " * (3 - len(dots))
        # Use fixed width to prevent text jumping
        animated_text = f"{self.base_text}{dots:<3}"
        self.loading_label.config(text=animated_text)

        self.dot_count += 1

        # Continue animation
        self.after(600, self.animate_text)

    def update_progress(self, value, status_text=None):
        """Update progress with optional status message"""
        self.percentage.config(text=f"{value}%")

        # Update status message if provided
        if status_text:
            self.status_label.config(text=status_text)

        # Change circle color based on progress
        if value >= 100:
            self.canvas.delete("all")
            center_x, center_y = 40, 40
            radius = 25
            # Draw complete circle (green)
            self.canvas.create_oval(center_x - radius, center_y - radius,
                                  center_x + radius, center_y + radius,
                                  outline="#28a745", width=4)
            # Add checkmark
            self.canvas.create_text(center_x, center_y, text="\u2713", 
                                  font=("Segoe UI", 20, "bold"), fill="#28a745")
            self.loading_label.config(text="Complete!", foreground="#28a745")
            self.status_label.config(text="Operation completed successfully")
            self.percentage.config(foreground="#28a745")

        self.update()

    def set_status(self, message):
        """Set custom status message"""
        self.status_label.config(text=message)
        self.update()

    def set_loading_text(self, text):
        """Change the main loading text"""
        self.base_text = text
        # Reset animation with proper formatting
        self.loading_label.config(text=f"{text}...")
        self.update()
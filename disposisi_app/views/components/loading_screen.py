import tkinter as tk
from tkinter import ttk
import math

class LoadingScreen(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self._destroyed = False
        self.title("Processing...")
        self.geometry("320x180")
        self.resizable(False, False)
        self.transient(parent)
        
        # Modern minimal styling
        self.configure(bg="#FAFAFA")
        self.overrideredirect(True)
        
        try:
            self.iconbitmap('JapekELEVATED.ico')
        except:
            pass

        self.grab_set()
        self.update_idletasks()
        self.center_window()
        self.create_ui()
        self.animate()

    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

    def create_ui(self):
        # Main container with subtle shadow effect
        main_frame = tk.Frame(self, bg="#FFFFFF", relief="flat", bd=0)
        main_frame.pack(fill="both", expand=True, padx=8, pady=8)
        
        # Add subtle border
        border_frame = tk.Frame(main_frame, bg="#F1F5F9", height=1)
        border_frame.pack(fill="x", side="top")
        
        # Content container
        content = tk.Frame(main_frame, bg="#FFFFFF")
        content.pack(fill="both", expand=True, padx=32, pady=32)
        
        # Minimal spinner
        self.spinner_canvas = tk.Canvas(content, width=32, height=32, 
                                       bg="#FFFFFF", highlightthickness=0)
        self.spinner_canvas.pack(pady=(0, 20))
        
        self.angle = 0
        self.draw_spinner()
        
        # Simple status text
        self.status_label = tk.Label(content,
                                   text="Processing...",
                                   font=("Segoe UI", 12),
                                   bg="#FFFFFF",
                                   fg="#374151")
        self.status_label.pack(pady=(0, 16))
        
        # Minimal progress bar
        progress_bg = tk.Frame(content, bg="#F3F4F6", height=2)
        progress_bg.pack(fill="x")
        
        self.progress_fill = tk.Frame(progress_bg, bg="#3B82F6", height=2)
        self.progress_fill.place(x=0, y=0, relheight=1, width=0)
        
        # Optional percentage (hidden by default for minimal look)
        self.percent_label = tk.Label(content,
                                    text="",
                                    font=("Segoe UI", 9),
                                    bg="#FFFFFF",
                                    fg="#9CA3AF")
        self.percent_label.pack(pady=(8, 0))

    def draw_spinner(self):
        """Draw minimal rotating dots"""
        if not self.spinner_canvas.winfo_exists():
            return
            
        self.spinner_canvas.delete("all")
        
        center_x, center_y = 16, 16
        radius = 12
        
        # Draw 8 dots around circle
        for i in range(8):
            angle = math.radians(i * 45 + self.angle)
            x = center_x + radius * math.cos(angle)
            y = center_y + radius * math.sin(angle)
            
            # Fade effect - dots closer to "front" are more opaque
            opacity = max(0.2, 1 - (i / 8))
            color = self.blend_colors("#3B82F6", "#FFFFFF", 1 - opacity)
            
            self.spinner_canvas.create_oval(x-2, y-2, x+2, y+2, 
                                          fill=color, outline="")

    def blend_colors(self, color1, color2, ratio):
        """Blend two hex colors"""
        def hex_to_rgb(hex_color):
            return tuple(int(hex_color[i:i+2], 16) for i in (1, 3, 5))
        
        def rgb_to_hex(rgb):
            return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        
        rgb1 = hex_to_rgb(color1)
        rgb2 = hex_to_rgb(color2)
        
        blended = tuple(int(rgb1[i] * ratio + rgb2[i] * (1 - ratio)) for i in range(3))
        return rgb_to_hex(blended)

    def animate(self):
        """Animate spinner rotation"""
        if getattr(self, '_destroyed', False):
            return
            
        self.angle = (self.angle + 15) % 360
        self.draw_spinner()
        self.after(80, self.animate)

    def update_progress(self, value, status_text=None):
        """Update progress"""
        if getattr(self, '_destroyed', False):
            return
            
        # Update progress bar
        container_width = self.progress_fill.master.winfo_width()
        if container_width > 1:
            target_width = int((value / 100) * container_width)
            self.progress_fill.place(width=target_width)
        
        # Update status
        if status_text:
            self.status_label.config(text=status_text)
        
        # Show percentage only if needed
        if value > 0:
            self.percent_label.config(text=f"{value}%")
        
        self.update_idletasks()

    def destroy(self):
        """Clean destroy"""
        self._destroyed = True
        try:
            super().destroy()
        except:
            pass

# Example usage
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide main window
    
    loading = LoadingScreen(root)
    
    # Simulate progress
    def simulate_progress():
        for i in range(0, 101, 10):
            if loading._destroyed:
                break
            loading.update_progress(i, f"Step {i//10 + 1}/10")
            root.update()
            root.after(200)
        
        if not loading._destroyed:
            loading.destroy()
        root.quit()
    
    root.after(100, simulate_progress)
    root.mainloop()
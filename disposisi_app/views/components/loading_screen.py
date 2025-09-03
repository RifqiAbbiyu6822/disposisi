import tkinter as tk
from tkinter import ttk, messagebox
import math
import threading
import time
from functools import wraps

class LoadingScreen(tk.Toplevel):
    def __init__(self, parent, title="Processing...", show_progress=True):
        super().__init__(parent)
        self._destroyed = False
        self.title(title)
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
        self.create_ui(show_progress)
        self.animate()

    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

    def create_ui(self, show_progress=True):
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
        
        if show_progress:
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
        else:
            self.progress_fill = None
            self.percent_label = None

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
            
        try:
            # Update progress bar
            if self.progress_fill and self.progress_fill.winfo_exists():
                container_width = self.progress_fill.master.winfo_width()
                if container_width > 1:
                    target_width = int((value / 100) * container_width)
                    self.progress_fill.place(width=target_width)
                
                # Show percentage only if needed
                if value > 0 and self.percent_label and self.percent_label.winfo_exists():
                    self.percent_label.config(text=f"{value}%")
            
            # Update status
            if status_text and self.status_label and self.status_label.winfo_exists():
                self.status_label.config(text=status_text)
            
            self.update_idletasks()
        except tk.TclError:
            # Widget sudah dihancurkan
            self._destroyed = True

    def destroy(self):
        """Clean destroy"""
        self._destroyed = True
        try:
            super().destroy()
        except:
            pass

class LoadingManager:
    """Global loading screen manager"""
    _instance = None
    _current_loading = None
    _loading_count = 0
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(LoadingManager, cls).__new__(cls)
        return cls._instance
    
    def show_loading(self, parent, title="Processing...", show_progress=True):
        """Show loading screen"""
        if self._current_loading and self._current_loading.winfo_exists():
            self._current_loading.destroy()
        
        self._current_loading = LoadingScreen(parent, title, show_progress)
        self._loading_count += 1
        return self._current_loading
    
    def hide_loading(self):
        """Hide current loading screen"""
        if self._current_loading and self._current_loading.winfo_exists():
            self._current_loading.destroy()
        self._current_loading = None
    
    def update_progress(self, value, status_text=None):
        """Update progress of current loading screen"""
        if self._current_loading and self._current_loading.winfo_exists():
            try:
                self._current_loading.update_progress(value, status_text)
            except tk.TclError:
                # Widget sudah dihancurkan, reset current loading
                self._current_loading = None

def with_loading_screen(title="Processing...", show_progress=True):
    """Decorator to automatically show loading screen during threading operations"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            # Find parent widget from args
            parent = None
            for arg in args:
                if hasattr(arg, 'winfo_toplevel'):
                    parent = arg
                    break
            
            if not parent:
                # Try to find from kwargs
                for key, value in kwargs.items():
                    if hasattr(value, 'winfo_toplevel'):
                        parent = value
                        break
            
            if not parent:
                return func(*args, **kwargs)
            
            loading_manager = LoadingManager()
            loading = loading_manager.show_loading(parent, title, show_progress)
            
            try:
                result = func(*args, **kwargs)
                return result
            finally:
                loading_manager.hide_loading()
        
        return wrapper
    return decorator

def threaded_with_loading(title="Processing...", show_progress=True):
    """Decorator to run function in thread with loading screen"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            # Find parent widget from args
            parent = None
            for arg in args:
                if hasattr(arg, 'winfo_toplevel'):
                    parent = arg
                    break
            
            if not parent:
                # Try to find from kwargs
                for key, value in kwargs.items():
                    if hasattr(value, 'winfo_toplevel'):
                        parent = value
                        break
            
            if not parent:
                # Run without loading screen
                threading.Thread(target=func, args=args, kwargs=kwargs, daemon=True).start()
                return
            
            loading_manager = LoadingManager()
            loading = loading_manager.show_loading(parent, title, show_progress)
            
            def threaded_func():
                try:
                    func(*args, **kwargs)
                finally:
                    loading.after(0, loading_manager.hide_loading)
            
            threading.Thread(target=threaded_func, daemon=True).start()
        
        return wrapper
    return decorator

# Enhanced messagebox with loading icon
class LoadingMessageBox:
    @staticmethod
    def showinfo(title, message, parent=None, **kwargs):
        if parent:
            try:
                parent.iconbitmap('JapekELEVATED.ico')
            except:
                pass
        return messagebox.showinfo(title, message, parent=parent, **kwargs)
    
    @staticmethod
    def showwarning(title, message, parent=None, **kwargs):
        if parent:
            try:
                parent.iconbitmap('JapekELEVATED.ico')
            except:
                pass
        return messagebox.showwarning(title, message, parent=parent, **kwargs)
    
    @staticmethod
    def showerror(title, message, parent=None, **kwargs):
        if parent:
            try:
                parent.iconbitmap('JapekELEVATED.ico')
            except:
                pass
        return messagebox.showerror(title, message, parent=parent, **kwargs)
    
    @staticmethod
    def askyesno(title, message, parent=None, **kwargs):
        if parent:
            try:
                parent.iconbitmap('JapekELEVATED.ico')
            except:
                pass
        return messagebox.askyesno(title, message, parent=parent, **kwargs)
    
    @staticmethod
    def askokcancel(title, message, parent=None, **kwargs):
        if parent:
            try:
                parent.iconbitmap('JapekELEVATED.ico')
            except:
                pass
        return messagebox.askokcancel(title, message, parent=parent, **kwargs)

# Global loading manager instance
loading_manager = LoadingManager()

# Utility function to replace messagebox calls
def replace_messagebox_calls():
    """Replace all messagebox calls with LoadingMessageBox calls"""
    import re
    
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
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
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
            
            # Replace messagebox calls
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
            
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
                
        except Exception as e:
            print(f"Error updating {file_path}: {e}")

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
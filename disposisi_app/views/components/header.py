from tkinter import ttk
import tkinter as tk

def create_header(parent):
    """Create a minimal header"""
    # Main header container - lebih compact
    header_frame = ttk.Frame(parent, style="Header.TFrame")
    header_frame.pack(fill="x", padx=0, pady=0)
    
    # Main content container dengan padding minimal
    header_content = ttk.Frame(header_frame, style="Header.TFrame")
    header_content.pack(fill="x", padx=15, pady=8)  # Reduced padding from 24,16 to 15,8
    
    # Single line layout untuk title dan info
    title_container = ttk.Frame(header_content, style="Header.TFrame")
    title_container.pack(fill="x")
    
    # Left side - Compact title
    left_section = ttk.Frame(title_container, style="Header.TFrame")
    left_section.pack(side="left", fill="y")
    
    # Compact title - smaller font and single line
    title_label = ttk.Label(
        left_section, 
        text="📋 Sistem Disposisi", 
        font=("Inter", 14, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 14, "bold"), 
        style="HeaderTitle.TLabel"
    )
    title_label.pack(anchor="w")
    
    # Right side - Compact status
    right_section = ttk.Frame(title_container, style="Header.TFrame")
    right_section.pack(side="right", fill="y")
    
    # Simple status indicator
    status_frame = ttk.Frame(right_section, style="Header.TFrame")
    status_frame.pack(anchor="e")
    
    # Compact version info
    version_label = ttk.Label(
        status_frame, 
        text="v2.0", 
        font=("Inter", 10, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 10, "bold"), 
        style="HeaderVersion.TLabel"
    )
    version_label.pack(side="right", padx=(10, 0))
    
    # Status dot
    status_dot = tk.Label(
        status_frame, 
        text="● Online", 
        font=("Segoe UI", 9), 
        fg="#10B981", 
        bg=getattr(parent.master, 'primary_color', '#1A1B23') if hasattr(parent.master, 'primary_color') else "#1A1B23"
    )
    status_dot.pack(side="right", padx=(0, 8))
    
    return header_frame

def add_header_styles(style, root):
    """Add minimal header styles"""
    
    # Get colors from root or use defaults
    primary_color = getattr(root, 'primary_color', '#1A1B23')
    card_color = getattr(root, 'card_color', '#FFFFFF')
    
    # Minimal header styles
    style.configure("HeaderTitle.TLabel", 
                   background=primary_color, 
                   foreground=card_color, 
                   font=("Inter", 14, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 14, "bold"))
    
    style.configure("HeaderVersion.TLabel", 
                   background=primary_color, 
                   foreground=card_color, 
                   font=("Inter", 10, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 10, "bold"))
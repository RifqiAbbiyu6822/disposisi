from tkinter import ttk
import tkinter as tk

def create_header(parent):
    # Main header container with modern glass effect
    header_frame = ttk.Frame(parent, style="HeaderGlass.TFrame", height=90)
    header_frame.pack(fill="x", padx=0, pady=0)
    header_frame.pack_propagate(False)
    
    # Gradient background simulation
    gradient_top = tk.Frame(header_frame, bg="#0F172A", height=45)
    gradient_top.pack(fill="x", side="top")
    
    gradient_bottom = tk.Frame(header_frame, bg="#1E293B", height=45)
    gradient_bottom.pack(fill="x", side="bottom")
    
    # Main content overlay
    content_overlay = tk.Frame(header_frame, bg="#0F172A")
    content_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
    
    # Content container with padding
    header_content = tk.Frame(content_overlay, bg="#0F172A")
    header_content.pack(fill="both", expand=True, padx=32, pady=20)
    
    # Left section - Modern branding
    left_section = tk.Frame(header_content, bg="#0F172A")
    left_section.pack(side="left", fill="both", expand=True)
    
    # Logo and title container
    brand_container = tk.Frame(left_section, bg="#0F172A")
    brand_container.pack(side="left", fill="y")
    
    # Logo with modern badge style
    logo_frame = tk.Frame(brand_container, bg="#3B82F6", width=50, height=50)
    logo_frame.pack(side="left", padx=(0, 20))
    logo_frame.pack_propagate(False)
    
    logo_label = tk.Label(logo_frame, text="📋", font=("Segoe UI", 24), 
                         bg="#3B82F6", fg="white")
    logo_label.place(relx=0.5, rely=0.5, anchor="center")
    
    # Title section
    title_container = tk.Frame(brand_container, bg="#0F172A")
    title_container.pack(side="left", fill="both", expand=True)
    
    # Main title
    title_label = tk.Label(
        title_container, 
        text="Sistem Disposisi Digital", 
        font=("Inter", 22, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 22, "bold"), 
        bg="#0F172A",
        fg="white"
    )
    title_label.pack(anchor="w")
    
    # Subtitle with modern styling
    subtitle_label = tk.Label(
        title_container, 
        text="Enterprise Document Management System", 
        font=("Inter", 12) if "Inter" in tk.font.families() else ("Segoe UI", 12), 
        bg="#0F172A",
        fg="#94A3B8"
    )
    subtitle_label.pack(anchor="w", pady=(4, 0))
    
    # Right section - Modern status indicators
    right_section = tk.Frame(header_content, bg="#0F172A")
    right_section.pack(side="right", fill="y")
    
    # Status container with glass effect
    status_container = tk.Frame(right_section, bg="#1E293B", relief="flat")
    status_container.pack(side="right", fill="y", padx=20, pady=5)
    
    # Inner status frame
    status_inner = tk.Frame(status_container, bg="#1E293B")
    status_inner.pack(padx=15, pady=8)
    
    # Connection status with modern indicator
    status_row = tk.Frame(status_inner, bg="#1E293B")
    status_row.pack(anchor="e", pady=(0, 4))
    
    # Animated status dot
    status_dot_outer = tk.Label(status_row, text="●", font=("Arial", 14), 
                               fg="#10B98130", bg="#1E293B")
    status_dot_outer.pack(side="left")
    
    status_dot = tk.Label(status_row, text="●", font=("Arial", 10), 
                         fg="#10B981", bg="#1E293B")
    status_dot.place(in_=status_dot_outer, relx=0.5, rely=0.5, anchor="center")
    
    status_text = tk.Label(
        status_row, 
        text="System Online", 
        font=("Inter", 10, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 10, "bold"), 
        fg="#10B981",
        bg="#1E293B"
    )
    status_text.pack(side="left", padx=(8, 0))
    
    # Version info with modern badge
    version_frame = tk.Frame(status_inner, bg="#1E293B")
    version_frame.pack(anchor="e")
    
    version_badge = tk.Frame(version_frame, bg="#3B82F6", relief="flat")
    version_badge.pack(side="right")
    
    version_label = tk.Label(
        version_badge, 
        text=" v2.0.1 ", 
        font=("Inter", 9, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 9, "bold"), 
        bg="#3B82F6",
        fg="white"
    )
    version_label.pack(padx=8, pady=2)
    
    # Add subtle animations
    def pulse_status():
        current_color = status_dot.cget("fg")
        new_color = "#10B981" if current_color == "#10B98180" else "#10B98180"
        status_dot.config(fg=new_color)
        header_frame.after(1000, pulse_status)
    
    pulse_status()
    
    return header_frame
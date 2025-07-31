from tkinter import ttk
import tkinter as tk

def create_header(parent):
    # Main header container with gradient effect
    header_frame = ttk.Frame(parent, style="Header.TFrame")
    header_frame.pack(fill="x", padx=0, pady=0)
    
    # Add subtle shadow/border at bottom
    shadow_frame = ttk.Frame(header_frame, height=1, style="HeaderShadow.TFrame")
    shadow_frame.pack(fill="x", side="bottom")
    
    # Main content container with proper padding
    header_content = ttk.Frame(header_frame, style="Header.TFrame")
    header_content.pack(fill="both", expand=True, padx=24, pady=16)
    
    # Left side - Logo and branding
    left_section = ttk.Frame(header_content, style="Header.TFrame")
    left_section.pack(side="left", fill="y")
    
    # Logo container with background accent
    logo_container = ttk.Frame(left_section, style="HeaderAccent.TFrame")
    logo_container.pack(side="left", padx=(0, 20))
    
    # Logo/Icon - using Unicode symbol for professional look
    logo_label = ttk.Label(logo_container, text="📋", font=("Segoe UI", 24), style="HeaderIcon.TLabel")
    logo_label.pack(padx=12, pady=8)
    
    # Title and subtitle container
    title_container = ttk.Frame(left_section, style="Header.TFrame")
    title_container.pack(side="left", fill="both", expand=True)
    
    # Main title with better typography
    title_label = ttk.Label(
        title_container, 
        text="Sistem Disposisi", 
        font=("Inter", 20, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 20, "bold"), 
        style="HeaderTitle.TLabel"
    )
    title_label.pack(anchor="w")
    
    # Subtitle with better spacing
    subtitle_label = ttk.Label(
        title_container, 
        text="Pembuatan dan Pelaporan Surat Disposisi", 
        font=("Inter", 11) if "Inter" in tk.font.families() else ("Segoe UI", 11), 
        style="HeaderSubtitle.TLabel"
    )
    subtitle_label.pack(anchor="w", pady=(4, 0))
    
    # Right side - Info and actions
    right_section = ttk.Frame(header_content, style="Header.TFrame")
    right_section.pack(side="right", fill="y")
    
    # Status indicator
    status_container = ttk.Frame(right_section, style="Header.TFrame")
    status_container.pack(side="right", padx=(20, 0))
    
    # Online status indicator
    status_frame = ttk.Frame(status_container, style="Header.TFrame")
    status_frame.pack(anchor="e", pady=(0, 8))
    
    # Status dot (green circle)
    status_dot = tk.Label(
        status_frame, 
        text="●", 
        font=("Segoe UI", 12), 
        fg="#10B981", 
        bg=parent.master.primary_color if hasattr(parent.master, 'primary_color') else "#1A1B23"
    )
    status_dot.pack(side="left")
    
    status_text = ttk.Label(
        status_frame, 
        text="Online", 
        font=("Inter", 9) if "Inter" in tk.font.families() else ("Segoe UI", 9), 
        style="HeaderStatus.TLabel"
    )
    status_text.pack(side="left", padx=(4, 0))
    
    # Version and build info
    info_container = ttk.Frame(status_container, style="Header.TFrame")
    info_container.pack(anchor="e")
    
    version_label = ttk.Label(
        info_container, 
        text="v2.0.1", 
        font=("Inter", 9, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 9, "bold"), 
        style="HeaderVersion.TLabel"
    )
    version_label.pack(anchor="e")
    
    build_label = ttk.Label(
        info_container, 
        text="Build 2024.07", 
        font=("Inter", 8) if "Inter" in tk.font.families() else ("Segoe UI", 8), 
        style="HeaderBuild.TLabel"
    )
    build_label.pack(anchor="e", pady=(2, 0))
    
    # Add hover effects and interactions
    def on_logo_hover(event):
        logo_label.configure(style="HeaderIconHover.TLabel")
    
    def on_logo_leave(event):
        logo_label.configure(style="HeaderIcon.TLabel")
    
    # Bind hover events
    logo_label.bind("<Enter>", on_logo_hover)
    logo_label.bind("<Leave>", on_logo_leave)
    
    return header_frame

# Additional styles to be added to your setup_styles function
def add_header_styles(style, root):
    """Add enhanced header styles to the existing style configuration"""
    
    # Get colors from root or use defaults
    primary_color = getattr(root, 'primary_color', '#1A1B23')
    accent_color = getattr(root, 'accent_color', '#4F46E5')
    card_color = getattr(root, 'card_color', '#FFFFFF')
    
    # Header shadow/border
    style.configure("HeaderShadow.TFrame", 
                   background="#E5E7EB", 
                   relief="flat")
    
    # Header accent container
    style.configure("HeaderAccent.TFrame", 
                   background=accent_color, 
                   relief="flat")
    
    # Enhanced header icon
    style.configure("HeaderIcon.TLabel", 
                   background=accent_color, 
                   foreground=card_color, 
                   font=("Segoe UI", 24))
    
    style.configure("HeaderIconHover.TLabel", 
                   background="#3730A3", 
                   foreground=card_color, 
                   font=("Segoe UI", 24))
    
    # Enhanced title styles
    style.configure("HeaderTitle.TLabel", 
                   background=primary_color, 
                   foreground=card_color, 
                   font=("Inter", 20, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 20, "bold"))
    
    style.configure("HeaderSubtitle.TLabel", 
                   background=primary_color, 
                   foreground="#D1D5DB", 
                   font=("Inter", 11) if "Inter" in tk.font.families() else ("Segoe UI", 11))
    
    # Status and info labels
    style.configure("HeaderStatus.TLabel", 
                   background=primary_color, 
                   foreground="#10B981", 
                   font=("Inter", 9) if "Inter" in tk.font.families() else ("Segoe UI", 9))
    
    style.configure("HeaderVersion.TLabel", 
                   background=primary_color, 
                   foreground=card_color, 
                   font=("Inter", 9, "bold") if "Inter" in tk.font.families() else ("Segoe UI", 9, "bold"))
    
    style.configure("HeaderBuild.TLabel", 
                   background=primary_color, 
                   foreground="#9CA3AF", 
                   font=("Inter", 8) if "Inter" in tk.font.families() else ("Segoe UI", 8))

# Usage example:
"""
# In your main application setup, after creating styles:
def setup_styles(root):
    # ... your existing setup_styles code ...
    
    # Add enhanced header styles
    add_header_styles(style, root)

# Then create header:
header = create_header(main_window)
"""
import tkinter as tk
from tkinter import ttk

def setup_styles(root):
    style = ttk.Style()
    style.theme_use('clam')
    
    # Ultra Modern Professional Color Palette - FIXED: No alpha values
    primary_color = '#0F172A'       # Slate 900 - Deep professional
    secondary_color = '#1E293B'     # Slate 800
    accent_color = '#3B82F6'        # Blue 500 - Modern accent
    accent_hover = '#2563EB'        # Blue 600
    background_color = '#F8FAFC'    # Slate 50 - Clean background
    surface_color = '#FFFFFF'       # Pure white surfaces
    card_hover = '#F1F5F9'         # Slate 100
    text_primary = '#0F172A'       # Slate 900
    text_secondary = '#475569'     # Slate 600
    text_muted = '#94A3B8'         # Slate 400
    border_color = '#E2E8F0'       # Slate 200
    border_focus = '#CBD5E1'       # Slate 300
    success_color = '#10B981'      # Emerald 500
    success_hover = '#059669'      # Emerald 600
    warning_color = '#F59E0B'      # Amber 500
    error_color = '#EF4444'        # Red 500
    info_color = '#06B6D4'         # Cyan 500

    # FIXED: Enhanced shadow and gradient colors without alpha
    shadow_light = '#F0F0F0'       # Light gray instead of alpha
    shadow_medium = '#E0E0E0'      # Medium gray instead of alpha
    shadow_dark = '#D0D0D0'        # Dark gray instead of alpha
    gradient_start = '#F8FAFC'
    gradient_end = '#F1F5F9'

    # Base configuration
    style.configure(".", 
                   background=background_color, 
                   foreground=text_primary, 
                   font=("Inter", 10))
    
    if isinstance(root, (tk.Tk, tk.Toplevel)):
        root.configure(bg=background_color)

    # Modern Header with Glass Effect
    style.configure("Header.TFrame", 
                   background=primary_color,
                   relief="flat")
    
    style.configure("HeaderGlass.TFrame",
                   background=primary_color,
                   relief="flat",
                   borderwidth=0)
    
    style.configure("Header.TLabel", 
                   background=primary_color, 
                   foreground=surface_color, 
                   font=("Inter", 18, "bold"))
    
    style.configure("HeaderSub.TLabel", 
                   background=primary_color, 
                   foreground="#CBD5E1", 
                   font=("Inter", 11))

    # Card Design with Hover Effects
    style.configure("Card.TFrame", 
                   background=surface_color, 
                   relief="flat", 
                   borderwidth=1,
                   bordercolor=border_color)
    
    style.map("Card.TFrame",
             background=[('active', card_hover)])

    # Enhanced LabelFrame with Modern Styling
    style.configure("TLabelframe", 
                   background=surface_color, 
                   borderwidth=1, 
                   bordercolor=border_color, 
                   relief="flat",
                   labelmargins=10)
    
    style.configure("TLabelframe.Label", 
                   background=surface_color, 
                   foreground=accent_color, 
                   font=("Inter", 12, "600"),
                   padding=(10, 5))

    # Typography System
    style.configure("TLabel", 
                   font=("Inter", 10), 
                   background=surface_color, 
                   foreground=text_primary)
    
    style.configure("Heading1.TLabel", 
                   font=("Inter", 24, "700"), 
                   background=background_color, 
                   foreground=primary_color)
    
    style.configure("Heading2.TLabel", 
                   font=("Inter", 18, "600"), 
                   background=background_color, 
                   foreground=primary_color)
    
    style.configure("Heading3.TLabel", 
                   font=("Inter", 14, "600"), 
                   background=surface_color, 
                   foreground=text_primary)
    
    style.configure("Caption.TLabel", 
                   font=("Inter", 9), 
                   background=surface_color, 
                   foreground=text_muted)

    # Modern Input Fields with Focus States
    style.configure("TEntry", 
                   font=("Inter", 11), 
                   fieldbackground=surface_color, 
                   bordercolor=border_color, 
                   lightcolor=border_color,
                   darkcolor=border_color,
                   insertcolor=accent_color,
                   padding=(12, 10), 
                   relief="flat", 
                   borderwidth=2)
    
    style.map("TEntry", 
             fieldbackground=[('focus', surface_color)],
             bordercolor=[('focus', accent_color)],
             lightcolor=[('focus', accent_color)],
             darkcolor=[('focus', accent_color)])
    
    style.configure("TCombobox", 
                   font=("Inter", 11), 
                   fieldbackground=surface_color, 
                   bordercolor=border_color,
                   lightcolor=border_color,
                   darkcolor=border_color,
                   arrowcolor=text_secondary,
                   padding=(12, 10),
                   relief="flat",
                   borderwidth=2)
    
    style.map("TCombobox",
             fieldbackground=[('focus', surface_color)],
             bordercolor=[('focus', accent_color)],
             lightcolor=[('focus', accent_color)],
             darkcolor=[('focus', accent_color)])

    # Advanced Button System
    style.configure("TButton",
                   font=("Inter", 11, "600"),
                   borderwidth=0,
                   relief="flat",
                   anchor="center",
                   padding=(20, 12))
    
    # Primary Button
    style.configure("Primary.TButton", 
                   font=("Inter", 11, "600"), 
                   background=accent_color, 
                   foreground="white", 
                   padding=(20, 12), 
                   borderwidth=0, 
                   relief="flat",
                   focuscolor='none')
    
    style.map("Primary.TButton", 
             background=[('pressed', accent_hover), 
                        ('active', accent_hover)],
             relief=[('pressed', 'flat'), ('active', 'flat')])
    
    # Secondary Button
    style.configure("Secondary.TButton", 
                   font=("Inter", 11), 
                   background=surface_color, 
                   foreground=text_primary, 
                   padding=(20, 12), 
                   borderwidth=2, 
                   bordercolor=border_color, 
                   relief="flat",
                   focuscolor='none')
    
    style.map("Secondary.TButton", 
             background=[('pressed', gradient_end), 
                        ('active', card_hover)],
             bordercolor=[('active', border_focus)])

    # Success Button
    style.configure("Success.TButton", 
                   font=("Inter", 11, "600"), 
                   background=success_color, 
                   foreground="white", 
                   padding=(20, 12), 
                   borderwidth=0, 
                   relief="flat",
                   focuscolor='none')
    
    style.map("Success.TButton", 
             background=[('pressed', success_hover), 
                        ('active', success_hover)])

    # Danger Button
    style.configure("Danger.TButton", 
                   font=("Inter", 11), 
                   background=error_color, 
                   foreground="white", 
                   padding=(16, 10), 
                   borderwidth=0, 
                   relief="flat",
                   focuscolor='none')
    
    style.map("Danger.TButton", 
             background=[('pressed', '#DC2626'), 
                        ('active', '#DC2626')])

    # Accent Button (New)
    style.configure("Accent.TButton", 
                   font=("Inter", 11, "600"), 
                   background=info_color, 
                   foreground="white", 
                   padding=(20, 12), 
                   borderwidth=0, 
                   relief="flat",
                   focuscolor='none')

    # Ghost Button (New)
    style.configure("Ghost.TButton", 
                   font=("Inter", 11), 
                   background=surface_color,  # FIXED: Changed from "transparent" 
                   foreground=accent_color, 
                   padding=(16, 10), 
                   borderwidth=1,
                   bordercolor=border_color,
                   relief="flat",
                   focuscolor='none')
    
    style.map("Ghost.TButton", 
             foreground=[('active', accent_hover)],
             background=[('active', card_hover)])  # FIXED: Added background mapping

    # Modern Notebook Tabs
    style.configure("TNotebook", 
                   background=background_color, 
                   borderwidth=0,
                   tabmargins=[0, 0, 0, 0])
    
    style.configure("TNotebook.Tab", 
                   font=("Inter", 11, "500"), 
                   padding=(24, 14), 
                   background=gradient_end, 
                   foreground=text_secondary, 
                   borderwidth=0, 
                   relief="flat")
    
    style.map("TNotebook.Tab",
             background=[('selected', surface_color)],
             foreground=[('selected', accent_color)],
             padding=[('selected', (24, 14, 24, 16))],
             relief=[('selected', 'flat')])

    # Enhanced Treeview with Zebra Striping
    style.configure("Treeview",
                  background=surface_color,
                  foreground=text_primary,
                  rowheight=36,
                  fieldbackground=surface_color,
                  borderwidth=1,
                  bordercolor=border_color,
                  font=("Inter", 10))
    
    style.configure("Treeview.Heading",
                  background=gradient_end,
                  foreground=text_primary,
                  relief="flat",
                  borderwidth=0,
                  font=("Inter", 11, "600"),
                  padding=(12, 8))
    
    style.map("Treeview",
             background=[('selected', accent_color)],
             foreground=[('selected', surface_color)])
    
    style.map("Treeview.Heading",
             background=[('active', card_hover)])

    # Modern Progress Bar
    style.configure("TProgressbar", 
                   troughcolor=gradient_end, 
                   background=accent_color, 
                   bordercolor=border_color,
                   lightcolor=accent_color,
                   darkcolor=accent_color,
                   borderwidth=0,
                   relief="flat")

    # Enhanced Checkbutton
    style.configure("TCheckbutton",
                   background=surface_color,
                   foreground=text_primary,
                   font=("Inter", 10),
                   focuscolor='none',
                   indicatorbackground=surface_color,
                   indicatorforeground=accent_color,
                   indicatorrelief="flat",
                   indicatordiameter=16)
    
    style.map("TCheckbutton",
             indicatorbackground=[('pressed', card_hover), 
                                ('active', gradient_end)])

    # Modern Scrollbar
    style.configure("TScrollbar",
                   background=gradient_end,
                   bordercolor=gradient_end,
                   arrowcolor=text_secondary,
                   troughcolor=background_color,
                   lightcolor=gradient_end,
                   darkcolor=gradient_end,
                   borderwidth=0,
                   relief="flat",
                   width=12)
    
    style.map("TScrollbar",
             background=[('active', border_color), 
                        ('pressed', text_secondary)])

    # Status Bar with Gradient
    style.configure("Status.TFrame", 
                   background=gradient_end, 
                   relief="flat",
                   borderwidth=0)
    
    style.configure("Status.TLabel", 
                   font=("Inter", 9), 
                   foreground=text_secondary, 
                   background=gradient_end)

    # Tooltip Style
    style.configure("Tooltip.TLabel",
                   background=primary_color,
                   foreground=surface_color,
                   font=("Inter", 9),
                   relief="flat",
                   borderwidth=1,
                   bordercolor=secondary_color)

    # Separator with subtle styling
    style.configure("TSeparator",
                   background=border_color)

    # Store extended color palette in root
    color_attrs = {
        'primary_color': primary_color,
        'secondary_color': secondary_color,
        'accent_color': accent_color,
        'accent_hover': accent_hover,
        'surface_color': surface_color,
        'background_color': background_color,
        'card_hover': card_hover,
        'text_primary': text_primary,
        'text_secondary': text_secondary,
        'text_muted': text_muted,
        'border_color': border_color,
        'border_focus': border_focus,
        'success_color': success_color,
        'warning_color': warning_color,
        'error_color': error_color,
        'info_color': info_color,
        'gradient_start': gradient_start,
        'gradient_end': gradient_end
    }
    
    for attr, value in color_attrs.items():
        setattr(root, attr, value)
    
    # Legacy attributes for compatibility
    setattr(root, 'canvas_bg', background_color)
    setattr(root, 'text_bg', surface_color)
    setattr(root, 'text_fg', text_primary)
    setattr(root, 'oddrow_color', gradient_end)
    setattr(root, 'evenrow_color', surface_color)
    setattr(root, 'blue_color', accent_color)
    setattr(root, 'yellow_gold_color', warning_color)
    setattr(root, 'white_color', surface_color)
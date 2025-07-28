import tkinter as tk
from tkinter import ttk

def setup_styles(root):
    style = ttk.Style()
    style.theme_use('clam')
    
    # Professional minimalist color palette
    primary_color = '#1A1B23'       # Deep charcoal (Primary)
    secondary_color = '#2D2E36'     # Lighter charcoal
    accent_color = '#4F46E5'        # Modern indigo (Accent)
    background_color = '#FAFAFA'    # Off-white (Background)
    card_color = '#FFFFFF'          # Pure white (Cards)
    text_color = '#1F2937'          # Dark gray text
    muted_color = '#6B7280'         # Medium gray (Muted text)
    border_color = '#E5E7EB'        # Light border
    success_color = '#10B981'       # Modern green
    warning_color = '#F59E0B'       # Amber
    error_color = '#EF4444'         # Modern red

    # Base style
    style.configure(".", 
                   background=background_color, 
                   foreground=text_color, 
                   font=("Segoe UI", 10))
    # Hanya set bg jika root adalah Tk atau Toplevel
    if isinstance(root, (tk.Tk, tk.Toplevel)):
        root.configure(bg=background_color)

    # Clean header design
    style.configure("Header.TFrame", background=primary_color)
    style.configure("Header.TLabel", 
                   background=primary_color, 
                   foreground=card_color, 
                   font=("Inter", 16, "bold"))
    style.configure("HeaderSub.TLabel", 
                   background=primary_color, 
                   foreground="#D1D5DB", 
                   font=("Inter", 10))

    # Minimal card design with subtle shadows
    style.configure("Card.TFrame", 
                   background=card_color, 
                   relief="flat", 
                   borderwidth=0)
    style.configure("TLabelframe", 
                   background=card_color, 
                   borderwidth=1, 
                   bordercolor=border_color, 
                   relief="solid")
    style.configure("TLabelframe.Label", 
                   background=card_color, 
                   foreground=text_color, 
                   font=("Inter", 11, "600"))

    # Clean typography
    style.configure("TLabel", 
                   font=("Inter", 10), 
                   background=card_color, 
                   foreground=text_color)
    style.configure("Heading.TLabel", 
                   font=("Inter", 12, "600"), 
                   background=card_color, 
                   foreground=text_color)

    # Modern input fields
    style.configure("TEntry", 
                   font=("Inter", 10), 
                   fieldbackground=card_color, 
                   bordercolor=border_color, 
                   padding=(12, 8), 
                   relief="solid", 
                   borderwidth=1)
    style.map("TEntry", 
             bordercolor=[('focus', accent_color)], 
             relief=[('focus', 'solid')], 
             borderwidth=[('focus', 2)])
    
    style.configure("TCombobox", 
                   font=("Inter", 10), 
                   fieldbackground=card_color, 
                   bordercolor=border_color, 
                   padding=(12, 8))

    # Modern button system
    style.configure("Primary.TButton", 
                   font=("Inter", 10, "bold"), 
                   background=accent_color, 
                   foreground="white", 
                   padding=(16, 10), 
                   borderwidth=0, 
                   relief="flat")
    style.map("Primary.TButton", 
             background=[('active', '#3730A3'), ('pressed', '#312E81')])
    
    style.configure("Secondary.TButton", 
                   font=("Inter", 10), 
                   background=card_color, 
                   foreground=text_color, 
                   padding=(16, 10), 
                   borderwidth=1, 
                   bordercolor=border_color, 
                   relief="solid")
    style.map("Secondary.TButton", 
             background=[('active', background_color), ('pressed', '#F3F4F6')])

    style.configure("Success.TButton", 
                   font=("Inter", 10, "bold"), 
                   background=success_color, 
                   foreground="white", 
                   padding=(16, 10), 
                   borderwidth=0, 
                   relief="flat")
    style.map("Success.TButton", 
             background=[('active', '#059669'), ('pressed', '#047857')])

    # Minimal tab design
    style.configure("TNotebook", 
                   background=background_color, 
                   borderwidth=0)
    style.configure("TNotebook.Tab", 
                   font=("Inter", 10), 
                   padding=(16, 12), 
                   background=background_color, 
                   foreground=muted_color, 
                   borderwidth=0, 
                   relief="flat")
    style.map("TNotebook.Tab",
             background=[('selected', card_color)],
             foreground=[('selected', text_color)])

    # Clean table design
    style.configure("Treeview",
                  background=card_color,
                  foreground=text_color,
                  rowheight=32,
                  fieldbackground=card_color,
                  borderwidth=0,
                  font=("Inter", 10))
    style.configure("Treeview.Heading",
                  background='#F9FAFB',
                  foreground=text_color,
                  relief="flat",
                  borderwidth=1,
                  bordercolor=border_color,
                  font=("Inter", 10, "600"))
    style.map("Treeview",
             background=[('selected', accent_color)],
             foreground=[('selected', card_color)])

    # Minimal progress bar
    style.configure("TProgressbar", 
                   troughcolor='#F3F4F6', 
                   background=accent_color, 
                   bordercolor=border_color,
                   borderwidth=0)

    # Clean status bar
    style.configure("Status.TFrame", 
                   background='#F9FAFB', 
                   relief="flat")
    style.configure("Status.TLabel", 
                   font=("Inter", 9), 
                   foreground=muted_color, 
                   background='#F9FAFB')

    # Store colors in root for non-ttk widgets
    setattr(root, 'canvas_bg', background_color)
    setattr(root, 'text_bg', card_color)
    setattr(root, 'text_fg', text_color)
    setattr(root, 'text_font', ("Inter", 11))
    setattr(root, 'oddrow_color', '#FAFAFA')
    setattr(root, 'evenrow_color', card_color)

    # Enhanced styles for InstruksiTable
    setattr(root, 'header_font', ("Inter", 12, "bold"))
    setattr(root, 'combobox_font', ("Inter", 11))
    setattr(root, 'dateentry_font', ("Inter", 13))
    setattr(root, 'active_row_bg', '#EFF6FF')
    setattr(root, 'inactive_row_bg', card_color)

    # Modern tooltip
    setattr(root, 'tooltip_bg', '#1F2937')
    setattr(root, 'tooltip_fg', card_color)
    setattr(root, 'tooltip_font', ("Inter", 9))
    
    # Status labels with better hierarchy
    setattr(root, 'status_label_font', ("Inter", 8))
    setattr(root, 'status_label_fg', accent_color)
    setattr(root, 'column_info_label_fg', success_color)
    setattr(root, 'structure_info_label_fg', '#8B5CF6')
    setattr(root, 'structure_info_label_font', ("Inter", 7))
    setattr(root, 'paging_info_label_font', ("Inter", 9))

    # === Custom Color Palette (Preserved) ===
    blue_color = '#1D4ED8'         # Deep blue dominan
    yellow_gold_color = '#FDE047'  # Bright golden yellow
    white_color = '#FFFFFF'        # Putih (untuk garis putus-putus dan background)
    # =========================================
    
    # Store custom colors in root
    setattr(root, 'blue_color', blue_color)
    setattr(root, 'yellow_gold_color', yellow_gold_color)
    setattr(root, 'white_color', white_color)
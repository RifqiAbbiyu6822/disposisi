import tkinter as tk

def attach_tooltip(widget, text, icon_path=None):
    def on_enter(event):
        x = widget.winfo_rootx() + 25
        y = widget.winfo_rooty() + 25
        tooltip = tk.Toplevel(widget)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{x}+{y}")
        if icon_path:
            try:
                tooltip.iconbitmap(icon_path)
            except:
                pass
        label = tk.Label(tooltip, text=text, background=getattr(widget.master, 'tooltip_bg', '#1e293b'), foreground=getattr(widget.master, 'tooltip_fg', 'white'), relief="solid", borderwidth=1, font=getattr(widget.master, 'tooltip_font', ("Segoe UI", 9)), padx=8, pady=4)
        label.pack()
        widget.tooltip = tooltip
    def on_leave(event):
        if hasattr(widget, 'tooltip') and widget.tooltip:
            widget.tooltip.destroy()
            widget.tooltip = None
    widget.bind("<Enter>", on_enter)
    widget.bind("<Leave>", on_leave)

def add_tooltips(input_widgets, tooltips, icon_path=None):
    for key, text in tooltips.items():
        widget = input_widgets.get(key)
        if widget:
            attach_tooltip(widget, text, icon_path) 
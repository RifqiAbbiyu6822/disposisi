from tkinter import ttk

def create_status_bar(parent):
    status_frame = ttk.Frame(parent, style="Status.TFrame", height=30)
    status_frame.pack(side="bottom", fill="x")
    status_frame.pack_propagate(False)
    status_message = ttk.Label(status_frame, text="Ready", style="Status.TLabel")
    status_message.pack(side="left", padx=15, pady=8)
    separator = ttk.Separator(status_frame, orient="vertical")
    separator.pack(side="right", fill="y", padx=10)
    version_label = ttk.Label(status_frame, text="v2.0", style="Status.TLabel")
    version_label.pack(side="right", padx=15, pady=8)
    return status_frame, status_message 
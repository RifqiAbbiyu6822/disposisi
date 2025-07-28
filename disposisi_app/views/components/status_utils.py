def update_status(status_message_widget, message, window=None):
    if status_message_widget:
        status_message_widget.config(text=message)
        # Auto-clear status after 3 detik
        if window:
            window.after(3000, lambda: status_message_widget.config(text="Ready")) 
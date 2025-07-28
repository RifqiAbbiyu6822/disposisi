def setup_touchpad_gestures(window, on_mouse_down, on_mouse_up, on_mouse_drag, on_double_click, on_mousewheel, on_horizontal_scroll, on_pinch_gesture_start, on_pinch_gesture):
    window.bind("<Button-1>", on_mouse_down)
    window.bind("<ButtonRelease-1>", on_mouse_up)
    window.bind("<B1-Motion>", on_mouse_drag)
    window.bind("<Double-Button-1>", on_double_click)
    window.bind("<MouseWheel>", on_mousewheel)
    window.bind("<Shift-MouseWheel>", on_horizontal_scroll)
    window.bind("<Control-Button-1>", on_pinch_gesture_start)
    window.bind("<Control-B1-Motion>", on_pinch_gesture) 
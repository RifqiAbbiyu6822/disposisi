def setup_shortcuts(window, save_to_pdf, clear_form, select_next_tab, select_prev_tab):
    window.bind_all("<Control-s>", lambda e: save_to_pdf())
    window.bind_all("<Control-S>", lambda e: save_to_pdf())
    window.bind_all("<Control-Shift-C>", lambda e: clear_form())
    window.bind_all("<Control-Tab>", lambda e: select_next_tab())
    window.bind_all("<Control-Shift-Tab>", lambda e: select_prev_tab()) 
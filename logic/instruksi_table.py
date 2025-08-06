import tkinter as tk
from tkinter import ttk, Text
from tkcalendar import DateEntry
import logging

logging.basicConfig(level=logging.WARNING)

class InstruksiTable:
    def __init__(self, parent, posisi_options, initial_data=None, use_grid=False, grid_kwargs=None):
        self.parent = parent
        self.posisi_options = posisi_options
        self.data = initial_data or [{"posisi": posisi_options[0], "instruksi": "", "tanggal": ""}]
        self.row_widgets = []
        self.tooltip = None  # Untuk tooltip
        self.frame = ttk.LabelFrame(parent, text="Instruksi", padding="10", style="Card.TFrame")
        if use_grid:
            if grid_kwargs is None:
                grid_kwargs = {"row": 0, "column": 0, "sticky": "nsew", "padx": 0, "pady": 0}
            self.frame.grid(**grid_kwargs)
        else:
            self.frame.pack(fill="x", pady=5)
        self.render_table()

    def _update_data_from_widgets(self):
        # Simpan semua input dari widget ke self.data
        new_data = []
        for w in self.row_widgets:
            posisi = w["posisi_cb"].get().strip()
            instruksi = w["instruksi_text"].get("1.0", tk.END).strip()
            tanggal = w["tanggal_entry"].get()
            new_data.append({"posisi": posisi, "instruksi": instruksi, "tanggal": tanggal})
        logging.debug("[DEBUG] Data dari widget sebelum update:", new_data)
        if new_data:
            self.data = new_data

    def render_table(self):
        # Jangan update data dari widget di sini!
        # Destroy old widgets
        for w in self.row_widgets:
            if "posisi_cb" in w: w["posisi_cb"].destroy()
            if "instruksi_text" in w: w["instruksi_text"].destroy()
            if "tanggal_entry" in w: w["tanggal_entry"].destroy()
            if "delete_cb" in w: w["delete_cb"].destroy()
        self.row_widgets = []
        for widget in self.frame.winfo_children():
            widget.destroy()
        root = self.frame.winfo_toplevel()
        header_font = ("Segoe UI", 12, "bold")
        ttk.Label(self.frame, text="Hapus", font=header_font).grid(row=0, column=0, padx=2, pady=2, sticky="nsew")
        ttk.Label(self.frame, text="Posisi", font=header_font).grid(row=0, column=1, padx=2, pady=2, sticky="nsew")
        ttk.Label(self.frame, text="Isi Instruksi", font=header_font).grid(row=0, column=2, padx=2, pady=2, sticky="nsew")
        ttk.Label(self.frame, text="Tanggal", font=header_font).grid(row=0, column=3, padx=2, pady=2, sticky="nsew")
        for i, row in enumerate(self.data):
            delete_var = tk.BooleanVar(value=False)
            delete_cb = ttk.Checkbutton(self.frame, variable=delete_var)
            delete_cb.grid(row=i+1, column=0, padx=2, pady=2, sticky="nsew")
            other_selected = set(r.get("posisi", "") for j, r in enumerate(self.data) if j != i and r.get("posisi", ""))
            available_options = [p for p in self.posisi_options if p not in other_selected or p == row.get("posisi", "")]
            posisi_cb = ttk.Combobox(self.frame, values=available_options, font=("Segoe UI", 11))
            posisi_cb.set(row.get("posisi", available_options[0] if available_options else ""))
            posisi_cb.grid(row=i+1, column=1, padx=2, pady=2, sticky="nsew")
            posisi_cb['state'] = 'normal'
            self.attach_tooltip(posisi_cb, "Pilih jabatan yang memberi instruksi")
            def limit_posisi(event, cb=posisi_cb):
                val = cb.get()
                if len(val) > 40:
                    cb.set(val[:40])
            posisi_cb.bind('<KeyRelease>', limit_posisi)
            def on_posisi_change(event, idx=i, cb=posisi_cb):
                self.data[idx]["posisi"] = cb.get()
                self.render_table()
            posisi_cb.bind('<<ComboboxSelected>>', on_posisi_change)
            instruksi_text = Text(self.frame, height=2, width=35, wrap="word", font=("Segoe UI", 13), borderwidth=1, relief="solid", highlightthickness=0, bg='#ffffff', fg='#1e293b')
            instruksi_text.grid(row=i+1, column=2, padx=2, pady=2, sticky="new")
            if row.get("instruksi", ""): instruksi_text.insert("1.0", row["instruksi"])
            self.attach_tooltip(instruksi_text, "Isi instruksi maksimal 200 karakter")
            def limit_instruksi(event, txt=instruksi_text):
                val = txt.get("1.0", tk.END)
                if len(val) > 200:
                    txt.delete("1.0", tk.END)
                    txt.insert("1.0", val[:200])
            instruksi_text.bind('<KeyRelease>', limit_instruksi)
            tanggal_entry = DateEntry(self.frame, width=18, date_pattern="dd-mm-yyyy", font=("Segoe UI", 13))
            tanggal_entry.grid(row=i+1, column=3, padx=2, pady=2, sticky="new")
            self.attach_tooltip(tanggal_entry, "Tanggal instruksi diberikan atau deadline")
            # Improved date handling
            tanggal_val = row.get("tanggal", "")
            if tanggal_val:
                try:
                    import datetime
                    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d.%m.%Y"):
                        try:
                            dt = datetime.datetime.strptime(tanggal_val.strip(), fmt)
                            tanggal_entry.set_date(dt.date())
                            break
                        except ValueError:
                            continue
                    else:
                        tanggal_entry.delete(0, tk.END)
                        tanggal_entry.insert(0, tanggal_val)
                except Exception as e:
                    logging.error(f"Error setting date: {e}")
                    tanggal_entry.delete(0, tk.END)
                    tanggal_entry.insert(0, tanggal_val)
            # Shortcut dan highlight
            posisi_cb.bind('<Tab>', lambda e, txt=instruksi_text: (txt.focus_set(), 'break'))
            instruksi_text.bind('<Tab>', lambda e, ent=tanggal_entry: (ent.focus_set(), 'break'))
            tanggal_entry.bind('<Tab>', lambda e, idx=i: (self.add_row(), self.row_widgets[-1]["posisi_cb"].focus_set(), 'break'))
            instruksi_text.bind('<Control-Return>', lambda e: (self.add_row(), self.row_widgets[-1]["posisi_cb"].focus_set(), 'break'))
            # Highlight baris aktif
            def on_focus_in(event, row_idx=i):
                for j, wdict in enumerate(self.row_widgets):
                    bg = '#e0f7fa' if j == row_idx else '#ffffff'
                    wdict["posisi_cb"].configure(background=bg)
                    wdict["instruksi_text"].configure(bg=bg)
                    wdict["tanggal_entry"].configure(background=bg)
            posisi_cb.bind('<FocusIn>', on_focus_in)
            instruksi_text.bind('<FocusIn>', on_focus_in)
            tanggal_entry.bind('<FocusIn>', on_focus_in)
            self.row_widgets.append({
                "delete_var": delete_var,
                "delete_cb": delete_cb,
                "posisi_cb": posisi_cb,
                "instruksi_text": instruksi_text,
                "tanggal_entry": tanggal_entry
            })

    def add_row(self):
        self._update_data_from_widgets()
        logging.debug("[DEBUG] Data sebelum tambah baris:", self.data)
        self.data.append({"posisi": self.posisi_options[0], "instruksi": "", "tanggal": ""})
        logging.debug("[DEBUG] Data setelah tambah baris:", self.data)
        self.render_table()

    def remove_selected_rows(self):
        self._update_data_from_widgets()
        new_data = []
        for i, w in enumerate(self.row_widgets):
            if not w["delete_var"].get():
                posisi = w["posisi_cb"].get().strip()
                instruksi = w["instruksi_text"].get("1.0", tk.END).strip()
                tanggal = w["tanggal_entry"].get()
                new_data.append({"posisi": posisi, "instruksi": instruksi, "tanggal": tanggal})
        if not new_data:
            new_data = [{"posisi": self.posisi_options[0], "instruksi": "", "tanggal": ""}]
        self.data = new_data
        self.render_table()

    def get_data(self):
        result = []
        for w in self.row_widgets:
            posisi = w["posisi_cb"].get().strip()
            instruksi = w["instruksi_text"].get("1.0", tk.END).strip()
            tanggal = w["tanggal_entry"].get()
            result.append({"posisi": posisi, "instruksi": instruksi, "tanggal": tanggal})
        return result

    def clear(self):
        self.data = [{"posisi": self.posisi_options[0], "instruksi": "", "tanggal": ""}]
        self.render_table()

    def kosongkan_semua_baris(self):
        # Mengosongkan semua baris tanpa menghapus baris
        for row in self.data:
            row["posisi"] = ""
            row["instruksi"] = ""
            row["tanggal"] = ""
        self.render_table()
        # Pastikan widget tanggal benar-benar kosong
        for w in self.row_widgets:
            try:
                w["tanggal_entry"].delete(0, 'end')
            except Exception:
                pass

    def attach_tooltip(self, widget, text):
        def on_enter(event):
            root = widget.winfo_toplevel()
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + 20
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            label = tk.Label(self.tooltip, text=text, background='#ffffe0', relief="solid", borderwidth=1, font=("Segoe UI", 9))
            label.pack(ipadx=4, ipady=2)
        def on_leave(event):
            if hasattr(self, 'tooltip') and self.tooltip:
                self.tooltip.destroy()
                self.tooltip = None
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave) 
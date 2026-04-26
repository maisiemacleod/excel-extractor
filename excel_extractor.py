"""
Excel Field Extractor v5
Drag & drop a folder, configure search parameters, extract cell values from Excel files.
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False


# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------

def find_excel_files(folder):
    results = []
    for root, dirs, files in os.walk(folder):
        for fname in files:
            if fname.lower().endswith(('.xlsx', '.xls')):
                results.append(os.path.join(root, fname))
    return results


# ---------------------------------------------------------------------------
# Sheet selection helpers
# ---------------------------------------------------------------------------

def select_sheet_xlsx(wb, sheet_mode, sheet_val):
    if sheet_mode == 'index':
        n = int(sheet_val)
        idx = (n - 1) if n > 0 else (len(wb.worksheets) + n)
        return wb.worksheets[idx]
    else:
        return wb[str(sheet_val)]


def select_sheet_xls(wb, sheet_mode, sheet_val):
    if sheet_mode == 'index':
        n = int(sheet_val)
        idx = (n - 1) if n > 0 else (wb.nsheets + n)
        return wb.sheet_by_index(idx)
    else:
        return wb.sheet_by_name(str(sheet_val))


# ---------------------------------------------------------------------------
# Cell iteration
# ---------------------------------------------------------------------------

def iter_cells_xlsx_ws(ws):
    for row in ws.iter_rows():
        for cell in row:
            yield cell.row, cell.column, cell.value


def iter_cells_xls_ws(ws):
    for r in range(ws.nrows):
        for c in range(ws.ncols):
            yield r + 1, c + 1, ws.cell_value(r, c)


# ---------------------------------------------------------------------------
# Match logic
# ---------------------------------------------------------------------------

def cell_matches(val, search_text, exact):
    if val is None:
        return False
    s = str(val)
    return s == search_text if exact else search_text in s


# ---------------------------------------------------------------------------
# Core processing
# ---------------------------------------------------------------------------

def process_file(filepath, search_text, exact, occurrence_n,
                 offset_r, offset_d, sheet_mode, sheet_val):
    fname = os.path.basename(filepath)
    ext = os.path.splitext(filepath)[1].lower()

    try:
        if ext == '.xlsx':
            if not HAS_OPENPYXL:
                return {"error": f"{fname}: openpyxl not installed"}
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            try:
                ws = select_sheet_xlsx(wb, sheet_mode, sheet_val)
                cells = list(iter_cells_xlsx_ws(ws))
            except Exception as e:
                wb.close()
                return {"error": f"{fname}: sheet not found ({e})"}
            wb.close()

        elif ext == '.xls':
            if not HAS_XLRD:
                return {"error": f"{fname}: xlrd not installed"}
            wb = xlrd.open_workbook(filepath)
            try:
                ws = select_sheet_xls(wb, sheet_mode, sheet_val)
                cells = list(iter_cells_xls_ws(ws))
            except Exception as e:
                return {"error": f"{fname}: sheet not found ({e})"}

        else:
            return None

    except Exception as e:
        return {"error": f"{fname}: {e}"}

    cells_dict = {(r, c): v for r, c, v in cells}
    matches = [(r, c, v) for r, c, v in cells if cell_matches(v, search_text, exact)]
    total = len(matches)

    if total == 0:
        return None

    try:
        idx = (occurrence_n - 1) if occurrence_n > 0 else (total + occurrence_n)
        match_row, match_col, match_val = matches[idx]
    except IndexError:
        return None

    read_val = cells_dict.get((match_row + offset_d, match_col + offset_r))

    return {
        "filename": fname,
        "field_content": str(match_val),
        "total_occurrences": total,
        "occur_seq": idx + 1,
        "match_row": match_row,
        "match_col": match_col,
        "read_value": str(read_val) if read_val is not None else "",
    }


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

_HEADERS = [
    "\u6587\u4ef6\u540d",
    "\u5b57\u6bb5\u5185\u5bb9",
    "\u5b57\u6bb5\u51fa\u73b0\u6b21\u6570",
    "\u51fa\u73b0\u7684\u6b21\u5e8f\u6570",
    "\u5339\u914d\u5355\u5143\u683c\u884c",
    "\u5339\u914d\u5355\u5143\u683c\u5217",
    "\u8bfb\u53d6\u5230\u7684\u503c",
]


def _row_values(row):
    return [
        row.get('filename', ''),
        row.get('field_content', ''),
        str(row.get('total_occurrences', '')),
        str(row.get('occur_seq', '')),
        str(row.get('match_row', '')),
        str(row.get('match_col', '')),
        row.get('read_value', ''),
    ]


def export_txt(rows, filepath):
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write('\t'.join(_HEADERS) + '\n')
        for row in rows:
            f.write('\t'.join(_row_values(row)) + '\n')


def export_csv(rows, filepath):
    import csv
    with open(filepath, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(_HEADERS)
        for row in rows:
            writer.writerow(_row_values(row))


def export_xlsx(rows, filepath):
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl not installed")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(_HEADERS)
    for row in rows:
        vals = _row_values(row)
        for i in (2, 3, 4, 5):
            try:
                vals[i] = int(vals[i])
            except (ValueError, TypeError):
                pass
        ws.append(vals)
    wb.save(filepath)


# ---------------------------------------------------------------------------
# Theme / style constants
# ---------------------------------------------------------------------------

BG = "#1e1e2e"           # dark base
SURFACE = "#2a2a3e"      # card / frame bg
ACCENT = "#7c6af7"       # purple accent
ACCENT2 = "#5a4fcf"      # darker accent (hover/press)
TEXT = "#cdd6f4"         # primary text
SUBTEXT = "#a6adc8"      # secondary text
GREEN = "#a6e3a1"        # success
RED = "#f38ba8"          # error
BORDER = "#45475a"       # border / separator
DROP_BG = "#2a2a3e"
DROP_ACTIVE = "#313149"
ENTRY_BG = "#313149"
BTN_BG = ACCENT
BTN_FG = "#ffffff"
ROW_ODD = "#252536"
ROW_EVEN = "#2a2a3e"
FONT_MAIN = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 11, "bold")
FONT_SMALL = ("Segoe UI", 9)


def apply_theme(root):
    style = ttk.Style(root)
    style.theme_use('default')

    style.configure(".", background=BG, foreground=TEXT,
                    font=FONT_MAIN, borderwidth=0, relief="flat")
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG, foreground=TEXT, font=FONT_MAIN)
    style.configure("TLabelframe", background=SURFACE, foreground=SUBTEXT,
                    font=FONT_BOLD, bordercolor=BORDER, relief="flat", padding=6)
    style.configure("TLabelframe.Label", background=SURFACE, foreground=ACCENT,
                    font=FONT_BOLD)

    style.configure("TCheckbutton", background=SURFACE, foreground=TEXT,
                    font=FONT_MAIN, focuscolor=SURFACE)
    style.map("TCheckbutton", background=[("active", SURFACE)])

    style.configure("TRadiobutton", background=SURFACE, foreground=TEXT,
                    font=FONT_MAIN, focuscolor=SURFACE)
    style.map("TRadiobutton", background=[("active", SURFACE)])

    style.configure("TEntry", fieldbackground=ENTRY_BG, foreground=TEXT,
                    insertcolor=TEXT, borderwidth=1, relief="flat",
                    padding=(6, 4))
    style.map("TEntry", fieldbackground=[("readonly", SURFACE)])

    style.configure("TSpinbox", fieldbackground=ENTRY_BG, foreground=TEXT,
                    background=ENTRY_BG, arrowcolor=SUBTEXT,
                    borderwidth=1, relief="flat", padding=(4, 4))

    style.configure("Accent.TButton", background=ACCENT, foreground=BTN_FG,
                    font=FONT_BOLD, padding=(16, 7), relief="flat",
                    borderwidth=0)
    style.map("Accent.TButton",
              background=[("active", ACCENT2), ("disabled", BORDER)],
              foreground=[("disabled", SUBTEXT)])

    style.configure("TButton", background=SURFACE, foreground=TEXT,
                    font=FONT_MAIN, padding=(12, 6), relief="flat",
                    borderwidth=1)
    style.map("TButton",
              background=[("active", ENTRY_BG)],
              foreground=[("disabled", SUBTEXT)])

    style.configure("TProgressbar", troughcolor=SURFACE, background=ACCENT,
                    borderwidth=0, thickness=6)

    style.configure("Treeview", background=ROW_ODD, foreground=TEXT,
                    fieldbackground=ROW_ODD, font=FONT_SMALL,
                    rowheight=24, borderwidth=0, relief="flat")
    style.configure("Treeview.Heading", background=SURFACE, foreground=ACCENT,
                    font=FONT_BOLD, borderwidth=0, relief="flat")
    style.map("Treeview",
              background=[("selected", ACCENT2)],
              foreground=[("selected", "#ffffff")])
    style.map("Treeview.Heading", background=[("active", ENTRY_BG)])

    style.configure("TScrollbar", background=SURFACE, troughcolor=BG,
                    borderwidth=0, arrowcolor=SUBTEXT, width=10)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel \u5b57\u6bb5\u63d0\u53d6\u5de5\u5177")
        self.root.configure(bg=BG)
        self.root.resizable(True, True)
        self.folder_path = tk.StringVar()
        self.results = []
        apply_theme(root)
        self._build_ui()

    def _section(self, parent, title):
        f = ttk.LabelFrame(parent, text=title)
        f.pack(fill='x', padx=12, pady=(6, 2))
        return f

    def _build_ui(self):
        # ---- Header bar ----
        header = tk.Frame(self.root, bg=ACCENT, height=48)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text="  Excel \u5b57\u6bb5\u63d0\u53d6\u5de5\u5177",
                 bg=ACCENT, fg="#ffffff", font=("Segoe UI", 13, "bold")).pack(side='left', padx=4)
        tk.Label(header, text="v5",
                 bg=ACCENT, fg="#d0cbff", font=("Segoe UI", 9)).pack(side='left')

        # Scrollable main area
        canvas = tk.Canvas(self.root, bg=BG, highlightthickness=0)
        vscroll = ttk.Scrollbar(self.root, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        main = tk.Frame(canvas, bg=BG)
        canvas_window = canvas.create_window((0, 0), window=main, anchor='nw')

        def _on_frame_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_configure(e):
            canvas.itemconfig(canvas_window, width=e.width)
        main.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # Mouse wheel scroll
        def _on_mousewheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        pad = {"padx": 10, "pady": 4}

        # ---- Drop zone ----
        frame_folder = self._section(main, "\u8f93\u5165\u6587\u4ef6\u5939")
        frame_folder.configure(labelanchor='nw')

        self.drop_frame = tk.Frame(frame_folder, bg=DROP_BG, relief="flat",
                                   highlightbackground=BORDER, highlightthickness=1,
                                   cursor="hand2")
        self.drop_frame.pack(fill='x', padx=4, pady=(4, 6))

        inner = tk.Frame(self.drop_frame, bg=DROP_BG)
        inner.pack(pady=14)

        icon_label = tk.Label(inner, text="\U0001f4c2", bg=DROP_BG,
                               font=("Segoe UI", 22), fg=ACCENT)
        icon_label.pack()

        if HAS_DND:
            drop_txt = "\u62d6\u653e\u6587\u4ef6\u5939\u5230\u6b64\u5904\uff0c\u6216\u70b9\u51fb\u9009\u62e9"
        else:
            drop_txt = "\u70b9\u51fb\u9009\u62e9\u6587\u4ef6\u5939"
        self.drop_hint = tk.Label(inner, text=drop_txt, bg=DROP_BG,
                                   fg=SUBTEXT, font=FONT_SMALL)
        self.drop_hint.pack()

        self.drop_path_label = tk.Label(self.drop_frame, text="",
                                         bg=DROP_BG, fg=GREEN,
                                         font=("Segoe UI", 9, "bold"),
                                         wraplength=580)
        self.drop_path_label.pack(pady=(0, 6))

        for w in (self.drop_frame, inner, icon_label, self.drop_hint):
            w.bind("<Button-1>", self._choose_folder)
            w.bind("<Enter>", lambda e: self.drop_frame.config(bg=DROP_ACTIVE,
                highlightbackground=ACCENT))
            w.bind("<Leave>", lambda e: self.drop_frame.config(bg=DROP_BG,
                highlightbackground=BORDER))

        if HAS_DND:
            try:
                self.drop_frame.drop_target_register(DND_FILES)
                self.drop_frame.dnd_bind('<<Drop>>', self._on_drop)
                self.drop_frame.dnd_bind('<<DragEnter>>', lambda e: e.action)
                self.drop_frame.dnd_bind('<<DragOver>>', lambda e: e.action)
            except Exception:
                pass

        btn_row = tk.Frame(frame_folder, bg=SURFACE)
        btn_row.pack(fill='x', padx=4, pady=(0, 4))
        ttk.Button(btn_row, text="\U0001f4c1  \u9009\u62e9\u6587\u4ef6\u5939",
                   command=self._choose_folder).pack(side='left', padx=4)

        # ---- Parameters: two-column layout ----
        frame_params = self._section(main, "\u67e5\u627e\u53c2\u6570")

        left = tk.Frame(frame_params, bg=SURFACE)
        left.pack(side='left', fill='both', expand=True, padx=(4, 2), pady=4)
        right = tk.Frame(frame_params, bg=SURFACE)
        right.pack(side='left', fill='both', expand=True, padx=(2, 4), pady=4)

        def lbl(parent, text, row, col, **kw):
            tk.Label(parent, text=text, bg=SURFACE, fg=SUBTEXT,
                     font=FONT_SMALL).grid(row=row, column=col, sticky='e',
                                           padx=(4,2), pady=3)

        # Left column
        lbl(left, "\u5b9a\u4f4d\u5185\u5bb9 X:", 0, 0)
        self.search_text = tk.StringVar()
        ent = ttk.Entry(left, textvariable=self.search_text, width=24)
        ent.grid(row=0, column=1, sticky='ew', padx=(0,4), pady=3)

        self.exact_match = tk.BooleanVar(value=False)
        cb = ttk.Checkbutton(left, text="\u7cbe\u786e\u5339\u914d",
                              variable=self.exact_match)
        cb.grid(row=0, column=2, padx=4, pady=3)

        lbl(left, "\u7b2c\u51e0\u6b21\u51fa\u73b0 n:", 1, 0)
        self.occurrence_n = tk.IntVar(value=1)
        ttk.Spinbox(left, from_=-9999, to=9999, textvariable=self.occurrence_n,
                    width=7).grid(row=1, column=1, sticky='w', padx=(0,4), pady=3)
        tk.Label(left, text="\u8d1f=\u5012\u6570", bg=SURFACE,
                 fg=SUBTEXT, font=FONT_SMALL).grid(row=1, column=2, padx=4)

        left.columnconfigure(1, weight=1)

        # Right column
        lbl(right, "\u5411\u53f3\u79fb\u52a8 r:", 0, 0)
        self.offset_r = tk.IntVar(value=0)
        ttk.Spinbox(right, from_=-9999, to=9999, textvariable=self.offset_r,
                    width=7).grid(row=0, column=1, sticky='w', padx=(0,4), pady=3)

        lbl(right, "\u5411\u4e0b\u79fb\u52a8 d:", 1, 0)
        self.offset_d = tk.IntVar(value=0)
        ttk.Spinbox(right, from_=-9999, to=9999, textvariable=self.offset_d,
                    width=7).grid(row=1, column=1, sticky='w', padx=(0,4), pady=3)

        right.columnconfigure(1, weight=1)

        # ---- Sheet selection ----
        frame_sheet = self._section(main, "Sheet \u9009\u62e9")

        sheet_inner = tk.Frame(frame_sheet, bg=SURFACE)
        sheet_inner.pack(fill='x', padx=4, pady=4)

        self.sheet_mode = tk.StringVar(value='index')

        ttk.Radiobutton(sheet_inner, text="\u7b2c\u51e0\u4e2a Sheet",
                        variable=self.sheet_mode, value='index',
                        command=self._update_sheet_mode).grid(row=0, column=0, padx=4, pady=3, sticky='w')
        self.sheet_index = tk.IntVar(value=1)
        self.sheet_index_spin = ttk.Spinbox(sheet_inner, from_=-9999, to=9999,
                                             textvariable=self.sheet_index, width=7)
        self.sheet_index_spin.grid(row=0, column=1, padx=4, pady=3, sticky='w')
        tk.Label(sheet_inner, text="\u8d1f=\u5012\u6570", bg=SURFACE,
                 fg=SUBTEXT, font=FONT_SMALL).grid(row=0, column=2, padx=4)

        ttk.Radiobutton(sheet_inner, text="Sheet \u540d\u79f0",
                        variable=self.sheet_mode, value='name',
                        command=self._update_sheet_mode).grid(row=1, column=0, padx=4, pady=3, sticky='w')
        self.sheet_name = tk.StringVar()
        self.sheet_name_entry = ttk.Entry(sheet_inner, textvariable=self.sheet_name, width=20)
        self.sheet_name_entry.grid(row=1, column=1, columnspan=2, padx=4, pady=3, sticky='ew')
        sheet_inner.columnconfigure(1, weight=1)
        self._update_sheet_mode()

        # ---- Export format + Run button row ----
        action_frame = tk.Frame(main, bg=BG)
        action_frame.pack(fill='x', padx=12, pady=8)

        fmt_frame = tk.Frame(action_frame, bg=BG)
        fmt_frame.pack(side='left')
        tk.Label(fmt_frame, text="\u5bfc\u51fa\u683c\u5f0f:", bg=BG,
                 fg=SUBTEXT, font=FONT_SMALL).pack(side='left', padx=(0,6))
        self.export_fmt = tk.StringVar(value="xlsx")
        for fmt in ["TXT", "CSV", "XLSX"]:
            ttk.Radiobutton(fmt_frame, text=fmt,
                            variable=self.export_fmt, value=fmt.lower()).pack(side='left', padx=4)

        self.run_btn = ttk.Button(action_frame, text="\u25b6  \u5f00\u59cb\u63d0\u53d6",
                                   style="Accent.TButton", command=self._run)
        self.run_btn.pack(side='right', padx=4)

        # Export button right next to run button (always visible)
        self.export_btn_top = ttk.Button(action_frame, text="\u2193  \u5bfc\u51fa\u7ed3\u679c",
                                          command=self._export, state='disabled')
        self.export_btn_top.pack(side='right', padx=4)

        # ---- Progress ----
        prog_frame = tk.Frame(main, bg=BG)
        prog_frame.pack(fill='x', padx=12, pady=(0, 4))

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(prog_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill='x')

        self.status_label = tk.Label(prog_frame, text="", bg=BG,
                                      fg=SUBTEXT, font=FONT_SMALL, anchor='w')
        self.status_label.pack(fill='x', pady=(2, 0))

        # ---- Results table ----
        frame_table = tk.Frame(main, bg=BG)
        frame_table.pack(fill='both', expand=True, padx=12, pady=(4, 4))

        tk.Label(frame_table, text="\u63d0\u53d6\u7ed3\u679c",
                 bg=BG, fg=ACCENT, font=FONT_BOLD).pack(anchor='w', pady=(0, 4))

        tree_frame = tk.Frame(frame_table, bg=SURFACE,
                               highlightbackground=BORDER, highlightthickness=1)
        tree_frame.pack(fill='both', expand=True)

        cols = ["\u6587\u4ef6\u540d", "\u5b57\u6bb5\u5185\u5bb9", "\u51fa\u73b0\u6b21\u6570",
                "\u6b21\u5e8f\u6570", "\u884c", "\u5217", "\u8bfb\u53d6\u503c"]
        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=12)
        col_widths = [160, 200, 80, 70, 50, 50, 160]
        for col, w in zip(cols, col_widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, minwidth=40, anchor='w')

        self.tree.tag_configure('odd', background=ROW_ODD)
        self.tree.tag_configure('even', background=ROW_EVEN)
        self.tree.tag_configure('error', background="#3d1a1a", foreground=RED)

        self.tree.pack(fill='both', expand=True, side='left')
        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        sb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=sb.set)

        # ---- Bottom bar ----
        bottom = tk.Frame(main, bg=BG)
        bottom.pack(fill='x', padx=12, pady=(4, 12))

        self.result_count = tk.Label(bottom, text="", bg=BG,
                                      fg=SUBTEXT, font=FONT_SMALL)
        self.result_count.pack(side='left')

        self.export_btn = ttk.Button(bottom, text="\u2193  \u5bfc\u51fa\u7ed3\u679c",
                                      command=self._export, state='disabled')
        self.export_btn.pack(side='right', padx=4)

    def _update_sheet_mode(self):
        mode = self.sheet_mode.get()
        if mode == 'index':
            self.sheet_index_spin.config(state='normal')
            self.sheet_name_entry.config(state='disabled')
        else:
            self.sheet_index_spin.config(state='disabled')
            self.sheet_name_entry.config(state='normal')

    def _choose_folder(self, event=None):
        path = filedialog.askdirectory(title="\u9009\u62e9\u6587\u4ef6\u5939")
        if path:
            self._set_folder(path)

    def _set_folder(self, path):
        self.folder_path.set(path)
        short = path if len(path) < 60 else "..." + path[-57:]
        self.drop_path_label.config(text="\u2714  " + short)
        self.drop_hint.config(fg=GREEN)

    def _on_drop(self, event):
        path = event.data.strip()
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]
        if os.path.isdir(path):
            self._set_folder(path)
        else:
            messagebox.showwarning("\u63d0\u793a", "\u8bf7\u62d6\u653e\u6587\u4ef6\u5939\uff0c\u800c\u975e\u6587\u4ef6")
        return event.action

    def _run(self):
        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("\u63d0\u793a", "\u8bf7\u5148\u9009\u62e9\u6587\u4ef6\u5939")
            return
        search_text = self.search_text.get()
        if not search_text:
            messagebox.showwarning("\u63d0\u793a", "\u8bf7\u8f93\u5165\u5b9a\u4f4d\u5185\u5bb9 X")
            return

        sheet_mode = self.sheet_mode.get()
        sheet_val = self.sheet_index.get() if sheet_mode == 'index' else self.sheet_name.get()
        if sheet_mode == 'name' and not sheet_val:
            messagebox.showwarning("\u63d0\u793a", "\u8bf7\u8f93\u5165 Sheet \u540d\u79f0")
            return

        self.run_btn.config(state='disabled')
        self.export_btn.config(state='disabled')
        self.export_btn_top.config(state='disabled')
        self.results = []
        self.result_count.config(text="")
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.progress_var.set(0)

        def worker():
            files = find_excel_files(folder)
            total = len(files)
            if total == 0:
                self.root.after(0, lambda: self._finish([], "\u672a\u627e\u5230 Excel \u6587\u4ef6"))
                return

            results = []
            for i, fp in enumerate(files):
                res = process_file(
                    fp, search_text,
                    self.exact_match.get(),
                    self.occurrence_n.get(),
                    self.offset_r.get(),
                    self.offset_d.get(),
                    sheet_mode, sheet_val,
                )
                if res:
                    results.append(res)
                pct = (i + 1) / total * 100
                self.root.after(0, lambda p=pct, f=fp: self._update_progress(p, os.path.basename(f)))

            msg = f"\u5904\u7406 {total} \u4e2a\u6587\u4ef6\uff0c\u5339\u914d {len(results)} \u6761\u7ed3\u679c"
            self.root.after(0, lambda: self._finish(results, msg))

        threading.Thread(target=worker, daemon=True).start()

    def _update_progress(self, pct, fname):
        self.progress_var.set(pct)
        self.status_label.config(text=f"\u6b63\u5728\u5904\u7406: {fname}")

    def _finish(self, results, msg):
        self.results = results
        self.status_label.config(text=msg)
        self.progress_var.set(100)
        for i, row in enumerate(results):
            tag = 'error' if 'error' in row else ('odd' if i % 2 == 0 else 'even')
            if 'error' in row:
                self.tree.insert('', 'end', values=(row['error'], '', '', '', '', '', ''), tags=(tag,))
            else:
                self.tree.insert('', 'end', values=(
                    row['filename'],
                    row['field_content'],
                    row['total_occurrences'],
                    row['occur_seq'],
                    row['match_row'],
                    row['match_col'],
                    row['read_value'],
                ), tags=(tag,))
        self.run_btn.config(state='normal')
        if results:
            self.export_btn.config(state='normal')
            self.export_btn_top.config(state='normal')
            self.result_count.config(
                text=f"\u5171 {len(results)} \u6761\u7ed3\u679c",
                fg=GREEN
            )

    def _export(self):
        fmt = self.export_fmt.get()
        ftypes = {
            "txt": [("Text files", "*.txt")],
            "csv": [("CSV files", "*.csv")],
            "xlsx": [("Excel files", "*.xlsx")],
        }
        path = filedialog.asksaveasfilename(
            defaultextension=f".{fmt}",
            filetypes=ftypes.get(fmt, [("All files", "*.*")]),
            title="\u4fdd\u5b58\u7ed3\u679c"
        )
        if not path:
            return
        try:
            if fmt == 'txt':
                export_txt(self.results, path)
            elif fmt == 'csv':
                export_csv(self.results, path)
            elif fmt == 'xlsx':
                export_xlsx(self.results, path)
            messagebox.showinfo("\u5bfc\u51fa\u6210\u529f", f"\u5df2\u5bfc\u51fa\u5230:\n{path}")
        except Exception as e:
            messagebox.showerror("\u5bfc\u51fa\u5931\u8d25", str(e))


def main():
    if HAS_DND:
        try:
            root = TkinterDnD.Tk()
        except Exception:
            root = tk.Tk()
    else:
        root = tk.Tk()
    root.geometry("860x720")
    root.minsize(640, 520)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

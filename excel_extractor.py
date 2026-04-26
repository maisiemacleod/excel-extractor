"""
Excel Field Extractor v3
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

    # Build lookup dict for offset reads
    cells_dict = {(r, c): v for r, c, v in cells}

    # Find all matching cells (row-first order)
    matches = [(r, c, v) for r, c, v in cells if cell_matches(v, search_text, exact)]
    total = len(matches)

    if total == 0:
        return None

    # Select n-th occurrence (1-based; negative counts from end)
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
    "\u6587\u4ef6\u540d",          # 文件名
    "\u5b57\u6bb5\u5185\u5bb9",    # 字段内容
    "\u5b57\u6bb5\u51fa\u73b0\u6b21\u6570",  # 字段出现次数
    "\u51fa\u73b0\u7684\u6b21\u5e8f\u6570",  # 出现的次序数
    "\u5339\u914d\u5355\u5143\u683c\u884c",  # 匹配单元格行
    "\u5339\u914d\u5355\u5143\u683c\u5217",  # 匹配单元格列
    "\u8bfb\u53d6\u5230\u7684\u503c",        # 读取到的值
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
        # Keep numeric columns as numbers in xlsx
        for i in (2, 3, 4, 5):
            try:
                vals[i] = int(vals[i])
            except (ValueError, TypeError):
                pass
        ws.append(vals)
    wb.save(filepath)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel \u5b57\u6bb5\u63d0\u53d6\u5de5\u5177")
        self.root.resizable(True, True)
        self.folder_path = tk.StringVar()
        self.results = []
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # --- Folder selection ---
        frame_folder = ttk.LabelFrame(self.root, text="\u8f93\u5165\u6587\u4ef6\u5939")
        frame_folder.pack(fill='x', **pad)

        if HAS_DND:
            drop_text = "\u62d6\u653e\u6587\u4ef6\u5939\u5230\u6b64\u5904\uff0c\u6216\u70b9\u51fb\u9009\u62e9"
        else:
            drop_text = "\u70b9\u51fb\u300c\u9009\u62e9\u6587\u4ef6\u5939\u300d\u6309\u9215\u9009\u62e9\uff08\u62d6\u653e\u4e0d\u53ef\u7528\uff0c\u8bf7\u5b89\u88c5 tkinterdnd2\uff09"

        self.drop_label = tk.Label(
            frame_folder,
            text=drop_text,
            bg="#e8f0fe", relief="groove", height=3, cursor="hand2"
        )
        self.drop_label.pack(fill='x', padx=4, pady=4)
        self.drop_label.bind("<Button-1>", self._choose_folder)

        if HAS_DND:
            try:
                self.drop_label.drop_target_register(DND_FILES)
                self.drop_label.dnd_bind('<<Drop>>', self._on_drop)
                self.drop_label.dnd_bind('<<DragEnter>>', lambda e: e.action)
                self.drop_label.dnd_bind('<<DragOver>>', lambda e: e.action)
            except Exception:
                pass  # DND registration failed; button fallback still works

        entry_folder = ttk.Entry(frame_folder, textvariable=self.folder_path, state='readonly')
        entry_folder.pack(fill='x', padx=4, pady=(0, 4))

        btn_choose = ttk.Button(frame_folder, text="\u9009\u62e9\u6587\u4ef6\u5939", command=self._choose_folder)
        btn_choose.pack(side='left', padx=4, pady=(0, 4))

        # --- Search parameters ---
        frame_params = ttk.LabelFrame(self.root, text="\u67e5\u627e\u53c2\u6570")
        frame_params.pack(fill='x', **pad)

        ttk.Label(frame_params, text="\u5b9a\u4f4d\u5185\u5bb9 X:").grid(row=0, column=0, sticky='e', **pad)
        self.search_text = tk.StringVar()
        ttk.Entry(frame_params, textvariable=self.search_text, width=30).grid(row=0, column=1, sticky='ew', **pad)

        self.exact_match = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_params, text="\u7cbe\u786e\u67e5\u627e", variable=self.exact_match).grid(row=0, column=2, **pad)

        ttk.Label(frame_params, text="\u7b2c\u51e0\u6b21\u51fa\u73b0 n:").grid(row=1, column=0, sticky='e', **pad)
        self.occurrence_n = tk.IntVar(value=1)
        ttk.Spinbox(frame_params, from_=-9999, to=9999, textvariable=self.occurrence_n, width=8).grid(row=1, column=1, sticky='w', **pad)
        ttk.Label(frame_params, text="(\u8d1f\u6570=\u5012\u6570\uff0c-1=\u6700\u540e)").grid(row=1, column=2, **pad)

        ttk.Label(frame_params, text="\u5411\u53f3\u79fb\u52a8 r:").grid(row=2, column=0, sticky='e', **pad)
        self.offset_r = tk.IntVar(value=0)
        ttk.Spinbox(frame_params, from_=-9999, to=9999, textvariable=self.offset_r, width=8).grid(row=2, column=1, sticky='w', **pad)

        ttk.Label(frame_params, text="\u5411\u4e0b\u79fb\u52a8 d:").grid(row=3, column=0, sticky='e', **pad)
        self.offset_d = tk.IntVar(value=0)
        ttk.Spinbox(frame_params, from_=-9999, to=9999, textvariable=self.offset_d, width=8).grid(row=3, column=1, sticky='w', **pad)

        frame_params.columnconfigure(1, weight=1)

        # --- Sheet selection ---
        frame_sheet = ttk.LabelFrame(self.root, text="Sheet \u9009\u62e9")
        frame_sheet.pack(fill='x', **pad)

        self.sheet_mode = tk.StringVar(value='index')

        rb_idx = ttk.Radiobutton(frame_sheet, text="\u7b2c\u51e0\u4e2a Sheet",
                                  variable=self.sheet_mode, value='index',
                                  command=self._update_sheet_mode)
        rb_idx.grid(row=0, column=0, **pad)

        self.sheet_index = tk.IntVar(value=1)
        self.sheet_index_spin = ttk.Spinbox(frame_sheet, from_=-9999, to=9999,
                                             textvariable=self.sheet_index, width=8)
        self.sheet_index_spin.grid(row=0, column=1, sticky='w', **pad)
        ttk.Label(frame_sheet, text="(\u8d1f\u6570=\u5012\u6570\uff0c-1=\u6700\u540e)").grid(row=0, column=2, **pad)

        rb_name = ttk.Radiobutton(frame_sheet, text="Sheet \u540d\u79f0",
                                   variable=self.sheet_mode, value='name',
                                   command=self._update_sheet_mode)
        rb_name.grid(row=1, column=0, **pad)

        self.sheet_name = tk.StringVar()
        self.sheet_name_entry = ttk.Entry(frame_sheet, textvariable=self.sheet_name, width=20)
        self.sheet_name_entry.grid(row=1, column=1, sticky='ew', **pad)

        frame_sheet.columnconfigure(1, weight=1)
        self._update_sheet_mode()

        # --- Export format ---
        frame_export = ttk.LabelFrame(self.root, text="\u5bfc\u51fa\u683c\u5f0f")
        frame_export.pack(fill='x', **pad)

        self.export_fmt = tk.StringVar(value="xlsx")
        for fmt in ["txt", "csv", "xlsx"]:
            ttk.Radiobutton(frame_export, text=fmt.upper(), variable=self.export_fmt, value=fmt).pack(side='left', padx=8, pady=4)

        # --- Run button ---
        self.run_btn = ttk.Button(self.root, text="\u5f00\u59cb\u63d0\u53d6", command=self._run)
        self.run_btn.pack(**pad)

        # --- Progress ---
        self.progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(self.root, variable=self.progress_var, maximum=100).pack(fill='x', **pad)

        self.status_label = ttk.Label(self.root, text="")
        self.status_label.pack(**pad)

        # --- Results table ---
        frame_table = ttk.LabelFrame(self.root, text="\u63d0\u53d6\u7ed3\u679c")
        frame_table.pack(fill='both', expand=True, **pad)

        cols = ["\u6587\u4ef6\u540d", "\u5b57\u6bb5\u5185\u5bb9", "\u51fa\u73b0\u6b21\u6570",
                "\u6b21\u5e8f\u6570", "\u884c", "\u5217", "\u8bfb\u53d6\u503c"]
        self.tree = ttk.Treeview(frame_table, columns=cols, show='headings', height=10)
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110, anchor='w')
        self.tree.pack(fill='both', expand=True, side='left')

        sb = ttk.Scrollbar(frame_table, orient='vertical', command=self.tree.yview)
        sb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=sb.set)

        # --- Export button ---
        self.export_btn = ttk.Button(self.root, text="\u5bfc\u51fa\u7ed3\u679c",
                                      command=self._export, state='disabled')
        self.export_btn.pack(**pad)

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
            self.folder_path.set(path)
            self.drop_label.config(text=path)

    def _on_drop(self, event):
        path = event.data.strip()
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]
        if os.path.isdir(path):
            self.folder_path.set(path)
            self.drop_label.config(text=path)
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
        self.results = []
        for row in self.tree.get_children():
            self.tree.delete(row)

        def worker():
            files = find_excel_files(folder)
            total = len(files)
            if total == 0:
                self.root.after(0, lambda: self._finish([], "0 \u4e2aExcel\u6587\u4ef6"))
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

            msg = f"\u5904\u7406 {total} \u4e2a\u6587\u4ef6\uff0c\u627e\u5230 {len(results)} \u6761\u7ed3\u679c"
            self.root.after(0, lambda: self._finish(results, msg))

        threading.Thread(target=worker, daemon=True).start()

    def _update_progress(self, pct, fname):
        self.progress_var.set(pct)
        self.status_label.config(text=f"\u6b63\u5728\u5904\u7406: {fname}")

    def _finish(self, results, msg):
        self.results = results
        self.status_label.config(text=msg)
        self.progress_var.set(100)
        for row in results:
            if 'error' in row:
                self.tree.insert('', 'end', values=(row['error'], '', '', '', '', '', ''))
            else:
                self.tree.insert('', 'end', values=(
                    row['filename'],
                    row['field_content'],
                    row['total_occurrences'],
                    row['occur_seq'],
                    row['match_row'],
                    row['match_col'],
                    row['read_value'],
                ))
        self.run_btn.config(state='normal')
        if results:
            self.export_btn.config(state='normal')

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
            messagebox.showinfo("\u5bfc\u51fa\u6210\u529f", f"\u5df2\u5bfc\u51fa\u5230: {path}")
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
    root.geometry("820x680")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

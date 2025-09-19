import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import numpy as np
import matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

matplotlib.use("TkAgg")

INDEPENDENTS = [
    "m(g)",
    "Water (g)",
    "T (\u00B0C)",
    "P CO2 (bar)",
    "t (min)",
    "PDR (MPa/s)",
]

DEPENDENTS = [
    "\u00F8 (\u00B5m)",
    "N\u1D65 (cells\u00B7cm^3)",
    "\u03C1 foam (g/cm^3)",
    "OC (%)",
    "DSC Tm (\u00B0C)",
    "DSC Tg (\u00B0C)",
    "DSC Xc (%)",
]

ALLOWED_COLUMNS = INDEPENDENTS + DEPENDENTS


class HeatmapModule:
    def __init__(self, parent, settings_manager=None):
        self.root = tk.Toplevel(parent) if isinstance(parent, tk.Tk) else parent
        self.root.title("Heatmap Analysis")
        self.settings = settings_manager

        self.df_sheets: dict[str, pd.DataFrame] = {}
        self.current_sheet: str | None = None
        self.heatmap_rendered = False

        self.file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()

        self._build_ui()
        default_file = ''
        if self.settings is not None:
            last_file = str(self.settings.get('last_heatmap_file', '') or '')
            if last_file:
                default_file = last_file
            else:
                fallback = str(self.settings.get('last_output_file', '') or '')
                if fallback:
                    default_file = fallback
        if default_file:
            self.file_var.set(default_file)



    def _build_ui(self):
        container = ttk.Frame(self.root, padding=12)
        container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        file_row = ttk.Frame(container)
        file_row.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        file_row.columnconfigure(1, weight=1)
        ttk.Label(file_row, text="All_Results Excel:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_row, textvariable=self.file_var).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=6)
        ttk.Button(file_row, text="Browse", command=self._browse_file).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(file_row, text="Load", command=self._load_file).grid(row=0, column=3, padx=(6, 0))

        sheet_row = ttk.Frame(container)
        sheet_row.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        ttk.Label(sheet_row, text="Sheet:").grid(row=0, column=0, sticky=tk.W)
        self.sheet_combo = ttk.Combobox(sheet_row, textvariable=self.sheet_var, state="disabled")
        self.sheet_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(6, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_change)
        sheet_row.columnconfigure(1, weight=1)

        body = ttk.Frame(container)
        body.grid(row=2, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        container.rowconfigure(2, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        controls = ttk.Frame(body)
        controls.grid(row=0, column=0, sticky=(tk.N, tk.S), padx=(0, 12))
        controls.columnconfigure(0, weight=1)
        ttk.Label(controls, text="Columns (independent + dependent):").grid(row=0, column=0, sticky=tk.W)
        self.columns_list = tk.Listbox(controls, selectmode=tk.EXTENDED, exportselection=False, height=20, width=28, state='disabled')
        self.columns_list.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.W, tk.E), pady=(4, 4))
        self.columns_list.bind("<<ListboxSelect>>", self._on_selection_change)
        scroll = ttk.Scrollbar(controls, orient=tk.VERTICAL, command=self.columns_list.yview)
        scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.columns_list.configure(yscrollcommand=scroll.set)
        controls.rowconfigure(1, weight=1)
        self.select_all_btn = ttk.Button(controls, text="Select all", command=self._select_all, state='disabled')
        self.select_all_btn.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(6, 2))
        self.clear_btn = ttk.Button(controls, text="Clear selection", command=self._clear_selection, state='disabled')
        self.clear_btn.grid(row=3, column=0, sticky=(tk.W, tk.E))

        plot_frame = ttk.Frame(body)
        plot_frame.grid(row=0, column=1, sticky=(tk.N, tk.S, tk.W, tk.E))
        plot_frame.rowconfigure(0, weight=1)
        plot_frame.columnconfigure(0, weight=1)
        self.fig = Figure(figsize=(11.5, 7.5), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.heatmap_colorbar = None

        actions = ttk.Frame(container)
        actions.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        actions.columnconfigure((0, 1, 2), weight=1)
        self.render_btn = ttk.Button(actions, text="Render Heatmap", command=self._render_heatmap, state='disabled')
        self.render_btn.grid(row=0, column=0, padx=(0, 6), sticky=(tk.W, tk.E))
        self.save_btn = ttk.Button(actions, text="Save Heatmap", command=self._save_heatmap, state='disabled')
        self.save_btn.grid(row=0, column=1, padx=6, sticky=(tk.W, tk.E))
        self.copy_btn = ttk.Button(actions, text="Copy Heatmap", command=self._copy_heatmap, state='disabled')
        self.copy_btn.grid(row=0, column=2, padx=6, sticky=(tk.W, tk.E))

        self.status_var = tk.StringVar(value="Load an All_Results workbook to begin.")
        ttk.Label(container, textvariable=self.status_var, foreground="gray").grid(row=4, column=0, sticky=tk.W, pady=(6, 0))

    # ---------- File handling ----------
    def _browse_file(self):
        initialdir = None
        preset = self.file_var.get().strip()
        if preset and os.path.isdir(preset):
            initialdir = preset
        elif preset and os.path.isfile(preset):
            initialdir = os.path.dirname(preset)
        filename = filedialog.askopenfilename(
            title="Select All_Results Excel",
            initialdir=initialdir,
            filetypes=[("Excel files", "All_Results_*.xlsx"), ("Excel", "*.xlsx;*.xls")],
        )
        if filename:
            self.file_var.set(filename)

    def _load_file(self):
        path = self.file_var.get().strip()
        if not path:
            messagebox.showwarning("Missing file", "Please select an All_Results_*.xlsx file.")
            return
        if os.path.isdir(path):
            import glob
            candidates = sorted(glob.glob(os.path.join(path, "All_Results_*.xlsx")))
            if not candidates:
                messagebox.showwarning("Not found", "No All_Results_*.xlsx in the selected folder.")
                return
            path = candidates[-1]
            self.file_var.set(path)
        if not os.path.exists(path):
            messagebox.showerror("File not found", path)
            return

        try:
            sheets = pd.read_excel(path, sheet_name=None, engine="openpyxl")
        except Exception as exc:
            messagebox.showerror("Read error", str(exc))
            return

        valid = {}
        for sheet_name, df in sheets.items():
            if not isinstance(df, pd.DataFrame):
                continue
            if 'Water (g)' not in df.columns and 'Water' in df.columns:
                df = df.rename(columns={'Water': 'Water (g)'})
            cols = [c for c in ALLOWED_COLUMNS if c in df.columns]
            if len(cols) < 2:
                continue
            valid[sheet_name] = df[cols].copy()

        if not valid:
            messagebox.showerror("No usable sheets", "None of the sheets contain at least two allowed columns (independent/dependent).")
            return

        self.df_sheets = valid
        options = sorted(valid.keys(), key=lambda n: n.lower())
        self.sheet_combo.configure(state="readonly", values=options)
        self.sheet_var.set(options[0])
        self.current_sheet = options[0]
        self._populate_columns()
        self.render_btn.configure(state='normal')
        self.status_var.set(f"Loaded {len(options)} sheet(s) from {os.path.basename(path)}")
        self.heatmap_rendered = False
        self.save_btn.configure(state='disabled')
        self.copy_btn.configure(state='disabled')
        if self.settings is not None:
            self.settings.set("last_heatmap_file", path)

    def _on_sheet_change(self, _event=None):
        new_sheet = self.sheet_var.get()
        if new_sheet == self.current_sheet:
            return
        self.current_sheet = new_sheet
        self._populate_columns()
        self.heatmap_rendered = False
        self.save_btn.configure(state='disabled')
        self.copy_btn.configure(state='disabled')
        self.status_var.set(f"Sheet set to {new_sheet}. Select columns and render.")

    def _populate_columns(self):
        df = self.df_sheets.get(self.current_sheet)
        self.columns_list.configure(state='normal')
        self.columns_list.delete(0, tk.END)
        if df is None:
            self.columns_list.configure(state='disabled')
            self.select_all_btn.configure(state='disabled')
            self.clear_btn.configure(state='disabled')
            return
        cols = [c for c in df.columns if c in ALLOWED_COLUMNS]
        for col in cols:
            self.columns_list.insert(tk.END, col)
        self.columns_list.selection_set(0, tk.END)
        self.columns_list.configure(state='normal')
        self.select_all_btn.configure(state='normal')
        self.clear_btn.configure(state='normal')
        self._on_selection_change()

    # ---------- Column selection helpers ----------
    def _select_all(self):
        if self.columns_list['state'] == 'disabled':
            return
        self.columns_list.select_set(0, tk.END)
        self._on_selection_change()

    def _clear_selection(self):
        if self.columns_list['state'] == 'disabled':
            return
        self.columns_list.selection_clear(0, tk.END)
        self._on_selection_change()

    def _on_selection_change(self, _event=None):
        self.heatmap_rendered = False
        self.save_btn.configure(state='disabled')
        self.copy_btn.configure(state='disabled')

    # ---------- Heatmap rendering ----------
    def _render_heatmap(self):
        if self.current_sheet is None:
            messagebox.showwarning("No sheet", "Load a workbook and choose a sheet first.")
            return
        df = self.df_sheets.get(self.current_sheet)
        if df is None:
            messagebox.showwarning("No data", "The selected sheet has no usable data.")
            return
        selection = self.columns_list.curselection() if self.columns_list['state'] == 'normal' else ()
        if selection:
            cols = [self.columns_list.get(i) for i in selection]
        else:
            cols = list(df.columns)
        cols = [c for c in cols if c in df.columns]
        unique_cols = []
        for col in cols:
            if col not in unique_cols:
                unique_cols.append(col)
        if len(unique_cols) < 2:
            messagebox.showwarning("Not enough columns", "Select at least two columns for the heatmap.")
            return
        data = df[unique_cols].dropna(how='all')
        data = data.loc[:, data.nunique(dropna=True) > 1]
        if data.shape[1] < 2:
            messagebox.showwarning("Not enough columns", "At least two non-constant columns are required for the heatmap.")
            return

        corr = data.corr(method='spearman')
        if corr.isna().all().all():
            messagebox.showwarning("Invalid correlation", "Unable to compute Spearman correlation for the current selection.")
            return

        self.ax.clear()
        if self.heatmap_colorbar is not None:
            try:
                self.heatmap_colorbar.remove()
            except Exception:
                pass
            self.heatmap_colorbar = None

        im = self.ax.imshow(corr, cmap='coolwarm', vmin=-1, vmax=1)
        labels = list(corr.columns)
        ticks = list(range(len(labels)))
        self.ax.set_xticks(ticks)
        self.ax.set_xticklabels(labels, rotation=45, ha='right')
        self.ax.set_yticks(ticks)
        self.ax.set_yticklabels(labels)
        self.ax.set_title('Spearman correlation heatmap')

        for i, row_label in enumerate(labels):
            for j, col_label in enumerate(labels):
                value = corr.iat[i, j]
                display = '-' if pd.isna(value) else f"{value:.2f}"
                self.ax.text(j, i, display, ha='center', va='center', color='black', fontsize=9)

        self.fig.tight_layout()
        try:
            self.heatmap_colorbar = self.fig.colorbar(im, ax=self.ax, fraction=0.046, pad=0.04)
            self.heatmap_colorbar.set_label('Spearman $r_s$')
        except Exception:
            pass

        self.canvas.draw_idle()
        preview = ', '.join(labels[:6])
        if len(labels) > 6:
            preview += ', ...'
        self.status_var.set(f"Sheet: {self.current_sheet} | Heatmap columns ({len(labels)}): {preview}")
        self.heatmap_rendered = True
        self.save_btn.configure(state='normal')
        self.copy_btn.configure(state='normal')

    # ---------- Export helpers ----------
    def _default_filename(self, prefix: str, ext: str) -> str:
        sheet = (self.current_sheet or 'sheet').replace(' ', '_')
        ts = pd.Timestamp.now().strftime('%Y%m%d-%H%M')
        return f"{prefix}_{sheet}_{ts}.{ext}"

    def _save_heatmap(self):
        if not self.heatmap_rendered:
            messagebox.showwarning("No heatmap", "Render the heatmap before saving.")
            return
        default = self._default_filename('heatmap', 'png')
        filename = filedialog.asksaveasfilename(
            title="Save Heatmap",
            defaultextension='.png',
            initialfile=default,
            filetypes=[('PNG', '*.png'), ('TIFF', '*.tiff;*.tif'), ('SVG', '*.svg')],
        )
        if not filename:
            return
        try:
            self.fig.savefig(filename, dpi=300, facecolor='white', bbox_inches='tight')
            messagebox.showinfo("Saved", f"Heatmap saved to:\n{filename}")
        except Exception as exc:
            messagebox.showerror("Save error", str(exc))

    def _copy_heatmap(self):
        if not self.heatmap_rendered:
            messagebox.showwarning("No heatmap", "Render the heatmap before copying.")
            return
        if os.name != 'nt':
            messagebox.showwarning("Unsupported", "Clipboard copy is only available on Windows.")
            return
        try:
            import win32clipboard
            import win32con
        except ImportError:
            messagebox.showerror("Copy error", "pywin32 is required for clipboard support on Windows.")
            return
        import io as _io
        from PIL import Image

        buffer = _io.BytesIO()
        try:
            self.fig.savefig(buffer, format='png', dpi=300, facecolor='white', bbox_inches='tight')
            buffer.seek(0)
            image = Image.open(buffer).convert('RGB')
            with _io.BytesIO() as output:
                image.save(output, 'BMP')
                data = output.getvalue()[14:]
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32con.CF_DIB, data)
            finally:
                win32clipboard.CloseClipboard()
            messagebox.showinfo("Copied", "Heatmap copied to the clipboard.")
        except Exception as exc:
            messagebox.showerror("Copy error", str(exc))
        finally:
            buffer.close()


# Helper to show module standalone

def show(parent=None, settings_manager=None):
    root = parent if parent is not None else tk.Tk()
    HeatmapModule(root, settings_manager)
    if parent is None:
        root.mainloop()

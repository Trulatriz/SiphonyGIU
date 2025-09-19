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
        controls.columnconfigure(1, weight=1)
        ttk.Label(controls, text='Independent columns:').grid(row=0, column=0, sticky=tk.W)
        ttk.Label(controls, text='Dependent columns:').grid(row=0, column=1, sticky=tk.W)

        indep_container = ttk.Frame(controls)
        indep_container.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.W, tk.E), pady=(4, 4))
        indep_container.columnconfigure(0, weight=1)
        self.indep_list = tk.Listbox(indep_container, selectmode=tk.EXTENDED, exportselection=False, height=16, width=26, state='disabled')
        self.indep_list.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.indep_list.bind('<<ListboxSelect>>', self._on_selection_change)
        indep_scroll = ttk.Scrollbar(indep_container, orient=tk.VERTICAL, command=self.indep_list.yview)
        indep_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.indep_list.configure(yscrollcommand=indep_scroll.set)

        dep_container = ttk.Frame(controls)
        dep_container.grid(row=1, column=1, sticky=(tk.N, tk.S, tk.W, tk.E), pady=(4, 4))
        dep_container.columnconfigure(0, weight=1)
        self.dep_list = tk.Listbox(dep_container, selectmode=tk.EXTENDED, exportselection=False, height=16, width=26, state='disabled')
        self.dep_list.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.dep_list.bind('<<ListboxSelect>>', self._on_selection_change)
        dep_scroll = ttk.Scrollbar(dep_container, orient=tk.VERTICAL, command=self.dep_list.yview)
        dep_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.dep_list.configure(yscrollcommand=dep_scroll.set)

        controls.rowconfigure(1, weight=1)
        ttk.Label(controls, text='Correlation method:').grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(6, 0))
        self.method_var = tk.StringVar(value='Spearman')
        self.method_combo = ttk.Combobox(controls, textvariable=self.method_var, values=['Spearman', 'Pearson', 'Distance (dCor)'], state='readonly')
        self.method_combo.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E))
        self.method_combo.bind('<<ComboboxSelected>>', self._on_selection_change)

        self.select_all_btn = ttk.Button(controls, text='Select all', command=self._select_all, state='disabled')
        self.select_all_btn.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(6, 2))
        self.clear_btn = ttk.Button(controls, text='Clear selection', command=self._clear_selection, state='disabled')
        self.clear_btn.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=(6, 2))


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
        self.dep_list.configure(state='normal')
        self.dep_list.delete(0, tk.END)
        if df is None:
            self.dep_list.configure(state='disabled')
            self.select_all_btn.configure(state='disabled')
            self.clear_btn.configure(state='disabled')
            return
        cols = [c for c in df.columns if c in ALLOWED_COLUMNS]
        for col in cols:
            self.dep_list.insert(tk.END, col)
        self.dep_list.selection_set(0, tk.END)
        self.dep_list.configure(state='normal')
        self.select_all_btn.configure(state='normal')
        self.clear_btn.configure(state='normal')
        self._on_selection_change()

    # ---------- Column selection helpers ----------
    def _select_all(self):
        for lst in (self.indep_list, self.dep_list):
            if lst['state'] != 'disabled':
                lst.select_set(0, tk.END)
        self._on_selection_change()

    def _clear_selection(self):
        for lst in (self.indep_list, self.dep_list):
            if lst['state'] != 'disabled':
                lst.selection_clear(0, tk.END)
        self._on_selection_change()

    def _on_selection_change(self, _event=None):
        self.heatmap_rendered = False
        self.save_btn.configure(state='disabled')
        self.copy_btn.configure(state='disabled')
        self.status_var.set('Adjust column selection and render the heatmap.')

    def _populate_columns(self):
        df = self.df_sheets.get(self.current_sheet)
        indep = [c for c in INDEPENDENTS if df is not None and c in df.columns]
        dep = [c for c in DEPENDENTS if df is not None and c in df.columns]

        def fill(listbox, columns):
            if not columns:
                listbox.delete(0, tk.END)
                listbox.configure(state='disabled')
                return False
            listbox.configure(state='normal')
            listbox.delete(0, tk.END)
            for col in columns:
                listbox.insert(tk.END, col)
            listbox.selection_set(0, tk.END)
            return True

        has_indep = fill(self.indep_list, indep)
        has_dep = fill(self.dep_list, dep)

        if not (has_indep or has_dep):
            self.select_all_btn.configure(state='disabled')
            self.clear_btn.configure(state='disabled')
            self._on_selection_change()
            return

        self.select_all_btn.configure(state='normal')
        self.clear_btn.configure(state='normal')
        self._on_selection_change()

    def _get_selected_columns(self):
        cols = []
        for lst in (self.indep_list, self.dep_list):
            if lst['state'] == 'disabled':
                continue
            selection = lst.curselection()
            if selection:
                for idx in selection:
                    value = lst.get(idx)
                    if value not in cols:
                        cols.append(value)
        if not cols:
            for lst in (self.indep_list, self.dep_list):
                if lst['state'] == 'disabled':
                    continue
                for idx in range(lst.size()):
                    value = lst.get(idx)
                    if value not in cols:
                        cols.append(value)
        return cols

    # ---------- Heatmap rendering ----------

    def _render_heatmap(self):
        if self.current_sheet is None:
            messagebox.showwarning('No sheet', 'Load a workbook and choose a sheet first.')
            return
        df = self.df_sheets.get(self.current_sheet)
        if df is None:
            messagebox.showwarning('No data', 'The selected sheet has no usable data.')
            return

        cols = self._get_selected_columns()
        available = [c for c in cols if c in df.columns]
        if len(available) < 2:
            messagebox.showwarning('Not enough columns', 'Select at least two columns for the heatmap.')
            return

        data = df[available].apply(pd.to_numeric, errors='coerce')
        data = data.dropna(how='all')
        data = data.loc[:, data.nunique(dropna=True) > 1]
        if data.shape[1] < 2:
            messagebox.showwarning('Not enough columns', 'At least two non-constant columns are required for the heatmap.')
            return

        method = self.method_var.get()
        try:
            corr = self._compute_correlation_matrix(data, method)
        except Exception as exc:
            messagebox.showerror('Correlation error', str(exc))
            return

        if corr.isna().all().all():
            messagebox.showwarning('Invalid correlation', 'Unable to compute the selected correlation for the current selection.')
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
        self.ax.set_title(f"{method} correlation heatmap")

        for i, row_label in enumerate(labels):
            for j, col_label in enumerate(labels):
                value = corr.iat[i, j]
                display = '-' if pd.isna(value) else f"{value:.2f}"
                self.ax.text(j, i, display, ha='center', va='center', color='black', fontsize=9)

        self.fig.tight_layout()
        try:
            self.heatmap_colorbar = self.fig.colorbar(im, ax=self.ax, fraction=0.046, pad=0.04)
            self.heatmap_colorbar.set_label('Correlation value')
        except Exception:
            pass

        self.canvas.draw_idle()
        preview = ', '.join(labels[:6])
        if len(labels) > 6:
            preview += ', ...'
        self.status_var.set(f"Sheet: {self.current_sheet} | {method} columns ({len(labels)}): {preview}")
        self.heatmap_rendered = True
        self.save_btn.configure(state='normal')
        self.copy_btn.configure(state='normal')

    def _compute_correlation_matrix(self, data: pd.DataFrame, method: str) -> pd.DataFrame:
        method_lower = method.lower()
        if method_lower.startswith('pearson'):
            return data.corr(method='pearson')
        if method_lower.startswith('spearman'):
            return data.corr(method='spearman')
        if 'dcor' in method_lower:
            labels = list(data.columns)
            result = pd.DataFrame(np.eye(len(labels)), index=labels, columns=labels, dtype=float)
            for i in range(len(labels)):
                for j in range(i + 1, len(labels)):
                    val = self._distance_correlation(data.iloc[:, i], data.iloc[:, j])
                    result.iat[i, j] = result.iat[j, i] = val
            return result
        raise ValueError(f"Unknown correlation method: {method}")

    @staticmethod
    def _distance_correlation(series_a: pd.Series, series_b: pd.Series) -> float:
        mask = series_a.notna() & series_b.notna()
        x = series_a[mask].to_numpy(dtype=float)
        y = series_b[mask].to_numpy(dtype=float)
        if len(x) < 2 or len(y) < 2:
            return float('nan')

        def _centered_distance_matrix(vec):
            vec = vec.reshape(-1, 1)
            dist = np.abs(vec - vec.T)
            row_mean = dist.mean(axis=1, keepdims=True)
            col_mean = dist.mean(axis=0, keepdims=True)
            grand_mean = dist.mean()
            return dist - row_mean - col_mean + grand_mean

        A = _centered_distance_matrix(x)
        B = _centered_distance_matrix(y)
        dcov = np.sqrt(np.mean(A * B))
        dvar_x = np.sqrt(np.mean(A * A))
        dvar_y = np.sqrt(np.mean(B * B))
        if dvar_x <= 1e-12 or dvar_y <= 1e-12:
            return float('nan')
        return dcov / np.sqrt(dvar_x * dvar_y)

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

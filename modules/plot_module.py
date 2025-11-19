import os
import json
import math
import io
import datetime as _dt
from dataclasses import dataclass

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import numpy as np
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from .plot_shared import (
    OKABE_ITO,
    RHO_FOAM_G,
    RHO_FOAM_KG,
    RHO_REL,
    DESV_RHO_FOAM_G,
    DESV_RHO_FOAM_KG,
    INDEPENDENTS,
    INDEPENDENT_TO_COLUMN,
    INDEPENDENT_COLUMNS,
    INDEPENDENT_LATEX,
    DEPENDENT_LABELS,
    DEPENDENT_MAP,
    DEPENDENT_COLUMN_TO_LABEL,
    DEPENDENT_COLUMNS,
    DEVIATIONS,
    LEGACY_DEPENDENT_LABELS,
    dependent_latex,
    independent_latex,
    friendly_column_name,
    augment_density_columns,
)


# Ensure TkAgg backend for embedding
matplotlib.use("TkAgg")






def _is_number_series(s: pd.Series) -> bool:
    try:
        pd.to_numeric(s.dropna(), errors="raise")
        return True
    except Exception:
        return False


def _as_float_array(s: pd.Series) -> np.ndarray:
    try:
        return pd.to_numeric(s, errors="coerce").to_numpy()
    except Exception:
        return s.to_numpy()


def _natural_sort_key(val):
    # Try numeric first
    try:
        return float(val)
    except Exception:
        pass
    # Extract first number within string for ordering like "10 g"
    if isinstance(val, str):
        import re
        m = re.search(r"[-+]?[0-9]*\.?[0-9]+", val.replace(",", "."))
        if m:
            try:
                return float(m.group(0))
            except Exception:
                pass
    # Fallback to string representation
    return str(val)


class _Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        self.widget.bind("<Enter>", self._show)
        self.widget.bind("<Leave>", self._hide)

    def _show(self, _event=None):
        if self.tip is not None:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(self.tip, text=self.text, background="lightyellow", relief="solid", borderwidth=1)
        label.pack(ipadx=4, ipady=2)

    def _hide(self, _event=None):
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None


@dataclass
class Constraint:
    exact: str = ""
    min_val: str = ""
    max_val: str = ""

    def to_filter(self, series: pd.Series):
        # Returns boolean mask for applying on DataFrame
        if self.exact.strip() != "":
            # Try numeric comparison first; fallback to string equality
            s_numeric = pd.to_numeric(series, errors="coerce")
            try:
                v = float(self.exact)
                return (s_numeric - v).abs() < 1e-12
            except Exception:
                return series.astype(str) == str(self.exact)
        else:
            lo = self.min_val.strip()
            hi = self.max_val.strip()
            if lo == "" and hi == "":
                # No constraint provided
                return pd.Series([True] * len(series), index=series.index)
            s_numeric = pd.to_numeric(series, errors="coerce")
            mask = pd.Series([True] * len(series), index=series.index)
            if lo != "":
                try:
                    v = float(lo)
                    mask &= s_numeric >= v
                except Exception:
                    # Non-numeric: no-op lower bound
                    pass
            if hi != "":
                try:
                    v = float(hi)
                    mask &= s_numeric <= v
                except Exception:
                    # Non-numeric: no-op upper bound
                    pass
            return mask


class PlotModule:
    """Interactive publication-quality scatter plotter for All_Results_* Excel files.

    - Validates exact headers per CombineModule
    - Enforces constancy rule (all independents not X nor Group must be constrained)
    - Optional error bars when deviation columns exist
    - Color-blind–friendly defaults; monochrome preview option
    - Export figure (PNG/TIFF/SVG) and filtered data (CSV + JSON settings)
    """

    def __init__(self, parent, settings_manager=None, default_all_results_glob: str | None = None):
        self.root = tk.Toplevel(parent) if isinstance(parent, tk.Tk) else parent
        self.root.title("Publication Plots (Scatter)")
        self.settings = settings_manager
        self.last_state_cache = {}
        self._suspend_state_events = False
        if self.settings is not None:
            cache = self.settings.get('plot_last_state', {})
            if isinstance(cache, dict):
                self.last_state_cache = dict(cache)

        # Data
        self.df_all = pd.DataFrame()
        self.df_filtered = pd.DataFrame()
        self.df_sheets = {}
        self.active_sheet_name = None
        self._sheet_frames = {}
        self._sheet_labels = {}
        self.sheet_notebook = None

        # UI variables
        self.file_var = tk.StringVar(value=default_all_results_glob or "")
        self.x_var = tk.StringVar()
        self.y_var = tk.StringVar()
        self.group_var = tk.StringVar(value="<None>")
        self.errorbar_var = tk.BooleanVar(value=False)
        self.mono_var = tk.BooleanVar(value=False)
        self.connect_lines_var = tk.BooleanVar(value=True)
        self.dpi_var = tk.IntVar(value=600)
        self.xmin_var = tk.StringVar()
        self.xmax_var = tk.StringVar()
        self.ymin_var = tk.StringVar()
        self.ymax_var = tk.StringVar()
        self.hline_enabled_var = tk.BooleanVar(value=False)
        self.hline_value_var = tk.StringVar()
        self.vline_enabled_var = tk.BooleanVar(value=False)
        self.vline_value_var = tk.StringVar()

        # Constraint variables per independent
        self.constraints = {name: Constraint() for name in INDEPENDENTS}

        self.hline_value_var.trace_add("write", self._on_reference_line_value_change)
        self.vline_value_var.trace_add("write", self._on_reference_line_value_change)

        self._build_ui()
        self._apply_default_fonts()

    # ---------- UI ----------
    def _build_ui(self):
        container = ttk.Frame(self.root, padding=10)
        container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # File select row
        file_row = ttk.Frame(container)
        file_row.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        file_row.columnconfigure(1, weight=1)
        ttk.Label(file_row, text="All_Results Excel:").grid(row=0, column=0, sticky=tk.W, padx=(0, 6))
        entry = ttk.Entry(file_row, textvariable=self.file_var)
        entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        ttk.Button(file_row, text="Browse", command=self._browse_file).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(file_row, text="Load", command=self._load_file).grid(row=0, column=3, padx=(6, 0))

        # Sheet tabs (populated after loading)
        self.sheet_notebook = ttk.Notebook(container)
        self.sheet_notebook.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.sheet_notebook.bind("<<NotebookTabChanged>>", self._on_sheet_tab_changed)
        self.sheet_notebook.grid_remove()

        # Selections row
        sel = ttk.Frame(container)
        sel.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        for i in range(8):
            sel.columnconfigure(i, weight=1)
        ttk.Label(sel, text="Y:").grid(row=0, column=0, sticky=tk.W)
        self.y_combo = ttk.Combobox(sel, textvariable=self.y_var, values=DEPENDENT_LABELS, state="readonly")
        self.y_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=6)
        self.y_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_y_change())

        ttk.Label(sel, text="X:").grid(row=0, column=2, sticky=tk.W)
        self.x_combo = ttk.Combobox(sel, textvariable=self.x_var, values=INDEPENDENTS, state="readonly")
        self.x_combo.grid(row=0, column=3, sticky=(tk.W, tk.E), padx=6)
        self.x_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_axes_change())

        ttk.Label(sel, text="Group:").grid(row=0, column=4, sticky=tk.W)
        self.group_combo = ttk.Combobox(sel, textvariable=self.group_var, values=["<None>"] + INDEPENDENTS, state="readonly")
        self.group_combo.grid(row=0, column=5, sticky=(tk.W, tk.E), padx=6)
        self.group_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_group_change())

        self.err_chk = ttk.Checkbutton(sel, text="Error bars", variable=self.errorbar_var, command=self._on_option_change)
        self.err_chk.grid(row=0, column=6, sticky=tk.W)
        _Tooltip(
            self.err_chk,
            "Available only for Y in {Ø (µm), Nᵥ (cells·cm³), ρf (g/cm³), ρf (kg/m³)} when the matching deviation column is present."
        )

        self.mono_chk = ttk.Checkbutton(sel, text="Monochrome preview", variable=self.mono_var, command=self._on_option_change)
        self.mono_chk.grid(row=0, column=7, sticky=tk.W)

        # Optional reference lines controls
        self.hline_chk = ttk.Checkbutton(
            sel,
            text="Horizontal line",
            variable=self.hline_enabled_var,
            command=self._on_reference_line_toggle,
        )
        self.hline_chk.grid(row=1, column=0, sticky=tk.W, pady=(6, 0))
        ttk.Label(sel, text="Y =").grid(row=1, column=1, sticky=tk.E, pady=(6, 0))
        self.hline_entry = ttk.Entry(sel, textvariable=self.hline_value_var, width=12, state="disabled")
        self.hline_entry.grid(row=1, column=2, sticky=(tk.W, tk.E), padx=(4, 12), pady=(6, 0))
        _Tooltip(self.hline_entry, "Draws a red horizontal guide across the plot at the specified Y value.")

        self.vline_chk = ttk.Checkbutton(
            sel,
            text="Vertical line",
            variable=self.vline_enabled_var,
            command=self._on_reference_line_toggle,
        )
        self.vline_chk.grid(row=1, column=4, sticky=tk.W, pady=(6, 0))
        ttk.Label(sel, text="X =").grid(row=1, column=5, sticky=tk.E, pady=(6, 0))
        self.vline_entry = ttk.Entry(sel, textvariable=self.vline_value_var, width=12, state="disabled")
        self.vline_entry.grid(row=1, column=6, sticky=(tk.W, tk.E), padx=(4, 0), pady=(6, 0))
        _Tooltip(self.vline_entry, "Draws a red vertical guide across the plot at the specified X value.")

        # Second-row display option: connect or not connect points
        self.connect_chk = ttk.Checkbutton(
            sel,
            text="Connect points with lines",
            variable=self.connect_lines_var,
            command=self._on_option_change,
        )
        self.connect_chk.grid(row=1, column=7, sticky=tk.W, pady=(6, 0))

        # Constraints frame
        const_frame = ttk.LabelFrame(container, text="Fixed Independents (Constancy rule)", padding=10)
        const_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        const_frame.columnconfigure(1, weight=1)
        headers = ["Variable", "Exact value (pick one)"]
        for j, h in enumerate(headers):
            ttk.Label(const_frame, text=h, foreground="gray").grid(row=0, column=j, sticky=tk.W, padx=4)

        # combobox per independent except PDR
        self.constraint_rows = {}
        row_idx = 1
        for name in INDEPENDENTS:
            if name == "PDR (MPa/s)":
                continue
            ttk.Label(const_frame, text=name).grid(row=row_idx, column=0, sticky=tk.W, padx=4, pady=2)
            cb = ttk.Combobox(const_frame, state="disabled")
            cb.bind("<<ComboboxSelected>>", lambda _e, var=name: self._on_constraint_change(var))
            cb.grid(row=row_idx, column=1, sticky=(tk.W, tk.E), padx=4)
            self.constraint_rows[name] = cb
            row_idx += 1

        # Actions row
        actions = ttk.Frame(container)
        actions.grid(row=4, column=0, sticky=(tk.W, tk.E))
        ttk.Button(actions, text="Render Plot", command=self._render_plot).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(actions, text="Save Figure", command=self._save_figure).grid(row=0, column=1, padx=6)
        ttk.Button(actions, text="Copy Figure", command=self._copy_figure).grid(row=0, column=2, padx=6)
        ttk.Button(actions, text="Export Data", command=self._export_data).grid(row=0, column=3, padx=6)
        ttk.Label(actions, text="DPI: 600 (fixed)").grid(row=0, column=4, padx=(12, 2))

        # Optional axis limits
        ttk.Label(actions, text="X min:").grid(row=1, column=0, sticky=tk.E, pady=(4, 0))
        ttk.Entry(actions, textvariable=self.xmin_var, width=10).grid(row=1, column=1, sticky=tk.W, pady=(4, 0))
        ttk.Label(actions, text="X max:").grid(row=1, column=2, sticky=tk.E, pady=(4, 0))
        ttk.Entry(actions, textvariable=self.xmax_var, width=10).grid(row=1, column=3, sticky=tk.W, pady=(4, 0))
        ttk.Label(actions, text="Y min:").grid(row=1, column=4, sticky=tk.E, pady=(4, 0))
        ttk.Entry(actions, textvariable=self.ymin_var, width=10).grid(row=1, column=5, sticky=tk.W, pady=(4, 0))
        ttk.Label(actions, text="Y max:").grid(row=1, column=6, sticky=tk.E, pady=(4, 0))
        ttk.Entry(actions, textvariable=self.ymax_var, width=10).grid(row=1, column=7, sticky=tk.W, pady=(4, 0))

        # Canvas
        canvas_frame = ttk.Frame(container)
        canvas_frame.grid(row=5, column=0, sticky=(tk.N, tk.S, tk.W, tk.E), pady=(10, 0))
        container.rowconfigure(5, weight=1)
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        self.fig = Figure(figsize=(8.5, 8.0), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=canvas_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))

        # Annotations area (under-plot echo)
        self.info_var = tk.StringVar(value="")
        self.info_label = ttk.Label(container, textvariable=self.info_var, foreground="gray")
        self.info_label.grid(row=6, column=0, sticky=(tk.W), pady=(6, 0))
        # Disable controls until file is loaded
        self._set_controls_state("disabled")

        # Prefill last path if available
        if self.settings is not None and not self.file_var.get().strip():
            last = self.settings.get("last_output_file", "")
            if last and os.path.exists(last):
                # Suggest All_Results in same folder
                folder = os.path.dirname(last)
                self.file_var.set(folder)

    def _apply_default_fonts(self):
        matplotlib.rcParams.update({
            "font.family": "DejaVu Sans",
            "axes.titlesize": 12,
            "axes.labelsize": 12,
            "xtick.labelsize": 10,
            "ytick.labelsize": 10,
            "legend.fontsize": 10,
        })

    def _set_controls_state(self, state: str):
        # Only set the state of the primary controls here.
        # Constraint comboboxes are managed separately by _populate_constraint_options/_apply_constraint_enablement.
        for w in [self.x_combo, self.y_combo, self.group_combo, self.err_chk, self.mono_chk, self.connect_chk]:
            try:
                w.configure(state=state)
            except Exception:
                pass

    def _drop_embedded_headers(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df
        key_cols = ["Label", "m(g)", "Water (g)", "T (\u00B0C)", "Psat (MPa)", "t (min)"]
        present = [col for col in key_cols if col in df.columns]
        if not present:
            return df
        match_flags = []
        for col in present:
            token = str(col).strip().casefold()
            series = df[col].astype(str).str.strip().str.casefold()
            match_flags.append(series == token)
        if not match_flags:
            return df
        match_sum = match_flags[0].astype(int)
        for flag in match_flags[1:]:
            match_sum += flag.astype(int)
        min_matches = 3 if len(present) >= 3 else len(present)
        header_mask = match_sum >= min_matches
        if header_mask.any():
            df = df.loc[~header_mask].copy()
        return df

    # ---------- Sheet management ----------
    def _clear_sheet_tabs(self):
        if not self.sheet_notebook:
            return
        for child in list(self.sheet_notebook.winfo_children()):
            self.sheet_notebook.forget(child)
            child.destroy()
        self._sheet_frames.clear()
        self._sheet_labels.clear()
        self.sheet_notebook.grid_remove()

    @staticmethod
    def _normalize_sheet_name(name: str) -> str:
        return name.strip().lower().replace('_', '').replace(' ', '')

    def _sheet_display_name(self, sheet_name: str) -> str:
        normalized = self._normalize_sheet_name(sheet_name)
        if normalized in {'allresults', 'general'}:
            return 'General'
        return sheet_name

    def _build_sheet_tabs(self):
        if not self.sheet_notebook:
            return
        self._clear_sheet_tabs()
        if not self.df_sheets:
            return
        for sheet_name in sorted(self.df_sheets.keys(), key=lambda n: n.lower()):
            df = self.df_sheets[sheet_name]
            display = self._sheet_display_name(sheet_name)
            self._sheet_labels[sheet_name] = display
            tab_text = f"{display} ({len(df)})"
            frame = ttk.Frame(self.sheet_notebook)
            frame.sheet_name = sheet_name  # type: ignore[attr-defined]
            ttk.Label(
                frame,
                text=f"{len(df)} rows ready for plotting",
                foreground='gray',
                wraplength=280,
                justify=tk.LEFT,
            ).pack(anchor='w', padx=12, pady=8)
            self.sheet_notebook.add(frame, text=tab_text)
            self._sheet_frames[frame] = sheet_name
        self.sheet_notebook.grid()

    def _activate_sheet(self, sheet_name: str, *, reset_axes: bool, select_tab: bool):
        if sheet_name not in self.df_sheets:
            return
        if self.active_sheet_name == sheet_name and not reset_axes:
            return
        self.active_sheet_name = sheet_name
        self.df_all = self._drop_embedded_headers(augment_density_columns(self.df_sheets[sheet_name].copy()))
        self.df_filtered = pd.DataFrame()
        self.constraints = {name: Constraint() for name in INDEPENDENTS}
        self._suspend_state_events = True
        try:
            for cb in self.constraint_rows.values():
                cb.set('')
            self._populate_constraint_options()
            if reset_axes:
                self.x_var.set(INDEPENDENTS[0])
                self.y_var.set(DEPENDENT_LABELS[0])
                self.group_var.set('<None>')
                self.errorbar_var.set(False)
                self.mono_var.set(False)
                self.hline_enabled_var.set(False)
                self.hline_value_var.set("")
                self.vline_enabled_var.set(False)
                self.vline_value_var.set("")
            self._apply_stored_state(sheet_name)
        finally:
            self._suspend_state_events = False
        self._update_errorbar_state()
        self._update_reference_line_controls()
        self._apply_constraint_enablement()
        self._persist_current_state()
        self.info_var.set('')
        try:
            self.ax.clear()
        except Exception:
            pass
        try:
            self.canvas.draw_idle()
        except Exception:
            pass
        if select_tab and self.sheet_notebook:
            for frame, actual in self._sheet_frames.items():
                if actual == sheet_name:
                    try:
                        self.sheet_notebook.select(frame)
                    except Exception:
                        pass
                    break

    def _on_sheet_tab_changed(self, _event=None):
        if not self.sheet_notebook:
            return
        current = self.sheet_notebook.select()
        if not current:
            return
        frame = self.sheet_notebook.nametowidget(current)
        sheet_name = self._sheet_frames.get(frame)
        if sheet_name:
            self._activate_sheet(sheet_name, reset_axes=False, select_tab=False)

    # ---------- State persistence ----------
    def _state_storage_key(self, sheet_name: str | None) -> str | None:
        if not sheet_name:
            return None
        file_path = self.file_var.get().strip()
        if not file_path:
            return None
        try:
            normalized = os.path.abspath(file_path)
        except Exception:
            normalized = file_path
        return f"{normalized}::{sheet_name}"

    def _collect_current_state(self) -> dict:
        state = {
            'x': self.x_var.get(),
            'y': self.y_var.get(),
            'group': self.group_var.get(),
            'errorbars': bool(self.errorbar_var.get()),
            'monochrome': bool(self.mono_var.get()),
            'connect_lines': bool(self.connect_lines_var.get()),
            'dpi': int(self.dpi_var.get()),
            'hline_enabled': bool(self.hline_enabled_var.get()),
            'hline_value': self.hline_value_var.get(),
            'vline_enabled': bool(self.vline_enabled_var.get()),
            'vline_value': self.vline_value_var.get(),
            'x_min': self.xmin_var.get(),
            'x_max': self.xmax_var.get(),
            'y_min': self.ymin_var.get(),
            'y_max': self.ymax_var.get(),
            'constraints': {},
        }
        for name, cb in self.constraint_rows.items():
            state['constraints'][name] = cb.get()
        return state

    def _apply_stored_state(self, sheet_name: str):
        key = self._state_storage_key(sheet_name)
        if not key:
            return
        state = self.last_state_cache.get(key)
        if not isinstance(state, dict):
            return
        self._suspend_state_events = True
        try:
            x_val = state.get('x')
            if x_val and x_val in INDEPENDENTS:
                self.x_var.set(x_val)
            y_val = state.get('y')
            if y_val:
                if y_val in DEPENDENT_LABELS:
                    self.y_var.set(y_val)
                elif y_val in DEPENDENT_COLUMN_TO_LABEL:
                    self.y_var.set(DEPENDENT_COLUMN_TO_LABEL[y_val])
            group_val = state.get('group')
            valid_groups = ["<None>"] + INDEPENDENTS
            if group_val in valid_groups:
                self.group_var.set(group_val)
            self.errorbar_var.set(bool(state.get('errorbars')))
            self.mono_var.set(bool(state.get('monochrome')))
            self.connect_lines_var.set(bool(state.get('connect_lines', True)))
            dpi_val = state.get('dpi')
            # Force 600 dpi for publication consistency; ignore older 300 dpi states.
            self.dpi_var.set(600)
            self.xmin_var.set(state.get('x_min', ''))
            self.xmax_var.set(state.get('x_max', ''))
            self.ymin_var.set(state.get('y_min', ''))
            self.ymax_var.set(state.get('y_max', ''))
            self.hline_enabled_var.set(bool(state.get('hline_enabled')))
            self.hline_value_var.set(state.get('hline_value', ''))
            self.vline_enabled_var.set(bool(state.get('vline_enabled')))
            self.vline_value_var.set(state.get('vline_value', ''))
            constraints = state.get('constraints', {})
            if isinstance(constraints, dict):
                for name, cb in self.constraint_rows.items():
                    val = constraints.get(name)
                    if val and isinstance(val, str):
                        values = cb.cget('values')
                        if val in values:
                            cb.set(val)
        finally:
            self._suspend_state_events = False

    def _persist_current_state(self):
        if self._suspend_state_events:
            return
        key = self._state_storage_key(self.active_sheet_name)
        if not key:
            return
        state = self._collect_current_state()
        self.last_state_cache[key] = state
        if self.settings is not None:
            self.settings.set('plot_last_state', dict(self.last_state_cache))

    def _on_option_change(self):
        self._persist_current_state()

    def _on_constraint_change(self, _name):
        if not self._suspend_state_events:
            self._persist_current_state()

    # ---------- Events ----------
    def _browse_file(self):
        initialdir = None
        preset = self.file_var.get()
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
            cand = sorted(glob.glob(os.path.join(path, "All_Results_*.xlsx")))
            if not cand:
                messagebox.showwarning("Not found", "No All_Results_*.xlsx in the selected folder.")
                return
            path = cand[-1]
            self.file_var.set(path)
        elif ("*" in path) or ("?" in path):
            import glob
            matches = sorted(glob.glob(path))
            if not matches:
                messagebox.showwarning("Not found", f"No files match pattern:\n{path}")
                return
            path = matches[-1]
            self.file_var.set(path)
        if not os.path.exists(path):
            messagebox.showerror("File not found", path)
            return

        try:
            sheets_raw = pd.read_excel(path, sheet_name=None, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Read error", f"Failed to read workbook: {e}")
            return

        required = set(INDEPENDENT_COLUMNS + DEPENDENT_COLUMNS)
        required.discard("Psat (MPa)")
        valid_sheets = {}
        for sheet_name, df in sheets_raw.items():
            if not isinstance(df, pd.DataFrame):
                continue
            if 'Water (g)' not in df.columns and 'Water' in df.columns:
                df = df.rename(columns={'Water': 'Water (g)'})
            df = self._drop_embedded_headers(augment_density_columns(df))
            for legacy, modern in LEGACY_DEPENDENT_LABELS.items():
                if legacy in df.columns and modern not in df.columns:
                    df = df.rename(columns={legacy: modern})
            missing = [c for c in required if c not in df.columns]
            if missing:
                for col in missing:
                    df[col] = pd.NA
            valid_sheets[sheet_name] = df.copy()

        if not valid_sheets:
            expected = "\n- ".join(sorted(friendly_column_name(col) for col in required))
            messagebox.showerror(
                "Invalid workbook",
                f"No sheet contains the required columns. Expected:\n- {expected}",
            )
            self._set_controls_state("disabled")
            self._clear_sheet_tabs()
            return

        self.df_sheets = valid_sheets

        default_sheet = None
        for name in valid_sheets.keys():
            if self._normalize_sheet_name(name) == "allresults":
                default_sheet = name
                break
        if default_sheet is None:
            default_sheet = next(iter(sorted(valid_sheets.keys(), key=lambda n: n.lower())))

        self.x_combo.configure(values=INDEPENDENTS)
        self.y_combo.configure(values=DEPENDENT_LABELS)
        self.group_combo.configure(values=["<None>"] + INDEPENDENTS)

        self._build_sheet_tabs()
        self._activate_sheet(default_sheet, reset_axes=True, select_tab=True)
        self._set_controls_state("readonly")

        messagebox.showinfo(
            "Loaded",
            f"Loaded {len(valid_sheets)} sheet{'s' if len(valid_sheets) != 1 else ''} from '{os.path.basename(path)}'.",
        )
    def _on_axes_change(self):
        # Ensure group is not same as X
        cur_g = self.group_var.get()
        if cur_g == self.x_var.get():
            self.group_var.set("<None>")
        self._update_errorbar_state()
        self._apply_constraint_enablement()
        self._persist_current_state()

    def _on_group_change(self):
        # No specific action besides errorbar state
        self._update_errorbar_state()
        self._apply_constraint_enablement()
        self._persist_current_state()

    def _on_y_change(self):
        self._update_errorbar_state()
        self._persist_current_state()

    def _update_errorbar_state(self):
        y_label = self.y_var.get().strip()
        yerr_col = DEVIATIONS.get(y_label)
        can_err = bool(yerr_col and yerr_col in getattr(self.df_all, "columns", []))
        state = "normal" if can_err else "disabled"
        self.err_chk.configure(state=state)
        if not can_err:
            self.errorbar_var.set(False)
            _Tooltip(self.err_chk, "Enable only for Y in {\u00F8, N\u1D65, \u03C1f (g/cm^3), \u03C1f (kg/m^3)} with deviation column present.")

    def _on_reference_line_toggle(self):
        self._update_reference_line_controls()
        if not self._suspend_state_events:
            self._persist_current_state()

    def _on_reference_line_value_change(self, *_):
        if not self._suspend_state_events:
            self._persist_current_state()

    def _update_reference_line_controls(self):
        state_h = "normal" if self.hline_enabled_var.get() else "disabled"
        state_v = "normal" if self.vline_enabled_var.get() else "disabled"
        if hasattr(self, "hline_entry"):
            self.hline_entry.configure(state=state_h)
        if hasattr(self, "vline_entry"):
            self.vline_entry.configure(state=state_v)

    def _populate_constraint_options(self):
        # Fill combobox options with unique values from data for each independent (except PDR)
        if self.df_all.empty:
            return
        for name, cb in self.constraint_rows.items():
            column = INDEPENDENT_TO_COLUMN.get(name, name)
            if column not in self.df_all.columns:
                cb.configure(state="disabled", values=[])
                continue
            vals = self.df_all[column].dropna().unique().tolist()
            try:
                vals_sorted = sorted(vals, key=_natural_sort_key)
            except Exception:
                vals_sorted = vals
            # Ensure values are strings for Tk
            cb_values = [str(v) for v in vals_sorted]
            cb.configure(state="readonly", values=cb_values)
            cb.set("")
        self._apply_constraint_enablement()

    def _apply_constraint_enablement(self):
        # Disable combobox for X and Group variables; others enabled (once values loaded)
        x = self.x_var.get()
        g = None if self.group_var.get() == "<None>" else self.group_var.get()
        for name, cb in self.constraint_rows.items():
            if name == x or (g and name == g):
                cb.configure(state="disabled")
            else:
                vals = cb.cget("values")
                has_vals = bool(vals) and len(vals) > 0
                if cb.cget("state") == "disabled" and has_vals:
                    cb.configure(state="readonly")

    # ---------- Core logic ----------
    def _collect_constraints(self):
        res = {}
        for name, cb in self.constraint_rows.items():
            val = (cb.get() or "").strip()
            res[name] = Constraint(exact=val)
        return res

    def _apply_constancy_rule(self, df: pd.DataFrame, x_display: str, group_display: str | None, constraints: dict):
        # Exclude X and optional group from constancy rule (and PDR globally)
        remaining_pairs = []
        for display in INDEPENDENTS:
            if display == "PDR (MPa/s)":
                continue
            if display == x_display:
                continue
            if group_display and display == group_display:
                continue
            column = INDEPENDENT_TO_COLUMN.get(display, display)
            if column not in df.columns:
                continue
            remaining_pairs.append((display, column))

        # Enforce that each remaining has some constraint (unless constant already)
        missing = []
        for display, column in remaining_pairs:
            c: Constraint = constraints.get(display, Constraint())
            if c.exact.strip() == "":
                if df[column].dropna().nunique() > 1:
                    missing.append(display)
        if missing:
            raise ValueError(
                "Constancy rule: select an exact value for these variables: " + ", ".join(missing)
            )

        # Build mask by ANDing all constraints provided
        mask = pd.Series([True] * len(df), index=df.index)
        for display, column in remaining_pairs:
            c: Constraint = constraints.get(display, Constraint())
            if c.exact.strip() != "":
                series = df[column]
                s_numeric = pd.to_numeric(series, errors="coerce")
                try:
                    vnum = float(c.exact)
                    mask &= (s_numeric - vnum).abs() < 1e-12
                except Exception:
                    mask &= series.astype(str) == c.exact
        return df[mask].copy(), [display for display, _ in remaining_pairs]

    def _ensure_n_requirements(self, df: pd.DataFrame, group_col: str | None):
        total = int(len(df))
        if total < 2:
            detail = ""
            if group_col and total > 0 and group_col in df.columns:
                counts = df[group_col].value_counts(dropna=False).to_dict()
                detail = f" (counts by {group_col}: {counts})"
            raise ValueError(
                f"Insufficient data: at least 2 points required after filtering (n={total})." + detail
            )
        if group_col:
            bad = []
            counts = {}
            for g, gdf in df.groupby(group_col):
                counts[str(g)] = len(gdf)
                if len(gdf) < 2:
                    bad.append(f"{g} (n={len(gdf)})")
            if bad:
                message = "Groups with < 2 points: " + ", ".join(bad)
                if counts:
                    message += f". Counts: {counts}"
                raise ValueError(message)

    def _prepare_plot_data(self, df: pd.DataFrame, x_name: str, y_name: str, group_name: str | None):
        if group_name:
            # Sort groups by natural numeric order
            keys = sorted(df[group_name].dropna().unique().tolist(), key=_natural_sort_key)
            groups = [(k, df[df[group_name] == k].sort_values(by=x_name)) for k in keys]
        else:
            groups = [(None, df.sort_values(by=x_name))]
        return groups

    def _style_axes(self, ax):
        ax.set_facecolor("white")
        ax.grid(False)
        # Keep only left and bottom spines
        for spine in ["top", "right"]:
            ax.spines[spine].set_visible(False)
        for spine in ["left", "bottom"]:
            ax.spines[spine].set_linewidth(0.8)

    def _group_styles(self, n_groups: int, monochrome: bool):
        markers = ["o", "s", "^", "D", "v", "P", "X", "+", "*", "h"]
        linestyles = ["-", "--", "-.", ":"]
        styles = []
        for i in range(n_groups):
            color = "#000000" if monochrome else OKABE_ITO[i % len(OKABE_ITO)]
            marker = markers[i % len(markers)]
            # After palette cycles, also vary line style to ensure 2 channels differ
            style = linestyles[(i // len(OKABE_ITO)) % len(linestyles)] if not monochrome else linestyles[i % len(linestyles)]
            styles.append((color, marker, style))
        return styles

    def _maybe_jitter(self, x_vals: np.ndarray):
        # Apply small jitter if many overlaps in X
        if x_vals.size == 0:
            return x_vals
        rng = np.nanmax(x_vals) - np.nanmin(x_vals)
        if not np.isfinite(rng) or rng == 0:
            return x_vals
        # Detect duplicates
        _, counts = np.unique(np.round(x_vals, 12), return_counts=True)
        if (counts > 1).any():
            amp = 0.03 * rng  # 3% of range (peak-to-peak)
            noise = (np.random.rand(len(x_vals)) - 0.5) * amp
            return x_vals + noise
        return x_vals

    @staticmethod
    def _optional_float(value):
        if value is None:
            return None
        text = str(value).strip().replace(",", ".")
        if text == "":
            return None
        try:
            return float(text)
        except ValueError:
            return None

    def _render_plot(self):
        if self.df_all.empty:
            messagebox.showwarning("No data", "Load an All_Results Excel first.")
            return

        x_display = self.x_var.get().strip()
        x_column = INDEPENDENT_TO_COLUMN.get(x_display, x_display)
        y_label = self.y_var.get().strip()
        y_column = DEPENDENT_MAP.get(y_label, y_label)
        grp = self.group_var.get().strip()
        group_display = None if grp == "<None>" else grp
        group_column = INDEPENDENT_TO_COLUMN.get(group_display, group_display) if group_display else None

        if x_column not in self.df_all.columns:
            messagebox.showerror("Missing column", f"Column '{x_column}' not found in the selected sheet.")
            return
        if group_column and (group_column not in self.df_all.columns):
            messagebox.showerror("Missing column", f"Grouping column '{group_column}' not found in the selected sheet.")
            return

        hline_value = None
        if self.hline_enabled_var.get():
            raw = self.hline_value_var.get()
            if not raw.strip():
                messagebox.showerror("Horizontal line", "Enter a Y value for the horizontal guide or disable the option.")
                return
            hline_value = self._optional_float(raw)
            if hline_value is None:
                messagebox.showerror("Horizontal line", f"'{raw}' is not a valid numeric Y value.")
                return

        vline_value = None
        if self.vline_enabled_var.get():
            raw = self.vline_value_var.get()
            if not raw.strip():
                messagebox.showerror("Vertical line", "Enter an X value for the vertical guide or disable the option.")
                return
            vline_value = self._optional_float(raw)
            if vline_value is None:
                messagebox.showerror("Vertical line", f"'{raw}' is not a valid numeric X value.")
                return

        # Collect and apply constraints
        constraints = self._collect_constraints()
        # Persist last used constraints for export summary
        self.constraints.update(constraints)
        self._persist_current_state()
        try:
            filtered, _remaining = self._apply_constancy_rule(self.df_all, x_display, group_display, constraints)
        except ValueError as e:
            messagebox.showerror("Constancy rule", str(e))
            return

        # Keep only relevant columns
        cols = [x_column, y_column]
        if group_column:
            cols.append(group_column)
        yerr_name = DEVIATIONS.get(y_label)
        if self.errorbar_var.get() and yerr_name in filtered.columns:
            cols.append(yerr_name)
        # For export, we'll keep deviation columns even if not used
        self.df_filtered = filtered.copy()

        # Guard for n
        try:
            self._ensure_n_requirements(filtered.dropna(subset=[x_column, y_column]), group_column)
        except ValueError as e:
            messagebox.showerror("Not enough points", str(e))
            return

        # Prepare plot
        self.ax.clear()
        self._style_axes(self.ax)

        groups = self._prepare_plot_data(filtered, x_column, y_column, group_column)
        styles = self._group_styles(len(groups), self.mono_var.get())
        connect = bool(self.connect_lines_var.get())

        for (gidx, (gval, gdf)) in enumerate(groups):
            color, marker, linestyle = styles[gidx]

            x = _as_float_array(gdf[x_column])
            y = _as_float_array(gdf[y_column])

            # Optional connecting lines within each group in ascending X
            if connect:
                self.ax.plot(
                    x,
                    y,
                    linestyle=linestyle,
                    linewidth=1.5,
                    color=color,
                    alpha=1.0,
                    antialiased=True,
                )

            # Scatter markers (edge black 0.6pt, size ~ 42pt^2)
            self.ax.scatter(
                x,
                y,
                s=42,
                marker=marker,
                facecolor=color,
                edgecolor="black",
                linewidths=0.6,
                alpha=0.9,
                zorder=3,
            )

            # Error bars (vertical only)
            if self.errorbar_var.get() and (yerr_name in gdf.columns):
                yerr = _as_float_array(gdf[yerr_name])
                self.ax.errorbar(
                    x,
                    y,
                    yerr=yerr,
                    fmt="none",
                    ecolor=color,
                    elinewidth=1.1,
                    capsize=3.5,
                    alpha=0.8,
                    zorder=2,
                )

        if hline_value is not None:
            self.ax.axhline(
                hline_value,
                color="red",
                linewidth=1.2,
                linestyle="--",
                alpha=0.9,
                zorder=1,
            )
        if vline_value is not None:
            self.ax.axvline(
                vline_value,
                color="red",
                linewidth=1.2,
                linestyle="--",
                alpha=0.9,
                zorder=1,
            )

        # Labels (exact headers with units)
        self.ax.set_xlabel(independent_latex(x_display))
        y_display_label = DEPENDENT_COLUMN_TO_LABEL.get(y_column, y_label)
        self.ax.set_ylabel(dependent_latex(y_display_label))

        # Optional fixed axis limits
        x_min = self._optional_float(self.xmin_var.get())
        x_max = self._optional_float(self.xmax_var.get())
        y_min = self._optional_float(self.ymin_var.get())
        y_max = self._optional_float(self.ymax_var.get())
        if x_min is not None or x_max is not None:
            cur = self.ax.get_xlim()
            self.ax.set_xlim(x_min if x_min is not None else cur[0], x_max if x_max is not None else cur[1])
        if y_min is not None or y_max is not None:
            cur = self.ax.get_ylim()
            self.ax.set_ylim(y_min if y_min is not None else cur[0], y_max if y_max is not None else cur[1])

        # Legend
        if group_display:
            handles, labels = [], []
            for (gidx, (gval, _)) in enumerate(groups):
                color, marker, linestyle = styles[gidx]
                legend_linestyle = linestyle if connect else ""
                legend_linewidth = 1.5 if connect else 0.0
                h = matplotlib.lines.Line2D(
                    [0],
                    [0],
                    color=color,
                    marker=marker,
                    linestyle=legend_linestyle,
                    linewidth=legend_linewidth,
                    markerfacecolor=color,
                    markeredgecolor="black",
                    markeredgewidth=0.6,
                )
                handles.append(h)
                group_label = independent_latex(group_display)
                labels.append(f"{group_label} = {gval}")
            leg = self.ax.legend(
                handles, labels,
                loc="upper left",
                bbox_to_anchor=(1.02, 1.0),
                frameon=False,
                handlelength=2.0,
                borderaxespad=0.0,
            )
        else:
            # No legend when not grouped
            pass

        # Tight layout with room on right for legend
        self.fig.subplots_adjust(left=0.12, right=0.78, bottom=0.12, top=0.95)

        # Under-plot annotations: n per group and fixed variables
        n_info = self._n_info_text(groups)
        fixed_info = self._fixed_info_text(group_display, constraints)
        sheet_label = self._sheet_labels.get(self.active_sheet_name, self.active_sheet_name or "<no sheet>")
        self.info_var.set(f"Sheet: {sheet_label}    |    {n_info}    |    Fixed: {fixed_info}")

        self.canvas.draw_idle()

    def _n_info_text(self, groups):
        if len(groups) == 1 and groups[0][0] is None:
            return f"n: {len(groups[0][1])}"
        counts = [len(gdf) for _, gdf in groups]
        return "n per group: " + "/".join(str(n) for n in counts)

    def _fixed_info_text(self, group_name, constraints: dict):
        parts = []
        for v in INDEPENDENTS:
            if v == "PDR (MPa/s)":
                continue
            if v == self.x_var.get() or (group_name and v == group_name):
                continue
            c: Constraint = constraints.get(v, Constraint())
            if c.exact.strip() != "":
                parts.append(f"{v}={c.exact}")
            else:
                parts.append(f"{v}=<unspecified>")
        return "; ".join(parts) if parts else "(none)"

    # ---------- Export ----------
    def _default_filename(self, ext: str) -> str:
        x = self.x_var.get().replace(" ", "_")
        y = self.y_var.get().replace(" ", "_")
        g = self.group_var.get()
        by = f"_by_{g.replace(' ', '_')}" if g and g != "<None>" else ""
        ts = _dt.datetime.now().strftime("%Y%m%d-%H%M")
        return f"scatter_{y}_vs_{x}{by}_{ts}.{ext}"

    def _copy_figure(self):
        if self.df_all.empty:
            messagebox.showwarning("No plot", "Render a plot before copying.")
            return
        if os.name != "nt":
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
            self.fig.savefig(
                buffer,
                format='png',
                dpi=int(self.dpi_var.get()),
                facecolor='white',
            )
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
            messagebox.showinfo("Copied", "Figure copied to the clipboard.")
        except Exception as e:
            messagebox.showerror("Copy error", str(e))
        finally:
            buffer.close()
    def _save_figure(self):
        if self.df_all.empty:
            messagebox.showwarning("No plot", "Render a plot before saving.")
            return
        # Choose format via dialog filetypes; default PNG
        default_name = self._default_filename("png")
        filename = filedialog.asksaveasfilename(
            title="Save Figure",
            defaultextension=".png",
            initialfile=default_name,
            filetypes=[
                ("PNG", "*.png"),
                ("TIFF", "*.tiff;*.tif"),
                ("SVG", "*.svg"),
                ("PDF", "*.pdf"),
            ],
        )
        if not filename:
            return
        # Ensure white background
        try:
            self.fig.savefig(filename, dpi=int(self.dpi_var.get()), facecolor="white")
            messagebox.showinfo("Saved", f"Figure saved to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Save error", str(e))

    def _export_data(self):
        if self.df_filtered.empty:
            messagebox.showwarning("No data", "Render a plot first (data depends on filters).")
            return
        default_csv = self._default_filename("csv")
        csv_path = filedialog.asksaveasfilename(
            title="Export filtered data (CSV)",
            defaultextension=".csv",
            initialfile=default_csv,
            filetypes=[("CSV", "*.csv")],
        )
        if not csv_path:
            return
        try:
            # Determine if deviation column used
            y_label = self.y_var.get().strip()
            y_column = DEPENDENT_MAP.get(y_label, y_label)
            yerr_name = DEVIATIONS.get(y_label)
            x_display = self.x_var.get().strip()
            x_column = INDEPENDENT_TO_COLUMN.get(x_display, x_display)
            used_cols = [x_column, y_column]
            group_display = None if self.group_var.get() == "<None>" else self.group_var.get().strip()
            group_column = INDEPENDENT_TO_COLUMN.get(group_display, group_display) if group_display else None
            if group_column:
                used_cols.append(group_column)
            # Always include deviation columns if present
            if yerr_name and (yerr_name in self.df_filtered.columns):
                used_cols.append(yerr_name)
            # Also include all constraints columns for reproducibility
            for v in INDEPENDENTS:
                column = INDEPENDENT_TO_COLUMN.get(v, v)
                if column not in used_cols:
                    used_cols.append(column)
            used_cols = [c for c in used_cols if c in self.df_filtered.columns]
            self.df_filtered[used_cols].to_csv(csv_path, index=False)

            # Sidecar JSON settings
            settings = {
                "file": self.file_var.get(),
                "x": x_display,
                "y": y_label,
                "group": group_display,
                "error_bars": bool(self.errorbar_var.get()),
                "yerr_column": yerr_name if yerr_name in self.df_filtered.columns else None,
                "dpi": int(self.dpi_var.get()),
                "palette": "Okabe-Ito" if not self.mono_var.get() else "monochrome",
                "monochrome": bool(self.mono_var.get()),
                "horizontal_line": {
                    "enabled": bool(self.hline_enabled_var.get()),
                    "value": self._optional_float(self.hline_value_var.get()),
                },
                "vertical_line": {
                    "enabled": bool(self.vline_enabled_var.get()),
                    "value": self._optional_float(self.vline_value_var.get()),
                },
                "axis_limits": {
                    "x_min": self._optional_float(self.xmin_var.get()),
                    "x_max": self._optional_float(self.xmax_var.get()),
                    "y_min": self._optional_float(self.ymin_var.get()),
                    "y_max": self._optional_float(self.ymax_var.get()),
                },
                "constraints": {
                    v: {"exact": self.constraints[v].exact, "min": self.constraints[v].min_val, "max": self.constraints[v].max_val}
                    for v in INDEPENDENTS
                },
                "n_total": int(len(self.df_filtered)),
            }
            json_path = os.path.splitext(csv_path)[0] + ".json"
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)

            messagebox.showinfo("Exported", f"CSV saved to:\n{csv_path}\nSettings saved to:\n{json_path}")
        except Exception as e:
            messagebox.showerror("Export error", str(e))


# Helper to show the module as a standalone window (optional)
def show(parent=None, settings_manager=None):
    root = parent if parent is not None else tk.Tk()
    PlotModule(root, settings_manager)
    if parent is None:
        root.mainloop()

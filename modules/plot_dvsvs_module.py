import os
import json
import datetime as _dt

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import math

import pandas as pd
import numpy as np
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib import lines as mlines, colors as mcolors

from .plot_shared import (
    OKABE_ITO,
    INDEPENDENTS,
    INDEPENDENT_TO_COLUMN,
    INDEPENDENT_COLUMNS,
    INDEPENDENT_LATEX,
    DEPENDENT_LABELS,
    DEPENDENT_MAP,
    DEPENDENT_COLUMN_TO_LABEL,
    DEPENDENT_COLUMNS,
    DEPENDENT_TO_DEVIATION,
    DEVIATIONS,
    LEGACY_DEPENDENT_LABELS,
    friendly_column_name,
    dependent_latex,
    augment_density_columns,
)

matplotlib.use("TkAgg")

COLOR_BINS = 5
SHAPE_BINS = 5
SHAPE_MARKERS = ["o", "s", "^", "D", "v", "P", "X", "h", "*"]
MISSING_LABEL = "Missing"
CATEGORY_DECIMALS = {
    "Water (g)": 1,
}


def _natural_sort_key(val):
    try:
        return float(val)
    except Exception:
        pass
    if isinstance(val, str):
        import re

        m = re.search(r"[-+]?[0-9]*\.?[0-9]+", val.replace(",", "."))
        if m:
            try:
                return float(m.group(0))
            except Exception:
                pass
    return str(val)


def _as_float_array(series: pd.Series) -> np.ndarray:
    try:
        return pd.to_numeric(series, errors="coerce").to_numpy()
    except Exception:
        return series.to_numpy()


def _format_number(value: float) -> str:
    if value is None or not np.isfinite(value):
        return "-"
    magnitude = abs(value)
    if 0 < magnitude < 1e-2 or magnitude >= 1e4:
        return f"{value:.3g}"
    return f"{value:.4f}".rstrip("0").rstrip(".")


def _interval_label(left: float, right: float) -> str:
    return f"[{_format_number(left)}, {_format_number(right)}]"


def _parse_float(text: str) -> float | None:
    text = (text or "").strip()
    if text == "":
        return None
    try:
        return float(text.replace(",", "."))
    except ValueError as exc:
        raise ValueError(f"Invalid numeric value: '{text}'") from exc


def _color_palette(count: int, monochrome: bool) -> list[str]:
    if count <= 0:
        return []
    if monochrome:
        values = np.linspace(0.25, 0.8, count)
        return [mcolors.to_hex((v, v, v)) for v in values]
    base = OKABE_ITO[1:] + [OKABE_ITO[0]]
    return [base[i % len(base)] for i in range(count)]


def _format_category_value(value, decimals: int | None = None) -> str:
    if value is None:
        return MISSING_LABEL
    try:
        num = float(value)
        if not math.isfinite(num):
            return MISSING_LABEL
        if decimals is not None:
            return f"{num:.{decimals}f}"
        if abs(num - round(num)) < 1e-6:
            return str(int(round(num)))
        return f"{num:.3g}"
    except Exception:
        return str(value)


class DependentScatterModule:
    """Publication-quality dependent-vs-dependent scatter plotter for All_Results_* Excel files."""

    def __init__(self, parent, settings_manager=None, default_all_results_glob: str | None = None):
        self.root = tk.Toplevel(parent) if isinstance(parent, tk.Tk) else parent
        self.root.title("Publication Plots (Dependent vs Dependent)")
        self.settings = settings_manager
        self.last_state_cache = {}
        self._suspend_state_events = False
        if self.settings is not None:
            cache = self.settings.get("plot_dvd_last_state", {})
            if isinstance(cache, dict):
                self.last_state_cache = dict(cache)

        # Data containers
        self.df_all = pd.DataFrame()
        self.df_filtered = pd.DataFrame()
        self.df_sheets = {}
        self.active_sheet_name = None
        self._sheet_frames = {}
        self._sheet_labels = {}
        self.sheet_notebook = None

        # UI vars
        self.file_var = tk.StringVar(value=default_all_results_glob or "")
        self.x_var = tk.StringVar()
        self.y_var = tk.StringVar()
        self.color_var = tk.StringVar(value="<None>")
        self.shape_var = tk.StringVar(value="<None>")
        self.errorbar_var = tk.BooleanVar(value=False)
        self.mono_var = tk.BooleanVar(value=False)
        self.dpi_var = tk.IntVar(value=300)

        # Filters and encoding info
        self.filter_controls: dict[str, dict] = {}
        self.filter_defaults: dict[str, tuple[float | None, float | None]] = {}
        self.last_filters: dict[str, dict] = {}
        self.last_encoding_info: dict[str, dict | None] = {"color": None, "shape": None}
        self._current_legends: list = []

        self._build_ui()
        self._apply_default_fonts()

    # ---------- UI ----------
    def _build_ui(self):
        container = ttk.Frame(self.root, padding=10)
        container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # File row
        file_row = ttk.Frame(container)
        file_row.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        file_row.columnconfigure(1, weight=1)
        ttk.Label(file_row, text="All_Results Excel:").grid(row=0, column=0, sticky=tk.W, padx=(0, 6))
        entry = ttk.Entry(file_row, textvariable=self.file_var)
        entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        ttk.Button(file_row, text="Browse", command=self._browse_file).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(file_row, text="Load", command=self._load_file).grid(row=0, column=3, padx=(6, 0))

        # Sheet tabs
        self.sheet_notebook = ttk.Notebook(container)
        self.sheet_notebook.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.sheet_notebook.bind("<<NotebookTabChanged>>", self._on_sheet_tab_changed)
        self.sheet_notebook.grid_remove()

        # Axis selections
        axes_row = ttk.Frame(container)
        axes_row.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        for i in range(6):
            axes_row.columnconfigure(i, weight=1)

        ttk.Label(axes_row, text="Y:").grid(row=0, column=0, sticky=tk.W)
        self.y_combo = ttk.Combobox(axes_row, textvariable=self.y_var, values=DEPENDENT_LABELS, state="readonly")
        self.y_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=6)
        self.y_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_axis_change())

        ttk.Label(axes_row, text="X:").grid(row=0, column=2, sticky=tk.W)
        self.x_combo = ttk.Combobox(axes_row, textvariable=self.x_var, values=DEPENDENT_LABELS, state="readonly")
        self.x_combo.grid(row=0, column=3, sticky=(tk.W, tk.E), padx=6)
        self.x_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_axis_change())

        # Encodings
        enc_row = ttk.Frame(container)
        enc_row.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        for i in range(6):
            enc_row.columnconfigure(i, weight=1)

        ttk.Label(enc_row, text="Color by:").grid(row=0, column=0, sticky=tk.W)
        color_options = ["<None>"] + INDEPENDENTS
        self.color_combo = ttk.Combobox(enc_row, textvariable=self.color_var, values=color_options, state="readonly")
        self.color_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=6)
        self.color_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_encoding_change())

        ttk.Label(enc_row, text="Shape by:").grid(row=0, column=2, sticky=tk.W)
        self.shape_combo = ttk.Combobox(enc_row, textvariable=self.shape_var, values=color_options, state="readonly")
        self.shape_combo.grid(row=0, column=3, sticky=(tk.W, tk.E), padx=6)
        self.shape_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_encoding_change())

        self.err_chk = ttk.Checkbutton(enc_row, text="Show error bars", variable=self.errorbar_var, command=self._on_option_change)
        self.err_chk.grid(row=0, column=4, sticky=tk.W, padx=(6, 0))

        self.mono_chk = ttk.Checkbutton(
            enc_row, text="Monochrome preview", variable=self.mono_var, command=self._on_option_change
        )
        self.mono_chk.grid(row=0, column=5, sticky=tk.W, padx=(6, 0))

        # Filters
        self.filters_frame = ttk.LabelFrame(container, text="Filters (optional)", padding=10)
        self.filters_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.filters_frame.columnconfigure(1, weight=1)

        # Actions
        actions = ttk.Frame(container)
        actions.grid(row=5, column=0, sticky=(tk.W, tk.E))
        ttk.Button(actions, text="Render Plot", command=self._render_plot).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(actions, text="Save Figure", command=self._save_figure).grid(row=0, column=1, padx=6)
        ttk.Button(actions, text="Copy Figure", command=self._copy_figure).grid(row=0, column=2, padx=6)
        ttk.Button(actions, text="Export Data", command=self._export_data).grid(row=0, column=3, padx=6)
        ttk.Label(actions, text="DPI:").grid(row=0, column=4, padx=(12, 2))
        ttk.Radiobutton(actions, text="300", variable=self.dpi_var, value=300, command=self._on_option_change).grid(row=0, column=5)
        ttk.Radiobutton(actions, text="600", variable=self.dpi_var, value=600, command=self._on_option_change).grid(row=0, column=6)

        # Canvas
        canvas_frame = ttk.Frame(container)
        canvas_frame.grid(row=6, column=0, sticky=(tk.N, tk.S, tk.W, tk.E), pady=(10, 0))
        container.rowconfigure(6, weight=1)
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        self.fig = Figure(figsize=(8.5, 8.0), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=canvas_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))

        # Info label
        self.info_var = tk.StringVar(value="")
        self.info_label = ttk.Label(container, textvariable=self.info_var, foreground="gray")
        self.info_label.grid(row=7, column=0, sticky=tk.W, pady=(6, 0))

        self._set_controls_state("disabled")

    def _apply_default_fonts(self):
        matplotlib.rcParams.update(
            {
                "font.family": "DejaVu Sans",
                "axes.titlesize": 12,
                "axes.labelsize": 12,
                "xtick.labelsize": 10,
                "ytick.labelsize": 10,
                "legend.fontsize": 10,
            }
        )

    def _set_controls_state(self, state: str):
        for widget in [self.x_combo, self.y_combo, self.color_combo, self.shape_combo, self.err_chk, self.mono_chk]:
            try:
                widget.configure(state=state)
            except Exception:
                pass

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
        return name.strip().lower().replace("_", "").replace(" ", "")

    def _sheet_display_name(self, sheet_name: str) -> str:
        normalized = self._normalize_sheet_name(sheet_name)
        if normalized in {"allresults", "general"}:
            return "General"
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
                foreground="gray",
                wraplength=280,
                justify=tk.LEFT,
            ).pack(anchor="w", padx=12, pady=8)
            self.sheet_notebook.add(frame, text=tab_text)
            self._sheet_frames[frame] = sheet_name
        self.sheet_notebook.grid()

    def _activate_sheet(self, sheet_name: str, *, reset_axes: bool, select_tab: bool):
        if sheet_name not in self.df_sheets:
            return
        if self.active_sheet_name == sheet_name and not reset_axes:
            return

        self.active_sheet_name = sheet_name
        self.df_all = augment_density_columns(self.df_sheets[sheet_name].copy())
        self.df_filtered = pd.DataFrame()
        self.last_filters = {}

        self._suspend_state_events = True
        try:
            if reset_axes:
                default_y = DEPENDENT_LABELS[0] if DEPENDENT_LABELS else ""
                default_x = DEPENDENT_LABELS[1] if len(DEPENDENT_LABELS) > 1 else default_y
                self.y_var.set(default_y)
                self.x_var.set(default_x)
                self.color_var.set("<None>")
                self.shape_var.set("<None>")
                self.mono_var.set(False)
            self._apply_stored_state(sheet_name)
        finally:
            self._suspend_state_events = False

        self._refresh_filter_controls()
        self._persist_current_state()
        self.info_var.set("")
        try:
            self.ax.clear()
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
        return {
            "x": self.x_var.get(),
            "y": self.y_var.get(),
            "color": self.color_var.get(),
            "shape": self.shape_var.get(),
            "errorbars": bool(self.errorbar_var.get()),
            "monochrome": bool(self.mono_var.get()),
            "dpi": int(self.dpi_var.get()),
            "filters": {col: {"min": ctrl["min_var"].get(), "max": ctrl["max_var"].get()} for col, ctrl in self.filter_controls.items()},
        }

    def _apply_stored_state(self, sheet_name: str):
        key = self._state_storage_key(sheet_name)
        if not key:
            return
        state = self.last_state_cache.get(key)
        if not isinstance(state, dict):
            return

        self._suspend_state_events = True
        try:
            x_val = state.get("x")
            if x_val in DEPENDENT_LABELS:
                self.x_var.set(x_val)
            y_val = state.get("y")
            if y_val in DEPENDENT_LABELS:
                self.y_var.set(y_val)
            color_val = state.get("color")
            if color_val in (["<None>"] + INDEPENDENTS):
                self.color_var.set(color_val)
            shape_val = state.get("shape")
            if shape_val in (["<None>"] + INDEPENDENTS):
                self.shape_var.set(shape_val)
            self.errorbar_var.set(bool(state.get("errorbars")))
            self.mono_var.set(bool(state.get("monochrome")))
            dpi_val = state.get("dpi")
            if dpi_val in (300, 600):
                self.dpi_var.set(dpi_val)
        finally:
            self._suspend_state_events = False

        self._refresh_filter_controls()

        filters = state.get("filters", {})
        for column, values in filters.items():
            ctrl = self.filter_controls.get(column)
            if not ctrl:
                continue
            ctrl["min_var"].set(values.get("min", ""))
            ctrl["max_var"].set(values.get("max", ""))

    def _persist_current_state(self):
        if self._suspend_state_events:
            return
        key = self._state_storage_key(self.active_sheet_name)
        if not key:
            return
        self.last_state_cache[key] = self._collect_current_state()
        if self.settings is not None:
            self.settings.set("plot_dvd_last_state", dict(self.last_state_cache))
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

            candidates = sorted(glob.glob(os.path.join(path, "All_Results_*.xlsx")))
            if not candidates:
                messagebox.showwarning("Not found", "No All_Results_*.xlsx in the selected folder.")
                return
            path = candidates[-1]
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
            if "Water (g)" not in df.columns and "Water" in df.columns:
                df = df.rename(columns={"Water": "Water (g)"})
            df = augment_density_columns(df)
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

        options = ["<None>"] + INDEPENDENTS
        self.x_combo.configure(values=DEPENDENT_LABELS)
        self.y_combo.configure(values=DEPENDENT_LABELS)
        self.color_combo.configure(values=options)
        self.shape_combo.configure(values=options)

        self._build_sheet_tabs()
        self._activate_sheet(default_sheet, reset_axes=True, select_tab=True)
        self._set_controls_state("readonly")

        messagebox.showinfo(
            "Loaded",
            f"Loaded {len(valid_sheets)} sheet{'s' if len(valid_sheets) != 1 else ''} from '{os.path.basename(path)}'.",
        )

    def _on_axis_change(self):
        if not self._suspend_state_events:
            self._refresh_filter_controls()
            self._persist_current_state()

    def _on_encoding_change(self):
        if not self._suspend_state_events:
            self._refresh_filter_controls()
            self._persist_current_state()

    def _on_option_change(self):
        if not self._suspend_state_events:
            self._persist_current_state()

    # ---------- Filters ----------
    def _is_numeric_column(self, column: str) -> bool:
        try:
            numeric = pd.to_numeric(self.df_all[column], errors="coerce")
            return numeric.notna().sum() > 0
        except Exception:
            return False

    def _column_min_max(self, column: str) -> tuple[float | None, float | None]:
        try:
            numeric = pd.to_numeric(self.df_all[column], errors="coerce").dropna()
        except Exception:
            return (None, None)
        if numeric.empty:
            return (None, None)
        return (float(numeric.min()), float(numeric.max()))

    def _reset_filter(self, column: str):
        defaults = self.filter_defaults.get(column, (None, None))
        ctrl = self.filter_controls.get(column)
        if not ctrl:
            return
        min_val, max_val = defaults
        ctrl["min_var"].set("" if min_val is None else _format_number(min_val))
        ctrl["max_var"].set("" if max_val is None else _format_number(max_val))
        self._persist_current_state()

    def _refresh_filter_controls(self):
        for child in self.filters_frame.winfo_children():
            child.destroy()
        existing = {col: (ctrl["min_var"].get(), ctrl["max_var"].get()) for col, ctrl in self.filter_controls.items()}
        self.filter_controls.clear()
        self.filter_defaults.clear()

        if self.df_all.empty:
            ttk.Label(self.filters_frame, text="Load data to enable filters.", foreground="gray").grid(row=0, column=0, sticky=tk.W)
            return

        desired: list[tuple[str, str]] = []
        y_disp = self.y_var.get().strip()
        x_disp = self.x_var.get().strip()
        if y_disp in DEPENDENT_MAP:
            y_col = DEPENDENT_MAP[y_disp]
            if self._is_numeric_column(y_col):
                desired.append((y_disp, y_col))
        if x_disp in DEPENDENT_MAP:
            x_col = DEPENDENT_MAP[x_disp]
            if self._is_numeric_column(x_col):
                desired.append((x_disp, x_col))

        color_disp = self.color_var.get().strip()
        if color_disp and color_disp != "<None>":
            color_col = INDEPENDENT_TO_COLUMN.get(color_disp, color_disp)
            if self._is_numeric_column(color_col):
                desired.append((color_disp, color_col))

        shape_disp = self.shape_var.get().strip()
        if shape_disp and shape_disp != "<None>":
            shape_col = INDEPENDENT_TO_COLUMN.get(shape_disp, shape_disp)
            if self._is_numeric_column(shape_col):
                desired.append((shape_disp, shape_col))

        ordered = []
        seen = set()
        for display, column in desired:
            if column not in seen:
                ordered.append((display, column))
                seen.add(column)

        if not ordered:
            ttk.Label(
                self.filters_frame,
                text="No numeric filters available for the current selection.",
                foreground="gray",
            ).grid(row=0, column=0, sticky=tk.W)
            return

        for row_idx, (display, column) in enumerate(ordered):
            default_min, default_max = self._column_min_max(column)
            self.filter_defaults[column] = (default_min, default_max)
            prev_min, prev_max = existing.get(column, ("", ""))
            if not prev_min:
                prev_min = "" if default_min is None else _format_number(default_min)
            if not prev_max:
                prev_max = "" if default_max is None else _format_number(default_max)

            ttk.Label(self.filters_frame, text=f"{display}:").grid(row=row_idx, column=0, sticky=tk.W, padx=(0, 6), pady=2)
            min_var = tk.StringVar(value=prev_min)
            max_var = tk.StringVar(value=prev_max)
            min_entry = ttk.Entry(self.filters_frame, textvariable=min_var, width=12)
            max_entry = ttk.Entry(self.filters_frame, textvariable=max_var, width=12)
            min_entry.grid(row=row_idx, column=1, sticky=tk.W, padx=(0, 6))
            max_entry.grid(row=row_idx, column=2, sticky=tk.W, padx=(0, 6))
            ttk.Label(self.filters_frame, text="to").grid(row=row_idx, column=1, sticky=tk.E, padx=(0, 36))
            ttk.Button(
                self.filters_frame,
                text="Reset",
                command=lambda col=column: self._reset_filter(col),
                width=8,
            ).grid(row=row_idx, column=3, sticky=tk.W)

            self.filter_controls[column] = {
                "display": display,
                "min_var": min_var,
                "max_var": max_var,
            }

    def _apply_filters(self, df: pd.DataFrame) -> pd.DataFrame:
        filtered = df
        self.last_filters = {}
        for column, ctrl in self.filter_controls.items():
            try:
                numeric = pd.to_numeric(filtered[column], errors="coerce")
            except Exception:
                continue
            if numeric.notna().sum() == 0:
                continue
            min_str = ctrl["min_var"].get()
            max_str = ctrl["max_var"].get()
            try:
                min_val = _parse_float(min_str)
                max_val = _parse_float(max_str)
            except ValueError as exc:
                raise ValueError(f"{ctrl['display']}: {exc}") from exc
            mask = pd.Series(True, index=filtered.index)
            if min_val is not None:
                mask &= numeric >= min_val
            if max_val is not None:
                mask &= numeric <= max_val
            filtered = filtered[mask]
            self.last_filters[column] = {"display": ctrl["display"], "min": min_val, "max": max_val}
        return filtered
    # ---------- Encodings ----------
    def _prepare_color_encoding(self, df: pd.DataFrame, display: str, column: str | None):
        if not display or display == "<None>" or not column:
            base_color = _color_palette(1, self.mono_var.get())[0] if len(df) else OKABE_ITO[1]
            series = pd.Series(["__all__"] * len(df), index=df.index)
            color_map = {"__all__": base_color}
            metadata = {"variable": None, "type": "none", "labels": []}
            legend = []
            return series, color_map, metadata, legend

        raw = df[column]
        numeric = pd.to_numeric(raw, errors="coerce")
        metadata = {"variable": display}

        treat_categorical = display in INDEPENDENTS
        if not treat_categorical and (numeric.notna().sum() <= 1 or numeric.dropna().nunique() <= COLOR_BINS):
            treat_categorical = True

        if not treat_categorical and numeric.notna().sum() > 1 and numeric.dropna().nunique() > 1:
            bins = min(COLOR_BINS, numeric.dropna().nunique())
            try:
                categorized = pd.qcut(numeric, q=bins, duplicates="drop")
            except ValueError:
                categorized = pd.cut(numeric, bins, duplicates="drop")
            intervals = categorized.cat.categories
            label_map = {interval: _interval_label(interval.left, interval.right) for interval in intervals}
            labeled = categorized.map(label_map)
            series = pd.Series(MISSING_LABEL, index=df.index)
            series.loc[labeled.index] = labeled
            metadata.update(
                {
                    "type": "numeric",
                    "bins": len(intervals),
                    "labels": [label_map[i] for i in intervals],
                    "edges": [[float(i.left), float(i.right)] for i in intervals],
                }
            )
        else:
            decimals = CATEGORY_DECIMALS.get(display)
            series = raw.apply(
                lambda v: MISSING_LABEL if pd.isna(v) else _format_category_value(v, decimals)
            )
            if series.nunique() <= 1:
                value = series.iloc[0] if not series.empty else "All"
                series = pd.Series([value] * len(df), index=df.index)
            metadata.update({"type": "categorical"})

        unique_labels = []
        for lbl in series:
            if lbl not in unique_labels:
                unique_labels.append(lbl)
        unique_labels = sorted(unique_labels, key=_natural_sort_key)

        if len(unique_labels) > len(OKABE_ITO):
            allowed = unique_labels[: len(OKABE_ITO) - 1]
            series = series.apply(lambda x: x if x in allowed or x == MISSING_LABEL else "Other")
            unique_labels = []
            for lbl in series:
                if lbl not in unique_labels:
                    unique_labels.append(lbl)
            unique_labels = sorted(unique_labels, key=_natural_sort_key)

        non_special = [lbl for lbl in unique_labels if lbl not in ("__all__", MISSING_LABEL)]
        color_map: dict[str, str] = {}
        if display == "Water (g)" and non_special:
            preferred = ["#000000", "#F0E442", "#56B4E9", "#009E73", "#D55E00", "#CC79A7"]
            try:
                ordered = sorted(non_special, key=lambda lbl: float(lbl))
            except Exception:
                ordered = list(non_special)
            for idx, lbl in enumerate(ordered):
                color_map[lbl] = preferred[idx % len(preferred)]
        else:
            palette = _color_palette(len(non_special), self.mono_var.get())
            for lbl, color in zip(non_special, palette):
                color_map[lbl] = color

        if MISSING_LABEL in unique_labels:
            color_map[MISSING_LABEL] = "#999999" if not self.mono_var.get() else "#777777"
        if "__all__" in unique_labels:
            if display == "Water (g)" and non_special:
                base_color = "#000000"
            else:
                base_color = _color_palette(1, self.mono_var.get())[0] if non_special else OKABE_ITO[1]
            color_map["__all__"] = base_color

        variable_display = INDEPENDENT_LATEX.get(display, display)
        legend = []
        for lbl in unique_labels:
            if lbl == "__all__":
                continue
            if lbl == MISSING_LABEL:
                value_text = "Missing"
            else:
                value_text = lbl
            legend.append((f"{variable_display} = {value_text}", color_map.get(lbl, OKABE_ITO[1])))

        metadata["labels"] = unique_labels
        metadata["palette"] = color_map.copy()
        return series, color_map, metadata, legend

    def _prepare_shape_encoding(self, df: pd.DataFrame, display: str, column: str | None):
        if not display or display == "<None>" or not column:
            series = pd.Series(["__all__"] * len(df), index=df.index)
            shape_map = {"__all__": "o"}
            metadata = {"variable": None, "type": "none", "labels": []}
            legend = []
            return series, shape_map, metadata, legend

        raw = df[column]
        numeric = pd.to_numeric(raw, errors="coerce")
        metadata = {"variable": display}

        treat_categorical = display in INDEPENDENTS
        if not treat_categorical and (numeric.notna().sum() <= 1 or numeric.dropna().nunique() <= SHAPE_BINS):
            treat_categorical = True

        if not treat_categorical and numeric.notna().sum() > 1 and numeric.dropna().nunique() > 1:
            bins = min(SHAPE_BINS, len(SHAPE_MARKERS), numeric.dropna().nunique())
            try:
                categorized = pd.qcut(numeric, q=bins, duplicates="drop")
            except ValueError:
                categorized = pd.cut(numeric, bins, duplicates="drop")
            intervals = categorized.cat.categories
            label_map = {interval: _interval_label(interval.left, interval.right) for interval in intervals}
            labeled = categorized.map(label_map)
            series = pd.Series(MISSING_LABEL, index=df.index)
            series.loc[labeled.index] = labeled
            metadata.update(
                {
                    "type": "numeric",
                    "bins": len(intervals),
                    "labels": [label_map[i] for i in intervals],
                    "edges": [[float(i.left), float(i.right)] for i in intervals],
                }
            )
        else:
            decimals = CATEGORY_DECIMALS.get(display)
            series = raw.apply(
                lambda v: MISSING_LABEL if pd.isna(v) else _format_category_value(v, decimals)
            )
            metadata.update({"type": "categorical"})

        unique_labels = []
        for lbl in series:
            if lbl not in unique_labels:
                unique_labels.append(lbl)
        unique_labels = sorted(unique_labels, key=_natural_sort_key)

        if len(unique_labels) > len(SHAPE_MARKERS):
            allowed = unique_labels[: len(SHAPE_MARKERS) - 1]
            series = series.apply(lambda x: x if x in allowed or x == MISSING_LABEL else "Other")
            unique_labels = []
            for lbl in series:
                if lbl not in unique_labels:
                    unique_labels.append(lbl)
            unique_labels = sorted(unique_labels, key=_natural_sort_key)

        non_special = [lbl for lbl in unique_labels if lbl not in ("__all__", MISSING_LABEL)]
        shape_map: dict[str, str] = {}
        for idx, lbl in enumerate(non_special):
            shape_map[lbl] = SHAPE_MARKERS[idx % len(SHAPE_MARKERS)]
        if MISSING_LABEL in unique_labels:
            shape_map[MISSING_LABEL] = "o"
        if "__all__" in unique_labels:
            shape_map["__all__"] = "o"

        variable_display = INDEPENDENT_LATEX.get(display, display)
        legend = []
        for lbl in unique_labels:
            if lbl == "__all__":
                continue
            if lbl == MISSING_LABEL:
                value_text = "Missing"
            else:
                value_text = lbl
            legend.append((f"{variable_display} = {value_text}", shape_map.get(lbl, "o")))

        metadata["labels"] = unique_labels
        metadata["markers"] = shape_map.copy()
        return series, shape_map, metadata, legend

    # ---------- Plotting ----------
    def _style_axes(self, ax):
        ax.set_facecolor("white")
        ax.grid(False)
        for spine in ["top", "right"]:
            ax.spines[spine].set_visible(False)
        for spine in ["left", "bottom"]:
            ax.spines[spine].set_linewidth(0.8)

    def _render_plot(self):
        if self.df_all.empty:
            messagebox.showwarning("No data", "Load an All_Results Excel first.")
            return

        x_display = self.x_var.get().strip()
        y_display = self.y_var.get().strip()
        if x_display not in DEPENDENT_MAP or y_display not in DEPENDENT_MAP:
            messagebox.showwarning("Missing selection", "Select both X and Y dependent variables.")
            return

        x_column = DEPENDENT_MAP[x_display]
        y_column = DEPENDENT_MAP[y_display]

        color_display = self.color_var.get().strip()
        color_column = INDEPENDENT_TO_COLUMN.get(color_display, color_display) if color_display != "<None>" else None
        shape_display = self.shape_var.get().strip()
        shape_column = INDEPENDENT_TO_COLUMN.get(shape_display, shape_display) if shape_display != "<None>" else None
        yerr_column = None

        try:
            filtered = self._apply_filters(self.df_all)
        except ValueError as exc:
            messagebox.showerror("Invalid filter", str(exc))
            return

        if filtered.empty:
            messagebox.showwarning("No data", "Filters removed all rows.")
            return

        if self.errorbar_var.get():
            candidate = DEPENDENT_TO_DEVIATION.get(y_display)
            if candidate and candidate in filtered.columns:
                yerr_column = candidate

        x_numeric = pd.to_numeric(filtered[x_column], errors="coerce")
        y_numeric = pd.to_numeric(filtered[y_column], errors="coerce")
        valid = filtered[x_numeric.notna() & y_numeric.notna()].copy()
        if len(valid) < 2:
            messagebox.showerror("Not enough points", "At least 2 valid points required after filtering.")
            return

        color_series, color_map, color_meta, color_legend = self._prepare_color_encoding(valid, color_display, color_column)
        shape_series, shape_map, shape_meta, shape_legend = self._prepare_shape_encoding(valid, shape_display, shape_column)

        plot_df = valid.copy()
        plot_df["_color_label"] = color_series
        plot_df["_shape_label"] = shape_series
        self.df_filtered = plot_df.copy()
        if color_display != "<None>":
            self.df_filtered["Color category"] = color_series
        if shape_display != "<None>":
            self.df_filtered["Shape category"] = shape_series
        if yerr_column and yerr_column in plot_df.columns:
            self.df_filtered[yerr_column] = plot_df[yerr_column]

        self.ax.clear()
        self._style_axes(self.ax)
        self._current_legends = []

        groups = plot_df.groupby(["_color_label", "_shape_label"])
        for (color_key, shape_key), group in groups:
            color = color_map.get(color_key, OKABE_ITO[1])
            marker = shape_map.get(shape_key, "o")
            x_vals = _as_float_array(group[x_column])
            y_vals = _as_float_array(group[y_column])
            self.ax.scatter(
                x_vals,
                y_vals,
                s=48,
                marker=marker,
                facecolor=color,
                edgecolor="black",
                linewidths=0.6,
                alpha=0.9,
                zorder=3,
            )
            if yerr_column and yerr_column in group.columns:
                yerr_vals = _as_float_array(group[yerr_column])
                self.ax.errorbar(
                    x_vals,
                    y_vals,
                    yerr=yerr_vals,
                    fmt="none",
                    ecolor=color,
                    elinewidth=1.1,
                    capsize=3.5,
                    alpha=0.8,
                    zorder=2,
                )

        self.ax.set_xlabel(dependent_latex(x_display))
        self.ax.set_ylabel(dependent_latex(y_display))

        legend_count = 0
        right_margin = 0.88

        if color_display != "<None>" and color_legend:
            color_handles = [
                mlines.Line2D(
                    [0], [0],
                    marker="o",
                    linestyle="",
                    color=color,
                    markerfacecolor=color,
                    markeredgecolor="black",
                    markeredgewidth=0.6,
                    label=label,
                )
                for label, color in color_legend
            ]
            leg1 = self.ax.legend(
                color_handles,
                [h.get_label() for h in color_handles],
                loc="upper left",
                bbox_to_anchor=(1.02, 1.0),
                frameon=False,
                borderaxespad=0.0,
            )
            self.ax.add_artist(leg1)
            self._current_legends.append(leg1)
            legend_count += 1

        if shape_display != "<None>" and shape_legend:
            anchor_y = 1.0 if legend_count == 0 else 0.5
            shape_handles = [
                mlines.Line2D(
                    [0], [0],
                    marker=marker,
                    linestyle="",
                    color="#666666",
                    markerfacecolor="#666666" if not self.mono_var.get() else "#555555",
                    markeredgecolor="black",
                    markeredgewidth=0.6,
                    label=label,
                )
                for label, marker in shape_legend
            ]
            leg2 = self.ax.legend(
                shape_handles,
                [h.get_label() for h in shape_handles],
                loc="upper left",
                bbox_to_anchor=(1.02, anchor_y),
                frameon=False,
                borderaxespad=0.0,
            )
            self.ax.add_artist(leg2)
            self._current_legends.append(leg2)
            legend_count += 1

        if legend_count == 0:
            right_margin = 0.88
        elif legend_count == 1:
            right_margin = 0.80
        else:
            right_margin = 0.72

        self.fig.subplots_adjust(right=right_margin)

        sheet_label = self._sheet_labels.get(self.active_sheet_name, self.active_sheet_name or "<no sheet>")
        self.info_var.set(f"Sheet: {sheet_label}    |    Points: {len(valid)}")

        self.canvas.draw_idle()
        self.last_encoding_info = {"color": color_meta, "shape": shape_meta}
        self._persist_current_state()
    def _default_filename(self, ext: str) -> str:
        x = self.x_var.get().replace(" ", "_")
        y = self.y_var.get().replace(" ", "_")
        suffix = ""
        color = self.color_var.get()
        shape = self.shape_var.get()
        if color and color != "<None>":
            suffix += f"_col_{color.replace(' ', '_')}"
        if shape and shape != "<None>":
            suffix += f"_shape_{shape.replace(' ', '_')}"
        ts = _dt.datetime.now().strftime("%Y%m%d-%H%M")
        return f"dvsd_{y}_vs_{x}{suffix}_{ts}.{ext}"

    def _save_figure(self):
        if self.df_filtered.empty:
            messagebox.showwarning("No plot", "Render a plot before saving.")
            return
        default_name = self._default_filename("png")
        filename = filedialog.asksaveasfilename(
            title="Save Figure",
            defaultextension=".png",
            initialfile=default_name,
            filetypes=[
                ("PNG", "*.png"),
                ("TIFF", "*.tiff;*.tif"),
                ("SVG", "*.svg"),
            ],
        )
        if not filename:
            return
        try:
            self.fig.savefig(
                filename,
                dpi=int(self.dpi_var.get()),
                facecolor="white",
                bbox_inches="tight",
                bbox_extra_artists=self._current_legends,
            )
            messagebox.showinfo("Saved", f"Figure saved to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Save error", str(e))

    def _copy_figure(self):
        if self.df_filtered.empty:
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
                format="png",
                dpi=int(self.dpi_var.get()),
                facecolor="white",
                bbox_inches="tight",
                bbox_extra_artists=self._current_legends,
            )
            buffer.seek(0)
            image = Image.open(buffer).convert("RGB")
            with _io.BytesIO() as output:
                image.save(output, "BMP")
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

    def _export_data(self):
        if self.df_filtered.empty:
            messagebox.showwarning("No data", "Render a plot first (data depends on filters).")
            return
        x_display = self.x_var.get().strip()
        y_display = self.y_var.get().strip()
        x_column = DEPENDENT_MAP.get(x_display, x_display)
        y_column = DEPENDENT_MAP.get(y_display, y_display)
        color_display = self.color_var.get().strip()
        color_column = INDEPENDENT_TO_COLUMN.get(color_display, color_display) if color_display != "<None>" else None
        shape_display = self.shape_var.get().strip()
        shape_column = INDEPENDENT_TO_COLUMN.get(shape_display, shape_display) if shape_display != "<None>" else None

        used_cols = [col for col in [x_column, y_column] if col in self.df_filtered.columns]
        for col in [color_column, shape_column]:
            if col and col in self.df_filtered.columns and col not in used_cols:
                used_cols.append(col)
        for extra in ["Color category", "Shape category"]:
            if extra in self.df_filtered.columns and extra not in used_cols:
                used_cols.append(extra)

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
            self.df_filtered[used_cols].to_csv(csv_path, index=False)
            settings = {
                "file": self.file_var.get(),
                "sheet": self.active_sheet_name,
                "x": x_display,
                "y": y_display,
                "color": color_display if color_display != "<None>" else None,
                "shape": shape_display if shape_display != "<None>" else None,
                "monochrome": bool(self.mono_var.get()),
                "filters": self.last_filters,
                "color_encoding": self.last_encoding_info.get("color"),
                "shape_encoding": self.last_encoding_info.get("shape"),
                "dpi": int(self.dpi_var.get()),
                "n_total": int(len(self.df_filtered)),
            }
            json_path = os.path.splitext(csv_path)[0] + ".json"
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
            messagebox.showinfo("Exported", f"CSV saved to:\n{csv_path}\nSettings saved to:\n{json_path}")
        except Exception as e:
            messagebox.showerror("Export error", str(e))


def show(parent=None, settings_manager=None):
    root = parent if parent is not None else tk.Tk()
    DependentScatterModule(root, settings_manager)
    if parent is None:
        root.mainloop()

import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

import numpy as np
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib import colors as mcolors

from .foam_type_manager import FoamTypeManager


class TGATextParser:
    numeric_pattern = r"-?[\d,\.]+(?:e[+-]?\d+)?"

    @staticmethod
    def _to_float(value):
        try:
            return float(str(value).strip().replace(",", "."))
        except Exception:
            return None

    def parse_file(self, file_path):
        with open(file_path, "r", encoding="latin-1", errors="ignore") as handle:
            full_text = handle.read()
        lines = full_text.splitlines()

        sample_name, mass_mg = self._parse_sample_info(full_text, file_path)
        curves = self._parse_curve_blocks(lines)
        if not curves:
            raise ValueError("No se encontraron bloques 'Curve Values' válidos.")

        mass_curve = self._select_mass_curve(curves)
        if mass_curve is None or mass_curve.empty:
            raise ValueError("No se pudo extraer la curva de pérdida de masa.")

        derivative_curve = self._select_derivative_curve(curves)
        derivative_scale_factor = 100.0
        if derivative_curve is None or derivative_curve.empty:
            derivative_curve = self._build_derivative_from_mass(mass_curve)
            derivative_scale_factor = 1.0

        results = self._parse_results(full_text, derivative_curve, derivative_scale_factor)
        return {
            "sample_name": sample_name,
            "mass_mg": mass_mg,
            "mass_curve": mass_curve,
            "derivative_curve": derivative_curve,
            "derivative_scale_factor": derivative_scale_factor,
            "results": results,
        }

    def _parse_sample_info(self, full_text, file_path):
        sample_name = os.path.splitext(os.path.basename(file_path))[0]
        mass_mg = None
        match = re.search(r"Sample:\s*\n\s*(.*?),\s*(" + self.numeric_pattern + r")\s*mg", full_text, flags=re.IGNORECASE)
        if match:
            sample_name = match.group(1).strip()
            mass_mg = self._to_float(match.group(2))
        return sample_name, mass_mg

    def _parse_curve_blocks(self, lines):
        curves = []
        i = 0
        while i < len(lines):
            if lines[i].strip() != "Curve Values:":
                i += 1
                continue
            if i + 1 >= len(lines):
                break
            header_line = lines[i + 1]
            if "Index" not in header_line:
                i += 1
                continue
            columns = self._parse_header_columns(header_line)
            if not columns:
                i += 1
                continue
            i += 3
            rows = []
            while i < len(lines):
                raw = lines[i].strip()
                if not raw:
                    i += 1
                    if i < len(lines) and not lines[i].strip():
                        break
                    continue
                tokens = re.split(r"\s+", raw)
                if len(tokens) < len(columns):
                    break
                if not re.fullmatch(r"-?\d+", tokens[0]):
                    break
                parsed_row = []
                valid = True
                for token in tokens[: len(columns)]:
                    value = self._to_float(token)
                    if value is None:
                        valid = False
                        break
                    parsed_row.append(value)
                if valid:
                    rows.append(parsed_row)
                i += 1
            if rows:
                df = pd.DataFrame(rows, columns=columns)
                curves.append(df)
        return curves

    def _parse_header_columns(self, line):
        compact = re.sub(r"\s+", " ", line.strip())
        if "x value" in compact and "y value" in compact:
            return ["Index", "t", "Ts", "Tr", "x value", "y value"]
        if "Value" in compact:
            return ["Index", "t", "Ts", "Tr", "Value"]
        return []

    def _select_mass_curve(self, curves):
        for curve in curves:
            if "Value" in curve.columns:
                return curve
        return curves[0] if curves else None

    def _select_derivative_curve(self, curves):
        for curve in curves:
            if "x value" in curve.columns and "y value" in curve.columns:
                return curve
        return None

    def _build_derivative_from_mass(self, mass_curve):
        df = mass_curve.copy()
        x = self._temperature_axis(df).to_numpy(dtype=float)
        y = df["Value"].to_numpy(dtype=float)
        order = np.argsort(x)
        x_sorted = x[order]
        y_sorted = y[order]
        grad = np.gradient(y_sorted, x_sorted)
        derived = pd.DataFrame({"x value": x_sorted, "y value": grad})
        return derived

    def _parse_results(self, full_text, derivative_curve, derivative_scale_factor):
        left_limit = self._first_match_float(full_text, rf"Left\s+Limit\s+({self.numeric_pattern})\s*(?:°C|Â°C)")
        right_limit = self._first_match_float(full_text, rf"Right\s+Limit\s+({self.numeric_pattern})\s*(?:°C|Â°C)")
        inflect = self._first_match_float(full_text, rf"Inflect\.\s*Pt\.\s+({self.numeric_pattern})\s*(?:°C|Â°C)")
        midpoint = self._first_match_float(full_text, rf"Midpoint\s+({self.numeric_pattern})\s*(?:°C|Â°C)")
        step_pct = self._first_match_float(full_text, rf"Step\s+({self.numeric_pattern})\s*%")

        td = inflect if inflect is not None else midpoint
        peak_value = None
        x_der = derivative_curve["x value"].to_numpy(dtype=float)
        y_der = derivative_curve["y value"].to_numpy(dtype=float)
        if td is None and len(x_der):
            peak_idx = int(np.argmax(np.abs(y_der)))
            td = float(x_der[peak_idx])
            peak_value = float(y_der[peak_idx] * derivative_scale_factor)
        elif td is not None and len(x_der):
            peak_idx = int(np.argmin(np.abs(x_der - td)))
            peak_value = float(y_der[peak_idx] * derivative_scale_factor)

        return {
            "left_limit": left_limit,
            "right_limit": right_limit,
            "step_pct": step_pct,
            "td": td,
            "peak_derivative": peak_value,
        }

    def _first_match_float(self, text, pattern):
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if not match:
            return None
        return self._to_float(match.group(1))

    def _temperature_axis(self, df):
        if "Tr" in df.columns:
            return df["Tr"]
        if "Ts" in df.columns:
            return df["Ts"]
        return df.iloc[:, 0]


class TGAImageEditor:
    tab_names = ("Mass Loss", "Derivative (DTG)")
    export_figsize = (8.5, 5.4)
    export_dpi = 600
    tab_colors = {
        "Mass Loss": "#0072B2",
        "Derivative (DTG)": "#CC79A7",
    }

    def __init__(self, root):
        self.root = root
        self.root.title("TGA Image Editor")
        self.root.geometry("1560x980")
        self.root.minsize(1320, 860)

        self.foam_manager = FoamTypeManager()
        self.parser = TGATextParser()
        self.parsed = None

        self.filepath_var = tk.StringVar()
        self.summary_var = tk.StringVar(value="No file loaded")
        self.status_var = tk.StringVar(value="Ready")

        self.axes = {}
        self.figures = {}
        self.canvases = {}
        self.controls = {}
        self.tab_plot_limits = {}

        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        top = ttk.LabelFrame(main, text="Input TGA file", padding=10)
        top.pack(fill=tk.X, pady=(0, 10))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="TXT file:").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(top, textvariable=self.filepath_var, state="readonly").grid(row=0, column=1, sticky=(tk.W, tk.E))
        ttk.Button(top, text="Browse", command=self.open_file).grid(row=0, column=2, padx=(8, 0))
        ttk.Button(top, text="Reload", command=self.reload_current_file).grid(row=0, column=3, padx=(8, 0))
        ttk.Button(top, text="Export 2 PNG", command=lambda: self.export_all("png")).grid(row=0, column=4, padx=(8, 0))
        ttk.Button(top, text="Export 2 PDF", command=lambda: self.export_all("pdf")).grid(row=0, column=5, padx=(8, 0))

        ttk.Label(top, textvariable=self.summary_var).grid(row=1, column=0, columnspan=6, sticky=tk.W, pady=(8, 0))

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        for tab_name in self.tab_names:
            self._build_tab(tab_name)

        ttk.Label(main, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, pady=(8, 0))

    def _build_tab(self, tab_name):
        tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(tab, text=tab_name)

        tab.columnconfigure(0, weight=4)
        tab.columnconfigure(1, weight=2)
        tab.rowconfigure(0, weight=1)

        plot_frame = ttk.Frame(tab)
        plot_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))

        fig = Figure(figsize=(8.5, 5.4), dpi=110)
        ax = fig.add_subplot(111)
        canvas = FigureCanvasTkAgg(fig, master=plot_frame)
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        ctr = ttk.LabelFrame(tab, text="Labels and export", padding=10)
        ctr.grid(row=0, column=1, sticky=(tk.N, tk.S, tk.W, tk.E), padx=(10, 0))
        ctr.columnconfigure(0, weight=1)
        ctr.configure(width=360, height=760)
        ctr.grid_propagate(False)

        show_temp = tk.BooleanVar(value=True)
        show_metric = tk.BooleanVar(value=True)
        curve_color = tk.StringVar(value=self.tab_colors.get(tab_name, "#1f77b4"))
        baseline_color = tk.StringVar(value=self.tab_colors.get(tab_name, "#1f77b4"))
        line_width = tk.DoubleVar(value=2.2)
        font_size = tk.DoubleVar(value=28.0)
        invert_x = tk.BooleanVar(value=False)
        show_baseline = tk.BooleanVar(value=True)
        left_limit = tk.StringVar(value="")
        right_limit = tk.StringVar(value="")
        x_temp = tk.DoubleVar(value=0.0)
        y_temp = tk.DoubleVar(value=0.0)
        x_metric = tk.DoubleVar(value=0.0)
        y_metric = tk.DoubleVar(value=0.0)

        temp_chk = ttk.Checkbutton(ctr, text="Show Td label", variable=show_temp, command=lambda n=tab_name: self.redraw_tab(n))
        temp_chk.grid(row=0, column=0, sticky=tk.W)
        metric_chk = ttk.Checkbutton(ctr, text="Show metric label", variable=show_metric, command=lambda n=tab_name: self.redraw_tab(n))
        metric_chk.grid(row=1, column=0, sticky=tk.W, pady=(0, 6))

        ttk.Label(ctr, text="Curve color (HEX)").grid(row=2, column=0, sticky=tk.W)
        c_row = ttk.Frame(ctr)
        c_row.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        c_row.columnconfigure(0, weight=1)
        ttk.Entry(c_row, textvariable=curve_color).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(c_row, text="Pick", width=6, command=lambda n=tab_name: self.pick_color(n, "curve_color")).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(ctr, text="Baseline color (HEX)").grid(row=4, column=0, sticky=tk.W)
        b_row = ttk.Frame(ctr)
        b_row.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        b_row.columnconfigure(0, weight=1)
        ttk.Entry(b_row, textvariable=baseline_color).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(b_row, text="Pick", width=6, command=lambda n=tab_name: self.pick_color(n, "baseline_color")).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(ctr, text="Line width").grid(row=6, column=0, sticky=tk.W)
        lw_spin = tk.Spinbox(ctr, from_=0.5, to=6.0, increment=0.1, textvariable=line_width, command=lambda n=tab_name: self.redraw_tab(n))
        lw_spin.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        lw_spin.bind("<KeyRelease>", lambda _e, n=tab_name: self.redraw_tab(n))

        ttk.Label(ctr, text="Font size").grid(row=8, column=0, sticky=tk.W)
        fs_spin = tk.Spinbox(ctr, from_=10.0, to=40.0, increment=0.5, textvariable=font_size, command=lambda n=tab_name: self.redraw_tab(n))
        fs_spin.grid(row=9, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        fs_spin.bind("<KeyRelease>", lambda _e, n=tab_name: self.redraw_tab(n))

        ttk.Checkbutton(ctr, text="Invert X axis (high → low)", variable=invert_x, command=lambda n=tab_name: self.redraw_tab(n)).grid(row=10, column=0, sticky=tk.W, pady=(0, 6))
        baseline_chk = ttk.Checkbutton(ctr, text="Show baseline", variable=show_baseline, command=lambda n=tab_name: self.redraw_tab(n))
        baseline_chk.grid(row=11, column=0, sticky=tk.W, pady=(0, 6))

        ttk.Label(ctr, text="Integration left limit (°C)").grid(row=12, column=0, sticky=tk.W)
        left_entry = ttk.Entry(ctr, textvariable=left_limit)
        left_entry.grid(row=13, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        left_entry.bind("<KeyRelease>", lambda _e, n=tab_name: self.redraw_tab(n))

        ttk.Label(ctr, text="Integration right limit (°C)").grid(row=14, column=0, sticky=tk.W)
        right_entry = ttk.Entry(ctr, textvariable=right_limit)
        right_entry.grid(row=15, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        right_entry.bind("<KeyRelease>", lambda _e, n=tab_name: self.redraw_tab(n))

        ttk.Label(ctr, text="Td label X").grid(row=16, column=0, sticky=tk.W)
        temp_x_scale = ttk.Scale(ctr, variable=x_temp, from_=0, to=1, command=lambda _v, n=tab_name: self.redraw_tab(n))
        temp_x_scale.grid(row=17, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(ctr, text="Td label Y").grid(row=18, column=0, sticky=tk.W)
        temp_y_scale = ttk.Scale(ctr, variable=y_temp, from_=0, to=1, command=lambda _v, n=tab_name: self.redraw_tab(n))
        temp_y_scale.grid(row=19, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(ctr, text="Metric label X").grid(row=20, column=0, sticky=tk.W)
        metric_x_scale = ttk.Scale(ctr, variable=x_metric, from_=0, to=1, command=lambda _v, n=tab_name: self.redraw_tab(n))
        metric_x_scale.grid(row=21, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(ctr, text="Metric label Y").grid(row=22, column=0, sticky=tk.W)
        metric_y_scale = ttk.Scale(ctr, variable=y_metric, from_=0, to=1, command=lambda _v, n=tab_name: self.redraw_tab(n))
        metric_y_scale.grid(row=23, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Button(ctr, text="Export PNG", command=lambda n=tab_name: self.export_tab(n, "png")).grid(row=24, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        ttk.Button(ctr, text="Export PDF", command=lambda n=tab_name: self.export_tab(n, "pdf")).grid(row=25, column=0, sticky=(tk.W, tk.E))

        self.axes[tab_name] = ax
        self.figures[tab_name] = fig
        self.canvases[tab_name] = canvas
        self.controls[tab_name] = {
            "show_temp": show_temp,
            "show_metric": show_metric,
            "curve_color": curve_color,
            "baseline_color": baseline_color,
            "line_width": line_width,
            "font_size": font_size,
            "invert_x": invert_x,
            "show_baseline": show_baseline,
            "left_limit": left_limit,
            "right_limit": right_limit,
            "x_temp": x_temp,
            "y_temp": y_temp,
            "x_metric": x_metric,
            "y_metric": y_metric,
            "temp_x_scale": temp_x_scale,
            "temp_y_scale": temp_y_scale,
            "metric_x_scale": metric_x_scale,
            "metric_y_scale": metric_y_scale,
        }

        if tab_name == "Mass Loss":
            temp_chk.state(["disabled"])
            metric_chk.state(["disabled"])
            baseline_chk.state(["disabled"])
        else:
            metric_chk.state(["disabled"])

    def open_file(self):
        initial_dir = None
        current_foam = self.foam_manager.get_current_foam_type()
        suggestions = self.foam_manager.get_suggested_paths("TGA", current_foam)
        if suggestions and suggestions.get("input_folder"):
            initial_dir = suggestions.get("input_folder")

        path = filedialog.askopenfilename(
            title="Select TGA TXT file",
            initialdir=initial_dir,
            filetypes=[("TGA text files", "*.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        self.filepath_var.set(path)
        self._load_file(path)

    def reload_current_file(self):
        path = self.filepath_var.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("No file", "Select a TGA TXT file first.")
            return
        self._load_file(path)

    def _load_file(self, file_path):
        try:
            self.parsed = self.parser.parse_file(file_path)
            self._update_summary()
            self._prepare_controls()
            for tab_name in self.tab_names:
                self.redraw_tab(tab_name)
            self.status_var.set(f"Loaded {os.path.basename(file_path)}")
        except Exception as exc:
            messagebox.showerror("Error", f"Could not parse TGA file:\n{exc}")
            self.status_var.set("Error loading file")

    def _update_summary(self):
        if not self.parsed:
            self.summary_var.set("No file loaded")
            return
        sample = self.parsed["sample_name"]
        mass = self.parsed["mass_mg"]
        mass_text = f"{mass:.3f} mg" if isinstance(mass, float) else "n/a"
        td = self.parsed["results"].get("td")
        td_text = f"{td:.2f} °C" if isinstance(td, float) else "n/a"
        self.summary_var.set(f"Sample: {sample} | Mass: {mass_text} | Td: {td_text}")

    def _prepare_controls(self):
        if not self.parsed:
            return
        self.tab_plot_limits = {}
        left_limit = self.parsed["results"].get("left_limit")
        right_limit = self.parsed["results"].get("right_limit")
        td = self.parsed["results"].get("td")

        x_ranges = []
        for tab_name in self.tab_names:
            x_data, _ = self._tab_data(tab_name)
            if len(x_data):
                x_ranges.append((float(np.min(x_data)), float(np.max(x_data))))
        shared_x = None
        if x_ranges:
            x0 = min(v[0] for v in x_ranges)
            x1 = max(v[1] for v in x_ranges)
            if x0 == x1:
                x1 = x0 + 1.0
            x_pad = 0.03 * (x1 - x0)
            shared_x = (x0 - x_pad, x1 + x_pad)

        for tab_name in self.tab_names:
            x_data, y_data = self._tab_data(tab_name)
            if len(x_data) == 0:
                continue
            x_min, x_max = float(np.min(x_data)), float(np.max(x_data))
            y_min, y_max = float(np.min(y_data)), float(np.max(y_data))
            if x_min == x_max:
                x_max = x_min + 1.0
            if y_min == y_max:
                y_max = y_min + 1.0
            y_pad = 0.08 * (y_max - y_min)
            x_lim = shared_x if shared_x else (x_min, x_max)
            self.tab_plot_limits[tab_name] = (x_lim[0], x_lim[1], y_min - y_pad, y_max + y_pad)

            ctr = self.controls[tab_name]
            ctr["temp_x_scale"].configure(from_=x_min, to=x_max)
            ctr["temp_y_scale"].configure(from_=y_min, to=y_max)
            ctr["metric_x_scale"].configure(from_=x_min, to=x_max)
            ctr["metric_y_scale"].configure(from_=y_min, to=y_max)

            if tab_name == "Mass Loss":
                ctr["x_temp"].set(td if td is not None else x_min + 0.62 * (x_max - x_min))
                ctr["y_temp"].set(y_min + 0.82 * (y_max - y_min))
                ctr["x_metric"].set(x_min + 0.15 * (x_max - x_min))
                ctr["y_metric"].set(y_min + 0.16 * (y_max - y_min))
                ctr["show_temp"].set(False)
                ctr["show_metric"].set(False)
                ctr["show_baseline"].set(False)
            else:
                ctr["x_temp"].set(x_min + 0.12 * (x_max - x_min))
                ctr["y_temp"].set(y_min + 0.08 * (y_max - y_min))
                ctr["x_metric"].set(x_min + 0.20 * (x_max - x_min))
                ctr["y_metric"].set(y_min + 0.12 * (y_max - y_min))
                ctr["show_temp"].set(True)
                ctr["show_metric"].set(False)

            ctr["curve_color"].set(self.tab_colors.get(tab_name, "#1f77b4"))
            ctr["baseline_color"].set(self.tab_colors.get(tab_name, "#1f77b4"))
            ctr["line_width"].set(2.2)
            ctr["font_size"].set(28.0)
            ctr["invert_x"].set(False)
            if tab_name != "Mass Loss":
                ctr["show_baseline"].set(True)
            ctr["left_limit"].set("" if left_limit is None else f"{left_limit:.3f}")
            ctr["right_limit"].set("" if right_limit is None else f"{right_limit:.3f}")

    def _tab_data(self, tab_name):
        if not self.parsed:
            return np.array([]), np.array([])
        if tab_name == "Mass Loss":
            curve = self.parsed["mass_curve"]
            x_series = curve["Tr"] if "Tr" in curve.columns else curve["Ts"]
            y_series = curve["Value"]
        else:
            curve = self.parsed["derivative_curve"]
            scale_factor = float(self.parsed.get("derivative_scale_factor", 1.0))
            x_series = curve["x value"] if "x value" in curve.columns else (curve["Tr"] if "Tr" in curve.columns else curve["Ts"])
            y_series = (curve["y value"] if "y value" in curve.columns else curve["Value"]) * scale_factor
        return x_series.to_numpy(dtype=float), y_series.to_numpy(dtype=float)

    def redraw_tab(self, tab_name):
        fig = self.figures.get(tab_name)
        ax = self.axes.get(tab_name)
        canvas = self.canvases.get(tab_name)
        if not fig or not ax or not canvas:
            return
        ax.clear()
        ax.set_facecolor("white")

        if not self.parsed:
            ax.text(0.5, 0.5, "Load a TGA TXT file", ha="center", va="center", transform=ax.transAxes, fontsize=12)
            fig.tight_layout()
            canvas.draw_idle()
            return

        x_data, y_data = self._tab_data(tab_name)
        if len(x_data) == 0:
            ax.text(0.5, 0.5, "No data for this plot", ha="center", va="center", transform=ax.transAxes, fontsize=12)
            fig.tight_layout()
            canvas.draw_idle()
            return

        ctr = self.controls[tab_name]
        color = self.resolve_color(ctr["curve_color"].get(), self.tab_colors.get(tab_name, "#1f77b4"))
        baseline_color = self.resolve_color(ctr["baseline_color"].get(), color)
        lw = self.read_float(ctr["line_width"], 2.2, minimum=0.4)
        fs = self.read_float(ctr["font_size"], 28.0, minimum=8.0)

        ax.plot(x_data, y_data, color=color, linewidth=lw)
        ax.set_xlabel("Temperature (°C)", fontname="DejaVu Sans", fontsize=fs)
        if tab_name == "Mass Loss":
            ax.set_ylabel("Mass (%)", fontname="DejaVu Sans", fontsize=fs)
        else:
            ax.set_ylabel("d(Mass)/dT (%/°C)", fontname="DejaVu Sans", fontsize=fs)
        ax.tick_params(direction="in", top=True, right=True, labelsize=max(8.0, fs - 2.0))
        ax.grid(False)
        if ctr["invert_x"].get():
            ax.invert_xaxis()

        limits = self.tab_plot_limits.get(tab_name)
        if limits:
            x0, x1, y0, y1 = limits
            ax.set_xlim(x0, x1)
            ax.set_ylim(y0, y1)

        left, right = self.get_limits(ctr)
        if tab_name != "Mass Loss" and ctr["show_baseline"].get() and left is not None and right is not None and left != right:
            sorted_idx = np.argsort(x_data)
            xs = x_data[sorted_idx]
            ys = y_data[sorted_idx]
            x_left = float(np.clip(left, np.min(x_data), np.max(x_data)))
            x_right = float(np.clip(right, np.min(x_data), np.max(x_data)))
            y_left = float(np.interp(x_left, xs, ys))
            y_right = float(np.interp(x_right, xs, ys))
            ax.plot([x_left, x_right], [y_left, y_right], linestyle="--", dashes=(6, 4), linewidth=max(1.2, 0.8 * lw), color=baseline_color, zorder=5)

        td = self.parsed["results"].get("td")
        if tab_name == "Derivative (DTG)" and ctr["show_temp"].get() and td is not None:
            ax.text(
                ctr["x_temp"].get(),
                ctr["y_temp"].get(),
                f"$T_{{d}}={td:.2f}$ °C",
                color="#000000",
                fontsize=max(8.0, fs - 1.0),
                fontname="DejaVu Sans",
            )

        fig.tight_layout()
        canvas.draw_idle()

    def get_limits(self, ctr):
        left = self.read_float(ctr["left_limit"], None)
        right = self.read_float(ctr["right_limit"], None)
        return left, right

    def read_float(self, variable, fallback, minimum=None):
        try:
            value = float(variable.get())
        except Exception:
            return fallback
        if minimum is not None:
            return max(minimum, value)
        return value

    def resolve_color(self, text, fallback):
        value = (text or "").strip()
        if not value:
            return fallback
        try:
            mcolors.to_rgba(value)
            return value
        except Exception:
            return fallback

    def pick_color(self, tab_name, key):
        ctr = self.controls.get(tab_name)
        if not ctr or key not in ctr:
            return
        initial = ctr[key].get()
        chosen = colorchooser.askcolor(color=initial, title=f"Select {tab_name} {key.replace('_', ' ')}")
        if chosen[1]:
            ctr[key].set(chosen[1])
            self.redraw_tab(tab_name)

    def export_tab(self, tab_name, ext):
        if not self.parsed:
            messagebox.showwarning("No data", "Load a TGA TXT file first.")
            return
        sample = self.parsed["sample_name"].replace(" ", "_")
        default_name = f"{sample}_{tab_name.replace(' ', '_').replace('(', '').replace(')', '')}.{ext}"
        out = filedialog.asksaveasfilename(
            title=f"Save {tab_name}",
            defaultextension=f".{ext}",
            initialfile=default_name,
            filetypes=[(ext.upper(), f"*.{ext}"), ("All files", "*.*")],
        )
        if not out:
            return
        try:
            self._save_fixed_figure(self.figures[tab_name], out)
            self.status_var.set(f"Saved {os.path.basename(out)}")
        except Exception as exc:
            messagebox.showerror("Error", f"Could not save figure:\n{exc}")
            self.status_var.set("Error saving figure")

    def export_all(self, ext):
        if not self.parsed:
            messagebox.showwarning("No data", "Load a TGA TXT file first.")
            return
        out_dir = filedialog.askdirectory(title=f"Select folder to export 2 {ext.upper()} figures")
        if not out_dir:
            return
        sample = self.parsed["sample_name"].replace(" ", "_")
        errors = []
        for tab_name in self.tab_names:
            try:
                filename = f"{sample}_{tab_name.replace(' ', '_').replace('(', '').replace(')', '')}.{ext}"
                out = os.path.join(out_dir, filename)
                self._save_fixed_figure(self.figures[tab_name], out)
            except Exception as exc:
                errors.append(f"{tab_name}: {exc}")
        if errors:
            messagebox.showerror("Export finished with errors", "\n".join(errors))
            self.status_var.set("Exported with errors")
        else:
            self.status_var.set(f"Saved 2 {ext.upper()} figures")
            messagebox.showinfo("Export completed", f"2 {ext.upper()} figures exported to:\n{out_dir}")

    def _save_fixed_figure(self, figure, output_path):
        figure.canvas.draw()
        figure.savefig(output_path, dpi=self.export_dpi)


class TGAImageModule:
    def __init__(self, root):
        self.editor = TGAImageEditor(root)

import io
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

import numpy as np
import pandas as pd
from PIL import Image
from matplotlib import colors as mcolors
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

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
        index = 0
        while index < len(lines):
            if lines[index].strip() != "Curve Values:":
                index += 1
                continue
            if index + 1 >= len(lines):
                break
            header_line = lines[index + 1]
            if "Index" not in header_line:
                index += 1
                continue
            columns = self._parse_header_columns(header_line)
            if not columns:
                index += 1
                continue
            index += 3
            rows = []
            while index < len(lines):
                raw = lines[index].strip()
                if not raw:
                    index += 1
                    if index < len(lines) and not lines[index].strip():
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
                index += 1
            if rows:
                curves.append(pd.DataFrame(rows, columns=columns))
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
        x = self._temperature_axis(mass_curve).to_numpy(dtype=float)
        y = mass_curve["Value"].to_numpy(dtype=float)
        order = np.argsort(x)
        x_sorted = x[order]
        y_sorted = y[order]
        grad = np.gradient(y_sorted, x_sorted)
        return pd.DataFrame({"x value": x_sorted, "y value": grad})

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
    export_figsize = (9.8, 8.908333333333333)
    export_dpi = 600
    mass_color_default = "#0072B2"
    dtg_color_default = "#CC79A7"
    export_layout = {"left": 0.14, "right": 0.84, "bottom": 0.17, "top": 0.97}

    def __init__(self, root):
        self.root = root
        self.root.title("TGA Image Editor")
        self.root.geometry("1560x1180")
        self.root.minsize(1320, 920)

        self.foam_manager = FoamTypeManager()
        self.parser = TGATextParser()
        self.parsed = None

        self.filepath_var = tk.StringVar()
        self.summary_var = tk.StringVar(value="No file loaded")
        self.status_var = tk.StringVar(value="Ready")

        self.figure = None
        self.ax_mass = None
        self.ax_dtg = None
        self.canvas = None
        self.controls = {}
        self.plot_limits = None

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
        ttk.Button(top, text="Export PNG", command=lambda: self.export_figure("png")).grid(row=0, column=4, padx=(8, 0))
        ttk.Button(top, text="Export PDF", command=lambda: self.export_figure("pdf")).grid(row=0, column=5, padx=(8, 0))
        ttk.Button(top, text="Copy Image", command=self.copy_image).grid(row=0, column=6, padx=(8, 0))
        ttk.Label(top, textvariable=self.summary_var).grid(row=1, column=0, columnspan=7, sticky=tk.W, pady=(8, 0))

        content = ttk.Frame(main)
        content.pack(fill=tk.BOTH, expand=True)
        content.columnconfigure(0, weight=4)
        content.columnconfigure(1, weight=2)
        content.rowconfigure(0, weight=1)

        plot_frame = ttk.LabelFrame(content, text="TGA Overlay (Mass + DTG)", padding=8)
        plot_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.figure = Figure(figsize=(8.5, 5.4), dpi=110)
        self.ax_mass = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        ctr = ttk.LabelFrame(content, text="Labels and style", padding=10)
        ctr.grid(row=0, column=1, sticky=(tk.N, tk.S, tk.W, tk.E), padx=(10, 0))
        ctr.columnconfigure(0, weight=1)
        ctr.configure(width=360, height=980)
        ctr.grid_propagate(False)

        show_td = tk.BooleanVar(value=True)
        show_td_line = tk.BooleanVar(value=True)
        show_legend = tk.BooleanVar(value=True)
        mass_color = tk.StringVar(value=self.mass_color_default)
        dtg_color = tk.StringVar(value=self.dtg_color_default)
        line_width = tk.DoubleVar(value=2.2)
        font_size = tk.DoubleVar(value=28.0)
        invert_x = tk.BooleanVar(value=False)
        x_td = tk.DoubleVar(value=0.0)
        y_td = tk.DoubleVar(value=0.0)
        legend_x = tk.DoubleVar(value=0.84)
        legend_y = tk.DoubleVar(value=0.50)

        ttk.Checkbutton(ctr, text="Show Td label", variable=show_td, command=self.redraw_plot).grid(row=0, column=0, sticky=tk.W, pady=(0, 6))
        ttk.Checkbutton(ctr, text="Show Td line", variable=show_td_line, command=self.redraw_plot).grid(row=1, column=0, sticky=tk.W, pady=(0, 6))
        ttk.Checkbutton(ctr, text="Show legend", variable=show_legend, command=self.redraw_plot).grid(row=2, column=0, sticky=tk.W, pady=(0, 6))

        ttk.Label(ctr, text="Mass color (HEX)").grid(row=3, column=0, sticky=tk.W)
        mass_row = ttk.Frame(ctr)
        mass_row.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        mass_row.columnconfigure(0, weight=1)
        ttk.Entry(mass_row, textvariable=mass_color).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(mass_row, text="Pick", width=6, command=lambda: self.pick_color("mass_color")).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(ctr, text="DTG color (HEX)").grid(row=5, column=0, sticky=tk.W)
        dtg_row = ttk.Frame(ctr)
        dtg_row.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        dtg_row.columnconfigure(0, weight=1)
        ttk.Entry(dtg_row, textvariable=dtg_color).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(dtg_row, text="Pick", width=6, command=lambda: self.pick_color("dtg_color")).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(ctr, text="Line width").grid(row=7, column=0, sticky=tk.W)
        lw_spin = tk.Spinbox(ctr, from_=0.5, to=6.0, increment=0.1, textvariable=line_width, command=self.redraw_plot)
        lw_spin.grid(row=8, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        lw_spin.bind("<KeyRelease>", lambda _e: self.redraw_plot())

        ttk.Label(ctr, text="Font size").grid(row=9, column=0, sticky=tk.W)
        fs_spin = tk.Spinbox(ctr, from_=10.0, to=40.0, increment=0.5, textvariable=font_size, command=self.redraw_plot)
        fs_spin.grid(row=10, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        fs_spin.bind("<KeyRelease>", lambda _e: self.redraw_plot())

        ttk.Checkbutton(ctr, text="Invert X axis (high → low)", variable=invert_x, command=self.redraw_plot).grid(row=11, column=0, sticky=tk.W, pady=(0, 8))

        ttk.Label(ctr, text="Td label X").grid(row=12, column=0, sticky=tk.W)
        td_x_scale = ttk.Scale(ctr, variable=x_td, from_=0, to=1, command=lambda _v: self.redraw_plot())
        td_x_scale.grid(row=13, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(ctr, text="Td label Y").grid(row=14, column=0, sticky=tk.W)
        td_y_scale = ttk.Scale(ctr, variable=y_td, from_=0, to=1, command=lambda _v: self.redraw_plot())
        td_y_scale.grid(row=15, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(ctr, text="Legend X (axes)").grid(row=16, column=0, sticky=tk.W)
        legend_x_scale = ttk.Scale(ctr, variable=legend_x, from_=0.0, to=1.0, command=lambda _v: self.redraw_plot())
        legend_x_scale.grid(row=17, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(ctr, text="Legend Y (axes)").grid(row=18, column=0, sticky=tk.W)
        legend_y_scale = ttk.Scale(ctr, variable=legend_y, from_=0.0, to=1.0, command=lambda _v: self.redraw_plot())
        legend_y_scale.grid(row=19, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        self.controls = {
            "show_td": show_td,
            "show_td_line": show_td_line,
            "show_legend": show_legend,
            "mass_color": mass_color,
            "dtg_color": dtg_color,
            "line_width": line_width,
            "font_size": font_size,
            "invert_x": invert_x,
            "x_td": x_td,
            "y_td": y_td,
            "legend_x": legend_x,
            "legend_y": legend_y,
            "td_x_scale": td_x_scale,
            "td_y_scale": td_y_scale,
            "legend_x_scale": legend_x_scale,
            "legend_y_scale": legend_y_scale,
        }

        ttk.Label(main, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, pady=(8, 0))

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
            self.redraw_plot()
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
        td = self.parsed["results"].get("td")
        mass_text = f"{mass:.3f} mg" if isinstance(mass, float) else "n/a"
        td_text = f"{td:.2f} °C" if isinstance(td, float) else "n/a"
        self.summary_var.set(f"Sample: {sample} | Mass: {mass_text} | Td: {td_text}")

    def _mass_data(self):
        if not self.parsed:
            return np.array([]), np.array([])
        curve = self.parsed["mass_curve"]
        x = (curve["Tr"] if "Tr" in curve.columns else curve["Ts"]).to_numpy(dtype=float)
        y = curve["Value"].to_numpy(dtype=float)
        return x, y

    def _dtg_data(self):
        if not self.parsed:
            return np.array([]), np.array([])
        curve = self.parsed["derivative_curve"]
        scale_factor = float(self.parsed.get("derivative_scale_factor", 1.0))
        x_series = curve["x value"] if "x value" in curve.columns else (curve["Tr"] if "Tr" in curve.columns else curve["Ts"])
        y_series = (curve["y value"] if "y value" in curve.columns else curve["Value"]) * scale_factor
        return x_series.to_numpy(dtype=float), y_series.to_numpy(dtype=float)

    def _prepare_controls(self):
        x_mass, y_mass = self._mass_data()
        x_dtg, y_dtg = self._dtg_data()
        if len(x_mass) == 0 or len(x_dtg) == 0:
            return

        x_min = min(float(np.min(x_mass)), float(np.min(x_dtg)))
        x_max = max(float(np.max(x_mass)), float(np.max(x_dtg)))
        if x_min == x_max:
            x_max = x_min + 1.0
        x_pad = 0.03 * (x_max - x_min)

        ym_min, ym_max = float(np.min(y_mass)), float(np.max(y_mass))
        yd_min, yd_max = float(np.min(y_dtg)), float(np.max(y_dtg))
        if ym_min == ym_max:
            ym_max = ym_min + 1.0
        if yd_min == yd_max:
            yd_max = yd_min + 1.0
        ym_pad = 0.08 * (ym_max - ym_min)
        yd_pad = 0.08 * (yd_max - yd_min)

        self.plot_limits = {
            "x": (x_min - x_pad, x_max + x_pad),
            "mass_y": (ym_min - ym_pad, ym_max + ym_pad),
            "dtg_y": (yd_min - yd_pad, yd_max + yd_pad),
        }

        ctr = self.controls
        ctr["td_x_scale"].configure(from_=x_min, to=x_max)
        ctr["td_y_scale"].configure(from_=yd_min, to=yd_max)
        ctr["mass_color"].set(self.mass_color_default)
        ctr["dtg_color"].set(self.dtg_color_default)
        ctr["line_width"].set(2.2)
        ctr["font_size"].set(28.0)
        ctr["invert_x"].set(False)
        ctr["show_td"].set(True)
        ctr["show_td_line"].set(True)
        ctr["show_legend"].set(True)
        ctr["x_td"].set(x_min + 0.12 * (x_max - x_min))
        ctr["y_td"].set(yd_min + 0.08 * (yd_max - yd_min))
        ctr["legend_x"].set(0.84)
        ctr["legend_y"].set(0.50)

    def redraw_plot(self):
        if self.figure is None or self.ax_mass is None or self.canvas is None:
            return
        self.ax_mass.clear()
        if self.ax_dtg is not None:
            try:
                self.ax_dtg.remove()
            except Exception:
                pass
            self.ax_dtg = None

        if not self.parsed:
            self.ax_mass.text(0.5, 0.5, "Load a TGA TXT file", ha="center", va="center", transform=self.ax_mass.transAxes, fontsize=12)
            self._apply_layout()
            self.canvas.draw_idle()
            return

        x_mass, y_mass = self._mass_data()
        x_dtg, y_dtg = self._dtg_data()
        if len(x_mass) == 0 or len(x_dtg) == 0:
            self.ax_mass.text(0.5, 0.5, "No data for this plot", ha="center", va="center", transform=self.ax_mass.transAxes, fontsize=12)
            self._apply_layout()
            self.canvas.draw_idle()
            return

        ctr = self.controls
        fs = self._read_float(ctr["font_size"], 28.0, minimum=8.0)
        lw = self._read_float(ctr["line_width"], 2.2, minimum=0.4)
        mass_color = self._resolve_color(ctr["mass_color"].get(), self.mass_color_default)
        dtg_color = self._resolve_color(ctr["dtg_color"].get(), self.dtg_color_default)

        mass_lw = max(0.6, lw)
        dtg_lw = max(0.6, lw * 0.78)

        mass_line, = self.ax_mass.plot(
            x_mass,
            y_mass,
            color=mass_color,
            linewidth=mass_lw,
            linestyle="-",
            alpha=1.0,
            zorder=3,
            label="Mass loss",
        )
        self.ax_mass.set_xlabel("Temperature (°C)", fontname="DejaVu Sans", fontsize=fs)
        self.ax_mass.set_ylabel("Mass (%)", color="#000000", fontname="DejaVu Sans", fontsize=fs)
        self.ax_mass.tick_params(axis="x", direction="in", top=True, labelsize=max(8.0, fs - 2.0))
        self.ax_mass.tick_params(axis="y", direction="in", left=True, right=False, colors="#000000", labelsize=max(8.0, fs - 2.0))
        self.ax_mass.grid(False)

        self.ax_dtg = self.ax_mass.twinx()
        dtg_line, = self.ax_dtg.plot(
            x_dtg,
            y_dtg,
            color=dtg_color,
            linewidth=dtg_lw,
            linestyle="--",
            dashes=(6, 4),
            alpha=0.75,
            zorder=2,
            label="DTG",
        )
        self.ax_dtg.set_ylabel("d(Mass)/dT (%/°C)", color="#000000", fontname="DejaVu Sans", fontsize=fs)
        self.ax_dtg.tick_params(axis="y", direction="in", left=False, right=True, colors="#000000", labelsize=max(8.0, fs - 2.0))
        self.ax_dtg.grid(False)
        if ctr["show_legend"].get():
            legend = self.ax_mass.legend(
                handles=[mass_line, dtg_line],
                loc="center",
                bbox_to_anchor=(ctr["legend_x"].get(), ctr["legend_y"].get()),
                bbox_transform=self.ax_mass.transAxes,
                frameon=True,
                fancybox=False,
                framealpha=1.0,
                fontsize=max(8.0, fs - 10.0),
            )
            legend.get_frame().set_edgecolor("#000000")
            legend.get_frame().set_linewidth(0.9)

        if self.plot_limits:
            x0, x1 = self.plot_limits["x"]
            ym0, ym1 = self.plot_limits["mass_y"]
            yd0, yd1 = self.plot_limits["dtg_y"]
            self.ax_mass.set_xlim(x0, x1)
            self.ax_mass.set_ylim(ym0, ym1)
            self.ax_dtg.set_ylim(yd0, yd1)

        if ctr["invert_x"].get():
            self.ax_mass.invert_xaxis()
            self.ax_dtg.invert_xaxis()

        td = self.parsed["results"].get("td")
        if ctr["show_td_line"].get() and td is not None:
            self.ax_mass.axvline(
                td,
                color="#6e6e6e",
                linestyle="--",
                dashes=(3, 3),
                linewidth=max(0.8, 0.45 * lw),
                alpha=0.55,
                zorder=1,
            )
        if ctr["show_td"].get() and td is not None:
            self.ax_dtg.text(
                ctr["x_td"].get(),
                ctr["y_td"].get(),
                f"$T_{{d}}={td:.2f}$ °C",
                color="#000000",
                fontsize=max(8.0, fs - 1.0),
                fontname="DejaVu Sans",
            )

        self._apply_layout()
        self.canvas.draw_idle()

    def _resolve_color(self, text, fallback):
        value = (text or "").strip()
        if not value:
            return fallback
        try:
            mcolors.to_rgba(value)
            return value
        except Exception:
            return fallback

    def _read_float(self, variable, fallback, minimum=None):
        try:
            value = float(variable.get())
        except Exception:
            return fallback
        if minimum is not None:
            return max(minimum, value)
        return value

    def pick_color(self, key):
        if key not in self.controls:
            return
        initial = self.controls[key].get()
        chosen = colorchooser.askcolor(color=initial, title=f"Select {key.replace('_', ' ')}")
        if chosen[1]:
            self.controls[key].set(chosen[1])
            self.redraw_plot()

    def export_figure(self, ext):
        if not self.parsed:
            messagebox.showwarning("No data", "Load a TGA TXT file first.")
            return
        sample = self.parsed["sample_name"].replace(" ", "_")
        default_name = f"{sample}_TGA_overlay.{ext}"
        out = filedialog.asksaveasfilename(
            title="Save TGA overlay figure",
            defaultextension=f".{ext}",
            initialfile=default_name,
            filetypes=[(ext.upper(), f"*.{ext}"), ("All files", "*.*")],
        )
        if not out:
            return
        try:
            self._save_current_figure(out)
            self.status_var.set(f"Saved {os.path.basename(out)}")
        except Exception as exc:
            messagebox.showerror("Error", f"Could not save figure:\n{exc}")
            self.status_var.set("Error saving figure")

    def copy_image(self):
        if self.figure is None:
            return
        original_size = tuple(self.figure.get_size_inches())
        try:
            self.figure.set_size_inches(*self.export_figsize, forward=True)
            self._apply_layout()
            buffer = io.BytesIO()
            self.figure.savefig(buffer, format="png", dpi=self.export_dpi, bbox_inches=None, pad_inches=0)
            buffer.seek(0)
            image = Image.open(buffer).convert("RGB")
            bmp = io.BytesIO()
            image.save(bmp, format="BMP")
            data = bmp.getvalue()[14:]
            import win32clipboard

            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            finally:
                win32clipboard.CloseClipboard()
            self.status_var.set("Copied TGA image to clipboard")
        except Exception as exc:
            messagebox.showerror("Clipboard error", f"Could not copy image:\n{exc}")
        finally:
            self.figure.set_size_inches(*original_size, forward=True)
            self._apply_layout()
            self.figure.canvas.draw_idle()

    def _apply_layout(self):
        if self.figure is None:
            return
        self.figure.subplots_adjust(**self.export_layout)

    def _save_current_figure(self, output_path):
        if self.figure is None:
            return
        original_size = tuple(self.figure.get_size_inches())
        try:
            self.figure.set_size_inches(*self.export_figsize, forward=True)
            self._apply_layout()
            self.figure.canvas.draw()
            self.figure.savefig(output_path, dpi=self.export_dpi, bbox_inches=None, pad_inches=0)
        finally:
            self.figure.set_size_inches(*original_size, forward=True)
            self._apply_layout()
            self.figure.canvas.draw_idle()




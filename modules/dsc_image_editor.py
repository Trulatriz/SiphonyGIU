import os
import re
import io
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

import numpy as np
import pandas as pd
from PIL import Image
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib import colors as mcolors
from matplotlib.figure import Figure

from .foam_type_manager import FoamTypeManager


class DSCTextParser:
    phase_order = ("1st Heating", "Cooling", "2nd Heating")
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

        sample_name, mass_mg, results_body = self._parse_header(full_text)
        data = self._parse_numeric_data(full_text)
        if data.empty:
            raise ValueError("No numeric data rows were found in the file.")

        segments = self._split_segments(data)
        parsed_results = self._parse_results(results_body or "")

        return {
            "sample_name": sample_name or os.path.splitext(os.path.basename(file_path))[0],
            "mass_mg": mass_mg,
            "data": data,
            "segments": segments,
            "results": parsed_results,
        }

    def _parse_header(self, full_text):
        sample_name = None
        mass_mg = None
        sample_match = re.search(r"Sample:\s*(.*?),\s*([\d,\.]+)\s*mg", full_text)
        if sample_match:
            sample_name = sample_match.group(1).replace("DSC", "").strip()
            mass_mg = self._to_float(sample_match.group(2))
        parts = full_text.split("Results:")
        results_body = parts[1] if len(parts) > 1 else ""
        return sample_name, mass_mg, results_body

    def _parse_numeric_data(self, full_text):
        rows = []
        for line in full_text.splitlines():
            stripped = line.strip()
            if not stripped:
                continue
            tokens = re.split(r"\s+", stripped)
            if len(tokens) < 5:
                continue
            index_token = tokens[0]
            if not re.fullmatch(r"-?\d+", index_token):
                continue
            index_value = int(index_token)
            numeric_values = [self._to_float(tokens[i]) for i in range(1, 5)]
            if any(value is None for value in numeric_values):
                continue
            rows.append((index_value, *numeric_values))

        return pd.DataFrame(rows, columns=["Index", "t", "Ts", "Tr", "Value"])

    def _split_segments(self, data_frame):
        total = len(data_frame)
        if total < 9:
            chunk = max(1, total // 3)
            split_points = [chunk, min(chunk * 2, total)]
        else:
            temperature = data_frame["Ts"].to_numpy(dtype=float)
            dtemp = np.diff(temperature)
            if len(dtemp) >= 5:
                kernel = np.ones(5, dtype=float) / 5.0
                dtemp_smooth = np.convolve(dtemp, kernel, mode="same")
            else:
                dtemp_smooth = dtemp

            median_abs = float(np.median(np.abs(dtemp_smooth))) if len(dtemp_smooth) else 0.0
            epsilon = max(1e-6, median_abs * 0.2)
            sign = np.zeros_like(dtemp_smooth, dtype=int)
            sign[dtemp_smooth > epsilon] = 1
            sign[dtemp_smooth < -epsilon] = -1

            prev = 1
            for i in range(len(sign)):
                if sign[i] == 0:
                    sign[i] = prev
                else:
                    prev = sign[i]

            for i in range(len(sign) - 2, -1, -1):
                if sign[i] == 0:
                    sign[i] = sign[i + 1]

            turn_heat_to_cool = None
            turn_cool_to_heat = None
            for i in range(1, len(sign)):
                if turn_heat_to_cool is None and sign[i - 1] > 0 and sign[i] < 0:
                    turn_heat_to_cool = i
                elif turn_heat_to_cool is not None and sign[i - 1] < 0 and sign[i] > 0:
                    turn_cool_to_heat = i
                    break

            if turn_heat_to_cool is None:
                turn_heat_to_cool = int(np.argmax(temperature))
            if turn_cool_to_heat is None:
                search_start = min(turn_heat_to_cool + 1, total - 1)
                turn_cool_to_heat = int(np.argmin(temperature[search_start:]) + search_start)

            split_points = sorted(
                [
                    max(1, min(turn_heat_to_cool + 1, total - 2)),
                    max(2, min(turn_cool_to_heat + 1, total - 1)),
                ]
            )

            if split_points[1] <= split_points[0]:
                split_points = [total // 3, (2 * total) // 3]

        s1 = max(1, min(split_points[0], total - 2))
        s2 = max(s1 + 1, min(split_points[1], total - 1))
        phase_slices = [
            data_frame.iloc[:s1].copy(),
            data_frame.iloc[s1:s2].copy(),
            data_frame.iloc[s2:].copy(),
        ]
        return {phase: segment for phase, segment in zip(self.phase_order, phase_slices)}

    def _parse_results(self, results_body):
        semi_events = self._parse_semicrystalline_events(results_body)
        amorphous_events = self._parse_amorphous_events(results_body)
        if semi_events:
            return {
                "mode": "semicrystalline",
                "events": semi_events,
                "semicrystalline_events": semi_events,
                "amorphous_events": amorphous_events,
            }
        if amorphous_events:
            return {
                "mode": "amorphous",
                "events": amorphous_events,
                "semicrystalline_events": semi_events,
                "amorphous_events": amorphous_events,
            }
        return {
            "mode": "unknown",
            "events": [],
            "semicrystalline_events": semi_events,
            "amorphous_events": amorphous_events,
        }

    def _parse_semicrystalline_events(self, text):
        blocks = re.split(r"(?=Crystallinity\s+-?[\d,\.]+\s*%)", text, flags=re.IGNORECASE)
        candidate_blocks = [block for block in blocks if re.search(r"Crystallinity\s+-?[\d,\.]+\s*%", block, flags=re.IGNORECASE)]
        if not candidate_blocks:
            return []

        events = []
        for idx, phase in enumerate(self.phase_order):
            block = candidate_blocks[idx] if idx < len(candidate_blocks) else ""
            peak = self._first_match_float(block, rf"Peak\s+({self.numeric_pattern})\s*(?:°C|Â°C|Ã‚Â°C)")
            crystallinity = self._first_match_float(block, rf"Crystallinity\s+({self.numeric_pattern})\s*%")
            left_limit = self._first_match_float(block, rf"Left\s+Limit\s+({self.numeric_pattern})\s*(?:°C|Â°C|Ã‚Â°C)")
            right_limit = self._first_match_float(block, rf"Right\s+Limit\s+({self.numeric_pattern})\s*(?:°C|Â°C|Ã‚Â°C)")
            onset = self._first_match_float(block, rf"Onset\s+({self.numeric_pattern})\s*(?:°C|Â°C|Ã‚Â°C)")
            if peak is None and crystallinity is None and onset is None:
                continue
            events.append(
                {
                    "phase": phase,
                    "temp_label": "Tc" if phase == "Cooling" else "Tm",
                    "temperature": peak if peak is not None else onset,
                    "crystallinity": crystallinity,
                    "left_limit": left_limit,
                    "right_limit": right_limit,
                }
            )
        return events

    def _parse_amorphous_events(self, text):
        pattern = re.compile(
            rf"Glass Transition.*?Midpoint ISO\s+({self.numeric_pattern})\s*(?:°C|Â°C|Ã‚Â°C).*?Delta cp\s+({self.numeric_pattern})\s*Jg\^-1K\^-1",
            re.IGNORECASE | re.DOTALL,
        )
        matches = list(pattern.finditer(text))
        if not matches:
            return []

        events = []
        phase_map = {
            0: "1st Heating",
            1: "2nd Heating",
        }
        for idx, match in enumerate(matches):
            phase = phase_map.get(idx)
            if not phase:
                continue
            events.append(
                {
                    "phase": phase,
                    "temp_label": "Tg",
                    "temperature": self._to_float(match.group(1)),
                    "crystallinity": self._to_float(match.group(2)),
                    "left_limit": None,
                    "right_limit": None,
                }
            )
        return events

    def _first_match_float(self, text, pattern):
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if not match:
            return None
        return self._to_float(match.group(1))


class DSCImageEditor:
    phase_order = ("1st Heating", "Cooling", "2nd Heating")
    combined_phase = "Cooling + 1st Heating"
    export_figsize = (9.8, 8.908333333333333)
    export_dpi = 600
    export_layout = {"left": 0.14, "right": 0.84, "bottom": 0.17, "top": 0.97}
    phase_colors = {
        "1st Heating": "#E69F00",
        "Cooling": "#56B4E9",
        "2nd Heating": "#009E73",
    }

    def __init__(self, root):
        self.root = root
        self.root.title("DSC Image Editor")
        self.root.geometry("1560x1180")
        self.root.minsize(1320, 920)

        self.foam_manager = FoamTypeManager()
        self.parser = DSCTextParser()
        self.parsed = None

        self.filepath_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.classification_var = tk.StringVar(value="auto")

        self.phase_axes = {}
        self.phase_figures = {}
        self.phase_canvases = {}
        self.phase_controls = {}
        self.phase_plot_limits = {}

        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        top = ttk.LabelFrame(main, text="Input DSC file", padding=10)
        top.pack(fill=tk.X, pady=(0, 10))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="TXT file:").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(top, textvariable=self.filepath_var, state="readonly").grid(row=0, column=1, sticky=(tk.W, tk.E))
        ttk.Button(top, text="Browse", command=self.open_file).grid(row=0, column=2, padx=(8, 0))
        ttk.Button(top, text="Reload", command=self.reload_current_file).grid(row=0, column=3, padx=(8, 0))
        ttk.Button(top, text="Export 4 PNG", command=lambda: self.export_all_phases("png")).grid(row=0, column=4, padx=(8, 0))
        ttk.Button(top, text="Export 4 PDF", command=lambda: self.export_all_phases("pdf")).grid(row=0, column=5, padx=(8, 0))
        ttk.Label(top, text="Classification:").grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        mode_combo = ttk.Combobox(
            top,
            textvariable=self.classification_var,
            values=("auto", "semicrystalline", "amorphous"),
            state="readonly",
            width=16,
        )
        mode_combo.grid(row=1, column=1, sticky=tk.W, pady=(8, 0))
        mode_combo.bind("<<ComboboxSelected>>", lambda _e: self.on_classification_changed())

        self.summary_var = tk.StringVar(value="No file loaded")
        ttk.Label(top, textvariable=self.summary_var).grid(row=2, column=0, columnspan=6, sticky=tk.W, pady=(8, 0))

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        for phase in (*self.phase_order, self.combined_phase):
            self._build_phase_tab(phase)

        ttk.Label(main, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, pady=(8, 0))

    def _build_phase_tab(self, phase):
        tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(tab, text=phase)

        tab.columnconfigure(0, weight=4)
        tab.columnconfigure(1, weight=2)
        tab.rowconfigure(0, weight=1)

        plot_frame = ttk.Frame(tab)
        plot_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))

        fig = Figure(figsize=(8.5, 5.4), dpi=110)
        ax = fig.add_subplot(111)
        canvas = FigureCanvasTkAgg(fig, master=plot_frame)
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        controls = ttk.LabelFrame(tab, text="Labels and export", padding=10)
        controls.grid(row=0, column=1, sticky=(tk.N, tk.S, tk.W, tk.E), padx=(10, 0))
        controls.columnconfigure(0, weight=1)
        controls.configure(width=360, height=980)
        controls.grid_propagate(False)

        if phase == self.combined_phase:
            self._build_combined_controls(phase, fig, ax, canvas, controls)
            return

        show_temp_var = tk.BooleanVar(value=True)
        show_cryst_var = tk.BooleanVar(value=True)
        curve_color_var = tk.StringVar(value=self.phase_colors.get(phase, "#1f77b4"))
        baseline_color_var = tk.StringVar(value=self.phase_colors.get(phase, "#1f77b4"))
        line_width_var = tk.DoubleVar(value=2.2)
        font_size_var = tk.DoubleVar(value=28.0)
        invert_x_var = tk.BooleanVar(value=(phase == "Cooling"))
        show_baseline_var = tk.BooleanVar(value=True)
        left_limit_var = tk.StringVar(value="")
        right_limit_var = tk.StringVar(value="")
        x_temp_var = tk.DoubleVar(value=0.0)
        y_temp_var = tk.DoubleVar(value=0.0)
        x_cryst_var = tk.DoubleVar(value=0.0)
        y_cryst_var = tk.DoubleVar(value=0.0)

        ttk.Checkbutton(
            controls,
            text="Show T label",
            variable=show_temp_var,
            command=lambda p=phase: self.redraw_phase(p),
        ).grid(row=0, column=0, sticky=tk.W)

        ttk.Checkbutton(
            controls,
            text="Show χc/Δcp label",
            variable=show_cryst_var,
            command=lambda p=phase: self.redraw_phase(p),
        ).grid(row=1, column=0, sticky=tk.W, pady=(0, 6))

        ttk.Label(controls, text="Curve color (HEX)").grid(row=2, column=0, sticky=tk.W)
        curve_row = ttk.Frame(controls)
        curve_row.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        curve_row.columnconfigure(0, weight=1)
        ttk.Entry(curve_row, textvariable=curve_color_var).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(curve_row, text="Pick", width=6, command=lambda p=phase: self.pick_color_for_phase(p, "curve_color")).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(controls, text="Baseline color (HEX)").grid(row=4, column=0, sticky=tk.W)
        baseline_row = ttk.Frame(controls)
        baseline_row.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        baseline_row.columnconfigure(0, weight=1)
        ttk.Entry(baseline_row, textvariable=baseline_color_var).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(baseline_row, text="Pick", width=6, command=lambda p=phase: self.pick_color_for_phase(p, "baseline_color")).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(controls, text="Line width").grid(row=6, column=0, sticky=tk.W)
        line_width_spin = tk.Spinbox(
            controls,
            from_=0.5,
            to=6.0,
            increment=0.1,
            textvariable=line_width_var,
            command=lambda p=phase: self.redraw_phase(p),
        )
        line_width_spin.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        line_width_spin.bind("<KeyRelease>", lambda _e, p=phase: self.redraw_phase(p))

        ttk.Label(controls, text="Font size").grid(row=8, column=0, sticky=tk.W)
        font_size_spin = tk.Spinbox(
            controls,
            from_=10.0,
            to=40.0,
            increment=0.5,
            textvariable=font_size_var,
            command=lambda p=phase: self.redraw_phase(p),
        )
        font_size_spin.grid(row=9, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        font_size_spin.bind("<KeyRelease>", lambda _e, p=phase: self.redraw_phase(p))

        ttk.Checkbutton(
            controls,
            text="Invert X axis (high to low)",
            variable=invert_x_var,
            command=lambda p=phase: self.redraw_phase(p),
        ).grid(row=10, column=0, sticky=tk.W, pady=(0, 6))

        ttk.Checkbutton(
            controls,
            text="Show baseline",
            variable=show_baseline_var,
            command=lambda p=phase: self.redraw_phase(p),
        ).grid(row=11, column=0, sticky=tk.W, pady=(0, 6))

        ttk.Label(controls, text="Integration left limit (°C)").grid(row=12, column=0, sticky=tk.W)
        left_entry = ttk.Entry(controls, textvariable=left_limit_var)
        left_entry.grid(row=13, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        left_entry.bind("<KeyRelease>", lambda _e, p=phase: self.redraw_phase(p))

        ttk.Label(controls, text="Integration right limit (°C)").grid(row=14, column=0, sticky=tk.W)
        right_entry = ttk.Entry(controls, textvariable=right_limit_var)
        right_entry.grid(row=15, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        right_entry.bind("<KeyRelease>", lambda _e, p=phase: self.redraw_phase(p))

        ttk.Label(controls, text="Temp label X").grid(row=16, column=0, sticky=tk.W)
        temp_x_scale = ttk.Scale(controls, variable=x_temp_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        temp_x_scale.grid(row=17, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="Temp label Y").grid(row=18, column=0, sticky=tk.W)
        temp_y_scale = ttk.Scale(controls, variable=y_temp_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        temp_y_scale.grid(row=19, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="χc/Δcp label X").grid(row=20, column=0, sticky=tk.W)
        cryst_x_scale = ttk.Scale(controls, variable=x_cryst_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        cryst_x_scale.grid(row=21, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="χc/Δcp label Y").grid(row=22, column=0, sticky=tk.W)
        cryst_y_scale = ttk.Scale(controls, variable=y_cryst_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        cryst_y_scale.grid(row=23, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Button(controls, text="Export PNG", command=lambda p=phase: self.export_phase_figure(p, "png")).grid(row=24, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        ttk.Button(controls, text="Export PDF", command=lambda p=phase: self.export_phase_figure(p, "pdf")).grid(row=25, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        ttk.Button(controls, text="Copy Image", command=lambda p=phase: self.copy_phase_image(p)).grid(row=26, column=0, sticky=(tk.W, tk.E))
        copy_btn = ttk.Button(
            controls,
            text="Copy Tm/χc coords from 1st Heating",
            command=self.copy_heating_coordinates,
        )
        copy_btn.grid(row=27, column=0, sticky=(tk.W, tk.E), pady=(6, 0))
        if phase != "2nd Heating":
            copy_btn.state(["disabled"])

        self.phase_axes[phase] = ax
        self.phase_figures[phase] = fig
        self.phase_canvases[phase] = canvas
        self.phase_controls[phase] = {
            "show_temp": show_temp_var,
            "show_cryst": show_cryst_var,
            "curve_color": curve_color_var,
            "baseline_color": baseline_color_var,
            "line_width": line_width_var,
            "font_size": font_size_var,
            "invert_x": invert_x_var,
            "show_baseline": show_baseline_var,
            "left_limit": left_limit_var,
            "right_limit": right_limit_var,
            "x_temp": x_temp_var,
            "y_temp": y_temp_var,
            "x_cryst": x_cryst_var,
            "y_cryst": y_cryst_var,
            "temp_x_scale": temp_x_scale,
            "temp_y_scale": temp_y_scale,
            "cryst_x_scale": cryst_x_scale,
            "cryst_y_scale": cryst_y_scale,
        }

    def _build_combined_controls(self, phase, fig, ax, canvas, controls):
        cooling_color_var = tk.StringVar(value="#0072b2")
        heating_color_var = tk.StringVar(value="#d55e00")
        line_width_var = tk.DoubleVar(value=2.2)
        label_font_size_var = tk.DoubleVar(value=28.0)
        arrow_font_size_var = tk.DoubleVar(value=20.0)
        invert_x_var = tk.BooleanVar(value=False)
        show_cooling_arrow_var = tk.BooleanVar(value=True)
        show_heating_arrow_var = tk.BooleanVar(value=True)
        cooling_arrow_x_var = tk.DoubleVar(value=0.0)
        cooling_arrow_y_var = tk.DoubleVar(value=0.0)
        heating_arrow_x_var = tk.DoubleVar(value=0.0)
        heating_arrow_y_var = tk.DoubleVar(value=0.0)
        exo_x_var = tk.DoubleVar(value=0.0)
        exo_y_var = tk.DoubleVar(value=0.0)
        tc_x_var = tk.DoubleVar(value=0.0)
        tc_y_var = tk.DoubleVar(value=0.0)
        tm_x_var = tk.DoubleVar(value=0.0)
        tm_y_var = tk.DoubleVar(value=0.0)
        xc2_x_var = tk.DoubleVar(value=0.0)
        xc2_y_var = tk.DoubleVar(value=0.0)

        ttk.Label(controls, text="Cooling color (HEX)").grid(row=0, column=0, sticky=tk.W)
        cooling_row = ttk.Frame(controls)
        cooling_row.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        cooling_row.columnconfigure(0, weight=1)
        ttk.Entry(cooling_row, textvariable=cooling_color_var).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(cooling_row, text="Pick", width=6, command=lambda p=phase: self.pick_color_for_phase(p, "cooling_color")).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(controls, text="1st Heating color (HEX)").grid(row=2, column=0, sticky=tk.W)
        heating_row = ttk.Frame(controls)
        heating_row.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        heating_row.columnconfigure(0, weight=1)
        ttk.Entry(heating_row, textvariable=heating_color_var).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(heating_row, text="Pick", width=6, command=lambda p=phase: self.pick_color_for_phase(p, "heating_color")).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(controls, text="Line width").grid(row=4, column=0, sticky=tk.W)
        line_width_spin = tk.Spinbox(
            controls,
            from_=0.5,
            to=6.0,
            increment=0.1,
            textvariable=line_width_var,
            command=lambda p=phase: self.redraw_phase(p),
        )
        line_width_spin.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        line_width_spin.bind("<KeyRelease>", lambda _e, p=phase: self.redraw_phase(p))

        ttk.Label(controls, text="Label font size").grid(row=6, column=0, sticky=tk.W)
        label_font_size_spin = tk.Spinbox(
            controls,
            from_=10.0,
            to=40.0,
            increment=0.5,
            textvariable=label_font_size_var,
            command=lambda p=phase: self.redraw_phase(p),
        )
        label_font_size_spin.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        label_font_size_spin.bind("<KeyRelease>", lambda _e, p=phase: self.redraw_phase(p))

        ttk.Label(controls, text="Arrow font size").grid(row=8, column=0, sticky=tk.W)
        arrow_font_size_spin = tk.Spinbox(
            controls,
            from_=8.0,
            to=36.0,
            increment=0.5,
            textvariable=arrow_font_size_var,
            command=lambda p=phase: self.redraw_phase(p),
        )
        arrow_font_size_spin.grid(row=9, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        arrow_font_size_spin.bind("<KeyRelease>", lambda _e, p=phase: self.redraw_phase(p))

        ttk.Checkbutton(
            controls,
            text="Invert X axis (high to low)",
            variable=invert_x_var,
            command=lambda p=phase: self.redraw_phase(p),
        ).grid(row=10, column=0, sticky=tk.W, pady=(0, 6))

        ttk.Checkbutton(
            controls,
            text="Show Cooling arrow (left)",
            variable=show_cooling_arrow_var,
            command=lambda p=phase: self.redraw_phase(p),
        ).grid(row=11, column=0, sticky=tk.W)
        ttk.Checkbutton(
            controls,
            text="Show 1st Heating arrow (right)",
            variable=show_heating_arrow_var,
            command=lambda p=phase: self.redraw_phase(p),
        ).grid(row=12, column=0, sticky=tk.W, pady=(0, 6))

        ttk.Label(controls, text="Cooling arrow X").grid(row=13, column=0, sticky=tk.W)
        cooling_arrow_x_scale = ttk.Scale(controls, variable=cooling_arrow_x_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        cooling_arrow_x_scale.grid(row=14, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="Cooling arrow Y").grid(row=15, column=0, sticky=tk.W)
        cooling_arrow_y_scale = ttk.Scale(controls, variable=cooling_arrow_y_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        cooling_arrow_y_scale.grid(row=16, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="1st Heating arrow X").grid(row=17, column=0, sticky=tk.W)
        heating_arrow_x_scale = ttk.Scale(controls, variable=heating_arrow_x_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        heating_arrow_x_scale.grid(row=18, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="1st Heating arrow Y").grid(row=19, column=0, sticky=tk.W)
        heating_arrow_y_scale = ttk.Scale(controls, variable=heating_arrow_y_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        heating_arrow_y_scale.grid(row=20, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="EXO X").grid(row=21, column=0, sticky=tk.W)
        exo_x_scale = ttk.Scale(controls, variable=exo_x_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        exo_x_scale.grid(row=22, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="EXO Y").grid(row=23, column=0, sticky=tk.W)
        exo_y_scale = ttk.Scale(controls, variable=exo_y_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        exo_y_scale.grid(row=24, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="Tc label X").grid(row=25, column=0, sticky=tk.W)
        tc_x_scale = ttk.Scale(controls, variable=tc_x_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        tc_x_scale.grid(row=26, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="Tc label Y").grid(row=27, column=0, sticky=tk.W)
        tc_y_scale = ttk.Scale(controls, variable=tc_y_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        tc_y_scale.grid(row=28, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="Tm label X").grid(row=29, column=0, sticky=tk.W)
        tm_x_scale = ttk.Scale(controls, variable=tm_x_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        tm_x_scale.grid(row=30, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="Tm label Y").grid(row=31, column=0, sticky=tk.W)
        tm_y_scale = ttk.Scale(controls, variable=tm_y_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        tm_y_scale.grid(row=32, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="Xc (1st) label X").grid(row=33, column=0, sticky=tk.W)
        xc2_x_scale = ttk.Scale(controls, variable=xc2_x_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        xc2_x_scale.grid(row=34, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="Xc (1st) label Y").grid(row=35, column=0, sticky=tk.W)
        xc2_y_scale = ttk.Scale(controls, variable=xc2_y_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        xc2_y_scale.grid(row=36, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Button(controls, text="Export PNG", command=lambda p=phase: self.export_phase_figure(p, "png")).grid(row=37, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        ttk.Button(controls, text="Export PDF", command=lambda p=phase: self.export_phase_figure(p, "pdf")).grid(row=38, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        ttk.Button(controls, text="Copy Image", command=lambda p=phase: self.copy_phase_image(p)).grid(row=39, column=0, sticky=(tk.W, tk.E))

        self.phase_axes[phase] = ax
        self.phase_figures[phase] = fig
        self.phase_canvases[phase] = canvas
        self.phase_controls[phase] = {
            "cooling_color": cooling_color_var,
            "heating_color": heating_color_var,
            "line_width": line_width_var,
            "label_font_size": label_font_size_var,
            "arrow_font_size": arrow_font_size_var,
            "invert_x": invert_x_var,
            "show_cooling_arrow": show_cooling_arrow_var,
            "show_heating_arrow": show_heating_arrow_var,
            "cooling_arrow_x": cooling_arrow_x_var,
            "cooling_arrow_y": cooling_arrow_y_var,
            "heating_arrow_x": heating_arrow_x_var,
            "heating_arrow_y": heating_arrow_y_var,
            "exo_x": exo_x_var,
            "exo_y": exo_y_var,
            "tc_x": tc_x_var,
            "tc_y": tc_y_var,
            "tm_x": tm_x_var,
            "tm_y": tm_y_var,
            "xc2_x": xc2_x_var,
            "xc2_y": xc2_y_var,
            "cooling_arrow_x_scale": cooling_arrow_x_scale,
            "cooling_arrow_y_scale": cooling_arrow_y_scale,
            "heating_arrow_x_scale": heating_arrow_x_scale,
            "heating_arrow_y_scale": heating_arrow_y_scale,
            "exo_x_scale": exo_x_scale,
            "exo_y_scale": exo_y_scale,
            "tc_x_scale": tc_x_scale,
            "tc_y_scale": tc_y_scale,
            "tm_x_scale": tm_x_scale,
            "tm_y_scale": tm_y_scale,
            "xc2_x_scale": xc2_x_scale,
            "xc2_y_scale": xc2_y_scale,
        }

    def open_file(self):
        initial_dir = None
        current_foam = self.foam_manager.get_current_foam_type()
        suggestions = self.foam_manager.get_suggested_paths("DSC", current_foam)
        if suggestions and suggestions.get("input_folder"):
            initial_dir = suggestions.get("input_folder")

        file_path = filedialog.askopenfilename(
            title="Select DSC TXT file",
            initialdir=initial_dir,
            filetypes=[("DSC text files", "*.txt"), ("All files", "*.*")],
        )
        if not file_path:
            return
        self.filepath_var.set(file_path)
        self._load_file(file_path)

    def reload_current_file(self):
        current = self.filepath_var.get()
        if not current or not os.path.exists(current):
            messagebox.showwarning("No file", "Select a DSC TXT file first.")
            return
        self._load_file(current)

    def _load_file(self, file_path):
        try:
            self.parsed = self.parser.parse_file(file_path)
            self._update_summary()
            self._prepare_phase_controls()
            for phase in (*self.phase_order, self.combined_phase):
                self.redraw_phase(phase)
            self.status_var.set(f"Loaded {os.path.basename(file_path)}")
        except Exception as exc:
            messagebox.showerror("Error", f"Could not parse DSC file:\n{exc}")
            self.status_var.set("Error loading file")

    def _prepare_phase_controls(self):
        self.phase_plot_limits = {}
        shared_limits = self._shared_heating_limits()
        if shared_limits:
            self.phase_plot_limits["1st Heating"] = shared_limits
            self.phase_plot_limits["2nd Heating"] = shared_limits
        cooling_heating_limits = self._shared_cooling_heating_limits()
        if cooling_heating_limits:
            self.phase_plot_limits[self.combined_phase] = cooling_heating_limits

        for phase in self.phase_order:
            segment = self.parsed["segments"].get(phase)
            if segment is None or segment.empty:
                continue
            x_data = segment["Ts"].to_numpy(dtype=float)
            y_data = segment["Value"].to_numpy(dtype=float)
            x_min, x_max = float(np.min(x_data)), float(np.max(x_data))
            y_min, y_max = float(np.min(y_data)), float(np.max(y_data))
            if x_min == x_max:
                x_max = x_min + 1.0
            if y_min == y_max:
                y_max = y_min + 1.0
            if phase not in self.phase_plot_limits:
                x_pad = 0.03 * (x_max - x_min)
                y_pad = 0.08 * (y_max - y_min)
                self.phase_plot_limits[phase] = (x_min - x_pad, x_max + x_pad, y_min - y_pad, y_max + y_pad)

            controls = self.phase_controls[phase]
            controls["temp_x_scale"].configure(from_=x_min, to=x_max)
            controls["temp_y_scale"].configure(from_=y_min, to=y_max)
            controls["cryst_x_scale"].configure(from_=x_min, to=x_max)
            controls["cryst_y_scale"].configure(from_=y_min, to=y_max)

            event = self._event_for_phase(phase)
            if phase in ("1st Heating", "2nd Heating"):
                default_temp_x = x_min + 0.30 * (x_max - x_min)
                default_temp_y = y_min + 0.19 * (y_max - y_min)
                default_cryst_x = x_min + 0.32 * (x_max - x_min)
                default_cryst_y = y_min + 0.09 * (y_max - y_min)
            elif phase == "Cooling":
                default_temp_x = x_min + 0.60 * (x_max - x_min)
                default_temp_y = y_min + 0.84 * (y_max - y_min)
                default_cryst_x = x_min + 0.57 * (x_max - x_min)
                default_cryst_y = y_min + 0.74 * (y_max - y_min)
            else:
                default_temp_x = event["temperature"] if event and event.get("temperature") is not None else x_min + 0.7 * (x_max - x_min)
                default_temp_y = y_min + 0.9 * (y_max - y_min)
                default_cryst_x = x_min + 0.1 * (x_max - x_min)
                default_cryst_y = y_min + 0.1 * (y_max - y_min)

            controls["x_temp"].set(default_temp_x)
            controls["y_temp"].set(default_temp_y)
            controls["x_cryst"].set(default_cryst_x)
            controls["y_cryst"].set(default_cryst_y)
            controls["curve_color"].set(self.phase_colors.get(phase, "#1f77b4"))
            controls["baseline_color"].set(self.phase_colors.get(phase, "#1f77b4"))
            controls["line_width"].set(2.2)
            controls["font_size"].set(28.0)
            controls["invert_x"].set(phase == "Cooling")
            controls["show_baseline"].set(True)
            controls["left_limit"].set("" if not event or event.get("left_limit") is None else f"{event['left_limit']:.3f}")
            controls["right_limit"].set("" if not event or event.get("right_limit") is None else f"{event['right_limit']:.3f}")

        # Keep 2nd Heating text coordinates identical to 1st Heating by default
        src = self.phase_controls.get("1st Heating")
        dst = self.phase_controls.get("2nd Heating")
        if src and dst:
            dst["x_temp"].set(src["x_temp"].get())
            dst["y_temp"].set(src["y_temp"].get())
            dst["x_cryst"].set(src["x_cryst"].get())
            dst["y_cryst"].set(src["y_cryst"].get())

        combined_controls = self.phase_controls.get(self.combined_phase)
        if combined_controls and cooling_heating_limits:
            x0, x1, y0, y1 = cooling_heating_limits
            x_span = x1 - x0
            y_span = y1 - y0
            combined_controls["cooling_arrow_x_scale"].configure(from_=x0, to=x1)
            combined_controls["cooling_arrow_y_scale"].configure(from_=y0, to=y1)
            combined_controls["heating_arrow_x_scale"].configure(from_=x0, to=x1)
            combined_controls["heating_arrow_y_scale"].configure(from_=y0, to=y1)
            combined_controls["exo_x_scale"].configure(from_=x0, to=x1)
            combined_controls["exo_y_scale"].configure(from_=y0, to=y1)
            combined_controls["cooling_color"].set("#0072b2")
            combined_controls["heating_color"].set("#d55e00")
            combined_controls["line_width"].set(2.2)
            combined_controls["label_font_size"].set(28.0)
            combined_controls["arrow_font_size"].set(20.0)
            combined_controls["invert_x"].set(False)
            combined_controls["show_cooling_arrow"].set(True)
            combined_controls["show_heating_arrow"].set(True)
            combined_controls["cooling_arrow_x"].set(x0 + 0.82 * x_span)
            combined_controls["cooling_arrow_y"].set(y0 + 0.57 * y_span)
            combined_controls["heating_arrow_x"].set(x0 + 0.24 * x_span)
            combined_controls["heating_arrow_y"].set(y0 + 0.35 * y_span)
            combined_controls["exo_x"].set(x0 + 0.12 * x_span)
            combined_controls["exo_y"].set(y0 + 0.06 * y_span)
            combined_controls["tc_x_scale"].configure(from_=x0, to=x1)
            combined_controls["tc_y_scale"].configure(from_=y0, to=y1)
            combined_controls["tm_x_scale"].configure(from_=x0, to=x1)
            combined_controls["tm_y_scale"].configure(from_=y0, to=y1)
            combined_controls["xc2_x_scale"].configure(from_=x0, to=x1)
            combined_controls["xc2_y_scale"].configure(from_=y0, to=y1)
            combined_controls["tc_x"].set(x0 + 0.23 * x_span)
            combined_controls["tc_y"].set(y0 + 0.84 * y_span)
            combined_controls["tm_x"].set(x0 + 0.30 * x_span)
            combined_controls["tm_y"].set(y0 + 0.17 * y_span)
            combined_controls["xc2_x"].set(x0 + 0.33 * x_span)
            combined_controls["xc2_y"].set(y0 + 0.09 * y_span)

    def copy_heating_coordinates(self):
        src = self.phase_controls.get("1st Heating")
        dst = self.phase_controls.get("2nd Heating")
        if not src or not dst:
            return
        dst["x_temp"].set(src["x_temp"].get())
        dst["y_temp"].set(src["y_temp"].get())
        dst["x_cryst"].set(src["x_cryst"].get())
        dst["y_cryst"].set(src["y_cryst"].get())
        self.redraw_phase("2nd Heating")

    def _event_for_phase(self, phase):
        active_results = self._get_active_results()
        if not self.parsed:
            return None
        for event in active_results["events"]:
            if event.get("phase") == phase:
                return event
        return None

    def _shared_heating_limits(self):
        if not self.parsed:
            return None
        seg1 = self.parsed["segments"].get("1st Heating")
        seg2 = self.parsed["segments"].get("2nd Heating")
        if seg1 is None or seg2 is None or seg1.empty or seg2.empty:
            return None

        x_all = np.concatenate([
            seg1["Ts"].to_numpy(dtype=float),
            seg2["Ts"].to_numpy(dtype=float),
        ])
        y_all = np.concatenate([
            seg1["Value"].to_numpy(dtype=float),
            seg2["Value"].to_numpy(dtype=float),
        ])
        x_min, x_max = float(np.min(x_all)), float(np.max(x_all))
        y_min, y_max = float(np.min(y_all)), float(np.max(y_all))
        if x_min == x_max:
            x_max = x_min + 1.0
        if y_min == y_max:
            y_max = y_min + 1.0
        x_pad = 0.03 * (x_max - x_min)
        y_pad = 0.08 * (y_max - y_min)
        return (x_min - x_pad, x_max + x_pad, y_min - y_pad, y_max + y_pad)

    def _shared_cooling_heating_limits(self):
        if not self.parsed:
            return None
        seg_cooling = self.parsed["segments"].get("Cooling")
        seg_heating = self.parsed["segments"].get("1st Heating")
        if seg_cooling is None or seg_heating is None or seg_cooling.empty or seg_heating.empty:
            return None

        x_all = np.concatenate([
            seg_cooling["Ts"].to_numpy(dtype=float),
            seg_heating["Ts"].to_numpy(dtype=float),
        ])
        y_all = np.concatenate([
            seg_cooling["Value"].to_numpy(dtype=float),
            seg_heating["Value"].to_numpy(dtype=float),
        ])
        x_min, x_max = float(np.min(x_all)), float(np.max(x_all))
        y_min, y_max = float(np.min(y_all)), float(np.max(y_all))
        if x_min == x_max:
            x_max = x_min + 1.0
        if y_min == y_max:
            y_max = y_min + 1.0
        x_pad = 0.03 * (x_max - x_min)
        y_pad = 0.08 * (y_max - y_min)
        return (x_min - x_pad, x_max + x_pad, y_min - y_pad, y_max + y_pad)

    def _get_active_results(self):
        if not self.parsed:
            return {"mode": "unknown", "events": []}
        parsed_results = self.parsed["results"]
        selected_mode = self.classification_var.get()
        if selected_mode == "auto":
            return {"mode": parsed_results.get("mode", "unknown"), "events": parsed_results.get("events", [])}
        if selected_mode == "semicrystalline":
            return {"mode": "semicrystalline", "events": parsed_results.get("semicrystalline_events", [])}
        return {"mode": "amorphous", "events": parsed_results.get("amorphous_events", [])}

    def _update_summary(self):
        if not self.parsed:
            self.summary_var.set("No file loaded")
            return
        sample = self.parsed["sample_name"]
        mass = self.parsed["mass_mg"]
        mass_text = f"{mass:.3f} mg" if isinstance(mass, float) else "n/a"
        auto_mode = self.parsed["results"]["mode"]
        active = self._get_active_results()
        self.summary_var.set(
            f"Sample: {sample} | Mass: {mass_text} | Auto: {auto_mode} | Selected: {active['mode']} ({len(active['events'])} events)"
        )

    def on_classification_changed(self):
        self._update_summary()
        if not self.parsed:
            return
        self._prepare_phase_controls()
        for phase in (*self.phase_order, self.combined_phase):
            self.redraw_phase(phase)

    def redraw_phase(self, phase):
        figure = self.phase_figures.get(phase)
        axis = self.phase_axes.get(phase)
        canvas = self.phase_canvases.get(phase)
        if not figure or not axis or not canvas:
            return

        axis.clear()
        axis.set_facecolor("white")

        if phase == self.combined_phase:
            self._redraw_combined_phase(figure, axis, canvas, phase)
            return

        if not self.parsed:
            axis.text(0.5, 0.5, "Load a DSC TXT file", ha="center", va="center", transform=axis.transAxes, fontsize=12)
            self._apply_fixed_layout(figure)
            canvas.draw_idle()
            return

        segment = self.parsed["segments"].get(phase)
        if segment is None or segment.empty:
            axis.text(0.5, 0.5, "No data for this phase", ha="center", va="center", transform=axis.transAxes, fontsize=12)
            self._apply_fixed_layout(figure)
            canvas.draw_idle()
            return

        x_data = segment["Ts"].to_numpy(dtype=float)
        y_data = segment["Value"].to_numpy(dtype=float)
        controls = self.phase_controls[phase]
        color = self.resolve_color(controls["curve_color"].get(), self.phase_colors[phase])
        baseline_color = self.resolve_color(controls["baseline_color"].get(), color)
        line_width = self.read_float_value(controls["line_width"], 2.2, minimum=0.4)
        font_size = self.read_float_value(controls["font_size"], 20.0, minimum=8.0)
        event = self._event_for_phase(phase)
        mode = self._get_active_results()["mode"]

        axis.plot(x_data, y_data, color=color, linewidth=line_width)
        axis.set_xlabel("Temperature (°C)", fontname="DejaVu Sans", fontsize=font_size)
        axis.set_ylabel("Heat Flow ($W \\cdot g^{-1}$)", fontname="DejaVu Sans", fontsize=font_size)
        axis.tick_params(direction="in", top=True, right=True, labelsize=max(8.0, font_size - 2.0))
        axis.grid(False)

        left, right = self.get_integration_limits(controls, event)
        if controls["show_baseline"].get() and left is not None and right is not None and left != right:
            x_sorted_idx = np.argsort(x_data)
            x_sorted = x_data[x_sorted_idx]
            y_sorted = y_data[x_sorted_idx]
            x_left = float(np.clip(left, np.min(x_data), np.max(x_data)))
            x_right = float(np.clip(right, np.min(x_data), np.max(x_data)))
            y_left = float(np.interp(x_left, x_sorted, y_sorted))
            y_right = float(np.interp(x_right, x_sorted, y_sorted))
            axis.plot(
                [x_left, x_right],
                [y_left, y_right],
                linestyle="--",
                dashes=(6, 4),
                linewidth=max(1.2, 0.8 * line_width),
                color=baseline_color,
                zorder=5,
            )

        if event and controls["show_temp"].get() and event.get("temperature") is not None:
            temp_label = event.get("temp_label", "T")
            latex_temp_label = {
                "Tm": "T_{m}",
                "Tc": "T_{c}",
                "Tg": "T_{g}",
            }.get(temp_label, temp_label)
            temp_text = f"${latex_temp_label}={event['temperature']:.1f}$ °C"
            axis.text(
                controls["x_temp"].get(),
                controls["y_temp"].get(),
                temp_text,
                color="#000000",
                fontsize=max(8.0, font_size - 1.0),
                fontname="DejaVu Sans",
            )

        if event and controls["show_cryst"].get() and event.get("crystallinity") is not None:
            if mode == "amorphous":
                cryst_text = f"$\\Delta c_p={event['crystallinity']:.1f}$ J/gK"
            else:
                cryst_text = f"$\\chi_c={event['crystallinity']:.1f}\\%$"
            axis.text(
                controls["x_cryst"].get(),
                controls["y_cryst"].get(),
                cryst_text,
                color="#000000",
                fontsize=max(8.0, font_size - 1.0),
                fontname="DejaVu Sans",
            )

        phase_limits = self.phase_plot_limits.get(phase)
        if phase_limits:
            x0, x1, y0, y1 = phase_limits
            axis.set_xlim(x0, x1)
            axis.set_ylim(y0, y1)
        if controls["invert_x"].get():
            axis.invert_xaxis()

        self._apply_fixed_layout(figure)
        canvas.draw_idle()

    def _redraw_combined_phase(self, figure, axis, canvas, phase):
        if not self.parsed:
            axis.text(0.5, 0.5, "Load a DSC TXT file", ha="center", va="center", transform=axis.transAxes, fontsize=12)
            self._apply_fixed_layout(figure)
            canvas.draw_idle()
            return

        cooling = self.parsed["segments"].get("Cooling")
        heating = self.parsed["segments"].get("1st Heating")
        if cooling is None or heating is None or cooling.empty or heating.empty:
            axis.text(0.5, 0.5, "No data for Cooling + 1st Heating", ha="center", va="center", transform=axis.transAxes, fontsize=12)
            self._apply_fixed_layout(figure)
            canvas.draw_idle()
            return

        controls = self.phase_controls[phase]
        cooling_color = self.resolve_color(controls["cooling_color"].get(), self.phase_colors["Cooling"])
        heating_color = self.resolve_color(controls["heating_color"].get(), self.phase_colors["1st Heating"])
        line_width = self.read_float_value(controls["line_width"], 2.2, minimum=0.4)
        label_font_size = self.read_float_value(controls["label_font_size"], 28.0, minimum=8.0)
        arrow_font_size = self.read_float_value(controls["arrow_font_size"], 20.0, minimum=8.0)

        x_cool = cooling["Ts"].to_numpy(dtype=float)
        y_cool = cooling["Value"].to_numpy(dtype=float)
        x_heat = heating["Ts"].to_numpy(dtype=float)
        y_heat = heating["Value"].to_numpy(dtype=float)

        axis.plot(x_cool, y_cool, color=cooling_color, linewidth=line_width)
        axis.plot(x_heat, y_heat, color=heating_color, linewidth=line_width)
        axis.set_xlabel("Temperature (°C)", fontname="DejaVu Sans", fontsize=label_font_size)
        axis.set_ylabel("Heat Flow ($W \\cdot g^{-1}$)", fontname="DejaVu Sans", fontsize=label_font_size)
        axis.tick_params(direction="in", top=True, right=True, labelsize=max(8.0, label_font_size - 2.0))
        axis.grid(False)

        phase_limits = self.phase_plot_limits.get(phase)
        if phase_limits:
            x0, x1, y0, y1 = phase_limits
            axis.set_xlim(x0, x1)
            axis.set_ylim(y0, y1)
            x_span = x1 - x0
            y_span = y1 - y0
            arrow_len = 0.10 * x_span
            arrow_text_size = max(8.0, arrow_font_size)

            if controls["show_cooling_arrow"].get():
                cooling_x = controls["cooling_arrow_x"].get()
                cooling_y = controls["cooling_arrow_y"].get()
                axis.annotate(
                    "Cooling",
                    xy=(cooling_x - arrow_len, cooling_y),
                    xytext=(cooling_x, cooling_y),
                    fontsize=arrow_text_size,
                    fontname="DejaVu Sans",
                    color="#000000",
                    ha="left",
                    va="center",
                    arrowprops={"arrowstyle": "->", "lw": 1.1, "color": "#000000"},
                )

            if controls["show_heating_arrow"].get():
                heating_x = controls["heating_arrow_x"].get()
                heating_y = controls["heating_arrow_y"].get()
                axis.annotate(
                    "1st Heating",
                    xy=(heating_x + arrow_len, heating_y),
                    xytext=(heating_x, heating_y),
                    fontsize=arrow_text_size,
                    fontname="DejaVu Sans",
                    color="#000000",
                    ha="right",
                    va="center",
                    arrowprops={"arrowstyle": "->", "lw": 1.1, "color": "#000000"},
                )

            exo_x = controls["exo_x"].get()
            exo_y = controls["exo_y"].get()
            exo_arrow_len = 0.09 * y_span
            axis.annotate(
                "EXO",
                xy=(exo_x, exo_y + exo_arrow_len),
                xytext=(exo_x, exo_y),
                fontsize=max(8.0, arrow_font_size - 1.0),
                fontname="DejaVu Sans",
                color="#000000",
                ha="center",
                va="bottom",
                arrowprops={"arrowstyle": "->", "lw": 1.1, "color": "#000000"},
            )

        cooling_event = self._event_for_phase("Cooling")
        second_event = self._event_for_phase("1st Heating")
        left_cool, right_cool = self.get_integration_limits({}, cooling_event)
        left_heat, right_heat = self.get_integration_limits({}, second_event)
        if cooling_event and left_cool is not None and right_cool is not None and left_cool != right_cool:
            x_sorted_idx = np.argsort(x_cool)
            x_sorted = x_cool[x_sorted_idx]
            y_sorted = y_cool[x_sorted_idx]
            x_left = float(np.clip(left_cool, np.min(x_cool), np.max(x_cool)))
            x_right = float(np.clip(right_cool, np.min(x_cool), np.max(x_cool)))
            y_left = float(np.interp(x_left, x_sorted, y_sorted))
            y_right = float(np.interp(x_right, x_sorted, y_sorted))
            axis.plot([x_left, x_right], [y_left, y_right], linestyle="--", dashes=(6, 4), linewidth=max(1.2, 0.8 * line_width), color=cooling_color, zorder=5)
        if second_event and left_heat is not None and right_heat is not None and left_heat != right_heat:
            x_sorted_idx = np.argsort(x_heat)
            x_sorted = x_heat[x_sorted_idx]
            y_sorted = y_heat[x_sorted_idx]
            x_left = float(np.clip(left_heat, np.min(x_heat), np.max(x_heat)))
            x_right = float(np.clip(right_heat, np.min(x_heat), np.max(x_heat)))
            y_left = float(np.interp(x_left, x_sorted, y_sorted))
            y_right = float(np.interp(x_right, x_sorted, y_sorted))
            axis.plot([x_left, x_right], [y_left, y_right], linestyle="--", dashes=(6, 4), linewidth=max(1.2, 0.8 * line_width), color=heating_color, zorder=5)

        mode = self._get_active_results()["mode"]
        if cooling_event and cooling_event.get("temperature") is not None:
            axis.text(
                controls["tc_x"].get(),
                controls["tc_y"].get(),
                f"$T_{{c}}={cooling_event['temperature']:.1f}$ °C",
                color="#000000",
                fontsize=max(8.0, label_font_size - 1.0),
                fontname="DejaVu Sans",
            )
        if second_event and second_event.get("temperature") is not None:
            second_temp_label = "T_{g}" if second_event.get("temp_label") == "Tg" else "T_{m}"
            axis.text(
                controls["tm_x"].get(),
                controls["tm_y"].get(),
                f"${second_temp_label}={second_event['temperature']:.1f}$ °C",
                color="#000000",
                fontsize=max(8.0, label_font_size - 1.0),
                fontname="DejaVu Sans",
            )
        if second_event and second_event.get("crystallinity") is not None:
            second_metric = (
                f"$\\Delta c_p={second_event['crystallinity']:.1f}$ J/gK"
                if mode == "amorphous"
                else f"$\\chi_c={second_event['crystallinity']:.1f}\\%$"
            )
            axis.text(
                controls["xc2_x"].get(),
                controls["xc2_y"].get(),
                second_metric,
                color="#000000",
                fontsize=max(8.0, label_font_size - 1.0),
                fontname="DejaVu Sans",
            )

        if controls["invert_x"].get():
            axis.invert_xaxis()

        self._apply_fixed_layout(figure)
        canvas.draw_idle()

    def pick_color_for_phase(self, phase, key):
        controls = self.phase_controls.get(phase)
        if not controls or key not in controls:
            return
        initial = controls[key].get()
        chosen = colorchooser.askcolor(color=initial, title=f"Select {phase} {key.replace('_', ' ')} color")
        hex_color = chosen[1]
        if hex_color:
            controls[key].set(hex_color)
            self.redraw_phase(phase)

    def resolve_color(self, proposed, fallback):
        value = (proposed or "").strip()
        if not value:
            return fallback
        try:
            mcolors.to_rgba(value)
            return value
        except Exception:
            return fallback

    def read_float_value(self, variable, fallback, minimum=None):
        try:
            value = float(variable.get())
        except Exception:
            return fallback
        if minimum is not None:
            return max(minimum, value)
        return value

    def get_integration_limits(self, controls, event):
        left_var = controls.get("left_limit") if isinstance(controls, dict) else None
        right_var = controls.get("right_limit") if isinstance(controls, dict) else None
        left = self.read_float_value(left_var, None) if left_var is not None else None
        right = self.read_float_value(right_var, None) if right_var is not None else None
        if left is None and event:
            left = event.get("left_limit")
        if right is None and event:
            right = event.get("right_limit")
        return left, right

    def export_phase_figure(self, phase, ext):
        if not self.parsed:
            messagebox.showwarning("No data", "Load a DSC TXT file first.")
            return
        sample = self.parsed["sample_name"].replace(" ", "_")
        default_name = f"{sample}_{phase.replace(' ', '_')}.{ext}"
        output_path = filedialog.asksaveasfilename(
            title=f"Save {phase} figure",
            defaultextension=f".{ext}",
            initialfile=default_name,
            filetypes=[(ext.upper(), f"*.{ext}"), ("All files", "*.*")],
        )
        if not output_path:
            return

        try:
            self._save_fixed_figure(self.phase_figures[phase], output_path)
            self.status_var.set(f"Saved {os.path.basename(output_path)}")
        except Exception as exc:
            messagebox.showerror("Error", f"Could not save figure:\n{exc}")
            self.status_var.set("Error saving figure")

    def export_all_phases(self, ext):
        if not self.parsed:
            messagebox.showwarning("No data", "Load a DSC TXT file first.")
            return
        output_dir = filedialog.askdirectory(title=f"Select folder to export 4 {ext.upper()} figures")
        if not output_dir:
            return
        sample = self.parsed["sample_name"].replace(" ", "_")
        errors = []
        for phase in (*self.phase_order, self.combined_phase):
            try:
                filename = f"{sample}_{phase.replace(' ', '_')}.{ext}"
                path = os.path.join(output_dir, filename)
                self._save_fixed_figure(self.phase_figures[phase], path)
            except Exception as exc:
                errors.append(f"{phase}: {exc}")

        if errors:
            messagebox.showerror("Export finished with errors", "\n".join(errors))
            self.status_var.set("Exported with errors")
        else:
            self.status_var.set(f"Saved 4 {ext.upper()} figures")
            messagebox.showinfo("Export completed", f"4 {ext.upper()} figures exported to:\n{output_dir}")

    def _save_fixed_figure(self, figure, output_path):
        original_size = tuple(figure.get_size_inches())
        try:
            figure.set_size_inches(*self.export_figsize, forward=True)
            self._apply_fixed_layout(figure)
            figure.canvas.draw()
            figure.savefig(output_path, dpi=self.export_dpi, bbox_inches=None, pad_inches=0)
        finally:
            figure.set_size_inches(*original_size, forward=True)
            self._apply_fixed_layout(figure)
            figure.canvas.draw_idle()

    def _apply_fixed_layout(self, figure):
        figure.subplots_adjust(**self.export_layout)

    def copy_phase_image(self, phase):
        figure = self.phase_figures.get(phase)
        if figure is None:
            return
        original_size = tuple(figure.get_size_inches())
        try:
            figure.set_size_inches(*self.export_figsize, forward=True)
            self._apply_fixed_layout(figure)
            buffer = io.BytesIO()
            figure.savefig(buffer, format="png", dpi=self.export_dpi, bbox_inches=None, pad_inches=0)
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
            self.status_var.set(f"Copied {phase} image to clipboard")
        except Exception as exc:
            messagebox.showerror("Clipboard error", f"Could not copy image:\n{exc}")
        finally:
            figure.set_size_inches(*original_size, forward=True)
            self._apply_fixed_layout(figure)
            figure.canvas.draw_idle()


class DSCImageModule:
    def __init__(self, root):
        self.editor = DSCImageEditor(root)





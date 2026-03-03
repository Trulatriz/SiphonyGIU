import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

import numpy as np
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib import colors as mcolors
from matplotlib.figure import Figure

from .foam_type_manager import FoamTypeManager


class DSCTextParser:
    phase_order = ("1st Heating", "Cooling", "2nd Heating")

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
            raise ValueError("No se encontraron filas de datos numéricos en el fichero.")

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
        if len(data_frame) < 9:
            chunk = max(1, len(data_frame) // 3)
            split_points = [chunk, min(chunk * 2, len(data_frame))]
        else:
            index_values = data_frame["Index"].to_numpy()
            diffs = np.diff(index_values)
            candidate_positions = np.where(diffs > 1)[0]
            if len(candidate_positions) >= 2:
                top = candidate_positions[np.argsort(diffs[candidate_positions])[-2:]]
                split_points = sorted((top + 1).tolist())
            else:
                split_points = [
                    len(data_frame) // 3,
                    (2 * len(data_frame)) // 3,
                ]

        s1 = max(1, min(split_points[0], len(data_frame) - 2))
        s2 = max(s1 + 1, min(split_points[1], len(data_frame) - 1))
        phase_slices = [
            data_frame.iloc[:s1].copy(),
            data_frame.iloc[s1:s2].copy(),
            data_frame.iloc[s2:].copy(),
        ]
        return {phase: segment for phase, segment in zip(self.phase_order, phase_slices)}

    def _parse_results(self, results_body):
        semi_events = self._parse_semicrystalline_events(results_body)
        if semi_events:
            return {"mode": "semicrystalline", "events": semi_events}
        amorphous_events = self._parse_amorphous_events(results_body)
        if amorphous_events:
            return {"mode": "amorphous", "events": amorphous_events}
        return {"mode": "unknown", "events": []}

    def _parse_semicrystalline_events(self, text):
        blocks = re.split(r"(?=Crystallinity\s+-?[\d,\.]+\s*%)", text, flags=re.IGNORECASE)
        candidate_blocks = [block for block in blocks if re.search(r"Crystallinity\s+-?[\d,\.]+\s*%", block, flags=re.IGNORECASE)]
        if not candidate_blocks:
            return []

        events = []
        for idx, phase in enumerate(self.phase_order):
            block = candidate_blocks[idx] if idx < len(candidate_blocks) else ""
            peak = self._first_match_float(block, r"Peak\s+(-?[\d,\.]+)\s*(?:°C|Â°C)")
            crystallinity = self._first_match_float(block, r"Crystallinity\s+(-?[\d,\.]+)\s*%")
            left_limit = self._first_match_float(block, r"Left\s+Limit\s+(-?[\d,\.]+)\s*(?:°C|Â°C)")
            right_limit = self._first_match_float(block, r"Right\s+Limit\s+(-?[\d,\.]+)\s*(?:°C|Â°C)")
            onset = self._first_match_float(block, r"Onset\s+(-?[\d,\.]+)\s*(?:°C|Â°C)")
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
            r"Glass Transition.*?Midpoint ISO\s+(-?[\d,\.]+)\s*(?:°C|Â°C).*?Delta cp\s+(-?[\d,\.]+)\s*Jg\^-1K\^-1",
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
    phase_colors = {
        "1st Heating": "#E69F00",
        "Cooling": "#56B4E9",
        "2nd Heating": "#009E73",
    }

    def __init__(self, root):
        self.root = root
        self.root.title("DSC Image Editor")
        self.root.geometry("1560x980")
        self.root.minsize(1320, 860)

        self.foam_manager = FoamTypeManager()
        self.parser = DSCTextParser()
        self.parsed = None

        self.filepath_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")

        self.phase_axes = {}
        self.phase_figures = {}
        self.phase_canvases = {}
        self.phase_controls = {}

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
        ttk.Button(top, text="Export 3 PNG", command=lambda: self.export_all_phases("png")).grid(row=0, column=4, padx=(8, 0))
        ttk.Button(top, text="Export 3 PDF", command=lambda: self.export_all_phases("pdf")).grid(row=0, column=5, padx=(8, 0))

        self.summary_var = tk.StringVar(value="No file loaded")
        ttk.Label(top, textvariable=self.summary_var).grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=(8, 0))

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        for phase in self.phase_order:
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

        show_temp_var = tk.BooleanVar(value=True)
        show_cryst_var = tk.BooleanVar(value=True)
        curve_color_var = tk.StringVar(value=self.phase_colors.get(phase, "#1f77b4"))
        baseline_color_var = tk.StringVar(value="#555555")
        line_width_var = tk.DoubleVar(value=2.2)
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

        ttk.Label(controls, text="Integration left limit (°C)").grid(row=8, column=0, sticky=tk.W)
        left_entry = ttk.Entry(controls, textvariable=left_limit_var)
        left_entry.grid(row=9, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        left_entry.bind("<KeyRelease>", lambda _e, p=phase: self.redraw_phase(p))

        ttk.Label(controls, text="Integration right limit (°C)").grid(row=10, column=0, sticky=tk.W)
        right_entry = ttk.Entry(controls, textvariable=right_limit_var)
        right_entry.grid(row=11, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        right_entry.bind("<KeyRelease>", lambda _e, p=phase: self.redraw_phase(p))

        ttk.Label(controls, text="Temp label X").grid(row=12, column=0, sticky=tk.W)
        temp_x_scale = ttk.Scale(controls, variable=x_temp_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        temp_x_scale.grid(row=13, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="Temp label Y").grid(row=14, column=0, sticky=tk.W)
        temp_y_scale = ttk.Scale(controls, variable=y_temp_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        temp_y_scale.grid(row=15, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="χc/Δcp label X").grid(row=16, column=0, sticky=tk.W)
        cryst_x_scale = ttk.Scale(controls, variable=x_cryst_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        cryst_x_scale.grid(row=17, column=0, sticky=(tk.W, tk.E), pady=(0, 6))

        ttk.Label(controls, text="χc/Δcp label Y").grid(row=18, column=0, sticky=tk.W)
        cryst_y_scale = ttk.Scale(controls, variable=y_cryst_var, from_=0, to=1, command=lambda _v, p=phase: self.redraw_phase(p))
        cryst_y_scale.grid(row=19, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Button(controls, text="Export PNG", command=lambda p=phase: self.export_phase_figure(p, "png")).grid(row=20, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        ttk.Button(controls, text="Export PDF", command=lambda p=phase: self.export_phase_figure(p, "pdf")).grid(row=21, column=0, sticky=(tk.W, tk.E))

        self.phase_axes[phase] = ax
        self.phase_figures[phase] = fig
        self.phase_canvases[phase] = canvas
        self.phase_controls[phase] = {
            "show_temp": show_temp_var,
            "show_cryst": show_cryst_var,
            "curve_color": curve_color_var,
            "baseline_color": baseline_color_var,
            "line_width": line_width_var,
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
            mode = self.parsed["results"]["mode"]
            sample = self.parsed["sample_name"]
            mass = self.parsed["mass_mg"]
            mass_text = f"{mass:.3f} mg" if isinstance(mass, float) else "n/a"
            self.summary_var.set(f"Sample: {sample} | Mass: {mass_text} | Mode: {mode}")
            self._prepare_phase_controls()
            for phase in self.phase_order:
                self.redraw_phase(phase)
            self.status_var.set(f"Loaded {os.path.basename(file_path)}")
        except Exception as exc:
            messagebox.showerror("Error", f"Could not parse DSC file:\n{exc}")
            self.status_var.set("Error loading file")

    def _prepare_phase_controls(self):
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

            controls = self.phase_controls[phase]
            controls["temp_x_scale"].configure(from_=x_min, to=x_max)
            controls["temp_y_scale"].configure(from_=y_min, to=y_max)
            controls["cryst_x_scale"].configure(from_=x_min, to=x_max)
            controls["cryst_y_scale"].configure(from_=y_min, to=y_max)

            event = self._event_for_phase(phase)
            default_temp_x = event["temperature"] if event and event.get("temperature") is not None else x_min + 0.7 * (x_max - x_min)
            default_temp_y = y_min + 0.9 * (y_max - y_min)
            default_cryst_x = x_min + 0.1 * (x_max - x_min)
            default_cryst_y = y_min + 0.1 * (y_max - y_min)

            controls["x_temp"].set(default_temp_x)
            controls["y_temp"].set(default_temp_y)
            controls["x_cryst"].set(default_cryst_x)
            controls["y_cryst"].set(default_cryst_y)
            controls["curve_color"].set(self.phase_colors.get(phase, "#1f77b4"))
            controls["baseline_color"].set("#555555")
            controls["line_width"].set(2.2)
            controls["left_limit"].set("" if not event or event.get("left_limit") is None else f"{event['left_limit']:.3f}")
            controls["right_limit"].set("" if not event or event.get("right_limit") is None else f"{event['right_limit']:.3f}")

    def _event_for_phase(self, phase):
        if not self.parsed:
            return None
        for event in self.parsed["results"]["events"]:
            if event.get("phase") == phase:
                return event
        return None

    def redraw_phase(self, phase):
        figure = self.phase_figures.get(phase)
        axis = self.phase_axes.get(phase)
        canvas = self.phase_canvases.get(phase)
        if not figure or not axis or not canvas:
            return

        axis.clear()
        axis.set_facecolor("white")

        if not self.parsed:
            axis.text(0.5, 0.5, "Load a DSC TXT file", ha="center", va="center", transform=axis.transAxes, fontsize=12)
            figure.tight_layout()
            canvas.draw_idle()
            return

        segment = self.parsed["segments"].get(phase)
        if segment is None or segment.empty:
            axis.text(0.5, 0.5, "No data for this phase", ha="center", va="center", transform=axis.transAxes, fontsize=12)
            figure.tight_layout()
            canvas.draw_idle()
            return

        x_data = segment["Ts"].to_numpy(dtype=float)
        y_data = segment["Value"].to_numpy(dtype=float)
        controls = self.phase_controls[phase]
        color = self.resolve_color(controls["curve_color"].get(), self.phase_colors[phase])
        baseline_color = self.resolve_color(controls["baseline_color"].get(), "#555555")
        line_width = self.read_float_value(controls["line_width"], 2.2, minimum=0.4)
        event = self._event_for_phase(phase)
        mode = self.parsed["results"]["mode"]

        axis.plot(x_data, y_data, color=color, linewidth=line_width)
        axis.set_xlabel("Temperature (°C)", fontname="DejaVu Serif", fontsize=12)
        axis.set_ylabel("Heat Flow ($W \\cdot g^{-1}$)", fontname="DejaVu Serif", fontsize=12)
        axis.tick_params(direction="in", top=True, right=True, labelsize=10)
        axis.grid(False)
        axis.text(0.02, 0.96, "Exo ↑", transform=axis.transAxes, fontsize=11, fontname="DejaVu Serif", va="top")

        left, right = self.get_integration_limits(controls, event)
        if left is not None and right is not None and left < right:
            left_idx = int(np.argmin(np.abs(x_data - left)))
            right_idx = int(np.argmin(np.abs(x_data - right)))
            axis.plot(
                [x_data[left_idx], x_data[right_idx]],
                [y_data[left_idx], y_data[right_idx]],
                linestyle="--",
                linewidth=max(0.8, 0.6 * line_width),
                color=baseline_color,
            )

        if event and controls["show_temp"].get() and event.get("temperature") is not None:
            temp_label = event.get("temp_label", "T")
            temp_text = f"${temp_label}={event['temperature']:.2f}$ °C"
            axis.text(
                controls["x_temp"].get(),
                controls["y_temp"].get(),
                temp_text,
                color=color,
                fontsize=11,
                fontname="DejaVu Serif",
            )

        if event and controls["show_cryst"].get() and event.get("crystallinity") is not None:
            if mode == "amorphous":
                cryst_text = f"$\\Delta c_p={event['crystallinity']:.3f}$ J/gK"
            else:
                cryst_text = f"$\\chi_c={event['crystallinity']:.2f}\\%$"
            axis.text(
                controls["x_cryst"].get(),
                controls["y_cryst"].get(),
                cryst_text,
                color=color,
                fontsize=11,
                fontname="DejaVu Serif",
            )

        figure.tight_layout()
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
        left = self.read_float_value(controls["left_limit"], None)
        right = self.read_float_value(controls["right_limit"], None)
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
            self.phase_figures[phase].savefig(output_path, dpi=350, bbox_inches="tight")
            self.status_var.set(f"Saved {os.path.basename(output_path)}")
        except Exception as exc:
            messagebox.showerror("Error", f"Could not save figure:\n{exc}")
            self.status_var.set("Error saving figure")

    def export_all_phases(self, ext):
        if not self.parsed:
            messagebox.showwarning("No data", "Load a DSC TXT file first.")
            return
        output_dir = filedialog.askdirectory(title=f"Select folder to export 3 {ext.upper()} figures")
        if not output_dir:
            return
        sample = self.parsed["sample_name"].replace(" ", "_")
        errors = []
        for phase in self.phase_order:
            try:
                filename = f"{sample}_{phase.replace(' ', '_')}.{ext}"
                path = os.path.join(output_dir, filename)
                self.phase_figures[phase].savefig(path, dpi=350, bbox_inches="tight")
            except Exception as exc:
                errors.append(f"{phase}: {exc}")

        if errors:
            messagebox.showerror("Export finished with errors", "\n".join(errors))
            self.status_var.set("Exported with errors")
        else:
            self.status_var.set(f"Saved 3 {ext.upper()} figures")
            messagebox.showinfo("Export completed", f"3 {ext.upper()} figures exported to:\n{output_dir}")


class DSCImageModule:
    def __init__(self, root):
        self.editor = DSCImageEditor(root)

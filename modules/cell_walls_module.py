import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import cv2
import numpy as np
import pandas as pd
from PIL import Image, ImageTk

from .foam_type_manager import FoamTypeManager


@dataclass
class MaskEntry:
    mask_path: Path
    meta_path: Path
    file_id: str
    group_key: str
    microns_per_pixel: float


class CellWallsModule:
    def __init__(self, root):
        self.root = root
        self.foam_manager = FoamTypeManager()
        self.current_foam = self.foam_manager.get_current_foam_type()
        self.entries: list[MaskEntry] = []
        self.exclusions: dict[str, list[tuple[int, int, int, int]]] = {}
        self.inclusions: dict[str, list[tuple[int, int, int, int]]] = {}

        self._build_ui()
        self._load_suggested_paths()

    def _build_ui(self):
        self.root.title("Cell Wall Thickness")
        self.root.geometry("1100x790")

        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        paths = ttk.LabelFrame(main, text="Paths", padding=10)
        paths.pack(fill=tk.X, pady=(0, 8))
        paths.columnconfigure(1, weight=1)

        ttk.Label(paths, text="Input Folder:").grid(row=0, column=0, sticky=tk.W, pady=(0, 6))
        self.input_var = tk.StringVar()
        ttk.Entry(paths, textvariable=self.input_var).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=6, pady=(0, 6))
        ttk.Button(paths, text="Browse", command=self._browse_input).grid(row=0, column=2, pady=(0, 6))

        ttk.Label(paths, text="Output Folder:").grid(row=1, column=0, sticky=tk.W)
        self.output_var = tk.StringVar()
        ttk.Entry(paths, textvariable=self.output_var).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=6)
        ttk.Button(paths, text="Browse", command=self._browse_output).grid(row=1, column=2)

        opts = ttk.LabelFrame(main, text="Options", padding=10)
        opts.pack(fill=tk.X, pady=(0, 8))
        self.hist_mode_var = tk.StringVar(value="bins")
        ttk.Label(opts, text="Histogram mode:").grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(opts, text="Bins", variable=self.hist_mode_var, value="bins").grid(row=0, column=1, sticky=tk.W, padx=(8, 6))
        self.n_bins_var = tk.StringVar(value="256")
        ttk.Entry(opts, textvariable=self.n_bins_var, width=10).grid(row=0, column=2, sticky=tk.W, padx=(0, 12))
        ttk.Radiobutton(opts, text="Bin width (µm)", variable=self.hist_mode_var, value="bin_width").grid(row=0, column=3, sticky=tk.W, padx=(0, 6))
        self.bin_width_var = tk.StringVar(value="2.0")
        ttk.Entry(opts, textvariable=self.bin_width_var, width=10).grid(row=0, column=4, sticky=tk.W, padx=(0, 12))
        ttk.Label(opts, text="x min (µm):").grid(row=0, column=5, sticky=tk.W, padx=(4, 6))
        self.x_min_var = tk.StringVar(value="0.0")
        ttk.Entry(opts, textvariable=self.x_min_var, width=10).grid(row=0, column=6, sticky=tk.W, padx=(0, 12))
        ttk.Label(opts, text="x max (µm):").grid(row=0, column=7, sticky=tk.W, padx=(4, 6))
        self.x_max_var = tk.StringVar(value="")
        ttk.Entry(opts, textvariable=self.x_max_var, width=10).grid(row=0, column=8, sticky=tk.W, padx=(0, 12))
        ttk.Label(opts, text="(Histogram per image + combined by group)").grid(row=0, column=9, sticky=tk.W)

        buttons = ttk.Frame(main)
        buttons.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(buttons, text="Scan Input", command=self.scan_input).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Select All", command=self._select_all_entries).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="Select None", command=self._select_no_entries).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(buttons, text="Edit Crop Rectangles", command=self.open_crop_editor).pack(side=tk.LEFT, padx=8)
        ttk.Button(buttons, text="Run Analysis", command=self.run_analysis).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Open Output Folder", command=self._open_output_folder).pack(side=tk.LEFT, padx=8)

        table_frame = ttk.LabelFrame(main, text="Detected binary masks", padding=8)
        table_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("file_id", "group", "mpp", "mask", "meta", "inc", "exc")
        self.table = ttk.Treeview(table_frame, columns=cols, show="headings", height=18, selectmode="extended")
        self.table.heading("file_id", text="File ID")
        self.table.heading("group", text="Group (*)")
        self.table.heading("mpp", text="um/px")
        self.table.heading("mask", text="Binary mask")
        self.table.heading("meta", text="Metadata")
        self.table.heading("inc", text="Inclusion rects")
        self.table.heading("exc", text="Exclusion rects")
        self.table.column("file_id", width=220)
        self.table.column("group", width=110)
        self.table.column("mpp", width=80, anchor=tk.E)
        self.table.column("mask", width=240)
        self.table.column("meta", width=240)
        self.table.column("inc", width=95, anchor=tk.E)
        self.table.column("exc", width=95, anchor=tk.E)
        self.table.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.table.yview)
        self.table.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(main, textvariable=self.status_var, anchor=tk.W).pack(fill=tk.X, pady=(6, 0))

    def _refresh_table(self):
        for item in self.table.get_children():
            self.table.delete(item)
        for e in self.entries:
            inc = len(self.inclusions.get(e.mask_path.name, []))
            exc = len(self.exclusions.get(e.mask_path.name, []))
            self.table.insert(
                "",
                tk.END,
                values=(e.file_id, e.group_key, f"{e.microns_per_pixel:.6f}", e.mask_path.name, e.meta_path.name, inc, exc),
            )

    def _select_all_entries(self):
        for item in self.table.get_children():
            self.table.selection_add(item)

    def _select_no_entries(self):
        self.table.selection_remove(self.table.selection())

    def _get_entries_for_analysis(self) -> list[MaskEntry]:
        selected = self.table.selection()
        if not selected:
            return []
        by_mask = {e.mask_path.name: e for e in self.entries}
        out: list[MaskEntry] = []
        for item in selected:
            values = self.table.item(item, "values")
            if len(values) >= 4:
                mask_name = str(values[3])
                e = by_mask.get(mask_name)
                if e is not None:
                    out.append(e)
        return out

    def _load_suggested_paths(self):
        suggested = self.foam_manager.get_suggested_paths("CellWall", self.current_foam)
        if suggested.get("input_folder"):
            self.input_var.set(suggested["input_folder"])
        if suggested.get("output_folder"):
            self.output_var.set(suggested["output_folder"])

    def _browse_input(self):
        initial = self.input_var.get() or None
        folder = filedialog.askdirectory(title="Select Cell wall Input folder", initialdir=initial)
        if folder:
            self.input_var.set(folder)

    def _browse_output(self):
        initial = self.output_var.get() or None
        folder = filedialog.askdirectory(title="Select Cell wall Output folder", initialdir=initial)
        if folder:
            self.output_var.set(folder)

    def _open_output_folder(self):
        out = self.output_var.get().strip()
        if not out or not os.path.isdir(out):
            messagebox.showwarning("Output", "Output folder does not exist.")
            return
        os.startfile(out)

    @staticmethod
    def _extract_group_key(file_id: str) -> str:
        stem_normalized = file_id.replace("_", " ")
        m = re.search(r"\d{8}(?:-\d+)?", stem_normalized)
        if m:
            return m.group(0)
        digits = "".join(ch for ch in file_id if ch.isdigit())
        return digits if digits else file_id

    def scan_input(self):
        input_dir = Path(self.input_var.get().strip())
        if not input_dir.exists():
            messagebox.showerror("Input", "Input folder does not exist.")
            return

        found: list[MaskEntry] = []
        errors = []
        for mask_path in sorted(input_dir.glob("*_binary_mask.png")):
            meta_path = mask_path.with_suffix(".meta.json")
            if not meta_path.exists():
                errors.append(f"Missing meta: {meta_path.name}")
                continue
            try:
                with open(meta_path, "r", encoding="utf-8") as f:
                    meta = json.load(f)
            except Exception as exc:
                errors.append(f"Invalid meta {meta_path.name}: {exc}")
                continue

            required = ("file_id", "binary_mask_file", "microns_per_pixel")
            missing = [k for k in required if k not in meta]
            if missing:
                errors.append(f"{meta_path.name}: missing {', '.join(missing)}")
                continue
            if str(meta["binary_mask_file"]) != mask_path.name:
                errors.append(f"{meta_path.name}: binary_mask_file does not match mask filename")
                continue

            try:
                mpp = float(meta["microns_per_pixel"])
            except Exception:
                errors.append(f"{meta_path.name}: invalid microns_per_pixel")
                continue
            if mpp <= 0:
                errors.append(f"{meta_path.name}: microns_per_pixel must be > 0")
                continue

            file_id = str(meta["file_id"])
            group_key = self._extract_group_key(file_id)
            found.append(MaskEntry(mask_path=mask_path, meta_path=meta_path, file_id=file_id, group_key=group_key, microns_per_pixel=mpp))

        self.entries = found
        self._refresh_table()

        self.foam_manager.save_module_paths(
            "CellWall",
            self.current_foam,
            {"input_folder": str(input_dir), "output_folder": self.output_var.get().strip()},
        )

        msg = f"Found {len(self.entries)} valid binary masks."
        if errors:
            msg += f" Skipped {len(errors)} files."
        self.status_var.set(msg)
        if errors:
            messagebox.showwarning("Scan completed with warnings", "\n".join(errors[:20]))

    def open_crop_editor(self):
        if not self.entries:
            messagebox.showinfo("Crop Editor", "Scan input first.")
            return

        top = tk.Toplevel(self.root)
        top.title("Crop editor (inclusion/exclusion rectangles)")
        top.geometry("1360x860")

        header = ttk.Frame(top, padding=8)
        header.pack(fill=tk.X)
        ttk.Label(header, text="Image:").pack(side=tk.LEFT)
        image_names = [e.mask_path.name for e in self.entries]
        selected = self.table.selection()
        selected_mask_name = None
        if selected:
            values = self.table.item(selected[0], "values")
            if len(values) >= 4:
                selected_mask_name = str(values[3])
        if selected_mask_name in image_names:
            initial_idx = image_names.index(selected_mask_name)
        else:
            initial_idx = 0
        image_index = {"idx": initial_idx}
        image_var = tk.StringVar(value=image_names[initial_idx])
        image_map = {e.mask_path.name: e for e in self.entries}
        combo = ttk.Combobox(header, textvariable=image_var, values=list(image_map.keys()), state="readonly", width=58)
        combo.pack(side=tk.LEFT, padx=6)
        ttk.Button(header, text="Previous", command=lambda: step_image(-1)).pack(side=tk.LEFT, padx=(8, 4))
        ttk.Button(header, text="Next", command=lambda: step_image(1)).pack(side=tk.LEFT, padx=(0, 10))

        mode_var = tk.StringVar(value="exclude")
        ttk.Radiobutton(header, text="Add exclusion", variable=mode_var, value="exclude").pack(side=tk.LEFT, padx=(20, 6))
        ttk.Radiobutton(header, text="Add inclusion", variable=mode_var, value="include").pack(side=tk.LEFT)

        counts_var = tk.StringVar(value="")
        ttk.Label(header, textvariable=counts_var).pack(side=tk.RIGHT)

        work = ttk.Frame(top)
        work.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        work.columnconfigure(0, weight=1)
        work.columnconfigure(1, weight=1)
        work.rowconfigure(0, weight=1)

        canvas = tk.Canvas(work, bg="black", width=650, height=680)
        canvas.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        preview_canvas = tk.Canvas(work, bg="black", width=650, height=680)
        preview_canvas.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        btns = ttk.Frame(top, padding=(8, 0, 8, 8))
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="Clear exclusions (this image)", command=lambda: clear_rects("exclude")).pack(side=tk.LEFT)
        ttk.Button(btns, text="Clear inclusions (this image)", command=lambda: clear_rects("include")).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="Apply + Preview", command=lambda: apply_preview(show_info=True)).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text="Close", command=top.destroy).pack(side=tk.RIGHT)
        apply_state_var = tk.StringVar(value="Preview not applied yet")
        ttk.Label(top, textvariable=apply_state_var, anchor=tk.W).pack(fill=tk.X, padx=12, pady=(0, 8))

        state = {
            "pil_img": None,
            "tk_img": None,
            "preview_tk_img": None,
            "scale": 1.0,
            "offx": 0,
            "offy": 0,
            "drag_start": None,
            "drag_item": None,
        }

        def redraw():
            canvas.delete("all")
            entry = image_map[image_var.get()]
            arr = cv2.imread(str(entry.mask_path), cv2.IMREAD_GRAYSCALE)
            if arr is None:
                return
            img = Image.fromarray(arr).convert("RGB")
            cw = max(1, canvas.winfo_width())
            ch = max(1, canvas.winfo_height())
            scale = min(cw / img.width, ch / img.height)
            nw = max(1, int(img.width * scale))
            nh = max(1, int(img.height * scale))
            offx = (cw - nw) // 2
            offy = (ch - nh) // 2
            resized = img.resize((nw, nh), Image.NEAREST)
            tk_img = ImageTk.PhotoImage(resized)
            state["pil_img"] = img
            state["tk_img"] = tk_img
            state["scale"] = scale
            state["offx"] = offx
            state["offy"] = offy
            canvas.create_image(offx, offy, anchor=tk.NW, image=tk_img)

            key = entry.mask_path.name
            for rect in self.exclusions.get(key, []):
                x1, y1, x2, y2 = rect
                canvas.create_rectangle(
                    offx + int(x1 * scale),
                    offy + int(y1 * scale),
                    offx + int(x2 * scale),
                    offy + int(y2 * scale),
                    outline="red",
                    width=2,
                )
            for rect in self.inclusions.get(key, []):
                x1, y1, x2, y2 = rect
                canvas.create_rectangle(
                    offx + int(x1 * scale),
                    offy + int(y1 * scale),
                    offx + int(x2 * scale),
                    offy + int(y2 * scale),
                    outline="lime",
                    width=2,
                )
            counts_var.set(
                f"Inclusions: {len(self.inclusions.get(key, []))}    Exclusions: {len(self.exclusions.get(key, []))}"
            )
            apply_preview(show_info=False)

        def render_preview(arr_preview: np.ndarray):
            preview_canvas.delete("all")
            img = Image.fromarray(arr_preview).convert("RGB")
            cw = max(1, preview_canvas.winfo_width())
            ch = max(1, preview_canvas.winfo_height())
            scale = min(cw / img.width, ch / img.height)
            nw = max(1, int(img.width * scale))
            nh = max(1, int(img.height * scale))
            offx = (cw - nw) // 2
            offy = (ch - nh) // 2
            resized = img.resize((nw, nh), Image.NEAREST)
            state["preview_tk_img"] = ImageTk.PhotoImage(resized)
            preview_canvas.create_image(offx, offy, anchor=tk.NW, image=state["preview_tk_img"])

        def apply_preview(show_info: bool):
            entry = image_map[image_var.get()]
            arr = cv2.imread(str(entry.mask_path), cv2.IMREAD_GRAYSCALE)
            if arr is None:
                return
            solid = (arr > 127).astype(np.uint8)
            h, w = solid.shape
            key = entry.mask_path.name
            include_rects = self.inclusions.get(key, [])
            if include_rects:
                roi = np.zeros_like(solid, dtype=np.uint8)
                for x1, y1, x2, y2 in include_rects:
                    xx1, xx2 = max(0, min(x1, x2)), min(w, max(x1, x2))
                    yy1, yy2 = max(0, min(y1, y2)), min(h, max(y1, y2))
                    roi[yy1:yy2, xx1:xx2] = 1
            else:
                roi = np.ones_like(solid, dtype=np.uint8)
            for x1, y1, x2, y2 in self.exclusions.get(key, []):
                xx1, xx2 = max(0, min(x1, x2)), min(w, max(x1, x2))
                yy1, yy2 = max(0, min(y1, y2)), min(h, max(y1, y2))
                roi[yy1:yy2, xx1:xx2] = 0

            analyzed = ((solid > 0) & (roi > 0)).astype(np.uint8) * 255
            # White solid to be analyzed, black background/excluded.
            render_preview(analyzed)
            apply_state_var.set(
                f"Applied on {entry.mask_path.name}: inclusions={len(include_rects)}, exclusions={len(self.exclusions.get(key, []))}"
            )
            if show_info:
                messagebox.showinfo("Applied", "Inclusions/exclusions applied to preview.")

        def to_img_xy(x, y):
            img = state["pil_img"]
            if img is None:
                return None
            scale = state["scale"]
            offx = state["offx"]
            offy = state["offy"]
            ix = int((x - offx) / scale)
            iy = int((y - offy) / scale)
            ix = max(0, min(img.width - 1, ix))
            iy = max(0, min(img.height - 1, iy))
            return ix, iy

        def on_press(event):
            if state["pil_img"] is None:
                return
            p = to_img_xy(event.x, event.y)
            if p is None:
                return
            state["drag_start"] = p
            outline = "yellow" if mode_var.get() == "exclude" else "cyan"
            state["drag_item"] = canvas.create_rectangle(event.x, event.y, event.x, event.y, outline=outline, width=2)

        def on_move(event):
            if state["drag_item"] is None:
                return
            x1, y1, _, _ = canvas.coords(state["drag_item"])
            canvas.coords(state["drag_item"], x1, y1, event.x, event.y)

        def on_release(event):
            if state["drag_start"] is None:
                return
            p2 = to_img_xy(event.x, event.y)
            if p2 is None:
                return
            x1, y1 = state["drag_start"]
            x2, y2 = p2
            x1, x2 = sorted((x1, x2))
            y1, y2 = sorted((y1, y2))
            if (x2 - x1) >= 2 and (y2 - y1) >= 2:
                key = image_var.get()
                if mode_var.get() == "include":
                    self.inclusions.setdefault(key, []).append((x1, y1, x2, y2))
                else:
                    self.exclusions.setdefault(key, []).append((x1, y1, x2, y2))
            state["drag_start"] = None
            state["drag_item"] = None
            redraw()
            self._refresh_table()
            apply_state_var.set("Changes added. Click 'Apply + Preview' to confirm.")

        def clear_rects(kind: str):
            key = image_var.get()
            if kind == "include":
                self.inclusions[key] = []
            else:
                self.exclusions[key] = []
            redraw()
            self._refresh_table()
            apply_state_var.set("Rectangles cleared for current image.")

        def set_image(name: str):
            if name not in image_map:
                return
            image_var.set(name)
            combo.set(name)
            try:
                image_index["idx"] = image_names.index(name)
            except ValueError:
                image_index["idx"] = 0
            redraw()
            apply_state_var.set(f"Loaded: {name}")

        def step_image(step: int):
            n = len(image_names)
            if n == 0:
                return
            image_index["idx"] = (image_index["idx"] + step) % n
            set_image(image_names[image_index["idx"]])

        def on_combo_selected(_e=None):
            set_image(combo.get())

        combo.bind("<<ComboboxSelected>>", on_combo_selected)
        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_move)
        canvas.bind("<ButtonRelease-1>", on_release)
        canvas.bind("<Configure>", lambda _e: redraw())
        preview_canvas.bind("<Configure>", lambda _e: apply_preview(show_info=False))
        top.after(10, redraw)

    @staticmethod
    def _safe_sheet_name(name: str) -> str:
        name = re.sub(r"[:\\/?*\[\]]", "_", name).strip()
        return name[:31] if len(name) > 31 else name

    @staticmethod
    def _safe_file_token(name: str) -> str:
        token = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(name)).strip("_")
        return token if token else "unnamed"

    @staticmethod
    def _filter_values_by_range(values_um: np.ndarray, x_min_um: float, x_max_um: float | None) -> np.ndarray:
        vals = values_um[np.isfinite(values_um)]
        vals = vals[vals >= float(x_min_um)]
        if x_max_um is not None:
            vals = vals[vals <= float(x_max_um)]
        return vals.astype(np.float32)

    @staticmethod
    def _get_histogram_edges(values_um: np.ndarray, hist_mode: str, hist_value: float) -> np.ndarray:
        vmax = float(np.max(values_um)) if values_um.size > 0 else 1.0
        if vmax <= 0:
            vmax = 1.0

        if hist_mode == "bins":
            n_bins = int(hist_value)
            if n_bins < 2:
                n_bins = 2
            return np.linspace(0.0, vmax, n_bins + 1, dtype=float)

        bin_width_um = float(hist_value)
        if bin_width_um <= 0:
            bin_width_um = 1.0
        upper = max(bin_width_um, np.ceil(vmax / bin_width_um) * bin_width_um)
        edges = np.arange(0.0, upper + bin_width_um, bin_width_um, dtype=float)
        if edges.size < 2:
            edges = np.array([0.0, bin_width_um], dtype=float)
        return edges

    @staticmethod
    def _make_histogram(values_um: np.ndarray, hist_mode: str, hist_value: float) -> tuple[pd.DataFrame, np.ndarray, np.ndarray]:
        if values_um.size == 0:
            empty = pd.DataFrame(columns=["index", "bin start", "bin end", "count", "relative frequency"])
            return empty, np.array([], dtype=float), np.array([], dtype=float)

        edges = CellWallsModule._get_histogram_edges(values_um, hist_mode, hist_value)
        counts, edges = np.histogram(values_um, bins=edges)
        total = int(counts.sum())
        rel = (counts / total) if total > 0 else np.zeros_like(counts, dtype=float)
        df = pd.DataFrame(
            {
                "index": np.arange(1, len(counts) + 1, dtype=int),
                "bin start": edges[:-1],
                "bin end": edges[1:],
                "count": counts.astype(int),
                "relative frequency": rel,
            }
        )
        return df, edges, counts

    @staticmethod
    def _local_thickness_px(solid_mask: np.ndarray) -> np.ndarray:
        solid_bool = solid_mask.astype(bool)
        try:
            import porespy as ps  # type: ignore

            tmap = ps.filters.local_thickness(im=solid_bool, mode="hybrid")
            return np.asarray(tmap, dtype=np.float32)
        except Exception:
            dt = cv2.distanceTransform(solid_mask.astype(np.uint8), cv2.DIST_L2, 5)
            return (2.0 * dt).astype(np.float32)

    @staticmethod
    def _compute_roiwise_thickness_um(analyzed_solid: np.ndarray, microns_per_pixel: float) -> tuple[np.ndarray, np.ndarray]:
        """Compute local thickness independently per connected ROI and merge results.

        Returns:
            t_um_map: float32 map in microns (0 outside analyzed solid)
            values_um: 1D float32 values over analyzed solid pixels
        """
        solid = (analyzed_solid > 0).astype(np.uint8)
        h, w = solid.shape
        t_um_map = np.zeros((h, w), dtype=np.float32)
        values_list: list[np.ndarray] = []

        n_labels, labels, stats, _ = cv2.connectedComponentsWithStats(solid, connectivity=8)
        if n_labels <= 1:
            return t_um_map, np.array([], dtype=np.float32)

        for lab in range(1, n_labels):
            x = int(stats[lab, cv2.CC_STAT_LEFT])
            y = int(stats[lab, cv2.CC_STAT_TOP])
            ww = int(stats[lab, cv2.CC_STAT_WIDTH])
            hh = int(stats[lab, cv2.CC_STAT_HEIGHT])
            if ww <= 0 or hh <= 0:
                continue

            # Tight ROI for this component.
            comp = (labels[y : y + hh, x : x + ww] == lab).astype(np.uint8)
            if int(comp.sum()) == 0:
                continue

            tpx = CellWallsModule._local_thickness_px(comp)
            tum = (tpx * float(microns_per_pixel)).astype(np.float32)
            comp_mask = comp > 0
            t_um_map[y : y + hh, x : x + ww][comp_mask] = tum[comp_mask]
            vals = tum[comp_mask]
            vals = vals[np.isfinite(vals) & (vals >= 0)]
            if vals.size > 0:
                values_list.append(vals.astype(np.float32))

        if not values_list:
            return t_um_map, np.array([], dtype=np.float32)
        return t_um_map, np.concatenate(values_list).astype(np.float32)

    @staticmethod
    def _save_thickness_colormap(t_um: np.ndarray, analyzed_solid: np.ndarray, out_path: Path):
        vis = np.zeros((t_um.shape[0], t_um.shape[1], 3), dtype=np.uint8)
        mask = (analyzed_solid > 0) & np.isfinite(t_um) & (t_um > 0)
        if np.any(mask):
            vals = t_um[mask]
            vmin = float(np.percentile(vals, 1.0))
            vmax = float(np.percentile(vals, 99.5))
            if vmax <= vmin:
                vmin = float(np.min(vals))
                vmax = float(np.max(vals))
            if vmax <= vmin:
                vmax = vmin + 1.0

            # ImageJ-like warm rendering: avoid cyan tones and emphasize magenta/red/yellow.
            norm01 = np.clip((t_um - vmin) / (vmax - vmin), 0.0, 1.0)
            norm01 = np.power(norm01, 0.85)
            norm = (norm01 * 255.0).astype(np.uint8)
            colored = cv2.applyColorMap(norm, cv2.COLORMAP_PLASMA)
            vis[mask] = colored[mask]
        cv2.imwrite(str(out_path), vis)

    @staticmethod
    def _save_histogram_png(
        values_um: np.ndarray,
        hist_mode: str,
        hist_value: float,
        out_png: Path,
        title: str,
        x_min_um: float = 0.0,
        x_max_um: float | None = None,
        kde_bw_adjust: float = 2.4,
    ):
        try:
            import matplotlib.pyplot as plt
            import seaborn as sns
        except Exception:
            return

        vals = CellWallsModule._filter_values_by_range(values_um, x_min_um=x_min_um, x_max_um=x_max_um)
        if vals.size == 0:
            return

        edges = CellWallsModule._get_histogram_edges(vals, hist_mode, hist_value)
        counts, _ = np.histogram(vals, bins=edges)
        if counts.size == 0:
            return

        fig = plt.figure(figsize=(7.0, 5.0), dpi=150)
        ax = fig.add_subplot(111)
        sns.set_style("white")
        sns.histplot(
            vals,
            bins=edges,
            stat="probability",
            kde=True,
            kde_kws={
                "bw_adjust": float(max(0.1, kde_bw_adjust)),
                "cut": 0,
                "clip": (float(x_min_um), float(x_max_um) if x_max_um is not None else np.inf),
            },
            color="#7088ad",
            edgecolor="#2f2f2f",
            alpha=0.65,
            ax=ax,
        )
        ax.set_ylabel("Relative Frequency")
        ax.set_xlabel("Thickness (µm)")
        ax.set_title(title)
        xmax = float(x_max_um) if x_max_um is not None else float(edges[-1])
        ax.set_xlim(float(x_min_um), xmax)
        ax.grid(False)
        fig.tight_layout()
        fig.savefig(out_png, bbox_inches="tight")
        plt.close(fig)

    def _tune_group_histograms(
        self,
        by_group_values: dict[str, np.ndarray],
        initial_bins: int,
        default_x_min_um: float,
        default_x_max_um: float | None,
        default_kde_bw_adjust: float = 2.4,
    ) -> tuple[bool, dict[str, dict[str, float]]]:
        """Show combined histogram tuner one group at a time.

        Returns per-group params:
        {"bins": int, "x_min": float, "x_max": float|None, "kde_bw_adjust": float}
        """
        tuned: dict[str, dict[str, float]] = {}
        groups = sorted(by_group_values.keys())
        if not groups:
            return True, tuned

        try:
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            import seaborn as sns
        except Exception:
            # If plotting stack is missing, keep initial bins for all groups.
            for g in groups:
                tuned[g] = {
                    "bins": int(initial_bins),
                    "x_min": float(default_x_min_um),
                    "x_max": float(default_x_max_um) if default_x_max_um is not None else np.nan,
                    "kde_bw_adjust": float(max(0.1, default_kde_bw_adjust)),
                }
            return True, tuned

        sns.set_style("white")
        for idx, g in enumerate(groups, start=1):
            vals = by_group_values[g]
            if vals.size == 0:
                tuned[g] = {
                    "bins": int(initial_bins),
                    "x_min": float(default_x_min_um),
                    "x_max": float(default_x_max_um) if default_x_max_um is not None else np.nan,
                    "kde_bw_adjust": float(max(0.1, default_kde_bw_adjust)),
                }
                continue

            dlg = tk.Toplevel(self.root)
            dlg.title(f"Histogram tuning {idx}/{len(groups)} - {g}")
            dlg.geometry("940x730")
            dlg.transient(self.root)
            dlg.grab_set()

            top = ttk.Frame(dlg, padding=8)
            top.pack(fill=tk.X)
            ttk.Label(top, text=f"Combined sample: {g}").pack(side=tk.LEFT)
            ttk.Label(top, text="Adjust bins / x-min / x-max, then Accept and Next").pack(side=tk.RIGHT)

            ctrl = ttk.Frame(dlg, padding=(8, 0, 8, 8))
            ctrl.pack(fill=tk.X)
            ttk.Label(ctrl, text="Bins:").pack(side=tk.LEFT)
            bins_var = tk.IntVar(value=int(initial_bins))
            bins_scale = tk.Scale(ctrl, from_=5, to=300, orient=tk.HORIZONTAL, variable=bins_var, length=380)
            bins_scale.pack(side=tk.LEFT, padx=(6, 12))
            kde_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(ctrl, text="KDE", variable=kde_var).pack(side=tk.LEFT)
            ttk.Label(ctrl, text="KDE smooth:").pack(side=tk.LEFT, padx=(12, 4))
            bw_var = tk.StringVar(value=f"{float(max(0.1, default_kde_bw_adjust)):.2f}")
            ttk.Entry(ctrl, textvariable=bw_var, width=6).pack(side=tk.LEFT)
            ttk.Label(ctrl, text="x min (µm):").pack(side=tk.LEFT, padx=(12, 4))
            x_min_var = tk.StringVar(value=f"{float(default_x_min_um):.4f}")
            ttk.Entry(ctrl, textvariable=x_min_var, width=8).pack(side=tk.LEFT)
            ttk.Label(ctrl, text="x max (µm):").pack(side=tk.LEFT, padx=(12, 4))
            max_default = "" if default_x_max_um is None else f"{float(default_x_max_um):.4f}"
            x_max_var = tk.StringVar(value=max_default)
            ttk.Entry(ctrl, textvariable=x_max_var, width=8).pack(side=tk.LEFT)

            fig_host = ttk.Frame(dlg)
            fig_host.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
            state = {"canvas": None}

            def redraw():
                try:
                    cur_x_min = float(x_min_var.get())
                    cur_bw = float(bw_var.get())
                    cur_x_max_txt = x_max_var.get().strip()
                    cur_x_max = float(cur_x_max_txt) if cur_x_max_txt else None
                    if cur_x_min < 0:
                        raise ValueError
                    if cur_bw <= 0:
                        raise ValueError
                    if cur_x_max is not None and cur_x_max <= cur_x_min:
                        raise ValueError
                except Exception:
                    return
                vals_plot = self._filter_values_by_range(vals, cur_x_min, cur_x_max)
                if vals_plot.size == 0:
                    return
                if state["canvas"] is not None:
                    state["canvas"].get_tk_widget().destroy()
                fig = plt.Figure(figsize=(8.6, 5.8), dpi=100)
                ax = fig.add_subplot(111)
                sns.histplot(
                    vals_plot,
                    bins=int(bins_var.get()),
                    stat="probability",
                    kde=bool(kde_var.get()),
                    kde_kws={
                        "bw_adjust": float(max(0.1, cur_bw)),
                        "cut": 0,
                        "clip": (cur_x_min, cur_x_max if cur_x_max is not None else np.inf),
                    },
                    color="#7088ad",
                    edgecolor="#2f2f2f",
                    alpha=0.65,
                    ax=ax,
                )
                ax.set_xlabel("Thickness (µm)")
                ax.set_ylabel("Relative Frequency")
                ax.set_title(f"{g} (combined)")
                ax.set_xlim(cur_x_min, cur_x_max if cur_x_max is not None else float(np.max(vals_plot)))
                ax.grid(False)
                fig.tight_layout()
                state["canvas"] = FigureCanvasTkAgg(fig, master=fig_host)
                state["canvas"].draw()
                state["canvas"].get_tk_widget().pack(fill=tk.BOTH, expand=True)

            result = {"ok": None}

            def accept():
                try:
                    cur_x_min = float(x_min_var.get())
                    cur_bw = float(bw_var.get())
                    cur_x_max_txt = x_max_var.get().strip()
                    cur_x_max = float(cur_x_max_txt) if cur_x_max_txt else None
                    if cur_x_min < 0:
                        raise ValueError
                    if cur_bw <= 0:
                        raise ValueError
                    if cur_x_max is not None and cur_x_max <= cur_x_min:
                        raise ValueError
                except Exception:
                    messagebox.showerror("Histogram tuning", "Invalid histogram settings. Ensure x max > x min, x min >= 0, and KDE smooth > 0.")
                    return
                tuned[g] = {
                    "bins": int(bins_var.get()),
                    "x_min": float(cur_x_min),
                    "x_max": float(cur_x_max) if cur_x_max is not None else np.nan,
                    "kde_bw_adjust": float(max(0.1, cur_bw)),
                }
                result["ok"] = True
                dlg.destroy()

            def cancel():
                result["ok"] = False
                dlg.destroy()

            buttons = ttk.Frame(dlg, padding=(8, 0, 8, 8))
            buttons.pack(fill=tk.X)
            ttk.Button(buttons, text="Refresh", command=redraw).pack(side=tk.LEFT)
            ttk.Button(buttons, text="Accept and Next", command=accept).pack(side=tk.RIGHT)
            ttk.Button(buttons, text="Cancel", command=cancel).pack(side=tk.RIGHT, padx=(0, 8))

            bins_scale.bind("<ButtonRelease-1>", lambda _e: redraw())
            bw_var.trace_add("write", lambda *_a: redraw())
            x_min_var.trace_add("write", lambda *_a: redraw())
            x_max_var.trace_add("write", lambda *_a: redraw())
            redraw()
            dlg.protocol("WM_DELETE_WINDOW", cancel)
            self.root.wait_window(dlg)
            if result["ok"] is not True:
                return False, {}

        return True, tuned

    def run_analysis(self):
        if not self.entries:
            messagebox.showerror("Run", "No valid entries. Scan input first.")
            return
        selected_entries = self._get_entries_for_analysis()
        if not selected_entries:
            messagebox.showerror("Run", "Select one or more images in the table before running analysis.")
            return

        out_dir = Path(self.output_var.get().strip())
        if not out_dir:
            messagebox.showerror("Run", "Output folder is required.")
            return
        out_dir.mkdir(parents=True, exist_ok=True)

        try:
            hist_mode = self.hist_mode_var.get().strip()
            if hist_mode not in {"bins", "bin_width"}:
                raise ValueError
            if hist_mode == "bins":
                hist_value = int(self.n_bins_var.get())
                if hist_value <= 1:
                    raise ValueError
            else:
                hist_value = float(self.bin_width_var.get())
                if hist_value <= 0:
                    raise ValueError
            x_min_um = float(self.x_min_var.get())
            if x_min_um < 0:
                raise ValueError
            x_max_txt = self.x_max_var.get().strip()
            x_max_um = float(x_max_txt) if x_max_txt else None
            if x_max_um is not None and x_max_um <= x_min_um:
                raise ValueError
        except Exception:
            if self.hist_mode_var.get().strip() == "bins":
                messagebox.showerror("Run", "Bins must be an integer > 1.")
            else:
                messagebox.showerror("Run", "Bin width must be a positive number.")
            return

        # Keep a numeric value for downstream helpers.
        try:
            hist_value = float(hist_value)
        except Exception:
            hist_value = 256.0
            hist_mode = "bins"

        try:
            _ = self._get_histogram_edges(np.array([1.0], dtype=float), hist_mode, hist_value)
        except Exception:
            messagebox.showerror("Run", "Invalid histogram configuration.")
            return

        try:
            _ = int(hist_value) if hist_mode == "bins" else float(hist_value)
        except Exception:
            messagebox.showerror("Run", "Invalid histogram value.")
            return

        try:
            _ = hist_mode
        except Exception:
            messagebox.showerror("Run", "Invalid histogram mode.")
            return

        try:
            _ = hist_value
        except Exception:
            messagebox.showerror("Run", "Invalid histogram parameter.")
            return
        try:
            _ = x_min_um
        except Exception:
            messagebox.showerror("Run", "x min must be a valid number.")
            return
        try:
            _ = x_max_um
        except Exception:
            messagebox.showerror("Run", "x max must be a valid number or empty.")
            return

        try:
            if hist_mode == "bins" and int(hist_value) <= 1:
                raise ValueError
            if hist_mode == "bin_width" and float(hist_value) <= 0:
                raise ValueError
            if float(x_min_um) < 0:
                raise ValueError
            if x_max_um is not None and float(x_max_um) <= float(x_min_um):
                raise ValueError
        except Exception:
            messagebox.showerror("Run", "Histogram configuration is not valid.")
            return

        thickness_dir = out_dir / "local_thickness_maps"
        hist_png_dir = out_dir / "histograms_png"
        thickness_dir.mkdir(parents=True, exist_ok=True)
        hist_png_dir.mkdir(parents=True, exist_ok=True)

        self.foam_manager.save_module_paths(
            "CellWall",
            self.current_foam,
            {"input_folder": self.input_var.get().strip(), "output_folder": str(out_dir)},
        )

        by_group: dict[str, list[tuple[MaskEntry, np.ndarray, Path]]] = {}
        errors = []

        for entry in selected_entries:
            arr = cv2.imread(str(entry.mask_path), cv2.IMREAD_GRAYSCALE)
            if arr is None:
                errors.append(f"Could not read mask: {entry.mask_path.name}")
                continue

            solid = (arr > 127).astype(np.uint8)
            key = entry.mask_path.name
            h, w = solid.shape

            include_rects = self.inclusions.get(key, [])
            if include_rects:
                roi = np.zeros_like(solid, dtype=np.uint8)
                for x1, y1, x2, y2 in include_rects:
                    xx1, xx2 = max(0, min(x1, x2)), min(w, max(x1, x2))
                    yy1, yy2 = max(0, min(y1, y2)), min(h, max(y1, y2))
                    roi[yy1:yy2, xx1:xx2] = 1
            else:
                roi = np.ones_like(solid, dtype=np.uint8)

            for x1, y1, x2, y2 in self.exclusions.get(key, []):
                xx1, xx2 = max(0, min(x1, x2)), min(w, max(x1, x2))
                yy1, yy2 = max(0, min(y1, y2)), min(h, max(y1, y2))
                roi[yy1:yy2, xx1:xx2] = 0

            analyzed_solid = ((solid > 0) & (roi > 0)).astype(np.uint8)
            if int(analyzed_solid.sum()) == 0:
                errors.append(f"No analyzed solid pixels after crop masks: {entry.mask_path.name}")
                continue

            tum, values = self._compute_roiwise_thickness_um(analyzed_solid, float(entry.microns_per_pixel))
            if values.size == 0:
                errors.append(f"No valid thickness values: {entry.mask_path.name}")
                continue

            map_path = thickness_dir / f"{entry.mask_path.stem}_local_thickness.png"
            self._save_thickness_colormap(tum, analyzed_solid, map_path)

            by_group.setdefault(entry.group_key, []).append((entry, values.astype(np.float32), map_path))

        # Tune bins on combined histograms, one sample(group) at a time.
        initial_bins = int(hist_value) if hist_mode == "bins" else 256
        by_group_values = {
            g: (np.concatenate([vals for _, vals, _ in items]) if items else np.array([], dtype=np.float32))
            for g, items in by_group.items()
        }
        ok_tune, tuned_params = self._tune_group_histograms(
            by_group_values,
            initial_bins=initial_bins,
            default_x_min_um=float(x_min_um),
            default_x_max_um=x_max_um,
            default_kde_bw_adjust=3.0,
        )
        if not ok_tune:
            self.status_var.set("Analysis cancelled during histogram tuning.")
            return

        generated = 0
        for group_key, items in by_group.items():
            excel_path = out_dir / f"histogram_{group_key}.xlsx"
            group_hist_mode = "bins"
            group_cfg = tuned_params.get(group_key, {"bins": float(initial_bins), "x_min": float(x_min_um), "x_max": np.nan, "kde_bw_adjust": 3.0})
            group_hist_value = float(group_cfg.get("bins", initial_bins))
            group_x_min = float(group_cfg.get("x_min", x_min_um))
            gxmax = group_cfg.get("x_max", np.nan)
            group_x_max = None if (gxmax is None or (isinstance(gxmax, float) and np.isnan(gxmax))) else float(gxmax)
            group_kde_bw = float(max(0.1, float(group_cfg.get("kde_bw_adjust", 3.0))))
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                all_values = []
                images_rows = []

                for entry, values, map_path in items:
                    vsel = self._filter_values_by_range(values, group_x_min, group_x_max)
                    all_values.append(vsel)
                    df_img, _, _ = self._make_histogram(vsel, group_hist_mode, group_hist_value)
                    sheet = self._safe_sheet_name(f"{entry.file_id}_log")
                    df_img.to_excel(writer, sheet_name=sheet, index=False)
                    hist_path = hist_png_dir / f"hist_cellwall{self._safe_file_token(entry.file_id)}.png"
                    self._save_histogram_png(
                        vsel,
                        group_hist_mode,
                        group_hist_value,
                        hist_path,
                        title=entry.file_id,
                        x_min_um=group_x_min,
                        x_max_um=group_x_max,
                        kde_bw_adjust=group_kde_bw,
                    )
                    images_rows.append(
                        {
                            "file_id": entry.file_id,
                            "group_key": entry.group_key,
                            "binary_mask_file": entry.mask_path.name,
                            "microns_per_pixel": entry.microns_per_pixel,
                            "n_values": int(values.size),
                            "inclusions": len(self.inclusions.get(entry.mask_path.name, [])),
                            "exclusions": len(self.exclusions.get(entry.mask_path.name, [])),
                            "local_thickness_map_png": str(map_path.name),
                            "histogram_png": str(hist_path.name),
                            "hist_bins_used": int(group_hist_value),
                            "hist_x_min_um": float(group_x_min),
                            "hist_x_max_um": (float(group_x_max) if group_x_max is not None else ""),
                            "hist_kde_bw_adjust": group_kde_bw,
                        }
                    )

                combined = np.concatenate(all_values) if all_values else np.array([], dtype=np.float32)
                df_combined, _, _ = self._make_histogram(combined, group_hist_mode, group_hist_value)
                df_combined.to_excel(writer, sheet_name=self._safe_sheet_name(f"histogram_{group_key}"), index=False)
                pd.DataFrame(images_rows).to_excel(writer, sheet_name=self._safe_sheet_name("images_used"), index=False)
                # Combined PNG per group, named with same label logic as histogram_<label>.xlsx
                combined_png = hist_png_dir / f"hist_cellwall{self._safe_file_token(group_key)}.png"
                self._save_histogram_png(
                    combined,
                    group_hist_mode,
                    group_hist_value,
                    combined_png,
                    title=f"{group_key} (combined)",
                    x_min_um=group_x_min,
                    x_max_um=group_x_max,
                    kde_bw_adjust=group_kde_bw,
                )
            generated += 1

        status = f"Completed. Generated {generated} workbook(s) in {out_dir}."
        if errors:
            status += f" Warnings: {len(errors)}."
            messagebox.showwarning("Cell wall analysis warnings", "\n".join(errors[:30]))
        self.status_var.set(status)
        self._refresh_table()

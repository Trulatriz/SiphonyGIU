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
        ttk.Label(opts, text="Number of bins:").grid(row=0, column=0, sticky=tk.W)
        self.n_bins_var = tk.StringVar(value="256")
        ttk.Entry(opts, textvariable=self.n_bins_var, width=12).grid(row=0, column=1, sticky=tk.W, padx=(6, 12))
        ttk.Label(opts, text="(Histogram per image + combined by group)").grid(row=0, column=2, sticky=tk.W)

        buttons = ttk.Frame(main)
        buttons.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(buttons, text="Scan Input", command=self.scan_input).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Edit Crop Rectangles", command=self.open_crop_editor).pack(side=tk.LEFT, padx=8)
        ttk.Button(buttons, text="Run Analysis", command=self.run_analysis).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Open Output Folder", command=self._open_output_folder).pack(side=tk.LEFT, padx=8)

        table_frame = ttk.LabelFrame(main, text="Detected binary masks", padding=8)
        table_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("file_id", "group", "mpp", "mask", "meta", "inc", "exc")
        self.table = ttk.Treeview(table_frame, columns=cols, show="headings", height=18)
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
        top.geometry("1080x820")

        header = ttk.Frame(top, padding=8)
        header.pack(fill=tk.X)
        ttk.Label(header, text="Image:").pack(side=tk.LEFT)
        image_var = tk.StringVar(value=self.entries[0].mask_path.name)
        image_map = {e.mask_path.name: e for e in self.entries}
        combo = ttk.Combobox(header, textvariable=image_var, values=list(image_map.keys()), state="readonly", width=58)
        combo.pack(side=tk.LEFT, padx=6)

        mode_var = tk.StringVar(value="exclude")
        ttk.Radiobutton(header, text="Add exclusion", variable=mode_var, value="exclude").pack(side=tk.LEFT, padx=(20, 6))
        ttk.Radiobutton(header, text="Add inclusion", variable=mode_var, value="include").pack(side=tk.LEFT)

        counts_var = tk.StringVar(value="")
        ttk.Label(header, textvariable=counts_var).pack(side=tk.RIGHT)

        canvas = tk.Canvas(top, bg="black", width=1020, height=680)
        canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        btns = ttk.Frame(top, padding=(8, 0, 8, 8))
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="Clear exclusions (this image)", command=lambda: clear_rects("exclude")).pack(side=tk.LEFT)
        ttk.Button(btns, text="Clear inclusions (this image)", command=lambda: clear_rects("include")).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="Close", command=top.destroy).pack(side=tk.RIGHT)

        state = {"pil_img": None, "tk_img": None, "scale": 1.0, "offx": 0, "offy": 0, "drag_start": None, "drag_item": None}

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

        def clear_rects(kind: str):
            key = image_var.get()
            if kind == "include":
                self.inclusions[key] = []
            else:
                self.exclusions[key] = []
            redraw()
            self._refresh_table()

        combo.bind("<<ComboboxSelected>>", lambda _e: redraw())
        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_move)
        canvas.bind("<ButtonRelease-1>", on_release)
        canvas.bind("<Configure>", lambda _e: redraw())
        top.after(10, redraw)

    @staticmethod
    def _safe_sheet_name(name: str) -> str:
        name = re.sub(r"[:\\/?*\[\]]", "_", name).strip()
        return name[:31] if len(name) > 31 else name

    @staticmethod
    def _make_histogram(values_um: np.ndarray, n_bins: int) -> tuple[pd.DataFrame, np.ndarray, np.ndarray]:
        if values_um.size == 0:
            empty = pd.DataFrame(columns=["index", "bin start", "bin end", "count", "relative frequency"])
            return empty, np.array([], dtype=float), np.array([], dtype=float)

        vmax = float(np.max(values_um))
        if vmax <= 0:
            vmax = 1.0
        counts, edges = np.histogram(values_um, bins=int(n_bins), range=(0.0, vmax))
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
    def _save_thickness_colormap(t_um: np.ndarray, analyzed_solid: np.ndarray, out_path: Path):
        vis = np.zeros((t_um.shape[0], t_um.shape[1], 3), dtype=np.uint8)
        mask = (analyzed_solid > 0) & np.isfinite(t_um) & (t_um > 0)
        if np.any(mask):
            vals = t_um[mask]
            vmax = float(np.percentile(vals, 99.5))
            if vmax <= 0:
                vmax = float(np.max(vals))
            if vmax <= 0:
                vmax = 1.0
            norm = np.clip((t_um / vmax) * 255.0, 0, 255).astype(np.uint8)
            colored = cv2.applyColorMap(norm, cv2.COLORMAP_TURBO)
            vis[mask] = colored[mask]
        cv2.imwrite(str(out_path), vis)

    @staticmethod
    def _save_histogram_png(values_um: np.ndarray, n_bins: int, out_png: Path, title: str):
        try:
            import matplotlib.pyplot as plt
            from matplotlib import cm
        except Exception:
            return

        if values_um.size == 0:
            return

        counts, edges = np.histogram(values_um, bins=n_bins, range=(0.0, float(np.max(values_um))))
        if counts.size == 0:
            return
        rel = (counts / counts.sum()) if counts.sum() > 0 else np.zeros_like(counts, dtype=float)
        widths = np.diff(edges)
        centers = (edges[:-1] + edges[1:]) / 2.0
        mode_idx = int(np.argmax(counts))
        mode_value = float(centers[mode_idx])
        mean_value = float(np.mean(values_um))
        std_value = float(np.std(values_um))
        min_value = float(np.min(values_um))
        max_value = float(np.max(values_um))
        bin_width = float(widths[0]) if widths.size > 0 else 0.0

        fig = plt.figure(figsize=(7.5, 6.0), dpi=150)
        gs = fig.add_gridspec(3, 1, height_ratios=[6, 0.45, 1.8], hspace=0.16)
        ax = fig.add_subplot(gs[0])
        ax.bar(edges[:-1], rel, width=widths, align="edge", color="lightgray", edgecolor="black", linewidth=0.5)
        ax.set_ylabel("Relative Frequency")
        ax.set_title(title)
        ax.set_xlim(edges[0], edges[-1])

        cax = fig.add_subplot(gs[1])
        gradient = np.linspace(0, 1, 512).reshape(1, -1)
        cax.imshow(gradient, aspect="auto", cmap=cm.turbo, extent=[edges[0], edges[-1], 0, 1])
        cax.set_yticks([])
        cax.set_xlabel("Thickness (um)")
        cax.set_xticks([edges[0], edges[-1]])

        tax = fig.add_subplot(gs[2])
        tax.axis("off")
        left_txt = (
            f"N: {int(values_um.size)}\n"
            f"Mean: {mean_value:.4f}\n"
            f"StdDev: {std_value:.4f}\n"
            f"Bins: {int(n_bins)}"
        )
        right_txt = (
            f"Min: {min_value:.4f}\n"
            f"Max: {max_value:.4f}\n"
            f"Mode: {mode_value:.4f} ({int(counts[mode_idx])})\n"
            f"Bin Width: {bin_width:.4f}"
        )
        tax.text(0.02, 0.95, left_txt, va="top", ha="left", fontsize=9)
        tax.text(0.52, 0.95, right_txt, va="top", ha="left", fontsize=9)
        fig.savefig(out_png, bbox_inches="tight")
        plt.close(fig)

    def run_analysis(self):
        if not self.entries:
            messagebox.showerror("Run", "No valid entries. Scan input first.")
            return

        out_dir = Path(self.output_var.get().strip())
        if not out_dir:
            messagebox.showerror("Run", "Output folder is required.")
            return
        out_dir.mkdir(parents=True, exist_ok=True)

        try:
            n_bins = int(self.n_bins_var.get())
            if n_bins <= 1:
                raise ValueError
        except Exception:
            messagebox.showerror("Run", "Number of bins must be an integer > 1.")
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

        by_group: dict[str, list[tuple[MaskEntry, np.ndarray, Path, Path]]] = {}
        errors = []

        for entry in self.entries:
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

            tpx = self._local_thickness_px(analyzed_solid)
            tum = tpx * float(entry.microns_per_pixel)
            values = tum[(analyzed_solid > 0) & np.isfinite(tum) & (tum > 0)]
            if values.size == 0:
                errors.append(f"No valid thickness values: {entry.mask_path.name}")
                continue

            map_path = thickness_dir / f"{entry.mask_path.stem}_local_thickness.png"
            self._save_thickness_colormap(tum, analyzed_solid, map_path)

            hist_path = hist_png_dir / f"{entry.mask_path.stem}_histogram.png"
            self._save_histogram_png(values, n_bins, hist_path, title=entry.file_id)

            by_group.setdefault(entry.group_key, []).append((entry, values.astype(np.float32), map_path, hist_path))

        generated = 0
        for group_key, items in by_group.items():
            excel_path = out_dir / f"histogram_{group_key}.xlsx"
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                all_values = []
                images_rows = []

                for entry, values, map_path, hist_path in items:
                    all_values.append(values)
                    df_img, _, _ = self._make_histogram(values, n_bins)
                    sheet = self._safe_sheet_name(f"{entry.file_id}_log")
                    df_img.to_excel(writer, sheet_name=sheet, index=False)
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
                        }
                    )

                combined = np.concatenate(all_values) if all_values else np.array([], dtype=np.float32)
                df_combined, _, _ = self._make_histogram(combined, n_bins)
                df_combined.to_excel(writer, sheet_name=self._safe_sheet_name(f"histogram_{group_key}"), index=False)
                pd.DataFrame(images_rows).to_excel(writer, sheet_name=self._safe_sheet_name("images_used"), index=False)
            generated += 1

        status = f"Completed. Generated {generated} workbook(s) in {out_dir}."
        if errors:
            status += f" Warnings: {len(errors)}."
            messagebox.showwarning("Cell wall analysis warnings", "\n".join(errors[:30]))
        self.status_var.set(status)
        self._refresh_table()


import json
import os
import re
from dataclasses import dataclass
from datetime import datetime
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
        # Key: mask filename, Value: list[tuple[x1, y1, x2, y2]]
        self.exclusions: dict[str, list[tuple[int, int, int, int]]] = {}

        self._build_ui()
        self._load_suggested_paths()

    def _build_ui(self):
        self.root.title("Cell Wall Thickness")
        self.root.geometry("1050x760")

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
        ttk.Label(opts, text="Bin width (µm):").grid(row=0, column=0, sticky=tk.W)
        self.bin_width_var = tk.StringVar(value="2.0")
        ttk.Entry(opts, textvariable=self.bin_width_var, width=12).grid(row=0, column=1, sticky=tk.W, padx=(6, 12))
        ttk.Label(opts, text="(Histogram per image + combined by group)").grid(row=0, column=2, sticky=tk.W)

        buttons = ttk.Frame(main)
        buttons.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(buttons, text="Scan Input", command=self.scan_input).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Edit Exclusions", command=self.open_exclusion_editor).pack(side=tk.LEFT, padx=8)
        ttk.Button(buttons, text="Run Analysis", command=self.run_analysis).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Open Output Folder", command=self._open_output_folder).pack(side=tk.LEFT, padx=8)

        table_frame = ttk.LabelFrame(main, text="Detected binary masks", padding=8)
        table_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("file_id", "group", "mpp", "mask", "meta")
        self.table = ttk.Treeview(table_frame, columns=cols, show="headings", height=18)
        self.table.heading("file_id", text="File ID")
        self.table.heading("group", text="Group (*)")
        self.table.heading("mpp", text="µm/px")
        self.table.heading("mask", text="Binary mask")
        self.table.heading("meta", text="Metadata")
        self.table.column("file_id", width=230)
        self.table.column("group", width=120)
        self.table.column("mpp", width=80, anchor=tk.E)
        self.table.column("mask", width=260)
        self.table.column("meta", width=260)
        self.table.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.table.yview)
        self.table.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(main, textvariable=self.status_var, anchor=tk.W).pack(fill=tk.X, pady=(6, 0))

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
        for item in self.table.get_children():
            self.table.delete(item)
        for e in self.entries:
            self.table.insert("", tk.END, values=(e.file_id, e.group_key, f"{e.microns_per_pixel:.6f}", e.mask_path.name, e.meta_path.name))

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

    def open_exclusion_editor(self):
        if not self.entries:
            messagebox.showinfo("Exclusions", "Scan input first.")
            return

        top = tk.Toplevel(self.root)
        top.title("Exclusion editor (rectangles)")
        top.geometry("980x760")

        header = ttk.Frame(top, padding=8)
        header.pack(fill=tk.X)
        ttk.Label(header, text="Image:").pack(side=tk.LEFT)
        image_var = tk.StringVar(value=self.entries[0].mask_path.name)
        image_map = {e.mask_path.name: e for e in self.entries}
        combo = ttk.Combobox(header, textvariable=image_var, values=list(image_map.keys()), state="readonly", width=60)
        combo.pack(side=tk.LEFT, padx=6)

        canvas = tk.Canvas(top, bg="black", width=940, height=620)
        canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        btns = ttk.Frame(top, padding=(8, 0, 8, 8))
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="Clear rectangles (this image)", command=lambda: clear_rects()).pack(side=tk.LEFT)
        ttk.Button(btns, text="Close", command=top.destroy).pack(side=tk.RIGHT)

        state = {
            "pil_img": None,
            "tk_img": None,
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

            for rect in self.exclusions.get(entry.mask_path.name, []):
                x1, y1, x2, y2 = rect
                rx1 = offx + int(x1 * scale)
                ry1 = offy + int(y1 * scale)
                rx2 = offx + int(x2 * scale)
                ry2 = offy + int(y2 * scale)
                canvas.create_rectangle(rx1, ry1, rx2, ry2, outline="red", width=2)

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
            state["drag_item"] = canvas.create_rectangle(event.x, event.y, event.x, event.y, outline="yellow", width=2)

        def on_move(event):
            if state["drag_item"] is None:
                return
            canvas.coords(state["drag_item"], canvas.coords(state["drag_item"])[0], canvas.coords(state["drag_item"])[1], event.x, event.y)

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
                self.exclusions.setdefault(key, []).append((x1, y1, x2, y2))
            state["drag_start"] = None
            state["drag_item"] = None
            redraw()

        def clear_rects():
            key = image_var.get()
            self.exclusions[key] = []
            redraw()

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
    def _histogram_dataframe(values_um: np.ndarray, bin_width_um: float) -> pd.DataFrame:
        if values_um.size == 0:
            return pd.DataFrame(columns=["bin_start_um", "bin_end_um", "bin_center_um", "count", "relative_frequency"])
        vmax = float(np.max(values_um))
        upper = max(bin_width_um, np.ceil(vmax / bin_width_um) * bin_width_um)
        edges = np.arange(0.0, upper + bin_width_um, bin_width_um)
        counts, edges = np.histogram(values_um, bins=edges)
        total = counts.sum()
        rel = (counts / total) if total > 0 else np.zeros_like(counts, dtype=float)
        centers = (edges[:-1] + edges[1:]) / 2.0
        return pd.DataFrame(
            {
                "bin_start_um": edges[:-1],
                "bin_end_um": edges[1:],
                "bin_center_um": centers,
                "count": counts,
                "relative_frequency": rel,
            }
        )

    @staticmethod
    def _local_thickness_px(solid_mask: np.ndarray) -> np.ndarray:
        solid_bool = solid_mask.astype(bool)
        try:
            import porespy as ps  # type: ignore

            tmap = ps.filters.local_thickness(im=solid_bool, mode="hybrid")
            return np.asarray(tmap, dtype=np.float32)
        except Exception:
            # Fallback: 2*distance transform
            dt = cv2.distanceTransform(solid_mask.astype(np.uint8), cv2.DIST_L2, 5)
            return (2.0 * dt).astype(np.float32)

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
            bin_width = float(self.bin_width_var.get())
            if bin_width <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Run", "Bin width must be a positive number.")
            return

        # Save current paths for convenience
        self.foam_manager.save_module_paths(
            "CellWall",
            self.current_foam,
            {"input_folder": self.input_var.get().strip(), "output_folder": str(out_dir)},
        )

        by_group: dict[str, list[tuple[MaskEntry, np.ndarray]]] = {}
        errors = []

        for entry in self.entries:
            arr = cv2.imread(str(entry.mask_path), cv2.IMREAD_GRAYSCALE)
            if arr is None:
                errors.append(f"Could not read mask: {entry.mask_path.name}")
                continue

            solid = (arr > 127).astype(np.uint8)
            for rect in self.exclusions.get(entry.mask_path.name, []):
                x1, y1, x2, y2 = rect
                solid[y1:y2, x1:x2] = 0

            tpx = self._local_thickness_px(solid)
            tum = tpx * float(entry.microns_per_pixel)
            values = tum[(solid > 0) & np.isfinite(tum) & (tum > 0)]
            if values.size == 0:
                errors.append(f"No valid thickness values: {entry.mask_path.name}")
                continue

            by_group.setdefault(entry.group_key, []).append((entry, values.astype(np.float32)))

        generated = 0
        for group_key, items in by_group.items():
            excel_path = out_dir / f"histogram_{group_key}.xlsx"
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                combined_values = []
                for entry, values in items:
                    combined_values.append(values)
                    df_img = self._histogram_dataframe(values, bin_width)
                    sheet = self._safe_sheet_name(f"{entry.file_id}_log")
                    df_img.to_excel(writer, sheet_name=sheet, index=False)

                all_vals = np.concatenate(combined_values) if combined_values else np.array([], dtype=np.float32)
                df_combined = self._histogram_dataframe(all_vals, bin_width)
                sheet_combined = self._safe_sheet_name(f"histogram_{group_key}")
                df_combined.to_excel(writer, sheet_name=sheet_combined, index=False)

                meta_rows = []
                for entry, values in items:
                    meta_rows.append(
                        {
                            "file_id": entry.file_id,
                            "binary_mask_file": entry.mask_path.name,
                            "microns_per_pixel": entry.microns_per_pixel,
                            "n_values": int(values.size),
                            "excluded_rectangles": len(self.exclusions.get(entry.mask_path.name, [])),
                        }
                    )
                pd.DataFrame(meta_rows).to_excel(writer, sheet_name=self._safe_sheet_name("images_used"), index=False)
            generated += 1

        status = f"Completed. Generated {generated} workbook(s) in {out_dir}."
        if errors:
            status += f" Warnings: {len(errors)}."
            messagebox.showwarning("Cell wall analysis warnings", "\n".join(errors[:30]))
        self.status_var.set(status)


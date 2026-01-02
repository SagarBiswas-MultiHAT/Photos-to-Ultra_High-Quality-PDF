import os
import queue
import threading
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

from PIL import Image, ImageOps
from reportlab.lib.pagesizes import A4, letter
from reportlab.pdfgen import canvas

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    import img2pdf  # type: ignore
except Exception:
    img2pdf = None


SUPPORTED_IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".gif",
    ".tif",
    ".tiff",
    ".webp",
}


def _format_bytes(num_bytes: int) -> str:
    value = float(num_bytes)
    for unit in ["B", "KB", "MB", "GB"]:
        if value < 1024:
            return f"{value:.0f} {unit}"
        value /= 1024
    return f"{value:.1f} TB"


def _safe_int(value: str, default: int) -> int:
    try:
        parsed = int(value)
        return parsed
    except Exception:
        return default


def _safe_float(value: Any) -> Optional[float]:
    try:
        return float(value)
    except Exception:
        return None


def _exif_rational_to_float(value: object) -> Optional[float]:
    # EXIF rationals may come as (num, den) tuples.
    if isinstance(value, tuple) and len(value) == 2:
        num = _safe_float(value[0])
        den = _safe_float(value[1])
        if num is None or den in (None, 0.0):
            return None
        return num / den
    return _safe_float(value)


def _detect_image_dpi(image: Image.Image) -> Optional[int]:
    # Try PIL info first
    try:
        dpi_info = image.info.get("dpi")
        if isinstance(dpi_info, tuple) and len(dpi_info) >= 2:
            x = _safe_float(dpi_info[0])
            y = _safe_float(dpi_info[1])
            if x and y and x > 0 and y > 0:
                dpi = int(round((x + y) / 2.0))
                if 72 <= dpi <= 1200:
                    return dpi
    except Exception:
        pass

    # EXIF: XResolution(282), YResolution(283), ResolutionUnit(296)
    try:
        exif = image.getexif()
        xres = _exif_rational_to_float(exif.get(282))
        yres = _exif_rational_to_float(exif.get(283))
        unit = int(exif.get(296, 2))  # 2=inches, 3=cm
        if xres and yres and xres > 0 and yres > 0:
            res = (xres + yres) / 2.0
            if unit == 3:
                res = res * 2.54
            dpi = int(round(res))
            if 72 <= dpi <= 1200:
                return dpi
    except Exception:
        pass

    return None


def _image_filetypes() -> List[Tuple[str, str]]:
    patterns = " ".join([f"*{ext}" for ext in sorted(SUPPORTED_IMAGE_EXTENSIONS)])
    return [("Image files", patterns), ("All files", "*")]


def _ensure_parent_dir(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def _transpose_exif(image: Image.Image) -> Image.Image:
    try:
        return ImageOps.exif_transpose(image)
    except Exception:
        return image


def _apply_exif_orientation(image: Image.Image) -> Tuple[Image.Image, bool]:
    try:
        exif = image.getexif()
        orientation = int(exif.get(274, 1))  # 274 = Orientation
    except Exception:
        orientation = 1
    if orientation != 1:
        return _transpose_exif(image), True
    return image, False


def _flatten_alpha(image: Image.Image, background_rgb=(255, 255, 255)) -> Image.Image:
    if image.mode in ("RGBA", "LA") or (image.mode == "P" and "transparency" in image.info):
        base = Image.new("RGB", image.size, background_rgb)
        rgba = image.convert("RGBA")
        base.paste(rgba, mask=rgba.getchannel("A"))
        return base
    if image.mode != "RGB":
        return image.convert("RGB")
    return image


def _has_transparency(image: Image.Image) -> bool:
    if image.mode in ("RGBA", "LA"):
        return True
    if image.mode == "P" and "transparency" in image.info:
        return True
    return False


@dataclass(frozen=True)
class ExportOptions:
    output_mode: str  # "combined" | "separate"
    combined_pdf_path: Optional[Path]
    output_folder: Optional[Path]

    page_size_mode: str  # "dpi" | "match_pixels" | "a4" | "letter"
    dpi_for_page: int
    margin_points: float

    embed_mode: str  # "keep_original" | "lossless_png" | "jpeg_high"
    jpeg_quality: int

    auto_rotate: bool
    set_metadata: bool
    title: str
    author: str


def _compute_page_size_points(
    image_px: Tuple[int, int],
    *,
    page_size_mode: str,
    dpi_for_page: int,
    fallback_pagesize: Tuple[float, float],
) -> Tuple[float, float]:
    w_px, h_px = image_px
    if page_size_mode == "match_pixels":
        return float(w_px), float(h_px)
    if page_size_mode == "dpi":
        dpi = max(1, dpi_for_page)
        return (w_px / dpi) * 72.0, (h_px / dpi) * 72.0
    return float(fallback_pagesize[0]), float(fallback_pagesize[1])


def _fit_rect_preserve_aspect(
    src_w: float,
    src_h: float,
    dst_w: float,
    dst_h: float,
) -> Tuple[float, float]:
    if src_w <= 0 or src_h <= 0 or dst_w <= 0 or dst_h <= 0:
        return 0.0, 0.0
    scale = min(dst_w / src_w, dst_h / src_h)
    return src_w * scale, src_h * scale


class ImageToPDFApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Photo(s) to Ultra-High Quality PDF")
        self.root.minsize(760, 520)

        self._images: List[Path] = []
        self._worker_thread: Optional[threading.Thread] = None
        self._events: "queue.Queue[tuple]" = queue.Queue()
        self._stop_requested = False

        self._build_ui()
        self._set_idle_state()
        self._announce_quality_backend()
        self._poll_events()

    def _announce_quality_backend(self) -> None:
        # If img2pdf is missing, we can still produce PDFs via reportlab, but
        # img2pdf provides the strongest lossless guarantee.
        if img2pdf is None:
            self.status_var.set("Ready. (Lossless backend not found; using fallback)")
            try:
                messagebox.showwarning(
                    "Lossless engine not installed",
                    "For the best ultra-high quality (lossless) PDF output, install 'img2pdf'.\n\n"
                    "Run: pip install -r requirements.txt",
                )
            except Exception:
                pass
        else:
            self.status_var.set("Ready. (Lossless backend: img2pdf)")

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        container = ttk.Frame(self.root, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        main = ttk.Frame(container)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=0)
        main.rowconfigure(0, weight=1)

        # Left pane: file list
        left = ttk.LabelFrame(main, text="Selected photos", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)

        list_frame = ttk.Frame(left)
        list_frame.grid(row=0, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(list_frame, activestyle="dotbox", selectmode=tk.EXTENDED)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=scroll.set)

        buttons = ttk.Frame(left)
        buttons.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        for i in range(6):
            buttons.columnconfigure(i, weight=1)

        ttk.Button(buttons, text="Add…", command=self.add_images).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="Remove", command=self.remove_selected).grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="Up", command=lambda: self.move_selected(-1)).grid(row=0, column=2, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="Down", command=lambda: self.move_selected(1)).grid(row=0, column=3, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="Clear", command=self.clear_all).grid(row=0, column=4, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="Sort A→Z", command=self.sort_images).grid(row=0, column=5, sticky="ew")

        tools = ttk.Frame(left)
        tools.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        tools.columnconfigure(0, weight=1)
        ttk.Button(tools, text="PDF → Photos…", command=self.open_pdf_to_photos).grid(row=0, column=0, sticky="w")

        self.summary_label = ttk.Label(left, text="No files selected.")
        self.summary_label.grid(row=3, column=0, sticky="w", pady=(8, 0))

        # Right pane: options
        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="ns")

        export = ttk.LabelFrame(right, text="Export", padding=10)
        export.grid(row=0, column=0, sticky="ew")
        export.columnconfigure(0, weight=1)

        self.output_mode = tk.StringVar(value="combined")
        mode_row = ttk.Frame(export)
        mode_row.grid(row=0, column=0, sticky="ew")
        ttk.Radiobutton(mode_row, text="Single PDF (multi-page)", variable=self.output_mode, value="combined", command=self._refresh_output_fields).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(mode_row, text="One PDF per photo", variable=self.output_mode, value="separate", command=self._refresh_output_fields).grid(row=1, column=0, sticky="w")

        self.combined_path_var = tk.StringVar(value="")
        combined_row = ttk.Frame(export)
        combined_row.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        combined_row.columnconfigure(0, weight=1)
        ttk.Label(combined_row, text="Output file:").grid(row=0, column=0, sticky="w")
        self.combined_entry = ttk.Entry(combined_row, textvariable=self.combined_path_var, width=42)
        self.combined_entry.grid(row=1, column=0, sticky="ew", padx=(0, 6))
        self.combined_browse = ttk.Button(combined_row, text="Browse…", command=self.browse_output_file)
        self.combined_browse.grid(row=1, column=1, sticky="ew")

        self.output_folder_var = tk.StringVar(value="")
        folder_row = ttk.Frame(export)
        folder_row.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        folder_row.columnconfigure(0, weight=1)
        ttk.Label(folder_row, text="Output folder:").grid(row=0, column=0, sticky="w")
        self.folder_entry = ttk.Entry(folder_row, textvariable=self.output_folder_var, width=42)
        self.folder_entry.grid(row=1, column=0, sticky="ew", padx=(0, 6))
        self.folder_browse = ttk.Button(folder_row, text="Browse…", command=self.browse_output_folder)
        self.folder_browse.grid(row=1, column=1, sticky="ew")

        quality = ttk.LabelFrame(right, text="Quality & page sizing", padding=10)
        quality.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        quality.columnconfigure(0, weight=1)

        self.page_size_mode = tk.StringVar(value="auto_dpi")
        ttk.Label(quality, text="Page size:").grid(row=0, column=0, sticky="w")
        self.page_mode_combo = ttk.Combobox(
            quality,
            textvariable=self.page_size_mode,
            values=[
                "auto_dpi",
                "dpi",
                "match_pixels",
                "a4",
                "letter",
            ],
            state="readonly",
            width=20,
        )
        self.page_mode_combo.grid(row=1, column=0, sticky="w")
        self.page_mode_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_page_fields())
        ttk.Label(quality, text="auto_dpi uses photo metadata (fallback to DPI below). dpi = manual.").grid(row=2, column=0, sticky="w", pady=(2, 0))

        dpi_row = ttk.Frame(quality)
        dpi_row.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(dpi_row, text="DPI (fallback/manual):").grid(row=0, column=0, sticky="w")
        self.dpi_var = tk.StringVar(value="300")
        self.dpi_entry = ttk.Entry(dpi_row, textvariable=self.dpi_var, width=8)
        self.dpi_entry.grid(row=0, column=1, sticky="w", padx=(6, 0))
        ttk.Label(dpi_row, text="(Recommended: 300 or 600)").grid(row=0, column=2, sticky="w", padx=(6, 0))

        margin_row = ttk.Frame(quality)
        margin_row.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(margin_row, text="Margin (pt):").grid(row=0, column=0, sticky="w")
        self.margin_var = tk.StringVar(value="0")
        ttk.Entry(margin_row, textvariable=self.margin_var, width=8).grid(row=0, column=1, sticky="w", padx=(6, 0))
        ttk.Label(margin_row, text="(0 = edge-to-edge)").grid(row=0, column=2, sticky="w", padx=(6, 0))

        self.embed_mode = tk.StringVar(value="jpeg_high")
        ttk.Label(quality, text="Embedding:").grid(row=5, column=0, sticky="w", pady=(8, 0))
        embed_combo = ttk.Combobox(
            quality,
            textvariable=self.embed_mode,
            values=[
                "keep_original",  # best for JPEG/PNG (no recompress)
                "lossless_png",   # convert everything to PNG (bigger files)
                "jpeg_high",      # re-encode as high-quality JPEG
            ],
            state="readonly",
            width=20,
        )
        embed_combo.grid(row=6, column=0, sticky="w")

        jpg_row = ttk.Frame(quality)
        jpg_row.grid(row=7, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(jpg_row, text="JPEG quality:").grid(row=0, column=0, sticky="w")
        self.jpeg_quality_var = tk.StringVar(value="100")
        ttk.Entry(jpg_row, textvariable=self.jpeg_quality_var, width=8).grid(row=0, column=1, sticky="w", padx=(6, 0))
        ttk.Label(jpg_row, text="(only used for jpeg_high)").grid(row=0, column=2, sticky="w", padx=(6, 0))

        misc = ttk.LabelFrame(right, text="Options", padding=10)
        misc.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        misc.columnconfigure(0, weight=1)

        self.auto_rotate_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(misc, text="Auto-rotate using EXIF", variable=self.auto_rotate_var).grid(row=0, column=0, sticky="w")

        self.set_metadata_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(misc, text="Set PDF metadata", variable=self.set_metadata_var).grid(row=1, column=0, sticky="w", pady=(6, 0))

        meta = ttk.Frame(misc)
        meta.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        meta.columnconfigure(1, weight=1)
        ttk.Label(meta, text="Title:").grid(row=0, column=0, sticky="w")
        self.title_var = tk.StringVar(value="")
        ttk.Entry(meta, textvariable=self.title_var).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Label(meta, text="Author:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.author_var = tk.StringVar(value="")
        ttk.Entry(meta, textvariable=self.author_var).grid(row=1, column=1, sticky="ew", padx=(6, 0), pady=(6, 0))

        # Bottom: action + progress
        bottom = ttk.Frame(container)
        bottom.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        bottom.columnconfigure(0, weight=1)
        bottom.columnconfigure(1, weight=0)

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(bottom, textvariable=self.status_var).grid(row=0, column=0, sticky="w")

        self.progress = ttk.Progressbar(bottom, mode="determinate", length=240)
        self.progress.grid(row=0, column=1, sticky="e", padx=(10, 0))

        action = ttk.Frame(container)
        action.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        action.columnconfigure(0, weight=1)
        action.columnconfigure(1, weight=0)
        action.columnconfigure(2, weight=0)

        self.convert_button = ttk.Button(action, text="Convert", command=self.start_convert)
        self.convert_button.grid(row=0, column=1, sticky="e", padx=(0, 6))
        self.cancel_button = ttk.Button(action, text="Cancel", command=self.request_cancel)
        self.cancel_button.grid(row=0, column=2, sticky="e")

        self._refresh_output_fields()
        self._refresh_page_fields()

    def open_pdf_to_photos(self) -> None:
        if fitz is None:
            messagebox.showwarning(
                "PDF engine not installed",
                "To convert PDF pages to photos, install 'PyMuPDF'.\n\nRun: pip install -r requirements.txt",
            )
            return
        PDFToPhotosWindow(self.root)

    def _set_idle_state(self) -> None:
        self.cancel_button.state(["disabled"])
        self.progress["value"] = 0
        self.progress["maximum"] = 1
        self._update_summary()

    def _set_busy_state(self) -> None:
        self.cancel_button.state(["!disabled"])

    def _refresh_output_fields(self) -> None:
        mode = self.output_mode.get()
        if mode == "combined":
            self.combined_entry.state(["!disabled"])
            self.combined_browse.state(["!disabled"])
            self.folder_entry.state(["disabled"])
            self.folder_browse.state(["disabled"])
        else:
            self.combined_entry.state(["disabled"])
            self.combined_browse.state(["disabled"])
            self.folder_entry.state(["!disabled"])
            self.folder_browse.state(["!disabled"])

    def _refresh_page_fields(self) -> None:
        mode = self.page_size_mode.get()
        # Only DPI modes use the DPI entry.
        if mode in {"auto_dpi", "dpi"}:
            self.dpi_entry.state(["!disabled"])
        else:
            self.dpi_entry.state(["disabled"])

    def _update_summary(self) -> None:
        if not self._images:
            self.summary_label.configure(text="No files selected.")
            self.convert_button.state(["disabled"])
            return

        total_bytes = 0
        for path in self._images:
            try:
                total_bytes += path.stat().st_size
            except Exception:
                pass

        self.summary_label.configure(text=f"{len(self._images)} file(s) • {_format_bytes(total_bytes)}")
        self.convert_button.state(["!disabled"])

        # Auto-fill default output paths if blank
        if self.output_mode.get() == "combined" and not self.combined_path_var.get().strip():
            first = self._images[0]
            default = first.with_suffix(".pdf")
            self.combined_path_var.set(str(default))
        if self.output_mode.get() == "separate" and not self.output_folder_var.get().strip():
            self.output_folder_var.set(str(self._images[0].parent))

    def _selected_indices(self) -> List[int]:
        return list(self.listbox.curselection())

    def _refresh_listbox(self, preserve_selection: Optional[List[int]] = None) -> None:
        self.listbox.delete(0, tk.END)
        for p in self._images:
            self.listbox.insert(tk.END, str(p))
        if preserve_selection:
            for idx in preserve_selection:
                if 0 <= idx < len(self._images):
                    self.listbox.selection_set(idx)
        self._update_summary()

    def add_images(self) -> None:
        paths = filedialog.askopenfilenames(title="Select photo(s)", filetypes=_image_filetypes())
        if not paths:
            return
        new_paths: List[Path] = []
        for p in paths:
            path = Path(p)
            if path.suffix.lower() not in SUPPORTED_IMAGE_EXTENSIONS:
                continue
            if path not in self._images and path.exists():
                new_paths.append(path)
        if not new_paths:
            return
        self._images.extend(new_paths)
        self._refresh_listbox(preserve_selection=list(range(len(self._images) - len(new_paths), len(self._images))))

    def remove_selected(self) -> None:
        indices = self._selected_indices()
        if not indices:
            return
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self._images):
                del self._images[idx]
        self._refresh_listbox()

    def clear_all(self) -> None:
        self._images.clear()
        self._refresh_listbox()
        self.combined_path_var.set("")

    def sort_images(self) -> None:
        if not self._images:
            return
        self._images.sort(key=lambda p: (p.name.lower(), str(p).lower()))
        self._refresh_listbox()

    def move_selected(self, direction: int) -> None:
        indices = self._selected_indices()
        if not indices:
            return

        if direction < 0:
            ordered = indices
        else:
            ordered = list(reversed(indices))

        for idx in ordered:
            new_idx = idx + direction
            if not (0 <= idx < len(self._images)):
                continue
            if not (0 <= new_idx < len(self._images)):
                continue
            self._images[idx], self._images[new_idx] = self._images[new_idx], self._images[idx]

        new_sel = [max(0, min(len(self._images) - 1, i + direction)) for i in indices]
        self._refresh_listbox(preserve_selection=new_sel)

    def browse_output_file(self) -> None:
        initial = self.combined_path_var.get().strip() or "output.pdf"
        initial_dir = str(Path(initial).parent) if initial else str(Path.home())
        chosen = filedialog.asksaveasfilename(
            title="Save PDF as",
            defaultextension=".pdf",
            initialdir=initial_dir,
            filetypes=[("PDF", "*.pdf")],
        )
        if chosen:
            self.combined_path_var.set(chosen)

    def browse_output_folder(self) -> None:
        initial = self.output_folder_var.get().strip() or str(Path.home())
        chosen = filedialog.askdirectory(title="Choose output folder", initialdir=initial)
        if chosen:
            self.output_folder_var.set(chosen)

    def _build_options(self) -> ExportOptions:
        output_mode = self.output_mode.get()

        combined_path = self.combined_path_var.get().strip()
        output_folder = self.output_folder_var.get().strip()

        combined_pdf_path = Path(combined_path) if combined_path else None
        output_folder_path = Path(output_folder) if output_folder else None

        page_size_mode = self.page_size_mode.get()
        dpi_for_page = max(1, _safe_int(self.dpi_var.get().strip(), 300))
        margin_points = float(_safe_int(self.margin_var.get().strip(), 0))

        embed_mode = self.embed_mode.get()
        jpeg_quality = min(100, max(1, _safe_int(self.jpeg_quality_var.get().strip(), 100)))

        auto_rotate = bool(self.auto_rotate_var.get())
        set_metadata = bool(self.set_metadata_var.get())
        title = self.title_var.get().strip()
        author = self.author_var.get().strip()

        return ExportOptions(
            output_mode=output_mode,
            combined_pdf_path=combined_pdf_path,
            output_folder=output_folder_path,
            page_size_mode=page_size_mode,
            dpi_for_page=dpi_for_page,
            margin_points=margin_points,
            embed_mode=embed_mode,
            jpeg_quality=jpeg_quality,
            auto_rotate=auto_rotate,
            set_metadata=set_metadata,
            title=title,
            author=author,
        )

    def _validate_before_convert(self) -> Optional[ExportOptions]:
        if not self._images:
            messagebox.showwarning("No files", "Please add at least one photo.")
            return None

        options = self._build_options()
        if options.output_mode == "combined":
            if options.combined_pdf_path is None:
                messagebox.showwarning("Output missing", "Please choose an output PDF file.")
                return None
        else:
            if options.output_folder is None:
                messagebox.showwarning("Output missing", "Please choose an output folder.")
                return None

        if options.page_size_mode not in {"auto_dpi", "dpi", "match_pixels", "a4", "letter"}:
            messagebox.showwarning("Invalid setting", "Invalid page size mode.")
            return None
        if options.embed_mode not in {"keep_original", "lossless_png", "jpeg_high"}:
            messagebox.showwarning("Invalid setting", "Invalid embedding mode.")
            return None

        return options

    def start_convert(self) -> None:
        if self._worker_thread and self._worker_thread.is_alive():
            return

        options = self._validate_before_convert()
        if options is None:
            return

        self._stop_requested = False
        self._set_busy_state()
        self.status_var.set("Converting…")
        self.progress["maximum"] = max(1, len(self._images))
        self.progress["value"] = 0

        images = list(self._images)
        self._worker_thread = threading.Thread(
            target=self._convert_worker,
            args=(images, options),
            daemon=True,
        )
        self._worker_thread.start()

    def request_cancel(self) -> None:
        if self._worker_thread and self._worker_thread.is_alive():
            self._stop_requested = True
            self.status_var.set("Cancel requested…")

    def _poll_events(self) -> None:
        try:
            while True:
                event = self._events.get_nowait()
                kind = event[0]
                if kind == "progress":
                    current, total, msg = event[1], event[2], event[3]
                    self.progress["maximum"] = max(1, total)
                    self.progress["value"] = current
                    self.status_var.set(msg)
                elif kind == "done":
                    output_paths = event[1]
                    self._set_idle_state()
                    if output_paths:
                        if len(output_paths) == 1:
                            messagebox.showinfo("Done", f"Saved PDF:\n{output_paths[0]}")
                            self.status_var.set(f"Saved: {output_paths[0]}")
                        else:
                            messagebox.showinfo("Done", f"Saved {len(output_paths)} PDF files.")
                            self.status_var.set(f"Saved {len(output_paths)} PDF files.")
                    else:
                        self.status_var.set("No output written.")
                elif kind == "cancelled":
                    self._set_idle_state()
                    self.status_var.set("Cancelled.")
                    messagebox.showinfo("Cancelled", "Conversion cancelled.")
                elif kind == "error":
                    self._set_idle_state()
                    self.status_var.set("Error.")
                    messagebox.showerror("Error", str(event[1]))
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)

    def _convert_worker(self, images: List[Path], options: ExportOptions) -> None:
        try:
            if img2pdf is not None:
                if options.output_mode == "combined":
                    out = self._export_combined_img2pdf(images, options)
                    self._events.put(("done", [str(out)]))
                else:
                    outs = self._export_separate_img2pdf(images, options)
                    self._events.put(("done", [str(p) for p in outs]))
            else:
                # Fallback to reportlab if img2pdf isn't installed.
                if options.output_mode == "combined":
                    out = self._export_combined(images, options)
                    self._events.put(("done", [str(out)]))
                else:
                    outs = self._export_separate(images, options)
                    self._events.put(("done", [str(p) for p in outs]))
        except _Cancelled:
            self._events.put(("cancelled",))
        except Exception as exc:
            self._events.put(("error", exc))

    def _layout_fun_for_img2pdf(self, options: ExportOptions):
        # img2pdf expects layout_fun(imgwidthpx, imgheightpx, ndpi) ->
        # (pagewidthpt, pageheightpt, imgwidthpt, imgheightpt)
        margin = max(0.0, float(options.margin_points))

        def layout_fun(imgwidthpx: int, imgheightpx: int, ndpi) -> Tuple[float, float, float, float]:
            # ndpi is usually a tuple (xdpi, ydpi)
            page_mode = options.page_size_mode

            # Determine base page size (in points)
            if page_mode in {"a4", "letter"}:
                base_page_w, base_page_h = (A4 if page_mode == "a4" else letter)
                page_w = float(base_page_w)
                page_h = float(base_page_h)
                avail_w = max(1.0, page_w - 2 * margin)
                avail_h = max(1.0, page_h - 2 * margin)
                # Fit image into available area, preserve aspect ratio.
                scale = min(avail_w / max(1.0, float(imgwidthpx)), avail_h / max(1.0, float(imgheightpx)))
                img_w = float(imgwidthpx) * scale
                img_h = float(imgheightpx) * scale
                return page_w, page_h, img_w, img_h

            if page_mode == "match_pixels":
                img_w = float(imgwidthpx)
                img_h = float(imgheightpx)
                return img_w + 2 * margin, img_h + 2 * margin, img_w, img_h

            # dpi / auto_dpi
            dpi = None
            if page_mode == "auto_dpi":
                try:
                    xdpi = _safe_float(ndpi[0]) if ndpi is not None else None
                    ydpi = _safe_float(ndpi[1]) if ndpi is not None else None
                    if xdpi and ydpi and xdpi > 0 and ydpi > 0:
                        dpi = (xdpi + ydpi) / 2.0
                except Exception:
                    dpi = None
                if dpi is None or dpi < 72 or dpi > 1200:
                    dpi = float(options.dpi_for_page)
            else:
                dpi = float(options.dpi_for_page)

            dpi = max(1.0, dpi)
            img_w = (float(imgwidthpx) / dpi) * 72.0
            img_h = (float(imgheightpx) / dpi) * 72.0
            return img_w + 2 * margin, img_h + 2 * margin, img_w, img_h

        return layout_fun

    def _prepare_image_for_img2pdf(self, img_path: Path, options: ExportOptions) -> Tuple[Path, List[Path]]:
        # Returns (path_to_use, temp_files_to_cleanup)
        # For keep_original we prefer passing the file as-is to img2pdf.
        if options.embed_mode == "keep_original":
            return img_path, []

        temp_files: List[Path] = []
        with Image.open(img_path) as im:
            if options.auto_rotate:
                im = ImageOps.exif_transpose(im)

            if options.embed_mode == "jpeg_high":
                # JPEG has no alpha; if present, flatten for best viewer compatibility.
                if _has_transparency(im):
                    im = _flatten_alpha(im)
                im = im.convert("RGB") if im.mode != "RGB" else im
                tmp = tempfile.NamedTemporaryFile(prefix="img2pdf_", suffix=".jpg", delete=False)
                tmp_path = Path(tmp.name)
                tmp.close()
                temp_files.append(tmp_path)
                im.save(
                    str(tmp_path),
                    format="JPEG",
                    quality=options.jpeg_quality,
                    subsampling=0,
                    optimize=True,
                )
                return tmp_path, temp_files

            # lossless_png
            tmp = tempfile.NamedTemporaryFile(prefix="img2pdf_", suffix=".png", delete=False)
            tmp_path = Path(tmp.name)
            tmp.close()
            temp_files.append(tmp_path)
            # Convert to RGB(A) PNG; preserve alpha if present.
            if im.mode not in {"RGB", "RGBA", "L", "LA"}:
                im = im.convert("RGBA" if _has_transparency(im) else "RGB")
            im.save(str(tmp_path), format="PNG", optimize=False)
            return tmp_path, temp_files

    def _cleanup_temp_files(self, temp_files: List[Path]) -> None:
        for p in temp_files:
            try:
                p.unlink(missing_ok=True)
            except Exception:
                pass

    def _dpi_note_for_file(self, img_path: Path, options: ExportOptions) -> str:
        mode = options.page_size_mode
        if mode not in {"auto_dpi", "dpi"}:
            return ""
        if mode == "dpi":
            return f" • DPI: {options.dpi_for_page} (manual)"

        # auto_dpi
        try:
            with Image.open(img_path) as im:
                detected = _detect_image_dpi(im)
        except Exception:
            detected = None

        if detected is None:
            return f" • DPI: {options.dpi_for_page} (fallback)"
        return f" • DPI: {detected} (from photo)"

    def _export_combined_img2pdf(self, images: List[Path], options: ExportOptions) -> Path:
        assert img2pdf is not None
        if options.combined_pdf_path is None:
            raise ValueError("Output PDF path is required")
        _ensure_parent_dir(options.combined_pdf_path)

        self._events.put(("progress", 0, max(1, len(images)), "Preparing images…"))

        prepared: List[Path] = []
        temps: List[Path] = []
        try:
            for idx, img_path in enumerate(images, start=1):
                if self._stop_requested:
                    raise _Cancelled()
                note = self._dpi_note_for_file(img_path, options)
                self._events.put(("progress", idx - 1, len(images), f"Preparing {idx}/{len(images)}: {img_path.name}{note}"))
                p, t = self._prepare_image_for_img2pdf(img_path, options)
                prepared.append(p)
                temps.extend(t)

            self._events.put(("progress", len(images), len(images), "Writing PDF…"))

            rotation_arg = None
            if options.auto_rotate and img2pdf is not None:
                rotation_arg = img2pdf.Rotation.ifvalid

            layout_fun = self._layout_fun_for_img2pdf(options)

            with open(options.combined_pdf_path, "wb") as f:
                # default_dpi is used by img2pdf when input has no DPI metadata.
                img2pdf.convert(
                    prepared,
                    outputstream=f,
                    layout_fun=layout_fun,
                    rotation=rotation_arg,
                    default_dpi=options.dpi_for_page,
                    title=options.title if options.set_metadata and options.title else None,
                    author=options.author if options.set_metadata and options.author else None,
                )

            return options.combined_pdf_path
        finally:
            self._cleanup_temp_files(temps)

    def _export_separate_img2pdf(self, images: List[Path], options: ExportOptions) -> List[Path]:
        assert img2pdf is not None
        if options.output_folder is None:
            raise ValueError("Output folder is required")
        options.output_folder.mkdir(parents=True, exist_ok=True)

        outputs: List[Path] = []
        total = len(images)
        for idx, img_path in enumerate(images, start=1):
            if self._stop_requested:
                raise _Cancelled()
            note = self._dpi_note_for_file(img_path, options)
            self._events.put(("progress", idx - 1, total, f"Preparing {idx}/{total}: {img_path.name}{note}"))

            p, temps = self._prepare_image_for_img2pdf(img_path, options)
            try:
                self._events.put(("progress", idx, total, f"Writing {idx}/{total}: {img_path.stem}.pdf"))

                rotation_arg = None
                if options.auto_rotate and img2pdf is not None:
                    rotation_arg = img2pdf.Rotation.ifvalid
                layout_fun = self._layout_fun_for_img2pdf(options)

                out_path = options.output_folder / f"{img_path.stem}.pdf"
                with open(out_path, "wb") as f:
                    img2pdf.convert(
                        [p],
                        outputstream=f,
                        layout_fun=layout_fun,
                        rotation=rotation_arg,
                        default_dpi=options.dpi_for_page,
                        title=options.title if options.set_metadata and options.title else None,
                        author=options.author if options.set_metadata and options.author else None,
                    )
                outputs.append(out_path)
            finally:
                self._cleanup_temp_files(temps)

        return outputs

    def _export_combined(self, images: List[Path], options: ExportOptions) -> Path:
        if options.combined_pdf_path is None:
            raise ValueError("Output PDF path is required")
        _ensure_parent_dir(options.combined_pdf_path)

        pdf = canvas.Canvas(str(options.combined_pdf_path))
        if options.set_metadata:
            if options.title:
                pdf.setTitle(options.title)
            if options.author:
                pdf.setAuthor(options.author)

        self._render_images_to_canvas(pdf, images, options, overall_total=len(images), progress_offset=0)
        pdf.save()
        return options.combined_pdf_path

    def _export_separate(self, images: List[Path], options: ExportOptions) -> List[Path]:
        if options.output_folder is None:
            raise ValueError("Output folder is required")
        options.output_folder.mkdir(parents=True, exist_ok=True)

        outputs: List[Path] = []
        overall_total = len(images)
        for idx, img_path in enumerate(images, start=1):
            if self._stop_requested:
                raise _Cancelled()
            out_path = options.output_folder / f"{img_path.stem}.pdf"
            pdf = canvas.Canvas(str(out_path))
            if options.set_metadata:
                if options.title:
                    pdf.setTitle(options.title)
                if options.author:
                    pdf.setAuthor(options.author)
            self._render_images_to_canvas(
                pdf,
                [img_path],
                options,
                overall_total=overall_total,
                progress_offset=idx - 1,
            )
            pdf.save()
            outputs.append(out_path)
        return outputs

    def _render_images_to_canvas(
        self,
        pdf: canvas.Canvas,
        images: List[Path],
        options: ExportOptions,
        *,
        overall_total: int,
        progress_offset: int = 0,
    ) -> None:
        temp_paths: List[Path] = []
        try:
            for i, img_path in enumerate(images, start=1):
                if self._stop_requested:
                    raise _Cancelled()

                current = progress_offset + i
                note = self._dpi_note_for_file(img_path, options)
                self._events.put(("progress", current, overall_total, f"Processing {current}/{overall_total}: {img_path.name}{note}"))

                with Image.open(img_path) as im:
                    rotated = False
                    if options.auto_rotate:
                        im, rotated = _apply_exif_orientation(im)

                    # Only flatten/convert when we must (e.g., transparency) to better preserve originals.
                    alpha = _has_transparency(im)
                    processed = False
                    if alpha:
                        im = _flatten_alpha(im)
                        processed = True

                    # Choose PDF page size
                    fallback = A4
                    if options.page_size_mode == "letter":
                        fallback = letter
                    elif options.page_size_mode == "a4":
                        fallback = A4

                    page_mode = options.page_size_mode
                    effective_dpi = options.dpi_for_page
                    if page_mode == "auto_dpi":
                        detected = _detect_image_dpi(im)
                        effective_dpi = detected if detected is not None else options.dpi_for_page
                        # Internally treat as dpi sizing with an automatically chosen DPI.
                        page_mode = "dpi"

                    page_w, page_h = _compute_page_size_points(
                        im.size,
                        page_size_mode=page_mode,
                        dpi_for_page=effective_dpi,
                        fallback_pagesize=fallback,
                    )

                    pdf.setPageSize((page_w, page_h))

                    # Prepare image source for reportlab
                    draw_path = img_path
                    must_rewrite = False
                    suffix = img_path.suffix.lower()
                    if options.embed_mode == "lossless_png":
                        must_rewrite = True
                    elif options.embed_mode == "jpeg_high":
                        must_rewrite = True
                    else:
                        # keep_original: only rewrite if we changed pixels (rotation/alpha) or unsupported for direct embedding
                        if rotated or processed:
                            must_rewrite = True
                        # Some encodings/modes (e.g., CMYK JPEG, palette PNG) can embed poorly; normalize.
                        if im.mode not in {"RGB", "L"}:
                            must_rewrite = True
                        if suffix not in {".jpg", ".jpeg", ".png"}:
                            must_rewrite = True

                    if must_rewrite:
                        if options.embed_mode == "jpeg_high":
                            # JPEG cannot store alpha; if there was alpha we already flattened.
                            tmp = tempfile.NamedTemporaryFile(prefix="img2pdf_", suffix=".jpg", delete=False)
                            tmp_path = Path(tmp.name)
                            tmp.close()
                            temp_paths.append(tmp_path)
                            im_rgb = im if im.mode == "RGB" else im.convert("RGB")
                            im_rgb.save(
                                str(tmp_path),
                                format="JPEG",
                                quality=options.jpeg_quality,
                                subsampling=0,
                                optimize=True,
                            )
                            draw_path = tmp_path
                        else:
                            tmp = tempfile.NamedTemporaryFile(prefix="img2pdf_", suffix=".png", delete=False)
                            tmp_path = Path(tmp.name)
                            tmp.close()
                            temp_paths.append(tmp_path)
                            # Lossless PNG (or keep_original but needs rewrite)
                            im_png = im if im.mode == "RGB" else im.convert("RGB")
                            im_png.save(str(tmp_path), format="PNG", optimize=False)
                            draw_path = tmp_path

                    # Draw with margins + preserve aspect ratio
                    margin = max(0.0, float(options.margin_points))
                    avail_w = max(1.0, page_w - 2 * margin)
                    avail_h = max(1.0, page_h - 2 * margin)
                    img_w_pts, img_h_pts = _fit_rect_preserve_aspect(
                        float(im.width),
                        float(im.height),
                        avail_w,
                        avail_h,
                    )
                    x = margin + (avail_w - img_w_pts) / 2.0
                    y = margin + (avail_h - img_h_pts) / 2.0

                    # NOTE: For keep_original with JPEG/PNG, using file path preserves original encoding (no recompress).
                    pdf.drawImage(str(draw_path), x, y, width=img_w_pts, height=img_h_pts)
                    pdf.showPage()
        finally:
            for p in temp_paths:
                try:
                    p.unlink(missing_ok=True)
                except Exception:
                    pass


class PDFToPhotosWindow:
    def __init__(self, parent: tk.Tk):
        self.window = tk.Toplevel(parent)
        self.window.title("PDF to Photos")
        self.window.minsize(720, 440)

        # Helps keep this window associated with the main app on Windows,
        # and makes focus behavior after file dialogs more predictable.
        try:
            self.window.transient(parent)
        except Exception:
            pass

        self._pdfs: List[Path] = []
        self._worker_thread: Optional[threading.Thread] = None
        self._events: "queue.Queue[tuple]" = queue.Queue()
        self._stop_requested = False

        self._build_ui()
        self._set_idle_state()
        self._poll_events()

        self._bring_to_front()

    def _bring_to_front(self) -> None:
        # File dialogs can steal focus; force this window back on top.
        try:
            self.window.deiconify()
            self.window.lift()
            self.window.focus_force()

            # Temporary topmost toggle is a common Tk workaround on Windows.
            self.window.attributes("-topmost", True)
            self.window.after(10, lambda: self.window.attributes("-topmost", False))
        except Exception:
            pass

    def _build_ui(self) -> None:
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)

        container = ttk.Frame(self.window, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        main = ttk.Frame(container)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=0)
        main.rowconfigure(0, weight=1)

        left = ttk.LabelFrame(main, text="Selected PDFs", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)

        list_frame = ttk.Frame(left)
        list_frame.grid(row=0, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(list_frame, activestyle="dotbox", selectmode=tk.EXTENDED)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=scroll.set)

        buttons = ttk.Frame(left)
        buttons.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        for i in range(4):
            buttons.columnconfigure(i, weight=1)

        ttk.Button(buttons, text="Add…", command=self.add_pdfs).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="Remove", command=self.remove_selected).grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="Clear", command=self.clear_all).grid(row=0, column=2, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="Open output folder", command=self.open_output_folder).grid(row=0, column=3, sticky="ew")

        self.summary_label = ttk.Label(left, text="No files selected.")
        self.summary_label.grid(row=2, column=0, sticky="w", pady=(8, 0))

        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="ns")

        export = ttk.LabelFrame(right, text="Export", padding=10)
        export.grid(row=0, column=0, sticky="ew")
        export.columnconfigure(0, weight=1)

        self.output_folder_var = tk.StringVar(value="")
        folder_row = ttk.Frame(export)
        folder_row.grid(row=0, column=0, sticky="ew")
        folder_row.columnconfigure(0, weight=1)
        ttk.Label(folder_row, text="Output folder:").grid(row=0, column=0, sticky="w")
        self.folder_entry = ttk.Entry(folder_row, textvariable=self.output_folder_var, width=42)
        self.folder_entry.grid(row=1, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(folder_row, text="Browse…", command=self.browse_output_folder).grid(row=1, column=1, sticky="ew")

        quality = ttk.LabelFrame(right, text="Quality", padding=10)
        quality.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        quality.columnconfigure(0, weight=1)

        self.format_var = tk.StringVar(value="png")
        ttk.Label(quality, text="Image format:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            quality,
            textvariable=self.format_var,
            values=["png", "jpg"],
            state="readonly",
            width=10,
        ).grid(row=1, column=0, sticky="w")
        ttk.Label(quality, text="png = lossless, jpg = smaller files").grid(row=2, column=0, sticky="w", pady=(2, 0))

        dpi_row = ttk.Frame(quality)
        dpi_row.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(dpi_row, text="Render DPI:").grid(row=0, column=0, sticky="w")
        self.dpi_var = tk.StringVar(value="300")
        ttk.Entry(dpi_row, textvariable=self.dpi_var, width=8).grid(row=0, column=1, sticky="w", padx=(6, 0))
        ttk.Label(dpi_row, text="(300 recommended, 600 for extra sharp)").grid(row=0, column=2, sticky="w", padx=(6, 0))

        jpg_row = ttk.Frame(quality)
        jpg_row.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(jpg_row, text="JPEG quality:").grid(row=0, column=0, sticky="w")
        self.jpeg_quality_var = tk.StringVar(value="95")
        ttk.Entry(jpg_row, textvariable=self.jpeg_quality_var, width=8).grid(row=0, column=1, sticky="w", padx=(6, 0))
        ttk.Label(jpg_row, text="(only used for jpg)").grid(row=0, column=2, sticky="w", padx=(6, 0))

        bottom = ttk.Frame(container)
        bottom.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        bottom.columnconfigure(0, weight=1)
        bottom.columnconfigure(1, weight=0)

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(bottom, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(bottom, mode="determinate", length=240)
        self.progress.grid(row=0, column=1, sticky="e", padx=(10, 0))

        action = ttk.Frame(container)
        action.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        action.columnconfigure(0, weight=1)
        action.columnconfigure(1, weight=0)
        action.columnconfigure(2, weight=0)

        self.convert_button = ttk.Button(action, text="Convert", command=self.start_convert)
        self.convert_button.grid(row=0, column=1, sticky="e", padx=(0, 6))
        self.cancel_button = ttk.Button(action, text="Cancel", command=self.request_cancel)
        self.cancel_button.grid(row=0, column=2, sticky="e")

    def _set_idle_state(self) -> None:
        self.cancel_button.state(["disabled"])
        self.progress["value"] = 0
        self.progress["maximum"] = 1
        self._update_summary()

    def _set_busy_state(self) -> None:
        self.cancel_button.state(["!disabled"])

    def _update_summary(self) -> None:
        if not self._pdfs:
            self.summary_label.configure(text="No files selected.")
            self.convert_button.state(["disabled"])
            return
        total_bytes = 0
        for path in self._pdfs:
            try:
                total_bytes += path.stat().st_size
            except Exception:
                pass
        self.summary_label.configure(text=f"{len(self._pdfs)} file(s) • {_format_bytes(total_bytes)}")
        self.convert_button.state(["!disabled"])
        if not self.output_folder_var.get().strip():
            self.output_folder_var.set(str(self._pdfs[0].parent))

    def _refresh_listbox(self) -> None:
        self.listbox.delete(0, tk.END)
        for p in self._pdfs:
            self.listbox.insert(tk.END, str(p))
        self._update_summary()

    def add_pdfs(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select PDF(s)",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*")],
            parent=self.window,
        )
        if not paths:
            self._bring_to_front()
            return
        new_paths: List[Path] = []
        for p in paths:
            path = Path(p)
            if path.suffix.lower() != ".pdf":
                continue
            if path not in self._pdfs and path.exists():
                new_paths.append(path)
        if not new_paths:
            return
        self._pdfs.extend(new_paths)
        self._refresh_listbox()
        self._bring_to_front()

    def remove_selected(self) -> None:
        indices = list(self.listbox.curselection())
        if not indices:
            return
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self._pdfs):
                del self._pdfs[idx]
        self._refresh_listbox()

    def clear_all(self) -> None:
        self._pdfs.clear()
        self._refresh_listbox()

    def browse_output_folder(self) -> None:
        initial = self.output_folder_var.get().strip() or str(Path.home())
        chosen = filedialog.askdirectory(title="Choose output folder", initialdir=initial, parent=self.window)
        if chosen:
            self.output_folder_var.set(chosen)
        self._bring_to_front()

    def open_output_folder(self) -> None:
        folder = self.output_folder_var.get().strip()
        if not folder:
            return
        p = Path(folder)
        if not p.exists():
            return
        try:
            os.startfile(str(p))
        except Exception:
            pass

    def _validate_before_convert(self) -> Optional[Tuple[List[Path], Path, str, int, int]]:
        if fitz is None:
            messagebox.showwarning(
                "PDF engine not installed",
                "To convert PDF pages to photos, install 'PyMuPDF'.\n\nRun: pip install -r requirements.txt",
            )
            return None
        if not self._pdfs:
            messagebox.showwarning("No files", "Please add at least one PDF.")
            return None

        out_folder_raw = self.output_folder_var.get().strip()
        if not out_folder_raw:
            messagebox.showwarning("Output missing", "Please choose an output folder.")
            return None
        out_folder = Path(out_folder_raw)
        out_folder.mkdir(parents=True, exist_ok=True)

        fmt = self.format_var.get().strip().lower()
        if fmt not in {"png", "jpg"}:
            messagebox.showwarning("Invalid setting", "Invalid image format.")
            return None

        dpi = max(36, _safe_int(self.dpi_var.get().strip(), 300))
        jpeg_quality = min(100, max(1, _safe_int(self.jpeg_quality_var.get().strip(), 95)))
        return (list(self._pdfs), out_folder, fmt, dpi, jpeg_quality)

    def start_convert(self) -> None:
        if self._worker_thread and self._worker_thread.is_alive():
            return
        validated = self._validate_before_convert()
        if validated is None:
            return
        pdfs, out_folder, fmt, dpi, jpeg_quality = validated

        self._stop_requested = False
        self._set_busy_state()
        self.status_var.set("Converting…")
        self.progress["maximum"] = 1
        self.progress["value"] = 0

        self._worker_thread = threading.Thread(
            target=self._convert_worker,
            args=(pdfs, out_folder, fmt, dpi, jpeg_quality),
            daemon=True,
        )
        self._worker_thread.start()

    def request_cancel(self) -> None:
        if self._worker_thread and self._worker_thread.is_alive():
            self._stop_requested = True
            self.status_var.set("Cancel requested…")

    def _poll_events(self) -> None:
        try:
            while True:
                event = self._events.get_nowait()
                kind = event[0]
                if kind == "progress":
                    current, total, msg = event[1], event[2], event[3]
                    self.progress["maximum"] = max(1, total)
                    self.progress["value"] = current
                    self.status_var.set(msg)
                elif kind == "done":
                    written = int(event[1])
                    self._set_idle_state()
                    self.status_var.set(f"Done. Saved {written} image(s).")
                    messagebox.showinfo("Done", f"Saved {written} image(s).")
                elif kind == "cancelled":
                    self._set_idle_state()
                    self.status_var.set("Cancelled.")
                    messagebox.showinfo("Cancelled", "Conversion cancelled.")
                elif kind == "error":
                    self._set_idle_state()
                    self.status_var.set("Error.")
                    messagebox.showerror("Error", str(event[1]))
        except queue.Empty:
            pass
        self.window.after(100, self._poll_events)

    def _convert_worker(self, pdfs: List[Path], out_folder: Path, fmt: str, dpi: int, jpeg_quality: int) -> None:
        try:
            assert fitz is not None
            total_pages = 0
            for pdf_path in pdfs:
                if self._stop_requested:
                    raise _Cancelled()
                with fitz.open(str(pdf_path)) as doc:
                    total_pages += int(doc.page_count)

            if total_pages <= 0:
                self._events.put(("done", 0))
                return

            current_page = 0
            written = 0
            zoom = float(dpi) / 72.0
            matrix = fitz.Matrix(zoom, zoom)

            for pdf_path in pdfs:
                if self._stop_requested:
                    raise _Cancelled()
                with fitz.open(str(pdf_path)) as doc:
                    base = pdf_path.stem
                    page_count = int(doc.page_count)
                    for page_number in range(page_count):
                        if self._stop_requested:
                            raise _Cancelled()

                        current_page += 1
                        self._events.put(
                            (
                                "progress",
                                current_page - 1,
                                total_pages,
                                f"Rendering {current_page}/{total_pages}: {pdf_path.name} (page {page_number + 1}/{page_count})",
                            )
                        )

                        page = doc.load_page(page_number)
                        pix = page.get_pixmap(matrix=matrix, alpha=False)

                        out_name = f"{base}_page_{page_number + 1:03d}.{fmt}"
                        out_path = out_folder / out_name

                        if fmt == "png":
                            pix.save(str(out_path))
                        else:
                            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                            img.save(
                                str(out_path),
                                format="JPEG",
                                quality=jpeg_quality,
                                subsampling=0,
                                optimize=True,
                            )
                        written += 1

            self._events.put(("progress", total_pages, total_pages, "Finalizing…"))
            self._events.put(("done", written))
        except _Cancelled:
            self._events.put(("cancelled",))
        except Exception as exc:
            self._events.put(("error", exc))


class _Cancelled(Exception):
    pass


def main() -> None:
    root = tk.Tk()
    try:
        style = ttk.Style()
        # Prefer native-looking themes on Windows when available.
        for theme in ("vista", "xpnative", "clam"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
    except Exception:
        pass
    ImageToPDFApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

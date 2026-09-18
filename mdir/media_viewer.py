from __future__ import annotations

import io
import threading
import math
import time
from pathlib import Path
from typing import Callable, Mapping, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk, ImageChops
from .ui_dispatch import install_dispatch
from .background import LatestWorker
from .excel_pdf_preview import _diagnostic

from .excel_viewer import copy_pil_image_to_windows_clipboard

PDF_EXTENSIONS = {".pdf"}
IMAGE_EXTENSIONS = {
    ".png", ".jpg", ".jpeg", ".jfif", ".webp", ".bmp", ".gif",
    ".tif", ".tiff", ".ico",
}


def media_kind_for_path(path: Path | str) -> Optional[str]:
    suffix = Path(path).suffix.lower()
    if suffix in PDF_EXTENSIONS:
        return "pdf"
    if suffix in IMAGE_EXTENSIONS:
        return "image"
    return None


def fit_zoom_for_size(
    source_width: int,
    source_height: int,
    viewport_width: int,
    viewport_height: int,
    *,
    margin: int = 14,
    allow_upscale: bool = False,
) -> float:
    """Return the zoom needed to fit an image inside the visible canvas.

    The default is a hybrid "1:1 / Fit" policy: images that already fit are
    kept at true 1:1 (100%), while oversized images are reduced just enough to
    fit in the current viewport.  Small images are therefore never enlarged
    unless ``allow_upscale`` is explicitly requested.
    """
    sw = max(1, int(source_width))
    sh = max(1, int(source_height))
    vw = max(1, int(viewport_width) - int(margin) * 2)
    vh = max(1, int(viewport_height) - int(margin) * 2)
    zoom = min(vw / sw, vh / sh)
    if not allow_upscale:
        zoom = min(1.0, zoom)
    return max(0.01, float(zoom))


_PDF_LOCK = threading.RLock()


def trim_excel_whitespace(image):
    """Crop only near-white outer margins, retaining a 12px safety border."""
    rgb = image.convert('RGB')
    difference = ImageChops.difference(rgb, Image.new('RGB', rgb.size, 'white'))
    # Even faint gray marks count as content; never trim the actual PDF file.
    bounds = difference.getbbox()
    if not bounds:
        return image
    left, top, right, bottom = bounds
    padding = 12
    return image.crop((max(0, left-padding), max(0, top-padding),
                       min(image.width, right+padding), min(image.height, bottom+padding)))


def _render_pdf_page(path, page_index, zoom):
    # PDFium is not thread safe, even across separate documents.
    with _PDF_LOCK:
        started = time.monotonic()
        result = _render_pdf_page_locked(path, page_index, zoom)
        _diagnostic('pdfium_render', source=path, page=page_index, elapsed=time.monotonic()-started)
        return result


def _render_pdf_page_locked(path: Path, page_index: int, zoom: float) -> tuple[Image.Image, int]:
    """Render one PDF page to a Pillow image and return (image, page_count).

    pypdfium2 is used instead of embedding Microsoft Office or requiring an
    external PDF application. It ships a PDFium-based renderer and works on
    Windows without a separate Poppler installation.
    """
    import pypdfium2 as pdfium

    document = pdfium.PdfDocument(str(path))
    try:
        page_count = len(document)
        if page_count <= 0:
            raise ValueError("This PDF contains no pages.")
        index = max(0, min(int(page_index), page_count - 1))
        page = document[index]
        try:
            # PDF points are 72 dpi. 96 dpi at 100% is a comfortable Windows
            # baseline while the UI zoom multiplier remains intuitive.
            scale = max(0.25, min(4.0, float(zoom))) * (96.0 / 72.0)
            width, height = page.get_size()
            scale = min(scale, math.sqrt(16_000_000 / max(1, width * height)))
            bitmap = page.render(scale=scale)
            try:
                image = bitmap.to_pil().convert("RGB")
            finally:
                try:
                    bitmap.close()
                except Exception:
                    pass
        finally:
            try:
                page.close()
            except Exception:
                pass
        return image, page_count
    finally:
        try:
            document.close()
        except Exception:
            pass


def _load_source_image(path: Path) -> Image.Image:
    with Image.open(path) as opened:
        # Animated formats are intentionally shown as their first frame.  Copying
        # keeps the data independent from the file handle after it closes.
        try:
            opened.seek(0)
        except Exception:
            pass
        image = opened.convert("RGBA") if opened.mode in {"P", "LA", "RGBA"} else opened.convert("RGB")
        return image.copy()


class MediaViewer(ttk.Frame):
    """Scrollable viewer for image files and rendered PDF pages.

    Rendering is done off the Tk thread.  Generation numbers invalidate stale
    results when the LEFT file selection changes quickly.
    """

    def __init__(
        self,
        parent,
        *,
        status_callback: Callable[[str], None],
        page_callback: Callable[[int, int, str], None],
        ui_palette: Mapping[str, str],
        zoom_callback: Optional[Callable[[float, str], None]] = None,
    ) -> None:
        super().__init__(parent)
        self._post = install_dispatch(self)
        self._worker = LatestWorker('xViewer-Media')
        self.bind('<Destroy>', lambda event: self._worker.close() if event.widget is self else None, add='+')
        self.trim_whitespace = False
        self._source_scale = 1.0
        self.status_callback = status_callback
        self.page_callback = page_callback
        self.ui_palette = dict(ui_palette)
        self.zoom_callback = zoom_callback
        self.failure_callback = None
        self.path: Optional[Path] = None
        self.kind: Optional[str] = None
        self.zoom = 1.0
        # Images start in hybrid 1:1/Fit mode: 100% when they fit, otherwise
        # shrink to the viewport. PDF pages keep their normal percentage zoom.
        self.image_fit_mode = True
        self.page_index = 0
        self.page_count = 0
        self.display_label: Optional[str] = None
        self._generation = 0
        self._source_image: Optional[Image.Image] = None
        self._display_image: Optional[Image.Image] = None
        self._tk_image: Optional[ImageTk.PhotoImage] = None
        self._resize_job: Optional[str] = None

        holder = ttk.Frame(self)
        holder.pack(fill="both", expand=True)
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            holder,
            bg=self.ui_palette.get("canvas_margin", self.ui_palette.get("window", "#ffffff")),
            highlightthickness=0,
            borderwidth=0,
        )
        ybar = ttk.Scrollbar(holder, orient="vertical", command=self.canvas.yview)
        xbar = ttk.Scrollbar(holder, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")

        self.canvas.bind("<MouseWheel>", self._mousewheel)
        self.canvas.bind("<Shift-MouseWheel>", self._shift_mousewheel)
        self.canvas.bind("<Control-MouseWheel>", self._ctrl_mousewheel)
        self.canvas.bind("<Button-1>", lambda _e: self.canvas.focus_set())
        self.canvas.bind("<Configure>", self._canvas_configured)

        self._show_message("Select a PDF or image file on the left.")

    def apply_palette(self, palette: Mapping[str, str]) -> None:
        self.ui_palette = dict(palette)
        try:
            self.canvas.configure(bg=self.ui_palette.get("canvas_margin", self.ui_palette.get("window", "#ffffff")))
            self.canvas.itemconfigure("message", fill=self.ui_palette.get("foreground", "#202020"))
        except Exception:
            pass

    def clear(self) -> None:
        self._worker.cancel()
        self._generation += 1
        if self._resize_job:
            try:
                self.after_cancel(self._resize_job)
            except Exception:
                pass
            self._resize_job = None
        self.path = None
        self.kind = None
        self.page_index = 0
        self.page_count = 0
        self.display_label = None
        self.image_fit_mode = True
        self._source_image = None
        self._display_image = None
        self._tk_image = None
        self.canvas.delete("all")
        self.canvas.configure(scrollregion=(0, 0, 1, 1))
        self.page_callback(0, 0, "")
        self._show_message("Select a PDF or image file on the left.")

    def show_message(self, text: str) -> None:
        """Display a centered transient message without changing the current file."""
        self._show_message(text)

    def _show_message(self, text: str) -> None:
        self.canvas.delete("all")
        try:
            width = max(300, self.canvas.winfo_width())
            height = max(200, self.canvas.winfo_height())
        except Exception:
            width, height = 800, 500
        self.canvas.create_text(
            width // 2,
            height // 2,
            text=text,
            fill=self.ui_palette.get("foreground", "#202020"),
            font=("Segoe UI", 13),
            justify="center",
            tags=("message",),
        )
        self.canvas.configure(scrollregion=(0, 0, width, height))

    def load(self, path: Path | str, kind: str, *, display_label: Optional[str] = None, trim_whitespace: bool = False) -> None:
        path = Path(path)
        kind = str(kind).lower()
        if kind not in {"pdf", "image"}:
            raise ValueError(f"Unsupported media kind: {kind}")
        self._generation += 1
        generation = self._generation
        self.path = path
        self.kind = kind
        self.display_label = display_label
        self.trim_whitespace = trim_whitespace
        self.page_index = 0
        self.page_count = 1 if kind == "image" else 0
        self.image_fit_mode = True
        self.zoom = 1.0
        self._source_scale = 1.0
        self._source_image = None
        self._display_image = None
        self._tk_image = None
        self._show_message(f"Loading {path.name} ...")
        self.status_callback(f"Loading {kind.upper()}: {path.name} ...")

        def worker() -> None:
            try:
                if kind == "image":
                    image = _load_source_image(path)
                    result = (image, 1)
                else:
                    result = _render_pdf_page(path, 0, 1.0)
            except Exception as exc:
                self._post(lambda error=exc: self._load_failed(generation, path, error))
                return
            self._post(lambda: self._load_complete(generation, path, kind, result[0], result[1]))

        self._worker.submit(worker)

    def _load_failed(self, generation: int, path: Path, exc: Exception) -> None:
        if generation != self._generation:
            return
        self.page_count = 0
        self._source_image = None
        self._display_image = None
        self._show_message(f"Could not open:\n{path.name}\n\n{exc}")
        self.status_callback(f"Open failed: {path.name} | {exc}")
        self.page_callback(0, 0, self.kind or "")
        if self.failure_callback is not None:
            self.failure_callback(path, exc)

    def _load_complete(
        self,
        generation: int,
        path: Path,
        kind: str,
        image: Image.Image,
        page_count: int,
    ) -> None:
        if generation != self._generation or self.path != path or self.kind != kind:
            return
        self.page_count = max(1, int(page_count))
        self._source_image = trim_excel_whitespace(image) if self.trim_whitespace and kind == 'pdf' else image
        self._source_scale = 1.0
        self._render_image_source()
        self.page_callback(self.page_index, self.page_count, kind)
        shown = self.display_label or str(path)
        if kind == "pdf":
            self.status_callback(f"{shown} | PDF | {self.page_count} page(s) | Page {self.page_index + 1}")
        else:
            pct = max(1, int(round(self.zoom * 100)))
            mode = f"Fit {pct}% (max 1:1)" if self.image_fit_mode else ("1:1 (100%)" if pct == 100 else f"{pct}%")
            self.status_callback(f"{shown} | Image | {image.width} × {image.height} px | {mode}")

    def _paint(self, image: Image.Image) -> None:
        self.canvas.delete("all")
        self._display_image = image
        self._tk_image = ImageTk.PhotoImage(image)
        margin = 14
        self.canvas.create_image(margin, margin, image=self._tk_image, anchor="nw", tags=("media",))
        self.canvas.configure(scrollregion=(0, 0, image.width + margin * 2, image.height + margin * 2))
        self.canvas.xview_moveto(0.0)
        self.canvas.yview_moveto(0.0)

    def _notify_zoom(self, mode: str) -> None:
        if self.zoom_callback is None:
            return
        try:
            self.zoom_callback(float(self.zoom), str(mode))
        except Exception:
            pass

    def _fit_zoom(self) -> float:
        source = self._source_image
        if source is None:
            return 1.0
        return fit_zoom_for_size(
            source.width / self._source_scale,
            source.height / self._source_scale,
            max(1, self.canvas.winfo_width()),
            max(1, self.canvas.winfo_height()),
            allow_upscale=self.kind == "pdf",
        )

    def _render_image_source(self) -> None:
        source = self._source_image
        if source is None:
            return
        if self.image_fit_mode:
            self.zoom = self._fit_zoom()
            mode = "fit"
        else:
            self.zoom = max(0.01, min(4.0, float(self.zoom)))
            mode = "manual" if abs(self.zoom - 1.0) > 1e-9 else "1:1"
        zoom = self.zoom / self._source_scale
        width = max(1, int(round(source.width * zoom)))
        height = max(1, int(round(source.height * zoom)))
        if width * height > 24_000_000:
            reduction = math.sqrt(24_000_000 / (width * height))
            width, height = max(1, int(width * reduction)), max(1, int(height * reduction))
        if width == source.width and height == source.height:
            display = source.copy()
        else:
            display = source.resize((width, height), Image.Resampling.LANCZOS)
        self._paint(display)
        self._notify_zoom(mode)

    def set_zoom(self, zoom: float) -> None:
        self.zoom = max(0.25, min(4.0, float(zoom)))
        if self.path is None or self.kind is None:
            return
        if self.kind == "image":
            # Choosing a numeric zoom leaves Fit mode.  The dedicated
            # 1:1 / Fit button can re-enter Fit at any time.
            self.image_fit_mode = False
            # Debounce repeated Ctrl+wheel changes so very large images do not
            # get resized for every single wheel event.
            if self._resize_job:
                try:
                    self.after_cancel(self._resize_job)
                except Exception:
                    pass
            self._resize_job = self.after(70, self._render_image_after_zoom)
        else:
            self.image_fit_mode = False
            if self._resize_job:
                self.after_cancel(self._resize_job)
            self._resize_job = self.after(100, self._render_pdf_current_page)

    def show_fit(self) -> None:
        """Fit an image to the viewport without enlarging it above 1:1."""
        if self.kind not in {"image", "pdf"} or self._source_image is None:
            return
        self.image_fit_mode = True
        self._schedule_image_render(20)

    def show_one_to_one(self) -> None:
        """Show an image at its original pixel size (100%)."""
        if self.kind not in {"image", "pdf"} or self._source_image is None:
            return
        self.image_fit_mode = False
        self.zoom = 1.0
        self._schedule_image_render(20)

    def toggle_fit_one_to_one(self) -> str:
        """Toggle image display between Fit and original 1:1 size."""
        if self.kind not in {"image", "pdf"}:
            return ""
        if self.image_fit_mode:
            self.show_one_to_one()
            return "1:1"
        self.show_fit()
        return "fit"

    def _schedule_image_render(self, delay: int = 70) -> None:
        if self._resize_job:
            try:
                self.after_cancel(self._resize_job)
            except Exception:
                pass
        self._resize_job = self.after(max(0, int(delay)), self._render_image_after_zoom)

    def _canvas_configured(self, _event=None) -> None:
        # In Fit mode the image follows the available RIGHT-pane size.  Debounce
        # resize storms while the user drags the pane splitter/window edge.
        if self.kind in {"image", "pdf"} and self.image_fit_mode and self._source_image is not None:
            self._schedule_image_render(90)

    def _render_image_after_zoom(self) -> None:
        self._resize_job = None
        self._render_image_source()

    def _render_pdf_current_page(self) -> None:
        if self.path is None or self.kind != "pdf":
            return
        self._generation += 1
        generation = self._generation
        path = self.path
        page_index = self.page_index
        zoom = 1.0 if self.image_fit_mode else self.zoom
        self._resize_job = None
        self._show_message(f"Rendering {path.name}\nPage {page_index + 1} ...")

        def worker() -> None:
            try:
                image, count = _render_pdf_page(path, page_index, zoom)
            except Exception as exc:
                self._post(lambda error=exc: self._load_failed(generation, path, error))
                return
            self._post(lambda: self._pdf_page_complete(generation, path, page_index, image, count, zoom))

        self._worker.submit(worker)

    def _pdf_page_complete(
        self,
        generation: int,
        path: Path,
        page_index: int,
        image: Image.Image,
        page_count: int,
        source_scale: float = 1.0,
    ) -> None:
        if generation != self._generation or self.path != path or self.kind != "pdf" or page_index != self.page_index:
            return
        self.page_count = max(1, int(page_count))
        self._source_image = trim_excel_whitespace(image) if self.trim_whitespace else image
        self._source_scale = source_scale
        self._render_image_source()
        self.page_callback(self.page_index, self.page_count, "pdf")
        self.status_callback(f"{path} | PDF | {self.page_count} page(s) | Page {self.page_index + 1}")

    def step_page(self, delta: int) -> None:
        if self.kind != "pdf" or self.path is None or self.page_count <= 0:
            return
        new_index = max(0, min(self.page_count - 1, self.page_index + int(delta)))
        if new_index == self.page_index:
            return
        self.page_index = new_index
        self.page_callback(self.page_index, self.page_count, "pdf")
        self._render_pdf_current_page()

    def first_page(self) -> None:
        if self.kind == "pdf" and self.page_count:
            self.page_index = 0
            self.page_callback(self.page_index, self.page_count, "pdf")
            self._render_pdf_current_page()

    def last_page(self) -> None:
        if self.kind == "pdf" and self.page_count:
            self.page_index = self.page_count - 1
            self.page_callback(self.page_index, self.page_count, "pdf")
            self._render_pdf_current_page()

    def current_image_for_copy(self) -> Optional[Image.Image]:
        if self.kind == "image" and self._source_image is not None:
            return self._source_image.copy()
        if self._display_image is not None:
            return self._display_image.copy()
        return None

    def copy_current(self) -> None:
        image = self.current_image_for_copy()
        if image is None:
            return
        try:
            copy_pil_image_to_windows_clipboard(image)
            label = self.path.name if self.path else "image"
            self.status_callback(f"Copied to Windows clipboard: {label}")
        except Exception as exc:
            messagebox.showerror("xViewer", f"Could not copy image:\n{exc}", parent=self)

    def save_current_image_as(self) -> None:
        image = self.current_image_for_copy()
        if image is None:
            return
        initial = (self.path.stem if self.path else "image")
        if self.kind == "pdf":
            initial += f"-page-{self.page_index + 1}"
        target = filedialog.asksaveasfilename(
            parent=self,
            title="Save displayed image as",
            initialfile=initial + ".png",
            defaultextension=".png",
            filetypes=[("PNG image", "*.png"), ("JPEG image", "*.jpg;*.jpeg"), ("All files", "*.*")],
        )
        if not target:
            return
        try:
            out = Path(target)
            save_image = image
            if out.suffix.lower() in {".jpg", ".jpeg"} and save_image.mode != "RGB":
                save_image = save_image.convert("RGB")
            save_image.save(out)
            self.status_callback(f"Saved image: {out}")
        except Exception as exc:
            messagebox.showerror("xViewer", f"Could not save image:\n{exc}", parent=self)

    def _mousewheel(self, event) -> str:
        # Windows wheel is multiples of 120; small fallback keeps other Tk ports usable.
        delta = int(getattr(event, "delta", 0) or 0)
        units = -max(1, abs(delta) // 120) if delta > 0 else max(1, abs(delta) // 120)
        self.canvas.yview_scroll(units * 3, "units")
        return "break"

    def _shift_mousewheel(self, event) -> str:
        delta = int(getattr(event, "delta", 0) or 0)
        units = -max(1, abs(delta) // 120) if delta > 0 else max(1, abs(delta) // 120)
        self.canvas.xview_scroll(units * 3, "units")
        return "break"

    def _ctrl_mousewheel(self, event) -> str:
        delta = int(getattr(event, "delta", 0) or 0)
        if delta > 0:
            self.set_zoom(self.zoom + 0.1)
        elif delta < 0:
            self.set_zoom(self.zoom - 0.1)
        return "break"

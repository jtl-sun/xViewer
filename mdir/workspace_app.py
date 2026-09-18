from __future__ import annotations

import io
import json
import os
import shutil
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from send2trash import send2trash

from .links import (
    LINK_KINDS,
    LINKS_CONFIG_PATH,
    MAX_LINKS,
    LinkDefinition,
    LinkManager,
    expand_link_text,
    load_links,
    load_mdir_links,
    save_links,
)

from .theme import THEME_CHOICES, effective_theme_name, normalize_theme_mode, palette_for
from .media_viewer import IMAGE_EXTENSIONS, PDF_EXTENSIONS, MediaViewer, media_kind_for_path

from .excel_viewer import (
    APP_TITLE,
    EXCEL_EXTENSIONS,
    MODERN_EXCEL_EXTENSIONS,
    EmbeddedImage,
    OpenPyxlSheetAdapter,
    VirtualSheet,
    WorkbookModel,
    XlrdSheetAdapter,
    legacy_xls_preview_path,
    legacy_xls_images,
    legacy_xls_pdf_preview_path,
    _merge_embedded_images,
    _column_name,
    _emu_to_pixels,
)

from . import __version__ as APP_VERSION
CONFIG_PATH = Path.home() / ".xexcel-viewer.json"


def _set_windows_app_user_model_id() -> None:
    """Give Windows a stable xViewer identity for taskbar/shortcut icon grouping."""
    if os.name != "nt":
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("jtl-sun.xViewer")
    except Exception:
        pass
MAX_QUICK_LINKS = MAX_LINKS
MAX_RECENT_DIRS = 30


# Backward-compatible helper names retained for the 2.3.x test/API surface.
def _normalize_quick_links(values: object) -> list[dict[str, object]]:
    from .links import parse_links
    return [
        {
            "label": link.label,
            "type": link.kind,
            "target": link.target,
            "args": list(link.args),
            "pane": link.pane,
        }
        for link in parse_links(values)
    ]


def _load_mdir_folder_links(path: Path | None = None) -> list[dict[str, str]]:
    links = load_mdir_links(path) if path is not None else load_mdir_links()
    return [{"label": link.label, "target": link.target} for link in links if link.kind == "folder"]


def _load_quick_links(path: Path = LINKS_CONFIG_PATH):
    return load_links(path)


def _save_quick_links(links, path: Path = LINKS_CONFIG_PATH) -> None:
    normalized: list[LinkDefinition] = []
    for item in links:
        if isinstance(item, LinkDefinition):
            normalized.append(item)
        elif isinstance(item, dict):
            normalized.append(LinkDefinition(
                str(item.get("label", "")).strip(),
                str(item.get("type", item.get("kind", "folder"))).strip().lower(),
                str(item.get("target", "")).strip(),
                tuple(str(x) for x in item.get("args", []) if isinstance(item.get("args", []), list)),
                str(item.get("pane", "active")).strip().lower(),
            ))
    save_links(normalized, path)


def _expand_quick_link_target(value: str, current: Path) -> Path:
    return Path(expand_link_text(value, current=current, selected=None, workbook=None, project=Path(__file__).resolve().parents[1]))


def _human_size(size: int) -> str:
    value = float(max(0, size))
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if value < 1024 or unit == "TB":
            if unit == "B":
                return f"{int(value):,}"
            return f"{value:,.1f} {unit}" if value < 10 else f"{value:,.0f} {unit}"
        value /= 1024
    return f"{size:,}"


def _centered_geometry(screen_width: int, screen_height: int, width: int, height: int) -> str:
    """Return a Tk geometry string centered on the primary display.

    The main xViewer window intentionally starts on the primary monitor rather
    than inheriting a stale OS/Tk position.  Coordinates are clamped so smaller
    screens never receive negative top/left positions.
    """
    sw = max(1, int(screen_width))
    sh = max(1, int(screen_height))
    w = max(1, int(width))
    h = max(1, int(height))
    x = max(0, (sw - w) // 2)
    y = max(0, (sh - h) // 2)
    return f"{w}x{h}+{x}+{y}"


def _recent_directory_key(path: Path | str) -> str:
    """Stable de-duplication key for the persistent recent-folder list."""
    try:
        text = str(Path(path).expanduser().resolve())
    except Exception:
        text = str(Path(path).expanduser())
    return os.path.normcase(os.path.normpath(text))


def _merge_recent_directories(current: Path, values: Iterable[Path | str], limit: int = MAX_RECENT_DIRS) -> list[Path]:
    """Return MRU directories with *current* first and duplicates removed."""
    result: list[Path] = []
    seen: set[str] = set()
    for value in (current, *tuple(values)):
        try:
            path = Path(value).expanduser()
            try:
                path = path.resolve()
            except Exception:
                pass
        except Exception:
            continue
        key = _recent_directory_key(path)
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(path)
        if len(result) >= max(1, int(limit)):
            break
    return result


def _file_name_matches_filter(name: str, query: str) -> bool:
    """Case-insensitive filename filter with simple AND terms.

    Examples:
    - ``010719`` matches any name containing 010719.
    - ``010719 TJ`` requires both terms to appear, in any order.
    Empty queries match everything.
    """
    terms = [part.casefold() for part in str(query or "").split() if part.strip()]
    if not terms:
        return True
    haystack = str(name or "").casefold()
    return all(term in haystack for term in terms)


def _search_arrow_target_index(item_count: int, current_index: Optional[int], direction: int) -> Optional[int]:
    """Return the LEFT-list row targeted when Up/Down leaves the search box.

    If there is no usable current row, Down starts at the first filtered result
    and Up starts at the last filtered result.  Otherwise the selection moves
    one visible row in the requested direction and is clamped at the ends.
    """
    count = max(0, int(item_count))
    if count <= 0:
        return None
    step = 1 if int(direction) >= 0 else -1
    if current_index is None or current_index < 0 or current_index >= count:
        return 0 if step > 0 else count - 1
    return max(0, min(count - 1, current_index + step))


def _left_file_tab_action() -> str:
    """LEFT file-list Tab policy.

    Tab is intentionally disabled while the FILES Treeview owns focus.  The
    workbook uses Tab for Excel-style cell movement, while main-pane switching
    is handled by Ctrl+Left/Ctrl+Right (or Alt+1/Alt+2).  Returning ``break``
    prevents Tk's default focus traversal from jumping to an invisible/unclear
    toolbar control.
    """
    return "break"


def _type_filter_accepts(
    path: Path | str,
    is_dir: bool,
    *,
    excel_only: bool,
    pdf_only: bool,
    images_only: bool,
) -> bool:
    """Return whether one LEFT-pane item passes the active type filter.

    Folders always remain visible for navigation.  The three ``... only``
    controls are exclusive in the GUI, but this helper intentionally supports
    unions as well so tests and future UI changes remain predictable.  If none
    of the controls is enabled, every file type is shown.
    """
    if is_dir:
        return True
    enabled = []
    if excel_only:
        enabled.append(EXCEL_EXTENSIONS)
    if pdf_only:
        enabled.append(PDF_EXTENSIONS)
    if images_only:
        enabled.append(IMAGE_EXTENSIONS)
    if not enabled:
        return True
    suffix = Path(path).suffix.lower()
    return any(suffix in extensions for extensions in enabled)


def _selection_workbook_action(paths: Iterable[Path | str]) -> str:
    """Return how the RIGHT viewer should react to LEFT selection.

    Excel keeps the historical ``"load"`` action for compatibility. PDFs and
    images use ``"load_pdf"`` / ``"load_image"``. A directory or unsupported
    file clears stale content. Zero/multi-selection keeps the current view.
    """
    items = [Path(value) for value in paths]
    if len(items) != 1:
        return "keep"
    path = items[0]
    if not path.is_file():
        return "clear"
    suffix = path.suffix.lower()
    if suffix in EXCEL_EXTENSIONS:
        return "load"
    if suffix in PDF_EXTENSIONS:
        return "load_pdf"
    if suffix in IMAGE_EXTENSIONS:
        return "load_image"
    return "clear"


def _send_paths_to_recycle_bin(paths: Iterable[Path | str]) -> tuple[list[Path], list[tuple[Path, Exception]]]:
    """Move paths to the OS recycle bin, returning successes and failures.

    Deletion is intentionally never downgraded to a permanent unlink/rmtree
    fallback.  If the recycle-bin operation fails, the item is left in place.
    """
    moved: list[Path] = []
    failed: list[tuple[Path, Exception]] = []
    for value in paths:
        path = Path(value)
        try:
            send2trash(str(path))
            moved.append(path)
        except Exception as exc:  # keep going so one bad item does not block the rest
            failed.append((path, exc))
    return moved, failed

def _parse_clipboard_value(text: str) -> Any:
    value = text.strip()
    if not value:
        return None
    if value.startswith("="):
        return value
    upper = value.upper()
    if upper == "TRUE":
        return True
    if upper == "FALSE":
        return False
    try:
        if value.startswith("0") and len(value) > 1 and not value.startswith("0."):
            return text
        return int(value)
    except ValueError:
        pass
    try:
        return float(value.replace(",", ""))
    except ValueError:
        return text


def _pane_hotkey_action(
    keysym: str,
    keycode: int,
    state: int,
    *,
    ctrl_down: Optional[bool] = None,
    alt_down: Optional[bool] = None,
) -> Optional[str]:
    """Return the requested main pane from one keyboard event.

    ``ctrl_down`` / ``alt_down`` are tri-state on purpose.  ``None`` means
    "trust Tk's modifier mask" (used on X11/Linux).  On Windows the caller
    passes the *physical* key state.  In that case we deliberately do not OR
    the value with Tk's state mask, because some Windows/Tk keyboard layouts
    can leave Alt-like state bits on ordinary number-key events.  That used to
    make typing 1 or 2 in the filename Search box jump to another pane.
    """
    key = str(keysym or "")
    code = int(keycode or 0)
    flags = int(state or 0)

    tk_ctrl = bool(flags & 0x0004)
    # X11/Tk commonly uses Mod1 (0x0008).  Windows/Tk may report Alt as
    # 0x20000, but on Windows we prefer the physical key state supplied above.
    tk_alt = bool((flags & 0x0008) or (flags & 0x20000))
    ctrl = tk_ctrl if ctrl_down is None else bool(ctrl_down)
    alt = tk_alt if alt_down is None else bool(alt_down)

    if ctrl and (key in {"Left", "KP_Left"} or code == 37):
        return "left"
    if ctrl and (key in {"Right", "KP_Right"} or code == 39):
        return "right"
    if alt and (key in {"1", "KP_1"} or code in {49, 97}):
        return "left"
    if alt and (key in {"2", "KP_2"} or code in {50, 98}):
        return "right"
    return None


class EditableOpenPyxlSheetAdapter(OpenPyxlSheetAdapter):
    def __init__(self, worksheet, images: list[EmbeddedImage], model: "EditableWorkbookModel") -> None:
        super().__init__(worksheet, images)
        self.model = model

    @property
    def editable(self) -> bool:
        return True

    def set_value(self, row: int, column: int, value: Any) -> None:
        cell = self.ws.cell(row, column)
        if cell.value == value:
            return
        cell.value = value
        self.max_row = max(self.max_row, row)
        self.max_column = max(self.max_column, column)
        self.model.dirty = True


class ReadOnlyXlrdSheetAdapter(XlrdSheetAdapter):
    @property
    def editable(self) -> bool:
        return False


class EditableWorkbookModel(WorkbookModel):
    """Workbook model used by the integrated right-hand editor.

    Modern Excel workbooks are loaded with formulas preserved and can be saved.
    Legacy .xls workbooks remain view/copy-only.
    """

    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self.show_formulas = True
        self.workbook = None
        self.sheets = []
        self.kind = "Excel"
        self.dirty = False
        self.editable = False
        self._backup_created = False
        self._load_editable()

    @staticmethod
    def _modern_images(worksheet) -> list[EmbeddedImage]:
        """Extract display copies while preserving openpyxl's image streams for save."""
        results: list[EmbeddedImage] = []
        for number, image in enumerate(getattr(worksheet, "_images", []), start=1):
            try:
                anchor = getattr(image, "anchor", None)
                marker = getattr(anchor, "_from", None)
                if marker is None:
                    continue
                row = int(marker.row) + 1
                column = int(marker.col) + 1
                x_offset = _emu_to_pixels(getattr(marker, "colOff", 0))
                y_offset = _emu_to_pixels(getattr(marker, "rowOff", 0))
                raw = bytes(image._data())
                if not raw:
                    continue
                image.ref = io.BytesIO(raw)
                extension = ".png"
                try:
                    extension = "." + str(getattr(image, "format", "png") or "png").lower().lstrip(".")
                except Exception:
                    pass
                width = max(1, int(round(float(getattr(image, "width", 96) or 96))))
                height = max(1, int(round(float(getattr(image, "height", 96) or 96))))
                ext = getattr(anchor, "ext", None)
                if ext is not None:
                    width = max(1, _emu_to_pixels(getattr(ext, "cx", 0)) or width)
                    height = max(1, _emu_to_pixels(getattr(ext, "cy", 0)) or height)
                to_marker = getattr(anchor, "_to", None)
                end_row = end_column = None
                end_x_offset = end_y_offset = 0
                if to_marker is not None:
                    end_row = int(to_marker.row) + 1
                    end_column = int(to_marker.col) + 1
                    end_x_offset = _emu_to_pixels(getattr(to_marker, "colOff", 0))
                    end_y_offset = _emu_to_pixels(getattr(to_marker, "rowOff", 0))
                results.append(
                    EmbeddedImage(
                        row=row,
                        column=column,
                        width=width,
                        height=height,
                        raw=raw,
                        extension=extension,
                        label=f"Image {number}",
                        x_offset=x_offset,
                        y_offset=y_offset,
                        end_row=end_row,
                        end_column=end_column,
                        end_x_offset=end_x_offset,
                        end_y_offset=end_y_offset,
                    )
                )
            except Exception:
                continue
        return results

    def _load_editable(self) -> None:
        suffix = self.path.suffix.lower()
        if suffix in MODERN_EXCEL_EXTENSIONS:
            from openpyxl import load_workbook

            keep_vba = suffix in {".xlsm", ".xltm"}
            workbook = load_workbook(
                self.path,
                data_only=False,
                read_only=False,
                keep_vba=keep_vba,
            )
            self.workbook = workbook
            self.sheets = [
                EditableOpenPyxlSheetAdapter(ws, self._modern_images(ws), self)
                for ws in workbook.worksheets
            ]
            self.kind = "Editable Excel"
            self.editable = True
            return

        if suffix == ".xls":
            preview = legacy_xls_preview_path(self.path)
            legacy_by_sheet = legacy_xls_images(self.path)
            if preview is not None:
                from openpyxl import load_workbook

                workbook = load_workbook(
                    preview,
                    data_only=False,
                    read_only=False,
                    keep_vba=False,
                )
                self.workbook = workbook
                self.sheets = [
                    OpenPyxlSheetAdapter(
                        ws,
                        _merge_embedded_images(self._modern_images(ws), legacy_by_sheet.get(ws.title, [])),
                    )
                    for ws in workbook.worksheets
                ]
                self.kind = "Legacy Excel (.xls) - read-only converted preview + direct picture export"
                self.editable = False
                return

            import xlrd

            workbook = xlrd.open_workbook(self.path, on_demand=True, formatting_info=False)
            self.workbook = workbook
            self.sheets = [
                ReadOnlyXlrdSheetAdapter(
                    workbook.sheet_by_index(i),
                    legacy_by_sheet.get(workbook.sheet_by_index(i).name, []),
                )
                for i in range(workbook.nsheets)
            ]
            self.kind = "Legacy Excel (.xls) - read only (cell fallback + direct picture export)"
            self.editable = False
            return

        raise ValueError(f"Unsupported Excel format: {suffix}")

    def save(self, target: Optional[Path] = None, *, create_backup: bool = True) -> Path:
        if not self.editable or self.workbook is None:
            raise RuntimeError("This workbook is read-only in xViewer.")
        target = Path(target or self.path)
        if target.suffix.lower() not in MODERN_EXCEL_EXTENSIONS:
            raise ValueError("Save as .xlsx, .xlsm, .xltx or .xltm.")

        if create_backup and target.resolve() == self.path.resolve() and self.path.exists() and not self._backup_created:
            backup_dir = self.path.parent / "xExcel_Backup"
            backup_dir.mkdir(exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            backup_name = f"{self.path.stem}-{stamp}{self.path.suffix}"
            shutil.copy2(self.path, backup_dir / backup_name)
            self._backup_created = True

        self.workbook.save(target)
        self.path = target
        self.dirty = False
        return target


class EditableVirtualSheet(VirtualSheet):
    def __init__(self, master, sheet, status_callback, selection_callback=None, ui_palette=None) -> None:
        self.selection_callback = selection_callback
        self._edit_entry: Optional[ttk.Entry] = None
        super().__init__(master, sheet, status_callback, ui_palette=ui_palette)

    @property
    def editable(self) -> bool:
        return bool(getattr(self.sheet, "editable", False))

    def _build_ui(self) -> None:
        super()._build_ui()
        self.body.bind("<Double-Button-1>", self._double_click, add="+")
        self.body.bind("<F2>", lambda _e: self.start_edit())
        self.body.bind("<Control-v>", self.paste_selection)
        self.body.bind("<Control-V>", self.paste_selection)
        self.body.bind("<Delete>", self.clear_selection)
        # Tab is reserved for normal Excel-style cell movement.  Main-pane
        # switching uses Ctrl+Left / Ctrl+Right. Alt+1 / Alt+2 remain direct shortcuts; F6 is kept for compatibility.
        self.body.bind("<Tab>", lambda _e: self._tab_cell(False))
        self.body.bind("<Shift-Tab>", lambda _e: self._tab_cell(True))
        self.body.bind("<ISO_Left_Tab>", lambda _e: self._tab_cell(True))

    def _report_cell(self) -> None:
        super()._report_cell()
        if callable(self.selection_callback):
            self.selection_callback(self)

    def _double_click(self, event) -> str:
        if self._image_from_event(event) is not None:
            return "break"
        self.active_cell = self._cell_from_event(event)
        self.anchor_cell = self.active_cell
        self.selected_image = None
        self.schedule_render()
        self.start_edit()
        return "break"

    def raw_active_value(self) -> Any:
        row, col = self.active_cell
        return self.sheet.value(row, col)

    def set_active_value(self, value: Any) -> bool:
        if not self.editable:
            self.status_callback("This .xls workbook is read-only. Use Open in Excel to edit it.")
            return False
        row, col = self.active_cell
        try:
            self.sheet.set_value(row, col, value)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not edit cell:\n{exc}", parent=self.winfo_toplevel())
            return False
        self._refresh_dimensions_if_needed()
        self.schedule_render()
        self._report_cell()
        return True

    def start_edit(self) -> None:
        if not self.editable:
            self.status_callback("This .xls workbook is read-only. Use Open in Excel to edit it.")
            return
        if self.selected_image is not None:
            return
        self.cancel_edit()
        row, col = self.active_cell
        merged = self._cell_anchor(row, col)
        end_row, end_col = row, col
        if merged is not None:
            row, col, end_row, end_col = merged
            self.active_cell = (row, col)
            self.anchor_cell = (row, col)

        left = self.body.canvasx(0)
        top = self.body.canvasy(0)
        x0 = self.columns.start(col) - left
        y0 = self.rows.start(row) - top
        x1 = self.columns.start(end_col) + self.columns.size(end_col) - left
        y1 = self.rows.start(end_row) + self.rows.size(end_row) - top

        entry = ttk.Entry(self.body)
        raw = self.sheet.value(row, col)
        entry.insert(0, "" if raw is None else str(raw))
        entry.place(
            x=max(0, int(x0 + 1)),
            y=max(0, int(y0 + 1)),
            width=max(50, int(x1 - x0 - 2)),
            height=max(22, int(y1 - y0 - 2)),
        )
        entry.focus_set()
        entry.selection_range(0, "end")
        entry.bind("<Return>", lambda _e: self.commit_edit())
        entry.bind("<Escape>", lambda _e: self.cancel_edit())
        entry.bind("<Tab>", lambda _e: self._commit_and_tab(False))
        entry.bind("<Shift-Tab>", lambda _e: self._commit_and_tab(True))
        entry.bind("<ISO_Left_Tab>", lambda _e: self._commit_and_tab(True))
        entry.bind("<FocusOut>", lambda _e: self.commit_edit())
        installer = getattr(self.winfo_toplevel(), "_apply_pane_nav_bindtag", None)
        if callable(installer):
            installer(entry)
        self._edit_entry = entry

    def _tab_cell(self, reverse: bool = False) -> str:
        """Move to the next visible cell horizontally, skipping hidden columns."""
        row, col = self.active_cell
        if reverse:
            new_col = self.columns.next_visible(col, -1)
            if new_col == col:
                new_row = self.rows.next_visible(row, -1)
                if new_row != row:
                    row = new_row
                    col = self.columns.last_visible()
            else:
                col = new_col
        else:
            new_col = self.columns.next_visible(col, 1)
            if new_col == col:
                new_row = self.rows.next_visible(row, 1)
                if new_row != row:
                    row = new_row
                    col = self.columns.first_visible()
            else:
                col = new_col
        self.active_cell = (row, col)
        self.anchor_cell = self.active_cell
        self.selected_image = None
        self._ensure_visible(row, col)
        self._report_cell()
        self.schedule_render()
        self.body.focus_set()
        return "break"

    def _commit_and_tab(self, reverse: bool = False) -> str:
        self.commit_edit()
        return self._tab_cell(reverse)

    def commit_edit(self) -> str:
        entry = self._edit_entry
        if entry is None:
            return "break"
        value = entry.get()
        self._edit_entry = None
        try:
            entry.destroy()
        except Exception:
            pass
        self.set_active_value(_parse_clipboard_value(value))
        self.body.focus_set()
        return "break"

    def cancel_edit(self) -> str:
        entry = self._edit_entry
        self._edit_entry = None
        if entry is not None:
            try:
                entry.destroy()
            except Exception:
                pass
        self.body.focus_set()
        return "break"

    def paste_selection(self, _event=None) -> str:
        if not self.editable:
            self.status_callback("This .xls workbook is read-only. Use Open in Excel to edit it.")
            return "break"
        try:
            text = self.clipboard_get()
        except Exception:
            self.status_callback("Clipboard does not contain text that can be pasted into cells.")
            return "break"
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if lines and lines[-1] == "":
            lines.pop()
        if not lines:
            return "break"
        start_row, start_col = self.active_cell
        changed = 0
        for r_off, line in enumerate(lines):
            fields = line.split("\t")
            for c_off, field in enumerate(fields):
                try:
                    self.sheet.set_value(start_row + r_off, start_col + c_off, _parse_clipboard_value(field))
                    changed += 1
                except Exception:
                    continue
        self.active_cell = (start_row + len(lines) - 1, start_col + max(len(line.split("\t")) for line in lines) - 1)
        self._refresh_dimensions_if_needed()
        self._ensure_visible(*self.active_cell)
        self.schedule_render()
        self.status_callback(f"Pasted {changed} cell(s). Press Ctrl+S to save the workbook.")
        if callable(self.selection_callback):
            self.selection_callback(self)
        return "break"

    def clear_selection(self, _event=None) -> str:
        if not self.editable:
            return "break"
        r1, c1 = self.anchor_cell
        r2, c2 = self.active_cell
        r1, r2 = sorted((r1, r2))
        c1, c2 = sorted((c1, c2))
        for row in range(r1, r2 + 1):
            for col in range(c1, c2 + 1):
                try:
                    self.sheet.set_value(row, col, None)
                except Exception:
                    pass
        self.schedule_render()
        self._report_cell()
        return "break"

    def _refresh_dimensions_if_needed(self) -> None:
        # Pasting/editing can extend the used range. Rebuild only when needed.
        if self.active_cell[0] > self.rows.count or self.active_cell[1] > self.columns.count:
            self._build_metrics()
            self.body.configure(scrollregion=(0, 0, self.columns.total, self.rows.total))
            self.col_header.configure(scrollregion=(0, 0, self.columns.total, 28))
            self.row_header.configure(scrollregion=(0, 0, 52, self.rows.total))

    def _right_click(self, event) -> str:
        image_index = self._image_from_event(event)
        if image_index is not None:
            return super()._right_click(event)
        cell = self._cell_from_event(event)
        self.active_cell = cell
        self.anchor_cell = cell
        self.selected_image = None
        self.schedule_render()
        self._report_cell()
        menu = tk.Menu(self, tearoff=False)
        if self.editable:
            menu.add_command(label="Edit Cell (F2)", command=self.start_edit)
            menu.add_separator()
        menu.add_command(label="Copy Cell / Range", command=self.copy_selection)
        if self.editable:
            menu.add_command(label="Paste", command=self.paste_selection)
            menu.add_command(label="Clear", command=self.clear_selection)
        menu.tk_popup(event.x_root, event.y_root)
        return "break"


class ExcelWorkspaceApp(tk.Tk):
    """Single-window workflow: many Excel files on the left, full workbook on the right."""

    def __init__(self, initial_path: Optional[Path] = None) -> None:
        _set_windows_app_user_model_id()
        super().__init__()
        self.title(f"xViewer {APP_VERSION}")
        self.minsize(1100, 650)
        self.geometry(_centered_geometry(self.winfo_screenwidth(), self.winfo_screenheight(), 1580, 920))
        self._set_icon()

        self.current_dir = self._initial_directory(initial_path)
        self.recent_dirs: list[Path] = self._initial_recent_dirs()
        self.history: list[Path] = []
        self.history_index = -1
        self._item_paths: dict[str, Path] = {}
        self._sort_key = "name"
        self._sort_reverse = False
        self._load_job: Optional[str] = None
        self._load_generation = 0
        self._loading_path: Optional[Path] = None
        self.model: Optional[EditableWorkbookModel] = None
        self.path: Optional[Path] = None
        self.sheet_views: list[EditableVirtualSheet] = []
        self.sheet_buttons: list[ttk.Button] = []
        self._last_selected_path: Optional[Path] = None
        self._active_panel = "left"
        self._pane_nav_bindtag = "xExcelPaneNavigation"
        self.quick_links = _load_quick_links()
        self._quick_link_buttons: list[ttk.Button] = []
        self._style: Optional[ttk.Style] = None
        self._theme_poll_job: Optional[str] = None
        self._effective_theme = ""
        self._directory_records: list[tuple[Path, bool, int, float]] = []
        self._records_dir: Optional[Path] = None
        self._filter_job: Optional[str] = None

        # mDIR-style right-button drag selection for the LEFT file list.
        # Right-drag is a "paint selection" gesture: every row crossed becomes
        # selected at most once per gesture. A row that was already selected
        # (especially the starting row) stays selected instead of being toggled off.
        self._right_drag_active = False
        self._right_drag_seen: set[str] = set()
        self._right_drag_last_index: Optional[int] = None
        self._right_drag_pointer_y = 0
        self._right_drag_scroll_job: Optional[str] = None

        self.path_var = tk.StringVar(value=str(self.current_dir))
        self.theme_mode = tk.StringVar(value=self._initial_theme_mode())
        self.filter_var = tk.StringVar(value="")
        initial_suffix = Path(initial_path).suffix.lower() if initial_path is not None and Path(initial_path).is_file() else ""
        self.excel_only = tk.BooleanVar(value=initial_suffix not in PDF_EXTENSIONS and initial_suffix not in IMAGE_EXTENSIONS)
        self.pdf_only = tk.BooleanVar(value=initial_suffix in PDF_EXTENSIONS)
        self.images_only = tk.BooleanVar(value=initial_suffix in IMAGE_EXTENSIONS)
        self.viewer_kind = "blank"
        self.zoom_value = tk.StringVar(value="100%")
        self.status_text = tk.StringVar(value="Select an Excel, PDF, or image file on the left. It will appear on the right.")
        self.cell_ref = tk.StringVar(value="A1")
        self.formula_value = tk.StringVar(value="")
        self.workbook_title = tk.StringVar(value="No file selected")

        self._build_ui()
        self._install_pane_nav_bindings()
        self._schedule_theme_poll()
        self.bind_all("<Control-l>", self._edit_path_dialog, add="+")
        self.bind_all("<Control-L>", self._edit_path_dialog, add="+")
        self.bind_all("<Control-f>", self._focus_file_search, add="+")
        self.bind_all("<Control-F>", self._focus_file_search, add="+")
        self._push_history(self.current_dir)
        self._remember_directory(self.current_dir, save=False)
        self.refresh_files()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(180, self._activate_left_panel)

        if initial_path is not None and initial_path.is_file():
            suffix = initial_path.suffix.lower()
            if suffix in EXCEL_EXTENSIONS:
                self.after(150, lambda: self.open_workbook(initial_path))
            elif suffix in PDF_EXTENSIONS:
                self.after(150, lambda: self.open_media(initial_path, "pdf"))
            elif suffix in IMAGE_EXTENSIONS:
                self.after(150, lambda: self.open_media(initial_path, "image"))

    def _set_icon(self) -> None:
        try:
            assets = Path(__file__).resolve().parent / "assets"
            icon_path = assets / "xviewer.ico"
            png_path = assets / "xviewer-icon.png"
            if os.name == "nt" and icon_path.exists():
                self.iconbitmap(default=str(icon_path))
            elif png_path.exists():
                self._window_icon_image = tk.PhotoImage(file=str(png_path))
                self.iconphoto(True, self._window_icon_image)
        except Exception:
            pass

    def _initial_directory(self, initial_path: Optional[Path]) -> Path:
        if initial_path is not None:
            p = Path(initial_path).expanduser()
            if p.is_file():
                return p.parent
            if p.is_dir():
                return p
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            saved = Path(str(data.get("last_dir", "")))
            if saved.is_dir():
                return saved
        except Exception:
            pass
        return Path.home()

    def _initial_recent_dirs(self) -> list[Path]:
        values: list[Path | str] = []
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            raw = data.get("recent_dirs", [])
            if isinstance(raw, list):
                values.extend(str(item) for item in raw if str(item).strip())
        except Exception:
            pass
        return _merge_recent_directories(self.current_dir, values)

    def _initial_theme_mode(self) -> str:
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return normalize_theme_mode(data.get("theme", "Bright"))
        except Exception:
            return "Bright"

    def _configure_theme_styles(self) -> None:
        palette = dict(palette_for(self.theme_mode.get()))
        effective = str(palette.get("name", effective_theme_name(self.theme_mode.get())))
        self._effective_theme = effective
        style = self._style or ttk.Style(self)
        self._style = style
        try:
            if effective == "Dark":
                style.theme_use("clam")
            else:
                style.theme_use("vista" if os.name == "nt" else "clam")
        except Exception:
            try:
                style.theme_use("clam")
            except Exception:
                pass

        bg = palette["window"]
        panel = palette["panel"]
        field = palette["field"]
        fg = palette["foreground"]
        muted = palette["muted"]
        button = palette["button"]
        active = palette["button_active"]
        border = palette["border"]
        selection = palette["selection"]
        selection_fg = palette["selection_foreground"]

        self.configure(background=bg)
        style.configure("TFrame", background=bg)
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("Muted.TLabel", background=bg, foreground=muted)
        style.configure("TButton", background=button, foreground=fg, bordercolor=border, focusthickness=1, focuscolor=border)
        style.map("TButton", background=[("active", active), ("pressed", selection)], foreground=[("pressed", selection_fg)])
        style.configure("TCheckbutton", background=bg, foreground=fg)
        style.map("TCheckbutton", background=[("active", bg)], foreground=[("active", fg)])
        style.configure("TEntry", fieldbackground=field, foreground=fg, bordercolor=border, insertcolor=fg)
        style.configure("TCombobox", fieldbackground=field, background=button, foreground=fg, arrowcolor=fg, bordercolor=border)
        style.map("TCombobox", fieldbackground=[("readonly", field)], foreground=[("readonly", fg)], selectbackground=[("readonly", selection)], selectforeground=[("readonly", selection_fg)])
        style.configure("Treeview", background=field, fieldbackground=field, foreground=fg, bordercolor=border, lightcolor=border, darkcolor=border)
        style.map("Treeview", background=[("selected", selection)], foreground=[("selected", selection_fg)])
        style.configure("Treeview.Heading", background=button, foreground=fg, relief="flat", bordercolor=border)
        style.map("Treeview.Heading", background=[("active", active)])
        style.configure("TLabelframe", background=bg, foreground=fg, bordercolor=border, lightcolor=border, darkcolor=border)
        style.configure("TLabelframe.Label", background=bg, foreground=fg)
        style.configure("TNotebook", background=bg, bordercolor=border)
        style.configure("TNotebook.Tab", background=button, foreground=fg, padding=(8, 4))
        style.map("TNotebook.Tab", background=[("selected", field), ("active", active)], foreground=[("selected", fg)])
        style.configure("TScrollbar", background=button, troughcolor=panel, bordercolor=border, arrowcolor=fg)

        # Tk widgets and future context menus/listboxes do not use ttk styles.
        self.option_add("*background", bg)
        self.option_add("*foreground", fg)
        self.option_add("*Menu.background", field)
        self.option_add("*Menu.foreground", fg)
        self.option_add("*Menu.activeBackground", selection)
        self.option_add("*Menu.activeForeground", selection_fg)
        self.option_add("*TCombobox*Listbox.background", field)
        self.option_add("*TCombobox*Listbox.foreground", fg)
        self.option_add("*TCombobox*Listbox.selectBackground", selection)
        self.option_add("*TCombobox*Listbox.selectForeground", selection_fg)
        self._apply_windows_titlebar(effective == "Dark")

    def _apply_windows_titlebar(self, dark: bool, window=None) -> None:
        if os.name != "nt":
            return
        try:
            import ctypes
            target = window or self
            target.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(target.winfo_id())
            value = ctypes.c_int(1 if dark else 0)
            # Windows 10/11 use attribute 20; older builds used 19.
            for attribute in (20, 19):
                try:
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attribute, ctypes.byref(value), ctypes.sizeof(value))
                    break
                except Exception:
                    continue
        except Exception:
            pass

    def _apply_theme_runtime(self, *, save: bool = False) -> None:
        self._configure_theme_styles()
        palette = dict(palette_for(self.theme_mode.get()))
        try:
            self.main_pane.configure(bg=palette["divider"])
        except Exception:
            pass
        try:
            self.path_bar.configure(
                bg=palette["path_bg"],
                highlightbackground=palette["path_border"],
                highlightcolor=palette["path_focus"],
            )
            self.path_bar_inner.configure(bg=palette["path_bg"])
            self._rebuild_path_bar()
        except Exception:
            pass
        try:
            self.sheet_strip_canvas.configure(bg=palette["window"])
        except Exception:
            pass
        for view in list(self.sheet_views):
            try:
                view.apply_ui_theme(palette)
            except Exception:
                pass
        try:
            self.media_view.apply_palette(palette)
        except Exception:
            pass
        if save:
            self._save_config()
        self._schedule_theme_poll()

    def _theme_changed(self, _event=None) -> None:
        self.theme_mode.set(normalize_theme_mode(self.theme_mode.get()))
        self._apply_theme_runtime(save=True)
        self.status_text.set(f"Theme: {self.theme_mode.get()} ({self._effective_theme})")

    def _schedule_theme_poll(self) -> None:
        if self._theme_poll_job is not None:
            try:
                self.after_cancel(self._theme_poll_job)
            except Exception:
                pass
            self._theme_poll_job = None
        if normalize_theme_mode(self.theme_mode.get()) == "System":
            self._theme_poll_job = self.after(1500, self._poll_system_theme)

    def _poll_system_theme(self) -> None:
        self._theme_poll_job = None
        if normalize_theme_mode(self.theme_mode.get()) != "System":
            return
        effective = effective_theme_name("System")
        if effective != self._effective_theme:
            self._apply_theme_runtime(save=False)
        else:
            self._schedule_theme_poll()

    def _build_ui(self) -> None:
        self.option_add("*Font", ("Segoe UI", 10))
        self._style = ttk.Style(self)
        self._configure_theme_styles()
        palette = dict(palette_for(self.theme_mode.get()))

        # mDIR-style top shortcut strip. Every configured link is a first-class
        # button (Folder / File / Program / Web / Action / Command), in the
        # same order as mDIR's Link Manager.
        quick = ttk.Frame(self, padding=(6, 5, 6, 3))
        quick.pack(fill="x")
        ttk.Button(quick, text="Edit Links", command=self._edit_quick_links).pack(side="left", padx=(0, 4))
        self.quick_links_frame = ttk.Frame(quick)
        self.quick_links_frame.pack(side="left", fill="x", expand=True)
        ttk.Button(quick, text="Reload", command=self._reload_quick_links).pack(side="right", padx=(4, 0))
        self.theme_combo = ttk.Combobox(quick, textvariable=self.theme_mode, values=THEME_CHOICES, state="readonly", width=8)
        self.theme_combo.pack(side="right", padx=(3, 0))
        self.theme_combo.bind("<<ComboboxSelected>>", self._theme_changed)
        ttk.Label(quick, text="Theme:").pack(side="right", padx=(8, 0))
        self._rebuild_quick_links()

        # Use the classic Tk PanedWindow here rather than ttk.Panedwindow.
        # On some Windows/Tk builds the themed paned window could start with
        # the left child collapsed to zero width.  minsize + an explicit sash
        # position guarantees that BOTH work areas are visible at startup.
        pane = tk.PanedWindow(
            self,
            orient=tk.HORIZONTAL,
            # Keep the divider easy to drag without making it visually heavy.
            # The previous raised 7 px sash looked like a dark border between
            # FILES and WORKBOOK on Windows. A flat, light 4 px sash keeps the
            # resize affordance while blending with the surrounding UI.
            sashrelief=tk.FLAT,
            sashwidth=4,
            sashpad=0,
            bd=0,
            relief=tk.FLAT,
            bg=palette["divider"],
        )
        pane.pack(fill="both", expand=True, padx=6, pady=(0, 4))
        self.main_pane = pane

        left = ttk.LabelFrame(pane, text="LEFT — Files", padding=(5, 4))
        right = ttk.LabelFrame(pane, text="RIGHT — Workbook / PDF / Image Viewer", padding=(5, 4))
        self.left_panel = left
        self.right_panel = right
        pane.add(left, minsize=340, stretch="always")
        pane.add(right, minsize=520, stretch="always")
        self.after_idle(self._position_main_sash)
        self.after(250, self._position_main_sash)

        # LEFT: mDIR-style navigation controls and single full-path strip.
        # The green path looks like mDIR, but each directory name remains
        # clickable so the user can jump directly to any parent folder.
        left_nav = ttk.Frame(left)
        left_nav.pack(fill="x", pady=(0, 3))
        ttk.Button(left_nav, text="Back", width=6, command=self.go_back).pack(side="left", padx=(0, 2))
        ttk.Button(left_nav, text="Up", width=5, command=self.go_up).pack(side="left", padx=2)
        self._add_drive_buttons(left_nav)
        ttk.Button(left_nav, text="Refresh", command=self.refresh_files).pack(side="right", padx=(4, 0))

        path_green = palette["path_bg"]
        path_row = ttk.Frame(left)
        path_row.pack(fill="x", pady=(0, 4))
        self.path_bar = tk.Canvas(
            path_row,
            height=28,
            bg=path_green,
            relief="flat",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=palette["path_border"],
            highlightcolor=palette["path_focus"],
        )
        self.path_bar.pack(side="left", fill="x", expand=True)
        self.path_history_button = ttk.Button(
            path_row, text="▼", width=3, command=self._show_recent_directories
        )
        self.path_history_button.pack(side="right", padx=(3, 0))
        self.path_bar_inner = tk.Frame(self.path_bar, bg=path_green)
        self.path_bar_window = self.path_bar.create_window(
            7, 3, anchor="nw", window=self.path_bar_inner
        )
        self.path_bar.bind("<Double-Button-1>", self._edit_path_dialog)
        self.path_bar.bind("<Configure>", lambda _e: self._update_path_bar_scrollregion())
        self._rebuild_path_bar()

        left_tools = ttk.Frame(left)
        left_tools.pack(fill="x", pady=(0, 2))
        ttk.Label(left_tools, text="Files", font=("Segoe UI", 10, "bold")).pack(side="left")
        self.clear_filter_button = ttk.Button(left_tools, text="×", width=3, command=self._clear_file_search)
        self.clear_filter_button.pack(side="right", padx=(3, 0))
        self.search_entry = ttk.Entry(left_tools, textvariable=self.filter_var, width=24)
        self.search_entry.pack(side="right")
        ttk.Label(left_tools, text="Search:").pack(side="right", padx=(0, 4))

        type_filters = ttk.Frame(left)
        type_filters.pack(fill="x", pady=(0, 4))
        ttk.Label(type_filters, text="Show:").pack(side="left")
        ttk.Checkbutton(
            type_filters, text="Excel only", variable=self.excel_only,
            command=lambda: self._type_filter_changed("excel")
        ).pack(side="left", padx=(6, 0))
        ttk.Checkbutton(
            type_filters, text="PDF only", variable=self.pdf_only,
            command=lambda: self._type_filter_changed("pdf")
        ).pack(side="left", padx=(8, 0))
        ttk.Checkbutton(
            type_filters, text="Images only", variable=self.images_only,
            command=lambda: self._type_filter_changed("image")
        ).pack(side="left", padx=(8, 0))
        self.search_entry.bind("<Return>", self._search_enter)
        self.search_entry.bind("<Escape>", self._clear_file_search)
        # After typing a filename filter, Up/Down should immediately leave the
        # search box and continue browsing the filtered FILES list.
        self.search_entry.bind("<Up>", self._search_arrow_to_files)
        self.search_entry.bind("<Down>", self._search_arrow_to_files)
        self.filter_var.trace_add("write", self._schedule_filter_refresh)

        cols = ("name", "ext", "size", "modified")
        # Keep the file table inside its own grid container.  Mixing left/right
        # and bottom packing in one parent can produce surprising geometry on
        # some Windows themes; the grid below is deterministic.
        file_area = ttk.Frame(left)
        file_area.pack(fill="both", expand=True)
        file_area.columnconfigure(0, weight=1)
        file_area.rowconfigure(0, weight=1)

        self.file_tree = ttk.Treeview(file_area, columns=cols, show="headings", selectmode="extended")
        self.file_tree.heading("name", text="Name", command=lambda: self.sort_files("name"))
        self.file_tree.heading("ext", text="Ext", command=lambda: self.sort_files("ext"))
        self.file_tree.heading("size", text="Size", command=lambda: self.sort_files("size"))
        self.file_tree.heading("modified", text="Modified", command=lambda: self.sort_files("modified"))
        self.file_tree.column("name", width=250, minwidth=120, anchor="w")
        self.file_tree.column("ext", width=65, minwidth=45, anchor="center")
        self.file_tree.column("size", width=95, minwidth=70, anchor="e")
        self.file_tree.column("modified", width=145, minwidth=120, anchor="center")
        ybar = ttk.Scrollbar(file_area, orient="vertical", command=self.file_tree.yview)
        xbar = ttk.Scrollbar(file_area, orient="horizontal", command=self.file_tree.xview)
        self.file_tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.file_tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        self.file_tree.bind("<<TreeviewSelect>>", self._file_selected)
        self.file_tree.bind("<Double-Button-1>", self._file_activated)
        # mDIR-style right-button drag: press on a row, keep the right button
        # down, and drag across rows to toggle them into/out of the selection.
        # Fast motion fills skipped intermediate rows and edge dragging scrolls.
        self.file_tree.bind("<ButtonPress-3>", self._file_right_drag_start)
        self.file_tree.bind("<B3-Motion>", self._file_right_drag_motion)
        self.file_tree.bind("<ButtonRelease-3>", self._file_right_drag_end)
        # Selection alone loads Excel in the RIGHT pane. Enter is deliberately
        # reserved for opening the original file with its OS-associated app
        # (normally Microsoft Excel on Windows).
        self.file_tree.bind("<Return>", self._file_enter_open_default)
        # Do not let Tab leave the LEFT file list and wander through toolbar
        # controls. Pane switching already has dedicated shortcuts, and Tab is
        # reserved for Excel-style cell movement in the RIGHT workbook.
        self.file_tree.bind("<Tab>", self._block_left_file_tab)
        self.file_tree.bind("<Shift-Tab>", self._block_left_file_tab)
        self.file_tree.bind("<ISO_Left_Tab>", self._block_left_file_tab)
        self.file_tree.bind("<Delete>", self._delete_selected_items)
        self.file_tree.bind("<KP_Delete>", self._delete_selected_items)
        self.file_tree.bind("<BackSpace>", lambda _e: self.go_up())
        self.file_tree.bind("<F5>", lambda _e: self.refresh_files(rescan=True))
        for key in ("<Up>", "<Down>", "<Home>", "<End>", "<Prior>", "<Next>"):
            self.file_tree.bind(key, self._left_navigation_key)

        # RIGHT: workbook editor. Keep the action buttons on their own top
        # line and show the workbook name on a dedicated second line. This
        # prevents long filenames from squeezing the editing controls and
        # makes the right pane easier to scan.
        editor_toolbar = ttk.Frame(right, padding=(0, 0, 0, 3))
        editor_toolbar.pack(fill="x")
        self.save_btn = ttk.Button(editor_toolbar, text="Save  Ctrl+S", command=self.save_workbook, state="disabled")
        self.save_btn.pack(side="left", padx=(0, 2))
        self.save_as_btn = ttk.Button(editor_toolbar, text="Save As", command=self.save_as, state="disabled")
        self.save_as_btn.pack(side="left", padx=2)
        self.copy_btn = ttk.Button(editor_toolbar, text="Copy", command=self.copy_current, state="disabled")
        self.copy_btn.pack(side="left", padx=(8, 2))
        self.paste_btn = ttk.Button(editor_toolbar, text="Paste", command=self.paste_current, state="disabled")
        self.paste_btn.pack(side="left", padx=2)
        self.copy_image_btn = ttk.Button(editor_toolbar, text="Copy Image", command=self.copy_image_current, state="disabled")
        self.copy_image_btn.pack(side="left", padx=2)
        self.open_btn = ttk.Button(editor_toolbar, text="Open", command=self.open_default, state="disabled")
        self.open_btn.pack(side="left", padx=(8, 2))
        ttk.Label(editor_toolbar, text="Zoom:").pack(side="left", padx=(12, 2))
        zoom_box = ttk.Combobox(editor_toolbar, textvariable=self.zoom_value, values=["50%", "75%", "90%", "100%", "110%", "125%", "150%", "175%", "200%"], width=6, state="readonly")
        zoom_box.pack(side="left")
        zoom_box.bind("<<ComboboxSelected>>", lambda _e: self._zoom_changed())
        self.fit_1to1_btn = ttk.Button(
            editor_toolbar,
            text="1:1 / Fit",
            command=self._toggle_media_fit_one_to_one,
            state="disabled",
        )
        self.fit_1to1_btn.pack(side="left", padx=(4, 0))

        # Dedicated workbook-name row directly under the buttons. The label
        # expands across the pane so long buyer filenames remain readable.
        workbook_name_row = ttk.Frame(right, padding=(0, 0, 0, 4))
        workbook_name_row.pack(fill="x")
        self.workbook_name_label = ttk.Label(
            workbook_name_row,
            textvariable=self.workbook_title,
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        )
        self.workbook_name_label.pack(side="left", fill="x", expand=True)

        formula = ttk.Frame(right, padding=(0, 0, 0, 4))
        formula.pack(fill="x")
        self.formula_frame = formula
        ttk.Label(formula, textvariable=self.cell_ref, width=9, anchor="center", relief="sunken").pack(side="left", padx=(0, 4))
        self.formula_entry = ttk.Entry(formula, textvariable=self.formula_value)
        self.formula_entry.pack(side="left", fill="x", expand=True)
        self.formula_entry.bind("<Return>", self._formula_commit)
        self.formula_entry.bind("<Escape>", lambda _e: self._sync_formula_bar())

        # Always-accessible sheet strip.  ttk.Notebook tabs remain below, while
        # this strip contains every worksheet name and can scroll horizontally.
        sheet_nav = ttk.Frame(right, padding=(0, 0, 0, 4))
        sheet_nav.pack(fill="x")
        self.sheet_nav_frame = sheet_nav
        self.sheet_count_text = tk.StringVar(value="Sheets: 0")
        ttk.Label(sheet_nav, textvariable=self.sheet_count_text, width=11).pack(side="left")
        ttk.Button(sheet_nav, text="◀", width=3, command=lambda: self._step_sheet(-1)).pack(side="left", padx=(0, 1))
        ttk.Button(sheet_nav, text="▶", width=3, command=lambda: self._step_sheet(1)).pack(side="left", padx=(0, 4))
        sheet_strip_holder = ttk.Frame(sheet_nav)
        sheet_strip_holder.pack(side="left", fill="x", expand=True)
        self.sheet_strip_canvas = tk.Canvas(sheet_strip_holder, height=32, highlightthickness=0, borderwidth=0, bg=palette["window"])
        self.sheet_strip_scroll = ttk.Scrollbar(sheet_strip_holder, orient="horizontal", command=self.sheet_strip_canvas.xview)
        self.sheet_strip_canvas.configure(xscrollcommand=self.sheet_strip_scroll.set)
        self.sheet_strip_canvas.pack(fill="x", expand=True)
        self.sheet_strip_scroll.pack(fill="x")
        self.sheet_strip_inner = ttk.Frame(self.sheet_strip_canvas)
        self._sheet_strip_window = self.sheet_strip_canvas.create_window((0, 0), window=self.sheet_strip_inner, anchor="nw")
        self.sheet_strip_inner.bind("<Configure>", self._sheet_strip_configured)
        self.sheet_strip_canvas.bind("<Shift-MouseWheel>", self._sheet_strip_mousewheel)

        self.notebook = ttk.Notebook(right)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", lambda _e: self._sheet_changed())

        self.blank = ttk.Frame(self.notebook)
        ttk.Label(
            self.blank,
            text=(
                "Select an Excel, PDF, or image file on the left.\n\n"
                "Excel files open as full editable workbooks.\n"
                "PDF files open page-by-page, and image files open in the media viewer.\n"
                "Ctrl+Left goes to FILES; Ctrl+Right goes to the RIGHT viewer.\n"
                "Alt+1 jumps to FILES; Alt+2 jumps to the RIGHT viewer.\n"
                "Arrow keys select files on the left and move cells in Excel on the right.\n"
                "Tab / Shift+Tab moves to the next / previous Excel cell."
            ),
            anchor="center",
            justify="center",
            font=("Segoe UI", 13),
        ).pack(fill="both", expand=True)
        self.notebook.add(self.blank, text="Workbook")

        # PDF/Image controls occupy the same RIGHT content area as the Excel
        # formula/sheet/notebook stack. They are packed only in media mode.
        self.media_nav = ttk.Frame(right, padding=(0, 0, 0, 4))
        self.media_type_text = tk.StringVar(value="")
        self.media_page_text = tk.StringVar(value="")
        ttk.Label(self.media_nav, textvariable=self.media_type_text, width=10).pack(side="left")
        self.media_first_btn = ttk.Button(self.media_nav, text="|◀", width=3, command=lambda: self._media_first_page())
        self.media_first_btn.pack(side="left", padx=(0, 1))
        self.media_prev_btn = ttk.Button(self.media_nav, text="◀", width=3, command=lambda: self._media_step_page(-1))
        self.media_prev_btn.pack(side="left", padx=(0, 1))
        ttk.Label(self.media_nav, textvariable=self.media_page_text, width=16, anchor="center").pack(side="left", padx=4)
        self.media_next_btn = ttk.Button(self.media_nav, text="▶", width=3, command=lambda: self._media_step_page(1))
        self.media_next_btn.pack(side="left", padx=(0, 1))
        self.media_last_btn = ttk.Button(self.media_nav, text="▶|", width=3, command=lambda: self._media_last_page())
        self.media_last_btn.pack(side="left", padx=(0, 8))
        self.media_save_image_btn = ttk.Button(self.media_nav, text="Save Image As", command=self._media_save_image_as)
        self.media_save_image_btn.pack(side="left", padx=(8, 0))

        self.media_view = MediaViewer(
            right,
            status_callback=self._set_status,
            page_callback=self._media_page_changed,
            ui_palette=palette,
            zoom_callback=self._media_zoom_changed,
        )
        self._apply_pane_nav_bindtag(self.media_view)

        status = ttk.Label(self, textvariable=self.status_text, anchor="w", relief="sunken", padding=(7, 4))
        status.pack(fill="x", side="bottom")

        self.bind_all("<Control-s>", lambda _e: self.save_workbook())
        self.bind_all("<Control-S>", lambda _e: self.save_workbook())
        self.bind_all("<Control-plus>", lambda _e: self.adjust_zoom(0.1))
        self.bind_all("<Control-equal>", lambda _e: self.adjust_zoom(0.1))
        self.bind_all("<Control-minus>", lambda _e: self.adjust_zoom(-0.1))
        self.bind_all("<Control-0>", lambda _e: self.set_zoom(1.0))
        # Dedicated main-pane shortcuts.  Do not steal Tab from Excel cells.
        # Ctrl+Left / Ctrl+Right are the primary directional pane shortcuts.
        # Alt+1 / Alt+2 remain direct pane shortcuts, and F6 stays as a legacy toggle.
        self.bind_all("<Control-Left>", self._shortcut_left_panel)
        self.bind_all("<Control-Right>", self._shortcut_right_panel)
        self.bind_all("<Alt-KeyPress-1>", self._shortcut_alt_left_panel)
        self.bind_all("<Alt-KeyPress-2>", self._shortcut_alt_right_panel)
        self.bind_all("<F6>", self._toggle_panel)
        self.bind_all("<Control-Prior>", lambda _e: self._step_sheet(-1))
        self.bind_all("<Control-Next>", lambda _e: self._step_sheet(1))
        self.bind_all("<Button-1>", self._track_panel_click, add="+")

    def _block_left_file_tab(self, _event=None) -> str:
        """Keep keyboard focus in LEFT FILES when Tab/Shift+Tab is pressed."""
        self._active_panel = "left"
        self._set_panel_indicator("left")
        self.status_text.set(
            "Tab is disabled in LEFT FILES | Use Ctrl+Right or Alt+2 for WORKBOOK"
        )
        return _left_file_tab_action()

    def _install_pane_nav_bindings(self) -> None:
        """Install pane shortcuts before normal widget/class key handling.

        Windows widgets such as Canvas, Entry and Treeview may consume arrow
        keys before an ``all`` binding is reached.  A custom bindtag is placed
        FIRST on every widget and receives a generic KeyPress.  On Windows the
        dispatcher also checks GetKeyState so Ctrl/Alt shortcuts do not depend
        on Tk's modifier-mask quirks.
        """
        tag = self._pane_nav_bindtag
        self.bind_class(tag, "<KeyPress>", self._pane_hotkey_dispatch)
        # Exact sequences are kept as an additional fast path/fallback.
        self.bind_class(tag, "<Control-KeyPress-Left>", self._shortcut_left_panel)
        self.bind_class(tag, "<Control-KeyPress-Right>", self._shortcut_right_panel)
        self.bind_class(tag, "<Alt-KeyPress-1>", self._shortcut_alt_left_panel)
        self.bind_class(tag, "<Alt-KeyPress-2>", self._shortcut_alt_right_panel)
        self._apply_pane_nav_bindtag(self)

    @staticmethod
    def _win_modifier_down(vk: int) -> bool:
        if os.name != "nt":
            return False
        try:
            import ctypes

            # GetAsyncKeyState reflects the physical key *now* and avoids stale
            # modifier bits that can be present in Tk events on some layouts.
            return bool(ctypes.windll.user32.GetAsyncKeyState(int(vk)) & 0x8000)
        except Exception:
            return False

    def _pane_hotkey_dispatch(self, event) -> Optional[str]:
        if os.name == "nt":
            ctrl_physical: Optional[bool] = self._win_modifier_down(0x11)  # VK_CONTROL
            alt_physical: Optional[bool] = self._win_modifier_down(0x12)   # VK_MENU / Alt
        else:
            ctrl_physical = None
            alt_physical = None

        action = _pane_hotkey_action(
            getattr(event, "keysym", ""),
            getattr(event, "keycode", 0),
            getattr(event, "state", 0),
            ctrl_down=ctrl_physical,
            alt_down=alt_physical,
        )
        if action == "left":
            return self._shortcut_left_panel(event)
        if action == "right":
            return self._shortcut_right_panel(event)
        return None

    def _alt_shortcut_is_really_down(self, event) -> bool:
        """Validate Alt for Alt+1/Alt+2 without trusting stale Windows masks."""
        if os.name == "nt":
            return self._win_modifier_down(0x12)
        return bool(int(getattr(event, "state", 0) or 0) & 0x0008)

    def _shortcut_alt_left_panel(self, event=None) -> Optional[str]:
        if event is not None and not self._alt_shortcut_is_really_down(event):
            return None
        return self._shortcut_left_panel(event)

    def _shortcut_alt_right_panel(self, event=None) -> Optional[str]:
        if event is not None and not self._alt_shortcut_is_really_down(event):
            return None
        return self._shortcut_right_panel(event)

    def _bind_pane_shortcuts_direct(self, widget) -> None:
        """Second line of defense for Windows/Tk focus edge cases."""
        try:
            widget.bind("<Control-KeyPress-Left>", self._shortcut_left_panel, add="+")
            widget.bind("<Control-KeyPress-Right>", self._shortcut_right_panel, add="+")
            widget.bind("<Alt-KeyPress-1>", self._shortcut_alt_left_panel, add="+")
            widget.bind("<Alt-KeyPress-2>", self._shortcut_alt_right_panel, add="+")
        except Exception:
            pass

    def _apply_pane_nav_bindtag(self, widget) -> None:
        try:
            tags = list(widget.bindtags())
            tag = self._pane_nav_bindtag
            if tag in tags:
                tags.remove(tag)
            widget.bindtags((tag, *tags))
            self._bind_pane_shortcuts_direct(widget)
            for child in widget.winfo_children():
                self._apply_pane_nav_bindtag(child)
        except Exception:
            pass

    def _position_main_sash(self) -> None:
        """Give both panes a safe visible width on Windows and Linux."""
        pane = getattr(self, "main_pane", None)
        if pane is None:
            return
        try:
            self.update_idletasks()
            width = int(pane.winfo_width())
            if width <= 10:
                self.after(100, self._position_main_sash)
                return
            # Roughly one third for the file browser, while guaranteeing enough
            # room for the workbook and keeping the file list useful.
            target = max(360, min(540, int(width * 0.34)))
            target = min(target, max(340, width - 520))
            pane.sash_place(0, target, 1)
        except Exception:
            pass

    @staticmethod
    def _is_descendant(widget, parent) -> bool:
        current = widget
        while current is not None:
            if current is parent:
                return True
            try:
                name = current.winfo_parent()
                current = current.nametowidget(name) if name else None
            except Exception:
                return False
        return False

    def _track_panel_click(self, event) -> None:
        """Remember which main panel the user clicked without stealing child focus."""
        try:
            if self._is_descendant(event.widget, self.left_panel):
                self._active_panel = "left"
                self._set_panel_indicator("left")
            elif self._is_descendant(event.widget, self.right_panel):
                self._active_panel = "right"
                self._set_panel_indicator("right")
        except Exception:
            pass

    def _ensure_left_selection(self) -> None:
        children = self.file_tree.get_children()
        if not children:
            return
        selected = self.file_tree.selection()
        iid = selected[0] if selected else children[0]
        self.file_tree.selection_set(iid)
        self.file_tree.focus(iid)
        self.file_tree.see(iid)

    def _set_panel_indicator(self, active: str) -> None:
        try:
            self.left_panel.configure(text="LEFT — Files  [ACTIVE]" if active == "left" else "LEFT — Files")
            self.right_panel.configure(text="RIGHT — Workbook / PDF / Image Viewer  [ACTIVE]" if active == "right" else "RIGHT — Workbook / PDF / Image Viewer")
        except Exception:
            pass

    def _activate_left_panel(self) -> None:
        self._active_panel = "left"
        self._set_panel_indicator("left")
        self._ensure_left_selection()
        self.file_tree.focus_set()
        self.status_text.set("ACTIVE: LEFT FILES | Up/Down selects a file | Ctrl+Right = VIEWER | Alt+2 = VIEWER")

    def _activate_right_panel(self) -> None:
        self._active_panel = "right"
        self._set_panel_indicator("right")
        if self.viewer_kind in {"pdf", "image"}:
            self.media_view.canvas.focus_set()
            self.status_text.set("ACTIVE: RIGHT VIEWER | Ctrl+Left = FILES | Alt+1 = FILES")
            return
        view = self.current_view()
        if view is not None:
            view.body.focus_set()
            view._report_cell()
        else:
            self.notebook.focus_set()
            self.status_text.set("ACTIVE: RIGHT WORKBOOK | Ctrl+Left = FILES | Alt+1 = FILES")

    def _toggle_panel(self, _event=None) -> str:
        # Legacy F6 pane toggle; Tab remains available to Excel.
        if self._active_panel == "right":
            self._commit_pending_edits()
            self._activate_left_panel()
        else:
            self._activate_right_panel()
        return "break"

    def _shortcut_left_panel(self, _event=None) -> str:
        if self._active_panel == "right":
            self._commit_pending_edits()
        self._activate_left_panel()
        return "break"

    def _shortcut_right_panel(self, _event=None) -> str:
        self._activate_right_panel()
        return "break"

    def _cancel_file_right_drag(self, *, finalize: bool = False) -> None:
        """Stop a LEFT-pane right-button drag gesture and cancel edge scrolling."""
        was_active = self._right_drag_active
        self._right_drag_active = False
        if self._right_drag_scroll_job is not None:
            try:
                self.after_cancel(self._right_drag_scroll_job)
            except Exception:
                pass
            self._right_drag_scroll_job = None
        self._right_drag_seen.clear()
        self._right_drag_last_index = None
        if finalize and was_active:
            self._file_selected()

    def _file_right_drag_row_at(self, y: int) -> str:
        """Return the Treeview row under *y*, including a useful edge fallback."""
        try:
            iid = self.file_tree.identify_row(int(y))
            if iid:
                return iid
            height = max(1, int(self.file_tree.winfo_height()))
            # Near the heading/top edge, identify the first visible body row;
            # near the bottom edge, identify the last visible body row.
            probe = 30 if int(y) < 30 else max(1, height - 4)
            return self.file_tree.identify_row(probe) or ""
        except Exception:
            return ""

    def _file_right_drag_toggle_iid(self, iid: str) -> None:
        """Select one row during a right-drag without deselecting existing rows.

        The method name is kept for compatibility with the existing event path,
        but the gesture is intentionally additive now. This fixes the case
        where a drag starts on an already-selected row: the starting row must
        remain selected while all crossed rows are added to the selection.
        """
        if not iid or iid in self._right_drag_seen:
            return
        self.file_tree.selection_add(iid)
        self._right_drag_seen.add(iid)
        try:
            self.file_tree.focus(iid)
            self.file_tree.see(iid)
        except Exception:
            pass

    def _file_right_drag_process_y(self, y: int) -> None:
        """Select the row at y and every skipped row since the last drag event."""
        iid = self._file_right_drag_row_at(y)
        if not iid:
            return
        children = list(self.file_tree.get_children())
        try:
            index = children.index(iid)
        except ValueError:
            return
        previous = self._right_drag_last_index
        if previous is None:
            indices = [index]
        elif index == previous:
            indices = [index]
        else:
            step = 1 if index > previous else -1
            indices = list(range(previous + step, index + step, step))
        for pos in indices:
            if 0 <= pos < len(children):
                self._file_right_drag_toggle_iid(children[pos])
        self._right_drag_last_index = index

    def _file_right_drag_start(self, event) -> str:
        iid = self.file_tree.identify_row(event.y)
        if not iid:
            self._cancel_file_right_drag(finalize=False)
            return "break"
        self._active_panel = "left"
        self._set_panel_indicator("left")
        self.file_tree.focus_set()
        self._cancel_file_right_drag(finalize=False)
        self._right_drag_active = True
        self._right_drag_pointer_y = int(event.y)
        self._file_right_drag_process_y(event.y)
        self.status_text.set(
            "RIGHT-DRAG selection: drag over rows to add them to the selection; release the right mouse button to finish."
        )
        return "break"

    def _file_right_drag_motion(self, event) -> str:
        if not self._right_drag_active:
            return "break"
        self._right_drag_pointer_y = int(event.y)
        self._file_right_drag_process_y(event.y)
        height = max(1, int(self.file_tree.winfo_height()))
        if event.y < 34 or event.y > height - 26:
            if self._right_drag_scroll_job is None:
                self._right_drag_scroll_job = self.after(55, self._file_right_drag_autoscroll_step)
        elif self._right_drag_scroll_job is not None:
            try:
                self.after_cancel(self._right_drag_scroll_job)
            except Exception:
                pass
            self._right_drag_scroll_job = None
        return "break"

    def _file_right_drag_autoscroll_step(self) -> None:
        self._right_drag_scroll_job = None
        if not self._right_drag_active:
            return
        height = max(1, int(self.file_tree.winfo_height()))
        y = int(self._right_drag_pointer_y)
        direction = -1 if y < 34 else (1 if y > height - 26 else 0)
        if not direction:
            return
        try:
            self.file_tree.yview_scroll(direction, "units")
            self.update_idletasks()
        except Exception:
            return
        probe_y = 32 if direction < 0 else max(1, height - 5)
        self._file_right_drag_process_y(probe_y)
        if self._right_drag_active:
            self._right_drag_scroll_job = self.after(55, self._file_right_drag_autoscroll_step)

    def _file_right_drag_end(self, event) -> str:
        if self._right_drag_active:
            self._right_drag_pointer_y = int(event.y)
            self._file_right_drag_process_y(event.y)
        self._cancel_file_right_drag(finalize=True)
        return "break"

    def _left_navigation_key(self, event) -> str:
        """Deterministic file navigation for the left pane."""
        self._active_panel = "left"
        children = list(self.file_tree.get_children())
        if not children:
            return "break"
        selected = self.file_tree.selection()
        current = selected[0] if selected else children[0]
        try:
            index = children.index(current)
        except ValueError:
            index = 0
        key = event.keysym
        page = max(1, int(max(1, self.file_tree.winfo_height()) / 24) - 2)
        if key == "Up":
            index -= 1
        elif key == "Down":
            index += 1
        elif key == "Home":
            index = 0
        elif key == "End":
            index = len(children) - 1
        elif key == "Prior":
            index -= page
        elif key == "Next":
            index += page
        index = max(0, min(len(children) - 1, index))
        iid = children[index]
        self.file_tree.selection_set(iid)
        self.file_tree.focus(iid)
        self.file_tree.see(iid)
        # TreeviewSelect normally fires, but call directly so keyboard behavior is
        # consistent on all Windows/Tk builds.  The loader debounces duplicates.
        self._file_selected()
        return "break"

    def _sheet_strip_configured(self, _event=None) -> None:
        try:
            bbox = self.sheet_strip_canvas.bbox("all")
            if bbox:
                self.sheet_strip_canvas.configure(scrollregion=bbox)
        except Exception:
            pass

    def _sheet_strip_mousewheel(self, event) -> str:
        delta = -1 if event.delta > 0 else 1
        self.sheet_strip_canvas.xview_scroll(delta * 3, "units")
        return "break"

    def _rebuild_sheet_strip(self) -> None:
        for child in self.sheet_strip_inner.winfo_children():
            child.destroy()
        self.sheet_buttons.clear()
        count = len(self.sheet_views)
        self.sheet_count_text.set(f"Sheets: {count}")
        for index, view in enumerate(self.sheet_views):
            title = str(view.sheet.title)
            button = ttk.Button(
                self.sheet_strip_inner,
                text=title,
                command=lambda i=index: self._select_sheet_index(i),
                takefocus=False,
            )
            button.pack(side="left", padx=(0, 2), pady=1)
            self._apply_pane_nav_bindtag(button)
            self.sheet_buttons.append(button)
        self.sheet_strip_inner.update_idletasks()
        self._sheet_strip_configured()
        self._sync_sheet_strip()

    def _select_sheet_index(self, index: int) -> None:
        if not self.sheet_views:
            return
        index = max(0, min(len(self.sheet_views) - 1, index))
        self._active_panel = "right"
        self.notebook.select(self.sheet_views[index])
        self._sync_sheet_strip(index)
        self.sheet_views[index].body.focus_set()

    def _step_sheet(self, delta: int):
        if not self.sheet_views:
            return "break"
        current = self.current_view()
        try:
            index = self.sheet_views.index(current) if current is not None else 0
        except ValueError:
            index = 0
        index = (index + delta) % len(self.sheet_views)
        self._select_sheet_index(index)
        return "break"

    def _sync_sheet_strip(self, index: Optional[int] = None) -> None:
        if not self.sheet_views or not self.sheet_buttons:
            return
        if index is None:
            current = self.current_view()
            try:
                index = self.sheet_views.index(current) if current is not None else 0
            except ValueError:
                index = 0
        for i, button in enumerate(self.sheet_buttons):
            button.state(["disabled"] if i == index else ["!disabled"])
        try:
            self.sheet_strip_inner.update_idletasks()
            button = self.sheet_buttons[index]
            content_width = max(1, self.sheet_strip_inner.winfo_reqwidth())
            viewport = max(1, self.sheet_strip_canvas.winfo_width())
            left = button.winfo_x()
            right = left + button.winfo_width()
            current_left = self.sheet_strip_canvas.canvasx(0)
            if left < current_left:
                self.sheet_strip_canvas.xview_moveto(max(0.0, left / content_width))
            elif right > current_left + viewport:
                target = max(0, right - viewport)
                self.sheet_strip_canvas.xview_moveto(min(1.0, target / content_width))
        except Exception:
            pass

    def _rebuild_quick_links(self) -> None:
        if not hasattr(self, "quick_links_frame"):
            return
        for child in self.quick_links_frame.winfo_children():
            child.destroy()
        self._quick_link_buttons.clear()
        for link in self.quick_links[:MAX_QUICK_LINKS]:
            if not isinstance(link, LinkDefinition):
                continue
            button = ttk.Button(
                self.quick_links_frame,
                text=link.label,
                command=lambda item=link: self._activate_link(item),
            )
            button.pack(side="left", padx=2)
            self._quick_link_buttons.append(button)

    def _reload_quick_links(self) -> None:
        self.quick_links = load_links()
        self._rebuild_quick_links()
        self.status_text.set(f"Reloaded {len(self.quick_links)} link(s) from {LINKS_CONFIG_PATH}")

    def _expanded_link_text(self, value: str) -> str:
        return expand_link_text(
            value,
            current=self.current_dir,
            selected=self.selected_file_path(),
            workbook=self.path,
            project=Path(__file__).resolve().parents[1],
        )

    def _focus_link_pane(self, pane: str, *, folder: bool = False) -> None:
        if pane == "left":
            self._activate_left_panel()
        elif pane == "right":
            self._activate_right_panel()
        elif folder:
            # A folder link always changes the only file browser in xExcel, so
            # make that result immediately keyboard-ready unless the user
            # explicitly requested the workbook pane.
            self._activate_left_panel()

    def _open_link_folder(self, link: LinkDefinition) -> None:
        target = Path(self._expanded_link_text(link.target))
        if target.is_file():
            target = target.parent
        if not target.is_dir():
            raise NotADirectoryError(target)
        self.navigate(target)
        self._focus_link_pane(link.pane, folder=True)
        self.status_text.set(f"Shortcut: {link.label} -> {target}")

    def _launch_link_process(self, link: LinkDefinition, *, command: bool = False) -> None:
        target = self._expanded_link_text(link.target)
        if command:
            if os.name == "nt":
                args = ["powershell.exe", "-NoExit", "-Command", target]
                flags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
                subprocess.Popen(args, cwd=str(self.current_dir), creationflags=flags)
            else:
                subprocess.Popen(["/bin/sh", "-lc", target], cwd=str(self.current_dir), start_new_session=True)
        else:
            arguments = [target, *(self._expanded_link_text(arg) for arg in link.args)]
            flags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0) if os.name == "nt" else 0
            kwargs = {"cwd": str(self.current_dir)}
            if os.name == "nt" and flags:
                kwargs["creationflags"] = flags
            elif os.name != "nt":
                kwargs["start_new_session"] = True
            subprocess.Popen(arguments, **kwargs)
        self.status_text.set(f"Launched shortcut: {link.label}")

    def _run_link_action(self, link: LinkDefinition) -> None:
        action = link.target.strip().lower()
        if action == "powershell_here":
            command = "powershell.exe" if os.name == "nt" else "/bin/sh"
            flags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0) if os.name == "nt" else 0
            kwargs = {"cwd": str(self.current_dir)}
            if os.name == "nt" and flags:
                kwargs["creationflags"] = flags
            elif os.name != "nt":
                kwargs["start_new_session"] = True
            subprocess.Popen([command], **kwargs)
        elif action == "refresh_all":
            self.refresh_files()
            view = self.current_view()
            if view is not None:
                view.schedule_render()
        elif action == "search":
            self._focus_file_search()
        elif action in {"home", "go_home"}:
            self.go_home()
        elif action in {"open_in_excel", "open_default"}:
            self.open_default()
        elif action in {"save", "save_workbook"}:
            self.save_workbook()
        elif action == "copy":
            self.copy_current()
        elif action == "paste":
            self.paste_current()
        elif action == "copy_image":
            self.copy_image_current()
        elif action in {"edit_links", "links"}:
            self._edit_quick_links()
        elif action == "toggle_preview":
            self.status_text.set("xExcel uses the full workbook pane instead of mDIR's preview toggle.")
        elif action == "toggle_ai_terminal":
            self.status_text.set("AI/File panel is not part of the Excel-focused xExcel layout.")
        elif action == "hidden_system":
            self.status_text.set("Hidden/System toggle is not exposed in xViewer.")
        else:
            raise ValueError(f"unsupported action: {link.target}")
        self._focus_link_pane(link.pane)

    def _activate_link(self, link: LinkDefinition) -> None:
        try:
            if link.kind == "folder":
                self._open_link_folder(link)
            elif link.kind == "file":
                target = Path(self._expanded_link_text(link.target))
                if not target.exists():
                    raise FileNotFoundError(target)
                if os.name == "nt":
                    os.startfile(str(target))  # type: ignore[attr-defined]
                else:
                    subprocess.Popen(["xdg-open", str(target)])
                self.status_text.set(f"Opened shortcut: {link.label}")
                self._focus_link_pane(link.pane)
            elif link.kind == "program":
                self._launch_link_process(link)
                self._focus_link_pane(link.pane)
            elif link.kind == "command":
                self._launch_link_process(link, command=True)
                self._focus_link_pane(link.pane)
            elif link.kind == "web":
                target = self._expanded_link_text(link.target)
                if not webbrowser.open(target, new=2):
                    raise OSError("the default browser did not accept the URL")
                self.status_text.set(f"Opened website: {link.label}")
                self._focus_link_pane(link.pane)
            elif link.kind == "action":
                self._run_link_action(link)
        except Exception as exc:
            messagebox.showwarning(APP_TITLE, f"Shortcut failed: {link.label}\n\n{exc}", parent=self)
            self.status_text.set(f"Shortcut failed: {link.label} ({exc})")

    def _import_mdir_links(self) -> None:
        imported = load_mdir_links()
        if not imported:
            messagebox.showinfo(APP_TITLE, "No mDIR links were found.", parent=self)
            return
        try:
            save_links(imported)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not save imported links:\n{exc}", parent=self)
            return
        self.quick_links = imported
        self._rebuild_quick_links()
        self.status_text.set(f"Imported {len(imported)} mDIR link(s).")

    def _edit_quick_links(self) -> None:
        def links_saved(links: list[LinkDefinition]) -> None:
            try:
                save_links(links)
            except Exception as exc:
                messagebox.showerror(APP_TITLE, f"Could not save links:\n{exc}", parent=self)
                return
            self.quick_links = list(links)
            self._rebuild_quick_links()
            self.status_text.set(f"Saved {len(links)} link(s): {LINKS_CONFIG_PATH}")

        LinkManager(self, self.quick_links, self.current_dir, links_saved)

    def _path_bar_targets(self) -> list[tuple[str, Path]]:
        """Return mDIR-style path segments from root to the current folder."""
        path = self.current_dir
        results: list[tuple[str, Path]] = []
        try:
            anchor = path.anchor
            if anchor:
                current = Path(anchor)
                label = anchor.rstrip("\\/") or anchor
                results.append((label, current))
                relative_parts = path.parts[1:]
            else:
                current = Path(path.parts[0]) if path.parts else path
                relative_parts = path.parts[1:]
                if path.parts:
                    results.append((path.parts[0], current))
            for part in relative_parts:
                current = current / part
                results.append((part, current))
        except Exception:
            results = [(path.name or str(path), path)]
        return results

    def _path_bar_open(self, target: Path) -> str:
        self.navigate(target)
        self._activate_left_panel()
        return "break"

    def _path_bar_cursor(self, hand: bool) -> None:
        if hasattr(self, "path_bar"):
            self.path_bar.configure(cursor="hand2" if hand else "arrow")

    def _update_path_bar_scrollregion(self) -> None:
        if not hasattr(self, "path_bar") or not hasattr(self, "path_bar_inner"):
            return
        try:
            self.path_bar.update_idletasks()
            width = max(1, self.path_bar_inner.winfo_reqwidth() + 14)
            height = max(26, self.path_bar_inner.winfo_reqheight() + 6)
            self.path_bar.configure(scrollregion=(0, 0, width, height))
            if width > max(1, self.path_bar.winfo_width()):
                self.path_bar.xview_moveto(1.0)
            else:
                self.path_bar.xview_moveto(0.0)
        except Exception:
            pass

    def _rebuild_path_bar(self) -> None:
        if not hasattr(self, "path_bar_inner"):
            return
        for child in self.path_bar_inner.winfo_children():
            child.destroy()
        palette = dict(palette_for(self.theme_mode.get()))
        green = palette["path_bg"]
        targets = self._path_bar_targets()
        for index, (label, target) in enumerate(targets):
            segment = tk.Label(
                self.path_bar_inner,
                text=label,
                bg=green,
                fg=palette["path_foreground"],
                activebackground=green,
                activeforeground=palette["path_hover"],
                font=("Consolas", 11, "bold"),
                bd=0,
                padx=0,
                pady=0,
                cursor="hand2",
            )
            segment.pack(side="left")
            segment.bind("<Button-1>", lambda _e, p=target: self._path_bar_open(p))
            segment.bind("<Enter>", lambda e, c=palette["path_hover"]: e.widget.configure(fg=c))
            segment.bind("<Leave>", lambda e, c=palette["path_foreground"]: e.widget.configure(fg=c))
            slash = tk.Label(
                self.path_bar_inner,
                text="\\",
                bg=green,
                fg=palette["path_foreground"],
                font=("Consolas", 11, "bold"),
                bd=0,
                padx=0,
                pady=0,
            )
            slash.pack(side="left")
        self.after_idle(self._update_path_bar_scrollregion)

    def _edit_path_dialog(self, _event=None) -> str:
        """Open a compact path-entry dialog without cluttering the mDIR-style bar."""
        window = tk.Toplevel(self)
        window.title("Go to folder")
        window.transient(self)
        window.resizable(True, False)
        self.after_idle(lambda: self._apply_windows_titlebar(self._effective_theme == "Dark", window))
        value = tk.StringVar(value=str(self.current_dir))
        body = ttk.Frame(window, padding=10)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Folder:").pack(side="left", padx=(0, 6))
        entry = ttk.Entry(body, textvariable=value, width=72)
        entry.pack(side="left", fill="x", expand=True)

        def go(_evt=None):
            target = Path(value.get()).expanduser()
            if target.is_dir():
                window.destroy()
                self.navigate(target)
                self._activate_left_panel()
            else:
                messagebox.showwarning(APP_TITLE, f"Folder not found:\n{target}", parent=window)
            return "break"

        ttk.Button(body, text="Go", command=go).pack(side="left", padx=(6, 0))
        entry.bind("<Return>", go)
        entry.bind("<Escape>", lambda _e: window.destroy())
        entry.focus_set()
        entry.selection_range(0, "end")
        window.grab_set()
        return "break"

    def _remember_directory(self, path: Path, *, save: bool = True) -> None:
        self.recent_dirs = _merge_recent_directories(path, self.recent_dirs)
        if save:
            self._save_config()

    def _open_recent_directory(self, path: Path) -> None:
        if not Path(path).is_dir():
            # Remove dead history entries silently so the list self-heals.
            key = _recent_directory_key(path)
            self.recent_dirs = [p for p in self.recent_dirs if _recent_directory_key(p) != key]
            self._save_config()
            messagebox.showwarning(APP_TITLE, f"Folder not found:\n{path}", parent=self)
            return
        self.navigate(Path(path))
        self._activate_left_panel()

    def _clear_recent_directories(self) -> None:
        # Keep the directory that is currently open; clear only the history.
        self.recent_dirs = [self.current_dir]
        self._save_config()
        self.status_text.set("Folder history cleared. The current folder is kept.")

    def _show_recent_directories(self, _event=None) -> str:
        menu = tk.Menu(self, tearoff=False)
        valid: list[Path] = []
        for path in self.recent_dirs:
            try:
                if Path(path).is_dir():
                    valid.append(Path(path))
            except Exception:
                continue
        # Current folder first, then most-recently used folders.
        for index, path in enumerate(valid[:MAX_RECENT_DIRS]):
            label = str(path)
            if index == 0 and _recent_directory_key(path) == _recent_directory_key(self.current_dir):
                label = f"✓  {label}"
            menu.add_command(label=label, command=lambda p=path: self._open_recent_directory(p))
        if valid:
            menu.add_separator()
        menu.add_command(label="Clear Folder History", command=self._clear_recent_directories)
        try:
            button = self.path_history_button
            x = button.winfo_rootx()
            y = button.winfo_rooty() + button.winfo_height()
            menu.tk_popup(x, y)
        finally:
            try:
                menu.grab_release()
            except Exception:
                pass
        return "break"

    def _add_drive_buttons(self, parent) -> None:
        if os.name != "nt":
            return
        import ctypes
        mask = ctypes.windll.kernel32.GetLogicalDrives()
        for i in range(26):
            if mask & (1 << i):
                root = Path(f"{chr(65+i)}:/")
                ttk.Button(parent, text=f"{chr(65+i)}:", width=3, command=lambda p=root: self.navigate(p)).pack(side="left", padx=1)

    def _push_history(self, path: Path) -> None:
        path = path.resolve()
        if self.history_index >= 0 and self.history[self.history_index] == path:
            return
        self.history = self.history[: self.history_index + 1]
        self.history.append(path)
        self.history_index = len(self.history) - 1

    def navigate(self, path: Path, *, add_history: bool = True) -> None:
        self._cancel_file_right_drag(finalize=False)
        try:
            path = Path(path).expanduser().resolve()
        except Exception:
            path = Path(path).expanduser()
        if not path.is_dir():
            messagebox.showwarning(APP_TITLE, f"Folder not found:\n{path}", parent=self)
            return
        self.current_dir = path
        self.path_var.set(str(path))
        self._rebuild_path_bar()
        if add_history:
            self._push_history(path)
        self._remember_directory(path, save=True)
        self.refresh_files(rescan=True)

    def go_back(self) -> None:
        if self.history_index <= 0:
            return
        self.history_index -= 1
        self.navigate(self.history[self.history_index], add_history=False)

    def go_up(self) -> None:
        parent = self.current_dir.parent
        if parent != self.current_dir:
            self.navigate(parent)

    def go_home(self) -> None:
        self.navigate(Path.home())

    def sort_files(self, key: str) -> None:
        if self._sort_key == key:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_key = key
            self._sort_reverse = False
        self.refresh_files(rescan=False)

    def _type_filter_changed(self, selected: str) -> None:
        """Keep Excel/PDF/Image ``only`` controls mutually exclusive.

        Turning the currently active control off leaves all three off, which
        intentionally means "show all file types". Folders are always visible.
        """
        selected = str(selected).lower()
        mapping = {
            "excel": self.excel_only,
            "pdf": self.pdf_only,
            "image": self.images_only,
        }
        active = mapping.get(selected)
        if active is not None and active.get():
            for name, var in mapping.items():
                if name != selected:
                    var.set(False)
        self.refresh_files(rescan=False)

    def _schedule_filter_refresh(self, *_args) -> None:
        if self._filter_job is not None:
            try:
                self.after_cancel(self._filter_job)
            except Exception:
                pass
        # Debounce typing so very large directories remain responsive.
        self._filter_job = self.after(100, lambda: self.refresh_files(rescan=False))

    def _focus_file_search(self, _event=None) -> str:
        self._activate_left_panel()
        try:
            self.search_entry.focus_set()
            self.search_entry.selection_range(0, "end")
        except Exception:
            pass
        self.status_text.set("Search filenames: type terms; Up/Down moves straight to the filtered file list; Esc clears it.")
        return "break"

    def _clear_file_search(self, _event=None) -> str:
        if self.filter_var.get():
            self.filter_var.set("")
        else:
            self.refresh_files(rescan=False)
        try:
            self.search_entry.focus_set()
        except Exception:
            pass
        return "break"

    def _search_enter(self, _event=None) -> str:
        # Flush a pending debounced filter so Enter always acts on the text
        # currently visible in the search box.
        self._flush_pending_file_filter()
        children = self.file_tree.get_children()
        if children:
            first = children[0]
            self.file_tree.selection_set(first)
            self.file_tree.focus(first)
            self.file_tree.see(first)
            self.file_tree.focus_set()
            self._active_panel = "left"
            self._file_selected()
        return "break"

    def _flush_pending_file_filter(self) -> None:
        """Apply the latest search text immediately instead of waiting for debounce."""
        if self._filter_job is not None:
            try:
                self.after_cancel(self._filter_job)
            except Exception:
                pass
            self._filter_job = None
            self.refresh_files(rescan=False)

    def _search_arrow_to_files(self, event) -> str:
        """Move focus from Search to FILES and continue with Up/Down browsing."""
        self._flush_pending_file_filter()
        children = list(self.file_tree.get_children())
        if not children:
            self.status_text.set(
                f'No filename matches for "{self.filter_var.get().strip()}" | Esc clears the search'
            )
            return "break"

        # The arrow press is a deliberate hand-off from Search to FILES.
        # Start at the nearest edge of the filtered result set rather than
        # reusing a stale selection from before the search text changed.
        direction = -1 if getattr(event, "keysym", "") == "Up" else 1
        target_index = _search_arrow_target_index(len(children), None, direction)
        if target_index is None:
            return "break"

        iid = children[target_index]
        self.file_tree.selection_set(iid)
        self.file_tree.focus(iid)
        self.file_tree.see(iid)
        self.file_tree.focus_set()
        self._active_panel = "left"
        self._file_selected()
        return "break"

    def _scan_directory_records(self) -> bool:
        records: list[tuple[Path, bool, int, float]] = []
        try:
            with os.scandir(self.current_dir) as entries:
                for entry in entries:
                    try:
                        path = Path(entry.path)
                        is_dir = entry.is_dir(follow_symlinks=False)
                        stat = entry.stat(follow_symlinks=False)
                        records.append((path, is_dir, stat.st_size if not is_dir else 0, stat.st_mtime))
                    except OSError:
                        continue
        except OSError as exc:
            self.status_text.set(f"Cannot read folder: {exc}")
            return False
        self._directory_records = records
        self._records_dir = self.current_dir
        return True

    def refresh_files(self, *, rescan: bool = True) -> None:
        self._cancel_file_right_drag(finalize=False)
        selected_paths = set(self.selected_file_paths())
        focused_path = self.selected_file_path()
        for iid in self.file_tree.get_children():
            self.file_tree.delete(iid)
        self._item_paths.clear()
        self._filter_job = None

        if rescan or self._records_dir != self.current_dir:
            if not self._scan_directory_records():
                return

        filter_text = self.filter_var.get().strip()
        base_records = []
        for rec in self._directory_records:
            path, is_dir, _size, _mtime = rec
            if not _type_filter_accepts(
                path,
                is_dir,
                excel_only=self.excel_only.get(),
                pdf_only=self.pdf_only.get(),
                images_only=self.images_only.get(),
            ):
                continue
            base_records.append(rec)

        records = [
            rec for rec in base_records
            if _file_name_matches_filter(rec[0].name, filter_text)
        ]

        def sort_value(rec):
            path, _is_dir, size, mtime = rec
            if self._sort_key == "ext":
                return path.suffix.lower()
            if self._sort_key == "size":
                return size
            if self._sort_key == "modified":
                return mtime
            return path.name.lower()

        directories = [rec for rec in records if rec[1]]
        files = [rec for rec in records if not rec[1]]
        directories.sort(key=sort_value, reverse=self._sort_reverse)
        files.sort(key=sort_value, reverse=self._sort_reverse)
        records = directories + files
        restore_iids: list[str] = []
        focus_iid: Optional[str] = None
        for index, (path, is_dir, size, mtime) in enumerate(records):
            iid = f"i{index}"
            self._item_paths[iid] = path
            ext = "<DIR>" if is_dir else path.suffix.lower().lstrip(".")
            values = (
                path.name,
                ext,
                "" if is_dir else _human_size(size),
                datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M"),
            )
            self.file_tree.insert("", "end", iid=iid, values=values)
            if path in selected_paths:
                restore_iids.append(iid)
            if focused_path is not None and path == focused_path:
                focus_iid = iid
        if restore_iids:
            self.file_tree.selection_set(*restore_iids)
            target = focus_iid or restore_iids[0]
            self.file_tree.focus(target)
            self.file_tree.see(target)
        if filter_text:
            self.status_text.set(
                f'{self.current_dir} | {len(records):,} match(es) / {len(base_records):,} item(s) | Search: "{filter_text}"'
            )
        else:
            self.status_text.set(
                f"{self.current_dir} | {len(records):,} item(s) | Ctrl+F searches filenames | Search Up/Down moves to files | Right-drag selects multiple | Select Excel/PDF/Image to view on the right"
            )

    def selected_file_paths(self) -> list[Path]:
        """Return all selected LEFT-pane items in visible list order."""
        selected = set(self.file_tree.selection())
        if not selected:
            return []
        return [
            self._item_paths[iid]
            for iid in self.file_tree.get_children()
            if iid in selected and iid in self._item_paths
        ]

    def selected_file_path(self) -> Optional[Path]:
        """Return the focused selected item, falling back to the first selection."""
        selection = tuple(self.file_tree.selection())
        if not selection:
            return None
        focus = self.file_tree.focus()
        if focus in selection and focus in self._item_paths:
            return self._item_paths[focus]
        return self._item_paths.get(selection[0])

    def _file_selected(self, _event=None) -> None:
        paths = self.selected_file_paths()
        if self._right_drag_active:
            self.status_text.set(
                f"RIGHT-DRAG selection: {len(paths)} item(s) selected | release right mouse button to finish"
            )
            return
        if not paths:
            return
        if len(paths) > 1:
            if self._load_job:
                try:
                    self.after_cancel(self._load_job)
                except Exception:
                    pass
                self._load_job = None
            self._last_selected_path = None
            self.status_text.set(
                f"{len(paths)} items selected in {self.current_dir} | Ctrl/Shift-click or RIGHT-DRAG selects multiple | Del moves selection to Recycle Bin"
            )
            return

        path = paths[0]
        same_selection = path == self._last_selected_path
        self._last_selected_path = path
        action = _selection_workbook_action(paths)

        if action in {"load", "load_pdf", "load_image"}:
            if self._load_job:
                try:
                    self.after_cancel(self._load_job)
                except Exception:
                    pass
            # Arrow-key browsing can be fast; wait briefly so only the final file opens.
            delay = 180 if same_selection else 240
            if action == "load":
                self._load_job = self.after(delay, lambda p=path: self.open_workbook(p))
            elif action == "load_pdf":
                self._load_job = self.after(delay, lambda p=path: self.open_media(p, "pdf"))
            else:
                self._load_job = self.after(delay, lambda p=path: self.open_media(p, "image"))
            return

        # A directory or unsupported file clears the old RIGHT content so a
        # workbook, PDF page, or image never appears to belong to the newly
        # selected item. Unsaved Excel edits keep the same Save/Discard/Cancel
        # protection used when moving between workbooks.
        if self.model is not None and not self._confirm_discard_or_save():
            self._reselect_current_workbook()
            self._last_selected_path = self.path
            return
        self._clear_current_workbook_view()
        if path.is_dir():
            self.status_text.set(
                f"{path.name} | Folder selected. Press Enter or double-click to open it. RIGHT viewer cleared."
            )
        elif path.is_file():
            self.status_text.set(
                f"{path.name} | Unsupported viewer type. Press Enter or double-click to open with the default application. RIGHT viewer cleared."
            )
        else:
            self.status_text.set(f"{path.name} | Item is no longer available. RIGHT viewer cleared.")

    def _show_excel_ui(self) -> None:
        """Show the Excel-specific controls in the RIGHT pane."""
        try:
            self.media_nav.pack_forget()
            self.media_view.pack_forget()
        except Exception:
            pass
        if not self.formula_frame.winfo_ismapped():
            self.formula_frame.pack(fill="x")
        if not self.sheet_nav_frame.winfo_ismapped():
            self.sheet_nav_frame.pack(fill="x")
        if not self.notebook.winfo_ismapped():
            self.notebook.pack(fill="both", expand=True)
        self.copy_btn.configure(state="normal")
        self.copy_image_btn.configure(state="normal")
        self.paste_btn.configure(state="normal" if self.model is not None and self.model.editable else "disabled")
        self.open_btn.configure(text="Open in Excel", state="normal" if self.path else "disabled")
        self.fit_1to1_btn.configure(state="disabled")

    def _show_media_ui(self, kind: str) -> None:
        """Show the PDF/image viewer while hiding Excel-only controls."""
        for widget in (self.formula_frame, self.sheet_nav_frame, self.notebook):
            try:
                widget.pack_forget()
            except Exception:
                pass
        if not self.media_nav.winfo_ismapped():
            self.media_nav.pack(fill="x")
        if not self.media_view.winfo_ismapped():
            self.media_view.pack(fill="both", expand=True)
        self.save_btn.configure(state="disabled")
        self.save_as_btn.configure(state="disabled")
        self.paste_btn.configure(state="disabled")
        self.copy_btn.configure(state="normal")
        self.copy_image_btn.configure(state="normal")
        self.open_btn.configure(text="Open PDF" if kind == "pdf" else "Open Image", state="normal")
        self.fit_1to1_btn.configure(state="normal" if kind == "image" else "disabled")

    def _media_page_changed(self, page_index: int, page_count: int, kind: str) -> None:
        if kind == "pdf" and page_count > 0:
            self.media_type_text.set("PDF")
            self.media_page_text.set(f"Page {page_index + 1} / {page_count}")
            self.media_first_btn.configure(state="normal" if page_index > 0 else "disabled")
            self.media_prev_btn.configure(state="normal" if page_index > 0 else "disabled")
            self.media_next_btn.configure(state="normal" if page_index + 1 < page_count else "disabled")
            self.media_last_btn.configure(state="normal" if page_index + 1 < page_count else "disabled")
        elif kind == "image":
            self.media_type_text.set("IMAGE")
            self.media_page_text.set("Image")
            for button in (self.media_first_btn, self.media_prev_btn, self.media_next_btn, self.media_last_btn):
                button.configure(state="disabled")
        else:
            self.media_type_text.set("")
            self.media_page_text.set("")
            for button in (self.media_first_btn, self.media_prev_btn, self.media_next_btn, self.media_last_btn):
                button.configure(state="disabled")

    def _media_step_page(self, delta: int) -> None:
        self.media_view.step_page(delta)

    def _media_first_page(self) -> None:
        self.media_view.first_page()

    def _media_last_page(self) -> None:
        self.media_view.last_page()

    def _media_save_image_as(self) -> None:
        self.media_view.save_current_image_as()

    def _media_zoom_changed(self, zoom: float, mode: str) -> None:
        """Keep the toolbar percentage synchronized with Image Fit/1:1."""
        if self.viewer_kind != "image":
            return
        pct = max(1, int(round(float(zoom) * 100)))
        self.zoom_value.set(f"{pct}%")
        if mode == "fit":
            self.status_text.set(
                f"{self.path} | Image | Fit {pct}% (max 1:1)" if self.path else f"Image | Fit {pct}%"
            )
        elif mode == "1:1":
            self.status_text.set(f"{self.path} | Image | 1:1 (100%)" if self.path else "Image | 1:1 (100%)")
        elif mode == "manual":
            self.status_text.set(f"{self.path} | Image | {pct}%" if self.path else f"Image | {pct}%")

    def _toggle_media_fit_one_to_one(self) -> None:
        if self.viewer_kind != "image":
            return
        self.media_view.toggle_fit_one_to_one()

    def open_media(self, path: Path, kind: str) -> None:
        """Open a PDF or image inside the RIGHT viewer."""
        path = Path(path)
        kind = str(kind).lower()
        if kind not in {"pdf", "image"} or not path.exists():
            return
        if self.path is not None and self.viewer_kind == kind:
            try:
                if path.resolve() == self.path.resolve():
                    return
            except Exception:
                pass
        if self.model is not None and not self._confirm_discard_or_save():
            self._reselect_current_workbook()
            return
        self._clear_current_workbook_view()
        self.viewer_kind = kind
        self.path = path
        self.zoom_value.set("100%")
        self._show_media_ui(kind)
        self.workbook_title.set(f"{path.name}   [{'PDF' if kind == 'pdf' else 'IMAGE'}]")
        self.title(f"xViewer {APP_VERSION} — {path.name}")
        self.media_view.zoom = 1.0
        self.media_view.load(path, kind)
        if self._active_panel == "left":
            self.file_tree.focus_set()

    def _clear_current_workbook_view(self) -> None:
        """Close the current workbook and restore the blank RIGHT placeholder.

        Incrementing ``_load_generation`` is important: a workbook may already
        be loading in a background thread when the LEFT selection moves to a
        directory or non-Excel file.  Late results must be discarded instead
        of repainting stale images after the pane has been cleared.
        """
        self._load_generation += 1
        self._loading_path = None
        if self._load_job:
            try:
                self.after_cancel(self._load_job)
            except Exception:
                pass
            self._load_job = None
        if self.model is not None:
            try:
                self.model.close()
            except Exception:
                pass
        self.model = None
        self.path = None
        self.viewer_kind = "blank"
        try:
            self.media_view.clear()
        except Exception:
            pass

        old_views = list(self.sheet_views)
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        self.sheet_views.clear()
        for view in old_views:
            try:
                view.destroy()
            except Exception:
                pass

        # ``blank`` is created once in _build_ui.  It may have been forgotten
        # when an Excel workbook was shown, so add it back every time we clear.
        try:
            self.notebook.add(self.blank, text="Workbook")
            self.notebook.select(self.blank)
        except Exception:
            pass

        self._rebuild_sheet_strip()
        self._show_excel_ui()
        self.workbook_title.set("No file selected")
        self.cell_ref.set("A1")
        self.formula_value.set("")
        self.save_btn.configure(state="disabled")
        self.save_as_btn.configure(state="disabled")
        self.copy_btn.configure(state="disabled")
        self.paste_btn.configure(state="disabled")
        self.copy_image_btn.configure(state="disabled")
        self.open_btn.configure(text="Open", state="disabled")
        self.title(f"xViewer {APP_VERSION}")

    @staticmethod
    def _delete_confirmation_text(paths: list[Path]) -> str:
        count = len(paths)
        preview = []
        for path in paths[:10]:
            marker = "[Folder] " if path.is_dir() else ""
            preview.append(f"  {marker}{path.name}")
        if count > 10:
            preview.append(f"  ... and {count - 10} more")
        noun = "item" if count == 1 else "items"
        return (
            f"Move {count} selected {noun} to the Recycle Bin?\n\n"
            + "\n".join(preview)
            + "\n\nYou can usually restore them later from the Recycle Bin."
        )

    def _delete_selected_items(self, _event=None) -> str:
        """Move all selected LEFT-pane items to the Recycle Bin after confirmation."""
        paths = self.selected_file_paths()
        if not paths:
            return "break"

        current_selected = False
        if self.path is not None:
            try:
                current_selected = any(p.resolve() == self.path.resolve() for p in paths)
            except Exception:
                current_selected = any(str(p) == str(self.path) for p in paths)

        if current_selected and self.model is not None and self.model.dirty:
            answer = messagebox.askyesnocancel(
                APP_TITLE,
                f"{self.path.name if self.path else 'The open workbook'} has unsaved changes.\n\n"
                "Save the changes before moving it to the Recycle Bin?",
                parent=self,
            )
            if answer is None:
                return "break"
            if answer and not self.save_workbook():
                return "break"

        if not messagebox.askyesno(
            APP_TITLE, self._delete_confirmation_text(paths), parent=self, icon="warning"
        ):
            return "break"

        children_before = list(self.file_tree.get_children())
        selected_iids = set(self.file_tree.selection())
        selected_indices = [i for i, iid in enumerate(children_before) if iid in selected_iids]
        next_index = min(selected_indices) if selected_indices else 0

        moved, failed = _send_paths_to_recycle_bin(paths)
        moved_keys = {_recent_directory_key(p) for p in moved}

        if self.path is not None and _recent_directory_key(self.path) in moved_keys:
            self._clear_current_workbook_view()

        self._last_selected_path = None
        self.refresh_files(rescan=True)

        children_after = list(self.file_tree.get_children())
        if children_after:
            iid = children_after[min(next_index, len(children_after) - 1)]
            self.file_tree.selection_set(iid)
            self.file_tree.focus(iid)
            self.file_tree.see(iid)
            self.file_tree.focus_set()

        if failed:
            details = "\n".join(f"{p.name}: {exc}" for p, exc in failed[:8])
            if len(failed) > 8:
                details += f"\n... and {len(failed) - 8} more"
            messagebox.showerror(
                APP_TITLE,
                f"Moved {len(moved)} item(s) to the Recycle Bin, but {len(failed)} item(s) could not be deleted.\n\n{details}",
                parent=self,
            )
        else:
            self.status_text.set(
                f"Moved {len(moved)} item(s) to the Recycle Bin | {self.current_dir}"
            )
        return "break"

    def _open_path_with_default_application(self, path: Path) -> None:
        """Open *path* using the operating system's associated application.

        On Windows this is the same association used by Explorer (for example,
        .xlsx normally opens in Microsoft Excel).  The helper is intentionally
        separate from the integrated RIGHT-pane workbook loader so Enter can
        open the original file without changing the browse workflow.
        """
        try:
            if os.name == "nt":
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
            self.status_text.set(f"Opened with default application: {path.name}")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not open file:\n{exc}", parent=self)

    def _file_enter_open_default(self, _event=None) -> str:
        """Enter in the LEFT pane opens a file externally; folders navigate."""
        path = self.selected_file_path()
        if path is None:
            return "break"
        if path.is_dir():
            self.navigate(path)
            return "break"
        self._open_path_with_default_application(path)
        return "break"

    def _file_activated(self, event=None) -> str:
        """Double-click in LEFT opens the pointed file with its default app.

        A single click/arrow selection already drives the integrated RIGHT viewer,
        so double-click mirrors Enter: folders navigate; files open externally.
        Resolve the row under the pointer first so the clicked item is used even on
        Tk builds where Treeview selection notification arrives slightly later.
        """
        if event is not None:
            try:
                iid = self.file_tree.identify_row(int(event.y))
            except Exception:
                iid = ""
            if iid and iid in self._item_paths:
                self.file_tree.selection_set(iid)
                self.file_tree.focus(iid)
                path = self._item_paths.get(iid)
            else:
                path = self.selected_file_path()
        else:
            path = self.selected_file_path()
        if path is None:
            return "break"
        if path.is_dir():
            self.navigate(path)
            return "break"
        self._open_path_with_default_application(path)
        return "break"

    def _reselect_current_workbook(self) -> None:
        if self.path is None:
            return
        for iid, candidate in self._item_paths.items():
            try:
                if candidate.resolve() == self.path.resolve():
                    self.file_tree.selection_set(iid)
                    self.file_tree.see(iid)
                    return
            except Exception:
                continue

    def _confirm_discard_or_save(self) -> bool:
        if self.model is None or not self.model.dirty:
            return True
        answer = messagebox.askyesnocancel(
            APP_TITLE,
            f"{self.path.name if self.path else 'Workbook'} has unsaved changes.\n\nSave before opening another workbook?",
            parent=self,
        )
        if answer is None:
            return False
        if answer:
            return self.save_workbook()
        return True

    def open_workbook(self, path: Path) -> None:
        self._load_job = None
        path = Path(path)
        if self.path is not None and path.resolve() == self.path.resolve() and self.model is not None:
            return
        if not self._confirm_discard_or_save():
            self._reselect_current_workbook()
            return
        if not path.exists():
            return
        try:
            self.media_view.clear()
        except Exception:
            pass
        self.viewer_kind = "excel"
        self._show_excel_ui()
        self._load_generation += 1
        generation = self._load_generation
        self._loading_path = path
        self.status_text.set(f"Loading full workbook: {path.name} ...")
        self.workbook_title.set(f"Loading {path.name} ...")

        def worker() -> None:
            # Legacy BIFF .xls files can contain drawing objects that neither
            # xlrd nor openpyxl can expose as images.  On Windows, prefer an
            # exact read-only PDF rendered by Microsoft Excel itself.  This
            # preserves old pictures/metafiles/OLE drawings visually.  If Excel
            # automation is unavailable, fall back to the normal cell/grid path.
            if path.suffix.lower() == ".xls":
                try:
                    exact_pdf = legacy_xls_pdf_preview_path(path)
                except Exception:
                    exact_pdf = None
                if exact_pdf is not None and exact_pdf.is_file():
                    self.after(0, lambda p=exact_pdf: self._load_legacy_xls_visual_complete(path, generation, p))
                    return
            try:
                model = EditableWorkbookModel(path)
            except Exception as exc:
                self.after(0, lambda: self._load_failed(path, generation, exc))
                return
            self.after(0, lambda: self._load_complete(path, generation, model))

        threading.Thread(target=worker, name="xExcel-Integrated-Workbook-Loader", daemon=True).start()

    def _load_failed(self, path: Path, generation: int, exc: Exception) -> None:
        if generation != self._load_generation:
            return
        self.status_text.set(f"Open failed: {path.name}")
        self.workbook_title.set("No workbook selected")
        messagebox.showerror(APP_TITLE, f"Could not open workbook:\n\n{path}\n\n{exc}", parent=self)

    def _load_legacy_xls_visual_complete(self, path: Path, generation: int, preview_pdf: Path) -> None:
        """Show an Excel-rendered exact visual preview for a legacy .xls file.

        Old BIFF workbooks frequently keep product photos as drawing-layer objects
        that disappear from Python-level workbook readers.  The PDF was rendered
        by Microsoft Excel from the original read-only workbook, so the RIGHT pane
        shows those pictures even when direct image extraction returns zero.
        """
        if generation != self._load_generation:
            return

        old = self.model
        self.model = None
        if old is not None:
            try:
                old.close()
            except Exception:
                pass

        old_views = list(self.sheet_views)
        for tab in self.notebook.tabs():
            try:
                self.notebook.forget(tab)
            except Exception:
                pass
        self.sheet_views.clear()
        for view in old_views:
            try:
                view.destroy()
            except Exception:
                pass

        self.path = path
        self.viewer_kind = "pdf"
        self._loading_path = None
        self.zoom_value.set("100%")
        self._show_media_ui("pdf")
        self.open_btn.configure(text="Open in Excel", state="normal")
        self.workbook_title.set(f"{path.name}   [READ ONLY (.xls) | EXACT VIEW]")
        self.title(f"xViewer {APP_VERSION} — {path.name}")
        self.media_view.zoom = 1.0
        self.media_view.load(preview_pdf, "pdf", display_label=f"{path} | Legacy .xls EXACT VIEW")
        self.status_text.set(
            f"{path} | Legacy .xls exact view rendered by Microsoft Excel | "
            "Images/drawings preserved visually | Open in Excel for cell editing/copying"
        )
        if self._active_panel == "left":
            self.file_tree.focus_set()

    def _load_complete(self, path: Path, generation: int, model: EditableWorkbookModel) -> None:
        if generation != self._load_generation:
            model.close()
            return
        old = self.model
        old_views = list(self.sheet_views)
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        self.sheet_views.clear()
        for view in old_views:
            try:
                view.destroy()
            except Exception:
                pass
        self.model = model
        self.path = path
        self.viewer_kind = "excel"
        self._show_excel_ui()
        for sheet in model.sheets:
            view = EditableVirtualSheet(self.notebook, sheet, self._set_status, self._sheet_selection_changed, ui_palette=palette_for(self.theme_mode.get()))
            self.sheet_views.append(view)
            self.notebook.add(view, text=sheet.title)
            self._apply_pane_nav_bindtag(view)
        if old is not None:
            old.close()
        editable_text = "EDITABLE" if model.editable else "READ ONLY (.xls)"
        image_count = sum(len(s.images) for s in model.sheets)
        self.workbook_title.set(f"{path.name}   [{editable_text}]")
        self.title(f"xViewer {APP_VERSION} — {path.name}")
        self.save_btn.configure(state="normal" if model.editable else "disabled")
        self.save_as_btn.configure(state="normal" if model.editable else "disabled")
        self.paste_btn.configure(state="normal" if model.editable else "disabled")
        self.copy_btn.configure(state="normal")
        self.copy_image_btn.configure(state="normal")
        self.open_btn.configure(text="Open in Excel", state="normal")
        self.status_text.set(
            f"{path} | {len(model.sheets)} sheet(s) | {image_count} embedded image(s) | "
            "Ctrl+Left FILES | Ctrl+Right WORKBOOK | Alt+1 FILES | Alt+2 WORKBOOK | Tab disabled in FILES / moves cells in WORKBOOK | arrows move files/cells | F2 edit | Ctrl+S save"
        )
        self._rebuild_sheet_strip()
        if self.sheet_views:
            self._sheet_selection_changed(self.sheet_views[0])
            # Loading a workbook must not steal focus while the user is browsing
            # the left file list with arrow keys.
            if self._active_panel == "right":
                self.sheet_views[0].body.focus_set()
            else:
                self.file_tree.focus_set()

    def _set_status(self, text: str) -> None:
        if self.model is not None and self.model.dirty:
            self.status_text.set(text + "   [UNSAVED]")
            if self.path is not None:
                self.workbook_title.set(f"{self.path.name} *   [EDITABLE]")
        else:
            self.status_text.set(text)

    def current_view(self) -> Optional[EditableVirtualSheet]:
        if not self.sheet_views:
            return None
        selected = self.notebook.select()
        if not selected:
            return self.sheet_views[0]
        widget = self.nametowidget(selected)
        return widget if isinstance(widget, EditableVirtualSheet) else None

    def _sheet_selection_changed(self, view: EditableVirtualSheet) -> None:
        row, col = view.active_cell
        self.cell_ref.set(f"{_column_name(col)}{row}")
        raw = view.raw_active_value()
        self.formula_value.set("" if raw is None else str(raw))
        if self.model is not None and self.model.dirty and self.path is not None:
            self.workbook_title.set(f"{self.path.name} *   [EDITABLE]")

    def _sync_formula_bar(self) -> str:
        view = self.current_view()
        if view is not None:
            self._sheet_selection_changed(view)
            view.body.focus_set()
        return "break"

    def _formula_commit(self, _event=None) -> str:
        view = self.current_view()
        if view is not None:
            if view.set_active_value(_parse_clipboard_value(self.formula_value.get())):
                self._sheet_selection_changed(view)
        return "break"

    def copy_current(self) -> None:
        if self.viewer_kind in {"pdf", "image"}:
            self.media_view.copy_current()
            return
        view = self.current_view()
        if view is not None:
            view.copy_selection()

    def paste_current(self) -> None:
        if self.viewer_kind != "excel":
            return
        view = self.current_view()
        if view is not None:
            view.paste_selection()

    def copy_image_current(self) -> None:
        if self.viewer_kind in {"pdf", "image"}:
            self.media_view.copy_current()
            return
        view = self.current_view()
        if view is not None:
            view.copy_selected_image()

    def _commit_pending_edits(self) -> None:
        view = self.current_view()
        if view is not None and view._edit_entry is not None:
            view.commit_edit()
        if self.focus_get() is self.formula_entry and view is not None:
            self._formula_commit()

    def save_workbook(self) -> bool:
        if self.model is None or not self.model.editable:
            return False
        self._commit_pending_edits()
        if not self.model.dirty:
            self.status_text.set("No changes to save.")
            return True
        try:
            target = self.model.save(create_backup=True)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not save workbook:\n{exc}", parent=self)
            return False
        self.path = target
        self.workbook_title.set(f"{target.name}   [EDITABLE]")
        self.title(f"xViewer {APP_VERSION} — {target.name}")
        self.status_text.set(f"Saved: {target} | A safety backup was created in xExcel_Backup before the first overwrite.")
        self.refresh_files(rescan=True)
        return True

    def save_as(self) -> bool:
        if self.model is None or not self.model.editable:
            return False
        self._commit_pending_edits()
        initial = self.path.name if self.path else "workbook.xlsx"
        target = filedialog.asksaveasfilename(
            parent=self,
            title="Save workbook as",
            initialfile=initial,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Macro-enabled Workbook", "*.xlsm"), ("All files", "*.*")],
        )
        if not target:
            return False
        try:
            saved = self.model.save(Path(target), create_backup=False)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not save workbook:\n{exc}", parent=self)
            return False
        self.path = saved
        self.workbook_title.set(f"{saved.name}   [EDITABLE]")
        self.title(f"xViewer {APP_VERSION} — {saved.name}")
        self.status_text.set(f"Saved copy: {saved}")
        return True

    def open_default(self) -> None:
        if self.path is None:
            return
        try:
            if os.name == "nt":
                os.startfile(str(self.path))  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", str(self.path)])
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not open default application:\n{exc}", parent=self)

    def _zoom_changed(self) -> None:
        try:
            self.set_zoom(float(self.zoom_value.get().rstrip("%")) / 100.0)
        except ValueError:
            self.zoom_value.set("100%")

    def set_zoom(self, zoom: float) -> None:
        zoom = max(0.5, min(2.0, zoom))
        self.zoom_value.set(f"{int(round(zoom * 100))}%")
        if self.viewer_kind in {"pdf", "image"}:
            self.media_view.set_zoom(zoom)
            return
        view = self.current_view()
        if view is not None:
            view.set_zoom(zoom)

    def adjust_zoom(self, delta: float) -> None:
        if self.viewer_kind in {"pdf", "image"}:
            current = self.media_view.zoom
        else:
            view = self.current_view()
            current = view.zoom if view is not None else 1.0
        self.set_zoom(current + delta)

    def _sheet_changed(self) -> None:
        view = self.current_view()
        if view is not None:
            self._sync_sheet_strip()
            if self._active_panel == "right":
                view.body.focus_set()
            self.zoom_value.set(f"{int(round(view.zoom * 100))}%")
            self._sheet_selection_changed(view)

    def _save_config(self) -> None:
        try:
            CONFIG_PATH.write_text(
                json.dumps(
                    {
                        "last_dir": str(self.current_dir),
                        "theme": normalize_theme_mode(self.theme_mode.get()),
                        "recent_dirs": [str(path) for path in self.recent_dirs[:MAX_RECENT_DIRS]],
                    },
                    indent=2,
                ),
                encoding="utf-8",
            )
        except Exception:
            pass

    def _on_close(self) -> None:
        if not self._confirm_discard_or_save():
            return
        self._save_config()
        if self._theme_poll_job is not None:
            try:
                self.after_cancel(self._theme_poll_job)
            except Exception:
                pass
            self._theme_poll_job = None
        if self.model is not None:
            self.model.close()
        self.destroy()


def self_check() -> int:
    try:
        import openpyxl  # noqa: F401
        import PIL  # noqa: F401
        from .excel_viewer import VirtualSheet  # noqa: F401
    except Exception as exc:
        print(f"xViewer self-check failed: {exc}", file=sys.stderr)
        return 1
    try:
        import xlrd  # noqa: F401
        legacy = "legacy .xls support OK"
    except Exception:
        legacy = "legacy .xls support unavailable until xlrd is installed"
    try:
        import pypdfium2  # noqa: F401
        pdf = "PDF viewer OK"
    except Exception:
        pdf = "PDF viewer unavailable until pypdfium2 is installed"
    print(f"xViewer {APP_VERSION} self-check OK ({legacy}; {pdf})")
    return 0


def launch_workspace(path: Optional[Path] = None) -> int:
    app = ExcelWorkspaceApp(path)
    app.mainloop()
    return 0


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if "--check" in args:
        return self_check()
    initial: Optional[Path] = None
    for value in args:
        if value.startswith("-"):
            continue
        p = Path(value).expanduser()
        if p.exists():
            initial = p
            break
    return launch_workspace(initial)


if __name__ == "__main__":
    raise SystemExit(main())

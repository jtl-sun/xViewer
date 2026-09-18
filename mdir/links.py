from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable, Iterable, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

MAX_LINKS = 16
LINK_KINDS = ("folder", "file", "program", "web", "action", "command")
PANE_VALUES = ("active", "left", "right")
LINKS_CONFIG_PATH = Path.home() / ".xexcel-viewer-links.json"
MDIR_LINKS_CONFIG_PATH = Path.home() / ".mdir-p-shortcuts.json"


@dataclass(frozen=True)
class LinkDefinition:
    label: str
    kind: str
    target: str
    args: tuple[str, ...] = ()
    pane: str = "active"


DEFAULT_LINKS = (
    LinkDefinition("Home", "folder", "{home}"),
    LinkDefinition("PowerShell", "action", "powershell_here"),
)


def _definition_from_value(value: object) -> LinkDefinition | None:
    if not isinstance(value, dict):
        return None
    label = str(value.get("label", "")).strip()
    kind = str(value.get("type", value.get("kind", "folder"))).strip().lower()
    target = str(value.get("target", "")).strip()
    pane = str(value.get("pane", "active")).strip().lower()
    raw_args = value.get("args", [])
    if not label or not target or kind not in LINK_KINDS:
        return None
    if pane not in PANE_VALUES:
        pane = "active"
    if not isinstance(raw_args, list):
        raw_args = []
    return LinkDefinition(
        label=label[:24],
        kind=kind,
        target=target,
        args=tuple(str(item) for item in raw_args),
        pane=pane,
    )


def parse_links(values: object) -> list[LinkDefinition]:
    if not isinstance(values, list):
        return []
    result: list[LinkDefinition] = []
    for value in values:
        link = _definition_from_value(value)
        if link is not None:
            result.append(link)
        if len(result) >= MAX_LINKS:
            break
    return result


def _serializable(links: Iterable[LinkDefinition]) -> list[dict[str, object]]:
    result: list[dict[str, object]] = []
    for link in tuple(links)[:MAX_LINKS]:
        value = asdict(link)
        value["type"] = value.pop("kind")
        value["args"] = list(link.args)
        result.append(value)
    return result


def save_links(links: Iterable[LinkDefinition], path: Path = LINKS_CONFIG_PATH) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(_serializable(links), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)
    return path


def load_mdir_links(path: Path = MDIR_LINKS_CONFIG_PATH) -> list[LinkDefinition]:
    try:
        return parse_links(json.loads(path.read_text(encoding="utf-8")))
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        return []


def _looks_legacy_folder_only(raw: object) -> bool:
    if not isinstance(raw, list) or not raw:
        return False
    for value in raw:
        if not isinstance(value, dict):
            return False
        # xExcel 2.3.x stored only label/target and sometimes kind/type=folder.
        kind = str(value.get("type", value.get("kind", "folder"))).lower()
        if kind != "folder":
            return False
        if set(value).issubset({"label", "target", "type", "kind"}) is False:
            return False
    return True


def _legacy_matches_mdir(legacy: list[LinkDefinition], mdir: list[LinkDefinition]) -> bool:
    mdir_folders = [(x.label, x.target) for x in mdir if x.kind == "folder"]
    legacy_pairs = [(x.label, x.target) for x in legacy]
    return bool(mdir_folders) and legacy_pairs == mdir_folders[: len(legacy_pairs)]


def load_links(
    path: Path = LINKS_CONFIG_PATH,
    mdir_path: Path = MDIR_LINKS_CONFIG_PATH,
) -> list[LinkDefinition]:
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
        links = parse_links(raw)
        if links:
            # Seamless 2.3.x -> 2.4 migration: when the old xExcel list is
            # simply the folder subset imported from mDIR, restore the whole
            # mDIR Link Manager list (Web/File/Action included).
            if _looks_legacy_folder_only(raw):
                mdir = load_mdir_links(mdir_path)
                if mdir and _legacy_matches_mdir(links, mdir):
                    return mdir
            return links
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        pass
    imported = load_mdir_links(mdir_path)
    return imported if imported else list(DEFAULT_LINKS)


def expand_link_text(
    value: str,
    *,
    current: Path,
    selected: Path | None,
    workbook: Path | None,
    project: Path,
) -> str:
    replacements = {
        "{current}": str(current),
        "{left}": str(current),
        "{right}": str(workbook.parent if workbook else current),
        "{selected}": str(selected or current),
        "{left_selected}": str(selected or current),
        "{right_selected}": str(workbook or current),
        "{home}": str(Path.home()),
        "{project}": str(project),
    }
    expanded = str(value)
    for token, replacement in replacements.items():
        expanded = expanded.replace(token, replacement)
    return os.path.expandvars(os.path.expanduser(expanded))


class LinkManager(tk.Toplevel):
    """Tk version of the mDIR Link Manager, with the same data model."""

    def __init__(
        self,
        parent: tk.Misc,
        links: Iterable[LinkDefinition],
        current_path: Path,
        on_save: Callable[[list[LinkDefinition]], None],
        *,
        mdir_path: Path = MDIR_LINKS_CONFIG_PATH,
    ) -> None:
        super().__init__(parent)
        self.title("xViewer — Link Manager")
        self.geometry("1120x650")
        self.minsize(900, 560)
        self.transient(parent)
        try:
            self.configure(background=ttk.Style(self).lookup("TFrame", "background"))
        except Exception:
            pass
        titlebar = getattr(parent, "_apply_windows_titlebar", None)
        if callable(titlebar):
            self.after_idle(lambda: titlebar(getattr(parent, "_effective_theme", "") == "Dark", self))
        self.grab_set()
        self.current_path = current_path
        self.on_save = on_save
        self.mdir_path = mdir_path
        self._syncing = False
        self._editing_index: Optional[int] = None
        self.drafts = [self._draft(link) for link in list(links)[:MAX_LINKS]]

        self.name_var = tk.StringVar()
        self.type_var = tk.StringVar(value="folder")
        self.pane_var = tk.StringVar(value="active")
        self.target_var = tk.StringVar()
        self.args_var = tk.StringVar(value="[]")
        self.status_var = tk.StringVar(value="Select a row, edit its fields, then Save. Esc cancels all changes.")

        self._build()
        self._wire_live_editor()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda _e: self.destroy())
        self._render_rows(0 if self.drafts else None)

    @staticmethod
    def _draft(link: LinkDefinition) -> dict[str, str]:
        return {
            "label": link.label,
            "kind": link.kind,
            "target": link.target,
            "pane": link.pane,
            "args": json.dumps(list(link.args), ensure_ascii=False),
        }

    def _build(self) -> None:
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(header, text="Link Manager", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Label(header, text="  mDIR-compatible shortcuts", style="Muted.TLabel").pack(side="left", pady=(6, 0))

        table_frame = ttk.Frame(outer)
        table_frame.grid(row=1, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        cols = ("name", "type", "target", "pane")
        self.table = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse", height=10)
        for key, title, width in (
            ("name", "Name", 150),
            ("type", "Type", 105),
            ("target", "URL / Path / Action", 590),
            ("pane", "Pane", 120),
        ):
            self.table.heading(key, text=title)
            self.table.column(key, width=width, minwidth=70, anchor="w")
        self.table.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.table.configure(yscrollcommand=scroll.set)
        self.table.bind("<<TreeviewSelect>>", self._row_selected)

        editor = ttk.Frame(outer, padding=(0, 12, 0, 0))
        editor.grid(row=2, column=0, sticky="ew")
        editor.columnconfigure(1, weight=1)
        editor.columnconfigure(2, weight=1)

        ttk.Label(editor, text="Name", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 12), pady=6)
        self.name_entry = ttk.Entry(editor, textvariable=self.name_var)
        self.name_entry.grid(row=0, column=1, columnspan=2, sticky="ew", pady=6)

        ttk.Label(editor, text="Type / Pane", font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky="w", padx=(0, 12), pady=6)
        self.type_combo = ttk.Combobox(editor, textvariable=self.type_var, values=[x.title() for x in LINK_KINDS], state="readonly", width=18)
        self.type_combo.grid(row=1, column=1, sticky="w", pady=6)
        self.pane_combo = ttk.Combobox(editor, textvariable=self.pane_var, values=["Active pane", "Left / Files", "Right / Workbook"], state="readonly", width=20)
        self.pane_combo.grid(row=1, column=2, sticky="w", padx=(12, 0), pady=6)

        ttk.Label(editor, text="Target", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", padx=(0, 12), pady=6)
        self.target_entry = ttk.Entry(editor, textvariable=self.target_var)
        self.target_entry.grid(row=2, column=1, columnspan=2, sticky="ew", pady=6)

        ttk.Label(editor, text="Arguments", font=("Segoe UI", 10, "bold")).grid(row=3, column=0, sticky="w", padx=(0, 12), pady=6)
        self.args_entry = ttk.Entry(editor, textvariable=self.args_var)
        self.args_entry.grid(row=3, column=1, columnspan=2, sticky="ew", pady=6)

        buttons = ttk.Frame(outer, padding=(0, 10, 0, 0))
        buttons.grid(row=3, column=0, sticky="ew")
        ttk.Button(buttons, text="Add", command=self._add).pack(side="left", padx=(0, 4))
        ttk.Button(buttons, text="Remove", command=self._remove).pack(side="left", padx=4)
        ttk.Button(buttons, text="Move Up", command=lambda: self._move(-1)).pack(side="left", padx=(12, 4))
        ttk.Button(buttons, text="Move Down", command=lambda: self._move(1)).pack(side="left", padx=4)
        ttk.Button(buttons, text="Browse File", command=self._browse_file).pack(side="left", padx=(12, 4))
        ttk.Button(buttons, text="Browse Folder", command=self._browse_folder).pack(side="left", padx=4)
        ttk.Button(buttons, text="Import mDIR", command=self._import_mdir).pack(side="right", padx=(8, 0))

        ttk.Label(outer, textvariable=self.status_var, style="Muted.TLabel").grid(row=4, column=0, sticky="w", pady=(8, 4))

        footer = ttk.Frame(outer)
        footer.grid(row=5, column=0, sticky="e", pady=(8, 0))
        ttk.Button(footer, text="Save", command=self._save, width=12).pack(side="left", padx=6)
        ttk.Button(footer, text="Cancel", command=self.destroy, width=12).pack(side="left")

    def _wire_live_editor(self) -> None:
        for variable in (self.name_var, self.type_var, self.pane_var, self.target_var, self.args_var):
            variable.trace_add("write", lambda *_args: self._sync_from_editor())

    def _display_type(self, kind: str) -> str:
        return kind.title()

    def _display_pane(self, pane: str) -> str:
        return {"active": "Active", "left": "Left / Files", "right": "Right / Workbook"}.get(pane, "Active")

    def _pane_internal(self, display: str) -> str:
        text = display.lower()
        if text.startswith("left"):
            return "left"
        if text.startswith("right"):
            return "right"
        return "active"

    def _type_internal(self, display: str) -> str:
        return display.lower().strip()

    def _render_rows(self, selected: Optional[int]) -> None:
        self._syncing = True
        try:
            for item in self.table.get_children(""):
                self.table.delete(item)
            for index, draft in enumerate(self.drafts):
                self.table.insert("", "end", iid=str(index), values=(
                    draft["label"],
                    self._display_type(draft["kind"]),
                    draft["target"],
                    self._display_pane(draft["pane"]),
                ))
            if selected is not None and self.drafts:
                selected = min(max(0, selected), len(self.drafts) - 1)
                self.table.selection_set(str(selected))
                self.table.focus(str(selected))
                self.table.see(str(selected))
                self._load_editor(selected)
            elif not self.drafts:
                self._editing_index = None
                self.name_var.set("")
                self.target_var.set("")
                self.args_var.set("[]")
        finally:
            self._syncing = False

    def _row_selected(self, _event=None) -> None:
        if self._syncing:
            return
        selected = self.table.selection()
        if not selected:
            return
        self._load_editor(int(selected[0]))

    def _load_editor(self, index: int) -> None:
        if not 0 <= index < len(self.drafts):
            return
        self._syncing = True
        try:
            self._editing_index = index
            draft = self.drafts[index]
            self.name_var.set(draft["label"])
            self.type_var.set(draft["kind"].title())
            self.pane_var.set(self._display_pane(draft["pane"]))
            self.target_var.set(draft["target"])
            self.args_var.set(draft["args"])
        finally:
            self._syncing = False

    def _sync_from_editor(self) -> None:
        if self._syncing or self._editing_index is None:
            return
        index = self._editing_index
        if not 0 <= index < len(self.drafts):
            return
        draft = self.drafts[index]
        draft["label"] = self.name_var.get()
        draft["kind"] = self._type_internal(self.type_var.get())
        draft["pane"] = self._pane_internal(self.pane_var.get())
        draft["target"] = self.target_var.get()
        draft["args"] = self.args_var.get()
        if self.table.exists(str(index)):
            self.table.item(str(index), values=(
                draft["label"], self._display_type(draft["kind"]), draft["target"], self._display_pane(draft["pane"])
            ))

    def _add(self) -> None:
        if len(self.drafts) >= MAX_LINKS:
            messagebox.showwarning("Link Manager", f"The shortcut bar supports up to {MAX_LINKS} links.", parent=self)
            return
        self.drafts.append({
            "label": "New Link",
            "kind": "folder",
            "target": str(self.current_path),
            "pane": "active",
            "args": "[]",
        })
        self._render_rows(len(self.drafts) - 1)
        self.name_entry.focus_set()
        self.name_entry.selection_range(0, "end")

    def _remove(self) -> None:
        if self._editing_index is None:
            return
        index = self._editing_index
        del self.drafts[index]
        self._editing_index = None
        self._render_rows(min(index, len(self.drafts) - 1) if self.drafts else None)

    def _move(self, delta: int) -> None:
        if self._editing_index is None:
            return
        index = self._editing_index
        other = index + delta
        if not 0 <= other < len(self.drafts):
            return
        self.drafts[index], self.drafts[other] = self.drafts[other], self.drafts[index]
        self._render_rows(other)

    def _browse_file(self) -> None:
        chosen = filedialog.askopenfilename(parent=self, initialdir=str(self.current_path), title="Choose file or program")
        if chosen:
            self.target_var.set(chosen)

    def _browse_folder(self) -> None:
        chosen = filedialog.askdirectory(parent=self, initialdir=str(self.current_path), title="Choose folder")
        if chosen:
            self.target_var.set(chosen)

    def _import_mdir(self) -> None:
        links = load_mdir_links(self.mdir_path)
        if not links:
            messagebox.showinfo("Link Manager", f"No mDIR links were found in:\n{self.mdir_path}", parent=self)
            return
        self.drafts = [self._draft(link) for link in links[:MAX_LINKS]]
        self._editing_index = None
        self._render_rows(0)
        self.status_var.set(f"Imported {len(self.drafts)} link(s) from mDIR. Save to apply them to xViewer.")

    def _validated(self) -> list[LinkDefinition] | None:
        self._sync_from_editor()
        result: list[LinkDefinition] = []
        for number, draft in enumerate(self.drafts, start=1):
            label = draft["label"].strip()
            kind = draft["kind"].strip().lower()
            pane = draft["pane"].strip().lower()
            target = draft["target"].strip()
            if not label or not target or kind not in LINK_KINDS or pane not in PANE_VALUES:
                messagebox.showwarning("Link Manager", f"Row {number} needs a name, valid type, pane, and target.", parent=self)
                return None
            try:
                args = json.loads(draft["args"].strip() or "[]")
                if not isinstance(args, list):
                    raise ValueError("Arguments must be a JSON list.")
            except (json.JSONDecodeError, ValueError) as exc:
                messagebox.showwarning("Link Manager", f"Row {number} has invalid Arguments:\n{exc}", parent=self)
                return None
            result.append(LinkDefinition(label[:24], kind, target, tuple(str(x) for x in args), pane))
        return result

    def _save(self) -> None:
        links = self._validated()
        if links is None:
            return
        self.on_save(links)
        self.destroy()

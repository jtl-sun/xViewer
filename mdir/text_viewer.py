from pathlib import Path
import json
import tkinter as tk
from tkinter import ttk
from .background import LatestWorker
from .ui_dispatch import install_dispatch

TEXT_EXTENSIONS = {'.txt', '.md', '.markdown', '.csv', '.json'}
MAX_TEXT_BYTES = 16 * 1024 * 1024


def read_text_preview(path):
    with Path(path).open('rb') as stream:
        raw = stream.read(MAX_TEXT_BYTES + 1)
    truncated = len(raw) > MAX_TEXT_BYTES
    raw = raw[:MAX_TEXT_BYTES]
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        text, encoding = raw.decode('utf-16', errors='replace'), 'utf-16'
    else:
        for encoding in ('utf-8-sig', 'utf-8', 'cp949', 'windows-1258', 'windows-1252'):
            try:
                text = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            text, encoding = raw.decode('utf-8', errors='replace'), 'utf-8 (replacement)'
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    if Path(path).suffix.lower() == '.json' and not truncated:
        try:
            text = json.dumps(json.loads(text), indent=2, ensure_ascii=False)
        except (ValueError, RecursionError):
            pass
    if truncated:
        text += '\n\n[Preview limited to the first 16 MiB. Open externally to read the complete file.]'
    return text, encoding


class TextViewer(ttk.Frame):
    def __init__(self, parent, *, status_callback, ui_palette):
        super().__init__(parent)
        self.status_callback = status_callback
        self._post = install_dispatch(self)
        self._worker = LatestWorker('xViewer-Text')
        self._generation = 0
        self.bind('<Destroy>', lambda e: self._worker.close() if e.widget is self else None, add='+')
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.text = tk.Text(self, wrap='none', undo=False, state='disabled', font=('Consolas', 11))
        ybar = ttk.Scrollbar(self, orient='vertical', command=self.text.yview)
        xbar = ttk.Scrollbar(self, orient='horizontal', command=self.text.xview)
        self.text.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.text.grid(row=0, column=0, sticky='nsew')
        ybar.grid(row=0, column=1, sticky='ns')
        xbar.grid(row=1, column=0, sticky='ew')
        for key in ('<Control-a>', '<Control-A>'):
            self.text.bind(key, self.select_all)
        for key in ('<Control-c>', '<Control-C>'):
            self.text.bind(key, self.copy_current)
        self.apply_palette(ui_palette)

    def apply_palette(self, palette):
        self.text.configure(bg=palette.get('cell_bg', palette.get('window', '#ffffff')),
                            fg=palette.get('foreground', '#202020'))

    def _set_text(self, value):
        self.text.configure(state='normal')
        self.text.delete('1.0', 'end')
        self.text.insert('1.0', value)
        self.text.configure(state='disabled')
        self.text.xview_moveto(0)
        self.text.yview_moveto(0)

    def clear(self):
        self._generation += 1
        self._worker.cancel()
        self._set_text('')

    def load(self, path):
        self.clear()
        generation = self._generation
        self._set_text('Loading text ...')
        def complete(value, encoding):
            if generation == self._generation:
                self._set_text(value)
                self.status_callback(f'{path} | Text | {encoding} | Read only | Ctrl+A select all / Ctrl+C copy')
        def worker():
            try:
                value, encoding = read_text_preview(path)
            except Exception as exc:
                value, encoding = f'Could not open text:\n{exc}', 'error'
            self._post(lambda: complete(value, encoding))
        self._worker.submit(worker)

    def select_all(self, _event=None):
        self.text.tag_add('sel', '1.0', 'end-1c')
        return 'break'

    def copy_current(self, _event=None):
        ranges = self.text.tag_ranges('sel')
        value = self.text.get(*ranges[:2]) if ranges else self.text.get('1.0', 'end-1c')
        self.clipboard_clear()
        self.clipboard_append(value)
        self.status_callback('Text copied to clipboard.')
        return 'break'

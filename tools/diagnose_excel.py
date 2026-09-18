"""Run an Excel preview diagnostic in the interactive Windows user session."""
from pathlib import Path
import json
import os
import sys
from datetime import datetime
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from mdir import __version__
from mdir.excel_pdf_preview import render_excel_pdf_cached, _cache_dir
from mdir.media_viewer import _render_pdf_page


def main():
    if len(sys.argv) > 1:
        source = Path(sys.argv[1]).resolve()
    else:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        name = filedialog.askopenfilename(title='Select the Excel file to diagnose',
            filetypes=[('Excel', '*.xls *.xlsx *.xlsm *.xltx *.xltm')])
        root.destroy()
        if not name:
            return 0
        source = Path(name).resolve()
    started = time.monotonic()
    result = render_excel_pdf_cached(source)
    report = dict(version=__version__, source=str(source), ok=bool(result.ok),
                  backend=result.backend, from_cache=result.from_cache,
                  seconds=result.elapsed, error=result.error)
    folder = _cache_dir().parent.parent/'logs'
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d-%H%M%S')
    if result.ok:
        try:
            image, count = _render_pdf_page(result.pdf_path, 0, 1.0)
            report.update(pages=count, first_page_size=image.size, first_render_seconds=time.monotonic()-started-result.elapsed)
            # The PDF itself remains in the local cache for manual comparison.
            report['pdf_path'] = str(result.pdf_path)
        except Exception as exc:
            report.update(ok=False, pdf_render_error=str(exc))
    target = folder/f'diagnostic-{stamp}.json'
    target.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding='utf-8')
    print(json.dumps(report, indent=2, ensure_ascii=True))
    print(f'Report: {target}')
    return 0 if report['ok'] else 1


if __name__ == '__main__':
    raise SystemExit(main())

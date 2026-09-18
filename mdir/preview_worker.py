"""Private, disposable Excel COM host. Never attach to the user's Excel session."""
from __future__ import annotations
import json
import sys
import time


def emit(**event):
    print(json.dumps(event, ensure_ascii=True), flush=True)


def configure_sheet(app, sheet):
    # Preserve paper size, margins and landscape; avoid Find/Shapes iteration.
    if sheet.Visible != -1:
        return
    try:
        app.PrintCommunication = False
        setup = sheet.PageSetup
        setup.PrintArea = ""
        setup.Zoom = False
        setup.FitToPagesWide = 1
        setup.FitToPagesTall = 1
    finally:
        app.PrintCommunication = True


def main():
    app = None
    com = None
    try:
        import pythoncom
        import win32com.client
        import win32process
        com = pythoncom
        com.CoInitialize()
        for line in sys.stdin:
            job = json.loads(line)
            book = None
            started = time.monotonic()
            stage = 'startup'
            def progress(name):
                nonlocal stage
                stage = name
                emit(stage=name, elapsed=time.monotonic() - started)
            try:
                if app is None:
                    progress('startup')
                    app = win32com.client.DispatchEx('Excel.Application')
                    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
                    emit(excel_pid=pid)
                    app.Visible = False
                    app.DisplayAlerts = False
                    app.EnableEvents = False
                    app.ScreenUpdating = False
                    app.AskToUpdateLinks = False
                    app.AutomationSecurity = 3
                progress('open')
                book = app.Workbooks.Open(
                    job['source'], UpdateLinks=0, ReadOnly=True, Password='',
                    WriteResPassword='', IgnoreReadOnlyRecommended=True,
                    Notify=False, AddToMru=False,
                )
                try:
                    app.Calculation = -4135  # manual; never recalculate on export
                    app.CalculateBeforeSave = False
                except Exception:
                    pass
                if job.get('layout', True):
                    progress('page_setup')
                    for sheet in book.Worksheets:
                        try:
                            configure_sheet(app, sheet)
                        except Exception as exc:
                            emit(warning='Page setup: ' + str(exc))
                progress('export')
                book.ExportAsFixedFormat(0, job['target'], 0, False, True,
                                         com.Missing, com.Missing, False)
                progress('close')
                book.Close(False)
                book = None
                emit(done=True, elapsed=time.monotonic() - started)
            except Exception as exc:
                # Report before cleanup: a Close/Quit call can itself hang.
                emit(done=False, stage=stage, error=str(exc), elapsed=time.monotonic() - started)
                return 1
    except Exception as exc:
        emit(done=False, stage='startup', error=str(exc))
        return 1
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if com is not None:
            com.CoUninitialize()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

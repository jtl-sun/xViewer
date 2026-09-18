from __future__ import annotations

import hashlib
import json
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
from dataclasses import dataclass
from typing import Optional

EXCEL_PREVIEW_EXTENSIONS = {'.xlsx', '.xlsm', '.xltx', '.xltm', '.xls'}
_CACHE_VERSION = '2.7.4-isolated-excel-v1'
FOREGROUND_TIMEOUT = 15.0
FALLBACK_TIMEOUT = 15.0


@dataclass(frozen=True)
class ExcelPdfResult:
    source: Path
    pdf_path: Optional[Path]
    from_cache: bool
    backend: str
    elapsed: float
    error: Optional[str] = None

    @property
    def ok(self):
        return self.pdf_path is not None and self.pdf_path.is_file() and not self.error


def _cache_dir():
    return Path(os.environ.get('LOCALAPPDATA') or tempfile.gettempdir()) / 'xViewer' / 'cache' / 'excel-pdf'


_LOG_LOCK = threading.Lock()


def _diagnostic(event, **details):
    # Local rotating diagnostics contain paths/timings, never workbook contents.
    try:
        with _LOG_LOCK:
            log = logging.getLogger('xviewer.preview')
            if not log.handlers:
                folder = _cache_dir().parent.parent / 'logs'
                folder.mkdir(parents=True, exist_ok=True)
                handler = RotatingFileHandler(folder / 'excel-preview.log', maxBytes=2_000_000, backupCount=2, encoding='utf-8')
                handler.setFormatter(logging.Formatter('%(asctime)s %(message)s'))
                log.addHandler(handler)
                log.setLevel(logging.INFO)
                log.propagate = False
            log.info(json.dumps(dict(event=event, **details), ensure_ascii=False, default=str))
    except Exception:
        pass


def excel_pdf_cache_path(source):
    source = Path(source).resolve()
    stat = source.stat()
    value = f'{_CACHE_VERSION}|{source}|{stat.st_size}|{stat.st_mtime_ns}'
    return _cache_dir() / (hashlib.sha256(value.encode('utf-8')).hexdigest()[:28] + '.pdf')


def _valid_pdf(path):
    try:
        with Path(path).open('rb') as stream:
            if stream.read(5) != b'%PDF-':
                return False
            stream.seek(0, 2)
            size = stream.tell()
            stream.seek(max(0, size - 2048))
            return b'%%EOF' in stream.read()
    except OSError:
        return False


def excel_pdf_cache_hit(source):
    try:
        target = excel_pdf_cache_path(source)
        return target if _valid_pdf(target) else None
    except OSError:
        return None


def prune_excel_pdf_cache(max_files=750, max_bytes=3 * 1024**3):
    try:
        entries = [(p.stat().st_mtime, p.stat().st_size, p) for p in _cache_dir().glob('*.pdf')]
        total, count = sum(row[1] for row in entries), len(entries)
        for _, size, path in sorted(entries):
            if count <= max_files and total <= max_bytes:
                break
            try:
                path.unlink()
                total -= size
                count -= 1
            except OSError:
                pass
    except OSError:
        pass


class _PrivateExcelHandle:
    """Retain a handle, not a reusable PID, to this worker's new Excel process."""
    def __init__(self, pid, started):
        self.handle = None
        handle = None
        try:
            import win32api
            import win32process
            handle = win32api.OpenProcess(0x1000 | 0x100000 | 1, False, pid)
            created = win32process.GetProcessTimes(handle)['CreationTime'].timestamp()
            if created >= started - 0.1:
                self.handle = handle
                handle = None
        except Exception as exc:
            _diagnostic('process_handle_unavailable', error=str(exc))
        finally:
            if handle is not None:
                handle.Close()

    def close(self):
        if self.handle is not None:
            try:
                import win32api
                win32api.TerminateProcess(self.handle, 1)
            except Exception:
                pass
            finally:
                self.handle.Close()
                self.handle = None


class _PersistentExcelCom:
    """Supervise disposable COM hosts so every COM call has a real deadline."""
    def __init__(self, command=None):
        executable = Path(sys.executable)
        if executable.name.lower() == 'pythonw.exe':
            executable = executable.with_name('python.exe')
        self.command = command or [str(executable), '-u', '-m', 'mdir.preview_worker']
        self.process = None
        self.events = None
        self.excel_handle = None
        self.error = ''
        self.stage = 'startup'

    def start(self):
        if self.process is not None and self.process.poll() is None:
            return
        self.stop()
        self.started = time.time()
        self.events = queue.Queue()
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        env['PYTHONPATH'] = str(Path(__file__).resolve().parent.parent) + os.pathsep + env.get('PYTHONPATH', '')
        self.process = subprocess.Popen(self.command, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL, text=True, encoding='utf-8', env=env,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        stream, events = self.process.stdout, self.events
        def read_events():
            try:
                for line in stream:
                    try:
                        events.put(json.loads(line))
                    except ValueError:
                        pass
            finally:
                stream.close()
                events.put(dict(done=False, error='Excel preview worker exited.'))
        threading.Thread(target=read_events, daemon=True).start()

    def render(self, source, target, *, cancel=None, timeout=FOREGROUND_TIMEOUT, layout=True):
        cancel = cancel or threading.Event()
        deadline = time.monotonic() + timeout
        self.error, self.stage = '', 'startup'
        try:
            self.start()
            self.process.stdin.write(json.dumps(dict(source=str(Path(source).resolve()), target=str(Path(target).resolve()), layout=layout)) + '\n')
            self.process.stdin.flush()
            while True:
                if cancel.is_set():
                    self.error = 'cancelled'
                    break
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    self.error = f'Excel preview timed out at {self.stage} ({timeout:g}s).'
                    break
                try:
                    event = self.events.get(timeout=min(0.05, remaining))
                except queue.Empty:
                    continue
                if 'excel_pid' in event:
                    self.excel_handle = _PrivateExcelHandle(event['excel_pid'], self.started)
                if 'stage' in event:
                    self.stage = event['stage']
                _diagnostic('worker', source=source, **event)
                if 'done' in event:
                    if event['done'] and _valid_pdf(target):
                        return True
                    self.error = event.get('error', 'Excel did not create a complete PDF.')
                    break
        except Exception as exc:
            self.error = str(exc)
        _diagnostic('render_failed', source=source, stage=self.stage, error=self.error)
        self.stop()
        return False

    def stop(self):
        process, self.process = self.process, None
        if process is not None:
            try:
                if process.poll() is None:
                    process.terminate()
                    process.wait(timeout=0.5)
            except Exception:
                try:
                    process.kill()
                    process.wait(timeout=0.5)
                except Exception:
                    pass
            try:
                process.stdin.close()
            except Exception:
                pass
        if self.excel_handle is not None:
            self.excel_handle.close()
            self.excel_handle = None


def _find_libreoffice():
    candidates = [shutil.which('soffice'), r'C:\Program Files\LibreOffice\program\soffice.exe',
                  r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']
    return next((str(p) for p in candidates if p and Path(p).is_file()), None)


def _libreoffice_render(source, target, *, cancel=None, timeout=FALLBACK_TIMEOUT):
    executable = _find_libreoffice()
    if not executable:
        return False
    cancel = cancel or threading.Event()
    with tempfile.TemporaryDirectory(prefix='xviewer-lo-') as folder:
        root = Path(folder)
        process = subprocess.Popen([executable, '-env:UserInstallation=' + (root / 'profile').as_uri(),
            '--headless', '--convert-to', 'pdf', '--outdir', str(root), str(Path(source).resolve())],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        try:
            deadline = time.monotonic() + timeout
            while process.poll() is None:
                if cancel.wait(0.05) or time.monotonic() >= deadline:
                    return False
            output = root / (Path(source).stem + '.pdf')
            if process.returncode == 0 and _valid_pdf(output):
                shutil.copyfile(output, target)
                return True
            return False
        finally:
            if process.poll() is None:
                if os.name == 'nt':
                    subprocess.run(['taskkill', '/PID', str(process.pid), '/T', '/F'],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=3,
                        creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                else:
                    process.kill()
                process.wait(timeout=3)


def render_excel_pdf_cached(source, *, session=None, cancel=None):
    source = Path(source)
    started = time.monotonic()
    cancel = cancel or threading.Event()
    own_session = session is None
    session = session or _PersistentExcelCom()
    temporary = None
    try:
        if cancel.is_set():
            raise RuntimeError('cancelled')
        if source.suffix.lower() not in EXCEL_PREVIEW_EXTENSIONS:
            raise ValueError('Unsupported Excel extension.')
        cached = excel_pdf_cache_hit(source)
        if cached is not None:
            try:
                os.utime(cached, None)
            except OSError:
                pass
            return ExcelPdfResult(source, cached, True, 'cache', time.monotonic() - started)
        target = excel_pdf_cache_path(source)
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_name(target.stem + '.' + uuid.uuid4().hex + '.tmp.pdf')
        backend = 'Microsoft Excel'
        ok = session.render(source, temporary, cancel=cancel)
        if not ok and not cancel.is_set():
            if session.stage == 'page_setup':
                # Exactly one bounded retry using the workbook's original layout.
                ok = session.render(source, temporary, cancel=cancel, timeout=FALLBACK_TIMEOUT, layout=False)
            else:
                backend = 'LibreOffice'
                ok = _libreoffice_render(source, temporary, cancel=cancel)
        if cancel.is_set():
            raise RuntimeError('cancelled')
        if not ok or not _valid_pdf(temporary):
            raise RuntimeError(session.error or 'Microsoft Excel or LibreOffice is required for faithful preview.')
        if target != excel_pdf_cache_path(source):
            raise RuntimeError('The workbook changed during preview; select it again.')
        os.replace(temporary, target)
        prune_excel_pdf_cache()
        _diagnostic('complete', source=source, backend=backend, elapsed=time.monotonic() - started)
        return ExcelPdfResult(source, target, False, backend, time.monotonic() - started)
    except Exception as exc:
        _diagnostic('failed', source=source, error=str(exc), elapsed=time.monotonic() - started)
        return ExcelPdfResult(source, None, False, 'failed', time.monotonic() - started, str(exc))
    finally:
        if own_session:
            session.stop()
        if temporary is not None:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass


class ExcelPdfPreviewEngine:
    """One active render and one replaceable foreground request. No prefetch."""
    def __init__(self):
        self._condition = threading.Condition()
        self._pending = None
        self._active_cancel = None
        self._closed = False
        self._thread = threading.Thread(target=self._worker, name='xViewer-Excel-PDF-Engine', daemon=True)
        self._thread.start()

    def request(self, source, callback=None, *, priority=0):
        if priority > 0:
            return None
        source = Path(source)
        cached = excel_pdf_cache_hit(source)
        with self._condition:
            if self._closed:
                return None
            if self._active_cancel is not None:
                self._active_cancel.set()
            self._pending = None if cached else (source, callback)
            self._condition.notify()
        if cached is not None and callback is not None:
            callback(ExcelPdfResult(source, cached, True, 'cache', 0.0))
        return cached

    def cancel(self):
        with self._condition:
            self._pending = None
            if self._active_cancel is not None:
                self._active_cancel.set()
            self._condition.notify()

    def prefetch(self, sources):
        pass  # Kept for compatibility; speculative conversion is disabled.

    def close(self):
        with self._condition:
            self._closed = True
            self.cancel()
            self._condition.notify_all()
        # Let the supervisor reap its child/private Excel before Python exits.
        # Never wait indefinitely for Office shutdown.
        if threading.current_thread() is not self._thread:
            self._thread.join(timeout=4)

    def _worker(self):
        session = _PersistentExcelCom()
        try:
            while True:
                with self._condition:
                    while self._pending is None and not self._closed:
                        self._condition.wait()
                    if self._closed:
                        return
                    source, callback = self._pending
                    self._pending = None
                    cancel = self._active_cancel = threading.Event()
                try:
                    result = render_excel_pdf_cached(source, session=session, cancel=cancel)
                except Exception as exc:
                    result = ExcelPdfResult(source, None, False, 'failed', 0.0, str(exc))
                if callback is not None and not cancel.is_set():
                    try:
                        callback(result)
                    except Exception as exc:
                        _diagnostic('callback_failed', error=str(exc))
        finally:
            session.stop()


_ENGINE = None
_ENGINE_LOCK = threading.Lock()


def get_excel_pdf_engine():
    global _ENGINE
    with _ENGINE_LOCK:
        if _ENGINE is None:
            _ENGINE = ExcelPdfPreviewEngine()
        return _ENGINE


def shutdown_excel_pdf_engine():
    global _ENGINE
    with _ENGINE_LOCK:
        engine, _ENGINE = _ENGINE, None
    if engine is not None:
        engine.close()

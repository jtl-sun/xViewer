import json
import os
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
import threading
import time
import unittest
from unittest.mock import patch, Mock

from PIL import Image, ImageDraw
from openpyxl import Workbook
from mdir import excel_pdf_preview as preview
from mdir.media_viewer import trim_excel_whitespace, _render_pdf_page
from mdir.text_viewer import read_text_preview
from mdir.workspace_app import EditableWorkbookModel, _type_filter_accepts
from mdir.preview_worker import configure_sheet


class PreviewReliabilityTests(unittest.TestCase):
    def test_shift_tab_survives_unsupported_iso_keysym(self):
        import tkinter as tk
        from mdir.workspace_app import _bind_shift_tab
        widget = Mock()
        def bind(sequence, callback):
            if sequence == '<ISO_Left_Tab>':
                raise tk.TclError('bad event type or keysym "ISO_Left_Tab"')
        widget.bind.side_effect = bind
        callback = Mock()
        _bind_shift_tab(widget, callback)
        self.assertEqual(widget.bind.call_args_list[0].args, ('<Shift-Tab>', callback))

    @unittest.skipUnless(os.name == 'nt', 'Windows process handle API')
    def test_only_new_private_excel_handle_can_be_terminated(self):
        from datetime import datetime, timezone
        for created, should_terminate in [(90, False), (101, True)]:
            handle = Mock()
            with patch('win32api.OpenProcess', return_value=handle), patch('win32process.GetProcessTimes',
                    return_value={'CreationTime': datetime.fromtimestamp(created, timezone.utc)}), patch('win32api.TerminateProcess') as terminate:
                owned = preview._PrivateExcelHandle(12345, 100)
                owned.close()
                if should_terminate:
                    terminate.assert_called_once_with(handle, 1)
                else:
                    terminate.assert_not_called()
                handle.Close.assert_called_once()

    def test_real_child_hang_has_deadline_and_is_reaped(self):
        session = preview._PersistentExcelCom([sys.executable, '-u', '-c', 'import time; time.sleep(30)'])
        started = time.monotonic()
        self.assertFalse(session.render('unused.xlsx', 'unused.pdf', timeout=0.3))
        self.assertLess(time.monotonic()-started, 3)
        self.assertIn('timed out', session.error)
        self.assertIsNone(session.process)

    def test_real_child_hang_is_cancelled_promptly(self):
        cancel = threading.Event()
        session = preview._PersistentExcelCom([sys.executable, '-u', '-c', 'import time; time.sleep(30)'])
        timer = threading.Timer(0.2, cancel.set)
        timer.start()
        started = time.monotonic()
        try:
            self.assertFalse(session.render('unused.xlsx', 'unused.pdf', cancel=cancel))
            self.assertEqual(session.error, 'cancelled')
            self.assertLess(time.monotonic()-started, 3)
        finally:
            timer.cancel()
            session.stop()

    def test_worker_can_restart_after_timeout(self):
        with TemporaryDirectory() as td:
            root = Path(td)
            session = preview._PersistentExcelCom([sys.executable, '-u', '-c', 'import time; time.sleep(30)'])
            self.assertFalse(session.render(root/'a.xlsx', root/'a.pdf', timeout=0.2))
            code = "import sys,json; from PIL import Image; job=json.loads(sys.stdin.readline()); Image.new('RGB',(50,50),'white').save(job['target'],'PDF'); print(json.dumps({'done':True}), flush=True); sys.stdin.readline()"
            session.command = [sys.executable, '-u', '-c', code]
            try:
                self.assertTrue(session.render(root/'a.xlsx', root/'a.pdf', timeout=3))
            finally:
                session.stop()

    def test_latest_foreground_replaces_queued_selection(self):
        started, cancelled, finished = threading.Event(), threading.Event(), threading.Event()
        calls, results = [], []
        def render(source, *, session, cancel):
            calls.append(source.name)
            if source.name == 'a.xlsx':
                started.set()
                self.assertTrue(cancel.wait(2))
                cancelled.set()
            return preview.ExcelPdfResult(source, None, False, 'test', 0, 'test')
        def callback(result):
            results.append(result.source.name)
            finished.set()
        with patch.object(preview, 'render_excel_pdf_cached', side_effect=render):
            engine = preview.ExcelPdfPreviewEngine()
            try:
                engine.request('a.xlsx', callback)
                self.assertTrue(started.wait(2))
                with engine._condition:
                    engine.request('b.xlsx', callback)
                    engine.request('c.xlsx', callback)
                self.assertTrue(cancelled.wait(2))
                self.assertTrue(finished.wait(2))
                self.assertEqual(calls, ['a.xlsx', 'c.xlsx'])
                self.assertEqual(results, ['c.xlsx'])
            finally:
                engine.close()
                engine._thread.join(2)

    def test_unexpected_render_exception_does_not_kill_engine(self):
        finished, results = threading.Event(), []
        def callback(result):
            results.append(result)
            finished.set()
        with patch.object(preview, 'render_excel_pdf_cached', side_effect=PermissionError('cache is read-only')):
            engine = preview.ExcelPdfPreviewEngine()
            try:
                for name in ('a.xlsx', 'b.xlsx'):
                    finished.clear()
                    engine.request(name, callback)
                    self.assertTrue(finished.wait(2))
                self.assertEqual(len(results), 2)
                self.assertTrue(engine._thread.is_alive())
                self.assertTrue(all('read-only' in r.error for r in results))
            finally:
                engine.close()
                engine._thread.join(2)

    def test_permission_error_is_returned_to_caller(self):
        with patch.object(preview, 'excel_pdf_cache_path', side_effect=PermissionError('denied')):
            result = preview.render_excel_pdf_cached('a.xlsx')
        self.assertFalse(result.ok)
        self.assertIn('denied', result.error)

    def test_partial_pdf_is_not_a_cache_hit(self):
        with TemporaryDirectory() as td, patch.object(preview, '_cache_dir', return_value=Path(td)/'cache'):
            source = Path(td)/'a.xls'
            source.write_bytes(b'test')
            target = preview.excel_pdf_cache_path(source)
            target.parent.mkdir()
            target.write_bytes(b'%PDF-1.4\n' + b'x'*256)
            self.assertIsNone(preview.excel_pdf_cache_hit(source))

    def test_source_change_during_export_is_not_cached(self):
        with TemporaryDirectory() as td, patch.object(preview, '_cache_dir', return_value=Path(td)/'cache'):
            source = Path(td)/'a.xls'
            source.write_bytes(b'initial')
            target = preview.excel_pdf_cache_path(source)
            def render(source, temporary, **kwargs):
                Image.new('RGB', (30, 30), 'white').save(temporary, 'PDF')
                source.write_bytes(b'changed-content')
                return True
            session = Mock(render=Mock(side_effect=render))
            result = preview.render_excel_pdf_cached(source, session=session)
            self.assertFalse(result.ok)
            self.assertIn('changed', result.error)
            self.assertFalse(target.exists())
            self.assertEqual(list(target.parent.iterdir()), [])

    def test_page_setup_preserves_orientation_paper_and_margins(self):
        class Obj: pass
        app, sheet, setup = Obj(), Obj(), Obj()
        sheet.Visible, sheet.PageSetup = -1, setup
        setup.Orientation, setup.PaperSize, setup.LeftMargin = 2, 9, 24
        configure_sheet(app, sheet)
        self.assertEqual((setup.Orientation, setup.PaperSize, setup.LeftMargin), (2, 9, 24))
        self.assertEqual((setup.PrintArea, setup.Zoom, setup.FitToPagesWide, setup.FitToPagesTall), ('', False, 1, 1))
        self.assertTrue(app.PrintCommunication)

    def test_trim_preserves_faint_content_and_padding(self):
        image = Image.new('RGB', (300, 200), 'white')
        ImageDraw.Draw(image).rectangle((100, 50, 199, 99), fill=(254, 254, 254))
        cropped = trim_excel_whitespace(image)
        self.assertEqual(cropped.size, (124, 74))
        self.assertEqual(image.size, (300, 200))
        blank = Image.new('RGB', (100, 100), 'white')
        self.assertEqual(trim_excel_whitespace(blank).size, blank.size)

    def test_read_only_native_fallback_does_not_call_legacy_com(self):
        with TemporaryDirectory() as td:
            path = Path(td)/'test.xlsx'
            book = Workbook()
            book.active['A1'] = 'unchanged'
            book.save(path)
            before = path.read_bytes()
            model = EditableWorkbookModel(path, allow_legacy_conversion=False, view_only=True)
            try:
                self.assertFalse(model.editable)
                with self.assertRaises(RuntimeError):
                    model.save()
                self.assertEqual(path.read_bytes(), before)
            finally:
                model.close()
        with patch('mdir.workspace_app.legacy_xls_preview_path') as converter, patch('mdir.workspace_app.legacy_xls_images') as images:
            with self.assertRaises(FileNotFoundError):
                EditableWorkbookModel(Path('missing.xls'), allow_legacy_conversion=False, view_only=True)
            converter.assert_not_called()
            images.assert_not_called()

    def test_text_encoding_and_json_format(self):
        with TemporaryDirectory() as td:
            path = Path(td)/'한글.txt'
            for encoding in ('utf-8-sig', 'utf-16', 'cp949'):
                path.write_bytes('한글 내용'.encode(encoding))
                self.assertEqual(read_text_preview(path)[0], '한글 내용')
            path = Path(td)/'sample.json'
            path.write_text('{"name":"한글","count":2}', encoding='utf-8')
            text, encoding = read_text_preview(path)
            self.assertIn('\n  "name": "한글"', text)

    def test_xlsx_named_xls_opens_read_only_without_changing_original(self):
        with TemporaryDirectory() as td:
            path = Path(td)/'supplier.xls'
            book = Workbook()
            book.active['A1'] = 'source'
            book.save(path)
            before = path.read_bytes()
            model = EditableWorkbookModel(path, allow_legacy_conversion=False)
            try:
                self.assertFalse(model.editable)
                self.assertEqual(model.sheets[0].value(1, 1), 'source')
            finally:
                model.close()
            self.assertEqual(path.read_bytes(), before)

    def test_filters_are_union_for_all_eight_combinations(self):
        for mask in range(8):
            opts = dict(excel_only=bool(mask&1), pdf_only=bool(mask&2), images_only=bool(mask&4))
            for path, bit in [('a.xlsx', 1), ('a.pdf', 2), ('a.png', 4), ('a.txt', 0)]:
                self.assertEqual(_type_filter_accepts(path, False, **opts), mask == 0 or bool(mask&bit))
            self.assertTrue(_type_filter_accepts('folder', True, **opts))


if __name__ == '__main__':
    unittest.main()

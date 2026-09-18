"""Hidden Tk integration checks: real loaders and event loop, no Office needed."""
from pathlib import Path
from tempfile import TemporaryDirectory
import time
import tkinter as tk
import unittest
from unittest.mock import patch, Mock
from PIL import Image
from openpyxl import Workbook
from mdir.workspace_app import ExcelWorkspaceApp
from mdir.excel_pdf_preview import ExcelPdfResult


class PreviewUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            root = tk.Tk()
            root.withdraw()
            root.destroy()
        except tk.TclError as exc:
            raise unittest.SkipTest(f'Tk display unavailable: {exc}')

    def setUp(self):
        self.temporary = TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.engine = Mock()
        self.patches = [patch('mdir.workspace_app.get_excel_pdf_engine', return_value=self.engine),
                        patch('mdir.workspace_app.ExcelWorkspaceApp._save_config'),
                        patch('mdir.workspace_app._load_quick_links', return_value=[])]
        for item in self.patches:
            item.start()
        self.app = ExcelWorkspaceApp(self.root)
        self.app.withdraw()
        self.errors = []
        self.app.report_callback_exception = lambda *args: self.errors.append(args)
        # Disable startup focus/list automation in these direct loader tests.
        for timer in self.app.tk.call('after', 'info'):
            script = str(self.app.tk.call('after', 'info', timer))
            if '_activate_left_panel' in script:
                self.app.after_cancel(timer)

    def tearDown(self):
        self.app._native_worker.close()
        if self.app.model is not None:
            self.app.model.close()
        for timer in self.app.tk.call('after', 'info'):
            # Cancel scheduling only; each owning widget deletes its own Tcl
            # callback commands during destroy (root must not delete a child's).
            self.app.tk.call('after', 'cancel', timer)
        self.app.destroy()
        for item in reversed(self.patches):
            item.stop()
        self.temporary.cleanup()
        self.assertEqual(self.errors, [])

    def pump(self, predicate, seconds=3):
        deadline = time.monotonic()+seconds
        while not predicate() and time.monotonic()<deadline:
            self.app.update()
            time.sleep(0.01)
        self.assertTrue(predicate(), 'UI completion timed out')

    def test_pdf_image_selection_and_error_delivery(self):
        for extension, kind in [('pdf', 'pdf'), ('png', 'image')]:
            path = self.root/f'good.{extension}'
            Image.new('RGB', (240, 100), 'white').save(path)
            with patch.object(self.app, 'selected_file_paths', return_value=[path]):
                self.app._file_selected()
            self.pump(lambda: self.app.media_view._display_image is not None)
            self.assertEqual(self.app.viewer_kind, kind)
            self.assertTrue(self.app.media_view.image_fit_mode)
            self.app._toggle_media_fit_one_to_one()
            self.assertFalse(self.app.media_view.image_fit_mode)
            self.app._toggle_media_fit_one_to_one()
            self.assertTrue(self.app.media_view.image_fit_mode)
        broken = self.root/'broken.pdf'
        broken.write_bytes(b'not a PDF')
        self.app.open_media(broken, 'pdf')
        self.pump(lambda: 'Open failed' in self.app.status_text.get())
        self.assertIsNone(self.app.media_view._display_image)

    def test_text_is_read_only_copy_selection_or_all(self):
        path = self.root/'note.txt'
        path.write_text('alpha\nbeta', encoding='utf-8')
        self.app.open_text(path)
        self.pump(lambda: 'alpha' in self.app.text_view.text.get('1.0','end'))
        self.assertEqual(str(self.app.text_view.text.cget('state')), 'disabled')
        with patch.object(self.app.text_view, 'clipboard_clear'), patch.object(self.app.text_view, 'clipboard_append') as copy:
            self.app.copy_current()
            copy.assert_called_with('alpha\nbeta')
            self.app.text_view.text.tag_add('sel', '1.0', '1.5')
            self.app.copy_current()
            copy.assert_called_with('alpha')
        self.app._activate_right_panel()
        self.app.excel_only.set(True)
        self.app.pdf_only.set(True)
        self.app._type_filter_changed('pdf')
        self.assertTrue(self.app.excel_only.get())
        self.assertTrue(self.app.pdf_only.get())

    def test_native_fallback_completes_read_only_and_failure_clears_loading(self):
        path = self.root/'good.xlsx'
        book = Workbook()
        book.active['A1'] = 'view only'
        book.save(path)
        self.app.path = self.app._loading_path = path
        self.app._open_workbook_native_fallback(path, self.app._load_generation, 'Office unavailable')
        self.pump(lambda: self.app.model is not None)
        self.assertFalse(self.app.model.editable)
        self.assertIsNone(self.app._loading_path)
        self.app._clear_current_workbook_view()
        path = self.root/'broken.xlsx'
        path.write_bytes(b'broken')
        self.app.path = self.app._loading_path = path
        self.app._open_workbook_native_fallback(path, self.app._load_generation, 'Office unavailable')
        self.pump(lambda: self.app.viewer_kind == 'failed')
        self.assertIsNone(self.app._loading_path)

    def test_stale_excel_result_does_not_overwrite_new_text(self):
        old_generation = self.app._load_generation
        path = self.root/'new.txt'
        path.write_text('new selection', encoding='utf-8')
        self.app.open_text(path)
        self.app._excel_pdf_render_complete(Path('old.xlsx'), old_generation,
            ExcelPdfResult(Path('old.xlsx'), None, False, 'test', 0, 'failed'))
        self.pump(lambda: 'new selection' in self.app.text_view.text.get('1.0', 'end'))
        self.assertEqual(self.app.path, path)
        self.assertEqual(self.app.viewer_kind, 'text')

    def test_corrupt_excel_pdf_cache_falls_back_to_read_only_cells(self):
        source = self.root/'source.xlsx'
        book = Workbook()
        book.active['A1'] = 'available'
        book.save(source)
        pdf = self.root/'broken-cache.pdf'
        pdf.write_bytes(b'%PDF-1.4\nnot actually a pdf\n%%EOF')
        self.app._display_excel_pdf_preview(source, self.app._load_generation,
            ExcelPdfResult(source, pdf, True, 'cache', 0))
        self.pump(lambda: self.app.model is not None)
        self.assertFalse(pdf.exists())
        self.assertFalse(self.app.model.editable)
        self.assertEqual(self.app.model.sheets[0].value(1, 1), 'available')


if __name__ == '__main__':
    unittest.main()

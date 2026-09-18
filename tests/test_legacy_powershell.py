"""Exercise the actual Windows PowerShell parser, without requiring Excel."""
import os
from pathlib import Path
import shutil
import subprocess
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from mdir import excel_viewer as viewer


@unittest.skipUnless(os.name == "nt" and shutil.which("powershell.exe"), "Windows PowerShell required")
class PowerShellExecutionTests(unittest.TestCase):
    def test_multiline_function_try_finally_and_unicode_path_execute(self):
        script = """$ErrorActionPreference = 'Stop'
function Write-Probe {
    [IO.File]::WriteAllText($env:XVIEWER_TEST_OUTPUT, 'executed')
}
try {
    Write-Probe
}
finally {
    [IO.File]::AppendAllText($env:XVIEWER_TEST_OUTPUT, '+finally')
}
"""
        with TemporaryDirectory() as td:
            target = Path(td) / "한글 space ' $ report.txt"
            env = dict(os.environ, XVIEWER_TEST_OUTPUT=str(target))
            completed = subprocess.run(
                viewer._powershell_script_command(shutil.which("powershell.exe"), script),
                stdin=subprocess.DEVNULL, capture_output=True, timeout=15, env=env,
                **viewer._hidden_subprocess_kwargs(),
            )
            self.assertEqual(completed.returncode, 0, completed.stderr)
            self.assertEqual(target.read_text(), "executed+finally")

    def test_terminating_error_is_not_reported_as_success(self):
        completed = subprocess.run(
            viewer._powershell_script_command(shutil.which("powershell.exe"),
                "$ErrorActionPreference = 'Stop'\ntry { throw 'probe failure' } finally { }"),
            stdin=subprocess.DEVNULL, capture_output=True, timeout=15,
            **viewer._hidden_subprocess_kwargs(),
        )
        self.assertNotEqual(completed.returncode, 0)

    def test_all_three_excel_paths_execute_complete_script(self):
        # Substitute only the Office operation, keeping the real subprocess
        # transport and parser used by PDF, XLSX and picture export.
        real_run = subprocess.run
        calls = []
        def probe(args, **kwargs):
            output = (kwargs['env'].get('XVIEWER_XLS_PDF')
                      or kwargs['env'].get('XVIEWER_XLS_TARGET')
                      or kwargs['env']['XVIEWER_XLS_MANIFEST'])
            kwargs['env']['XVIEWER_TEST_OUTPUT'] = output
            script = """try {
    [IO.File]::WriteAllText($env:XVIEWER_TEST_OUTPUT, ('x' * 200))
}
finally {
    [Console]::WriteLine('finished')
}
"""
            if args[-1] == '-':
                kwargs['input'] = script
            else:
                args = args[:-1] + [script]
            result = real_run(args, **kwargs)
            calls.append(result)
            return result
        with TemporaryDirectory() as td, patch.object(viewer.subprocess, 'run', side_effect=probe):
            root = Path(td)
            source = root / 'source.xls'
            self.assertTrue(viewer._render_xls_with_excel_pdf(source, root / 'preview.pdf'))
            self.assertTrue(viewer._convert_xls_with_excel_com(source, root / 'preview.xlsx'))
            self.assertTrue(viewer._extract_xls_images_with_excel_com(source, root / 'images.json', root / 'images'))
        self.assertEqual(len(calls), 3)
        self.assertTrue(all('finished' in result.stdout for result in calls))


if __name__ == '__main__':
    unittest.main()

# Changelog

## 2.6.7 — Windows icon cache fix and xViewer release hardening

- Fix the Desktop/Start Menu shortcut still showing the old xExcel-era icon after an update. The installer now copies the icon to a versioned path such as `%LOCALAPPDATA%\xViewer\icons\xviewer-2.6.7.ico` instead of overwriting the same cached filename.
- Delete and recreate `xViewer.lnk` during installation, then notify Windows Shell (`SHChangeNotify` and `ie4uinit -show`) so Explorer refreshes icon metadata without requiring a restart.
- Give the running process the explicit Windows AppUserModelID `jtl-sun.xViewer`, improving taskbar grouping/icon identity when launched through `pythonw.exe`.
- Keep the 2.6.6 centered startup, LEFT-list double-click external open, and the GPT-6 Astra-assisted legacy `.xls` execution repair unchanged.

## 2.6.6 — Centered startup, double-click external open, refreshed xViewer icon

- Center the main xViewer window on the primary display at startup while preserving the existing 1580×920 default size and minimum size.
- LEFT file-list double-click now mirrors Enter: folders navigate and files open with the Windows-associated application. Single-click/arrow selection continues to drive the integrated RIGHT viewer.
- Resolve the exact Treeview row under the double-click pointer before opening so the intended file opens reliably.
- Replace the old xExcel-era icon with a new xViewer icon representing Excel, PDF, and image viewing, and update the installer/shortcuts to use `xviewer.ico`.
- Preserve the 2.6.5 legacy `.xls` repair: PowerShell scripts are passed as one `-Command` argument in STA mode, avoiding the multiline `try/finally` stdin execution bug that prevented legacy image rendering helpers from actually running.
- Add regression coverage for centered geometry and double-click behavior.

## 2.6.5 — Legacy XLS helper execution fix

- Run complete PowerShell scripts as a single command instead of interactive stdin. The old transport could exit successfully without executing multiline try/finally blocks, silently disabling XLS PDF, XLSX and picture export.
- Use STA for Excel clipboard operations and invalidate the previous preview cache.
- Keep source workbooks read-only and retain the existing Excel-rendered PDF preview strategy.
- Add three Windows subprocess regression tests; isolate the converted-preview test from real Office automation.
- Actual Excel rendering still requires Microsoft Excel and a working interactive Windows user session. No native xlrd image support is added.

## 2.6.4

- Added Microsoft Excel native PDF `EXACT VIEW` for legacy `.xls` workbooks.
- `.xls` now prefers Excel-rendered fixed-format output on Windows, preserving old drawing-layer pictures that can be invisible to xlrd/openpyxl and Shape-by-Shape export.
- The original `.xls` remains read-only and is never saved or modified.
- Exact-view PDF is cached by source path/size/mtime/version; 2.6.3 legacy caches are automatically bypassed.
- Existing converted-grid and xlrd fallbacks remain available when Excel automation/PDF export is unavailable.
- Regression suite increased to 33 tests.


## 2.6.3

- Fixed legacy `.xls` workbooks whose product pictures still did not appear after the 2.6.2 `.xlsx` preview conversion.
- Added direct Microsoft Excel COM extraction of legacy picture-like Shapes to PNG sidecar files.
- Preserves the Shape's worksheet, top-left cell, pixel offset, width, and height so exported pictures can be overlaid in the RIGHT workbook view.
- Merges COM-exported pictures with openpyxl images while suppressing near-duplicate pictures.
- Direct picture export also works with the xlrd cell fallback when conversion is unavailable but Excel COM can still read the workbook.
- Bumped the legacy `.xls` cache format so image-less 2.6.2 caches are not reused.
- Added regression coverage for legacy image manifest loading and image de-duplication.

## 2.6.2

- Fixed missing embedded images in legacy `.xls` workbooks.
- Added read-only `.xls` preview conversion through installed Microsoft Excel (preferred) or LibreOffice (fallback).
- Original `.xls` files are never modified; xViewer converts to a cached temporary `.xlsx` preview for display only.
- Cache invalidates automatically when the source path, size, or modified time changes.
- If neither conversion engine is available, xViewer falls back to the previous xlrd cell-only legacy viewer.
- Added a regression test confirming converted `.xls` previews expose embedded images while remaining read-only.

## 2.6.1
- Renamed the visible product branding to **xViewer**.
- Added an Image Viewer **`1:1 / Fit`** button to the right of Zoom.
- Images now open in hybrid 1:1/Fit mode: no enlargement above 100%, but large images shrink to the available viewer.
- Fit recalculates when the viewer is resized.
- Numeric zoom exits Fit mode; the button returns to Fit or exact 1:1.
- Added regression coverage for fit zoom calculation.

## 2.6.0

- Added integrated PDF viewing in the RIGHT pane using pypdfium2.
- Added integrated image viewing for PNG/JPG/JPEG/JFIF/WEBP/BMP/GIF/TIF/TIFF/ICO.
- Added mutually exclusive `Excel only`, `PDF only`, and `Images only` LEFT filters; turning the active filter off shows all file types.
- Added PDF first/previous/next/last page navigation and page count.
- Added PDF/image zoom, full scroll, Clipboard image copy, Save Image As, and external Open buttons.
- Added safe switching among Excel, PDF, image, folder, and unsupported-file selections without stale RIGHT content.
- Added asynchronous PDF/image rendering with generation invalidation for rapid LEFT browsing.
- Extended Open With registration to supported PDF/image formats without changing Windows default associations.
- Added `pypdfium2>=4.30` dependency and media-render regression tests.
- Regression suite expanded from 25 to 28 tests.

## 2.5.13

- Disabled `Tab` and `Shift+Tab` while the LEFT FILES Treeview has keyboard focus.
- Prevents default Tk focus traversal from jumping to toolbar buttons with no obvious visual destination.
- Pane switching remains explicit via `Ctrl+Left` / `Ctrl+Right` and `Alt+1` / `Alt+2`.
- RIGHT WORKBOOK keeps normal Excel-style `Tab` / `Shift+Tab` cell navigation.
- Added regression coverage for the LEFT file-list Tab policy.

## 2.5.12

- Fixed stale RIGHT-pane workbook content when the LEFT selection moves from an Excel workbook to a folder or non-Excel file.
- Cancels pending/deferred workbook loads and invalidates in-flight background loaders before clearing the viewer.
- Restores the blank Workbook placeholder after clearing so no previous Excel image remains visible.
- Protects unsaved workbooks: Save / Discard / Cancel is respected before clearing the currently edited workbook.
- Added regression coverage for selection-to-viewer action decisions.

## 2.5.11

- LEFT Search box: `Up` / `Down` now immediately applies any pending filter, exits the search field, focuses the filtered file list, and continues selection navigation.
- `Down` enters the filtered list at the first result; `Up` enters it at the last result, avoiding stale pre-search selections.
- `Enter` now also flushes the current debounced search text before selecting the first result.
- Added regression coverage for search-to-file arrow navigation.

## 2.5.10

- Final GitHub-readiness audit and packaging cleanup.
- Unified app/viewer version reporting with the package `__version__`.
- Removed stale hard-coded installer version labels.
- Installer recreates incompatible pre-existing private venvs.
- Uninstaller removes stale Excel `OpenWithProgids` registrations.
- Added Windows GitHub Actions CI for Python 3.11 / 3.12 / 3.13.
- Added clean release ZIP + SHA256 packaging tool and a version-consistency regression test.

## 2.5.9

- Fixed LEFT-pane RIGHT-DRAG selection when the gesture starts on an already-selected file.
- The starting row remains selected and every crossed row is added to the selection.
- Existing selections are not accidentally toggled off.
- Fast-drag gap filling and edge auto-scroll remain enabled.

## 2.5.8

- Added mDIR-style right mouse drag selection in the LEFT file list.
- Fast drags include intermediate rows and edge dragging auto-scrolls.
- Workbook auto-loading pauses during the gesture and resumes after release.

## 2.5.7

- Added standard Ctrl-click / Shift-click multi-selection in the LEFT file list.
- Del moves selected files/folders to the Windows Recycle Bin after confirmation.
- Unsaved open workbooks are protected before deletion.

## 2.5.6

- Enter on a LEFT-pane file opens it with the Windows-associated application.
- Enter on a folder navigates into that folder.

## 2.5.5

- Added persistent recent-directory history with a compact dropdown beside the path bar.

## 2.5.4

- Fixed Ctrl+F filename search so ordinary number keys are not misread as Alt+1 / Alt+2 pane shortcuts on Windows.

## 2.5.3

- Added live LEFT-pane filename filtering, Ctrl+F focus, Esc clear, and cached directory filtering.

## 2.5.2

- Split the RIGHT WORKBOOK header into a button row and a separate filename/status row.

## 2.5.1

- Extended RIGHT WORKBOOK scroll extents to the full bottom/right edge of embedded images.

## 2.5.0

- Added Bright / Dark / System themes with persistent preference and Windows system-theme following.

## 2.4.1

- Softened the LEFT/RIGHT pane divider while preserving resize behavior.

## 2.4.0

- Added the mDIR-compatible Link Manager and full Folder / File / Program / Web / Action / Command shortcut model.

## 2.3.1

- Replaced breadcrumb buttons with the continuous green mDIR-style clickable path bar.

## 2.3.0

- Added mDIR folder shortcut import and editable quick links.

## 2.2.0

- Improved Excel layout fidelity: hidden rows/columns, row heights, column widths, gridline state, fonts, fills, borders, alignment, and image anchors.

## 2.1.4

- Hardened Ctrl+Left / Ctrl+Right and Alt+1 / Alt+2 pane switching on Windows/Tk.

## 2.1.3

- Made Ctrl+Left / Ctrl+Right the primary pane-navigation shortcuts while retaining Alt+1 / Alt+2 and F6.

## 2.1.2

- Added dedicated pane-navigation shortcuts and returned Tab / Shift+Tab to cell navigation.

## 2.1.1

- Fixed LEFT pane collapse at startup and stabilized two-pane geometry.

## 2.1.0

- Added mouse/keyboard pane focus, LEFT file arrow navigation, RIGHT cell arrow navigation, and full worksheet tab display.

## 2.0.0

- Introduced the integrated LEFT file browser + RIGHT full-workbook viewing/editing workspace.

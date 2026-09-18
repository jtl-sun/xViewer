# xViewer 2.6.7 Release Notes

## Shortcut/icon refresh fix

2.6.7 addresses the case where Windows continued to display the previous xExcel-style Desktop icon even though the packaged `xviewer.ico` had already been replaced. Windows Explorer caches shortcut icons heavily when the icon file path stays unchanged. The installer now installs each release icon under a versioned filename, removes/recreates the Desktop and Start Menu shortcuts, sends a Shell association-change notification, and asks `ie4uinit` to refresh visible icons. The running app also sets an explicit `jtl-sun.xViewer` AppUserModelID so the taskbar identity follows xViewer instead of the generic Python host.


## Startup and file-list usability

The main xViewer window now opens centered on the primary display. The default client size remains 1580×920 and the existing minimum size is unchanged.

LEFT file-list double-click now behaves like Enter: double-click a folder to enter it, or double-click a file to open it with the Windows-associated application. This is intentionally separate from the integrated RIGHT viewer, which continues to update from normal single-click or keyboard selection. The row under the mouse pointer is resolved directly before opening to avoid stale Treeview selection timing.

## New xViewer icon

The application icon has been refreshed for the broader xViewer identity. The new icon visually represents the three supported viewer families—Excel, PDF, and images—and is supplied as `xviewer-icon.png` plus a multi-resolution Windows `xviewer.ico`. The Windows installer, Desktop shortcut, Start Menu shortcut, Open With registration, and Tk window now use the new asset.

## Legacy XLS repair preserved

The GPT-6 Astra-assisted 2.6.5 repair was reviewed and retained. Its important fix is transport-level: Excel helper PowerShell scripts are no longer piped through interactive `powershell.exe -Command -`, which could return success without executing a complete multiline `try/finally` block. xViewer now passes the entire script as one command argument, uses STA for Excel clipboard automation, keeps source paths in environment variables, and invalidates older legacy caches. This allows the existing Excel-rendered PDF/XLSX/picture helper paths to actually execute.

## Validation

- Existing core regression suite plus new startup-centering and double-click tests.
- Python compile check.
- Application self-check.
- Clean release ZIP and SHA256 generation.

## Install

Close xViewer, extract the ZIP, and run `install_xViewer.bat`. The installer updates the application icon and shortcuts automatically.

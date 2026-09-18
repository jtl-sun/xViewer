$ErrorActionPreference = "Stop"

function Find-Python {
    $Py = Get-Command py -ErrorAction SilentlyContinue
    if ($Py) {
        $Resolved = & $Py.Source -3 -c "import sys; print(sys.executable)"
        if ($LASTEXITCODE -eq 0 -and $Resolved) { return $Resolved.Trim() }
    }
    $Python = Get-Command python -ErrorAction SilentlyContinue
    if ($Python) { return $Python.Source }
    throw "Python 3.11 or newer was not found. Install Python for Windows, then run install_xViewer.bat again."
}

function Notify-ShellIconRefresh {
    try {
        if (-not ("XViewerShellNotify" -as [type])) {
            Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class XViewerShellNotify {
    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
"@
        }
        # SHCNE_ASSOCCHANGED = 0x08000000, SHCNF_IDLIST = 0
        [XViewerShellNotify]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    } catch {
        Write-Host "Shell icon refresh notification was skipped: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }

    try {
        $Ie4uinit = Join-Path $env:SystemRoot "System32\ie4uinit.exe"
        if (Test-Path -LiteralPath $Ie4uinit) {
            Start-Process -FilePath $Ie4uinit -ArgumentList "-show" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
        }
    } catch { }
}

$Python = Find-Python
& $Python -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)"
if ($LASTEXITCODE -ne 0) { throw "xViewer requires Python 3.11 or newer." }

$InstallRoot = Join-Path $env:LOCALAPPDATA "xViewer"
$VenvRoot = Join-Path $InstallRoot "venv"
$VenvPython = Join-Path $VenvRoot "Scripts\python.exe"
$VenvPythonw = Join-Path $VenvRoot "Scripts\pythonw.exe"
$IconSource = Join-Path $PSScriptRoot "mdir\assets\xviewer.ico"
$IconRoot = Join-Path $InstallRoot "icons"
$BinRoot = Join-Path $InstallRoot "bin"

New-Item -ItemType Directory -Path $InstallRoot -Force | Out-Null
New-Item -ItemType Directory -Path $IconRoot -Force | Out-Null
New-Item -ItemType Directory -Path $BinRoot -Force | Out-Null

$NeedVenv = $true
if (Test-Path -LiteralPath $VenvPython) {
    try {
        & $VenvPython -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)"
        if ($LASTEXITCODE -eq 0) { $NeedVenv = $false }
    } catch {
        $NeedVenv = $true
    }
}

if ($NeedVenv) {
    if (Test-Path -LiteralPath $VenvRoot) {
        Write-Host "[1/4] Recreating incompatible private Python environment..." -ForegroundColor Yellow
        Remove-Item -LiteralPath $VenvRoot -Recurse -Force
    } else {
        Write-Host "[1/4] Creating private Python environment..." -ForegroundColor Green
    }
    & $Python -m venv $VenvRoot
    if ($LASTEXITCODE -ne 0) { throw "Could not create the private Python environment." }
} else {
    Write-Host "[1/4] Existing private Python environment is compatible." -ForegroundColor Green
}

Write-Host "[2/4] Installing xViewer..." -ForegroundColor Green
& $VenvPython -m pip install --disable-pip-version-check --prefer-binary --upgrade --force-reinstall $PSScriptRoot
if ($LASTEXITCODE -ne 0) { throw "Could not install xViewer." }

$Version = (& $VenvPython -P -c "import mdir; print(mdir.__version__)").Trim()
if (-not $Version) { throw "Could not determine installed xViewer version." }

# Windows aggressively caches shortcut icons by file path.  Never overwrite the
# same installed icon path between releases.  A versioned path guarantees that
# a new xViewer icon is picked up immediately after an update.
$InstalledIcon = Join-Path $IconRoot ("xviewer-" + $Version + ".ico")
if (Test-Path -LiteralPath $IconSource) {
    Copy-Item -LiteralPath $IconSource -Destination $InstalledIcon -Force
}
Remove-Item -LiteralPath (Join-Path $InstallRoot "xviewer.ico") -Force -ErrorAction SilentlyContinue

$Launcher = '@echo off' + "`r`n" + '"' + $VenvPythonw + '" -P -m mdir %*' + "`r`n"
Set-Content -LiteralPath (Join-Path $BinRoot "xviewer.cmd") -Value $Launcher -Encoding Ascii

Write-Host "[3/4] Creating Desktop and Start Menu shortcuts..." -ForegroundColor Green
$Shell = New-Object -ComObject WScript.Shell
$Desktop = [Environment]::GetFolderPath("Desktop")
$StartMenu = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
foreach ($ShortcutPath in @((Join-Path $Desktop "xViewer.lnk"),(Join-Path $StartMenu "xViewer.lnk"))) {
    # Delete the old .lnk first so Explorer cannot retain its previous icon metadata.
    Remove-Item -LiteralPath $ShortcutPath -Force -ErrorAction SilentlyContinue
    $Shortcut = $Shell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $VenvPythonw
    $Shortcut.Arguments = "-P -m mdir"
    $Shortcut.WorkingDirectory = [Environment]::GetFolderPath("UserProfile")
    if (Test-Path -LiteralPath $InstalledIcon) { $Shortcut.IconLocation = "$InstalledIcon,0" }
    $Shortcut.Description = "xViewer - browse Excel workbooks, PDFs, and images in one viewer"
    $Shortcut.Save()
}

Write-Host "[4/4] Registering Open with xViewer..." -ForegroundColor Green
$ProgId = "xViewer.Workbook"
$ProgRoot = "HKCU:\Software\Classes\$ProgId"
New-Item -Path $ProgRoot -Force | Out-Null
Set-Item -Path $ProgRoot -Value "Excel / PDF / Image - xViewer"
New-Item -Path "$ProgRoot\DefaultIcon" -Force | Out-Null
Set-Item -Path "$ProgRoot\DefaultIcon" -Value "`"$InstalledIcon`",0"
New-Item -Path "$ProgRoot\shell\open\command" -Force | Out-Null
Set-Item -Path "$ProgRoot\shell\open\command" -Value "`"$VenvPythonw`" -P -m mdir `"%1`""
foreach ($Ext in @(".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".pdf", ".png", ".jpg", ".jpeg", ".jfif", ".webp", ".bmp", ".gif", ".tif", ".tiff", ".ico")) {
    $OpenWith = "HKCU:\Software\Classes\$Ext\OpenWithProgids"
    New-Item -Path $OpenWith -Force | Out-Null
    New-ItemProperty -Path $OpenWith -Name $ProgId -Value "" -PropertyType String -Force | Out-Null
}

# Force Explorer/Start Menu to discard cached icon association data.  The
# versioned icon filename above is the primary cache-busting mechanism; these
# notifications make the change visible without requiring a sign-out/restart.
Notify-ShellIconRefresh

$Check = & $VenvPython -P -m mdir --check
if ($LASTEXITCODE -ne 0) { throw "Installation validation failed." }
Write-Host ""
Write-Host $Check -ForegroundColor Green
Write-Host "Installed successfully: $InstallRoot" -ForegroundColor Green
Write-Host "Shortcut icon: $InstalledIcon" -ForegroundColor Green

# Best-effort cleanup of the pre-rename xExcel Viewer installation. User
# settings live in the profile and remain available for compatibility.
$LegacyInstallRoot = Join-Path $env:LOCALAPPDATA "xExcel Viewer"
$LegacyDesktop = Join-Path ([Environment]::GetFolderPath("Desktop")) "xExcel Viewer.lnk"
$LegacyStart = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\xExcel Viewer.lnk"
Remove-Item -LiteralPath $LegacyDesktop -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $LegacyStart -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath "HKCU:\Software\Classes\xExcelViewer.Workbook" -Recurse -Force -ErrorAction SilentlyContinue
foreach ($Ext in @(".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".pdf", ".png", ".jpg", ".jpeg", ".jfif", ".webp", ".bmp", ".gif", ".tif", ".tiff", ".ico")) {
    Remove-ItemProperty -LiteralPath ("HKCU:\Software\Classes\" + $Ext + "\OpenWithProgids") -Name "xExcelViewer.Workbook" -ErrorAction SilentlyContinue
}
if ((Test-Path -LiteralPath $LegacyInstallRoot) -and ($LegacyInstallRoot -ne $InstallRoot)) {
    Remove-Item -LiteralPath $LegacyInstallRoot -Recurse -Force -ErrorAction SilentlyContinue
}

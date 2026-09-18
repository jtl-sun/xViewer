# xViewer 2.6.7


**2.6.7:** Windows가 이전 바로가기 아이콘을 캐시해서 새 xViewer 아이콘이 적용되지 않는 문제를 수정했습니다. 설치 시 아이콘을 `xviewer-<version>.ico`처럼 버전별 새 경로에 복사하고, 기존 `xViewer.lnk`를 삭제 후 다시 생성하며 Windows Shell 아이콘 캐시 갱신을 요청합니다. 실행 프로세스에도 `jtl-sun.xViewer` AppUserModelID를 지정해 작업표시줄 아이콘/그룹도 xViewer 정체성을 사용하도록 보강했습니다.


GitHub의 새 `xViewer` 저장소에서는 `main`에 새 버전을 커밋하면 Windows CI가 테스트/빌드한 뒤 해당 버전 태그와 GitHub Release를 자동 생성하도록 release workflow를 포함합니다.
**2.6.5:** Legacy `.xls` 표시 방식을 한 단계 더 보강했습니다. 일부 오래된 BIFF Excel 파일은 `.xlsx` 변환이나 Shape별 PNG 추출에서도 제품 이미지가 사라질 수 있습니다. 2.6.5는 Windows에서 Microsoft Excel이 설치되어 있으면 원본 `.xls`를 **read-only로 Excel 자체 PDF 렌더링**하여 RIGHT 패널에 `EXACT VIEW`로 우선 표시합니다. 이 방식은 Excel의 실제 drawing/printing engine을 사용하므로 오래된 picture/metafile/OLE drawing도 시각적으로 보존될 가능성이 가장 높습니다. Excel 자동화가 불가능하면 기존 cell/grid fallback을 그대로 사용합니다. 원본 `.xls`는 수정하지 않습니다.

**2.6.3:** Legacy `.xls` 이미지 표시를 다시 보완했습니다. 2.6.2의 `.xls → .xlsx` 변환만으로는 일부 오래된 BIFF/Escher 그림, WMF/EMF 계열 picture shape, OLE/linked picture가 openpyxl에 노출되지 않는 경우가 있었습니다. 2.6.3은 Microsoft Excel COM을 사용해 원본 `.xls`의 picture-like `Shape`를 직접 PNG로 렌더링하고, 원래 worksheet/cell 위치와 크기를 sidecar cache에 저장한 뒤 RIGHT workbook 위에 합성합니다. 기존 converted `.xlsx` preview의 이미지가 보이는 경우에는 중복을 피하고, converted preview에서 빠진 그림만 추가합니다. 2.6.2의 legacy cache key도 변경하여 이전에 생성된 이미지 없는 cache를 자동으로 재사용하지 않습니다. 원본 `.xls`는 계속 read-only입니다.

**2.6.2:** Legacy Excel `.xls` 이미지 표시를 보완했습니다. `.xls`는 `xlrd`만으로는 BIFF/Escher drawing layer의 embedded image를 읽을 수 없기 때문에, Windows에서 Microsoft Excel이 설치되어 있으면 원본을 수정하지 않고 백그라운드에서 임시 `.xlsx` preview로 변환하여 이미지와 레이아웃을 표시합니다. Excel이 없으면 LibreOffice를 자동 fallback으로 시도하며, 둘 다 없을 때만 기존 cell-only `.xls` view로 돌아갑니다. 변환 preview는 source path/size/modified time 기준으로 cache되어 같은 파일을 다시 볼 때 재변환을 줄입니다. 원본 `.xls`는 계속 read-only입니다.


**2.6.1:** 제품 이름을 `xViewer`로 정리하고, Image Viewer에 `1:1 / Fit` 버튼을 추가했습니다. 이미지 파일은 기본적으로 화면 안에 맞추되 작은 이미지는 확대하지 않는 **1:1/Fit 모드**로 열립니다. `1:1 / Fit` 버튼으로 원본 100%와 화면 맞춤을 즉시 전환할 수 있습니다.

**2.6.0:** PDF와 일반 이미지 파일을 xViewer 안에서 직접 볼 수 있도록 확장했습니다. LEFT에는 `Excel only`, `PDF only`, `Images only` 필터가 추가되며, 선택한 PDF는 페이지 단위로, PNG/JPG/WEBP/BMP/GIF/TIFF/ICO 등 이미지는 스크롤 가능한 RIGHT Viewer에서 표시됩니다. PDF 페이지와 이미지도 Zoom, Windows Clipboard 복사, 외부 기본 프로그램 열기를 지원합니다.

**2.5.13:** LEFT FILES Tab-key focus fix. When the file list owns focus, `Tab` / `Shift+Tab` are now intentionally ignored so focus does not jump unpredictably through toolbar controls. Use `Ctrl+Right` or `Alt+2` to move to WORKBOOK; Tab remains Excel-style cell navigation only inside WORKBOOK.

**2.5.12:** RIGHT-pane stale-workbook cleanup. Selecting a folder or a non-Excel file on the LEFT now clears any previously displayed workbook immediately, cancels pending workbook loads, and returns the RIGHT pane to its blank placeholder so an old Excel image cannot remain on screen.

**2.5.11:** Search-to-file keyboard navigation update. After typing in the LEFT Search box, pressing **Up** or **Down** immediately applies the current filter, leaves the search box, focuses the filtered FILES list, and moves file selection so browsing can continue without reaching for the mouse.

이 버전은 **왼쪽의 수많은 Excel/PDF/Image 파일을 빠르게 고르고, 오른쪽에서 내용을 확인하며 필요한 자료를 복사·정리**하는 업무 흐름에 집중합니다. Excel은 기존처럼 전체 Workbook 편집/저장을 지원하고, PDF와 일반 이미지는 읽기 전용 Viewer로 빠르게 확인합니다.

## 핵심 변화

- 기존 mDIR-P 2.26.15의 안정판 구조와 Excel Virtual Grid 코드를 기준으로 다시 단순화했습니다.
- 예전처럼 오른쪽에 `first 35 rows x 9 columns` Preview를 표시하지 않습니다.
- Excel 파일을 왼쪽에서 선택하면 **같은 창 오른쪽**에 전체 Workbook이 열립니다.
- 모든 Sheet 탭을 사용할 수 있습니다.
- 사용 영역 전체를 가상 그리드(Virtual Grid)로 스크롤합니다. 화면에 보이는 부분을 중심으로 그리므로 큰 파일도 전체 셀 UI를 한꺼번에 만들지 않습니다.
- `.xlsx/.xlsm/.xltx/.xltm`은 셀 편집과 저장을 지원합니다.
- `.xls`는 안전을 위해 읽기 전용입니다. Microsoft Excel(우선) 또는 LibreOffice가 있으면 임시 `.xlsx` preview로 변환하여 embedded image까지 표시하고, 둘 다 없으면 cell-only fallback으로 엽니다.
- Embedded image를 표시하고 선택한 이미지를 `Ctrl+C`로 Windows Clipboard에 복사할 수 있습니다.
- Legacy `.xls` preview cache: `%LOCALAPPDATA%\xViewer\cache\legacy-xls` (원본 path/size/modified time이 바뀌면 자동으로 새 cache를 사용).

## 화면 구성

**왼쪽:** 파일 브라우저
- `Ctrl+클릭` / `Shift+클릭` 멀티 선택
- **오른쪽 마우스 버튼을 누른 채 위/아래로 드래그**하면 시작 파일을 유지하면서 지나간 파일을 선택에 추가 (mDIR 방식)
- 빠르게 드래그해도 중간 행을 빠뜨리지 않고, 위/아래 가장자리에서는 자동 스크롤
- `Del`로 선택 항목을 Windows 휴지통으로 이동
- 폴더 이동 / Back / Up / Home / Drive 버튼
- Name / Ext / Size / Modified
- 이름 Filter
- `Excel only` / `PDF only` / `Images only` 필터
- 세 필터 중 하나를 켜면 해당 형식만 표시하며, 활성 필터를 다시 끄면 모든 파일 형식을 표시
- Excel 파일을 한 번 선택하면 약 0.28초 뒤 오른쪽에서 Workbook 로드

**오른쪽:** Excel / PDF / Image Viewer
- Workbook의 모든 Sheet
- 전체 행/열 스크롤
- 셀/범위 선택 후 `Ctrl+C`
- `F2` 또는 셀 더블클릭으로 직접 편집
- 상단 Formula Bar에서도 값 편집
- `Ctrl+V`로 탭/행 구조를 유지한 범위 Paste
- `Delete`로 선택 범위 지우기
- `Ctrl+S` 저장
- Save As
- Embedded image 클릭 후 `Ctrl+C`
- 이미지 우클릭 → Copy Image / Save Image As
- Zoom 50%~200%
- Open in Excel 버튼
- PDF: 페이지 이동(`|◀`, `◀`, `▶`, `▶|`), 전체 페이지 스크롤, Zoom 50%~200%, 현재 페이지를 이미지로 Clipboard 복사
- Image: PNG/JPG/JPEG/JFIF/WEBP/BMP/GIF/TIF/TIFF/ICO 표시, 스크롤/Zoom, 원본 이미지를 Clipboard 복사, Save Image As
- PDF/Image 선택 시 Excel 전용 Formula/Sheet UI는 숨겨지고 Media Viewer로 자동 전환

## 화면 테마

상단 오른쪽 `Theme:`에서 세 가지 모드를 선택할 수 있습니다.

- `Bright`: 기존 밝은 테마
- `Dark`: xExcel의 파일 목록, 툴바, 입력창, 탭, Link Manager, 상태바, 경로바와 Excel row/column header를 어둡게 표시
- `System`: Windows의 **앱 모드**를 따라 Bright/Dark를 자동 전환

선택한 테마는 `%USERPROFILE%\.xexcel-viewer.json`에 저장됩니다. `System`을 선택하면 프로그램 실행 중 Windows 테마 변경도 주기적으로 확인하여 반영합니다.

중요: **Excel Workbook의 실제 셀 색상과 서식은 원본을 우선합니다.** Dark 테마라고 해서 흰색 Excel 셀을 검게 뒤집지 않습니다. 이것은 오래된 buyer/spec sheet를 원본 MS Excel과 최대한 비슷하게 보여주기 위한 의도입니다.

## 저장 안전장치

원본에 처음 `Ctrl+S`로 덮어쓸 때 같은 폴더의 `xExcel_Backup` 폴더에 타임스탬프 백업을 한 번 자동 생성합니다.

예:

```text
D:\Buyer\010620-14 TJ.xlsx
D:\Buyer\xExcel_Backup\010620-14 TJ-20260917-153000.xlsx
```

오래된 `.xls` 파일은 xViewer에서 직접 저장하지 않습니다. 필요하면 `Open in Excel`로 편집하거나 새 `.xlsx`로 정리하는 것을 권장합니다.

## 설치

1. ZIP 전체 압축 해제
2. `install_xViewer.bat` 더블클릭
3. 바탕화면 `xViewer` 실행

기본 설치 위치:

```text
%LOCALAPPDATA%\xViewer
```

Python 3.11 이상이 필요합니다. 설치 후 프로그램은 `pythonw.exe`로 실행되므로 별도 Terminal 창이 뜨지 않습니다.

## 단축키

- `F2` / Double Click: 셀 편집
- `Ctrl+C`: 셀 범위 또는 선택된 embedded image 복사
- `Ctrl+V`: 셀 범위 붙여넣기
- `Delete`: 선택 셀 지우기
- `Ctrl+S`: 저장
- `Ctrl++ / Ctrl+- / Ctrl+0`: Zoom
- 왼쪽 파일 목록 `F5`: 새로고침

## 기준

- Base: mDIR-P 2.26.15 안정판
- xViewer: 2.6.5
- 목표: 대량의 **Excel buyer/work-sheet + PDF + reference image**를 빠르게 열어 내용과 이미지를 확인하고 다른 자료로 정리하는 전용 작업 프로그램


## 2.1 keyboard workflow

- **Mouse click**: click the left file list or the right workbook to activate that pane.
- **Ctrl+Left**: go directly to the left FILES pane.
- **Ctrl+Right**: go directly to the right WORKBOOK pane.
- **Alt+1**: secondary shortcut for the left FILES pane.
- **Alt+2**: secondary shortcut for the right WORKBOOK pane.
- **F6**: legacy pane-toggle shortcut kept for compatibility.
- **Tab / Shift+Tab (LEFT FILES)**: disabled; focus stays in the file list. Use `Ctrl+Right` or `Alt+2` to move to WORKBOOK.
- **Tab / Shift+Tab (WORKBOOK)**: move to the next / previous Excel cell; Tab is not used to switch panes.
- **Left FILES pane**: Up/Down selects the previous/next file. Home/End and PageUp/PageDown are supported.
- **Right WORKBOOK pane**: arrow keys move the active cell. Shift+arrow extends the selection.
- **Sheets**: every worksheet is listed in the horizontal Sheets bar. Click any sheet name, use the arrow buttons, or press Ctrl+PageUp/Ctrl+PageDown.

Workbook loading is deliberately prevented from stealing focus from the left pane, so you can keep moving through many files with the keyboard while their contents load on the right.


## 2.2.0 Excel-layout fidelity update

오른쪽 WORKBOOK 패널이 원본 Excel과 다르게 뒤엉켜 보이던 가장 큰 원인은 **숨김 행/열을 화면에서 실제로 숨기지 않고 모두 폭/높이를 주어 그렸던 것**입니다. 이 버전은 Excel의 저장된 레이아웃 정보를 더 충실히 반영합니다.

- 숨김 Row / Column은 0px로 처리하고 행/열 Header에서도 건너뜁니다.
- Column width / Row height / 기본 크기를 Excel 값에 맞춰 계산합니다.
- `Show Gridlines` 설정을 반영합니다. 원본 Excel에서 Gridline이 꺼져 있으면 xExcel에서도 빈 셀 격자를 강제로 그리지 않습니다.
- Cell border의 thin / medium / thick / double을 표시합니다.
- Font name / size / bold / italic / underline, Fill color, horizontal/vertical alignment, indent, wrap을 반영합니다.
- Excel의 기본 세로 정렬(bottom)과 숫자의 기본 오른쪽 정렬을 반영합니다.
- Theme / indexed color를 가능한 범위에서 해석합니다.
- Embedded image의 anchor offset 및 one-cell/two-cell anchor 크기를 반영합니다.
- 방향키와 Tab 이동 시 숨김 행/열을 건너뜁니다.
- 일반 텍스트는 빈 인접 셀 쪽으로 자연스럽게 보이도록 cell fill과 text를 분리 렌더링합니다.

이 렌더러는 Excel 자체를 내장한 것은 아니므로 모든 Office 기능을 100% pixel-identical하게 재현하지는 않습니다. 하지만 오래된 buyer/spec worksheet에서 흔한 **숨김 행·열, 병합 셀, 이미지, 테이블 border, 사용자 지정 폭/높이** 때문에 레이아웃이 무너지는 문제를 우선 해결하는 버전입니다.

## mDIR-compatible Link Manager (2.4.0)

The top shortcut bar and **Edit Links** window now use the same link model as mDIR. Links are shown in their saved order and can be **Folder, File, Program, Web, Action, or Command**. Each link also stores a target pane (`active`, `left/files`, or `right/workbook`) and optional JSON arguments.

The Link Manager uses the mDIR layout: a table with **Name / Type / URL-Path-Action / Pane**, then editable **Name, Type/Pane, Target, Arguments** fields, followed by **Add, Remove, Move Up, Move Down, Browse File, Browse Folder, Save, Cancel**. An **Import mDIR** button copies the complete mDIR link list, including non-folder items such as GitHub, program launchers, and `powershell_here`.

Links are stored in `%USERPROFILE%\.xexcel-viewer-links.json` using the mDIR-compatible JSON schema. If xExcel 2.3.x contains only the old folder-only list and it matches the folder subset of `%USERPROFILE%\.mdir-p-shortcuts.json`, 2.4.0 and later restore the complete mDIR link list on load.

Supported placeholders include `{home}`, `{current}`, `{selected}`, `{left}`, `{right}`, `{left_selected}`, `{right_selected}`, and `{project}`. The mDIR-style green path bar remains clickable by directory name, and `Ctrl+L` still opens direct path entry.

### Left-pane filename search
Use the small `▼` button at the right edge of the green path bar to reopen recently visited folders. xExcel keeps up to 30 unique directories in most-recent-first order in `%USERPROFILE%\.xexcel-viewer.json`; the list survives restarts. Choose any path to jump there, or select **Clear Folder History** to remove previous entries while keeping the current folder.

Use the `Search:` box above the file list to filter the current folder instantly. Matching is case-insensitive. Separate terms with spaces to require all terms, e.g. `010719 TJ`. `Ctrl+F` focuses the search box, `Enter` selects the first visible result, and `Esc` or the `×` button clears the filter. xExcel caches the current directory listing so typing does not rescan a large folder for every character.

## LEFT-pane multi-selection and Delete

- Click: select one item.
- Ctrl+click: add/remove individual items.
- Shift+click: select a continuous range.
- Del: move the complete selection to the Recycle Bin after confirmation.
- Enter: open a single selected file with its Windows-associated application; folders navigate into the folder.

Deletion is intentionally recycle-bin only; xExcel does not silently switch to permanent deletion when recycling fails.


## GitHub / 개발 검사

저장소에는 Windows GitHub Actions CI가 포함되어 있습니다. Push 또는 Pull Request 시 Python 3.11 / 3.12 / 3.13에서 다음을 자동 검사합니다.

```text
compileall
unittest
python -m mdir --check
python -m build
clean release ZIP packaging
```

로컬 최종 검사:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -e ".[dev]"
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -m mdir --check
.\.venv\Scripts\python.exe -m build
.\.venv\Scripts\python.exe tools\package_release.py
```

`tools/package_release.py`는 `__pycache__`, build/dist, virtual environment, 임시 파일을 제외한 설치용 ZIP과 `SHA256SUMS.txt`를 만듭니다.
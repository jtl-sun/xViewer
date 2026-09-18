# xViewer 2.7.4

xViewer는 Windows용 **Excel / PDF / Image / Text 통합 Viewer**입니다. LEFT 파일 리스트에서 많은 파일을 빠르게 훑어보고, RIGHT 패널에서 내용을 확인한 뒤 필요한 원본은 연결 프로그램이나 Microsoft Excel로 즉시 열 수 있습니다.

## 2.7.4 — 로딩 안정성과 통합 미리보기

- Excel 변환을 별도 프로세스에서 수행하고 작업 전체에 15초 제한을 적용합니다. 멈춘 작업을 취소한 뒤 다음 파일을 처리할 수 있습니다.
- xViewer 전용 Excel의 프로세스 핸들만 종료 대상으로 관리합니다. 실행 중인 모든 Excel을 종료하지 않습니다.
- 최신 선택을 우선 처리하고 이전 작업을 취소합니다. 인접 파일 자동 미리읽기는 중단했습니다.
- PDF·이미지 선택 오류, 백그라운드 오류 전달, 오래된 결과가 새 선택을 덮는 문제를 수정했습니다.
- 확장자가 `.xls`이지만 실제 내용은 `.xlsx`인 파일을 호환 보기에서 읽습니다.
- Excel·PDF·Images 필터를 함께 선택할 수 있습니다. 모두 해제하면 텍스트를 포함한 모든 파일을 표시합니다.
- 일반 PDF와 Excel PDF는 Fit으로 시작하며 1:1/Fit 전환을 지원합니다. 이미지 Fit은 최대 100%입니다.
- Excel PDF에만 화면 표시용 흰 여백 축소를 적용합니다. 원본 PDF·Excel은 수정하지 않습니다.
- TXT/MD/Markdown/CSV/JSON 읽기 전용 보기, 선택/전체 복사, 가로·세로 스크롤을 지원합니다.

미리보기 순서:

1. 원본 경로·크기·수정시간·캐시 버전이 같은 PDF가 있으면 재사용합니다.
2. 별도 작업 프로세스의 Microsoft Excel을 재사용하여 읽기 전용 PDF를 생성합니다(최대 15초).
3. PageSetup 단계 실패라면 원래 인쇄 설정으로 한 번 재시도합니다. 다른 실패는 설치된 LibreOffice로 변환합니다(추가 최대 15초).
4. 변환에 실패하면 openpyxl/xlrd 호환 보기로 전환합니다. 호환 보기도 읽기 전용이며 이미지나 배치가 다를 수 있습니다. 15초 안에 끝나지 않으면 오류를 표시합니다.

기존 PowerShell/legacy 그림 변환을 호환 보기에서 다시 호출하지 않습니다. 따라서 Excel 오류 뒤에 같은 변환을 여러 차례 장시간 반복하지 않습니다.

페이지 설정은 PrintArea 해제와 가로·세로 1페이지 맞춤만 적용합니다. 기존 방향, 용지, 여백은 유지합니다. 긴 시트는 한 페이지로 축소되므로 확대하거나 Excel에서 원본을 확인하세요.

## Excel 편집

2.7.4의 RIGHT Excel 화면은 정확성과 탐색 속도를 우선한 Preview입니다. 원본을 편집하려면 다음 중 하나를 사용합니다.

- LEFT에서 파일 선택 후 `Enter`
- LEFT에서 파일을 마우스 왼쪽 **더블클릭**
- RIGHT 상단 **Open in Excel**

PDF Preview 생성 엔진을 사용할 수 없는 환경에서는 이전 openpyxl/xlrd 기반 native workbook viewer가 compatibility fallback으로 남아 있습니다.

## 지원 형식

- Excel: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`
- PDF: `.pdf`
- Image: `.png`, `.jpg`, `.jpeg`, `.jfif`, `.webp`, `.bmp`, `.gif`, `.tif`, `.tiff`, `.ico`
- Text: `.txt`, `.md`, `.markdown`, `.csv`, `.json`

LEFT의 `Excel`, `PDF`, `Images`는 독립 필터입니다. 선택한 종류를 합쳐서 보여주며, 모두 끄면 전체 파일을 표시합니다.

텍스트는 UTF-8/BOM, UTF-16/BOM, CP949, Windows-1258/1252 순으로 읽습니다. JSON은 정렬해 표시하고 아주 큰 텍스트는 앞부분 16 MiB까지만 표시합니다. 인코딩 자동 판별이 애매한 파일은 원본 연결 프로그램으로 확인하세요.

## LEFT 파일 브라우저

- `↑ / ↓`: 파일 이동
- `Ctrl+F`: 파일명 검색
- Search에서 `↑ / ↓`: 검색창을 빠져나와 결과 목록으로 즉시 이동
- `Ctrl+Click`, `Shift+Click`: 멀티 선택
- 오른쪽 마우스 드래그: 지나가는 파일들을 연속 선택
- `Del`: Windows 휴지통으로 이동
- `Enter`: Windows 기본 연결 프로그램으로 열기
- 왼쪽 더블클릭: 폴더 이동 또는 기본 연결 프로그램으로 파일 열기
- LEFT에서 `Tab / Shift+Tab`: 의도적으로 비활성화
- `Ctrl+Left / Ctrl+Right`, `Alt+1 / Alt+2`: LEFT/RIGHT 패널 이동

## PDF / Image Viewer

- PDF 페이지 이동: 처음 / 이전 / 다음 / 마지막
- Zoom 50%~200%
- PDF는 Fit으로 시작하며 `1:1 / Fit` 전환 가능
- PDF 현재 페이지를 이미지로 Copy
- Image는 `1:1 / Fit` 모드가 기본
- Image Copy / Save Image As
- Bright / Dark / System theme

## 설치

Windows 11 권장, Python 3.11 이상.

압축을 푼 뒤:

```text
install_xViewer.bat
```

설치 프로그램은 private Python environment를 만들고 필요한 패키지를 설치합니다. Windows에서는 빠른 Excel Preview를 위해 `pywin32`도 설치됩니다.

## 개발 실행

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -e ".[dev]"
.\.venv\Scripts\python.exe -m mdir
```

Self-check:

```powershell
.\.venv\Scripts\python.exe -m mdir --check
```

Tests:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

## 주요 모듈

- `mdir/workspace_app.py` — dual-pane UI, file browser, selection, Excel preview routing
- `mdir/excel_pdf_preview.py` — 변환 제한시간/취소, 전용 Excel 관리, PDF 캐시
- `mdir/preview_worker.py` — 별도 프로세스에서 실행되는 Excel COM 작업
- `mdir/ui_dispatch.py`, `mdir/background.py` — 화면 결과 전달 및 제한된 작업 대기열
- `mdir/text_viewer.py` — 읽기 전용 텍스트 보기
- `mdir/media_viewer.py` — PDF/Image rendering and scrolling
- `mdir/excel_viewer.py` — previous native Excel renderer retained as fallback
- `mdir/links.py` — mDIR-compatible Link Manager
- `mdir/theme.py` — Bright/Dark/System theme

## Cache

Excel PDF Preview:

```text
%LOCALAPPDATA%\xViewer\cache\excel-pdf
```

Cache는 source path + size + modified time + cache version으로 구분됩니다. 원본 Excel이 수정되면 자동으로 새 Preview가 생성됩니다. 오래된 Preview는 cache pruning에 의해 자동 정리됩니다.

## 진단과 검증 범위

단계별 시간·오류는 `%LOCALAPPDATA%\xViewer\logs\excel-preview.log`에 남습니다(2MB × 최대 3개). 셀 내용은 기록하지 않지만 파일 경로는 포함됩니다.

설치 후 `diagnose_excel.bat`를 실행해 문제가 있는 Excel을 선택하면 실제 사용자 Windows 세션에서 변환을 검사하고 같은 로그 폴더에 JSON 보고서를 만듭니다. 생성 PDF 경로도 보고서에 기록됩니다. 파일을 배치 파일 위로 끌어다 놓아도 됩니다.

이번 작업의 자동 검사와 실제 파일 읽기 결과, 아직 확인하지 못한 범위는 `REVIEW-2.7.4.md`를 확인하세요. Office가 없는 환경, 암호 파일, 손상된 파일, 지원하지 않는 그림은 완전한 표시를 보장할 수 없습니다.

## License

MIT License

# Prior Tools Built for Maker: `loc-writer` and `outlook-cli-for-msds`

Primary-source analysis of `/Users/seungjoo/workspace/loc-writer` and `/Users/seungjoo/workspace/outlook-cli-for-msds` to inform Wayfinder tickets #7 (delivery formats & pre-flight check) and #8 (numerical accuracy & 화관법 compliance safeguards).

---

## 1. Executive Comparison

| Dimension | `loc-writer` (`/Users/seungjoo/workspace/loc-writer`) | `outlook-cli-for-msds` (`/Users/seungjoo/workspace/outlook-cli-for-msds`) |
| --- | --- | --- |
| **Domain / Job** | LoC lookup, DOCX generation, preview, & PDF export | Bulk MSDS email assembly, Drafts staging, & attachment match |
| **Stack** | Python 3.9+ (`PyQt6`, `python-docx`, `pywin32`) + Rust (`calamine 0.26`, `rusqlite 0.31`) | Python (`openpyxl`, `pywin32`, `argparse`, `logging`, `re`) |
| **Delivery** | Windows `release.zip`: `--onefile` EXEs (`loc_writer.exe` 47MB, `excel2db.exe` 2.2MB) + external `.xlsx`/`.docx`/`README.md` (`README.md:9-18`) | Single `--onefile` EXE (`dist/outlook-cli.exe`, `README.md:67-77`) + `outlook_config.json`, `usage.txt`, `signature/signature.html` |
| **Target PC** | Low-spec Win 11 64-bit, MS Word, zero dev runtime (`AGENTS.md:5-6`, `README.md:26-28`) | Windows PC with desktop Outlook MAPI logged in (`REQUIREMENTS.md`, `outlook_mail.py:30-53`) |
| **UX Pattern** | GUI, drag-and-drop Excel picker, HTML preview, `os.startfile` DOCX edit, 15-item FIFO history (`recent_history.json`) | CLI flags (`--excel`, `--msds-dir`, `--draft`) + interactive numbered prompt when run without args (`outlook_cli.py:90-189, 230-234`) |
| **Office COM** | `calamine` -> SQLite; `python-docx`; `DispatchEx("Word.Application")` PDF (`pdf_converter.py:14-28`) | `openpyxl` `.xlsx`; `Dispatch("Outlook.Application")` with `CoInitialize`/`CoUninitialize` (`outlook_mail.py:28-110`) |
| **Safety** | `ReadOnly=True`, `SaveChanges=False`, `upx=False` against AV false positives (`loc_writer.spec:43`) | Drafts-first (`mail.Save()`), primary sender guard, `제품코드_` prefix, `[base]_skipped_[ts].xlsx`, 7-day UTF-8 audit logs |
| **Compliance** | Fragile: trusts Excel formulas; ignores `#REF!`, hardcodes `100%`, skips cut-off checks | Structural only: required `Company`/`Contact`, `YYYY-MM-DD HH:MM:SS` timestamp, `제품코드_` attachment check |

---

## 2. Detailed Findings: `loc-writer`

### 2.1 Architecture, Languages, and Packaging
- **Hybrid Python + Rust split**:
  - `main.py:1-12` launches a `PyQt6` desktop app (`MainWindow`).
  - `tools/excel2db/src/main.rs:1-120` is a standalone Rust CLI (`calamine = "0.26"`, `rusqlite = "0.31"` with `features = ["bundled"]`, LTO release profile in `tools/excel2db/Cargo.toml:7-13`). It converts a large workbook (`★화학물질_20260107_samples.xlsx`, where `화학물질_20260506` has 47,518 rows and `LoC` has 1,014 rows) into `loc_database.db` in ~0.15s (`README.md:14, 37`), avoiding `openpyxl` startup lag on low-spec PCs (`AGENTS.md:5`).
- **PyInstaller `--onefile` + sidecar mutable files**:
  - `loc_writer.spec:14-51` builds `loc_writer.exe` (`console=False`, `excludes=['tkinter']`, `upx=False` with comment: *"UPX는 백신 오탐을 유발할 수 있어 끈다 — 비개발자 배포에서 지원 부담이 큼"* at line 42).
  - Only `src/gui/styles/theme.qss` is bundled into `_MEIPASS` (`loc_writer.spec:20`). User-editable files (`LOC - EMG - ProductName - Date.docx`, `★화학물질_*.xlsx`, `excel2db.exe`, `README.md`) are deliberately placed **beside** the `.exe` (`loc_writer.spec:7-12`, `README.md:126`).
  - `src/utils/paths.py:4-19` resolves resources in order: `base_dir` (`os.path.dirname(sys.executable)` when `sys.frozen`) -> `sys._MEIPASS` -> CWD.
  - `build_release.py:40-83` guards against non-Windows execution (`sys.platform != "win32"`), runs `cargo build --release` + `python -m PyInstaller loc_writer.spec --noconfirm`, assembles `release/` and `release.zip`, and optionally publishes via `gh release create`.
  - `requirements.txt:3-6` lists `PyQt6`, `python-docx`, `pywin32`, `pyinstaller` (versions unpinned).

### 2.2 Non-Developer UX, Korean Encoding, and Office Integration
- **Zero-CLI GUI & "Word Studio" metaphor**:
  - `src/gui/views/main_window.py:19-162` & `src/gui/styles/theme.qss:1-180` use Word Deep Navy (`#1b365d`), `'Segoe UI', 'Malgun Gothic', sans-serif`, plain Korean labels (`조회`, `PDF 변환`, `docx에서 직접 편집`), and a right-hand HTML preview (`QTextEdit`).
  - **Excel Picker modal**: `MainWindow.__init__` checks `self.db_builder.needs_setup()` (missing DB or `.xlsx` mtime newer than `.db` in `src/db/bridge.py:81-91`) and opens `ExcelPickerDialog` (`src/gui/views/excel_picker_dialog.py:59-203`) featuring a drag-and-drop `DropFrame` for `.xlsx` files and a background `DBWorkerThread(QThread)` with an indeterminate progress bar.
  - **Human-in-the-loop escape hatch**: Searching an `EMG CODE` auto-generates `LoC/LOC - {emg_code} - {product_name} - {MMM}{YYYY}.docx` (`src/loc_session.py:114-129, 171-176`). Clicking `docx에서 직접 편집` invokes `os.startfile(self._current_docx_path)` (`src/loc_session.py:98`) so the user can tweak formatting/text in MS Word before clicking `PDF 변환`.
  - **Quick re-lookup**: `recent_history.json` persists up to 15 FIFO codes (`src/loc_session.py:178-196`, `docs/adr/0002-json-recent-history-persistence.md`).
- **Word COM PDF conversion safeguards (`src/document/pdf_converter.py:4-28`)**:
  - Uses `win32com.client.DispatchEx("Word.Application")` (not `Dispatch`) so it launches an isolated Word instance instead of attaching to—and `Quit()`-ing—a user's open Word session (`pdf_converter.py:12-14`).
  - Sets `word.Visible = False`, `word.DisplayAlerts = 0` (`wdAlertsNone`), opens `ReadOnly=True`, exports with `ExportFormat=17` (PDF), closes with `SaveChanges=False`, and cleans up in `finally: word.Quit()`.
- **DOCX manipulation (`src/document/docx_handler.py:13-57`)**:
  - Directly mutates `doc.tables[0]` and `doc.tables[2]` of `LOC - EMG - ProductName - Date.docx`, cloning `t2.rows[1]._tr` via `copy.deepcopy` before `total_row._tr.addprevious(new_tr)` so the `Total Contents | 100%` bottom row stays intact.
- **Encoding & Error Handling Gaps**:
  - Subprocess output from `excel2db.exe` is decoded with `encoding="utf-8", errors="ignore"` (`src/db/bridge.py:115-116`) and only logged via `print(...)` (`bridge.py:103-125`), which is invisible in a `console=False` PyInstaller GUI build; `DatabaseSetup.select` only returns generic `"엑셀 변환 중 오류가 발생했습니다. 로그를 확인하세요."` (`src/db/setup.py:35`).
  - `tools/excel2db/src/main.rs:24-26` deletes the existing `loc_database.db` **before** opening the Excel file or verifying the `LoC` sheet exists—meaning a bad Excel file destroys a working DB.
  - `src/gui/views/main_window.py:238-247` is missing `return` in `except NotReadyError:` (also flagged in `scratchpad/architecture-review-20261013-132513.html:56`), causing a false-success dialog or `UnboundLocalError` on `pdf_path`.
  - `src/document/loc_document.py:47-112` interpolates `product_name` and cell strings into preview HTML without `html.escape()`.

### 2.3 Numerical, Formula, and Compliance Safeguards (Observed Reality)
Inspection of `★화학물질_20260107_samples.xlsx`, `tools/excel2db/src/main.rs`, `src/models/loc_data.py`, and `src/document/loc_document.py` reveals critical lessons for ticket #8:
1. **Excel formula / cached-value reliance & `#REF!` silent pass-through**:
   - In `★화학물질_20260107_samples.xlsx` (`LoC` sheet, 1,014 rows), almost all compliance columns (`A`, `D`, `F`, `H`, `I`–`S`) are Excel formulas (`VLOOKUP` into the 47,518-row `화학물질_20260506` sheet, and `IF(H="","CBI",...)` / `IF(F="","N","P")`).
   - Row 2 (`ABC123`, `11118-57-3`, `함량=90`) has **cached `#REF!` error values** (`t="e"`, `<v>#REF!</v>`) in columns `D`, `F`, `H`, and `I`–`S` due to broken `VLOOKUP` references.
   - `excel2db` (`tools/excel2db/src/main.rs:76-81`) converts every cell via `val.to_string().trim().to_string()` into `TEXT` columns in SQLite (`main.rs:38-53`) with **zero check** for `Data::Error` (`#REF!`, `#N/A`, `#VALUE!`), meaning broken formula errors silently flow into the LoC document.
2. **Column-index fragility & CBI leakage history**:
   - Commit `a1f2a21` (`fix: correct CBI column mapping and enforce direct PDF overwrite`) changed `excel2db` to read columns `8, 9, 10` (`I`/`J`/`K`, CBI-masked formulas) instead of unmasked columns `C`/`D`/`E`. Prior to that fix, CBI-protected ingredients were exposed because positional indices (`get_str(8)`, `get_str(9)`, etc.) had no header-name validation.
3. **Hardcoded `Total Contents = 100%` without arithmetic check**:
   - Both `LOC - EMG - ProductName - Date.docx` (table 2 last row) and `src/document/loc_document.py:101-104` hardcode `Total Contents` = `100%`. Neither `excel2db` nor `loc_writer` parses `content_pct` as a number, sums components, or flags sums != 100% (or mixed `"CBI"` / range strings like `10-20`).
4. **No 화관법 threshold cross-check**:
   - Sheet `화학물질_20260506` contains regulatory cut-off text in column `K` (`유해특성분류 및 혼합물 함량기준(%)`, e.g. `유독물질 : 인체급성유해성 : 25%, 인체만성유해성 : 0.1%, ...`), and `LoC` relies purely on manual/Excel `VLOOKUP` without program-level verification of `함량` vs. threshold %.
5. **Explicit null/default conventions**:
   - `GLOSSARY.md:7` and `src/models/loc_data.py:23-35` enforce keeping empty hazard cells as `""` (never `"N"`), while `chem_div` defaults to `"P"` when `None`/`""` (`loc_data.py:27`). Filename sanitization replaces `/` and `\` with `_` (`src/loc_session.py:172`).

---

## 3. Detailed Findings: `outlook-cli-for-msds`

### 3.1 Architecture, Languages, and Packaging
- **Pure Python CLI + Windows COM**:
  - Entrypoint `outlook_cli.py` + domain module `core/distribution.py` + Outlook COM adapter `infrastructure/outlook_mail.py`.
  - Uses `openpyxl` for `.xlsx` read/write and `pywin32` (`win32com.client`, `pythoncom`) for Outlook.
  - Built into a single EXE (`dist/outlook-cli.exe`) via `pyinstaller --onefile --clean --name outlook-cli --exclude-module tkinter --exclude-module unittest --exclude-module pydoc outlook_cli.py` (`README.md:76`). Unlike `loc-writer`, there is **no** `.spec` file, `requirements.txt`, or `pyproject.toml` committed, and `outlook_cli.py:68` uses `os.path.dirname(os.path.abspath(__file__))` for `outlook_config.json` (which resolves inside PyInstaller's temp `_MEIPASS` rather than beside the `.exe` under `--onefile`!).
  - `PYTHON.md:7-10` mandates `uv` (`uv venv`, `uv pip install`, `uv run`) for local dev/packaging.

### 3.2 Non-Developer UX, Korean Encoding, and Excel/Outlook Integration
- **Dual-mode CLI + interactive prompt (`outlook_cli.py:90-241`)**:
  - When run with no arguments (`len(sys.argv) == 1`) or without `--to`/`--excel`, drops into a Korean interactive wizard (`run_interactive_mode`) that lists files in CWD by number (`[1] ...`) so users don't have to type paths, and strips accidental quotes `.strip('"').strip("'")` from dragged-and-dropped file paths (`outlook_cli.py:132, 167, 264`, `core/distribution.py:100`).
  - Also includes a plain-text quick manual `usage.txt` alongside `README.md`.
- **Excel schema & resilient partial-batch processing (`core/distribution.py:21-251`)**:
  - Reads two sheets case-insensitively (`Recipients` and `TableData`) with `openpyxl.load_workbook(..., data_only=True)` (`core/distribution.py:167`).
  - Header lookup is case-insensitive (`str(h).strip().lower()`) and supports Korean/English aliases for deferred time (`deferredtime`, `deferred_time`, `deferred`, `예약시간`, `예약발송` at `core/distribution.py:113-119`), semicolon-separated `Cc`/`Bcc`/`Attachments`, and auto-ignores legacy `"번호"` column in `TableData` while auto-generating 1..N row numbers in HTML (`TableDataSheetParser` + `HtmlTableRenderer`, `core/distribution.py:39, 258-280`).
  - **Structural vs. per-row error split (`docs/adr/0003-mail-distribution-context-and-audit-log-retention.md:15-16`)**:
    - Structural errors (missing file, missing `Recipients` sheet or `To` column, missing `Email` in `TableData`) abort immediately with Korean `ValueError`/`FileNotFoundError`.
    - Per-recipient errors (missing `Company` or `Contact`, missing MSDS file for a product code, invalid `YYYY-MM-DD HH:MM:SS` deferred timestamp) **skip only that recipient**, log an error, and at `close()` export `[original]_skipped_[YYYYMMDD_HHMMSS].xlsx` with the exact `Recipients` + `TableData` structure so the user can fix and re-run only failed rows (`core/distribution.py:204-251, 506-530, 547-552`).
- **Korean & Windows file encoding handling**:
  - Signature loader (`core/distribution.py:408-414`) tries `["utf-8", "cp949", "euc-kr", "utf-16"]` sequentially when reading `signature.html` (which in `signature/signature.html` is Microsoft Word-exported HTML with `charset=utf-8` and `맑은 고딕`).
  - Audit logs (`logs/distribution_YYYYMMDD_HHMMSS.log`) force `encoding="utf-8"` (`core/distribution.py:436`) and auto-delete logs older than 7 days on startup (`_cleanup_old_logs`, lines 451-476).
  - Test runner (`tests/test_cli.py:30, 36`) explicitly sets `PYTHONIOENCODING="utf-8"` and `encoding="utf-8", errors="replace"` for Windows subprocess Korean output.
- **Outlook COM safety guardrails (`infrastructure/outlook_mail.py:28-110`, `docs/adr/0002`, `docs/adr/0004`)**:
  - Calls `pythoncom.CoInitialize()` / `CoUninitialize()`.
  - Validates `sender` against `outlook.Session.Accounts` (`SmtpAddress`), raising a Korean `ValueError` if unregistered; blocks sending from a non-primary account unless `--force-sender` is passed (`infrastructure/outlook_mail.py:41-60`).
  - Bulk mode (`outlook_cli.py:19`) sets `is_interactive=False`; when `--draft` is set, calls `mail.Save()` (background save to Drafts without popping up dozens of Outlook windows) instead of `mail.Display()`.
  - Forces `맑은 고딕` `11.0pt` and `MsoTableGrid` inline CSS (`border: solid windowtext 1.0pt`) so generated emails match Outlook's native Korean typography (`docs/adr/0004`).
- **Observed Code-Documentation Discrepancies**:
  - `README.md:57-58` documents `--non-interactive` and `--force-primary` flags, whereas `outlook_cli.py:205` implements `--force-sender` and has no `--non-interactive` CLI flag.
  - `test_recipients_v4.xlsx` lacks `Company`, `Contact`, and `Bcc` columns required by `RecipientValidator` (`core/distribution.py:337-342`) and `README.md:26-28`.

### 3.3 Numerical / Compliance Safeguards
- **Strict prefix matching for MSDS attachments (`core/distribution.py:323-328`, `docs/adr/0001`)**: Enforces `filename.lower().startswith(prod_code.lower() + "_")` so product code `P1003` never attaches `P10030000000000123213.pdf` (a real bug reproduced in `test_msds/` and `tests/test_distribution.py:88-118`). Missing any product's MSDS skips the recipient rather than sending an incomplete compliance email.
- **No numeric calculation**: Product tables contain codes/names only; no percentages, sums, or threshold arithmetic exist in `outlook-cli-for-msds`. (Note: in `core/distribution.py:121-126, 344-348`, if Excel stores `예약시간` as an Excel native `datetime` cell rather than a string, `validate_deferred` returns `True` without checking range/past-date validity.)

---

## 4. Concrete Lessons & Constraints for the Non-Developer Vibe-Coding Skill

### For Ticket #7 (Delivery Formats & Pre-Flight Capability Check)
1. **Default to zero-install, sidecar-configurable Windows `.exe` (with `upx=False`) or pure-local HTML/browser when possible**: Maker's PC is low-spec Windows 11 with no Python/dev runtime (`loc-writer/AGENTS.md:5`). PyInstaller `--onefile` (`console=False` for GUI, `upx=False` to prevent corporate AV quarantine) worked, **but** any user-editable assets (`.xlsx`, `.docx` templates, `signature.html`, `config.json`) must live **beside** the `.exe` via `sys.frozen` / `os.path.dirname(sys.executable)`—never `__file__` (a bug in `outlook_cli.py:68`) and never baked inside `_MEIPASS`.
2. **Pre-flight check must verify Windows Office/Outlook COM & build constraints**:
   - Both tools depend on Windows-only `pywin32` COM (`Word.Application` / `Outlook.Application`) and cannot even be built (`loc_writer.spec`) or integration-tested on macOS. A skill targeting Maker's PC must check up front: (a) macOS dev machine vs. Windows target machine, (b) whether MS Word / Desktop Outlook (classic MAPI COM, not New Outlook) is installed and logged in, (c) whether corporate AV blocks unsigned `.exe` or scripts.
3. **CLI + `--draft` + skipped-row recovery vs. GUI + preview**:
   - For high-consequence actions (bulk emailing, official LoCs), **always build a human-in-the-loop staging step**: `mail.Save()` to Drafts (`outlook-cli-for-msds`), HTML preview + `os.startfile` direct Word edit (`loc-writer`), and `[input]_skipped_[timestamp].xlsx` partial-retry export (`outlook-cli-for-msds`).
   - GUI (`console=False`) silently swallows `print()` diagnostics (`loc-writer/src/db/bridge.py`), whereas `outlook-cli-for-msds` writes UTF-8 timestamped audit logs (`logs/distribution_*.log`) with 7-day auto-cleanup. Future GUI tools must also write persistent Korean logs and surface actionable error details in dialogs.
4. **Input ergonomics for non-developers**:
   - Drag-and-drop `.xlsx` zones (`ExcelPickerDialog`), quote-stripping on pasted paths (`.strip('"').strip("'")`), `cp949`/`euc-kr`/`utf-8`/`utf-16` encoding fallback for user-supplied HTML/text files, header-name matching (instead of brittle column indices), and filename sanitization (`/` `\` replacement).

### For Ticket #8 (Numerical Accuracy & 화관법 Compliance Safeguards)
1. **Never trust Excel cached formula values blindly**:
   - `★화학물질_20260107_samples.xlsx` proved that real workbooks contain broken `VLOOKUP`s yielding `#REF!` cached error cells (`t="e"`), stale cached values when edited outside Excel, and `data_only=True` / `calamine` stringification (`val.to_string()`) silently passes `#REF!`, `#N/A`, `#VALUE!` straight into official compliance PDFs. **Rule**: Explicitly reject/quarantine Excel error values (`#REF!`, `#N/A`, `#VALUE!`, `#NAME?`, `#DIV/0!`) at ingest.
2. **Never hardcode totals (`100%`) or skip cross-checks**:
   - `loc-writer` hardcodes `Total Contents | 100%` in both template and HTML preview without summing `content_pct`. **Rule**: Parse numeric contents using `Decimal` (not binary `float`), verify component sums (with explicit tolerance only for rounded ranges, and explicit handling for `"CBI"` / range strings like `a~b`), and block or warn prominently on mismatch.
3. **Enforce header-based schema validation before positional extraction**:
   - `loc-writer` initially leaked raw chemical names/CAS/contents instead of CBI-masked columns (`a1f2a21`) because `excel2db` hardcoded column indices `8..18`. **Rule**: Validate expected header names (as `outlook-cli-for-msds` does) and explicitly test CBI masking invariants.
4. **Verify regulatory cut-off thresholds in code, not only via Excel `VLOOKUP`**:
   - `화학물질_20260506` stores mixture cut-off criteria (`유독물질 : ... 25%, ... 0.1%`), yet `loc-writer` performs zero programmatic check comparing `함량` against those thresholds or verifying `Chemical Division (P/N)`. Compliance tools must treat Excel formulas as untrusted inputs and implement independent validation/warning gates.
5. **Preserve original state on failed rebuilds**:
   - `excel2db` deletes `loc_database.db` before opening/validating the `.xlsx` (`tools/excel2db/src/main.rs:24-26`). **Rule**: Stage conversions to a temp file and atomically swap only after schema + error + numerical validation passes.

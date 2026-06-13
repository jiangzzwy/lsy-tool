# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Batch generation tool for 案件线索移送函 (case clue transfer letters). Reads an Excel worksheet of complaint orders, classifies each row into one of four categories (外卖/三方公司/三方个人/自营开票), fills category-specific Word templates, and produces a consolidated 台账 (ledger) Excel.

## Running the Application

- **GUI mode**: `python3 app.py` (uses dearpygui desktop UI)
- **CLI mode**: `python3 main.py` (headless, reads config.py directly)
- **Shortcut**: `./run.sh` (runs GUI mode)
- **Install deps**: `pip install -r requirements.txt` (openpyxl, python-docx, dearpygui, requests, pyinstaller)

## Packaging

- **macOS**: `python3 build.py --mac` → `dist/移送函生成工具.app`
- **Windows**: `python3 build.py --win` (run on Windows) → `dist/移送函生成工具.exe`
- **Clean**: `python3 build.py --clean`
- The `demands/` directory (templates + source Excel) is bundled into the app via `--add-data`
- `app_base_dir()` in `app.py` resolves the correct base directory whether running from source or a PyInstaller frozen bundle (`sys._MEIPASS`)

## Architecture

### Pipeline (4 steps, same in both CLI and GUI)

1. **Parse Excel** (`excel_parser.py`) — `parse_excel()` reads source xlsx into `RowData` objects; `classify_row()` determines category; `split_rows()` produces `SplitItem` list (one doc per shop when a row has multiple shops)
2. **Generate Word docs** (`word_generator.py`) — `generate_word_docs()` iterates `SplitItem` list, picks template by classification, fills placeholder paragraphs (subtitle, bureau, body, date), saves to `output/{category}/` subdirectories
3. **Generate ledger** (`ledger_generator.py`) — `generate_ledger()` writes one row per `SplitItem` into a new xlsx with fixed legal text columns and derived fields
4. **Bureau lookup** (`api_client.py`) — `CreditCodeService` wraps one of three `BaseApiClient` implementations:
   - `MockApiClient`: heuristic regex on address strings + special rules map (no API needed)
   - `TianyanchaApiClient`: real API lookup by credit code
   - `QichachaApiClient`: real API lookup by credit code
   Results are cached to `output/.credit_code_cache.json`

### Key Data Flow

`config.py` holds all settings (source file paths, template map, column indices, API provider). Both `main.py` and `app.py` import config directly. The GUI (`app.py`) overrides config values from UI inputs before running the pipeline.

### Classification Logic (`excel_parser.py:classify_row`)

- Category "M" starts with "外卖美食"/"外卖" → **外卖**
- Column H = "三方"/"三方/自营" with valid credit code → **三方公司**, else → **三方个人**
- Column H = "自营" → **自营开票**
- Anything else → **其他** (skipped from output)

### Bureau Lookup Fallback Chain

For 三方公司 and 自营开票: credit code API lookup → address heuristic fallback → placeholder text. For 三方个人: address heuristic only (no credit code expected). For 外卖: hardcoded `上海市杨浦区市场监督管理局`.

## Key Conventions

- Column references in `config.py.COL` are 1-based indices matching the source Excel layout
- Template Word files use specific paragraph positions (paragraph 2 = subtitle, paragraph 4 = bureau, paragraph 5 = body, paragraph 12 = date) — edits to templates must preserve this structure
- `demands/` directory contains source Excel and all 4 Word templates; `output/` receives generated files
- `ledger_generator.py` reuses `_get_bureau()` from `word_generator.py` via import to avoid duplicating bureau resolution logic
- The GUI runs the pipeline in a background thread (`threading.Thread(daemon=True)`) to avoid blocking the DearPyGui event loop
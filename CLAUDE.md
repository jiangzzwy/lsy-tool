# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Batch generation tool for 案件线索移送函 (case clue transfer letters). Reads an Excel worksheet of complaint orders, classifies each row into one of four categories (外卖/三方公司/三方个人/自营开票), fills category-specific Word templates, and produces a consolidated 台账 (ledger) Excel.

## Running the Application

- **GUI mode**: `python3 web.py` or `./run.sh` — Flask server on port 5004 + pywebview desktop window
- **CLI mode**: `python3 main.py` — headless, reads config.py directly
- **Import bureau data**: `python3 main.py import <filled_template.xlsx>` — bulk-import bureau mappings from a filled export template
- **Install deps**: `pip install -r requirements.txt`

## Packaging

- **macOS**: `python3 build.py --mac` → `dist/移送函批量生成工具.app`
- **Windows**: `python3 build.py --win` (on Windows) → `dist/移送函批量生成工具.exe`
- **Clean**: `python3 build.py --clean`
- The `demands/`, `web/templates/`, `web/static/` directories are bundled via `--add-data`
- `base_dir()` in `web.py` resolves the correct base directory whether running from source or a PyInstaller frozen bundle (`sys._MEIPASS`)
- CI: `.github/workflows/build-windows.yml` auto-builds Windows EXE on push to main and creates a GitHub Release

## Architecture

### Pipeline (5 steps, same in both CLI and GUI)

1. **Parse Excel** (`excel_parser.py`) — `parse_excel()` reads source xlsx into `RowData` objects; `classify_row()` determines category; `split_rows()` produces `SplitItem` list (one doc per shop when a row has multiple shops)
2. **Split rows** (`excel_parser.py`) — multi-shop rows split into one `SplitItem` per shop; `其他` category rows are dropped
3. **Pre-check bureau coverage** (`api_client.py:BureauDB.check_coverage`) — checks which credit codes have bureau mappings in the local SQLite DB; exports missing ones as a template Excel for user to fill in; CLI exits if coverage is incomplete
4. **Generate Word docs** (`word_generator.py`) — iterates `SplitItem` list, picks template by classification, fills placeholder paragraphs (subtitle, bureau, body, date), saves to `output/{category}/` subdirectories
5. **Generate ledger** (`ledger_generator.py`) — writes one row per `SplitItem` into a new xlsx with fixed legal text columns and derived fields

### Bureau Lookup System (`api_client.py`)

No API calls. Uses a local SQLite database (`output/.bureau_cache.db`) with an import/export workflow:

- **`BureauDB`** — SQLite-backed cache: `lookup(credit_code)` → bureau name; `check_coverage(items)` → missing codes report; `export_template(missing, path)` → Excel template for user to fill; `import_template(path)` → bulk-insert filled data
- **`MockApiClient`** — address heuristic fallback (regex pattern matching on address strings + `SPECIAL_RULES` dict). Used for rows without credit codes (三方个人) or when DB lookup fails
- **Fallback chain**: DB lookup by credit code → address heuristic → placeholder text
- The GUI provides a full mapping table CRUD (add/edit/delete entries) and an import-update button

### Key Data Flow

`config.py` holds all settings (source file paths, template map, column indices). Both `main.py` and `web.py` import config directly. The GUI (`web.py`) uses Flask REST API (`/api/*`) to run the pipeline in background threads.

### Classification Logic (`excel_parser.py:classify_row`)

- Category "M" starts with "外卖美食"/"外卖" → **外卖**
- Column H = "三方"/"三方/自营" with valid credit code → **三方公司**, else → **三方个人**
- Column H = "自营" → **自营开票**
- Anything else → **其他** (skipped from output)

### Bureau Lookup Fallback by Classification

- **外卖**: hardcoded `上海市杨浦区市场监督管理局`
- **三方公司** and **自营开票**: DB lookup by credit code → address heuristic → placeholder
- **三方个人**: address heuristic only (no credit code expected)

## Key Conventions

- Column references in `config.py.COL` are 1-based indices matching the source Excel layout
- Template Word files use specific paragraph positions identified by content matching (starts-with "京技管商务消移送", contains "市场监督管理局" + ends with "：", starts-with "我单位在举报调查中发现", date regex) — edits to templates must preserve this structure
- `demands/` directory contains source Excel and all 4 Word templates; `output/` receives generated files
- `ledger_generator.py` reuses `_get_bureau()` from `word_generator.py` via import to avoid duplicating bureau resolution logic
- The GUI runs the pipeline in a background thread (`threading.Thread(daemon=True)`) to avoid blocking the Flask event loop
- All bureau state is in SQLite; `config.py` retains legacy API provider constants but they are unused

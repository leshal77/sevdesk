# CLAUDE.md

## Project Overview

**sevDesk Extrakt** is a single-file Python CLI utility (v0.1.30) that extracts voucher data from the [sevDesk](https://sevdesk.de) accounting API, generates Excel reports, downloads PDF documents, and tags processed vouchers in sevDesk for tracking.

- **Language**: Python 3
- **Entry point**: `sevdesk_vouchers_v0.1.30.py`
- **Architecture**: Single-file script, functional style (no classes), 13 functions + `main()`

## Running the Script

```bash
python sevdesk_vouchers_v0.1.30.py \
  --begin DD.MM.YYYY \
  --end DD.MM.YYYY \
  --endRechnungsdatum DD.MM.YYYY \
  --type income|expense \
  --exportTag EXPORT_TAG_NAME \
  [--token API_TOKEN] \
  [--no-verify] \
  [--debug] \
  [--skip-pdf] \
  [--skip-tagging]
```

The API token can be passed via `--token` or the `SEVDESK_API_TOKEN` environment variable.

## Dependencies

No `requirements.txt` exists. Install manually:

```bash
pip install requests openpyxl
```

Standard library modules used: `argparse`, `datetime`, `sys`, `urllib3`, `os`, `shutil`, `re`, `base64`.

## Codebase Structure

The entire application lives in a single file: `sevdesk_vouchers_v0.1.30.py` (864 lines).

### Function Map

| Function | Lines | Purpose |
|---|---|---|
| `parse_date()` | 23–38 | Parse dates in DD.MM.YYYY or YYYY-MM-DD format |
| `convert_date_to_dir_format()` | 40–43 | Convert dates for directory names |
| `fetch_vouchers()` | 45–137 | Fetch vouchers from sevDesk API with creditDebit filtering |
| `fetch_tags_for_voucher()` | 139–192 | Get tags for a single voucher |
| `fetch_tags_for_all_vouchers()` | 194–231 | Batch fetch tags for all vouchers |
| `format_currency()` | 233–235 | Format amount with currency string |
| `sanitize_filename()` | 237–247 | Clean filenames of invalid characters |
| `tag_vouchers_in_sevdesk()` | 249–401 | Create tags on vouchers with creditDebit validation (v0.1.30 fix) |
| `download_voucher_pdf()` | 403–461 | Download a single voucher PDF (base64 or binary) |
| `download_pdfs_for_vouchers()` | 463–514 | Batch PDF downloading |
| `create_xlsx_sheet()` | 516–563 | Format and populate an Excel worksheet |
| `create_output_directory()` | 565–602 | Create output directory structure |
| `main()` | 604–864 | Orchestrate the full workflow |

### Key Architecture: Three-Stage Filtering Pipeline

The `main()` function implements a 3-stage filter on vouchers:

1. **Stage 1** — Invoice date (`voucherDate`) must fall between `--begin` and `--endRechnungsdatum`
2. **Stage 2** — Delivery date (`deliveryDate`) must fall between `--begin` and `--end`
3. **Stage 3** — Exclude vouchers that already have tags (previously exported)

Results are split into two categories:
- **Volltreffer** (matches) — passed all 3 stages, will be tagged and have PDFs downloaded
- **Abgelehnt** (rejected) — failed stage 2 or 3

### sevDesk API

- Base URL: `https://my.sevdesk.de/api/v1/`
- Key endpoints:
  - `GET /Voucher` — fetch vouchers (with `creditDebit` filter: `D`=income, `C`=expense)
  - `GET /Tag` — fetch tags for a voucher
  - `POST /Tag/Factory/create` — create a tag on a voucher
  - `GET /Voucher/{id}/downloadDocument` — download PDF (returns base64-encoded content)
- Auth: `Authorization` header with API token
- All requests use a persistent `requests.Session()`

### Output Structure

```
{exportTag}/
├── Ausgangsrechnungen/          (for --type income)
│   ├── ausgangsrechnungen_{begin}_{end}.xlsx
│   └── 1.pdf, 2.pdf, ...       (numbered to match Excel row numbers)
└── Eingangsrechnungen/          (for --type expense)
    ├── eingangsrechnungen_{begin}_{end}.xlsx
    └── 1.pdf, 2.pdf, ...
```

The Excel file has two worksheets: "Volltreffer" and "Abgelehnt", each with 8 columns: Nr., Rechn-Dat, Lief-Dat, Beleg-Nr, Lieferant/Kunde, Beschreibung, Betrag, Bereits exportiert (Tag).

## Code Conventions

- **Language**: Code comments, docstrings, CLI output, and variable names are in **German** (targeting German business users). Follow this convention.
- **No classes** — purely functional style with module-level functions
- **Session passing** — a `requests.Session` object is created once and passed to all API functions
- **Error handling** — try/except blocks with graceful degradation; errors are collected in lists and reported in summaries rather than raising exceptions
- **Logging** — prefix-based: `[info]`, `[debug]`, `[WARNING]`, `[ERROR]`. Debug output only shown when `--debug` is passed.
- **Progress indicators** — inline `\r` overwrites for batch operations (e.g., `"Tagging 5/100..."`)
- **Versioning** — stored in `__version__` variable and embedded in the filename (`sevdesk_vouchers_v{version}.py`)
- **No type annotations** — the codebase does not use Python type hints
- **4-space indentation** throughout

## Important Bug Fixes

- **v0.1.29**: Improved API filtering and debug outputs for creditDebit values
- **v0.1.30**: Added `creditDebit` validation before tagging to prevent incorrect classification; stores `creditDebit` in voucher data dict; skips vouchers with wrong type during tagging

## Development Notes

- **No tests, CI/CD, linting, or formatting tools** are configured
- **No package management** (no `requirements.txt`, `pyproject.toml`, or `setup.py`)
- **Cross-platform**: includes Windows UTF-8 console encoding workaround
- **SSL**: verification can be disabled via `--no-verify` for development/debugging
- When modifying the script, update `__version__` and the filename to match
- The `creditDebit` mapping is critical and unintuitive: `D` (Debit) = income/Ausgangsrechnungen, `C` (Credit) = expense/Eingangsrechnungen

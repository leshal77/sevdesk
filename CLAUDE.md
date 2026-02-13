# CLAUDE.md

## Project Overview

**sevDesk Extrakt** is a single-file Python CLI utility (v0.1.30) that extracts voucher data from the [sevDesk](https://sevdesk.de) accounting API, generates Excel reports, downloads PDF documents, and tags processed vouchers in sevDesk for tracking.

- **Language**: Python 3
- **Entry point**: `sevdesk_vouchers_v0.1.30.py`
- **Architecture**: Single-file script, functional style (no classes), 13 functions + `main()`
- **API spec**: `openapi.yaml` — full sevDesk OpenAPI 3.0 specification (17k lines)

## Business Context & Purpose

The script supports an **incremental export workflow** for a German accounting practice:

1. There are two distinct export categories that are run **separately** with completely different documents:
   - **Income** (`--type income`) — Ausgangsrechnungen (outgoing invoices, sales)
   - **Expense** (`--type expense`) — Eingangsrechnungen (incoming invoices, purchases)

2. Vouchers are selected through a **complex date-based filtering** pipeline (invoice date + delivery date ranges).

3. After a successful export, the script **tags each exported voucher** in sevDesk with the `--exportTag` value. On subsequent runs, already-tagged vouchers are excluded — enabling **incremental exports** (only new/unprocessed vouchers get exported).

### Known Critical Bug (reason for v0.1.30)

**Problem**: When running an income export with tag `incomeEXPORT_2025_DEZEMBER_01`, **expense vouchers were also getting tagged** with that income tag. This is a data corruption issue — expense vouchers become incorrectly marked as already exported under an income tag.

**Root cause**: The sevDesk API's `creditDebit` filter was not reliably filtering server-side. Vouchers with the wrong `creditDebit` value leaked through, and the tagging function did not validate the voucher type before applying the tag.

**Fix in v0.1.30**:
- `tag_vouchers_in_sevdesk()` now takes `expected_credit_debit` parameter and **skips** vouchers whose `creditDebit` doesn't match (line 324)
- `creditDebit` is stored in the voucher data dict so it's available at tagging time (line 718)
- Local safety-net filter in `main()` removes wrong-type vouchers after API fetch (line 651)

**This bug may not be fully resolved** — the local filtering is a workaround. The API itself may still return mixed results.

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

## sevDesk API Reference

Full API spec: `openapi.yaml` (OpenAPI 3.0)

- **Base URL**: `https://my.sevdesk.de/api/v1/`
- **Auth**: `Authorization` header with API token
- All requests use a persistent `requests.Session()`

### creditDebit Mapping (CRITICAL)

This mapping is **unintuitive** and the source of the cross-tagging bug:

| `creditDebit` value | Accounting term | German term | `--type` value | Meaning |
|---|---|---|---|---|
| `D` (Debit) | Money IN | Ausgangsrechnungen | `income` | You **sold** something |
| `C` (Credit) | Money OUT | Eingangsrechnungen | `expense` | You **bought** something |

The API's `creditDebit` query parameter is an enum `[C, D]` (see `openapi.yaml` VoucherModel, line ~16405). **The API filter is unreliable** — the script must filter locally as a safety net.

### Endpoints Used

#### `GET /Voucher` — Fetch vouchers
- **Params**: `limit` (int, default 1000), `creditDebit` (C|D), `embed` (e.g. `taxRule,supplier`), `startDate`, `endDate`
- **Response**: `{ "objects": [ VoucherResponse, ... ] }`
- Key VoucherResponse fields: `id`, `voucherDate`, `deliveryDate`, `creditDebit`, `sumGross`, `description`, `voucherNumber`, `supplier` (embedded object), `currency`
- **Caution**: API may return vouchers with wrong `creditDebit` despite the filter parameter

#### `GET /Tag` — Fetch tags for a voucher
- **Params**: `objectName=Voucher`, `objectId={voucherId}`
- **Response**: `{ "objects": [ { "id", "name", "objectName": "Tag", ... } ] }`
- Used to check if a voucher was already exported (has any tag)

#### `POST /Tag/Factory/create` — Create tag on voucher
- **Body**: `{ "name": "TAG_NAME", "object": { "id": voucherId, "objectName": "Voucher" } }`
- **Response**: Returns a `TagRelation` object (not just a Tag) linking the tag to the voucher
- **This is where the cross-tagging bug manifests** — must validate `creditDebit` before calling

#### `GET /Voucher/{voucherId}/downloadDocument` — Download PDF
- **Response**: JSON with `{ "objects": { "content": "base64...", "base64Encoded": true } }` or direct binary PDF
- Script handles both formats; timeout set to 60s for large documents

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

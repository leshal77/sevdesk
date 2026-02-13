# CLAUDE.md

## Project Overview

**sevDesk Extrakt** is a single-file Python CLI utility (v0.1.31) that extracts voucher data from the [sevDesk](https://sevdesk.de) accounting API, generates Excel reports, downloads PDF documents, and tags processed vouchers in sevDesk for tracking.

- **Language**: Python 3
- **Entry point**: `sevdesk_vouchers_v0.1.31.py`
- **Architecture**: Single-file script, functional style (no classes), 12 functions + `main()`
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
- `tag_vouchers_in_sevdesk()` now takes `expected_credit_debit` parameter and **skips** vouchers whose `creditDebit` doesn't match
- `creditDebit` is stored in the voucher data dict so it's available at tagging time
- Local safety-net filter in `main()` removes wrong-type vouchers after API fetch

### Known Critical Bug (reason for v0.1.31)

**Problem**: After running income export (step 1), expense export (step 2) reports ALL expense vouchers as "already tagged" with the income tag `incomeEXPORT_2025_DEZEMBER_01` and exports 0 Volltreffer. Expense vouchers are never exported.

**Root cause**: `GET /Tag?objectName=Voucher&objectId={id}` does NOT filter by `objectId`. The `objectName` and `objectId` parameters are **undocumented** (the official OpenAPI spec only defines `id` and `name` as parameters for `GET /Tag`). After the income export creates a tag on 9 income vouchers, `GET /Tag?objectName=Voucher&objectId={any_id}` returns that tag for ALL vouchers — it returns all tags associated with the Voucher object type globally, not for the specific voucher.

**Fix in v0.1.31**:
- Replaced `GET /Tag` (per-voucher, N API calls, buggy objectId filter) with `GET /TagRelation` (single API call, exact voucher↔tag mapping)
- `GET /TagRelation` returns `Model_TagCreateResponse` objects with `tag.id` AND `object.id`, allowing proper per-voucher tag resolution
- Removed `fetch_tags_for_voucher()` (the buggy per-voucher function)
- `fetch_tags_for_all_vouchers()` now uses two API calls: `GET /Tag` (id→name map) + `GET /TagRelation` (exact assignments)
- Performance improvement: 2 API calls total instead of N (one per voucher)

## Running the Script

```bash
python sevdesk_vouchers_v0.1.31.py \
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

The entire application lives in a single file: `sevdesk_vouchers_v0.1.31.py`.

### Function Map

| Function | Purpose |
|---|---|
| `parse_date()` | Parse dates in DD.MM.YYYY or YYYY-MM-DD format |
| `convert_date_to_dir_format()` | Convert dates for directory names |
| `fetch_vouchers()` | Fetch vouchers from sevDesk API with creditDebit filtering |
| `fetch_tags_for_all_vouchers()` | Batch fetch tags via TagRelation API (v0.1.31 fix) |
| `format_currency()` | Format amount with currency string |
| `sanitize_filename()` | Clean filenames of invalid characters |
| `tag_vouchers_in_sevdesk()` | Create tags on vouchers with creditDebit validation (v0.1.30 fix) |
| `download_voucher_pdf()` | Download a single voucher PDF (base64 or binary) |
| `download_pdfs_for_vouchers()` | Batch PDF downloading |
| `create_xlsx_sheet()` | Format and populate an Excel worksheet |
| `create_output_directory()` | Create output directory structure |
| `main()` | Orchestrate the full workflow |

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

#### `GET /Tag` — Fetch all tags (name registry)
- **Official params** (per OpenAPI spec): `id` (number), `name` (string)
- **Response**: `{ "objects": [ { "id", "name", "objectName": "Tag", ... } ] }`
- **WARNING**: `objectName` and `objectId` are NOT official parameters. Using them does NOT filter by specific voucher — it returns all tags globally. This was the root cause of the v0.1.31 bug.

#### `GET /TagRelation` — Fetch tag-to-object assignments (v0.1.31)
- **Params**: none (returns all relations)
- **Response**: `{ "objects": [ { "tag": {"id", "objectName": "Tag"}, "object": {"id", "objectName": "Voucher"} } ] }`
- Used in v0.1.31 to correctly determine which tags are assigned to which specific vouchers

#### `POST /Tag/Factory/create` — Create tag on voucher
- **Body**: `{ "name": "TAG_NAME", "object": { "id": voucherId, "objectName": "Voucher" } }`
- **Response**: Returns a `TagRelation` object (not just a Tag) linking the tag to the voucher
- Must validate `creditDebit` before calling (v0.1.30 fix)

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
- **v0.1.31**: Fixed expense export showing 0 Volltreffer due to buggy `GET /Tag` endpoint; replaced per-voucher `GET /Tag?objectId=` (undocumented, broken filter) with `GET /TagRelation` (exact voucher↔tag mapping); removed `fetch_tags_for_voucher()` function; also a performance improvement (2 API calls instead of N)

## Development Notes

- **No tests, CI/CD, linting, or formatting tools** are configured
- **No package management** (no `requirements.txt`, `pyproject.toml`, or `setup.py`)
- **Cross-platform**: includes Windows UTF-8 console encoding workaround
- **SSL**: verification can be disabled via `--no-verify` for development/debugging
- When modifying the script, update `__version__` and the filename to match
- The `creditDebit` mapping is critical and unintuitive: `D` (Debit) = income/Ausgangsrechnungen, `C` (Credit) = expense/Eingangsrechnungen

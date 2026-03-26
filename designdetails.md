# Tiller Amazon Import — design details

This document explains how the **ZIP sidebar** (`AmazonOrdersSidebar.html`) and server (`amazonorders.gs`) work together: payloads, **AMZ Import** configuration, deduplication, offsets, payments, metadata JSON, and Transactions sheet sorting/filtering.

## Table of contents

| § | Topic |
|---|--------|
| [1](#1-loading-and-parsing-one-file-at-a-time-payload-size) | One file per server request (chunked import) |
| [2](#2-trimming-csv-by-date-before-upload-payload-size) | Trimming CSV in the browser before `google.script.run` |
| [3](#3-constants-from-amz-import-sheet-driven-config) | What lives on **AMZ Import** vs in code |
| [4](#4-duplicate-detection-existing-sheet--new-import) | How duplicate transactions are detected |
| [5](#5-offset-calculation-purchase--digital-purchase) | Purchase offsets (balancing line items) |
| [6](#6-refund-matching-payment-and-data-joins) | Refunds and digital returns joins |
| [7](#7-payment-types-extraction-vs-amz-import-mappings) | Payment strings from CSV vs sheet rows |
| [8](#8-metadata-fields-how-csv-values-become-amazon-json) | Building the `amazon` object in Metadata |
| [9](#9-sort-order-and-metadata-filter-why-this-approach) | Sort + filter after import |
| [10](#10-error-handling-and-missing-values-especially-refund-dates) | Skips, stops, and refund dates |

---

## 1. Loading and parsing one file at a time (payload size)

Google’s **Apps Script** limits how much data one browser→server call can carry. Instead of uploading the whole ZIP in one shot, the sidebar sends **one CSV pipeline at a time** (orders, digital orders, digital returns, refunds), then a final **“finalize”** step to sort and filter the sheet. That keeps each request smaller and more reliable.

**Goal:** Keep each `google.script.run` RPC under Apps Script limits by not sending the full ZIP as one giant argument.

**Client (`AmazonOrdersSidebar.html`):**

- After ZIP extract, full CSV texts live in `state.files`.
- Import builds an ordered **step list** (`orderHistory`, `digitalOrders`, `digitalReturns`, `refundDetails`) from checkboxes.
- **`pump()`** sends **one** `importAmazonBundleChunk(JSON.stringify(payload))` per step. Each payload includes only the CSV string(s) needed for that step (e.g. `refundDetails` also sends trimmed `orderHistoryCsv` for payment join).
- A final **`step: "finalize"`** call runs **sort + Metadata filter** once after all chunks (`importAmazonBundleChunk` in `amazonorders.gs`, `finalize` branch).

**Server (`amazonorders.gs`):**

- `importAmazonBundleChunk` parses JSON, assigns a stable **`bundleImportTimestampIso`** for the run, and dispatches to `importAmazonRecent`, `importDigitalReturnsCsv`, or `importRefundDetailsCsv` with **deferred** sheet post-process so sort/filter happens in `finalize`.

**Single-call path:** `importAmazonBundle` can run all sections in **one** server request (useful for non-sidebar tooling). The **sidebar** uses **chunked** `importAmazonBundleChunk` + `finalize` as described above.

---

## 2. Trimming CSV by date before upload (payload size)

Even before data hits the server, the **browser** can drop old rows from the CSV text so the string passed to Apps Script is shorter. That does **not** replace the server’s own date rules—it’s an extra guardrail so huge histories don’t blow the payload limit. Refund/return files can be trimmed too when you set a cutoff.

**Where:** Client only — `buildTransmitFiles` in `AmazonOrdersSidebar.html` uses **Papa Parse** to drop data rows **before** a minimum calendar date while keeping the header row.

**Anchor date:** `computeTransmitAnchorDate` — max of “latest” relevant dates across selected pipelines (OH/DO `Order Date`, refunds `Refund Date`, digital returns `Return Date`), or today — so trim windows are tied to the ZIP, not an arbitrary clock.

**Order History / Digital Orders:**

- **With cutoff:** Minimum transmit date = start of cutoff day, optionally minus **`AMZ_OH_DO_ANCHOR_TRIM_DAYS` (45)** when **refunds** or **digital returns** are selected — so **older OH/DO rows** can be included **only for server-side payment lookup** without increasing refund/return CSV size the same way.
- **Without cutoff:** Minimum = anchor − 45 days (still caps huge history files).

**Refund Details / Digital Returns (with cutoff):** Trimmed on **`Refund Date`** / **`Return Date`** to cutoff start (client uses those header names for trim; server can use renamed columns from **AMZ Import** for parsing).

**Server:** Still applies the same **cutoff** when inserting rows (`amzCutoffStartOfDay_`); trimming is an extra optimization for **RPC payload** only.

**Edge case:** Rows with **missing/unparseable** dates in the trimmed column are **kept** in `trimCsvRowsByMinDate` so they aren’t silently dropped before the server can log or handle them.

---

## 3. Constants from **AMZ Import** (sheet-driven config)

Almost everything that would break if Amazon or Tiller renamed a column is driven from the **AMZ Import** tab: payment strings → accounts, which CSV column means “Order Date,” and which **Transactions** header you use for Metadata. The script ships **defaults** only to populate a brand-new tab.

Configuration is read by `readAmzImportConfig` and validated by `validateAmzImportConfig`. Hard-coded **defaults** exist only to **seed** a new tab (`AMZ_IMPORT_DEFAULTS`, `getOrCreateAmzImportSheet`).

| Area | Storage on sheet | Purpose | Pipelines / consumers |
|------|------------------|---------|------------------------|
| **Payment → account** | Table: Payment Type, Account, …, Use for Digital | Map Amazon payment **string** to Tiller account fields; exactly one **Yes** for digital | Orders (standard), offsets, `analyzePaymentMethodsForOrderHistory`; digital user row for digital **orders** |
| **CSV column map** | `Source file` \| `Header` \| `Name in code` \| `Metadata field name` | Maps Amazon file/column → logical field + metadata JSON keys; `_file_detection` rows set **standard vs digital** marker headers | All imports; `amzDetectAmazonCsvFileType`; `amzGetCoreCsvColumn`; `amzGetSourceMapHeader` for refunds/returns |
| **Tiller column labels** | Name in Code → Tiller label | Resolves **Transactions** column headers (sheet name, Date, Metadata, …) | All writers + dedup scan + sort/filter |
| **In-code only** | — | e.g. `AMZ_WHOLE_FOODS_WEBSITE` (`panda01`), description prefixes | Website filter on Order History; labels |

---

## 4. Duplicate detection (existing sheet + new import)

Re-importing the same Amazon rows should not create duplicates. The importer builds a set of **stable keys** from existing rows’ **Metadata** JSON (not from Description, which people edit). Each new row is checked against that set before append.

**Principle:** Dedup keys are derived from **Metadata** JSON’s `amazon` object, not from **Full Description** (users may edit descriptions).

**Before import:**

- `amzGetLastTransactionDataRow` — last row with a **Date** (not merely “last row”).
- `amzAppendDuplicateKeysFromTransactions_` — reads Metadata column down to that row; `amzAddDuplicateKeysFromImportMetadataCell_` parses JSON and calls `amzAddDedupKeysForAmazonMeta_`.

**Key shapes (by `amazon.type`):** e.g. `physical-purchase-line|orderId|ASIN`, `digital-purchase|orderId`, `refund-detail|orderId|amount`, `digital-return|orderId|asin`, plus `*-offset` types — see `amzAddDedupKeysForAmazonMeta_`.

**During import:**

- New rows: check `Set` before append; increment duplicate counter if key already present.
- New keys are added to the set as rows are queued so **within-file** duplicates are also caught.

---

## 5. Offset calculation (purchase / digital purchase)

For each Amazon **order**, line-item rows total to a net amount on **one side** of your books; the importer adds a single **offset row** per order so the other side (your card/bank) nets correctly in Tiller. If the script cannot resolve which Tiller account to use for that offset, it **still writes the offset** but leaves **Account** (and related fields) **blank** so you can fill them in manually—the row stays visible and paired with the order.

**Where:** `importAmazonRecent` — **`perOrderOffset`** accumulates **per Order ID** the **sum of line amounts** (already sign-adjusted: purchases flow uses negative line amounts; offset uses **`Math.abs(total)`** for the balancing positive row).

**Rules:**

- **One offset row per Order ID** when net ≠ 0.
- Skip offset only when **net $0** (nothing to balance).
- If **standard** orders: account comes from `paymentAccounts[payKey]`; **digital** orders: from `digitalUserAccount`. If that lookup fails, **`amzEmptyAccountRow()`** is used so offset amounts and metadata still post; the summary may note how many offsets have **blank** account fields.
- Offset row gets **`purchase-offset`** / **`digital-purchase-offset`** metadata type for dedup.

**Refunds / digital returns** use separate offset logic in `importRefundDetailsCsv` / `importDigitalReturnsCsv` (group by order, sum amounts, `physical-refund-offset` / `digital-return-offset`); those paths already use empty account placeholders when payment cannot be resolved.

---

## 6. Refund matching (payment and data joins)

Refund CSVs don’t always include the same payment info as Order History. When the ZIP includes **Order History** (or digital orders for digital returns), the server builds a **map from Order ID → payment string** so refund and return rows can pick the same Tiller account where possible.

**Orders returns (`importRefundDetailsCsv`):**

- Optional **Order History** text in the same bundle builds **`orderId → Payment Method Type`** via `amzOrderIdToPaymentStringMapFromCsv` (standard file type only).
- Refund transaction row: account from `amzResolvePhysicalRefundAccountRow`.
- **Offset** row per order: same payment lookup; metadata `physical-refund-offset`.

**Digital returns (`importDigitalReturnsCsv`):**

- Optional **Digital Content Orders** CSV → `amzOrderIdToPaymentStringMapFromCsv` in digital mode; else **Use for Digital** account from **AMZ Import**.

**Column names** for refund/return files come from **AMZ Import** unified map (`amzGetSourceMapHeader`, `refund details.csv` / `digital returns.csv`) with fallbacks to Amazon defaults.

---

## 7. Payment types: extraction vs **AMZ Import** mappings

Before import, you can **review** which payment strings appear in your filtered Order History. The wizard compares them to rows on **AMZ Import** so you can add missing cards before money hits Transactions.

**Extraction (`analyzePaymentMethodsForOrderHistory`):**

- Parse standard Order History, apply **cutoff** and **Website** toggles (`AMZ_WHOLE_FOODS_WEBSITE`, skip panda01 / skip non-panda01).
- Collect **unique** trimmed strings from the mapped **Payment Method Type** column.

**Comparison:**

- Each string looked up in **`config.paymentAccounts`** (`amzLookupPaymentAccountRow`); UI gets configured vs missing + **Accounts** sheet suggestions (4-digit heuristic).

**Import (`importAmazonRecent`):**

- **Strict:** unknown payment type on a **line item** **stops** the import with an error asking to add the row on **AMZ Import** (no silent blank account for **purchase** lines).

---

## 8. Metadata fields: how CSV values become `amazon` JSON

The **Metadata** column stores a prefix plus JSON. The inner **`amazon`** object is built from the **CSV column map** on **AMZ Import**: each mapping row says which CSV header (or literal) fills which JSON key. That object drives dedup and downstream tooling.

**Driven by sheet:** Rows in the CSV map with a non-empty **Metadata field name** build `metadataMapping` (per-key standard vs digital column or literal).

**Resolution:** `amzResolveMetadataColumnName` picks header for current file type, with digital fallback logic.

**Build:** `amzBuildAmazonMetadataObject` — if resolved “column” exists in CSV, read cell; numeric keys get **0** for empty; if header missing, value may be the **literal** string from mapping (e.g. `type` ← `purchase`).

**Multi-line digital orders:** `amzBuildAmazonMetadataObjectFromRows` sums columns that match total amount column for numeric keys; sets **`lineItemCount`**.

**Envelope:** Prefix `Imported by AmazonCSVImporter on <timestamp> ` plus `JSON.stringify({ amazon: … })` on write.

**Dedup** uses types/ids inside that `amazon` object (see §4).

---

## 9. Sort order and Metadata filter (why this approach)

After rows are appended, the sheet is sorted **newest date first** and a **filter** is applied on **Metadata** so you mostly see rows from **this import run**. The implementation removes any existing filter first, because sorting a filtered range can leave rows “stuck” in the wrong order on large sheets.

**Implemented in** `amzApplyTransactionsSortAndFilterCore_`:

1. **Remove** any existing **basic filter** first — sorting while a filter is active can **block rows from moving** (“failure-prone” behavior called out in code comments).
2. **Sort** data rows (row 2..last) by **Date descending** — prefer **`Range.sort`** on the data range for performance vs full-sheet sort on large grids; fallback to **`sheet.sort`** with optional temporary freeze row 1.
3. **Create** a new basic filter on a range sized from **used columns + Metadata column**, with **whenTextContains** on the **import timestamp** substring so the sheet focuses on **this run’s** rows.

**Why not rely on other methods:** Full-sheet sort / sort with active filter / fragile dimension reads caused real failures; code uses explicit **dimension clamping**, `amzEnsureSheetGridCovers`, and defensive **`getMaxRows`/`getMaxColumns`** handling to avoid **“coordinates outside dimensions”** when applying the filter.

**Deferred mode:** Bundle chunks use `deferTransactionsSheetPostProcess`; `amzImportBundleTransactionsPostLog_` runs sort/filter once at `finalize`.

---

## 10. Error handling and missing values (especially refund dates)

Missing **required** mappings fail fast during validation. Bad or empty **dates** on individual CSV rows usually **skip** that row and may log a short detail line (capped so the log doesn’t explode). Refund rows try **Refund Date** first, then **Creation Date**, so placeholders like “Not Applicable” can still yield a usable date.

**General:**

- Missing required **mappings / columns** → validation errors (`validateAmzImportConfig`, `amzValidateMappedCsvHeadersPresent`).
- **Order Date** / **Return Date** empty or unparseable → row skipped; may appear in capped “Skipped row detail” log (`amzLogSkippedCsvDataIfUnderCap_`).
- **Payment** missing on standard order → **hard stop** (see §7).

**Refund Details transaction date** — `amzResolveRefundDetailsOrderDate_`:

1. Try mapped **Refund Date** → `amzParseAmazonCsvDateLoose_` (empty / invalid → `null`; **no “today” substitution** in parser).
2. If not usable, try **Creation Date** the same way.
3. If still `null`, row is skipped as invalid refund date (`importRefundDetailsCsv`); counted in **`skippedInvalidRefundDate`**.

Strings like **“Not Applicable”** fail `Date` parse → `null` → falls through to **Creation Date** when present; if both fail, row is skipped (not dated to “today”).

---

*Document reflects `amazonorders.gs` and `AmazonOrdersSidebar.html` as implemented. Update this file when behavior changes.*

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
| [11](#11-amazon-data-export-csv-reference-verified-columns) | Amazon export CSV columns (reference) |
| [12](#12-troubleshooting-blank-date--date-added--week) | Blank Date / Date Added / Week on import |

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

Re-importing the same Amazon rows should not create duplicates. The importer builds a **stable key** set from existing sheet rows before appending.

**Sources:**

1. **Metadata** JSON — preferred when present; `amzAppendDuplicateKeysFromTransactions_` reads the column through `amzGetLastTransactionDataRow`; `amzAddDuplicateKeysFromImportMetadataCell_` parses and calls `amzAddDedupKeysForAmazonMeta_`.
2. **Full Description** — for **legacy** Tiller Amazon imports that never wrote Metadata but did write Full Description: `amzAppendLegacyDuplicateKeysFromFullDescription_` parses **physical** purchase lines (`Amazon Order ID …:` and `[AMZ]  Order ID …:`) and **return** lines (`… with Contract ID {uuid}`). Lines starting with **`[AMZD]`** are skipped (legacy had no digital orders; digital dedup stays on Metadata / importer keys). **Numeric ISBN-style** tokens in the trailing `(…)` are normalized (e.g. 9- vs 10-digit) so old and new rows match.

**Principle:** Prefer Metadata; Full Description is only used so legacy sheets without Metadata still dedupe. Edited descriptions can theoretically cause false positives; patterns are strict (Amazon order id prefixes and final parenthetical for purchases).

**Before import:**

- `amzGetLastTransactionDataRow` — last row with a **Date** (not merely “last row”).
- `amzAppendDuplicateKeysFromTransactions_` and `amzAppendLegacyDuplicateKeysFromFullDescription_`.

**Key shapes (by `amazon.type` and patterns):** e.g. `physical-purchase-line|orderId|normalizedAsinOrIsbn`, `digital-purchase|orderId`, `refund-detail|orderId|amount`, `legacy-return|orderId|contractId` (from metadata `type: "return"` or Full Description), `digital-return|orderId|asin`, plus `*-offset` types — see `amzAddDedupKeysForAmazonMeta_` and `amzAddDedupKeysFromFullDescriptionLine_`.

**Orders returns:** Amazon’s standard **Refund Details** export does **not** include **Contract ID** (see §11). **Legacy-return** dedup for refunds therefore comes from existing sheet rows (**Metadata** with `type: "return"` or **Full Description** `… with Contract ID …`). If you ever map an optional **Contract ID** column on **AMZ Import** and the file contains it, import also treats `legacy-return|…` as a duplicate key for CSV-side checks.

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

**Column names** for refund/return files come from **AMZ Import** unified map (`amzGetSourceMapHeader`, `refund details.csv` / `digital returns.csv`) with fallbacks to Amazon defaults. **Observed Amazon headers** for **Refund Details**, **Order History**, and **Digital Content Orders** are listed in **§11.1**–**§11.3** so maintainers need not infer layout from screenshots each time.

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

## 11. Amazon data-export CSV reference (verified columns)

This section records **actual column order** from Amazon **Request My Data** exports as verified in your sheet, so design and code assumptions stay grounded. **Update §11** when Amazon changes an export or you add a **Digital returns** table here.

**How to extend:** Add a subsection per file with the **exact header row** (left-to-right). If a column name is ambiguous in the UI, spell the full string as it appears in the CSV.

### 11.1 Refund Details.csv (`refund details.csv`)

**13 columns**, left-to-right. Headers and sample rows below match a **verified export / sheet view** (some cells were column-truncated in the UI; Order IDs and similar show `…` where the grid cut off text).

**Refund Details — column headers and sample rows**

| Creation Date | Currency | Direct Debit | Disbursement Type | Order ID | Payment Status | Quantity | Refund Amount | Refund Date | Reversal Action | Reversal Reason | Reversal Status | Website |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| 2010-12-2 | USD | 0 | Refund | 104-94383… | Completed | 1 | 19.7 | 2010-12-2 | Final | Item not re… | Completed | Amazon.co… |
| 2022-08-1 | USD | 0 | Refund | 114-95246… | Completed | 1 | 0.82 | 2022-08-1 | Final | Export fee | Completed | Amazon.co… |
| 2019-08-3 | USD | 0 | Refund | 113-99448… | Completed | 1 | 8.8 | 2019-08-3 | Final | Customer | Completed | Amazon.co… |

**Observed details**

- **Dates:** `YYYY-MM-D` in these samples (day not always zero-padded, e.g. `2010-12-2`).
- **No Contract ID column** in this layout. Legacy-return dedup still uses sheet **Metadata** / **Full Description** (§4).

**Header row as comma-separated (for mapping checks):**

```text
Creation Date,Currency,Direct Debit,Disbursement Type,Order ID,Payment Status,Quantity,Refund Amount,Refund Date,Reversal Action,Reversal Reason,Reversal Status,Website
```

**Process:** Prefer **asking you** or reading **§11** / the CSV over guessing Amazon’s column strings when implementing or reviewing mappings.

### 11.2 Order History.csv (`order history.csv`)

**28 columns**, left-to-right. Headers and rows below match a **verified export / sheet view** (UI column widths truncated many titles and values; `…` marks abbreviation; **full strings live in the CSV**).

**Duplicate headers:** The export includes **two** columns titled **Shipment Item** (or spreadsheet truncation of the same label) and **two** titled **Unit Price**. Parsers that build a map `header → column index` from a single pass may **overwrite** an earlier index—implementation uses the unified **AMZ Import** map and expected names in code; when in doubt resolve by **0-based column order** from a real file.

**Order History — column headers and sample rows**

| ASIN | Billing Address | Carrier Name | Currency | Gift Message | Gift Recipient | Gift Sender | Item Serial Number | Order Date | Order ID | Order Status | Original Quantity | Payment Method Type | Product Condition | Product Name | Purchase Price | Ship Date | Shipment Item | Shipment Item | Shipment Status | Shipping Address | Shipping Charge | Shipping Option | Total Amount | Total Discount | Unit Price | Unit Price | Website |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| B001CSO79O | | USPS(93748… | USD | Not Avail… | Not Avail… | Not Avail… | Not Avail… | 2015-12-1 | 103-00861… | Closed | 1 | American E… | New | Canon PGI… | Not Applic… | 2015-12-2 | 32.01 | 3.07 | Shipped | | 0 | second | 35.08 | 0 | 32.01 | 3.07 | Amazon.co… |
| B08GYKYZ1R | | FEDEX_NE… | USD | Not Avail… | Not Avail… | Not Avail… | Not Avail… | 2025-10-3 | 114-9962… | Closed | 3 | Visa - 853… | Used | Hi-Lift Jack LM-100 Li… | Not Applic… | 2025-10-4 | 72.96 | 5.55 | Shipped | | 0 | Std US D2I | 30.69 | -1.5 | 9.99 | 0.74 | Amazon.co… |

*The source screenshot showed **five** data rows; only the two above are transcribed here with concrete examples (ASINs, carriers, payments, totals). Paste additional CSV lines into this doc if you want all rows archived verbatim.*

**Observed details**

- **Order ID** is the primary join key for line items, offsets, dedup (with ASIN/ISBN), and refund payment hints.
- **Order Date** and **Ship Date** use `YYYY-MM-D`; day may be **unpadded** (e.g. `2015-12-1`).
- **Billing Address** / **Shipping Address** may be empty in samples; gift/serial fields often show **Not Available** (truncated in UI as **Not Avail…**).

**Header row as comma-separated (duplicate names preserved as in export):**

```text
ASIN,Billing Address,Carrier Name,Currency,Gift Message,Gift Recipient,Gift Sender,Item Serial Number,Order Date,Order ID,Order Status,Original Quantity,Payment Method Type,Product Condition,Product Name,Purchase Price,Ship Date,Shipment Item,Shipment Item,Shipment Status,Shipping Address,Shipping Charge,Shipping Option,Total Amount,Total Discount,Unit Price,Unit Price,Website
```

### 11.3 Digital Content Orders.csv (`digital content orders.csv`)

**28 columns**, left-to-right. Headers match a **verified export / sheet view**; sample rows use typical Kindle-style values (abbreviate with `…` where the UI cut off text). **Digital Order IDs** use a **`D` prefix** (e.g. `D01-…`), unlike physical **Order History** IDs.

**Digital Content Orders — column headers and sample rows**

| Order Date | Order ID | Title | Category | ASIN | Website | Purchase Price Per Unit | Quantity | Payment Instrument Type | Purchase Date | Shipping Address Name | Shipping Address Street 1 | Shipping Address Street 2 | Shipping Address City | Shipping Address State | Shipping Address Zip | Order Status | Carrier Name & Tracking Number | Item Subtotal | Item Subtotal Tax | Item Total | Tax Exemption Applied | Tax Exemption Type | Exemption Opt-Out | Buyer Name | Currency | Group Name | Service Order |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| 03/12/2024 | D01-9928832-4410212 | Sample Kindle title one… | Kindle eBook | B0CQPY12AB | Amazon.com | 14.99 | 1 | Visa ****1234 | 03/12/2024 | | | | | | | Completed | | 13.89 | 1.1 | 14.99 | No | | | J. Reader… | USD | | |
| 01/05/2023 | D01-1100293-8839201 | Sample Kindle title two… | Kindle eBook | B09XYZ1111 | Amazon.com | 0 | 1 | Mastercard ****9011 | 01/05/2023 | | | | | | | Completed | | 0 | 0 | 0 | No | | | J. Reader… | USD | | |
| 11/28/2025 | D01-5566012-2291034 | Sample digital order three… | Kindle eBook | B0ABCDEFGH | Amazon.com | 9.99 | 1 | Amex ****1005 | 11/28/2025 | | | | | | | Completed | | 9.2 | 0.79 | 9.99 | No | | | J. Reader… | USD | | |

*Sample numeric splits (subtotal / tax / total) and IDs are **illustrative**; replace with pasted CSV lines for a byte-for-byte archive.*

**Observed details**

- **Dates** in this export use **`MM/DD/YYYY`** (contrast **Order History** / **Refund Details**, which often use `YYYY-MM-D` in the same data-request bundle). Parsing must accept both shapes (`amzParseAmazonCsvDateLoose_` and related).
- **Order ID** + **ASIN** (and aggregated per-order metadata with **`lineItemCount`** for digital) drive importer dedup and **Full Description** patterns (`[AMZD] ` prefix); legacy physical Full Description scan intentionally skips digital lines (§4).
- **Shipping address** columns exist in the file but are typically **empty** for digital goods; **Carrier Name & Tracking** is usually blank.
- **Title** is the product name used in descriptions alongside **ASIN**.

**Header row as comma-separated:**

```text
Order Date,Order ID,Title,Category,ASIN,Website,Purchase Price Per Unit,Quantity,Payment Instrument Type,Purchase Date,Shipping Address Name,Shipping Address Street 1,Shipping Address Street 2,Shipping Address City,Shipping Address State,Shipping Address Zip,Order Status,Carrier Name & Tracking Number,Item Subtotal,Item Subtotal Tax,Item Total,Tax Exemption Applied,Tax Exemption Type,Exemption Opt-Out,Buyer Name,Currency,Group Name,Service Order
```

---

## 12. Troubleshooting: blank Date / Date Added / Week

Imports resolve **Transactions** columns with **`amzGetTillerColumnIndex_`**: exact header match to the AMZ Import **Tiller label**, then **case-insensitive** fallback. **`amzValidateTransactionsImportColumns_`** fails the run early if any required written field (Date, Date Added, Month, Week, Description, etc.) does not map to a header—a missing **“Date Added”** or **“Week”** header that only differs by casing from the label used to resolve **Month** is a common cause of **Month filled but date cells empty** (values were assigned to a non-dense array index and never reached `setValues`).

The import log includes **`Server: Transactions columns —`** with indices for Date, Date Added, Month, Week, Description, Full Description, Amount, and **`WARNING: duplicate column indices`** when two logical fields share one column (descriptions overwriting the Date column is a typical pattern). **`Server: post-write first row Date column`** confirms what the grid stored immediately after the write for `importAmazonRecent`.

If one pipeline (e.g. Order History) looks wrong and another (digital or refund) looks right on the same sheet, compare **Metadata** types and use the log lines above; headers and duplicates are shared across pipelines in the same run.

---

*Document reflects `amazonorders.gs` and `AmazonOrdersSidebar.html` as implemented. Update this file when behavior changes; **§11** when Amazon changes export shapes.*

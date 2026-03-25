# Tiller Tools

One GitHub repo and Apps Script project for **Tiller** (https://tiller.com/) spreadsheets:

- **Tiller Quick Search** — sidebar to filter your **Transactions** sheet by date, amount, description, account, and category (basic filter + helper columns).
- **Tiller Amazon Import** — sidebar wizard to import Amazon order CSVs from your data-export ZIP into **Transactions** (with **AMZ Import** configuration). Large exports are sent to Apps Script in **separate chunks** (one main CSV per round trip, plus a final step that sorts/filters Transactions) so big ZIPs stay reliable.

Repository: [github.com/daveinlosbarriles/TillerTools](https://github.com/daveinlosbarriles/TillerTools). Use **clasp** from this folder to push everything, or copy individual files into an Apps Script project as described below.

This software is produced by a Tiller user and is not affiliated with Tiller LLC.

## Privacy, terms, and contact

| Document | Link |
|----------|------|
| Privacy Policy | [PRIVACY.md](PRIVACY.md) · [on GitHub](https://github.com/daveinlosbarriles/TillerTools/blob/master/PRIVACY.md) |
| Terms of Service | [TERMS.md](TERMS.md) · [on GitHub](https://github.com/daveinlosbarriles/TillerTools/blob/master/TERMS.md) |

Questions: **tillertoolsbydave@gmail.com**

## Google Workspace add-on

This project is set up as a **Google Sheets Editor add-on**:

- **[`appsscript.json`](appsscript.json)** — `addOns.common` (name **Tiller Tools**, `logoUrl` → logo in `assets/` via GitHub raw), Sheets host, and **`tillerToolsOnHomepage`** for the side-panel welcome card.
- **[`Code.js`](Code.js)** — **`onInstall`** / **`onOpen`**, **`createMenu("Tiller Tools")`** with **Tiller Amazon Import** and **Tiller Quick Search** (works for sheet-bound scripts and for many add‑on installs; after install, also check **Extensions**).
- **OAuth scopes** (declare the same on the Google Cloud **OAuth consent screen**):
  - `https://www.googleapis.com/auth/spreadsheets.currentonly`
  - `https://www.googleapis.com/auth/script.container.ui`

Distribute via **Deploy → Test deployments** / **Google Workspace Marketplace**, link a **standard GCP project**, and add **test users** while the OAuth app is in *Testing*.

## clasp notes

- **[`.claspignore`](.claspignore)** — ignores a local **`TillerAmazonOrdersCSVImport/`** folder if you still have an old clone next to this repo (so it is not pushed to Apps Script).

---

## Tiller Quick Search — key features

**Search by:**
- **Date** – range or exact From / To
- **Amount** – min and max
- **Description** – “contains” text, with options like:
  - `Amazon | Walmart` – rows containing Amazon or Walmart
  - `^Vanguard` – name starts with Vanguard
  - `1234$` – name ends with 1234
  - `^Gas$` – exactly “Gas” (no “Gasoline”)
  - `Amazon.*Return|Return.*Amazon` – contains both Amazon and Return (any order)
  - `gas but not chevron` – contains “gas” but not “chevron”
- **Account** – choose one or more
- **Category** – one or more categories, sorted by group
- See all of your search criteria at once in the sidebar

<img width="3840" height="2063" alt="Tiller Tools screenshot" src="https://github.com/user-attachments/assets/80d36566-54d4-40f0-b86d-35040a49a8e1" />

---

## What you need

- A Google Sheet that uses **Tiller** (or has the same structure), with at least these sheets:
  - **Transactions** – with columns such as Date, Description, Category, Amount, Account
  - **Categories** – used for the category list in the sidebar
  - **Accounts** – used for the account list in the sidebar

---

## Installation

### Step 1: Open the Apps Script editor and add the script files

1. Open your Tiller Google Sheet.
2. In the menu, click **Extensions** → **Apps Script**.  
   A new tab opens with the Apps Script editor (code view).
3. If you see a default file like `Code.gs` with some sample code, you can replace it or add new files. For **both** Quick Search and Amazon import, add at least: **Code.gs**, **QuickSearchSidebar.gs**, **QuickSearch.html**, **amazonorders.gs**, **AmazonOrdersSidebar.html** (and optionally `AmazonOrdersDialog.html`). For **Quick Search only**, you can omit the Amazon files and remove the Amazon menu line from **Code.js**.

   **Creating or replacing files:**
   - **Code.gs** — Add **onInstall** (calling **onOpen**) and **onOpen** from **Code.js**, or paste the whole **Code.js** as **Code.gs**. The file also defines **`tillerToolsOnHomepage`** for the Workspace add-on card; merge carefully if you already have an **onOpen** handler. Apps Script uses the **.gs** extension on disk.
   - **QuickSearchSidebar.gs**  
     - Click **+** → **Script**, name it `QuickSearchSidebar`, then paste the contents of **QuickSearchSidebar.js** from this repo. It will appear as **QuickSearchSidebar.gs**.
   - **QuickSearch.html**  
     - Click **+** → **HTML**, name it `QuickSearch`, then paste the contents of **QuickSearch.html** from this repo. Leave the name as **QuickSearch** (no “.html” in the file list is fine).
   - **Amazon import (optional):** Add **amazonorders.gs** and **AmazonOrdersSidebar.html** the same way (script + HTML). Match the file names the script expects (`HtmlService.createHtmlOutputFromFile("AmazonOrdersSidebar")`).

4. Save everything: **File** → **Save** (or Ctrl+S). Give the project a name (e.g. “Tiller Tools”) if prompted.

   **Using clasp:** Clone this repo, run `clasp login` and `clasp clone` / link your Apps Script project, then `clasp push` so all `.gs` / `.html` files and `appsscript.json` stay in sync with GitHub.

5. **Permissions (first run):**  
   The first time you use the script (e.g. reload the sheet and open **Tiller Tools** → **Tiller Quick Search** or **Tiller Amazon Import**), Google will ask you to authorize the app:
   - Click **Review permissions**, choose your account, then **Advanced** → **Go to [project name] (unsafe)** (this is your own script).
   - Click **Allow**.  
   This lets the script read and filter your sheet. No data is sent to anyone else.

---

### Step 2: Turn on the Google Sheets API (v4)

Quick Search needs the **Sheets API** to work properly.

1. In the Apps Script editor (same tab as your code), click **Services** in the left sidebar (the “+” / add services icon).  
   - If you don’t see “Services”, try **Resources** → **Advanced Google Services** (older editor) or **Services** (new editor).
2. Find **Google Sheets API** (or “Sheets API”) in the list and turn it **On**.
3. If you’re in **Advanced Google Services**, also check the link at the bottom: **Google Cloud Console**. Click it and make sure the **Google Sheets API** is enabled for the same project. Then return to Apps Script.

Once this is on, you don’t need to do it again.

---

### Step 3: Open the tools

1. Go back to your Google Sheet tab and **reload the page** (F5 or refresh).
2. Open the **Tiller Tools** menu (**top menu bar**, or **Extensions** when using an installed Workspace add-on).
3. Choose:
   - **Tiller Quick Search** — sidebar for filters; click **Search** when ready.
   - **Tiller Amazon Import** — ZIP wizard; follow the steps to pick your Amazon export, categories, and import.

---

## First-time use: two new columns

The **first time** you run a search (or the first time the script needs the helper columns), it will add **two columns** to the right of your **Transactions** sheet:

- **QuickSearch** – shows TRUE/FALSE for each row (whether the row matches your current criteria).
- **QuickCriteria** – a hidden-style cell that stores the current criteria so the filter can work.

These columns are **normal sheet columns**. You can:

- **Delete them** anytime (e.g. to tidy the sheet). The next time you click **Search** in the sidebar, the script will add them again at the end of the sheet.
- **Leave them in place** – then each time you click **Search**, only the criteria and the TRUE/FALSE values update; the columns stay where they are.

You don’t need to edit these columns yourself; the script fills them in.

---

## How the filter works (basic filter)

Quick Search applies a **basic filter** to your **Transactions** sheet (the same kind you get from **Data** → **Create a filter**). It sets the filter so that only rows where **QuickSearch** is TRUE are shown.

Because it’s a normal basic filter:

- You can **add more criteria** using the filter dropdowns in the header row (e.g. filter by another column).
- You can **change or clear** the filter using the sheet’s filter icon and dropdowns as usual.
- Quick Search only controls the criteria that affect the **QuickSearch** column; the rest of the filter behavior is the same as any other filtered sheet.

---

## Where the lists come from (dependencies)

- **Category list** in the sidebar is read from your **Categories** sheet: the first column (Category) and second column (Group). The script uses data starting at row 2. A “(Blank)” option is added so you can search for transactions with no category.
- **Account list** in the sidebar is read from your **Accounts** sheet: the account names from column **J**, starting at row 2. Any row where column **Q** is **Hide** is skipped, so those accounts don’t appear in the list.

Your Transactions sheet must have columns that match what Tiller uses (e.g. Date, Description, Category, Amount, Account). The script finds them by the header names in row 1.

---

## Repo contents (for reference)

| File in repo   | In Apps Script   | Purpose |
|----------------|------------------|--------|
| Code.js        | Code.gs          | `onInstall` / `onOpen`, **Tiller Tools** menu → **Tiller Amazon Import**, **Tiller Quick Search**; Card homepage `tillerToolsOnHomepage` |
| QuickSearchSidebar.js | QuickSearchSidebar.gs | All Quick Search logic (criteria, filter, helper columns) |
| QuickSearch.html      | QuickSearch.html  | Sidebar UI |
| amazonorders.gs | amazonorders.gs | Amazon CSV import pipelines, **`importAmazonBundleChunk`**, finalize sort/filter |
| AmazonOrdersSidebar.html | AmazonOrdersSidebar.html | Amazon import **sidebar** (ZIP, JSZip in browser, chunked `google.script.run`) |
| AmazonOrdersDialog.html | AmazonOrdersDialog.html | Legacy modal UI (unused if menu opens the sidebar) |
| appsscript.json       | (project config) | Time zone, scopes, Sheets advanced service, **`addOns`** for Workspace listing |
| assets/tiller-tools-logo.png | — | Add-on icon; referenced by `logoUrl` (GitHub raw URL) |
| PRIVACY.md / TERMS.md | — | Privacy and terms (not deployed by clasp) |

### Architecture

**Quick Search** and **Amazon import** do not call each other; they share only the menu in **Code.js**. Each has its own header→column map (`getTillerColumnMap` vs `amzGetTillerColumnMap`). To ship **one** feature only, delete the other’s files and drop its menu item(s) from **Code.js**.

### Source of truth

All source for both tools is in **[TillerTools](https://github.com/daveinlosbarriles/TillerTools)**. Commit and push here; there is no separate mirror repo.

---

## Usage tips

- **Quick Search:** **Tiller Tools** → **Tiller Quick Search**, set date range, amount, description, account, and/or category, then click **Search**.
- **Amazon import:** **Tiller Tools** → **Tiller Amazon Import**, choose your orders ZIP, select which CSV pipelines to run, then complete payment-method review if prompted before **Import**.
- Use **Reset All** or the **×** next to each section to clear criteria.
- In **Description**, you can use plain text or simple **regex** (e.g. `Amazon|Walmart` for “Amazon or Walmart”). Use **` but not `** (with spaces) to require one phrase and exclude another (e.g. `gas but not chevron`). Click the **Description** label in the sidebar for examples.

---

## To Uninstall

1. **Remove the menu and script files:** Open **Extensions** → **Apps Script**. Remove the **onInstall** / **onOpen** code (or the Tiller Tools menu) from **Code.gs** so the menu no longer appears. Delete **QuickSearchSidebar.gs**, **QuickSearch.html**, and if used **amazonorders.gs** / **AmazonOrdersSidebar.html** / **AmazonOrdersDialog.html**.
2. **Optional – remove the helper columns:** On your **Transactions** sheet, delete the **QuickSearch** and **QuickCriteria** columns if they were added. You can delete them like any other columns (right‑click the column letter → Delete column).
3. **Optional – clear the filter:** If a filter is still applied to the Transactions sheet, turn it off via **Data** → **Turn off filter** (or use the filter icon in the toolbar).

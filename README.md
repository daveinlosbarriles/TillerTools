# Tiller Quick Search

A Google Apps Script sidebar for **Tiller** spreadsheets that adds a Quick Search panel. Filter your Transactions sheet by date, amount, description, account, and category. Quick Search uses your sheet’s normal basic filter and two helper columns (Match and Criteria) so you can combine its results with any other filter criteria you set directly on the sheet.

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
3. If you see a default file like `Code.gs` with some sample code, you can replace it or add new files. You will create three items in this project:
   - **Code.gs** – adds the “Tiller Tools” menu to your sheet  
   - **QuickSearchSidebar.gs** – all the Quick Search logic  
   - **QuickSearch.html** – the sidebar interface  

   **Creating or replacing files:**
   - **Code.gs**  
     - If **Code.gs** already exists: select it, delete its contents, and paste in the contents of the **Code.js** file from this repo.  
     - If it doesn’t exist: click the **+** next to “Files”, choose **Script**, name it `Code`, then paste the contents of **Code.js**.  
     - In Apps Script, script files are saved as **.gs** (not .js). The editor will show “Code.gs” once saved. That’s correct.
   - **QuickSearchSidebar.gs**  
     - Click **+** → **Script**, name it `QuickSearchSidebar`, then paste the contents of **QuickSearchSidebar.js** from this repo. It will appear as **QuickSearchSidebar.gs**.
   - **QuickSearch.html**  
     - Click **+** → **HTML**, name it `QuickSearch`, then paste the contents of **QuickSearch.html** from this repo. Leave the name as **QuickSearch** (no “.html” in the file list is fine).

4. Save everything: **File** → **Save** (or Ctrl+S). Give the project a name (e.g. “Tiller Quick Search”) if prompted.

5. **Permissions (first run):**  
   The first time you use the script (e.g. reload the sheet and open **Tiller Tools** → **Quick Search**), Google will ask you to authorize the app:
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

### Step 3: Use Quick Search

1. Go back to your Google Sheet tab and **reload the page** (F5 or refresh).
2. You should see **Tiller Tools** in the menu bar.
3. Click **Tiller Tools** → **Quick Search**.  
   The Quick Search sidebar opens on the right. Set your filters and click **Search**.

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
| Code.js        | Code.gs          | Adds “Tiller Tools” menu and “Quick Search” item |
| QuickSearchSidebar.js | QuickSearchSidebar.gs | All Quick Search logic (criteria, filter, helper columns) |
| QuickSearch.html      | QuickSearch.html  | Sidebar UI |
| appsscript.json       | (project config) | Script settings; usually auto-managed by Apps Script |

---

## Usage tips

- Open **Tiller Tools** → **Quick Search**, set date range, amount, description, account, and/or category, then click **Search**.
- Use **Reset All** or the **×** next to each section to clear criteria.
- In **Description**, you can use plain text or simple **regex** (e.g. `Amazon|Walmart` for “Amazon or Walmart”). Use **` but not `** (with spaces) to require one phrase and exclude another (e.g. `gas but not chevron`). Click the **Description** label in the sidebar for examples.

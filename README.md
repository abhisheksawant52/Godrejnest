# Godrej Nest · Community event ledger

A responsive React and Vite dashboard for Godrej Nest community event funds. Visitors can switch between events to see each event's totals, contributions, stall receipts, payments, and contribution trend. Event data is stored in separate JSON files under `src/events/` and bundled into the site at build time.

The dashboard includes a spoken “Welcome to Godrej Nest Community” greeting with a brief synthesized chime. It attempts to play when the page opens and provides a replay control. If browser autoplay is blocked, visitors can tap to hear it. The site prefers an available English voice identified as female; the available voices depend on the visitor's device and browser. The Godrej Nest wordmark is enlarged in the sidebar and footer.

## Run locally

Requirements: Node.js and npm. To import Excel workbooks, Python 3 is also required; the importer uses only Python's standard library.

```sh
npm install
npm run dev
```

## Build and preview

```sh
npm run build
npm run preview
```

The Vercel project configuration is in `vercel.json`. Deploy the built site after adding or updating event data. The live static site does not accept workbook uploads.

## Add an event from Excel

Each workbook must contain these worksheet tabs and headers:

- `Total Receipts`: `Name`, `Flat No`, `Amount`, `Date`, `Time`, `Remarks`
- `Stall Payment`: `Name`, `Flat No`, `Amount`, `Date`, `Time`, `Remarks`
- `Total Expenses`: `Date`, `Particular`, `Amount`, `Paid to`, `Remarks`

Import a workbook from PowerShell:

```powershell
npm run import:statement -- "C:\path\to\Diwali_2026.xlsx" diwali-2026 --name "Diwali 2026"
```

This creates `src/events/diwali-2026.json`. The event selector discovers JSON files in that folder automatically. Run a production build and deploy for the event to appear on the published site.

The event ID is a lowercase slug made of letters, numbers, and hyphens. Reusing an ID updates that event's JSON file. If omitted, the importer generates an ID from the workbook filename. The display name is optional and defaults to a title-cased version of the ID. The chart year is read from the display name or ID; if neither includes a year, the current year is used.

For example, to import or refresh the existing Navratri event, reuse its `navratri-2026` ID. Reusing the ID updates the existing event instead of creating a second event from the workbook filename:

```powershell
npm run import:statement -- "C:\Users\Abhishek\OneDrive\Desktop\Cultural Godrej Nest\Navratri_Day Contribution Statement 2026.xlsx" navratri-2026 --name "Navratri 2026"
```

The importer checks for the required tabs and headers, validates dates and positive amounts, calculates totals, and stops if the expense rows do not match the workbook's stated expense total. It keeps the anonymous donor label and omits that receipt's free-text remark. Review the generated JSON before building and deploying.

## Event data format

Each file in `src/events/` contains event metadata and the ledger data:

```json
{
  "id": "diwali-2026",
  "name": "Diwali 2026",
  "year": 2026,
  "receipts": [],
  "stalls": [],
  "expenses": [],
  "totals": {
    "contributions": 0,
    "stallFees": 0,
    "expenses": 0
  }
}
```

Use the importer to generate this structure so calculated totals and validation stay consistent. Event data is included in the generated site bundle; avoid putting information in these files that should not be public.

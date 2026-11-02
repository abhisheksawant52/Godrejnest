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

The Vercel project configuration is in `vercel.json`. Event JSON files are bundled into the deployed site. The dashboard also has an **Import Excel** button: imported events are saved in that browser's local storage and remain available there after reload. They are not uploaded to the server or shared with other visitors. To publish an event for everyone, use the command-line importer below and deploy the updated site.

## Add an event from Excel

Each workbook should contain worksheets with recognizable names and headers:

- `Total Receipts`: `Name`, `Flat No`, `Amount`, `Date`, `Time`, `Remarks`
- `Stall Payment`: `Name`, `Flat No`, `Amount`, `Date`, `Time`, `Remarks`
- `Total Expenses`: `Date`, `Particular`, `Amount`, `Paid to`, `Remarks`

Import a workbook from PowerShell:

```powershell
npm run import:statement -- "C:\path\to\Diwali_2026.xlsx" diwali-2026 --name "Diwali 2026"
```

This creates or refreshes `src/events/diwali-2026.json` and automatically runs a production build, updating the local website files in `dist/`. The event selector discovers JSON files in `src/events/` automatically. Deploy the updated site to publish the event.

The event ID is a lowercase slug made of letters, numbers, and hyphens. Reusing an ID updates that event's JSON file rather than adding a duplicate event. If omitted, the importer generates an ID from the workbook filename. The display name is optional and defaults to a title-cased version of the ID. The chart year is read from the display name or ID; if neither includes a year, the current year is used.

The importer reads every worksheet tab in the workbook. It identifies incoming payment tabs from their `Name`, `Amount`, and `Date` headers, identifies expense tabs from `Particular`/`Item`, `Amount`, and `Date`, and separates incoming tabs whose names contain “stall” into stall receipts. Other incoming tabs (including `Owner Registration Payment`) are included in contributions. Extra tabs without matching ledger headers are skipped. Each imported row retains its source tab name; the contribution register shows it in the **Source Sheet** column and expense details include it. The workbook tab names are therefore used both to classify data and to show its origin on the dashboard.

For example, to import or refresh the existing Navratri event, reuse its `navratri-2026` ID. Reusing the ID updates the existing event instead of creating a second event from the workbook filename:

```powershell
npm run import:statement -- "C:\Users\Abhishek\OneDrive\Desktop\Cultural Godrej Nest\Navratri_Day Contribution Statement 2026.xlsx" navratri-2026 --name "Navratri 2026"
npm run import:statement -- "C:\Users\Abhishek\OneDrive\Desktop\Cultural Godrej Nest\Cricket Contribution Statement 2026.xlsx" cricket-2026 --name "Cricket 2026"
```

The importer checks for the required tabs and headers, validates dates and positive amounts, calculates totals, and stops if the expense rows do not match the workbook's stated expense total. It keeps the anonymous donor label and omits that receipt's free-text remark. Review the generated JSON before building and deploying.

The dashboard upload accepts `.xlsx` workbooks directly. It reads worksheet names to identify stall payment sheets (a tab name containing “stall”), other incoming receipt sheets, and expense sheets. The event name comes from the workbook filename. Sheets need headers for `Name`, `Amount`, and `Date` for incoming payments, or `Particular`/`Item`, `Amount`, and `Date` for expenses. Imported browser data stays in that browser's local storage; use the CLI importer when the event needs to be included in a public deployment.

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

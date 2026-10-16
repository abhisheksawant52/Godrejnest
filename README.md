# Godrej Nest · Navratri community ledger

A responsive React and Vite dashboard for the Godrej Nest Navratri 2026 community funds summary. It presents aggregate contributions, a searchable, scrollable contribution register with all six receipt columns, recorded stall-fee receipts, outgoing event payments, and a date-based contribution trend. The receipt marked anonymous stays anonymous, with its identifying details withheld.

## Run locally

```sh
npm install
npm run dev
```

## Production build

```sh
npm run build
npm run preview
```

The Vercel project configuration is in `vercel.json`.

## Refresh dashboard data from the local workbook

The dashboard reads `src/statement.json`. After editing the Excel workbook, run this from the project folder in PowerShell:

```powershell
npm run import:statement
```

The importer defaults to `OneDrive\Desktop\Cultural Godrej Nest\Navratri_Day Contribution Statement 2026.xlsx` under your Windows user folder. To use a different file, pass its path:

```powershell
npm run import:statement -- "C:\path\to\Navratri_Day Contribution Statement 2026.xlsx"
```

The script checks the expected worksheet tabs and columns, validates dated amounts, recalculates totals, and stops if the expense rows do not match the workbook's stated total. It preserves the anonymous donor label and leaves out that receipt's free-text remark. After importing, review `src/statement.json`, then run `npm run build` before deploying.

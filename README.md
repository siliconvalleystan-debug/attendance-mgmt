# Attendance Calculator (React)

A React + Vite reimplementation of the attendance dashboard for summarising daily attendance from copied spreadsheet rows.

## Features
- Editable data grid that supports multi-row paste from Excel/Google Sheets using the columns `Employee ID`, `First Name`, `Department`, `Date`, `Time`, `Punch State`, `Work Code`, `Data Sources`.
- Date range filtering and configurable on-time threshold (default 08:30).
- Daily computations for earliest check-in, latest check-out, total hours, and on-time status.
- Calculation breakdown card plus an Excel export (`Info`, `Daily Summary`, `Calculation Steps`, `Punches`, `Raw Input` sheets).
- Sample dataset loader for quick demos.

## Scripts

```bash
# install dependencies
npm install

# run locally
npm run dev

# build for production
npm run build

# preview the production build
npm run preview
```

## Usage
1. Start the dev server with `npm run dev` and open the printed URL.
2. Paste attendance rows into the grid (tab/comma delimited) or click **Load Sample**.
3. Press **Process** to ingest the data and auto-fill the filter dates.
4. Adjust the filters or on-time threshold as needed.
5. Click **Export Results** to download the multi-sheet `.xlsx` workbook.

> Tip: If the app reports skipped rows, open the “Raw Input” sheet in the export to find lines with missing timestamps or punch labels, correct them in your source sheet, then paste again.

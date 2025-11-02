# Attendance Calculator

A lightweight browser-based tool for summarising daily attendance from copied spreadsheet rows.

## Features
- Type or paste tabular rows directly into the editable grid with the columns: `Employee ID`, `First Name`, `Department`, `Date`, `Time`, `Punch State`, `Work Code`, `Data Sources`.
- Filter results by date range and configure the on-time threshold (default 08:30).
- Calculates earliest check-in, latest check-out, on-time status, and total working hours per employee per day.
- Highlights late arrivals, shows quick summary cards, a mini step-by-step breakdown, and exports the processed results as a multi-sheet Excel workbook (info, daily summary, calculation steps, punch log, raw input).
- Includes a "Load Sample" shortcut to try the workflow instantly.

## Getting Started
1. Open `index.html` in any modern browser (Chrome, Firefox, Edge, Safari).
2. Copy the rows you want to analyse from Excel/Sheets and paste them straight into the first cell of the table (header row is already provided).
3. Click **Process** (or press `Ctrl/Cmd + Enter`) to convert the grid into daily summaries.
4. Adjust the filters and on-time limit as needed, then review the results table.
5. Click **Export Results** to download an `.xlsx` file containing the filtered summary, calculation steps, punch-level detail, and the original rows.

> Tip: The tool matches rows by the pasted `Punch State` values (e.g. "Check In", "Check Out"). If any rows are skipped, a warning will appear under the Process button—double-check the date/time or punch labels in those lines.
# attendance-mgmt

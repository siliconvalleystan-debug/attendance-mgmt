import React from 'react';
import dayjs from 'dayjs';
import customParseFormat from 'dayjs/plugin/customParseFormat';
import isSameOrAfter from 'dayjs/plugin/isSameOrAfter';
import isSameOrBefore from 'dayjs/plugin/isSameOrBefore';
import * as XLSX from 'xlsx';

dayjs.extend(customParseFormat);
dayjs.extend(isSameOrAfter);
dayjs.extend(isSameOrBefore);

const TABLE_COLUMNS = [
  'Employee ID',
  'First Name',
  'Department',
  'Date',
  'Time',
  'Punch State',
  'Work Code',
  'Data Sources',
];

const FIELD_CANDIDATES = {
  name: [
    'name',
    'employee',
    'employee name',
    'full name',
    'person',
    'staff name',
    'user name',
    'member',
    'first name',
  ],
  timestamp: [
    'timestamp',
    'date time',
    'datetime',
    'log time',
    'punch time',
    'attendance time',
    'record time',
  ],
  date: ['date', 'day', 'attendance date', 'work date'],
  time: ['time', 'clock time', 'attendance time'],
  type: ['type', 'event', 'action', 'status', 'direction', 'transaction type', 'check type', 'io', 'punch state'],
  checkIn: ['check in', 'time in', 'clock in', 'punch in', 'signin', 'sign in', 'in time', 'first in'],
  checkOut: ['check out', 'time out', 'clock out', 'punch out', 'signout', 'sign out', 'out time', 'last out'],
};

const TYPE_IN_KEYWORDS = ['check in', 'clock in', 'punch in', 'signin', 'sign in', 'login'];
const TYPE_OUT_KEYWORDS = ['check out', 'clock out', 'punch out', 'signout', 'sign out', 'logout'];
const IDENTIFIER_FIELDS = [
  'employee id',
  'employee code',
  'emp id',
  'employee number',
  'employee no',
  'id',
];

const INITIAL_ROWS = 12;

const BannerToneStyles = {
  info: 'border-sky-200 bg-sky-50 text-sky-700',
  success: 'border-emerald-200 bg-emerald-50 text-emerald-700',
  warn: 'border-amber-200 bg-amber-50 text-amber-700',
  error: 'border-red-200 bg-red-50 text-red-700',
};

function createEmptyRow() {
  return TABLE_COLUMNS.reduce((row, column) => {
    row[column] = '';
    return row;
  }, {});
}

export default function App() {
  const [gridRows, setGridRows] = React.useState(() =>
    Array.from({ length: INITIAL_ROWS }, () => createEmptyRow()),
  );
  const [banner, setBanner] = React.useState({ text: '', tone: 'info' });
  const [rawRows, setRawRows] = React.useState([]);
  const [records, setRecords] = React.useState([]);
  const [filteredResults, setFilteredResults] = React.useState([]);
  const [calculationDetails, setCalculationDetails] = React.useState([]);
  const [summary, setSummary] = React.useState({ employees: 0, onTimeRate: null, averageMinutes: null });
  const [startDate, setStartDate] = React.useState('');
  const [endDate, setEndDate] = React.useState('');
  const [onTimeLimit, setOnTimeLimit] = React.useState('08:30');
  const [lastFilterOptions, setLastFilterOptions] = React.useState(null);
  const [skippedRows, setSkippedRows] = React.useState(0);

  const handleCellChange = React.useCallback((rowIndex, column, value) => {
    setGridRows((prev) => {
      const next = [...prev];
      next[rowIndex] = { ...next[rowIndex], [column]: value };
      return next;
    });
  }, []);

  const handlePaste = React.useCallback((event, rowIndex, columnIndex) => {
    const text = event.clipboardData?.getData('text/plain');
    if (!text) return;
    event.preventDefault();

    const lines = text
      .split(/\r?\n/)
      .map((line) => line.trimEnd())
      .filter((line) => line.length > 0);
    if (!lines.length) return;

    const delimiter = detectDelimiter(lines[0]);

    setGridRows((prev) => {
      const next = [...prev];
      const requiredRows = rowIndex + lines.length;
      while (next.length < requiredRows) {
        next.push(createEmptyRow());
      }

      lines.forEach((line, lineOffset) => {
        const values = splitLine(line, delimiter);
        const targetRowIndex = rowIndex + lineOffset;
        const row = { ...next[targetRowIndex] };
        values.forEach((value, valueOffset) => {
          const targetColumnIndex = columnIndex + valueOffset;
          if (targetColumnIndex >= TABLE_COLUMNS.length) return;
          const columnKey = TABLE_COLUMNS[targetColumnIndex];
          row[columnKey] = value.trim();
        });
        next[targetRowIndex] = row;
      });

      return next;
    });
  }, []);

  const handleAddRow = React.useCallback(() => {
    setGridRows((prev) => [...prev, createEmptyRow()]);
  }, []);

  const handleLoadSample = React.useCallback(() => {
    const sampleRows = [
      ['1234', 'Amena', 'Operations', '2025-10-06', '08:04', 'Check In', 'HQ', 'Device A'],
      ['1234', 'Amena', 'Operations', '2025-10-06', '18:16', 'Check Out', 'HQ', 'Device A'],
      ['1234', 'Amena', 'Operations', '2025-10-07', '08:26', 'Check In', 'HQ', 'Device A'],
      ['1234', 'Amena', 'Operations', '2025-10-07', '17:53', 'Check Out', 'HQ', 'Device A'],
      ['5678', 'Jordan', 'Marketing', '2025-10-06', '08:42', 'Check In', 'HQ', 'Device B'],
      ['5678', 'Jordan', 'Marketing', '2025-10-06', '18:02', 'Check Out', 'HQ', 'Device B'],
    ];

    const filledRows = sampleRows.map((values) =>
      TABLE_COLUMNS.reduce((acc, column, index) => {
        acc[column] = values[index] ?? '';
        return acc;
      }, {}),
    );

    setGridRows([...filledRows, ...Array.from({ length: INITIAL_ROWS }, () => createEmptyRow())]);
    setBanner({ text: 'Loaded sample data. Click Process to generate the summary.', tone: 'info' });
  }, []);

  const collectRows = React.useCallback(() => {
    const rowsWithMeta = gridRows.map((row, index) => ({
      ...row,
      __lineNumber: index + 2,
    }));

    return rowsWithMeta.filter((row) => TABLE_COLUMNS.some((column) => (row[column] ?? '').trim().length > 0));
  }, [gridRows]);

  const autofillDateRange = React.useCallback(
    (normalizedRecords) => {
      if (!normalizedRecords.length) return;
      const dates = normalizedRecords
        .map((record) => record.timestamp?.startOf('day'))
        .filter((value) => value && value.isValid());

      if (!dates.length) return;
      const earliest = dates.reduce((min, candidate) => (candidate.isBefore(min) ? candidate : min));
      const latest = dates.reduce((max, candidate) => (candidate.isAfter(max) ? candidate : max));

      setStartDate((prev) => prev || earliest.format('YYYY-MM-DD'));
      setEndDate((prev) => prev || latest.format('YYYY-MM-DD'));
    },
    [setStartDate, setEndDate],
  );

  const applyFilters = React.useCallback(
    (recordsToFilter, optionsOverride) => {
      if (!recordsToFilter.length) {
        setFilteredResults([]);
        setCalculationDetails([]);
        setSummary({ employees: 0, onTimeRate: null, averageMinutes: null });
        setLastFilterOptions(null);
        return;
      }

      const filterOptions =
        optionsOverride ||
        {
          startDate: startDate ? dayjs(startDate).startOf('day') : null,
          endDate: endDate ? dayjs(endDate).endOf('day') : null,
          onTimeLimitMinutes: getMinutesFromTime(onTimeLimit),
        };

      const results = summarizeByEmployee(recordsToFilter, filterOptions);
      const computedSummary = computeSummary(results);
      const details = buildCalculationDetailEntries(results, filterOptions);

      setFilteredResults(results);
      setSummary(computedSummary);
      setCalculationDetails(details);
      setLastFilterOptions(filterOptions);
    },
    [startDate, endDate, onTimeLimit],
  );

  const handleProcess = React.useCallback(() => {
    try {
      const rows = collectRows();
      if (!rows.length) {
        setBanner({ text: 'Enter at least one data row before processing.', tone: 'error' });
        setRawRows([]);
        setRecords([]);
        setFilteredResults([]);
        setCalculationDetails([]);
        setSummary({ employees: 0, onTimeRate: null, averageMinutes: null });
        setSkippedRows(0);
        return;
      }

      let skipped = 0;
      const normalized = rows.flatMap((row, index) => {
        const normalizedRows = normalizeRow(row, index);
        if (!normalizedRows.length) {
          skipped += 1;
          console.warn(`Skipped row ${row.__lineNumber}: unable to parse required fields.`, row);
          return [];
        }
        return normalizedRows;
      });

      if (!normalized.length) {
        setBanner({
          text: 'No usable attendance records were found in the pasted data.',
          tone: 'error',
        });
        setRawRows(rows);
        setRecords([]);
        setFilteredResults([]);
        setCalculationDetails([]);
        setSummary({ employees: 0, onTimeRate: null, averageMinutes: null });
        setSkippedRows(skipped);
        return;
      }

      setRawRows(rows);
      setRecords(normalized);
      setSkippedRows(skipped);
      autofillDateRange(normalized);

      const tone = skipped > 0 ? 'warn' : 'success';
      const message =
        skipped > 0
          ? `Processed with ${skipped} row(s) skipped due to missing or invalid data.`
          : 'Processed successfully.';
      setBanner({ text: message, tone });

      applyFilters(normalized);
    } catch (error) {
      console.error(error);
      setBanner({
        text: error instanceof Error ? error.message : 'Failed to process the pasted data.',
        tone: 'error',
      });
    }
  }, [applyFilters, autofillDateRange, collectRows]);

  React.useEffect(() => {
    if (!records.length) return;
    applyFilters(records);
  }, [records, startDate, endDate, onTimeLimit, applyFilters]);

  const handleExport = React.useCallback(() => {
    if (!filteredResults.length) return;

    const filterOptions =
      lastFilterOptions ||
      {
        startDate: startDate ? dayjs(startDate).startOf('day') : null,
        endDate: endDate ? dayjs(endDate).endOf('day') : null,
        onTimeLimitMinutes: getMinutesFromTime(onTimeLimit),
      };

    const summaryData = filteredResults.map((item) => ({
      Name: item.name,
      Date: item.date,
      'Earliest Check-In': item.earliestIn ? item.earliestIn.format('HH:mm') : '',
      'Latest Check-Out': item.latestOut ? item.latestOut.format('HH:mm') : '',
      'On Time': item.isOnTime === null ? '' : item.isOnTime ? 'Yes' : 'No',
      'Total Hours': item.totalMinutes != null ? formatDuration(item.totalMinutes) : '',
    }));

    const detailEntries = buildCalculationDetailEntries(filteredResults, filterOptions);
    const detailSheetData = detailEntries.map((entry) => ({
      Name: entry.name,
      Date: entry.date,
      'Earliest Check-In': entry.earliestInText === '—' ? '' : entry.earliestInText,
      'Latest Check-Out': entry.latestOutText === '—' ? '' : entry.latestOutText,
      'On-Time Limit': entry.onTimeLimitLabel,
      'On Time?': entry.onTimeStatus,
      Explanation: entry.onTimeExplanation,
      'Total Hours': entry.totalText === '—' ? '' : entry.totalText,
      'Hours Calculation': entry.calcExpression,
    }));

    const punchesData = records
      .filter((record) => recordMatchesFilter(record, filterOptions))
      .sort((a, b) => {
        const nameCompare = a.name.localeCompare(b.name);
        if (nameCompare !== 0) return nameCompare;
        return a.timestamp.valueOf() - b.timestamp.valueOf();
      })
      .map((record) => {
        const rawRow = typeof record.sourceRowIndex === 'number' ? rawRows[record.sourceRowIndex] : null;
        const punchState =
          rawRow && rawRow['Punch State']
            ? rawRow['Punch State']
            : record.type === 'in'
            ? 'Check In'
            : 'Check Out';
        return {
          Name: record.name,
          'Employee ID': rawRow?.['Employee ID'] ?? '',
          'First Name': rawRow?.['First Name'] ?? '',
          Department: rawRow?.['Department'] ?? '',
          Date: record.timestamp.format('YYYY-MM-DD'),
          Time: record.timestamp.format('HH:mm'),
          'Punch State': punchState,
          'Work Code': rawRow?.['Work Code'] ?? '',
          'Data Sources': rawRow?.['Data Sources'] ?? '',
          'Source Row #': rawRow?.__lineNumber ?? '',
        };
      });

    const rawInputData = rawRows.map((row, index) => {
      const entry = { '#': index + 1 };
      TABLE_COLUMNS.forEach((column) => {
        entry[column] = row?.[column] ?? '';
      });
      return entry;
    });

    const summarySnapshot = summary ?? computeSummary(filteredResults);
    const onTimeLimitLabel =
      detailEntries[0]?.onTimeLimitLabel || formatMinutesAsTime(filterOptions.onTimeLimitMinutes);
    const infoRows = [
      ['Generated on', dayjs().format('YYYY-MM-DD HH:mm')],
      ['On-time limit', onTimeLimitLabel],
      ['Filter - From', filterOptions.startDate ? filterOptions.startDate.format('YYYY-MM-DD') : '—'],
      ['Filter - To', filterOptions.endDate ? filterOptions.endDate.format('YYYY-MM-DD') : '—'],
      ['Total employees (filtered)', summarySnapshot?.employees ?? 0],
      ['Total day rows', filteredResults.length],
      ['On-time rate', summarySnapshot?.onTimeRate != null ? `${summarySnapshot.onTimeRate}%` : '—'],
      [
        'Average hours / day',
        summarySnapshot?.averageMinutes != null ? formatDuration(summarySnapshot.averageMinutes) : '—',
      ],
      ['Skipped rows during import', skippedRows],
    ];

    const workbook = XLSX.utils.book_new();
    const infoSheet = XLSX.utils.aoa_to_sheet(infoRows);
    infoSheet['!cols'] = [{ wch: 28 }, { wch: 32 }];
    XLSX.utils.book_append_sheet(workbook, infoSheet, 'Info');

    const summarySheet = XLSX.utils.json_to_sheet(summaryData);
    summarySheet['!cols'] = [
      { wch: 32 },
      { wch: 14 },
      { wch: 16 },
      { wch: 16 },
      { wch: 12 },
      { wch: 15 },
    ];
    XLSX.utils.book_append_sheet(workbook, summarySheet, 'Daily Summary');

    if (detailSheetData.length) {
      const detailSheet = XLSX.utils.json_to_sheet(detailSheetData);
      detailSheet['!cols'] = [
        { wch: 32 },
        { wch: 14 },
        { wch: 16 },
        { wch: 16 },
        { wch: 16 },
        { wch: 10 },
        { wch: 50 },
        { wch: 15 },
        { wch: 20 },
      ];
      XLSX.utils.book_append_sheet(workbook, detailSheet, 'Calculation Steps');
    }

    if (punchesData.length) {
      const punchesSheet = XLSX.utils.json_to_sheet(punchesData);
      punchesSheet['!cols'] = [
        { wch: 32 },
        { wch: 14 },
        { wch: 16 },
        { wch: 14 },
        { wch: 10 },
        { wch: 14 },
        { wch: 16 },
        { wch: 18 },
        { wch: 18 },
        { wch: 12 },
      ];
      XLSX.utils.book_append_sheet(workbook, punchesSheet, 'Punches');
    }

    if (rawInputData.length) {
      const rawSheet = XLSX.utils.json_to_sheet(rawInputData);
      rawSheet['!cols'] = [{ wch: 6 }, ...TABLE_COLUMNS.map(() => ({ wch: 18 }))];
      XLSX.utils.book_append_sheet(workbook, rawSheet, 'Raw Input');
    }

    const timestamp = dayjs().format('YYYYMMDD-HHmmss');
    XLSX.writeFile(workbook, `attendance-summary-${timestamp}.xlsx`, { compression: true });
  }, [filteredResults, lastFilterOptions, onTimeLimit, rawRows, records, skippedRows, startDate, summary]);

  const employeesCount = summary?.employees ?? 0;
  const exportDisabled = filteredResults.length === 0;

  return (
    <div className="min-h-screen bg-slate-100 pb-16">
      <header className="bg-slate-900 text-white shadow">
        <div className="mx-auto flex max-w-5xl flex-col gap-3 px-6 py-8 md:flex-row md:items-center md:justify-between">
          <div>
            <h1 className="text-2xl font-semibold">Attendance Calculator</h1>
            <p className="text-sm text-slate-300">
              Paste attendance logs, filter by date, and analyse on-time performance.
            </p>
          </div>
          <div className="flex gap-2 text-sm">
            <button
              type="button"
              onClick={handleLoadSample}
              className="inline-flex items-center rounded-md border border-white/20 px-3 py-2 font-medium text-white transition hover:border-white/40 hover:bg-white/10"
            >
              Load Sample
            </button>
            <button
              type="button"
              onClick={handleExport}
              disabled={exportDisabled}
              className="inline-flex items-center rounded-md border border-white/20 px-3 py-2 font-medium text-white transition hover:border-white/40 hover:bg-white/10 disabled:pointer-events-none disabled:opacity-50"
            >
              Export Results
            </button>
          </div>
        </div>
      </header>

      <main className="mx-auto flex max-w-5xl flex-col gap-8 px-6 py-10">
        <section>
          <h2 className="text-lg font-semibold text-slate-700">1. Paste attendance data</h2>
          <p className="mb-4 text-sm text-slate-500">
            Copy rows from your spreadsheet (tab or comma separated) with the columns:{' '}
            <span className="font-medium">
              Employee ID, First Name, Department, Date, Time, Punch State, Work Code, Data Sources
            </span>
            .
          </p>
          <p className="mb-3 rounded-lg bg-slate-200/60 px-4 py-2 text-xs text-slate-600">
            Tip: click the first cell, then paste (<span className="font-semibold">Ctrl/Cmd + V</span>) directly from
            Excel or Google Sheets. Add extra rows with the button below if needed.
          </p>
          <div className="overflow-x-auto rounded-xl border border-slate-300 bg-white shadow-sm">
            <table className="min-w-full border-separate border-spacing-0 text-sm text-slate-700">
              <thead className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
                <tr>
                  {TABLE_COLUMNS.map((column) => (
                    <th key={column} className="border-b border-slate-200 px-4 py-3 font-semibold">
                      {column}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {gridRows.map((row, rowIndex) => (
                  <tr key={`row-${rowIndex}`} className="divide-x divide-slate-100">
                    {TABLE_COLUMNS.map((column, columnIndex) => (
                      <td key={`${rowIndex}-${column}`} className="px-0 py-0">
                        <input
                          type="text"
                          value={row[column]}
                          onChange={(event) => handleCellChange(rowIndex, column, event.target.value)}
                          onPaste={(event) => handlePaste(event, rowIndex, columnIndex)}
                          className="h-full w-full border-none px-4 py-3 font-mono text-xs text-slate-800 focus:outline-none focus:ring-2 focus:ring-slate-400"
                          placeholder={rowIndex === 0 ? column : ''}
                          spellCheck="false"
                        />
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="mt-3 flex flex-wrap items-center gap-3">
            <button
              type="button"
              onClick={handleAddRow}
              className="inline-flex items-center rounded-md border border-slate-300 px-3 py-2 text-xs font-medium text-slate-700 transition hover:border-slate-400 hover:bg-slate-100"
            >
              Add Row
            </button>
            <button
              type="button"
              onClick={handleProcess}
              className="inline-flex items-center rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white shadow transition hover:bg-slate-800"
            >
              Process
            </button>
            {banner.text && (
              <p
                className={`rounded-lg border px-3 py-2 text-sm font-medium ${
                  BannerToneStyles[banner.tone] ?? BannerToneStyles.info
                }`}
              >
                {banner.text}
              </p>
            )}
          </div>
        </section>

        <section className="grid gap-6 rounded-xl bg-white p-6 shadow md:grid-cols-2">
          <div className="space-y-4">
            <h2 className="text-lg font-semibold text-slate-700">2. Filter options</h2>
            <div className="grid grid-cols-2 gap-4">
              <label className="flex flex-col gap-1 text-sm font-medium text-slate-600">
                From
                <input
                  type="date"
                  value={startDate}
                  onChange={(event) => setStartDate(event.target.value)}
                  className="rounded-md border-slate-300"
                />
              </label>
              <label className="flex flex-col gap-1 text-sm font-medium text-slate-600">
                To
                <input
                  type="date"
                  value={endDate}
                  onChange={(event) => setEndDate(event.target.value)}
                  className="rounded-md border-slate-300"
                />
              </label>
            </div>
            <label className="flex flex-col gap-1 text-sm font-medium text-slate-600">
              On-time before
              <input
                type="time"
                value={onTimeLimit}
                onChange={(event) => setOnTimeLimit(event.target.value || '08:30')}
                className="rounded-md border-slate-300"
              />
            </label>
            <button
              type="button"
              onClick={() => applyFilters(records)}
              className="inline-flex items-center justify-center rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white shadow transition hover:bg-slate-800 disabled:pointer-events-none disabled:opacity-50"
              disabled={!records.length}
            >
              Filter
            </button>
          </div>
          <div className="space-y-4">
            <h2 className="text-lg font-semibold text-slate-700">Summary</h2>
            <div id="summaryCards" className="grid gap-4 sm:grid-cols-2">
              <div className="rounded-lg border border-slate-200 p-4">
                <p className="text-xs uppercase text-slate-500">Employees</p>
                <p className="mt-1 text-2xl font-semibold text-slate-900">{employeesCount || '—'}</p>
              </div>
              <div className="rounded-lg border border-slate-200 p-4">
                <p className="text-xs uppercase text-slate-500">On-time rate</p>
                <p className="mt-1 text-2xl font-semibold text-slate-900">
                  {summary?.onTimeRate != null ? `${summary.onTimeRate}%` : '—'}
                </p>
              </div>
              <div className="rounded-lg border border-slate-200 p-4 sm:col-span-2">
                <p className="text-xs uppercase text-slate-500">Average hours / day</p>
                <p className="mt-1 text-2xl font-semibold text-slate-900">
                  {summary?.averageMinutes != null ? formatDuration(summary.averageMinutes) : '—'}
                </p>
              </div>
              <div className="rounded-lg border border-slate-200 p-4 sm:col-span-2">
                <p className="text-xs uppercase text-slate-500">Calculation steps</p>
                <div className="mt-3 space-y-3 text-xs leading-relaxed text-slate-600">
                  {calculationDetails.length === 0 ? (
                    <p>
                      Paste data and click <span className="font-semibold text-slate-700">Process</span> to see how each
                      summary is derived.
                    </p>
                  ) : (
                    calculationDetails.slice(0, 5).map((detail) => (
                      <div key={`${detail.name}-${detail.date}`} className="rounded-md bg-slate-50 p-3">
                        <p className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                          {detail.name} • {detail.date}
                        </p>
                        <ul className="mt-2 space-y-1 text-[11px] text-slate-600">
                          <li>
                            <span className="font-medium text-slate-700">Earliest check-in:</span>{' '}
                            {detail.earliestInText}
                          </li>
                          <li>
                            <span className="font-medium text-slate-700">Latest check-out:</span>{' '}
                            {detail.latestOutText}
                          </li>
                          <li>
                            <span className="font-medium text-slate-700">Total hours:</span> {detail.totalText}{' '}
                            {detail.calcExpression !== 'N/A' ? `(${detail.calcExpression})` : ''}
                          </li>
                          <li>
                            <span className="font-medium text-slate-700">On-time check:</span>{' '}
                            {detail.onTimeExplanation}
                          </li>
                        </ul>
                      </div>
                    ))
                  )}
                  {calculationDetails.length > 5 && (
                    <p className="text-[11px] text-slate-500">
                      …and {calculationDetails.length - 5} more day(s).
                    </p>
                  )}
                </div>
              </div>
            </div>
          </div>
        </section>

        <section className="rounded-xl bg-white p-6 shadow">
          <div className="mb-4 flex items-center justify-between gap-4">
            <h2 className="text-lg font-semibold text-slate-700">3. Results</h2>
            <div className="text-sm text-slate-500" id="resultsCount">
              {filteredResults.length > 0
                ? `Showing ${filteredResults.length} day(s) for ${employeesCount} employee(s).`
                : 'Showing 0 results'}
            </div>
          </div>
          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-slate-200">
              <thead className="bg-slate-50 text-left text-xs uppercase tracking-wider text-slate-500">
                <tr>
                  <th className="px-4 py-3">Name</th>
                  <th className="px-4 py-3">Date</th>
                  <th className="px-4 py-3">Earliest Check-In</th>
                  <th className="px-4 py-3">Latest Check-Out</th>
                  <th className="px-4 py-3">On Time?</th>
                  <th className="px-4 py-3">Total Hours</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-sm text-slate-700">
                {filteredResults.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="px-4 py-6 text-center text-sm text-slate-500">
                      No records match the selected filters.
                    </td>
                  </tr>
                ) : (
                  filteredResults.map((item) => (
                    <tr key={`${item.name}-${item.date}`} className={item.isOnTime === false ? 'bg-red-50/40' : ''}>
                      <td className="px-4 py-3 font-medium text-slate-800">{item.name}</td>
                      <td className="px-4 py-3 text-slate-600">{item.date}</td>
                      <td className="px-4 py-3">{item.earliestIn ? item.earliestIn.format('HH:mm') : '—'}</td>
                      <td className="px-4 py-3">{item.latestOut ? item.latestOut.format('HH:mm') : '—'}</td>
                      <td className="px-4 py-3 font-medium">
                        {item.isOnTime === null ? (
                          <span className="text-slate-400">—</span>
                        ) : item.isOnTime ? (
                          <span className="text-emerald-600">✅ On time</span>
                        ) : (
                          <span className="text-red-600">❌ Late</span>
                        )}
                      </td>
                      <td className="px-4 py-3">{item.totalMinutes != null ? formatDuration(item.totalMinutes) : '—'}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </section>
      </main>
    </div>
  );
}

function detectDelimiter(line) {
  const delimiters = ['\t', ',', ';', '|'];
  let bestDelimiter = '\t';
  let bestCount = 0;

  for (const delimiter of delimiters) {
    const count = line.split(delimiter).length - 1;
    if (count > bestCount) {
      bestDelimiter = delimiter;
      bestCount = count;
    }
  }

  return bestCount === 0 ? '\t' : bestDelimiter;
}

function splitLine(line, delimiter) {
  const cells = [];
  let current = '';
  let insideQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];

    if (char === '"') {
      if (insideQuotes && line[index + 1] === '"') {
        current += '"';
        index += 1;
      } else {
        insideQuotes = !insideQuotes;
      }
      continue;
    }

    if (char === delimiter && !insideQuotes) {
      cells.push(current);
      current = '';
      continue;
    }

    current += char;
  }

  cells.push(current);
  return cells.map((cell) => cell.trim());
}

function normalizeRow(row, rowIndex) {
  const primaryName = readField(row, FIELD_CANDIDATES.name);
  const employeeId = readField(row, IDENTIFIER_FIELDS);
  let name = primaryName;

  if (!name && employeeId) {
    name = String(employeeId);
  }

  if (!name) return [];

  name = String(name).trim();
  if (!name) return [];

  if (employeeId && !name.includes(String(employeeId))) {
    name = `${name} (${employeeId})`;
  }

  const sourceRowIndex = rowIndex ?? null;

  const typeRaw = readField(row, FIELD_CANDIDATES.type);

  if (typeRaw) {
    const normalizedType = normalizeType(typeRaw);
    if (!normalizedType) return [];

    let timestamp = parseTimestamp(readField(row, FIELD_CANDIDATES.timestamp));
    if (!timestamp) {
      const rawDate = readField(row, FIELD_CANDIDATES.date);
      const rawTime = readField(row, FIELD_CANDIDATES.time);
      timestamp = parseDateAndTime(rawDate, rawTime);
    }
    if (!timestamp) return [];

    return [
      {
        name,
        type: normalizedType,
        timestamp,
        sourceRowIndex,
      },
    ];
  }

  const rawDate = readField(row, FIELD_CANDIDATES.date);
  const checkInValue = readField(row, FIELD_CANDIDATES.checkIn);
  const checkOutValue = readField(row, FIELD_CANDIDATES.checkOut);

  const events = [];

  const baseDate = rawDate ? parseDate(rawDate) : null;

  const checkInTimestamp = buildTimestampFromDateAndValue(baseDate, checkInValue);
  if (checkInTimestamp) {
    events.push({
      name,
      type: 'in',
      timestamp: checkInTimestamp,
      sourceRowIndex,
    });
  }

  const checkOutTimestamp = buildTimestampFromDateAndValue(baseDate, checkOutValue);
  if (checkOutTimestamp) {
    events.push({
      name,
      type: 'out',
      timestamp: checkOutTimestamp,
      sourceRowIndex,
    });
  }

  return events;
}

function readField(row, candidates) {
  if (!row || !candidates.length) return null;

  const normalizedEntries = Object.entries(row).map(([key, value]) => [normalizeKey(key), value]);

  for (const candidate of candidates) {
    const normalizedCandidate = normalizeKey(candidate);
    const directMatch = normalizedEntries.find(([normalizedKey]) => normalizedKey === normalizedCandidate);
    if (directMatch && directMatch[1] !== undefined && directMatch[1] !== null && directMatch[1] !== '') {
      return directMatch[1];
    }
  }

  for (const candidate of candidates) {
    const normalizedCandidate = normalizeKey(candidate);
    const partialMatch = normalizedEntries.find(
      ([normalizedKey, value]) => normalizedKey.includes(normalizedCandidate) && value !== undefined && value !== null && value !== '',
    );
    if (partialMatch) return partialMatch[1];
  }

  return null;
}

function normalizeKey(key) {
  return String(key ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function buildTimestampFromDateAndValue(baseDate, rawValue) {
  if (rawValue == null || rawValue === '') return null;

  if (rawValue instanceof Date) return dayjs(rawValue);

  if (typeof rawValue === 'number') {
    if (rawValue >= 1) {
      const fullTimestamp = parseTimestamp(rawValue);
      if (fullTimestamp) return fullTimestamp;
    }
    if (!baseDate) return null;
    const time = parseTime(rawValue);
    if (time) {
      return baseDate.hour(time.hour).minute(time.minute).second(time.second);
    }
    return null;
  }

  const rawText = String(rawValue).trim();
  if (!rawText) return null;

  if (valueHasDateComponent(rawText)) {
    const fullTimestamp = parseTimestamp(rawText);
    if (fullTimestamp) return fullTimestamp;
  }

  if (!baseDate) return null;

  const time = parseTime(rawText);
  if (time) {
    return baseDate.hour(time.hour).minute(time.minute).second(time.second);
  }
  return null;
}

function valueHasDateComponent(text) {
  return /\d{4}[./-]\d{1,2}[./-]\d{1,2}/.test(text) || /\d{1,2}[./-]\d{1,2}[./-]\d{2,4}/.test(text);
}

function normalizeType(value) {
  if (value == null) return null;
  const text = String(value).trim().toLowerCase();
  if (!text) return null;

  const directMatches = {
    in: ['in', 'i', 'time in', 'in time'],
    out: ['out', 'o', 'time out', 'out time'],
  };

  if (directMatches.in.includes(text)) return 'in';
  if (directMatches.out.includes(text)) return 'out';

  if (TYPE_IN_KEYWORDS.some((keyword) => text.includes(keyword))) return 'in';
  if (TYPE_OUT_KEYWORDS.some((keyword) => text.includes(keyword))) return 'out';
  return null;
}

function parseTimestamp(value) {
  if (!value && value !== 0) return null;

  if (value instanceof Date) return dayjs(value);
  if (typeof value === 'number' && Number.isFinite(value)) {
    const dateParts = XLSX.SSF.parse_date_code(value);
    if (!dateParts) return null;
    const jsDate = new Date(
      dateParts.y,
      Math.max(0, (dateParts.m || 1) - 1),
      dateParts.d || 1,
      dateParts.H || 0,
      dateParts.M || 0,
      dateParts.S || 0,
    );
    return dayjs(jsDate);
  }

  const cleaned = String(value).trim();
  if (!cleaned) return null;

  const supportedFormats = [
    'YYYY-MM-DD HH:mm:ss',
    'YYYY-MM-DD HH:mm',
    'YYYY/MM/DD HH:mm',
    'DD/MM/YYYY HH:mm',
    'MM/DD/YYYY HH:mm',
    'MM/DD/YYYY h:mm A',
    'DD-MM-YYYY HH:mm',
    'M/D/YYYY H:mm',
    'M/D/YYYY h:mm A',
    'D/M/YYYY H:mm',
    'D/M/YYYY h:mm A',
    'YYYY-MM-DDTHH:mm:ss',
    'YYYY-MM-DDTHH:mm:ssZ',
    'YYYY-MM-DDTHH:mm',
  ];

  for (const format of supportedFormats) {
    const parsed = dayjs(cleaned, format, true);
    if (parsed.isValid()) return parsed;
  }

  const fallback = dayjs(cleaned);
  return fallback.isValid() ? fallback : null;
}

function parseDateAndTime(dateValue, timeValue) {
  const date = parseDate(dateValue);
  if (!date) return null;

  if (!timeValue && timeValue !== 0) return null;

  const time = parseTime(timeValue);
  if (!time) return null;

  return date.hour(time.hour).minute(time.minute).second(time.second);
}

function parseDate(value) {
  if (!value && value !== 0) return null;

  if (value instanceof Date) return dayjs(value).startOf('day');
  if (typeof value === 'number' && Number.isFinite(value)) {
    const dateParts = XLSX.SSF.parse_date_code(value);
    if (!dateParts) return null;
    return dayjs(new Date(dateParts.y, Math.max(0, (dateParts.m || 1) - 1), dateParts.d || 1)).startOf('day');
  }

  const cleaned = String(value).trim();
  if (!cleaned) return null;

  const formats = ['YYYY-MM-DD', 'DD/MM/YYYY', 'MM/DD/YYYY', 'DD-MM-YYYY', 'M/D/YYYY', 'D/M/YYYY'];
  for (const format of formats) {
    const parsed = dayjs(cleaned, format, true);
    if (parsed.isValid()) return parsed.startOf('day');
  }

  const fallback = dayjs(cleaned);
  return fallback.isValid() ? fallback.startOf('day') : null;
}

function parseTime(value) {
  if (!value && value !== 0) return null;

  if (value instanceof Date) {
    return { hour: value.getHours(), minute: value.getMinutes(), second: value.getSeconds() };
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const minutes = Math.round(value * 24 * 60);
    const hour = Math.floor(minutes / 60);
    const minute = minutes % 60;
    return { hour, minute, second: 0 };
  }

  const cleaned = String(value).trim();
  if (!cleaned) return null;

  const formats = ['HH:mm', 'H:mm', 'HH:mm:ss', 'h:mm A', 'h:mm:ss A'];
  for (const format of formats) {
    const parsed = dayjs(cleaned, format, true);
    if (parsed.isValid()) {
      return { hour: parsed.hour(), minute: parsed.minute(), second: parsed.second() };
    }
  }

  const fallback = dayjs(cleaned);
  if (fallback.isValid()) {
    return { hour: fallback.hour(), minute: fallback.minute(), second: fallback.second() };
  }
  return null;
}

function summarizeByEmployee(records, options) {
  const grouped = new Map();

  for (const record of records) {
    const { name, timestamp, type } = record;

    if (options.startDate && timestamp.isBefore(options.startDate, 'day')) continue;
    if (options.endDate && timestamp.isAfter(options.endDate, 'day')) continue;

    const dayKey = timestamp.format('YYYY-MM-DD');
    const key = `${name}__${dayKey}`;
    if (!grouped.has(key)) {
      grouped.set(key, {
        name,
        date: timestamp.startOf('day'),
        checkIns: [],
        checkOuts: [],
      });
    }

    const group = grouped.get(key);
    if (type === 'in') {
      group.checkIns.push(timestamp);
    } else if (type === 'out') {
      group.checkOuts.push(timestamp);
    }
  }

  const results = [];

  for (const [, value] of grouped) {
    if (!value.checkIns.length && !value.checkOuts.length) continue;

    const earliestIn = value.checkIns.length
      ? value.checkIns.reduce((earliest, candidate) => (candidate.isBefore(earliest) ? candidate : earliest))
      : null;

    const latestOut = value.checkOuts.length
      ? value.checkOuts.reduce((latest, candidate) => (candidate.isAfter(latest) ? candidate : latest))
      : null;

    let totalMinutes = null;
    if (earliestIn && latestOut && latestOut.isAfter(earliestIn)) {
      totalMinutes = Math.round(latestOut.diff(earliestIn, 'minute', true));
    }

    let isOnTime = null;
    if (earliestIn) {
      const minutes = earliestIn.hour() * 60 + earliestIn.minute();
      isOnTime = minutes <= options.onTimeLimitMinutes;
    }

    results.push({
      name: value.name,
      date: value.date.format('YYYY-MM-DD'),
      earliestIn,
      latestOut,
      isOnTime,
      totalMinutes,
    });
  }

  return results.sort((a, b) => {
    if (a.name.toLowerCase() === b.name.toLowerCase()) {
      return dayjs(a.date).diff(dayjs(b.date));
    }
    return a.name.localeCompare(b.name);
  });
}

function computeSummary(results) {
  const uniqueEmployees = new Set(results.map((item) => item.name)).size;
  const onTimeRecords = results.filter((item) => item.isOnTime !== null);
  const onTimeCount = onTimeRecords.filter((item) => item.isOnTime).length;
  const onTimeRate =
    onTimeRecords.length > 0 ? Math.round((onTimeCount / onTimeRecords.length) * 100 * 10) / 10 : null;

  const totalMinutes = results
    .filter((item) => item.totalMinutes != null)
    .reduce((sum, item) => sum + item.totalMinutes, 0);

  const averageMinutesCount = results.filter((item) => item.totalMinutes != null).length;
  const averageMinutes = averageMinutesCount > 0 ? Math.round(totalMinutes / averageMinutesCount) : null;

  return { employees: uniqueEmployees, onTimeRate, averageMinutes };
}

function buildCalculationDetailEntries(results, options = {}) {
  if (!results.length) return [];
  const onTimeLimitMinutes =
    typeof options.onTimeLimitMinutes === 'number'
      ? options.onTimeLimitMinutes
      : getMinutesFromTime(options.onTimeLimitLabel ?? '08:30');
  const onTimeLimitLabel = formatMinutesAsTime(onTimeLimitMinutes);

  return results.map((item) => {
    const earliestInText = item.earliestIn ? item.earliestIn.format('HH:mm') : '—';
    const latestOutText = item.latestOut ? item.latestOut.format('HH:mm') : '—';
    const totalText = item.totalMinutes != null ? formatDuration(item.totalMinutes) : '—';
    const onTimeStatus = item.isOnTime === null ? '—' : item.isOnTime ? 'On Time' : 'Late';

    let onTimeExplanation;
    if (item.isOnTime === null) {
      onTimeExplanation = 'No check-in detected for this day.';
    } else if (item.isOnTime) {
      onTimeExplanation = `On time because ${earliestInText} ≤ ${onTimeLimitLabel}.`;
    } else {
      onTimeExplanation = `Late because ${earliestInText} > ${onTimeLimitLabel}.`;
    }

    const calcExpression =
      item.totalMinutes != null && item.earliestIn && item.latestOut
        ? `${latestOutText} − ${earliestInText}`
        : 'N/A';

    return {
      name: item.name,
      date: item.date,
      earliestInText,
      latestOutText,
      totalText,
      calcExpression,
      onTimeExplanation,
      onTimeStatus,
      onTimeLimitLabel,
      totalMinutes: item.totalMinutes,
      earliestIn: item.earliestIn,
      latestOut: item.latestOut,
      isOnTime: item.isOnTime,
    };
  });
}

function formatDuration(minutes) {
  const hours = Math.floor(minutes / 60);
  const remainingMinutes = Math.abs(minutes % 60);
  return `${hours}h ${remainingMinutes}m`;
}

function formatMinutesAsTime(minutes) {
  const hours = Math.floor(minutes / 60);
  const mins = Math.abs(minutes % 60);
  return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
}

function getMinutesFromTime(timeString) {
  const [hours, minutes] = timeString.split(':').map(Number);
  if (Number.isFinite(hours) && Number.isFinite(minutes)) {
    return hours * 60 + minutes;
  }
  return 8 * 60 + 30;
}

function recordMatchesFilter(record, options) {
  if (!options) return true;
  if (options.startDate && record.timestamp.isBefore(options.startDate, 'day')) return false;
  if (options.endDate && record.timestamp.isAfter(options.endDate, 'day')) return false;
  return true;
}

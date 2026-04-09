# Excel-Report-Generator

Vite app that ingests a large Excel file from a file input and maps unique column A values to unique column B/C values.

## Run locally

```bash
npm install
npm run dev
```

Open the local Vite URL, upload any common Excel-importable file (`.xlsx`, `.xlsm`, `.xlsb`, `.xltx`, `.xltm`, `.xlam`, `.xls`, `.xlt`, `.xml`, `.csv`, `.txt`, `.prn`, `.dif`, `.slk`, `.ods`, `.fods`).

The app treats the first row as headers/labels, then:
1. Displays unique column A values in alphabetical order.
2. Displays each A value's unique column B/C pairs alphabetically (by B, then C).
3. Places B and C values one row below A, with C shown immediately next to B.
4. Adds an "Entries Per Location" column that shows how many source rows match that exact A/B/C combination.
5. Adds a separate "Number Passed/Failed" column that shows K-column percentages for that same combination based on Entries Per Location (for example: `Value1: 55.56%, Value2: 44.44%`). If only one K value exists, it still shows the associated value label (for example: `Value1: 100%`).
6. Colors each mapped row by average attempts-to-pass across G-column test sequences (sequence resets when test numbers are no longer consecutive): no pass `#DA9694` only when there are zero passes, more than 4 attempts `red`, third/fourth attempt average `orange`, second attempt average `yellow`, first attempt average `green`.
7. Adds an "RLU Averages / Limits / Anomalies" column with: overall average of column E (shown as RLU), average RLU grouped by matching K values, the column F values as Limits (shown as `value Limit`), anomaly flags (non-numeric RLU values, negative RLU values, all-fail/no-pass cases, and IQR-based outliers with values/causes), and a "Data Integrity Risk" percentage. Risk is shown only when non-zero, and includes the specific alarming fail-to-pass sequence values next to it. Detection is loosened by using column F passing ranges so fail-to-pass drops are not flagged when the passing RLU value is still within the allowed range. "Anomalies" is shown only when anomalies are present.
8. Adds a separate "Calibration Review (O)" tab that groups rows by column O and scores calibration consistency using column E distributions plus G-sequence and K-pass behavior (including pass rate, average attempts-to-pass, sequence spread, and outlier rate).
9. Lets you export a styled `.xlsx` report using an "Export Excel Report" button. The export creates two sheets (`Mapped View` and `Calibration Review`) that mirror the two on-screen tabs, including column widths, row spacing, highlighted mapped row status colors, and wrapped analytics text.

### Export naming and save location

- On browsers that support the File System Access API, export opens the native save dialog so you can choose file name and location directly.
- On other browsers, export asks for the file name in a prompt and then downloads the file using the browser's default download behavior/location.

## Build

```bash
npm run build
```
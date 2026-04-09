# GH Copilot Continuity Notes

Last updated: 2026-04-09

## Project Snapshot

- App type: Vite + vanilla JS browser app
- Primary source file: src/main.js
- Styles: src/style.css
- Parser dependency: xlsx
- Export dependency: exceljs (lazy-loaded at export time)
- Purpose: import raw spreadsheet data, build mapped/analytics views, export styled 2-sheet Excel report

## Current Product Behavior

### Input and Processing

- Accepts common Excel-importable formats (`.xlsx`, `.xlsm`, `.xlsb`, `.xltx`, `.xltm`, `.xlam`, `.xls`, `.xlt`, `.xml`, `.csv`, `.txt`, `.prn`, `.dif`, `.slk`, `.ods`, `.fods`).
- Treats first non-empty A/B/C row as header row.
- Groups by A with unique B/C pair rows under each A.
- Computes:
  - Entries Per Location
  - Number Passed/Failed breakdown from K
  - Sequence-based attempt metrics from G + K
  - RLU analytics from E
  - Limits interpretation from F
  - Data Integrity Risk details with alarming sequence text

### Tabs in UI

- Mapped View
- Calibration Review (O)

### Calibration Review (O)

- Groups by O value.
- Uses E/G/K signals for scoring:
  - pass rate
  - average attempts to pass
  - RLU mean/std dev/CV
  - outlier rate
  - sequence spread
- Produces a risk score + assessment label.

## Export Behavior (Implemented)

- Export button enabled after successful file load.
- Exports a styled workbook with 2 sheets:
  - Mapped View
  - Calibration Review
- Save behavior:
  - Uses native file picker (`showSaveFilePicker`) when browser supports it (name + location dialog).
  - Falls back to filename prompt + browser download behavior when unsupported.

### Mapped Sheet Styling

- Header styling + frozen top row.
- Wrapped text and row spacing preserved.
- Column B auto-width based on widest B value (with min/max bounds).
- Column F (analytics text column) alignment:
  - horizontal center
  - vertical center
  - wrap enabled
- Dynamic row heights for mapped rows based on wrapped text estimate (prevents clipping of long analytics/limits content).
- Status colors applied to mapped result rows across B:E cells.
- Legend key written in column H with matching colors and labels:
  - No pass in sequence
  - Average attempts > 4
  - Average attempts > 2
  - Average attempts > 1
  - Average attempts = 1

### Calibration Sheet Styling

- Header styling + frozen top row.
- Border and row styling consistent with mapped export style.

## Recent Terminology/UX Changes

- Analytics text label updated from "K Avg" to "RLU Avg by K".
- Added export button in UI.
- Added tabbed report UI (Mapped + Calibration).

## Key Functions To Know (src/main.js)

- mapColumnAtoBAndCValues
- renderMappedResults
- buildCalibrationByO
- renderCalibrationResults
- getAverageAttemptsToPass
- getIntegrityRiskDetails
- exportReportWorkbook
- getAutoColumnWidth
- applyMappedRowHeights
- addMappedColorKeyInColumnH

## Validation Status

- Repeated `npm run build` checks completed successfully after each major change.
- Current known warning: large chunk warning due to exceljs bundle split; functional export is working.

## Session Timeline (Condensed Log)

- Built from a simple A-value mapping tool into a full two-tab analytics app.
- Added broader input format support and mapping refinements for A/B/C grouping.
- Added Entries Per Location and K-based pass/fail percentages.
- Added sequence-based row coloring from G+K behavior.
- Added E/F/K analytics text with anomaly and integrity-risk logic.
- Added Calibration Review tab grouped by O with E/G/K scoring.
- Added styled Excel export with two matching sheets.
- Added save dialog support (`showSaveFilePicker`) with prompt/download fallback.
- Added export layout polish:
  - Column B auto-width from data
  - Dynamic mapped row heights for wrapped content
  - Column F centered horizontally and vertically
  - Color legend key in column H

## Known Constraints

- Exact Excel visual rendering may still vary slightly by Excel version/font metrics.
- In fallback download mode (no File System Access API), folder location is controlled by browser settings.

## Fast Resume Checklist

1. Open src/main.js and GH_COPILOT_CONTEXT.md.
2. Run `npm install` (if needed) and `npm run dev`.
3. Load a representative source file and verify both tabs.
4. Export workbook and check:
   - Mapped View styling
   - Column B auto-width
   - Column F centered + wrapped + proper height
   - Column H legend
   - Calibration sheet populated
5. Run `npm run build` before handoff.

## Suggested Next Enhancements

- Optional export profile presets (default filename/template).
- Optional CSV export for each tab.
- Optional tuning controls for calibration risk thresholds.
- Optional tiny QA panel showing parsed row counts by source columns.

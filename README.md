# Excel Report Generator

A Vite + React application that ingests an Excel (`.xlsx`) file and displays the **unique values from Column A** directly on the page.

## Features

- 📂 File input — pick any `.xlsx` file from your machine
- 🔍 Parses the uploaded workbook entirely in the browser using [`read-excel-file`](https://www.npmjs.com/package/read-excel-file)
- ✅ Deduplicates all non-empty Column A values and shows them in an ordered list with a count badge
- ⚡ Loading indicator while large files are being processed
- 🛑 User-friendly error message for invalid / unsupported files

## Getting Started

```bash
npm install
npm run dev
```

Then open http://localhost:5173 in your browser, choose an `.xlsx` file, and the unique Column A values will appear on the page.

## Build

```bash
npm run build
```

The production-ready bundle is output to the `dist/` folder.


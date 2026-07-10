# Route Sales Report Builder

A browser-based tool that turns a **Route Wise Item Wise Outlet Wise Sales Report** (the master `.xlsx` export from your sales system) into a clean, branded, per-route **"Last Visit Outlet Wise Quantity Sales"** report — ready to download as a styled Excel file or a print-ready PDF.

Everything runs client-side in the browser. No backend, no server, no data leaves your machine.

---

## What it does

1. You upload the master sales report (`.xlsx`)
2. The app auto-detects every **Territory** and **Route** in the file
3. You pick a Territory + Route, choose a Last Visit Date, and pick which "last visit" rule to use
4. It builds an outlet-by-outlet, product-by-product breakdown showing quantities from the customer's most recent visit
5. You download it as a formatted Excel workbook (bordered table, bold header, merged outlet cells) or a PDF (via the browser's print dialog)

There's also a **Template Editor** screen where you can customize all the labels, column headers, report title, and brand accent color — these settings are saved in your browser and reused every time you generate a report.

---

## Features

- 📂 **Multi-file upload** — drag-and-drop or select several `.xlsx` files at once, each processed independently
- 🔍 Territory and Route dropdowns **auto-extracted** from each uploaded file — no manual typing
- 📅 A **calendar date picker** for the Last Visit Date, one per uploaded file (each file can have a different visit date)
- 🗂️ Click any uploaded file's card to make it "active" — that's the one shown in the main preview and exported by the header buttons
- ⚙️ Two strategies for resolving "last visit" quantity, set independently per file:
  - **Most recent visit with data** — scans PREVIOUS 1ST → 4TH VISIT columns and uses the first one with a quantity
  - **Strict** — only uses the PREVIOUS 1ST VISIT column
- 📊 Live stats: total sale value, outlet count, rows matched
- 📄 Paginated outlet table (4 outlets per page)
- 🖨️ Excel export with bold colored headers, cell borders, and merged outlet/row-number cells
- 🖨️ PDF export via a styled print view
- 📦 **Export All as Excel** — with multiple files uploaded, generate a separate styled workbook for every file that has valid route data, in one click
- 🎨 Row banding — optional alternating light-colored shading per outlet group (on by default), with 7 selectable light tones, applied consistently to the on-screen table, Excel export, and PDF export
- 🎨 Template Editor with 11 accent color swatches (including light tones) and fully editable labels/column headers, with a live preview
- 💾 Template settings persist across sessions via `localStorage`

---

## Tech stack

| Purpose | Library |
|---|---|
| UI framework | React (Vite) |
| Styling | Tailwind CSS v4 (`@tailwindcss/vite`) |
| Reading uploaded `.xlsx` | [`xlsx`](https://www.npmjs.com/package/xlsx) (SheetJS) |
| Writing styled `.xlsx` exports | [`exceljs`](https://www.npmjs.com/package/exceljs) |
| Icons | Material Symbols Outlined (Google Fonts) |
| Fonts | Manrope (headings), Work Sans (body), IBM Plex Sans (small caps labels) |

---

## Project setup (from scratch)

### 1. Create the Vite project

```bash
npm create vite@latest route-report-app -- --template react
cd route-report-app
npm install
```

### 2. Install Tailwind CSS v4

```bash
npm install -D tailwindcss @tailwindcss/vite
```

Update `vite.config.js`:

```js
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

export default defineConfig({
  plugins: [react(), tailwindcss()],
})
```

Replace the contents of `src/index.css` with:

```css
@import "tailwindcss";

.material-symbols-outlined {
  font-family: 'Material Symbols Outlined';
  font-variation-settings: 'FILL' 0, 'wght' 400, 'GRAD' 0, 'opsz' 24;
  vertical-align: middle;
}
```

### 3. Install the remaining packages

```bash
npm install xlsx exceljs
```

### 4. Add the fonts

Open `index.html` and add these inside `<head>`, below `<title>`:

```html
<link rel="preconnect" href="https://fonts.googleapis.com" />
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@600;700&family=Work+Sans:wght@400;500&family=IBM+Plex+Sans:wght@600&display=swap" rel="stylesheet" />
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:wght,FILL@100..700,0..1&display=swap" rel="stylesheet" />
```

Your full `index.html` should look like:

```html
<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <link rel="icon" type="image/svg+xml" href="/favicon.svg" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Route Sales Report Builder</title>

    <link rel="preconnect" href="https://fonts.googleapis.com" />
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@600;700&family=Work+Sans:wght@400;500&family=IBM+Plex+Sans:wght@600&display=swap" rel="stylesheet" />
    <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:wght,FILL@100..700,0..1&display=swap" rel="stylesheet" />
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.jsx"></script>
  </body>
</html>
```

### 5. Add the app code and logo

- Replace `src/App.jsx` with the project's `App.jsx`
- Place your brand logo at `public/logo.png`
- Confirm `src/main.jsx` imports the stylesheet: `import './index.css'`

### 6. Run it

```bash
npm run dev
```

Open the printed `http://localhost:5173` link.

### 7. (Recommended) Put it under version control

```bash
git init
git add .
git commit -m "Initial working version"
```

Commit again any time you make a change you're happy with — it gives you an instant rollback point if a future edit (yours, mine, or an AI tool's) breaks something:

```bash
git add .
git commit -m "describe what changed"
```

To undo uncommitted changes back to the last commit:

```bash
git checkout -- src/App.jsx
```

---

## How to use it

### Dashboard

1. **Step 1** — drop one or more Route Wise Item Wise Outlet Wise Sales Report `.xlsx` files into the upload box (or click to browse and multi-select). Each file appears as its own card below.
2. For each file's card, set:
   - **Territory** — the Route dropdown updates to only show routes within that territory
   - **Route**
   - **Last Visit Date** — via the calendar picker (click the field to open the month grid; each file keeps its own date)
   - **Which visit counts as "last"** — *Most recent visit with data* (recommended default) handles outlets that were visited on a fixed cycle but skipped in recent rounds; *Only PREVIOUS 1ST VISIT column* is stricter
3. Click a file's card to make it the **active** file — its outlet/product breakdown appears in the main panel on the right
4. Click **Download Excel** or **Download PDF** to export the active file's report
5. With two or more files uploaded, a **"Export All as Excel"** button appears in the sidebar — click it to generate a separate styled workbook for every file that has valid route data (downloads one after another)
6. Use the **×** on any card to remove that file, or **Reset Filters (active file)** to reset just the currently active file's territory/route/date/strategy back to defaults

### Template Editor

Switch to this tab to customize:
- Report title
- Territory / Route / Total sale value / Last visit date labels
- The 4 table column headers
- The brand accent color (11 preset swatches)
- **Outlet row banding** — toggle on/off, and pick from 7 light shading colors used to alternate-highlight each outlet's product group (matches the table, Excel export, and PDF export)

Click **Save Template** — this is stored in your browser and will be used automatically every time you generate a report going forward. **Reset to defaults** reverts everything back to the original Cherry Alloy theme and default labels.

---

## Understanding the source file format

The uploaded master report must have a header row where column A = `NO` and column E = `TERRITORY`. Below that is one sub-header row, then the data.

Each data row represents one product for one outlet visit cycle, with columns:

| Columns (0-indexed) | Field |
|---|---|
| 0 | NO |
| 1–3 | RSM / ASM / ASE |
| 4 | TERRITORY |
| 5 | ROUTE NUMBER |
| 6 | ROUTE NAME |
| 7 | OUTLET ID |
| 8 | OUTLET NAME |
| 9 | PRODUCT |
| 10–16 | PREVIOUS 4TH VISIT (Qty, Invoice Value, Return Qty, Return Value, Free Qty, Free Value, Net Sales) |
| 17–23 | PREVIOUS 3RD VISIT (same 7 sub-columns) |
| 24–30 | PREVIOUS 2ND VISIT (same 7 sub-columns) |
| 31–37 | PREVIOUS 1ST VISIT (same 7 sub-columns) |

NO, RSM, ASM, ASE, TERRITORY, ROUTE NUMBER, ROUTE NAME, OUTLET ID, and OUTLET NAME are only filled in on the **first** product row of each outlet — the app automatically carries these values down through the blank rows that follow, so each product row knows which outlet it belongs to.

---

## Known limitations

- **PDF export** relies on the browser's native print-to-PDF — formatting will look slightly different depending on which browser/print settings are used
- **Excel styling** requires `exceljs`; the lighter-weight `xlsx` library alone cannot write cell borders/fills
- Template settings are stored in `localStorage`, which is per-browser — they won't carry over if you switch browsers or computers
- No login/user accounts — this is a single-user local tool by design

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| `npm error Missing script: "dev"` | You're running the command from the wrong folder — `cd` into `route-report-app` first (the folder containing `package.json` with a `dev` script) |
| Blank page / page flashes then disappears | Open DevTools Console (Ctrl+Shift+I) and check for a red error — usually a leftover reference to an undefined variable or a stale copy of `App.jsx` |
| Styling looks unstyled/plain | Confirm `@tailwindcss/vite` is in `vite.config.js` plugins and `src/index.css` starts with `@import "tailwindcss";` |
| Icons show as text instead of symbols | Confirm the Material Symbols Outlined font link is in `index.html` and the `.material-symbols-outlined` CSS rule is in `index.css` |
| Old UI still showing after an update | Hard refresh (Ctrl+Shift+R), or stop the dev server and run `npm run dev` again; as a last resort delete `node_modules/.vite` to clear Vite's cache |
| "No rows had quantity" for a route | Try switching the last-visit strategy between "Most recent visit with data" and "Only PREVIOUS 1ST VISIT column" |
| A file's card shows a red error message | The uploaded file isn't in the expected Route Wise Item Wise Outlet Wise Sales Report layout — confirm it has a header row with `NO` in column A and `TERRITORY` in column E |
| "Export All as Excel" skipped a file | That file either had no valid Territory/Route selected, or its report had zero matching outlets for the selected strategy — check that file's card individually first |

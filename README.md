# Atlas Sheet Viewer

Static spreadsheet viewer built with React, Vite, and SheetJS for GitHub Pages deployment.

## What it does

- Opens local `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.csv`, `.tsv`, and `.ods` files.
- Parses workbooks inside a Web Worker so the UI stays responsive.
- Loads one sheet at a time and requests visible cell windows on demand.
- Renders the sheet through a virtualized grid instead of painting the full workbook DOM.
- Includes search for the active sheet, workbook metadata, and a sheet switcher.

## Large file strategy

This app is designed to stay usable with bigger spreadsheets, including files larger than 50 MB, by combining:

- worker-based parsing
- lazy sheet loading
- tiled cell-window fetching
- row and column virtualization

There is still a hard browser-memory ceiling. Very wide sheets, heavily formatted workbooks, or extremely dense files can still hit client-side limits because GitHub Pages is a static host and all parsing happens in the browser.

## Development

```bash
npm install
npm run dev
```

Production build:

```bash
npm run build
```

Preview the built app locally:

```bash
npm run preview
```

## Deploy to GitHub Pages

The repo includes [`.github/workflows/deploy.yml`](.github/workflows/deploy.yml), which builds the Vite app and deploys `dist/` to GitHub Pages.

To enable deployment:

1. Push the repository to GitHub.
2. Make sure the default branch is `main`, or update the workflow branch filter.
3. In GitHub, set Pages to use **GitHub Actions** as the source.
4. Push to `main` or run the workflow manually.

The Vite config uses `base: './'`, so the generated app works on project pages without hardcoding a repository name.

## Notes

- Search currently scans the active sheet only.
- Files are never uploaded anywhere by this app.
- The `xlsx` package currently reports one high-severity advisory in `npm audit`; review that before using this viewer for untrusted workbook workflows.

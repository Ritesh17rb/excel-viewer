# Atlas Sheet Viewer

Server-backed Excel viewer for large `.xlsx` files, deployed on Vercel.

Live app: `https://excel-viewer-fawn.vercel.app`

## What it does

- Visible upload button plus drag-and-drop for `.xlsx` files
- Direct browser uploads to Vercel Blob in production
- Streaming workbook ingestion with `exceljs`
- Server-side paging so the browser does not hold the whole workbook in memory
- Virtualized spreadsheet grid
- Server-backed search with jump-to-cell results
- Server-backed sorting
- Server-backed text filters
- Header-row mode for sort/filter operations

## Architecture

- Next.js App Router frontend and API routes
- `@vercel/blob` for uploaded workbooks and processed artifacts
- `exceljs` streaming reader for ingestion
- Server-generated paged workbook artifacts
- Derived sheet views for sort/filter operations

This project is intentionally not a browser-only parser. Large workbooks are parsed on the server, then served back to the UI as lightweight row pages and derived views.

## Current limitations

- Read-only viewer: no editing or saving workbook changes
- `.xlsx` only in the server-backed ingestion path
- No formula recalculation engine
- No charts, pivot tables, comments, or macros UI
- Formatting fidelity is limited compared with desktop Excel

## Development

Install dependencies and start the app:

```bash
npm install
npm run dev
```

Build:

```bash
npm run build
```

Lint:

```bash
npm run lint
```

Sample workbook ingestion test:

```bash
npm run test:sample -- /absolute/path/to/file.xlsx
```

## Deployment

The app is configured for Vercel.

Required setup:

1. Create or import the project in Vercel.
2. Connect a Vercel Blob store to the project.
3. Set `BLOB_ACCESS=private` or `BLOB_ACCESS=public`.
4. Ensure `BLOB_READ_WRITE_TOKEN` is available in the project environment.
5. Deploy.

This repo already includes [vercel.json](/home/ritesh/work/excel-viewer/vercel.json) so Vercel treats the project as a Next.js app.

## Important files

- [components/workbook-viewer.tsx](/home/ritesh/work/excel-viewer/components/workbook-viewer.tsx)
- [components/spreadsheet-grid.tsx](/home/ritesh/work/excel-viewer/components/spreadsheet-grid.tsx)
- [lib/server/workbook-ingest.ts](/home/ritesh/work/excel-viewer/lib/server/workbook-ingest.ts)
- [lib/server/workbook-views.ts](/home/ritesh/work/excel-viewer/lib/server/workbook-views.ts)
- [app/api/workbooks/[id]/search/route.ts](/home/ritesh/work/excel-viewer/app/api/workbooks/[id]/search/route.ts)
- [app/api/workbooks/[id]/view/route.ts](/home/ritesh/work/excel-viewer/app/api/workbooks/[id]/view/route.ts)

## Why this exists

The earlier GitHub Pages/browser-worker version was not reliable for large files like the 30 MB workbook tested in this repo. This version moves ingestion and heavy workbook operations to the server so large sheets remain usable in the browser.

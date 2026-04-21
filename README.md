# Atlas Sheet Viewer

Server-backed Excel viewer for large `.xlsx` files, built for Vercel.

## Architecture

- Next.js App Router frontend and API routes
- Direct browser uploads to Vercel Blob in deployed environments
- Streaming `.xlsx` ingestion with `exceljs`
- Server-side chunking into row pages
- Virtualized client grid that fetches only visible page data

This is intentionally not a browser-only parser anymore. The browser never tries to hold the full workbook in memory.

## Current scope

- Optimized for large `.xlsx` files
- Local development supports a direct local upload route
- Deployed environments should use Vercel Blob uploads
- Active-sheet search is currently disabled in the server-backed path

## Development

```bash
npm install
npm run dev
```

Production build:

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

## Vercel deployment

1. Push the repo to GitHub.
2. Import the repo into Vercel.
3. Create a Vercel Blob store and connect it to the project.
4. Set `BLOB_ACCESS=public` or `BLOB_ACCESS=private`.
5. Deploy.

If you use a private Blob store, keep `BLOB_READ_WRITE_TOKEN` available in the Vercel project environment.

## Notes

- The server-backed parser currently supports `.xlsx`.
- The old GitHub Pages worker-based approach was removed because it was not reliable for the 30 MB sample workbook.

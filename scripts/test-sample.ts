import { ingestWorkbook } from '../lib/server/workbook-ingest'

const samplePath = process.argv[2] ?? '/home/ritesh/work/excel-viewer/30mb.xlsx'

const start = performance.now()

const manifest = await ingestWorkbook({
  kind: 'local',
  fileName: samplePath.split('/').at(-1) ?? 'sample.xlsx',
  fileSize: 0,
  format: 'xlsx',
  localPath: samplePath,
})

console.log(
  JSON.stringify(
    {
      elapsedMs: Number((performance.now() - start).toFixed(1)),
      workbookId: manifest.id,
      sheets: manifest.sheetNames.map((sheetName) => manifest.sheets[sheetName]),
    },
    null,
    2,
  ),
)

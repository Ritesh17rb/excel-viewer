import { randomUUID } from 'node:crypto'

import ExcelJS from 'exceljs'

import { clamp, sanitizeSegment, toColumnLabel } from '@/lib/format'
import type {
  WorkbookManifest,
  WorkbookSheetManifest,
  WorkbookSheetPage,
  WorkbookUploadSource,
} from '@/lib/workbook-types'
import { openSourceStream, putWorkbookJson } from '@/lib/server/workbook-store'

const PAGE_SIZE = 500
const DEFAULT_COLUMN_WIDTH = 148
const MIN_COLUMN_WIDTH = 112
const MAX_COLUMN_WIDTH = 320
const WIDTH_SAMPLE_ROWS = 80
const WIDTH_SAMPLE_COLUMNS = 64

export async function ingestWorkbook(
  source: WorkbookUploadSource,
): Promise<WorkbookManifest> {
  if (source.format !== 'xlsx') {
    throw new Error('The server-backed viewer currently supports .xlsx files only.')
  }

  const workbookId = randomUUID()
  const manifest: WorkbookManifest = {
    id: workbookId,
    createdAt: new Date().toISOString(),
    source,
    performanceMode: 'server',
    sheetNames: [],
    sheets: {},
  }

  const input = await openSourceStream(source)
  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(input, {
    worksheets: 'emit',
    sharedStrings: 'cache',
    hyperlinks: 'ignore',
    styles: 'ignore',
  })

  let sheetIndex = 0

  for await (const worksheet of workbookReader) {
    const sheetName =
      (worksheet as unknown as { name?: string }).name ?? `Sheet ${sheetIndex + 1}`
    const sheetKey = `${String(sheetIndex + 1).padStart(2, '0')}-${sanitizeSegment(sheetName)}`
    const columnWidths: number[] = []
    const pagePaths: string[] = []
    const rowBuffer: string[][] = []

    let rowCount = 0
    let columnCount = 0
    let populatedCellCount = 0
    let nextExpectedRow = 1
    let pageStartRow = 0

    for await (const row of worksheet) {
      rowCount = Math.max(rowCount, row.number)

      while (nextExpectedRow < row.number) {
        pageStartRow = pushRow(rowBuffer, pageStartRow, nextExpectedRow - 1, [])
        nextExpectedRow += 1

        if (rowBuffer.length === PAGE_SIZE) {
          pagePaths.push(
            await flushPage({
              workbookId,
              sheetKey,
              sheetName,
              page: pagePaths.length,
              rowStart: pageStartRow,
              rowEnd: pageStartRow + rowBuffer.length,
              rows: rowBuffer.splice(0, rowBuffer.length),
            }),
          )
          pageStartRow = nextExpectedRow - 1
        }
      }

      const rowValues = Array.isArray(row.values) ? row.values : []
      const lastColumn = Math.max(0, rowValues.length - 1)
      const values: string[] = []

      for (let columnIndex = 1; columnIndex <= lastColumn; columnIndex += 1) {
        const value = stringifyCell(rowValues[columnIndex])

        if (value !== '') {
          populatedCellCount += 1
        }

        values.push(value)

        if (row.number <= WIDTH_SAMPLE_ROWS && columnIndex <= WIDTH_SAMPLE_COLUMNS) {
          const width = clamp(
            value
              ? value.slice(0, 48).length * 8.4 + 34
              : toColumnLabel(columnIndex - 1).length * 18 + 34,
            MIN_COLUMN_WIDTH,
            MAX_COLUMN_WIDTH,
          )
          columnWidths[columnIndex - 1] = Math.max(
            columnWidths[columnIndex - 1] ?? DEFAULT_COLUMN_WIDTH,
            width,
          )
        }
      }

      columnCount = Math.max(columnCount, values.length)
      pageStartRow = pushRow(
        rowBuffer,
        pageStartRow,
        row.number - 1,
        trimTrailingEmpty(values),
      )
      nextExpectedRow = row.number + 1

      if (rowBuffer.length === PAGE_SIZE) {
        pagePaths.push(
          await flushPage({
            workbookId,
            sheetKey,
            sheetName,
            page: pagePaths.length,
            rowStart: pageStartRow,
            rowEnd: pageStartRow + rowBuffer.length,
            rows: rowBuffer.splice(0, rowBuffer.length),
          }),
        )
        pageStartRow = nextExpectedRow - 1
      }
    }

    if (rowBuffer.length > 0) {
      pagePaths.push(
        await flushPage({
          workbookId,
          sheetKey,
          sheetName,
          page: pagePaths.length,
          rowStart: pageStartRow,
          rowEnd: pageStartRow + rowBuffer.length,
          rows: rowBuffer.splice(0, rowBuffer.length),
        }),
      )
    }

    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      columnWidths[columnIndex] = Math.max(
        columnWidths[columnIndex] ?? DEFAULT_COLUMN_WIDTH,
        clamp(
          toColumnLabel(columnIndex).length * 18 + 34,
          MIN_COLUMN_WIDTH,
          DEFAULT_COLUMN_WIDTH,
        ),
      )
    }

    const sheetManifest: WorkbookSheetManifest = {
      key: sheetKey,
      name: sheetName,
      rowCount,
      columnCount,
      populatedCellCount,
      pageSize: PAGE_SIZE,
      pageCount: pagePaths.length,
      columnWidths,
      pages: pagePaths,
      searchEnabled: true,
      headerRows: 0,
    }

    manifest.sheetNames.push(sheetName)
    manifest.sheets[sheetName] = sheetManifest
    sheetIndex += 1
  }

  await putWorkbookJson(`${manifest.id}/manifest.json`, manifest)

  return manifest
}

async function flushPage({
  workbookId,
  sheetKey,
  sheetName,
  page,
  rowStart,
  rowEnd,
  rows,
}: WorkbookSheetPage & {
  workbookId: string
  sheetKey: string
}): Promise<string> {
  const relativePath = `${workbookId}/${sheetKey}/page-${page}.json`

  await putWorkbookJson(relativePath, {
    sheetName,
    page,
    rowStart,
    rowEnd,
    rows,
  })

  return relativePath
}

function stringifyCell(value: unknown): string {
  if (value == null) {
    return ''
  }

  if (typeof value === 'string') {
    return value
  }

  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }

  if (value instanceof Date) {
    return value.toISOString()
  }

  if (typeof value === 'object' && value && 'text' in value) {
    const text = (value as { text?: string }).text
    return typeof text === 'string' ? text : ''
  }

  return String(value)
}

function trimTrailingEmpty(values: string[]): string[] {
  let last = values.length - 1

  while (last >= 0 && values[last] === '') {
    last -= 1
  }

  return values.slice(0, last + 1)
}

function pushRow(
  rowBuffer: string[][],
  pageStartRow: number,
  rowIndex: number,
  values: string[],
): number {
  if (rowBuffer.length === 0) {
    rowBuffer.push(values)
    return rowIndex
  }

  rowBuffer.push(values)
  return pageStartRow
}

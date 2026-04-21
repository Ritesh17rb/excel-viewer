/// <reference lib="webworker" />

import * as XLSX from 'xlsx'

import { clamp, toColumnLabel } from '../lib/format'
import type {
  SearchResult,
  SheetMetadata,
  WorkerRequest,
  WorkerResponse,
} from '../types/workbook'

const workerScope = self as DedicatedWorkerGlobalScope

const DEFAULT_COLUMN_WIDTH = 148
const MAX_COLUMN_WIDTH = 320
const MIN_COLUMN_WIDTH = 112
const STANDARD_WIDTH_SAMPLE_ROWS = 140
const STANDARD_WIDTH_SAMPLE_COLUMNS = 220
const LARGE_WIDTH_SAMPLE_ROWS = 60
const LARGE_WIDTH_SAMPLE_COLUMNS = 90
const LARGE_FILE_SIZE_BYTES = 18 * 1024 * 1024
const SEARCH_POPULATED_CELL_LIMIT = 250_000
const SEARCH_GRID_CELL_LIMIT = 5_000_000
const LARGE_SHEET_POPULATED_CELL_LIMIT = 350_000
const LARGE_SHEET_GRID_CELL_LIMIT = 8_000_000

let sourceBuffer: ArrayBuffer | null = null
let activeSheetName: string | null = null
let activeWorksheet: XLSX.WorkSheet | null = null
let workbookPerformanceMode: 'standard' | 'large' = 'standard'
const sheetMetadataCache = new Map<string, SheetMetadata>()

workerScope.addEventListener('message', (event: MessageEvent<WorkerRequest>) => {
  try {
    handleMessage(event.data)
  } catch (error) {
    post({
      type: 'error',
      message:
        error instanceof Error
          ? error.message
          : 'The spreadsheet could not be processed.',
    })
  }
})

function handleMessage(message: WorkerRequest): void {
  switch (message.type) {
    case 'load-workbook':
      loadWorkbook(message)
      return
    case 'load-sheet':
      loadSheet(message.sheetName)
      return
    case 'load-window':
      loadWindow(message)
      return
    case 'search-sheet':
      searchSheet(message.sheetName, message.query, message.limit)
      return
  }
}

function loadWorkbook(message: Extract<WorkerRequest, { type: 'load-workbook' }>) {
  sourceBuffer = message.buffer
  activeSheetName = null
  activeWorksheet = null
  sheetMetadataCache.clear()

  const workbook = XLSX.read(sourceBuffer, {
    type: 'array',
    bookProps: true,
    bookSheets: true,
    WTF: false,
  })

  workbookPerformanceMode =
    message.fileSize >= LARGE_FILE_SIZE_BYTES ? 'large' : 'standard'

  post({
    type: 'workbook-loaded',
    workbook: {
      fileName: message.fileName,
      fileSize: message.fileSize,
      format: message.format,
      loadedAt: new Date().toISOString(),
      sheetNames: workbook.SheetNames,
      performanceMode: workbookPerformanceMode,
    },
  })
}

function loadSheet(sheetName: string): void {
  if (!sourceBuffer) {
    throw new Error('Open a workbook before loading a sheet.')
  }

  const workbook = XLSX.read(sourceBuffer, {
    type: 'array',
    sheets: sheetName,
    raw: workbookPerformanceMode === 'large',
    cellFormula: false,
    cellHTML: false,
    cellNF: false,
    cellText: workbookPerformanceMode !== 'large',
    cellStyles: false,
    sheetStubs: false,
    WTF: false,
  })

  const worksheet = workbook.Sheets[sheetName]

  if (!worksheet) {
    throw new Error(`Sheet "${sheetName}" could not be loaded.`)
  }

  activeSheetName = sheetName
  activeWorksheet = worksheet

  const metadata = createSheetMetadata(sheetName, worksheet)
  sheetMetadataCache.set(sheetName, metadata)

  post({
    type: 'sheet-loaded',
    sheet: metadata,
  })
}

function loadWindow(
  message: Extract<WorkerRequest, { type: 'load-window' }>,
): void {
  const worksheet = ensureSheet(message.sheetName)
  const metadata = sheetMetadataCache.get(message.sheetName)

  if (!metadata) {
    throw new Error(`Sheet "${message.sheetName}" has no metadata.`)
  }

  const rowStart = clamp(message.rowStart, 0, metadata.rowCount)
  const rowEnd = clamp(Math.max(message.rowEnd, rowStart), rowStart, metadata.rowCount)
  const colStart = clamp(message.colStart, 0, metadata.columnCount)
  const colEnd = clamp(
    Math.max(message.colEnd, colStart),
    colStart,
    metadata.columnCount,
  )

  post({
    type: 'window-loaded',
    window: {
      key: message.key,
      sheetName: message.sheetName,
      rowStart,
      rowEnd,
      colStart,
      colEnd,
      values: extractWindow(worksheet, rowStart, rowEnd, colStart, colEnd),
    },
  })
}

function searchSheet(sheetName: string, query: string, limit: number): void {
  const normalizedQuery = query.trim().toLocaleLowerCase()

  if (!normalizedQuery) {
    post({
      type: 'search-results',
      sheetName,
      query,
      results: [],
      disabledReason: null,
    })
    return
  }

  const worksheet = ensureSheet(sheetName)
  const metadata = sheetMetadataCache.get(sheetName)

  if (!metadata) {
    throw new Error(`Sheet "${sheetName}" has no metadata.`)
  }

  if (!metadata.searchEnabled) {
    post({
      type: 'search-results',
      sheetName,
      query,
      results: [],
      disabledReason: metadata.searchDisabledReason,
    })
    return
  }

  const results: SearchResult[] = []

  for (const address in worksheet) {
    if (isMetadataKey(address)) {
      continue
    }

    const value = formatCell(worksheet[address] as XLSX.CellObject | undefined)

    if (!value || !value.toLocaleLowerCase().includes(normalizedQuery)) {
      continue
    }

    const position = XLSX.utils.decode_cell(address)

    results.push({
      rowIndex: position.r,
      columnIndex: position.c,
      address: `${toColumnLabel(position.c)}${position.r + 1}`,
      value,
    })

    if (results.length >= limit) {
      break
    }
  }

  post({
    type: 'search-results',
    sheetName,
    query,
    results,
    disabledReason: null,
  })
}

function ensureSheet(sheetName: string): XLSX.WorkSheet {
  if (activeSheetName === sheetName && activeWorksheet) {
    return activeWorksheet
  }

  loadSheet(sheetName)

  if (activeSheetName !== sheetName || !activeWorksheet) {
    throw new Error(`Sheet "${sheetName}" is unavailable.`)
  }

  return activeWorksheet
}

function createSheetMetadata(
  sheetName: string,
  worksheet: XLSX.WorkSheet,
): SheetMetadata {
  const ref = worksheet['!ref']

  if (!ref) {
    return {
      name: sheetName,
      rowCount: 0,
      columnCount: 0,
      range: null,
      columnWidths: [],
      populatedCellCount: 0,
      largeSheetMode: workbookPerformanceMode === 'large',
      searchEnabled: true,
      searchDisabledReason: null,
    }
  }

  const range = XLSX.utils.decode_range(ref)
  const rowCount = range.e.r + 1
  const columnCount = range.e.c + 1
  const gridCellCount = rowCount * columnCount
  const columnWidths = Array.from({ length: columnCount }, () => DEFAULT_COLUMN_WIDTH)
  const widthSampleRows =
    workbookPerformanceMode === 'large'
      ? LARGE_WIDTH_SAMPLE_ROWS
      : STANDARD_WIDTH_SAMPLE_ROWS
  const widthSampleColumns =
    workbookPerformanceMode === 'large'
      ? LARGE_WIDTH_SAMPLE_COLUMNS
      : STANDARD_WIDTH_SAMPLE_COLUMNS

  const sampleRowEnd = Math.min(range.e.r, widthSampleRows - 1)
  const sampleColumnEnd = Math.min(range.e.c, widthSampleColumns - 1)

  for (let columnIndex = 0; columnIndex <= sampleColumnEnd; columnIndex += 1) {
    columnWidths[columnIndex] = clamp(
      toColumnLabel(columnIndex).length * 18 + 34,
      MIN_COLUMN_WIDTH,
      DEFAULT_COLUMN_WIDTH,
    )
  }

  let populatedCellCount = 0

  for (const address in worksheet) {
    if (isMetadataKey(address)) {
      continue
    }

    populatedCellCount += 1

    const position = XLSX.utils.decode_cell(address)

    if (position.r > sampleRowEnd || position.c > sampleColumnEnd) {
      continue
    }

    const value = formatCell(worksheet[address] as XLSX.CellObject | undefined)

    if (!value) {
      continue
    }

    columnWidths[position.c] = Math.max(
      columnWidths[position.c],
      clamp(value.slice(0, 48).length * 8.4 + 34, MIN_COLUMN_WIDTH, MAX_COLUMN_WIDTH),
    )
  }

  const largeSheetMode =
    workbookPerformanceMode === 'large' ||
    populatedCellCount > LARGE_SHEET_POPULATED_CELL_LIMIT ||
    gridCellCount > LARGE_SHEET_GRID_CELL_LIMIT
  const searchEnabled =
    populatedCellCount <= SEARCH_POPULATED_CELL_LIMIT &&
    gridCellCount <= SEARCH_GRID_CELL_LIMIT
  const searchDisabledReason =
    populatedCellCount === 0
      ? null
      : searchEnabled
        ? null
        : 'Search is disabled for this sheet to keep the browser responsive in large-file mode.'

  return {
    name: sheetName,
    rowCount,
    columnCount,
    range: ref,
    columnWidths,
    populatedCellCount,
    largeSheetMode,
    searchEnabled,
    searchDisabledReason,
  }
}

function extractWindow(
  worksheet: XLSX.WorkSheet,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): string[][] {
  const rows: string[][] = []

  for (let rowIndex = rowStart; rowIndex < rowEnd; rowIndex += 1) {
    const row: string[] = []

    for (let columnIndex = colStart; columnIndex < colEnd; columnIndex += 1) {
      row.push(formatCell(getCell(worksheet, rowIndex, columnIndex)))
    }

    rows.push(row)
  }

  return rows
}

function getCell(
  worksheet: XLSX.WorkSheet,
  rowIndex: number,
  columnIndex: number,
): XLSX.CellObject | undefined {
  const address = XLSX.utils.encode_cell({
    r: rowIndex,
    c: columnIndex,
  })

  return worksheet[address] as XLSX.CellObject | undefined
}

function formatCell(cell: XLSX.CellObject | undefined): string {
  if (!cell) {
    return ''
  }

  if (typeof cell.w === 'string' && cell.w.length > 0) {
    return cell.w
  }

  if (typeof cell.v === 'string') {
    return cell.v
  }

  if (cell.v == null && typeof cell.f === 'string' && cell.f.length > 0) {
    return `=${cell.f}`
  }

  return cell.v == null ? '' : String(cell.v)
}

function isMetadataKey(address: string): boolean {
  return address.startsWith('!')
}

function post(message: WorkerResponse): void {
  workerScope.postMessage(message)
}

export {}

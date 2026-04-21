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
const WIDTH_SAMPLE_ROWS = 140
const WIDTH_SAMPLE_COLUMNS = 220

let sourceBuffer: ArrayBuffer | null = null
let workbookSheetNames: string[] = []
let activeSheetName: string | null = null
let activeWorksheet: XLSX.WorkSheet | null = null
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

  workbookSheetNames = workbook.SheetNames

  post({
    type: 'workbook-loaded',
    workbook: {
      fileName: message.fileName,
      fileSize: message.fileSize,
      format: message.format,
      loadedAt: new Date().toISOString(),
      sheetNames: workbookSheetNames,
    },
  })
}

function loadSheet(sheetName: string): void {
  if (!sourceBuffer) {
    throw new Error('Open a workbook before loading a sheet.')
  }

  const workbook = XLSX.read(sourceBuffer, {
    type: 'array',
    dense: true,
    sheets: sheetName,
    raw: false,
    cellFormula: false,
    cellHTML: false,
    cellNF: false,
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
    })
    return
  }

  const worksheet = ensureSheet(sheetName)
  const metadata = sheetMetadataCache.get(sheetName)

  if (!metadata) {
    throw new Error(`Sheet "${sheetName}" has no metadata.`)
  }

  const results: SearchResult[] = []

  outer: for (let rowIndex = 0; rowIndex < metadata.rowCount; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < metadata.columnCount; columnIndex += 1) {
      const value = formatCell(getCell(worksheet, rowIndex, columnIndex))

      if (!value) {
        continue
      }

      if (value.toLocaleLowerCase().includes(normalizedQuery)) {
        results.push({
          rowIndex,
          columnIndex,
          address: `${toColumnLabel(columnIndex)}${rowIndex + 1}`,
          value,
        })
      }

      if (results.length >= limit) {
        break outer
      }
    }
  }

  post({
    type: 'search-results',
    sheetName,
    query,
    results,
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
    }
  }

  const range = XLSX.utils.decode_range(ref)
  const rowCount = range.e.r + 1
  const columnCount = range.e.c + 1
  const columnWidths = Array.from({ length: columnCount }, () => DEFAULT_COLUMN_WIDTH)

  const sampleRowEnd = Math.min(range.e.r, WIDTH_SAMPLE_ROWS - 1)
  const sampleColumnEnd = Math.min(range.e.c, WIDTH_SAMPLE_COLUMNS - 1)

  for (let columnIndex = 0; columnIndex <= sampleColumnEnd; columnIndex += 1) {
    columnWidths[columnIndex] = clamp(
      toColumnLabel(columnIndex).length * 18 + 34,
      MIN_COLUMN_WIDTH,
      DEFAULT_COLUMN_WIDTH,
    )
  }

  for (let rowIndex = 0; rowIndex <= sampleRowEnd; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex <= sampleColumnEnd; columnIndex += 1) {
      const value = formatCell(getCell(worksheet, rowIndex, columnIndex))

      if (!value) {
        continue
      }

      columnWidths[columnIndex] = Math.max(
        columnWidths[columnIndex],
        clamp(value.slice(0, 48).length * 8.4 + 34, MIN_COLUMN_WIDTH, MAX_COLUMN_WIDTH),
      )
    }
  }

  return {
    name: sheetName,
    rowCount,
    columnCount,
    range: ref,
    columnWidths,
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
  const denseWorksheet = worksheet as unknown as Array<Array<XLSX.CellObject | undefined>>
  return denseWorksheet[rowIndex]?.[columnIndex]
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

function post(message: WorkerResponse): void {
  workerScope.postMessage(message)
}

export {}

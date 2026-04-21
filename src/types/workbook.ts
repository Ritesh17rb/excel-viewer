export interface WorkbookSummary {
  fileName: string
  fileSize: number
  format: string
  loadedAt: string
  sheetNames: string[]
}

export interface SheetMetadata {
  name: string
  rowCount: number
  columnCount: number
  range: string | null
  columnWidths: number[]
}

export interface GridWindow {
  key: string
  sheetName: string
  rowStart: number
  rowEnd: number
  colStart: number
  colEnd: number
  values: string[][]
}

export interface GridRange {
  rowStart: number
  rowEnd: number
  colStart: number
  colEnd: number
}

export interface CellPosition {
  rowIndex: number
  columnIndex: number
}

export interface SearchResult extends CellPosition {
  address: string
  value: string
}

export type WorkerRequest =
  | {
      type: 'load-workbook'
      buffer: ArrayBuffer
      fileName: string
      fileSize: number
      format: string
    }
  | {
      type: 'load-sheet'
      sheetName: string
    }
  | ({
      type: 'load-window'
      key: string
      sheetName: string
    } & GridRange)
  | {
      type: 'search-sheet'
      sheetName: string
      query: string
      limit: number
    }

export type WorkerResponse =
  | {
      type: 'workbook-loaded'
      workbook: WorkbookSummary
    }
  | {
      type: 'sheet-loaded'
      sheet: SheetMetadata
    }
  | {
      type: 'window-loaded'
      window: GridWindow
    }
  | {
      type: 'search-results'
      sheetName: string
      query: string
      results: SearchResult[]
    }
  | {
      type: 'error'
      message: string
    }

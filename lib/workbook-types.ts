export type BlobAccess = 'public' | 'private'
export type SortDirection = 'asc' | 'desc'

export interface WorkbookUploadSource {
  kind: 'blob' | 'local'
  fileName: string
  fileSize: number
  format: string
  url?: string
  pathname?: string
  access?: BlobAccess
  localPath?: string
}

export interface WorkbookSheetManifest {
  key: string
  name: string
  rowCount: number
  columnCount: number
  populatedCellCount: number
  pageSize: number
  pageCount: number
  columnWidths: number[]
  pages: string[]
  searchEnabled: boolean
  headerRows: number
}

export interface WorkbookManifest {
  id: string
  createdAt: string
  source: WorkbookUploadSource
  performanceMode: 'server'
  sheetNames: string[]
  sheets: Record<string, WorkbookSheetManifest>
}

export interface WorkbookSheetPage {
  sheetName: string
  page: number
  rowStart: number
  rowEnd: number
  rows: string[][]
}

export interface WorkbookFilterRule {
  columnIndex: number
  term: string
}

export interface WorkbookSortRule {
  columnIndex: number
  direction: SortDirection
}

export interface WorkbookSheetViewConfig {
  headerRows: number
  filters: WorkbookFilterRule[]
  sort: WorkbookSortRule | null
}

export interface WorkbookSheetViewManifest extends WorkbookSheetManifest {
  id: string
  kind: 'view'
  baseSheetName: string
  filters: WorkbookFilterRule[]
  sort: WorkbookSortRule | null
}

export interface WorkbookSearchMatch {
  rowIndex: number
  columnIndex: number
  address: string
  value: string
}

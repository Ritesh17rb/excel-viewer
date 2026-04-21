export type BlobAccess = 'public' | 'private'

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
  searchEnabled: false
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


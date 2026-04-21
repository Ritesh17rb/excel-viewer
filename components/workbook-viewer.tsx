'use client'

import { upload } from '@vercel/blob/client'
import { useMemo, useRef, useState } from 'react'

import { SpreadsheetGrid } from '@/components/spreadsheet-grid'
import {
  addressFromPosition,
  fileExtension,
  formatBytes,
  formatCount,
} from '@/lib/format'
import type {
  BlobAccess,
  WorkbookManifest,
  WorkbookSheetManifest,
  WorkbookSheetPage,
  WorkbookUploadSource,
} from '@/lib/workbook-types'

type UploadMode = 'blob' | 'local'
type ViewerStatus = 'idle' | 'uploading' | 'ingesting' | 'ready' | 'error'

interface WorkbookViewerProps {
  uploadMode: UploadMode
  blobAccess: BlobAccess
}

interface CellPosition {
  rowIndex: number
  columnIndex: number
}

const MAX_CACHED_PAGES = 80

export function WorkbookViewer({
  uploadMode,
  blobAccess,
}: WorkbookViewerProps) {
  const pageCacheRef = useRef(new Map<string, WorkbookSheetPage>())
  const pendingPageRequestsRef = useRef(new Set<string>())

  const [status, setStatus] = useState<ViewerStatus>('idle')
  const [uploadProgress, setUploadProgress] = useState(0)
  const [manifest, setManifest] = useState<WorkbookManifest | null>(null)
  const [activeSheetName, setActiveSheetName] = useState<string | null>(null)
  const [selectedCell, setSelectedCell] = useState<CellPosition | null>(null)
  const [cacheVersion, setCacheVersion] = useState(0)
  const [error, setError] = useState<string | null>(null)

  const activeSheet = useMemo<WorkbookSheetManifest | null>(() => {
    if (!manifest || !activeSheetName) {
      return null
    }

    return manifest.sheets[activeSheetName] ?? null
  }, [activeSheetName, manifest])

  const selectedAddress = selectedCell
    ? addressFromPosition(selectedCell.rowIndex, selectedCell.columnIndex)
    : 'None'

  const selectedValue =
    selectedCell && activeSheet
      ? getCellValue(activeSheet, selectedCell.rowIndex, selectedCell.columnIndex)
      : null

  async function handleFile(file: File) {
    resetViewer()

    if (fileExtension(file.name) !== 'xlsx') {
      setStatus('error')
      setError('The server-backed viewer currently supports .xlsx files only.')
      return
    }

    try {
      const source = await uploadWorkbook(file)

      setStatus('ingesting')

      const response = await fetch('/api/workbooks/ingest', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ source }),
      })

      const payload = (await response.json()) as
        | { manifest: WorkbookManifest }
        | { error: string }

      if (!response.ok || !('manifest' in payload)) {
        throw new Error(
          'error' in payload ? payload.error : 'Workbook ingestion failed.',
        )
      }

      setManifest(payload.manifest)
      setActiveSheetName(payload.manifest.sheetNames[0] ?? null)
      setStatus('ready')

      const firstSheetName = payload.manifest.sheetNames[0]
      if (firstSheetName) {
        requestVisibleRows(payload.manifest, firstSheetName, 0, 120)
      }
    } catch (caughtError) {
      setStatus('error')
      setError(
        caughtError instanceof Error
          ? caughtError.message
          : 'Workbook upload failed.',
      )
    }
  }

  async function uploadWorkbook(file: File): Promise<WorkbookUploadSource> {
    if (uploadMode === 'blob') {
      setStatus('uploading')

      const blob = await upload(
        `uploads/${crypto.randomUUID()}-${file.name.replace(/\s+/g, '-')}`,
        file,
        {
          access: blobAccess,
          handleUploadUrl: '/api/blob/upload',
          multipart: file.size > 25 * 1024 * 1024,
          onUploadProgress: ({ percentage }) => {
            setUploadProgress(percentage)
          },
        },
      )

      return {
        kind: 'blob',
        fileName: file.name,
        fileSize: file.size,
        format: fileExtension(file.name),
        url: blob.url,
        pathname: blob.pathname,
        access: blobAccess,
      }
    }

    setStatus('uploading')
    const formData = new FormData()
    formData.append('file', file)

    const response = await fetch('/api/uploads/local', {
      method: 'POST',
      body: formData,
    })
    const payload = (await response.json()) as
      | { source: WorkbookUploadSource }
      | { error: string }

    if (!response.ok || !('source' in payload)) {
      throw new Error(
        'error' in payload ? payload.error : 'Local upload failed.',
      )
    }

    return payload.source
  }

  function requestVisibleRows(
    nextManifest: WorkbookManifest,
    sheetName: string,
    rowStart: number,
    rowEnd: number,
  ) {
    const sheet = nextManifest.sheets[sheetName]

    if (!sheet || sheet.pageCount === 0) {
      return
    }

    const firstPage = Math.floor(rowStart / sheet.pageSize)
    const lastPage = Math.floor(Math.max(rowEnd - 1, rowStart) / sheet.pageSize)

    for (let page = firstPage; page <= lastPage; page += 1) {
      const cacheKey = `${sheetName}:${page}`

      if (
        pageCacheRef.current.has(cacheKey) ||
        pendingPageRequestsRef.current.has(cacheKey)
      ) {
        continue
      }

      pendingPageRequestsRef.current.add(cacheKey)

      void fetch(`/api/workbooks/${nextManifest.id}/sheet?name=${encodeURIComponent(sheetName)}&page=${page}`)
        .then(async (response) => {
          const payload = (await response.json()) as
            | { sheetPage: WorkbookSheetPage }
            | { error: string }

          if (!response.ok || !('sheetPage' in payload)) {
            throw new Error(
              'error' in payload ? payload.error : 'Sheet page request failed.',
            )
          }

          const cache = pageCacheRef.current
          cache.set(cacheKey, payload.sheetPage)

          while (cache.size > MAX_CACHED_PAGES) {
            const firstKey = cache.keys().next().value

            if (!firstKey) {
              break
            }

            cache.delete(firstKey)
          }

          setCacheVersion((current) => current + 1)
        })
        .catch((caughtError) => {
          setError(
            caughtError instanceof Error
              ? caughtError.message
              : 'A workbook page could not be loaded.',
          )
        })
        .finally(() => {
          pendingPageRequestsRef.current.delete(cacheKey)
        })
    }
  }

  function getCellValue(
    sheet: WorkbookSheetManifest,
    rowIndex: number,
    columnIndex: number,
  ): string | null {
    const page = Math.floor(rowIndex / sheet.pageSize)
    const cacheKey = `${sheet.name}:${page}`
    const sheetPage = pageCacheRef.current.get(cacheKey)

    if (!sheetPage) {
      return null
    }

    return sheetPage.rows[rowIndex - sheetPage.rowStart]?.[columnIndex] ?? ''
  }

  function resetViewer() {
    pageCacheRef.current.clear()
    pendingPageRequestsRef.current.clear()
    setCacheVersion(0)
    setUploadProgress(0)
    setManifest(null)
    setActiveSheetName(null)
    setSelectedCell(null)
    setError(null)
  }

  return (
    <main className="page-shell">
      <section className="hero-card">
        <div>
          <span className="eyebrow">Server-backed Vercel architecture</span>
          <h1>Large XLSX files parsed on the server, not in the browser.</h1>
          <p className="hero-copy">
            Uploads go straight to storage, the workbook is chunked server-side,
            and the grid only pulls the row pages you are actually viewing.
          </p>
        </div>

        <label className="upload-panel">
          <input
            className="visually-hidden"
            type="file"
            accept=".xlsx"
            onChange={(event) => {
              const [file] = event.currentTarget.files ?? []
              event.currentTarget.value = ''

              if (file) {
                void handleFile(file)
              }
            }}
          />

          <strong>Choose an .xlsx workbook</strong>
          <span>
            {uploadMode === 'blob'
              ? 'Direct browser upload to Vercel Blob'
              : 'Local development upload path'}
          </span>
          {status === 'uploading' ? (
            <small>Uploading {uploadProgress.toFixed(0)}%</small>
          ) : status === 'ingesting' ? (
            <small>Building paged workbook artifacts…</small>
          ) : (
            <small>Optimized for large row-heavy files like your 30 MB sample.</small>
          )}
        </label>
      </section>

      <section className="workspace-grid">
        <aside className="sidebar-card">
          <div className="metric-grid">
            <article>
              <span>Status</span>
              <strong>{status}</strong>
            </article>
            <article>
              <span>Mode</span>
              <strong>{uploadMode === 'blob' ? 'Blob + server' : 'Local + server'}</strong>
            </article>
            <article>
              <span>Workbook</span>
              <strong>{manifest ? manifest.source.fileName : 'None'}</strong>
            </article>
            <article>
              <span>Size</span>
              <strong>{manifest ? formatBytes(manifest.source.fileSize) : '0 B'}</strong>
            </article>
          </div>

          {manifest ? (
            <div className="sheet-list">
              {manifest.sheetNames.map((sheetName) => {
                const sheet = manifest.sheets[sheetName]

                return (
                  <button
                    key={sheetName}
                    className={`sheet-tab${sheetName === activeSheetName ? ' is-active' : ''}`}
                    type="button"
                    onClick={() => {
                      setActiveSheetName(sheetName)
                      setSelectedCell(null)
                      requestVisibleRows(manifest, sheetName, 0, 120)
                    }}
                  >
                    <strong>{sheetName}</strong>
                    <span>
                      {formatCount(sheet.rowCount)} rows • {formatCount(sheet.columnCount)} cols
                    </span>
                  </button>
                )
              })}
            </div>
          ) : (
            <p className="muted-copy">
              Load a workbook to inspect row pages served by the backend.
            </p>
          )}
        </aside>

        <section className="viewer-card">
          <div className="viewer-toolbar">
            <div>
              <span className="eyebrow eyebrow--small">Workbook</span>
              <h2>{activeSheet?.name ?? 'No sheet selected'}</h2>
            </div>

            <div className="stat-row">
              <div className="stat-pill">
                <span>Rows</span>
                <strong>{activeSheet ? formatCount(activeSheet.rowCount) : '0'}</strong>
              </div>
              <div className="stat-pill">
                <span>Cells</span>
                <strong>
                  {activeSheet ? formatCount(activeSheet.populatedCellCount) : '0'}
                </strong>
              </div>
              <div className="stat-pill">
                <span>Selection</span>
                <strong>{selectedAddress}</strong>
              </div>
            </div>
          </div>

          {activeSheet ? (
            <SpreadsheetGrid
              sheet={activeSheet}
              cacheVersion={cacheVersion}
              selectedCell={selectedCell}
              getCellValue={(rowIndex, columnIndex) =>
                getCellValue(activeSheet, rowIndex, columnIndex)
              }
              requestVisibleRows={(rowStart, rowEnd) => {
                if (manifest) {
                  requestVisibleRows(manifest, activeSheet.name, rowStart, rowEnd)
                }
              }}
              onSelectCell={setSelectedCell}
            />
          ) : (
            <div className="empty-state">
              <h3>Upload an XLSX file to start</h3>
              <p>
                This version uses a server-backed paging model so the browser never
                tries to hold the whole workbook in memory.
              </p>
            </div>
          )}

          <div className="inspector-card">
            <div className="inspector-header">
              <span className="eyebrow eyebrow--small">Cell inspector</span>
              <strong>{selectedAddress}</strong>
            </div>
            <p>{selectedCell ? selectedValue ?? 'Loading…' : 'Select a cell to inspect it.'}</p>
          </div>

          {error ? <p className="error-banner">{error}</p> : null}
        </section>
      </section>
    </main>
  )
}


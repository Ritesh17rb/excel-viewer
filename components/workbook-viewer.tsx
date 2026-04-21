'use client'

import { upload } from '@vercel/blob/client'
import { useMemo, useRef, useState } from 'react'

import { SpreadsheetGrid } from '@/components/spreadsheet-grid'
import {
  addressFromPosition,
  fileExtension,
  formatBytes,
  formatCount,
  toColumnLabel,
} from '@/lib/format'
import type {
  BlobAccess,
  SortDirection,
  WorkbookManifest,
  WorkbookSearchMatch,
  WorkbookSheetManifest,
  WorkbookSheetPage,
  WorkbookSheetViewManifest,
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

interface FilterDraft {
  id: string
  columnIndex: string
  term: string
}

const MAX_CACHED_PAGES = 80
const SEARCH_LIMIT = 50

export function WorkbookViewer({
  uploadMode,
  blobAccess,
}: WorkbookViewerProps) {
  const pendingPageRequestsRef = useRef(new Set<string>())
  const fileInputRef = useRef<HTMLInputElement>(null)
  const nextFilterIdRef = useRef(1)

  function createFilterDraft(): FilterDraft {
    const filter = {
      id: `filter-${nextFilterIdRef.current}`,
      columnIndex: '',
      term: '',
    }
    nextFilterIdRef.current += 1
    return filter
  }

  const [pageCache, setPageCache] = useState<Map<string, WorkbookSheetPage>>(
    () => new Map(),
  )
  const [status, setStatus] = useState<ViewerStatus>('idle')
  const [uploadProgress, setUploadProgress] = useState(0)
  const [manifest, setManifest] = useState<WorkbookManifest | null>(null)
  const [activeSheetName, setActiveSheetName] = useState<string | null>(null)
  const [activeView, setActiveView] = useState<WorkbookSheetViewManifest | null>(null)
  const [selectedCell, setSelectedCell] = useState<CellPosition | null>(null)
  const [cacheVersion, setCacheVersion] = useState(0)
  const [error, setError] = useState<string | null>(null)
  const [isDragging, setIsDragging] = useState(false)
  const [isApplyingView, setIsApplyingView] = useState(false)
  const [useHeaderRow, setUseHeaderRow] = useState(true)
  const [sortColumnIndex, setSortColumnIndex] = useState('')
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc')
  const [filters, setFilters] = useState<FilterDraft[]>([
    {
      id: 'filter-0',
      columnIndex: '',
      term: '',
    },
  ])
  const [searchQuery, setSearchQuery] = useState('')
  const [isSearching, setIsSearching] = useState(false)
  const [searchResults, setSearchResults] = useState<WorkbookSearchMatch[]>([])

  const activeSheet = useMemo<WorkbookSheetManifest | null>(() => {
    if (!manifest || !activeSheetName) {
      return null
    }

    return manifest.sheets[activeSheetName] ?? null
  }, [activeSheetName, manifest])

  const displaySheet = activeView ?? activeSheet
  const activeHeaderRows = activeView ? activeView.headerRows : useHeaderRow ? 1 : 0

  function getCellValue(
    source: WorkbookSheetManifest | WorkbookSheetViewManifest,
    rowIndex: number,
    columnIndex: number,
  ): string | null {
    const page = Math.floor(rowIndex / source.pageSize)
    const cacheKey = `${getSourceKey(source)}:${page}`
    const sheetPage = pageCache.get(cacheKey)

    if (!sheetPage) {
      return null
    }

    return sheetPage.rows[rowIndex - sheetPage.rowStart]?.[columnIndex] ?? ''
  }

  function requestVisibleRows(
    source: WorkbookSheetManifest | WorkbookSheetViewManifest,
    rowStart: number,
    rowEnd: number,
  ) {
    if (!manifest || source.pageCount === 0) {
      return
    }

    const firstPage = Math.floor(rowStart / source.pageSize)
    const lastPage = Math.floor(Math.max(rowEnd - 1, rowStart) / source.pageSize)
    const sourceKey = getSourceKey(source)

    for (let page = firstPage; page <= lastPage; page += 1) {
      const cacheKey = `${sourceKey}:${page}`

      if (
        pageCache.has(cacheKey) ||
        pendingPageRequestsRef.current.has(cacheKey)
      ) {
        continue
      }

      pendingPageRequestsRef.current.add(cacheKey)

      const params = new URLSearchParams({
        name: getSourceSheetName(source),
        page: String(page),
      })

      if (isViewManifest(source)) {
        params.set('view', source.id)
      }

      void fetch(`/api/workbooks/${manifest.id}/sheet?${params.toString()}`)
        .then(async (response) => {
          const payload = (await response.json()) as
            | { sheetPage: WorkbookSheetPage }
            | { error: string }

          if (!response.ok || !('sheetPage' in payload)) {
            throw new Error(
              'error' in payload ? payload.error : 'Sheet page request failed.',
            )
          }

          setPageCache((current) => {
            const next = new Map(current)
            next.set(cacheKey, payload.sheetPage)

            while (next.size > MAX_CACHED_PAGES) {
              const firstKey = next.keys().next().value

              if (!firstKey) {
                break
              }

              next.delete(firstKey)
            }

            return next
          })
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

  const selectedAddress = selectedCell
    ? addressFromPosition(selectedCell.rowIndex, selectedCell.columnIndex)
    : 'None'

  const selectedValue =
    selectedCell && displaySheet
      ? getCellValue(displaySheet, selectedCell.rowIndex, selectedCell.columnIndex)
      : null

  const columnOptions = displaySheet
    ? Array.from({ length: displaySheet.columnCount }, (_, index) => {
        const defaultLabel = toColumnLabel(index)

        if (activeHeaderRows === 0) {
          return {
            index,
            label: defaultLabel,
          }
        }

        const headerValue = getCellValue(displaySheet, 0, index)?.trim()
        return {
          index,
          label: headerValue ? `${defaultLabel} · ${headerValue}` : defaultLabel,
        }
      })
    : []

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

      const firstSheetName = payload.manifest.sheetNames[0] ?? null
      setManifest(payload.manifest)
      setActiveSheetName(firstSheetName)
      setStatus('ready')

      if (firstSheetName) {
        requestVisibleRows(payload.manifest.sheets[firstSheetName], 0, 120)
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

  async function applyView() {
    if (!manifest || !activeSheet) {
      return
    }

    const normalizedFilters = filters
      .filter((filter) => filter.columnIndex !== '' && filter.term.trim())
      .map((filter) => ({
        columnIndex: Number(filter.columnIndex),
        term: filter.term.trim(),
      }))

    const sort =
      sortColumnIndex !== ''
        ? {
            columnIndex: Number(sortColumnIndex),
            direction: sortDirection,
          }
        : null

    if (normalizedFilters.length === 0 && !sort) {
      setActiveView(null)
      setSearchResults([])
      requestVisibleRows(activeSheet, 0, 120)
      return
    }

    try {
      setIsApplyingView(true)
      setError(null)
      setSearchResults([])
      setSelectedCell(null)

      const response = await fetch(`/api/workbooks/${manifest.id}/view`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          sheetName: activeSheet.name,
          config: {
            headerRows: useHeaderRow ? 1 : 0,
            filters: normalizedFilters,
            sort,
          },
        }),
      })

      const payload = (await response.json()) as
        | { view: WorkbookSheetViewManifest }
        | { error: string }

      if (!response.ok || !('view' in payload)) {
        throw new Error(
          'error' in payload ? payload.error : 'View generation failed.',
        )
      }

      setActiveView(payload.view)
      requestVisibleRows(payload.view, 0, 120)
    } catch (caughtError) {
      setError(
        caughtError instanceof Error
          ? caughtError.message
          : 'Filter or sort could not be applied.',
      )
    } finally {
      setIsApplyingView(false)
    }
  }

  function clearDerivedView() {
    setActiveView(null)
    setSortColumnIndex('')
    setSortDirection('asc')
    setFilters([createFilterDraft()])
    setSearchResults([])
    setSelectedCell(null)

    if (activeSheet) {
      requestVisibleRows(activeSheet, 0, 120)
    }
  }

  async function runSearch() {
    if (!manifest || !activeSheetName) {
      return
    }

    const query = searchQuery.trim()

    if (!query) {
      setSearchResults([])
      return
    }

    try {
      setIsSearching(true)
      setError(null)

      const response = await fetch(`/api/workbooks/${manifest.id}/search`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          sheetName: activeSheetName,
          viewId: activeView?.id,
          query,
          limit: SEARCH_LIMIT,
        }),
      })

      const payload = (await response.json()) as
        | { matches: WorkbookSearchMatch[] }
        | { error: string }

      if (!response.ok || !('matches' in payload)) {
        throw new Error('error' in payload ? payload.error : 'Search failed.')
      }

      setSearchResults(payload.matches)
    } catch (caughtError) {
      setError(
        caughtError instanceof Error ? caughtError.message : 'Search failed.',
      )
    } finally {
      setIsSearching(false)
    }
  }

  function openSearchResult(match: WorkbookSearchMatch) {
    if (!displaySheet) {
      return
    }

    requestVisibleRows(
      displaySheet,
      Math.max(0, match.rowIndex - 60),
      Math.min(displaySheet.rowCount, match.rowIndex + 60),
    )
    setSelectedCell({
      rowIndex: match.rowIndex,
      columnIndex: match.columnIndex,
    })
  }

  function resetViewer() {
    pendingPageRequestsRef.current.clear()
    nextFilterIdRef.current = 1
    setPageCache(new Map())
    setCacheVersion(0)
    setUploadProgress(0)
    setManifest(null)
    setActiveSheetName(null)
    setActiveView(null)
    setSelectedCell(null)
    setError(null)
    setSearchQuery('')
    setSearchResults([])
    setUseHeaderRow(true)
    setSortColumnIndex('')
    setSortDirection('asc')
    setFilters([createFilterDraft()])
  }

  function handleSheetChange(sheetName: string) {
    if (!manifest) {
      return
    }

    setActiveSheetName(sheetName)
    setActiveView(null)
    setSelectedCell(null)
    setSearchQuery('')
    setSearchResults([])
    setUseHeaderRow(true)
    setSortColumnIndex('')
    setSortDirection('asc')
    setFilters([createFilterDraft()])
    requestVisibleRows(manifest.sheets[sheetName], 0, 120)
  }

  function onFileDrop(event: React.DragEvent<HTMLDivElement>) {
    event.preventDefault()
    setIsDragging(false)

    const file = event.dataTransfer.files?.[0]

    if (file) {
      void handleFile(file)
    }
  }

  return (
    <main className="page-shell">
      <section className="hero-card">
        <div>
          <span className="eyebrow">Server-backed Vercel architecture</span>
          <h1>Large XLSX files parsed on the server, not in the browser.</h1>
          <p className="hero-copy">
            Uploads go straight to storage, workbook pages stay on the backend,
            and the browser only requests the visible viewport, search hits,
            and derived sorted or filtered views.
          </p>
        </div>

        <div
          className={`upload-panel${isDragging ? ' is-dragging' : ''}`}
          onDragEnter={(event) => {
            event.preventDefault()
            setIsDragging(true)
          }}
          onDragOver={(event) => {
            event.preventDefault()
            setIsDragging(true)
          }}
          onDragLeave={(event) => {
            event.preventDefault()
            if (event.currentTarget === event.target) {
              setIsDragging(false)
            }
          }}
          onDrop={onFileDrop}
        >
          <input
            ref={fileInputRef}
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

          <strong>Upload an Excel workbook</strong>
          <span>
            Drag and drop a `.xlsx` file here or choose one explicitly.
          </span>

          <div className="upload-panel__actions">
            <button
              className="upload-button"
              type="button"
              onClick={() => fileInputRef.current?.click()}
            >
              Choose `.xlsx` file
            </button>
            <span className="upload-panel__meta">
              {uploadMode === 'blob'
                ? 'Direct browser upload to Vercel Blob'
                : 'Local development upload path'}
            </span>
          </div>

          {status === 'uploading' ? (
            <small>Uploading {uploadProgress.toFixed(0)}%</small>
          ) : status === 'ingesting' ? (
            <small>Building row pages, search, and view artifacts…</small>
          ) : (
            <small>
              Visible upload, large-file paging, search, sort, and filter controls
              are available below.
            </small>
          )}
        </div>
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
                    onClick={() => handleSheetChange(sheetName)}
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
              Upload a workbook to unlock server-backed search, sort, and filter
              views.
            </p>
          )}
        </aside>

        <section className="viewer-card">
          <div className="viewer-toolbar">
            <div>
              <span className="eyebrow eyebrow--small">Workbook</span>
              <h2>{displaySheet?.name ?? 'No sheet selected'}</h2>
            </div>

            <div className="stat-row">
              <div className="stat-pill">
                <span>Rows</span>
                <strong>{displaySheet ? formatCount(displaySheet.rowCount) : '0'}</strong>
              </div>
              <div className="stat-pill">
                <span>Cells</span>
                <strong>
                  {displaySheet ? formatCount(displaySheet.populatedCellCount) : '0'}
                </strong>
              </div>
              <div className="stat-pill">
                <span>Selection</span>
                <strong>{selectedAddress}</strong>
              </div>
            </div>
          </div>

          {displaySheet ? (
            <>
              <section className="workbench-toolbar">
                <div className="control-card">
                  <div className="control-card__header">
                    <span className="eyebrow eyebrow--small">Search</span>
                    <strong>{searchResults.length} hits</strong>
                  </div>
                  <div className="toolbar-row">
                    <input
                      className="toolbar-input"
                      type="search"
                      value={searchQuery}
                      placeholder="Search the current sheet or current filtered view"
                      onChange={(event) => setSearchQuery(event.currentTarget.value)}
                      onKeyDown={(event) => {
                        if (event.key === 'Enter') {
                          event.preventDefault()
                          void runSearch()
                        }
                      }}
                    />
                    <button
                      className="toolbar-button"
                      type="button"
                      onClick={() => void runSearch()}
                      disabled={isSearching}
                    >
                      {isSearching ? 'Searching…' : 'Search'}
                    </button>
                  </div>

                  {searchResults.length > 0 ? (
                    <div className="search-results">
                      {searchResults.map((match) => (
                        <button
                          key={`${match.address}:${match.value}`}
                          className="search-result"
                          type="button"
                          onClick={() => openSearchResult(match)}
                        >
                          <strong>{match.address}</strong>
                          <span>{match.value}</span>
                        </button>
                      ))}
                    </div>
                  ) : (
                    <p className="muted-copy muted-copy--compact">
                      Search scans the active sheet on the server and jumps the grid
                      to matching cells.
                    </p>
                  )}
                </div>

                <div className="control-card">
                  <div className="control-card__header">
                    <span className="eyebrow eyebrow--small">Sort and Filter</span>
                    <strong>
                      {activeView
                        ? `${formatCount(activeView.rowCount)} visible rows`
                        : 'Raw sheet'}
                    </strong>
                  </div>

                  <label className="toggle-row">
                    <input
                      type="checkbox"
                      checked={useHeaderRow}
                      onChange={(event) => setUseHeaderRow(event.currentTarget.checked)}
                    />
                    <span>Treat row 1 as the header row for sort and filter operations</span>
                  </label>

                  <div className="toolbar-grid">
                    <label className="toolbar-field">
                      <span>Sort column</span>
                      <select
                        className="toolbar-select"
                        value={sortColumnIndex}
                        onChange={(event) => setSortColumnIndex(event.currentTarget.value)}
                      >
                        <option value="">No sort</option>
                        {columnOptions.map((column) => (
                          <option key={column.index} value={column.index}>
                            {column.label}
                          </option>
                        ))}
                      </select>
                    </label>

                    <label className="toolbar-field">
                      <span>Direction</span>
                      <select
                        className="toolbar-select"
                        value={sortDirection}
                        onChange={(event) =>
                          setSortDirection(event.currentTarget.value as SortDirection)
                        }
                      >
                        <option value="asc">Ascending</option>
                        <option value="desc">Descending</option>
                      </select>
                    </label>
                  </div>

                  <div className="filter-stack">
                    {filters.map((filter, index) => (
                      <div key={filter.id} className="filter-row">
                        <label className="toolbar-field">
                          <span>Filter {index + 1}</span>
                          <select
                            className="toolbar-select"
                            value={filter.columnIndex}
                            onChange={(event) => {
                              const nextValue = event.currentTarget.value
                              setFilters((current) =>
                                current.map((entry) =>
                                  entry.id === filter.id
                                    ? { ...entry, columnIndex: nextValue }
                                    : entry,
                                ),
                              )
                            }}
                          >
                            <option value="">Choose column</option>
                            {columnOptions.map((column) => (
                              <option key={column.index} value={column.index}>
                                {column.label}
                              </option>
                            ))}
                          </select>
                        </label>

                        <label className="toolbar-field toolbar-field--grow">
                          <span>Contains</span>
                          <input
                            className="toolbar-input"
                            type="text"
                            value={filter.term}
                            placeholder="Type text to keep matching rows"
                            onChange={(event) => {
                              const nextValue = event.currentTarget.value
                              setFilters((current) =>
                                current.map((entry) =>
                                  entry.id === filter.id
                                    ? { ...entry, term: nextValue }
                                    : entry,
                                ),
                              )
                            }}
                          />
                        </label>

                        <button
                          className="toolbar-button toolbar-button--ghost"
                          type="button"
                          disabled={filters.length === 1}
                          onClick={() =>
                            setFilters((current) =>
                              current.length === 1
                                ? current
                                : current.filter((entry) => entry.id !== filter.id),
                            )
                          }
                        >
                          Remove
                        </button>
                      </div>
                    ))}
                  </div>

                  <div className="toolbar-row">
                    <button
                      className="toolbar-button toolbar-button--ghost"
                      type="button"
                      onClick={() =>
                        setFilters((current) => [...current, createFilterDraft()])
                      }
                    >
                      Add filter
                    </button>
                    <button
                      className="toolbar-button"
                      type="button"
                      onClick={() => void applyView()}
                      disabled={isApplyingView}
                    >
                      {isApplyingView ? 'Applying…' : 'Apply sort/filter'}
                    </button>
                    <button
                      className="toolbar-button toolbar-button--ghost"
                      type="button"
                      onClick={clearDerivedView}
                    >
                      Clear view
                    </button>
                  </div>

                  {activeView ? (
                    <p className="view-banner">
                      Showing a derived server view with
                      {activeView.sort
                        ? ` ${activeView.sort.direction} sort on ${
                            columnOptions.find(
                              (column) => column.index === activeView.sort?.columnIndex,
                            )?.label ?? toColumnLabel(activeView.sort.columnIndex)
                          }`
                        : ' no active sort'}
                      {activeView.filters.length > 0
                        ? ` and ${activeView.filters.length} active filter${activeView.filters.length === 1 ? '' : 's'}.`
                        : '.'}
                    </p>
                  ) : null}
                </div>
              </section>

              <SpreadsheetGrid
                sheet={displaySheet}
                cacheVersion={cacheVersion}
                headerRowCount={activeHeaderRows}
                selectedCell={selectedCell}
                getCellValue={(rowIndex, columnIndex) =>
                  getCellValue(displaySheet, rowIndex, columnIndex)
                }
                requestVisibleRows={(rowStart, rowEnd) => {
                  requestVisibleRows(displaySheet, rowStart, rowEnd)
                }}
                onSelectCell={setSelectedCell}
              />
            </>
          ) : (
            <div className="empty-state">
              <h3>Upload an XLSX file to start</h3>
              <p>
                This version exposes visible upload controls and keeps search,
                filtering, and sorting on the server so the browser does not freeze
                on larger workbooks.
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

function getSourceKey(source: WorkbookSheetManifest | WorkbookSheetViewManifest): string {
  return isViewManifest(source) ? `view:${source.id}` : `sheet:${source.name}`
}

function getSourceSheetName(
  source: WorkbookSheetManifest | WorkbookSheetViewManifest,
): string {
  return isViewManifest(source) ? source.baseSheetName : source.name
}

function isViewManifest(
  source: WorkbookSheetManifest | WorkbookSheetViewManifest,
): source is WorkbookSheetViewManifest {
  return 'kind' in source && source.kind === 'view'
}

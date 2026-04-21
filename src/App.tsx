import { useRef, useState } from 'react'

import './App.css'

import { SpreadsheetGrid } from './components/SpreadsheetGrid'
import { useSpreadsheetWorker } from './hooks/useSpreadsheetWorker'
import { addressFromPosition, formatBytes, formatCount } from './lib/format'
import type { CellPosition } from './types/workbook'

const ACCEPTED_FORMATS = '.xlsx,.xls,.xlsm,.xlsb,.csv,.tsv,.ods'

interface FocusTarget extends CellPosition {
  token: number
}

function describeStatus(status: string): string {
  switch (status) {
    case 'loading-workbook':
      return 'Reading workbook container'
    case 'loading-sheet':
      return 'Preparing sheet data'
    case 'ready':
      return 'Ready'
    case 'error':
      return 'Problem opening workbook'
    default:
      return 'Drop a spreadsheet to begin'
  }
}

function App() {
  const fileInputRef = useRef<HTMLInputElement>(null)
  const dragDepthRef = useRef(0)
  const focusTokenRef = useRef(0)

  const [isDragActive, setIsDragActive] = useState(false)
  const [searchInput, setSearchInput] = useState('')
  const [selectedCell, setSelectedCell] = useState<CellPosition | null>(null)
  const [focusCell, setFocusCell] = useState<FocusTarget | null>(null)

  const {
    status,
    workbook,
    activeSheetName,
    activeSheetMeta,
    sheetMetaByName,
    searchState,
    cacheVersion,
    error,
    openFile,
    selectSheet,
    requestVisibleTiles,
    getCellValue,
    runSearch,
  } = useSpreadsheetWorker()

  const selectedAddress = selectedCell
    ? addressFromPosition(selectedCell.rowIndex, selectedCell.columnIndex)
    : 'None'
  const selectedValue = selectedCell
    ? getCellValue(selectedCell.rowIndex, selectedCell.columnIndex)
    : null
  const searchDisabledReason = activeSheetMeta?.searchDisabledReason ?? null
  const isLargeMode =
    workbook?.performanceMode === 'large' || activeSheetMeta?.largeSheetMode === true

  async function handleFiles(fileList: FileList | null) {
    const [file] = fileList ?? []

    if (!file) {
      return
    }

    setSelectedCell(null)
    setFocusCell(null)
    setSearchInput('')

    await openFile(file)
  }

  function openPicker() {
    fileInputRef.current?.click()
  }

  function handleDragEnter(event: React.DragEvent<HTMLElement>) {
    event.preventDefault()
    dragDepthRef.current += 1
    setIsDragActive(true)
  }

  function handleDragLeave(event: React.DragEvent<HTMLElement>) {
    event.preventDefault()
    dragDepthRef.current = Math.max(0, dragDepthRef.current - 1)

    if (dragDepthRef.current === 0) {
      setIsDragActive(false)
    }
  }

  function handleDrop(event: React.DragEvent<HTMLElement>) {
    event.preventDefault()
    dragDepthRef.current = 0
    setIsDragActive(false)
    void handleFiles(event.dataTransfer.files)
  }

  function handleSearchSubmit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault()
    runSearch(searchInput)
  }

  function jumpToCell(rowIndex: number, columnIndex: number) {
    focusTokenRef.current += 1

    setFocusCell({
      rowIndex,
      columnIndex,
      token: focusTokenRef.current,
    })
    setSelectedCell({
      rowIndex,
      columnIndex,
    })
  }

  return (
    <div className="app-shell">
      <input
        ref={fileInputRef}
        className="visually-hidden"
        type="file"
        accept={ACCEPTED_FORMATS}
        onChange={(event) => {
          void handleFiles(event.currentTarget.files)
          event.currentTarget.value = ''
        }}
      />

      <header className="hero-panel">
        <div className="hero-panel__copy">
          <span className="eyebrow">GitHub Pages-ready spreadsheet viewer</span>
          <h1>Open big Excel workbooks directly in the browser.</h1>
          <p>
            Upload <code>.xlsx</code>, <code>.xls</code>, <code>.xlsb</code>,{' '}
            <code>.csv</code>, or <code>.ods</code> files. Parsing stays in a Web
            Worker, the grid is virtualized, and the file never leaves the device.
          </p>
        </div>

        <div className="hero-panel__actions">
          <div className="feature-pills" aria-label="Key capabilities">
            <span>Local-only</span>
            <span>Worker parsed</span>
            <span>Virtualized grid</span>
            <span>50 MB+ friendly</span>
          </div>

          <div className="hero-panel__cta-row">
            <button className="primary-button" type="button" onClick={openPicker}>
              Choose workbook
            </button>
            <p className="caption">
              Optimized for large local files. Actual limits still depend on
              browser memory and workbook complexity.
            </p>
          </div>
        </div>
      </header>

      <main className="workspace">
        <aside className="sidebar">
          <section
            className={`panel upload-panel${isDragActive ? ' is-active' : ''}`}
            onDragEnter={handleDragEnter}
            onDragLeave={handleDragLeave}
            onDragOver={(event) => event.preventDefault()}
            onDrop={handleDrop}
          >
            <div className="panel__header">
              <span className="panel__label">Drop Zone</span>
              <span className={`status-dot status-dot--${status}`}></span>
            </div>

            <h2>Load a workbook</h2>
            <p>
              Drag a spreadsheet here or browse the local filesystem. The viewer
              is fully static, so it can deploy on GitHub Pages without any
              backend.
            </p>

            <button className="secondary-button" type="button" onClick={openPicker}>
              Browse files
            </button>

            <p className="panel__support">
              Supports {ACCEPTED_FORMATS.replaceAll(',', ', ')}
            </p>
          </section>

          <section className="panel">
            <div className="panel__header">
              <span className="panel__label">Workbook</span>
              <span className="status-text" aria-live="polite">
                {describeStatus(status)}
              </span>
            </div>

            {workbook ? (
              <div className="stack">
                <div className="workbook-card">
                  <strong>{workbook.fileName}</strong>
                  <span>{workbook.format}</span>
                </div>

                <div className="metric-grid">
                  <article>
                    <span>File size</span>
                    <strong>{formatBytes(workbook.fileSize)}</strong>
                  </article>
                  <article>
                    <span>Sheets</span>
                    <strong>{formatCount(workbook.sheetNames.length)}</strong>
                  </article>
                  <article>
                    <span>Loaded</span>
                    <strong>
                      {new Date(workbook.loadedAt).toLocaleTimeString([], {
                        hour: '2-digit',
                        minute: '2-digit',
                      })}
                    </strong>
                  </article>
                  <article>
                    <span>Parsing</span>
                    <strong>
                      {workbook.performanceMode === 'large' ? 'Worker safe mode' : 'Worker'}
                    </strong>
                  </article>
                </div>
              </div>
            ) : (
              <p className="placeholder-copy">
                No workbook loaded yet. This viewer is meant for local inspection,
                not upload-and-store workflows.
              </p>
            )}
          </section>

          <section className="panel">
            <div className="panel__header">
              <span className="panel__label">Sheets</span>
              <span className="status-text">
                {workbook ? `${workbook.sheetNames.length} available` : 'Waiting'}
              </span>
            </div>

            {workbook ? (
              <div className="sheet-list" role="list">
                {workbook.sheetNames.map((sheetName) => {
                  const metadata = sheetMetaByName[sheetName]
                  const isActive = sheetName === activeSheetName

                  return (
                    <button
                      key={sheetName}
                      className={`sheet-list__item${isActive ? ' is-active' : ''}`}
                      type="button"
                      onClick={() => {
                        setSelectedCell(null)
                        setFocusCell(null)
                        setSearchInput('')
                        selectSheet(sheetName)
                      }}
                    >
                      <span>{sheetName}</span>
                      <small>
                        {metadata
                          ? `${formatCount(metadata.rowCount)} rows • ${formatCount(
                              metadata.columnCount,
                            )} cols`
                          : 'Load sheet'}
                      </small>
                    </button>
                  )
                })}
              </div>
            ) : (
              <p className="placeholder-copy">
                Loaded sheets appear here so you can jump between tabs quickly.
              </p>
            )}
          </section>
        </aside>

        <section className="viewer-panel">
          <div className="viewer-toolbar">
            <div className="viewer-toolbar__heading">
              <span className="panel__label">Viewer</span>
              <h2>{activeSheetName ?? 'Spreadsheet canvas'}</h2>
            </div>

            <form className="search-form" onSubmit={handleSearchSubmit}>
              <input
                type="search"
                placeholder={
                  activeSheetMeta?.searchEnabled === false
                    ? 'Search disabled for large sheets'
                    : 'Search current sheet'
                }
                value={searchInput}
                disabled={!activeSheetMeta || activeSheetMeta.searchEnabled === false}
                onChange={(event) => setSearchInput(event.currentTarget.value)}
              />
              <button
                className="primary-button primary-button--compact"
                type="submit"
                disabled={
                  !activeSheetMeta ||
                  activeSheetMeta.searchEnabled === false ||
                  searchState.status === 'loading'
                }
              >
                {searchState.status === 'loading' ? 'Searching…' : 'Search'}
              </button>
            </form>
          </div>

          <div className="viewer-statusbar">
            <div className="stat-chip">
              <span>Rows</span>
              <strong>{activeSheetMeta ? formatCount(activeSheetMeta.rowCount) : '0'}</strong>
            </div>
            <div className="stat-chip">
              <span>Columns</span>
              <strong>
                {activeSheetMeta ? formatCount(activeSheetMeta.columnCount) : '0'}
              </strong>
            </div>
            <div className="stat-chip">
              <span>Selection</span>
              <strong>{selectedAddress}</strong>
            </div>
            <div className="stat-chip">
              <span>Filled cells</span>
              <strong>
                {activeSheetMeta
                  ? formatCount(activeSheetMeta.populatedCellCount)
                  : '0'}
              </strong>
            </div>
            <div className="stat-chip">
              <span>Search</span>
              <strong>
                {activeSheetMeta?.searchEnabled === false
                  ? 'Guarded'
                  : searchState.status === 'done'
                  ? `${searchState.results.length} hits`
                  : searchState.status === 'loading'
                    ? 'Running'
                    : 'Idle'}
              </strong>
            </div>
          </div>

          {isLargeMode ? (
            <div className="info-banner">
              Large-file mode is active. The worker is using sparse cell lookups and
              tighter caching so big sheets do not allocate a full in-memory grid.
            </div>
          ) : null}

          {searchDisabledReason ? (
            <div className="info-banner info-banner--muted">{searchDisabledReason}</div>
          ) : null}

          {searchState.status === 'done' && searchState.query ? (
            <div className="search-results">
              {searchState.disabledReason ? (
                <p className="placeholder-copy">{searchState.disabledReason}</p>
              ) : searchState.results.length > 0 ? (
                searchState.results.map((result) => (
                  <button
                    key={`${result.address}:${result.value}`}
                    className="search-results__item"
                    type="button"
                    onClick={() => jumpToCell(result.rowIndex, result.columnIndex)}
                  >
                    <strong>{result.address}</strong>
                    <span>{result.value}</span>
                  </button>
                ))
              ) : (
                <p className="placeholder-copy">
                  No matches for <code>{searchState.query}</code> in the active
                  sheet.
                </p>
              )}
            </div>
          ) : null}

          <div className="viewer-stage">
            {activeSheetMeta ? (
              activeSheetMeta.rowCount > 0 && activeSheetMeta.columnCount > 0 ? (
                <SpreadsheetGrid
                  sheet={activeSheetMeta}
                  cacheVersion={cacheVersion}
                  selectedCell={selectedCell}
                  focusCell={focusCell}
                  getCellValue={getCellValue}
                  requestVisibleTiles={requestVisibleTiles}
                  onSelectCell={(cell) => setSelectedCell(cell)}
                />
              ) : (
                <div className="viewer-empty-state">
                  <div className="viewer-empty-state__art" />
                  <h3>The active sheet is empty</h3>
                  <p>This tab does not contain a populated cell range.</p>
                </div>
              )
            ) : (
              <div className="viewer-empty-state">
                <div className="viewer-empty-state__art" />
                <h3>Built for static deployment and large sheets</h3>
                <p>
                  Open a workbook to inspect it with worker-based parsing and a
                  virtualized grid that keeps scrolling fluid.
                </p>
              </div>
            )}
          </div>

          <div className="inspector">
            <div>
              <span className="panel__label">Cell inspector</span>
              <strong>{selectedAddress}</strong>
            </div>
            <p>
              {selectedCell
                ? selectedValue == null
                  ? 'Fetching the selected cell from the worker cache…'
                  : selectedValue || 'Blank cell'
                : 'Select a cell to inspect its value here.'}
            </p>
          </div>

          {error ? <p className="error-banner">{error}</p> : null}
        </section>
      </main>
    </div>
  )
}

export default App

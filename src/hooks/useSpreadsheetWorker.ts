import { startTransition, useEffect, useEffectEvent, useRef, useState } from 'react'

import { fileExtension } from '../lib/format'
import type {
  GridRange,
  GridWindow,
  SearchResult,
  SheetMetadata,
  WorkbookSummary,
  WorkerResponse,
} from '../types/workbook'

const TILE_ROWS = 60
const TILE_COLUMNS = 12
const INITIAL_ROW_PREFETCH = 96
const INITIAL_COLUMN_PREFETCH = 14
const LARGE_INITIAL_ROW_PREFETCH = 48
const LARGE_INITIAL_COLUMN_PREFETCH = 10
const STANDARD_MAX_CACHED_WINDOWS = 220
const LARGE_MAX_CACHED_WINDOWS = 96

type ViewerStatus =
  | 'idle'
  | 'loading-workbook'
  | 'loading-sheet'
  | 'ready'
  | 'error'

interface SearchState {
  status: 'idle' | 'loading' | 'done'
  query: string
  results: SearchResult[]
  disabledReason: string | null
}

function makeTileKey(
  sheetName: string,
  tileRowIndex: number,
  tileColumnIndex: number,
): string {
  return `${sheetName}:${tileRowIndex}:${tileColumnIndex}`
}

export function useSpreadsheetWorker() {
  const workerRef = useRef<Worker | null>(null)
  const workbookRef = useRef<WorkbookSummary | null>(null)
  const activeSheetRef = useRef<string | null>(null)
  const searchQueryRef = useRef('')
  const sheetMetadataRef = useRef<Record<string, SheetMetadata>>({})
  const windowCacheRef = useRef(new Map<string, GridWindow>())
  const pendingWindowRequestsRef = useRef(new Set<string>())

  const [status, setStatus] = useState<ViewerStatus>('idle')
  const [workbook, setWorkbook] = useState<WorkbookSummary | null>(null)
  const [activeSheetName, setActiveSheetName] = useState<string | null>(null)
  const [activeSheetMeta, setActiveSheetMeta] = useState<SheetMetadata | null>(null)
  const [sheetMetaByName, setSheetMetaByName] = useState<
    Record<string, SheetMetadata>
  >({})
  const [searchState, setSearchState] = useState<SearchState>({
    status: 'idle',
    query: '',
    results: [],
    disabledReason: null,
  })
  const [cacheVersion, setCacheVersion] = useState(0)
  const [error, setError] = useState<string | null>(null)

  const clearWindowCache = () => {
    windowCacheRef.current.clear()
    pendingWindowRequestsRef.current.clear()
    startTransition(() => {
      setCacheVersion((current) => current + 1)
    })
  }

  function requestVisibleTiles(range: GridRange, sheetName = activeSheetRef.current) {
    if (!sheetName) {
      return
    }

    const worker = workerRef.current
    const sheet = sheetMetadataRef.current[sheetName]

    if (!worker || !sheet || sheet.rowCount === 0 || sheet.columnCount === 0) {
      return
    }

    const rowStart = Math.max(0, Math.min(range.rowStart, sheet.rowCount - 1))
    const rowEnd = Math.max(rowStart + 1, Math.min(range.rowEnd, sheet.rowCount))
    const colStart = Math.max(0, Math.min(range.colStart, sheet.columnCount - 1))
    const colEnd = Math.max(colStart + 1, Math.min(range.colEnd, sheet.columnCount))

    const firstTileRow = Math.floor(rowStart / TILE_ROWS)
    const lastTileRow = Math.floor((rowEnd - 1) / TILE_ROWS)
    const firstTileColumn = Math.floor(colStart / TILE_COLUMNS)
    const lastTileColumn = Math.floor((colEnd - 1) / TILE_COLUMNS)

    for (let tileRow = firstTileRow; tileRow <= lastTileRow; tileRow += 1) {
      for (
        let tileColumn = firstTileColumn;
        tileColumn <= lastTileColumn;
        tileColumn += 1
      ) {
        const key = makeTileKey(sheetName, tileRow, tileColumn)

        if (
          windowCacheRef.current.has(key) ||
          pendingWindowRequestsRef.current.has(key)
        ) {
          continue
        }

        pendingWindowRequestsRef.current.add(key)

        worker.postMessage({
          type: 'load-window',
          key,
          sheetName,
          rowStart: tileRow * TILE_ROWS,
          rowEnd: Math.min((tileRow + 1) * TILE_ROWS, sheet.rowCount),
          colStart: tileColumn * TILE_COLUMNS,
          colEnd: Math.min((tileColumn + 1) * TILE_COLUMNS, sheet.columnCount),
        })
      }
    }
  }

  function selectSheet(sheetName: string) {
    if (!workerRef.current || sheetName === activeSheetRef.current) {
      return
    }

    activeSheetRef.current = sheetName
    searchQueryRef.current = ''
    clearWindowCache()

    startTransition(() => {
      setActiveSheetName(sheetName)
      setActiveSheetMeta(sheetMetadataRef.current[sheetName] ?? null)
      setSearchState({
        status: 'idle',
        query: '',
        results: [],
        disabledReason: null,
      })
      setStatus('loading-sheet')
      setError(null)
    })

    workerRef.current.postMessage({
      type: 'load-sheet',
      sheetName,
    })
  }

  async function openFile(file: File) {
    const worker = workerRef.current

    if (!worker) {
      return
    }

    try {
      const buffer = await file.arrayBuffer()

      workbookRef.current = null
      activeSheetRef.current = null
      searchQueryRef.current = ''
      sheetMetadataRef.current = {}
      clearWindowCache()

      startTransition(() => {
        setStatus('loading-workbook')
        setWorkbook(null)
        setActiveSheetName(null)
        setActiveSheetMeta(null)
        setSheetMetaByName({})
        setSearchState({
          status: 'idle',
          query: '',
          results: [],
          disabledReason: null,
        })
        setError(null)
      })

      worker.postMessage(
        {
          type: 'load-workbook',
          buffer,
          fileName: file.name,
          fileSize: file.size,
          format: fileExtension(file.name),
        },
        [buffer],
      )
    } catch (caughtError) {
      startTransition(() => {
        setStatus('error')
        setError(
          caughtError instanceof Error
            ? caughtError.message
            : 'The selected file could not be opened.',
        )
      })
    }
  }

  function runSearch(query: string) {
    const worker = workerRef.current
    const sheetName = activeSheetRef.current
    const normalizedQuery = query.trim()
    const activeSheet = sheetName ? sheetMetadataRef.current[sheetName] : null

    searchQueryRef.current = normalizedQuery

    if (!worker || !sheetName) {
      return
    }

    if (!normalizedQuery) {
      startTransition(() => {
        setSearchState({
          status: 'idle',
          query: '',
          results: [],
          disabledReason: null,
        })
      })
      return
    }

    if (activeSheet && !activeSheet.searchEnabled) {
      startTransition(() => {
        setSearchState({
          status: 'done',
          query: normalizedQuery,
          results: [],
          disabledReason: activeSheet.searchDisabledReason,
        })
      })
      return
    }

    startTransition(() => {
      setSearchState({
        status: 'loading',
        query: normalizedQuery,
        results: [],
        disabledReason: null,
      })
    })

    worker.postMessage({
      type: 'search-sheet',
      sheetName,
      query: normalizedQuery,
      limit: 60,
    })
  }

  function getCellValue(rowIndex: number, columnIndex: number): string | null {
    const sheetName = activeSheetRef.current

    if (!sheetName) {
      return null
    }

    const tileKey = makeTileKey(
      sheetName,
      Math.floor(rowIndex / TILE_ROWS),
      Math.floor(columnIndex / TILE_COLUMNS),
    )
    const window = windowCacheRef.current.get(tileKey)

    if (!window) {
      return null
    }

    return (
      window.values[rowIndex - window.rowStart]?.[columnIndex - window.colStart] ?? ''
    )
  }

  const handleWorkerMessage = useEffectEvent(
    (event: MessageEvent<WorkerResponse>) => {
      const message = event.data

      switch (message.type) {
        case 'workbook-loaded': {
          workbookRef.current = message.workbook
          sheetMetadataRef.current = {}
          setSheetMetaByName({})

          const [firstSheetName] = message.workbook.sheetNames
          activeSheetRef.current = firstSheetName ?? null

          startTransition(() => {
            setWorkbook(message.workbook)
            setActiveSheetName(firstSheetName ?? null)
            setActiveSheetMeta(null)
            setStatus(firstSheetName ? 'loading-sheet' : 'ready')
            setError(null)
          })

          if (firstSheetName) {
            workerRef.current?.postMessage({
              type: 'load-sheet',
              sheetName: firstSheetName,
            })
          }

          return
        }

        case 'sheet-loaded': {
          const nextSheetMetadata = {
            ...sheetMetadataRef.current,
            [message.sheet.name]: message.sheet,
          }

          sheetMetadataRef.current = nextSheetMetadata

          startTransition(() => {
            setSheetMetaByName(nextSheetMetadata)
          })

          if (message.sheet.name !== activeSheetRef.current) {
            return
          }

          startTransition(() => {
            setActiveSheetMeta(message.sheet)
            setStatus('ready')
            setError(null)
          })

          requestVisibleTiles(
            {
              rowStart: 0,
              rowEnd:
                message.sheet.largeSheetMode || workbookRef.current?.performanceMode === 'large'
                  ? LARGE_INITIAL_ROW_PREFETCH
                  : INITIAL_ROW_PREFETCH,
              colStart: 0,
              colEnd:
                message.sheet.largeSheetMode || workbookRef.current?.performanceMode === 'large'
                  ? LARGE_INITIAL_COLUMN_PREFETCH
                  : INITIAL_COLUMN_PREFETCH,
            },
            message.sheet.name,
          )

          return
        }

        case 'window-loaded': {
          pendingWindowRequestsRef.current.delete(message.window.key)

          if (message.window.sheetName !== activeSheetRef.current) {
            return
          }

          const cachedWindows = windowCacheRef.current
          cachedWindows.set(message.window.key, message.window)

          const cacheLimit =
            workbookRef.current?.performanceMode === 'large'
              ? LARGE_MAX_CACHED_WINDOWS
              : STANDARD_MAX_CACHED_WINDOWS

          while (cachedWindows.size > cacheLimit) {
            const oldestKey = cachedWindows.keys().next().value

            if (!oldestKey) {
              break
            }

            cachedWindows.delete(oldestKey)
          }

          startTransition(() => {
            setCacheVersion((current) => current + 1)
          })

          return
        }

        case 'search-results': {
          if (
            message.sheetName !== activeSheetRef.current ||
            message.query !== searchQueryRef.current
          ) {
            return
          }

          startTransition(() => {
            setSearchState({
              status: 'done',
              query: message.query,
              results: message.results,
              disabledReason: message.disabledReason,
            })
          })

          return
        }

        case 'error': {
          startTransition(() => {
            setStatus('error')
            setError(message.message)
          })
        }
      }
    },
  )

  useEffect(() => {
    const worker = new Worker(
      new URL('../workers/spreadsheet.worker.ts', import.meta.url),
      { type: 'module' },
    )

    const onMessage = (event: MessageEvent<WorkerResponse>) => {
      handleWorkerMessage(event)
    }

    workerRef.current = worker
    worker.addEventListener('message', onMessage)

    return () => {
      worker.removeEventListener('message', onMessage)
      worker.terminate()
      workerRef.current = null
    }
  }, [])

  return {
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
  }
}

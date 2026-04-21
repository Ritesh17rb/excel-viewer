'use client'

import { useEffect, useRef } from 'react'

import { useVirtualizer } from '@tanstack/react-virtual'

import { toColumnLabel } from '@/lib/format'
import type { WorkbookSheetManifest } from '@/lib/workbook-types'

interface CellPosition {
  rowIndex: number
  columnIndex: number
}

interface SpreadsheetGridProps {
  sheet: WorkbookSheetManifest
  cacheVersion: number
  headerRowCount: number
  selectedCell: CellPosition | null
  getCellValue: (rowIndex: number, columnIndex: number) => string | null
  requestVisibleRows: (rowStart: number, rowEnd: number) => void
  onSelectCell: (cell: CellPosition) => void
}

const ROW_HEADER_WIDTH = 84
const COLUMN_HEADER_HEIGHT = 48
const CELL_HEIGHT = 40

export function SpreadsheetGrid({
  sheet,
  cacheVersion,
  headerRowCount,
  selectedCell,
  getCellValue,
  requestVisibleRows,
  onSelectCell,
}: SpreadsheetGridProps) {
  const scrollRef = useRef<HTMLDivElement>(null)

  // eslint-disable-next-line react-hooks/incompatible-library
  const rowVirtualizer = useVirtualizer({
    count: sheet.rowCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => CELL_HEIGHT,
    overscan: 16,
  })

  const columnVirtualizer = useVirtualizer({
    horizontal: true,
    count: sheet.columnCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: (index) => sheet.columnWidths[index] ?? 148,
    overscan: 4,
  })

  const visibleRows = rowVirtualizer.getVirtualItems()
  const visibleColumns = columnVirtualizer.getVirtualItems()

  useEffect(() => {
    if (!visibleRows.length) {
      return
    }

    const first = visibleRows[0]?.index ?? 0
    const last = (visibleRows.at(-1)?.index ?? 0) + 1
    requestVisibleRows(Math.max(0, first - 40), Math.min(sheet.rowCount, last + 40))
  }, [requestVisibleRows, sheet.rowCount, visibleRows])

  useEffect(() => {
    if (!selectedCell) {
      return
    }

    rowVirtualizer.scrollToIndex(selectedCell.rowIndex, {
      align: 'auto',
    })
    columnVirtualizer.scrollToIndex(selectedCell.columnIndex, {
      align: 'auto',
    })
  }, [columnVirtualizer, rowVirtualizer, selectedCell])

  return (
    <div className="sheet-grid" data-cache-version={cacheVersion}>
      <div
        className="sheet-grid__corner"
        style={{ width: ROW_HEADER_WIDTH, height: COLUMN_HEADER_HEIGHT }}
      >
        #
      </div>

      <div
        className="sheet-grid__header"
        style={{ left: ROW_HEADER_WIDTH, height: COLUMN_HEADER_HEIGHT }}
      >
        <div
          className="sheet-grid__header-track"
          style={{
            width: columnVirtualizer.getTotalSize(),
            transform: `translateX(${-(columnVirtualizer.scrollOffset ?? 0)}px)`,
          }}
        >
          {visibleColumns.map((column) => (
            <div
              key={column.key}
              className={`sheet-grid__column-heading${
                selectedCell?.columnIndex === column.index ? ' is-selected' : ''
              }`}
              style={{
                width: column.size,
                height: COLUMN_HEADER_HEIGHT,
                transform: `translateX(${column.start}px)`,
              }}
            >
              {toColumnLabel(column.index)}
            </div>
          ))}
        </div>
      </div>

      <div
        className="sheet-grid__rows"
        style={{ top: COLUMN_HEADER_HEIGHT, width: ROW_HEADER_WIDTH }}
      >
        <div
          className="sheet-grid__rows-track"
          style={{
            height: rowVirtualizer.getTotalSize(),
            transform: `translateY(${-(rowVirtualizer.scrollOffset ?? 0)}px)`,
          }}
        >
          {visibleRows.map((row) => (
            <div
              key={row.key}
              className={`sheet-grid__row-heading${
                selectedCell?.rowIndex === row.index ? ' is-selected' : ''
              }`}
              style={{
                width: ROW_HEADER_WIDTH,
                height: row.size,
                transform: `translateY(${row.start}px)`,
              }}
            >
              {row.index + 1}
            </div>
          ))}
        </div>
      </div>

      <div
        ref={scrollRef}
        className="sheet-grid__viewport"
        style={{
          left: ROW_HEADER_WIDTH,
          top: COLUMN_HEADER_HEIGHT,
        }}
      >
        <div
          className="sheet-grid__canvas"
          style={{
            width: columnVirtualizer.getTotalSize(),
            height: rowVirtualizer.getTotalSize(),
          }}
        >
          {visibleRows.map((row) =>
            visibleColumns.map((column) => {
              const value = getCellValue(row.index, column.index)
              const isSelected =
                selectedCell?.rowIndex === row.index &&
                selectedCell?.columnIndex === column.index
              const isHeaderCell = row.index < headerRowCount

              return (
                <button
                  key={`${row.index}:${column.index}`}
                  className={`sheet-grid__cell${isSelected ? ' is-selected' : ''}${
                    value == null ? ' is-loading' : ''
                  }${isHeaderCell ? ' is-header' : ''}${
                    row.index % 2 === 1 ? ' is-alt' : ''
                  }`}
                  type="button"
                  onClick={() =>
                    onSelectCell({
                      rowIndex: row.index,
                      columnIndex: column.index,
                    })
                  }
                  style={{
                    width: column.size,
                    height: row.size,
                    transform: `translate(${column.start}px, ${row.start}px)`,
                  }}
                  title={value ?? 'Loading cell contents'}
                >
                  <span className="sheet-grid__cell-text">{value ?? ''}</span>
                </button>
              )
            }),
          )}
        </div>
      </div>
    </div>
  )
}

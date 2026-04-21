import { useEffect, useRef } from 'react'

import { useVirtualizer } from '@tanstack/react-virtual'

import { toColumnLabel } from '../lib/format'
import type { CellPosition, GridRange, SheetMetadata } from '../types/workbook'

interface FocusTarget extends CellPosition {
  token: number
}

interface SpreadsheetGridProps {
  sheet: SheetMetadata
  cacheVersion: number
  selectedCell: CellPosition | null
  focusCell: FocusTarget | null
  getCellValue: (rowIndex: number, columnIndex: number) => string | null
  requestVisibleTiles: (range: GridRange) => void
  onSelectCell: (cell: CellPosition) => void
}

const CELL_HEIGHT = 42

export function SpreadsheetGrid({
  sheet,
  cacheVersion,
  selectedCell,
  focusCell,
  getCellValue,
  requestVisibleTiles,
  onSelectCell,
}: SpreadsheetGridProps) {
  const scrollRef = useRef<HTMLDivElement>(null)

  // eslint-disable-next-line react-hooks/incompatible-library
  const rowVirtualizer = useVirtualizer({
    count: sheet.rowCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => CELL_HEIGHT,
    overscan: 12,
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
  const firstVisibleRow = visibleRows[0]?.index ?? 0
  const lastVisibleRow = (visibleRows.at(-1)?.index ?? 0) + 1
  const firstVisibleColumn = visibleColumns[0]?.index ?? 0
  const lastVisibleColumn = (visibleColumns.at(-1)?.index ?? 0) + 1
  const totalBodyWidth = columnVirtualizer.getTotalSize()
  const totalBodyHeight = rowVirtualizer.getTotalSize()

  useEffect(() => {
    if (!visibleRows.length || !visibleColumns.length) {
      return
    }

    requestVisibleTiles({
      rowStart: Math.max(0, firstVisibleRow - 18),
      rowEnd: Math.min(sheet.rowCount, lastVisibleRow + 18),
      colStart: Math.max(0, firstVisibleColumn - 4),
      colEnd: Math.min(sheet.columnCount, lastVisibleColumn + 4),
    })
  }, [
    firstVisibleColumn,
    firstVisibleRow,
    lastVisibleColumn,
    lastVisibleRow,
    requestVisibleTiles,
    sheet.columnCount,
    sheet.rowCount,
    visibleColumns.length,
    visibleRows.length,
  ])

  useEffect(() => {
    if (!focusCell) {
      return
    }

    rowVirtualizer.scrollToIndex(focusCell.rowIndex, {
      align: 'center',
    })
    columnVirtualizer.scrollToIndex(focusCell.columnIndex, {
      align: 'center',
    })

    requestVisibleTiles({
      rowStart: Math.max(0, focusCell.rowIndex - 24),
      rowEnd: Math.min(sheet.rowCount, focusCell.rowIndex + 24),
      colStart: Math.max(0, focusCell.columnIndex - 6),
      colEnd: Math.min(sheet.columnCount, focusCell.columnIndex + 6),
    })

    onSelectCell({
      rowIndex: focusCell.rowIndex,
      columnIndex: focusCell.columnIndex,
    })
  }, [
    columnVirtualizer,
    focusCell,
    onSelectCell,
    requestVisibleTiles,
    rowVirtualizer,
    sheet.columnCount,
    sheet.rowCount,
  ])

  return (
    <div className="sheet-grid" data-cache-version={cacheVersion}>
      <div className="sheet-grid__corner">#</div>

      <div className="sheet-grid__header">
        <div
          className="sheet-grid__header-track"
          style={{
            width: totalBodyWidth,
            transform: `translateX(${-(columnVirtualizer.scrollOffset ?? 0)}px)`,
          }}
        >
          {visibleColumns.map((column) => {
            const isSelected = selectedCell?.columnIndex === column.index

            return (
              <div
                key={column.key}
                className={`sheet-grid__column-heading${
                  isSelected ? ' is-selected' : ''
                }`}
                style={{
                  width: column.size,
                  transform: `translateX(${column.start}px)`,
                }}
              >
                {toColumnLabel(column.index)}
              </div>
            )
          })}
        </div>
      </div>

      <div className="sheet-grid__rows">
        <div
          className="sheet-grid__rows-track"
          style={{
            height: totalBodyHeight,
            transform: `translateY(${-(rowVirtualizer.scrollOffset ?? 0)}px)`,
          }}
        >
          {visibleRows.map((row) => {
            const isSelected = selectedCell?.rowIndex === row.index

            return (
              <div
                key={row.key}
                className={`sheet-grid__row-heading${
                  isSelected ? ' is-selected' : ''
                }`}
                style={{
                  height: row.size,
                  transform: `translateY(${row.start}px)`,
                }}
              >
                {row.index + 1}
              </div>
            )
          })}
        </div>
      </div>

      <div ref={scrollRef} className="sheet-grid__viewport">
        <div
          className="sheet-grid__canvas"
          style={{
            width: totalBodyWidth,
            height: totalBodyHeight,
          }}
        >
          {visibleRows.map((row) =>
            visibleColumns.map((column) => {
              const value = getCellValue(row.index, column.index)
              const isSelected =
                selectedCell?.rowIndex === row.index &&
                selectedCell?.columnIndex === column.index

              return (
                <button
                  key={`${row.index}:${column.index}`}
                  className={`sheet-grid__cell${
                    isSelected ? ' is-selected' : ''
                  }${value == null ? ' is-loading' : ''}`}
                  type="button"
                  title={value ?? 'Loading cell contents'}
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
                >
                  <span className="sheet-grid__cell-text">
                    {value == null ? '' : value}
                  </span>
                </button>
              )
            }),
          )}
        </div>
      </div>

      <div className="sheet-grid__meta">
        <span>{sheet.range ?? 'Empty range'}</span>
        <span>{sheet.rowCount} rows</span>
        <span>{sheet.columnCount} columns</span>
      </div>
    </div>
  )
}

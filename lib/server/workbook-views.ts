import { createHash } from 'node:crypto'

import { addressFromPosition } from '@/lib/format'
import { putWorkbookJson, readWorkbookJson, readWorkbookJsonIfExists } from '@/lib/server/workbook-store'
import type {
  WorkbookFilterRule,
  WorkbookSearchMatch,
  WorkbookSheetManifest,
  WorkbookSheetPage,
  WorkbookSheetViewConfig,
  WorkbookSheetViewManifest,
  WorkbookSortRule,
} from '@/lib/workbook-types'

const collator = new Intl.Collator('en', {
  numeric: true,
  sensitivity: 'base',
})

interface RowRecord {
  row: string[]
  originalRowIndex: number
}

export async function createWorkbookView(
  workbookId: string,
  sheet: WorkbookSheetManifest,
  config: WorkbookSheetViewConfig,
): Promise<WorkbookSheetViewManifest> {
  const normalizedConfig = normalizeViewConfig(config)
  const viewId = createHash('sha1')
    .update(JSON.stringify({ sheetName: sheet.name, config: normalizedConfig }))
    .digest('hex')
    .slice(0, 16)

  const manifestPath = `${workbookId}/views/${viewId}/manifest.json`
  const existing = await readWorkbookJsonIfExists<WorkbookSheetViewManifest>(manifestPath)

  if (existing) {
    return existing
  }

  const headerRows: string[][] = []
  const dataRows: RowRecord[] = []
  let populatedCellCount = 0

  for (const pagePath of sheet.pages) {
    const page = await readWorkbookJson<WorkbookSheetPage>(pagePath)

    page.rows.forEach((row, rowOffset) => {
      const rowIndex = page.rowStart + rowOffset
      const trimmedRow = trimTrailingEmpty(row)

      if (rowIndex < normalizedConfig.headerRows) {
        headerRows[rowIndex] = trimmedRow
        populatedCellCount += countPopulatedCells(trimmedRow)
        return
      }

      if (!matchesFilters(trimmedRow, normalizedConfig.filters)) {
        return
      }

      populatedCellCount += countPopulatedCells(trimmedRow)

      dataRows.push({
        row: trimmedRow,
        originalRowIndex: rowIndex,
      })
    })
  }

  if (normalizedConfig.sort) {
    const { columnIndex, direction } = normalizedConfig.sort

    dataRows.sort((left, right) => {
      const comparison = collator.compare(
        left.row[columnIndex] ?? '',
        right.row[columnIndex] ?? '',
      )

      if (comparison !== 0) {
        return direction === 'asc' ? comparison : -comparison
      }

      return left.originalRowIndex - right.originalRowIndex
    })
  }

  const pagePaths: string[] = []
  const rowBuffer: string[][] = []
  let pageStartRow = 0
  let outputRowIndex = 0

  for (const headerRow of headerRows) {
    if (rowBuffer.length === 0) {
      pageStartRow = outputRowIndex
    }

    rowBuffer.push(headerRow ?? [])
    outputRowIndex += 1

    if (rowBuffer.length === sheet.pageSize) {
      pagePaths.push(
        await flushViewPage({
          workbookId,
          viewId,
          sheetName: sheet.name,
          page: pagePaths.length,
          rowStart: pageStartRow,
          rowEnd: pageStartRow + rowBuffer.length,
          rows: rowBuffer.splice(0, rowBuffer.length),
        }),
      )
    }
  }

  for (const record of dataRows) {
    if (rowBuffer.length === 0) {
      pageStartRow = outputRowIndex
    }

    rowBuffer.push(record.row)
    outputRowIndex += 1

    if (rowBuffer.length === sheet.pageSize) {
      pagePaths.push(
        await flushViewPage({
          workbookId,
          viewId,
          sheetName: sheet.name,
          page: pagePaths.length,
          rowStart: pageStartRow,
          rowEnd: pageStartRow + rowBuffer.length,
          rows: rowBuffer.splice(0, rowBuffer.length),
        }),
      )
    }
  }

  if (rowBuffer.length > 0) {
    pagePaths.push(
      await flushViewPage({
        workbookId,
        viewId,
        sheetName: sheet.name,
        page: pagePaths.length,
        rowStart: pageStartRow,
        rowEnd: pageStartRow + rowBuffer.length,
        rows: rowBuffer.splice(0, rowBuffer.length),
      }),
    )
  }

  const viewManifest: WorkbookSheetViewManifest = {
    id: viewId,
    kind: 'view',
    key: `${sheet.key}-view-${viewId}`,
    name: sheet.name,
    baseSheetName: sheet.name,
    rowCount: outputRowIndex,
    columnCount: sheet.columnCount,
    populatedCellCount,
    pageSize: sheet.pageSize,
    pageCount: pagePaths.length,
    columnWidths: sheet.columnWidths,
    pages: pagePaths,
    searchEnabled: true,
    headerRows: normalizedConfig.headerRows,
    filters: normalizedConfig.filters,
    sort: normalizedConfig.sort,
  }

  await putWorkbookJson(manifestPath, viewManifest)

  return viewManifest
}

export async function loadWorkbookView(
  workbookId: string,
  viewId: string,
): Promise<WorkbookSheetViewManifest> {
  const manifest = await readWorkbookJson<WorkbookSheetViewManifest>(
    `${workbookId}/views/${viewId}/manifest.json`,
  )

  return manifest
}

export async function searchWorkbookSheet(
  workbookId: string,
  source: WorkbookSheetManifest | WorkbookSheetViewManifest,
  query: string,
  limit = 50,
): Promise<WorkbookSearchMatch[]> {
  const normalizedQuery = query.trim().toLowerCase()

  if (!normalizedQuery) {
    return []
  }

  const matches: WorkbookSearchMatch[] = []
  const cappedLimit = Math.min(Math.max(limit, 1), 200)

  for (const pagePath of source.pages) {
    const page = await readWorkbookJson<WorkbookSheetPage>(pagePath)

    for (let rowOffset = 0; rowOffset < page.rows.length; rowOffset += 1) {
      const rowIndex = page.rowStart + rowOffset
      const row = page.rows[rowOffset] ?? []

      for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
        const value = row[columnIndex] ?? ''

        if (!value || !value.toLowerCase().includes(normalizedQuery)) {
          continue
        }

        matches.push({
          rowIndex,
          columnIndex,
          address: addressFromPosition(rowIndex, columnIndex),
          value,
        })

        if (matches.length >= cappedLimit) {
          return matches
        }
      }
    }
  }

  return matches
}

function normalizeViewConfig(config: WorkbookSheetViewConfig): WorkbookSheetViewConfig {
  const filters = config.filters
    .map((rule) => ({
      columnIndex: Math.max(0, Math.trunc(rule.columnIndex)),
      term: rule.term.trim(),
    }))
    .filter((rule) => rule.term.length > 0)

  const sort: WorkbookSortRule | null = config.sort
    ? {
        columnIndex: Math.max(0, Math.trunc(config.sort.columnIndex)),
        direction: config.sort.direction === 'desc' ? 'desc' : 'asc',
      }
    : null

  return {
    headerRows: config.headerRows > 0 ? 1 : 0,
    filters,
    sort,
  }
}

function matchesFilters(row: string[], filters: WorkbookFilterRule[]): boolean {
  if (filters.length === 0) {
    return true
  }

  return filters.every((filter) => {
    const value = row[filter.columnIndex] ?? ''
    return value.toLowerCase().includes(filter.term.toLowerCase())
  })
}

async function flushViewPage(page: WorkbookSheetPage & {
  workbookId: string
  viewId: string
}): Promise<string> {
  const relativePath = `${page.workbookId}/views/${page.viewId}/page-${page.page}.json`

  await putWorkbookJson(relativePath, {
    sheetName: page.sheetName,
    page: page.page,
    rowStart: page.rowStart,
    rowEnd: page.rowEnd,
    rows: page.rows,
  })

  return relativePath
}

function countPopulatedCells(row: string[]): number {
  let total = 0

  for (const value of row) {
    if (value !== '') {
      total += 1
    }
  }

  return total
}

function trimTrailingEmpty(values: string[]): string[] {
  let last = values.length - 1

  while (last >= 0 && values[last] === '') {
    last -= 1
  }

  return values.slice(0, last + 1)
}

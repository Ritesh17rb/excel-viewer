import { NextResponse } from 'next/server'

import { loadWorkbookView } from '@/lib/server/workbook-views'
import { readWorkbookJson } from '@/lib/server/workbook-store'
import type {
  WorkbookManifest,
  WorkbookSheetManifest,
  WorkbookSheetPage,
  WorkbookSheetViewManifest,
} from '@/lib/workbook-types'

export const runtime = 'nodejs'

export async function GET(
  request: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params
    const { searchParams } = new URL(request.url)
    const sheetName = searchParams.get('name')
    const viewId = searchParams.get('view')
    const page = Number(searchParams.get('page') ?? '0')

    if (!sheetName || Number.isNaN(page) || page < 0) {
      return NextResponse.json(
        { error: 'Invalid sheet page request.' },
        { status: 400 },
      )
    }

    const manifest = await readWorkbookJson<WorkbookManifest>(`${id}/manifest.json`)
    let source: WorkbookSheetManifest | WorkbookSheetViewManifest | null = null

    if (viewId) {
      const view = await loadWorkbookView(id, viewId)

      if (view.baseSheetName !== sheetName) {
        return NextResponse.json({ error: 'View does not match sheet.' }, { status: 400 })
      }

      source = view
    } else {
      source = manifest.sheets[sheetName] ?? null
    }

    if (!source) {
      return NextResponse.json({ error: 'Sheet not found.' }, { status: 404 })
    }

    const pagePath = source.pages[page]

    if (!pagePath) {
      return NextResponse.json({ error: 'Page not found.' }, { status: 404 })
    }

    const sheetPage = await readWorkbookJson<WorkbookSheetPage>(pagePath)
    return NextResponse.json({ sheetPage })
  } catch (error) {
    return NextResponse.json(
      {
        error: error instanceof Error ? error.message : 'Sheet page unavailable.',
      },
      { status: 500 },
    )
  }
}

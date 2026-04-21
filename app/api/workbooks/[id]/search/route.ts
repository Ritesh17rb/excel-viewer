import { NextResponse } from 'next/server'

import { loadWorkbookView, searchWorkbookSheet } from '@/lib/server/workbook-views'
import { readWorkbookJson } from '@/lib/server/workbook-store'
import type { WorkbookManifest } from '@/lib/workbook-types'

export const runtime = 'nodejs'
export const maxDuration = 300

export async function POST(
  request: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params
    const body = (await request.json()) as {
      sheetName?: string
      viewId?: string
      query?: string
      limit?: number
    }

    if (!body.sheetName || !body.query) {
      return NextResponse.json(
        { error: 'Missing sheet name or search query.' },
        { status: 400 },
      )
    }

    const manifest = await readWorkbookJson<WorkbookManifest>(`${id}/manifest.json`)
    const sheet = manifest.sheets[body.sheetName]

    if (!sheet) {
      return NextResponse.json({ error: 'Sheet not found.' }, { status: 404 })
    }

    const source = body.viewId
      ? await loadWorkbookView(id, body.viewId)
      : sheet

    const matches = await searchWorkbookSheet(id, source, body.query, body.limit ?? 40)

    return NextResponse.json({ matches })
  } catch (error) {
    return NextResponse.json(
      {
        error: error instanceof Error ? error.message : 'Search failed.',
      },
      { status: 500 },
    )
  }
}

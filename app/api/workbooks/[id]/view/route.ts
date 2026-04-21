import { NextResponse } from 'next/server'

import { createWorkbookView } from '@/lib/server/workbook-views'
import { readWorkbookJson } from '@/lib/server/workbook-store'
import type { WorkbookManifest, WorkbookSheetViewConfig } from '@/lib/workbook-types'

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
      config?: WorkbookSheetViewConfig
    }

    if (!body.sheetName || !body.config) {
      return NextResponse.json(
        { error: 'Missing sheet name or view config.' },
        { status: 400 },
      )
    }

    const manifest = await readWorkbookJson<WorkbookManifest>(`${id}/manifest.json`)
    const sheet = manifest.sheets[body.sheetName]

    if (!sheet) {
      return NextResponse.json({ error: 'Sheet not found.' }, { status: 404 })
    }

    const view = await createWorkbookView(id, sheet, body.config)
    return NextResponse.json({ view })
  } catch (error) {
    return NextResponse.json(
      {
        error: error instanceof Error ? error.message : 'View generation failed.',
      },
      { status: 500 },
    )
  }
}

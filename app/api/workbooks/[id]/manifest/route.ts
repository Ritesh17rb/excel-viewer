import { NextResponse } from 'next/server'

import { readWorkbookJson } from '@/lib/server/workbook-store'
import type { WorkbookManifest } from '@/lib/workbook-types'

export const runtime = 'nodejs'

export async function GET(
  _request: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params
    const manifest = await readWorkbookJson<WorkbookManifest>(`${id}/manifest.json`)
    return NextResponse.json({ manifest })
  } catch (error) {
    return NextResponse.json(
      {
        error:
          error instanceof Error ? error.message : 'Workbook manifest not found.',
      },
      { status: 404 },
    )
  }
}


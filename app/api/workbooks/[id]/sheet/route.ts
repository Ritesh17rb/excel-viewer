import { NextResponse } from 'next/server'

import { readWorkbookJson } from '@/lib/server/workbook-store'
import type { WorkbookManifest, WorkbookSheetPage } from '@/lib/workbook-types'

export const runtime = 'nodejs'

export async function GET(
  request: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params
    const { searchParams } = new URL(request.url)
    const sheetName = searchParams.get('name')
    const page = Number(searchParams.get('page') ?? '0')

    if (!sheetName || Number.isNaN(page) || page < 0) {
      return NextResponse.json(
        { error: 'Invalid sheet page request.' },
        { status: 400 },
      )
    }

    const manifest = await readWorkbookJson<WorkbookManifest>(`${id}/manifest.json`)
    const sheet = manifest.sheets[sheetName]

    if (!sheet) {
      return NextResponse.json({ error: 'Sheet not found.' }, { status: 404 })
    }

    const pagePath = sheet.pages[page]

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

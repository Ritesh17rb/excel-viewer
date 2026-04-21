import { NextResponse } from 'next/server'

import { fileExtension } from '@/lib/format'
import { ingestWorkbook } from '@/lib/server/workbook-ingest'
import type { WorkbookUploadSource } from '@/lib/workbook-types'

export const runtime = 'nodejs'
export const maxDuration = 300

export async function POST(request: Request) {
  const body = (await request.json()) as {
    source?: WorkbookUploadSource
  }

  if (!body.source) {
    return NextResponse.json({ error: 'Missing workbook source.' }, { status: 400 })
  }

  const format = body.source.format || fileExtension(body.source.fileName)

  if (format !== 'xlsx') {
    return NextResponse.json(
      { error: 'The server-backed viewer currently supports .xlsx files only.' },
      { status: 400 },
    )
  }

  try {
    const manifest = await ingestWorkbook({
      ...body.source,
      format,
    })

    return NextResponse.json({ manifest })
  } catch (error) {
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : 'Workbook ingestion failed.',
      },
      { status: 500 },
    )
  }
}

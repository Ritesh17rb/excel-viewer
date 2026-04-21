import { NextResponse } from 'next/server'

import { fileExtension } from '@/lib/format'
import { saveLocalUpload } from '@/lib/server/workbook-store'

export const runtime = 'nodejs'

export async function POST(request: Request) {
  if (process.env.NODE_ENV === 'production') {
    return NextResponse.json(
      { error: 'Local uploads are only available during development.' },
      { status: 403 },
    )
  }

  const formData = await request.formData()
  const file = formData.get('file')

  if (!(file instanceof File)) {
    return NextResponse.json({ error: 'Missing file.' }, { status: 400 })
  }

  if (fileExtension(file.name) !== 'xlsx') {
    return NextResponse.json(
      { error: 'Only .xlsx files are supported in the server-backed viewer.' },
      { status: 400 },
    )
  }

  const source = await saveLocalUpload(file)
  return NextResponse.json({ source })
}


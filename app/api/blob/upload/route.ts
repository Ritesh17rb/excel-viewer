import { NextResponse } from 'next/server'
import {
  handleUpload,
  type HandleUploadBody,
} from '@vercel/blob/client'

export const runtime = 'nodejs'
export const maxDuration = 60

const ALLOWED_TYPES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
]

export async function POST(request: Request) {
  if (!process.env.BLOB_READ_WRITE_TOKEN) {
    return NextResponse.json(
      { error: 'BLOB_READ_WRITE_TOKEN is not configured.' },
      { status: 500 },
    )
  }

  const body = (await request.json()) as HandleUploadBody

  try {
    const response = await handleUpload({
      body,
      request,
      onBeforeGenerateToken: async () => ({
        allowedContentTypes: ALLOWED_TYPES,
        maximumSizeInBytes: 1024 * 1024 * 512,
        addRandomSuffix: true,
        validUntil: Date.now() + 5 * 60 * 1000,
      }),
    })

    return NextResponse.json(response)
  } catch (error) {
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : 'Client upload token generation failed.',
      },
      { status: 500 },
    )
  }
}

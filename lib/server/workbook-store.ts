import { createReadStream } from 'node:fs'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import path from 'node:path'
import { Readable, type Stream } from 'node:stream'
import type { ReadableStream as WebReadableStream } from 'node:stream/web'

import { get, put } from '@vercel/blob'

import type { BlobAccess, WorkbookUploadSource } from '@/lib/workbook-types'

const LOCAL_ROOT = path.join(process.cwd(), '.cache', 'workbooks')
const LOCAL_UPLOAD_ROOT = path.join(process.cwd(), '.cache', 'uploads')
const BLOB_ROOT = 'processed-workbooks'

export function hasBlobStore(): boolean {
  return Boolean(process.env.BLOB_READ_WRITE_TOKEN)
}

export function getBlobAccess(): BlobAccess {
  return process.env.BLOB_ACCESS === 'private' ? 'private' : 'public'
}

export async function putWorkbookJson<T>(relativePath: string, value: T): Promise<string> {
  const payload = JSON.stringify(value)

  if (hasBlobStore()) {
    await put(`${BLOB_ROOT}/${relativePath}`, payload, {
      access: getBlobAccess(),
      allowOverwrite: true,
      addRandomSuffix: false,
      contentType: 'application/json; charset=utf-8',
    })

    return `${BLOB_ROOT}/${relativePath}`
  }

  const targetPath = path.join(LOCAL_ROOT, relativePath)
  await mkdir(path.dirname(targetPath), { recursive: true })
  await writeFile(targetPath, payload, 'utf8')
  return targetPath
}

export async function readWorkbookJson<T>(relativePath: string): Promise<T> {
  if (hasBlobStore()) {
    const blob = await get(`${BLOB_ROOT}/${relativePath}`, {
      access: getBlobAccess(),
      useCache: false,
    })

    if (!blob || blob.statusCode !== 200 || !blob.stream) {
      throw new Error(`Artifact not found: ${relativePath}`)
    }

    const text = await new Response(blob.stream).text()
    return JSON.parse(text) as T
  }

  const text = await readFile(path.join(LOCAL_ROOT, relativePath), 'utf8')
  return JSON.parse(text) as T
}

export async function openSourceStream(
  source: WorkbookUploadSource,
): Promise<Stream> {
  if (source.kind === 'local') {
    if (!source.localPath) {
      throw new Error('Missing local path for uploaded workbook.')
    }

    return createReadStream(source.localPath)
  }

  if (source.access === 'private') {
    const blob = await get(source.pathname ?? source.url ?? '', {
      access: 'private',
      useCache: false,
    })

    if (!blob || blob.statusCode !== 200 || !blob.stream) {
      throw new Error('Private blob could not be opened.')
    }

    return Readable.fromWeb(blob.stream as unknown as WebReadableStream)
  }

  const response = await fetch(source.url ?? '')

  if (!response.ok || !response.body) {
    throw new Error('Uploaded blob could not be downloaded for ingestion.')
  }

  return Readable.fromWeb(response.body as unknown as WebReadableStream)
}

export async function saveLocalUpload(file: File): Promise<WorkbookUploadSource> {
  const fileBuffer = Buffer.from(await file.arrayBuffer())
  const localPath = path.join(LOCAL_UPLOAD_ROOT, `${Date.now()}-${file.name}`)

  await mkdir(path.dirname(localPath), { recursive: true })
  await writeFile(localPath, fileBuffer)

  return {
    kind: 'local',
    fileName: file.name,
    fileSize: file.size,
    format: path.extname(file.name).replace('.', '').toLowerCase(),
    localPath,
  }
}

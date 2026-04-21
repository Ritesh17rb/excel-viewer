const byteFormatter = new Intl.NumberFormat('en-US', {
  maximumFractionDigits: 1,
})

const countFormatter = new Intl.NumberFormat('en-US')

export function formatBytes(value: number): string {
  if (!Number.isFinite(value) || value <= 0) {
    return '0 B'
  }

  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const exponent = Math.min(
    Math.floor(Math.log(value) / Math.log(1024)),
    units.length - 1,
  )
  const amount = value / 1024 ** exponent

  return `${byteFormatter.format(amount)} ${units[exponent]}`
}

export function formatCount(value: number): string {
  return countFormatter.format(Math.max(0, Math.trunc(value)))
}

export function fileExtension(fileName: string): string {
  const [, extension = 'spreadsheet'] = /\.([^.]+)$/.exec(fileName) ?? []
  return extension.toLowerCase()
}

export function toColumnLabel(index: number): string {
  let current = index + 1
  let label = ''

  while (current > 0) {
    const remainder = (current - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    current = Math.floor((current - 1) / 26)
  }

  return label
}

export function addressFromPosition(
  rowIndex: number,
  columnIndex: number,
): string {
  return `${toColumnLabel(columnIndex)}${rowIndex + 1}`
}

export function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max)
}

export function sanitizeSegment(value: string): string {
  const normalized = value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')

  return normalized || 'sheet'
}


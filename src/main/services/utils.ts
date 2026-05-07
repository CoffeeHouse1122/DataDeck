import path from 'node:path'

export const MONTH_LABELS = ['Jan.', 'Feb.', 'Mar.', 'Apr.', 'May', 'Jun.', 'Jul.', 'Aug.', 'Sept.', 'Oct.', 'Nov.', 'Dec.']

export function normalizeText(value: unknown): string {
  if (value === null || value === undefined) {
    return ''
  }

  if (typeof value === 'object' && 'result' in (value as Record<string, unknown>)) {
    return normalizeText((value as { result: unknown }).result)
  }

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10)
  }

  return String(value).trim()
}

export function normalizeNumber(value: unknown): number {
  if (value === null || value === undefined || value === '') {
    return 0
  }

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0
  }

  if (typeof value === 'object' && 'result' in (value as Record<string, unknown>)) {
    return normalizeNumber((value as { result: unknown }).result)
  }

  const text = String(value).replace(/,/g, '').trim()
  if (!text) {
    return 0
  }

  if (text.endsWith('%')) {
    return Number.parseFloat(text.slice(0, -1)) / 100
  }

  const parsed = Number.parseFloat(text)
  return Number.isFinite(parsed) ? parsed : 0
}

export function toInteger(value: number): number {
  return Math.round(Number.isFinite(value) ? value : 0)
}

export function toPercent(value: number): number {
  return Number.isFinite(value) ? value : 0
}

export function columnLetter(index: number): string {
  let current = index
  let result = ''
  while (current > 0) {
    const remainder = (current - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    current = Math.floor((current - 1) / 26)
  }
  return result
}

export function deepClone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value))
}

export function detectReportKey(sourcePath: string, fallbackHeader: string): string {
  const fileMatch = sourcePath.match(/(\d{6})-(\d{6})/)
  if (fileMatch) {
    return fileMatch[2]
  }

  const headerMatch = fallbackHeader.match(/(\d{6})-(\d{6})/)
  if (headerMatch) {
    return headerMatch[2]
  }

  return new Date().toISOString().slice(0, 7).replace('-', '')
}

export function monthLabelFromKey(reportKey: string): string {
  const month = Number.parseInt(reportKey.slice(4, 6), 10)
  return MONTH_LABELS[Math.max(0, Math.min(11, month - 1))] ?? 'Jan.'
}

export function titleMonthFromKey(reportKey: string): string {
  const year = reportKey.slice(0, 4)
  const month = Number.parseInt(reportKey.slice(4, 6), 10)
  const longMonth = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][month - 1] ?? 'January'
  return `${longMonth} ${year}`
}

export function sanitizeFileStem(sourcePath: string): string {
  return path.basename(sourcePath, path.extname(sourcePath))
}

export function ensureArrayLength<T>(items: T[], size: number, filler: () => T): T[] {
  const next = [...items]
  while (next.length < size) {
    next.push(filler())
  }
  return next.slice(0, size)
}

export function buildOutputPath(outputDir: string, baseName: string, reportKey: string, extension: string): string {
  return path.join(outputDir, `${baseName}-${reportKey}.${extension}`)
}

export function compactName(value: string): string {
  return value.replace(/\s+/g, ' ').replace(/\([^)]*\)/g, '').replace(/\s+,/g, ',').trim()
}

export function normalizePersonName(value: string): string {
  return compactName(value).toLowerCase()
}

export function percentageDelta(current: number, previous: number): number {
  if (!previous) {
    return 0
  }
  return (current - previous) / previous
}

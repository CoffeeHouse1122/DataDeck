import type ExcelJS from 'exceljs'

type TableCell = {
  value: string
  align?: 'left' | 'center' | 'right'
  background?: string
  color?: string
  fontSize?: number
  weight?: 400 | 600 | 700
  barRatio?: number
  barColor?: string
}

type WorksheetRangeRenderOptions = {
  startRow: number
  endRow: number
  startCol: number
  endCol: number
  rowIndices?: number[]
  colIndices?: number[]
  cellOverlay?: (cell: ExcelJS.Cell, row: number, col: number) => {
    background?: string
    dataBarRatio?: number
    dataBarColor?: string
  }
  verticalMerges?: Array<{
    startRow: number
    endRow: number
    col: number
    value?: string
  }>
  colWidthOverrides?: Map<number, number>
  defaultColWidth?: number
  defaultRowHeight?: number
  fontScale?: number
  fontScaleOverrides?: Map<number, number>
  rowFontScaleOverrides?: Map<number, number>
  rowHeightOverrides?: Map<number, number>
  forceBold?: boolean
  gridStrokeColor?: string
  gridStrokeWidth?: number
}

type TableRenderOptions = {
  columns: Array<{ key: string; label: string; width: number }>
  rows: Array<Record<string, TableCell>>
  title?: string
  subtitle?: string
  rowHeight?: number
  headerHeight?: number
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function colorFromArgb(argb?: string, fallback = '#ffffff'): string {
  if (!argb) {
    return fallback
  }
  const rgb = argb.length === 8 ? argb.slice(2) : argb
  return `#${rgb}`
}

function cellFill(cell: ExcelJS.Cell): string {
  const fill = cell.fill as ExcelJS.Fill | undefined
  if (fill?.type === 'pattern' && fill.fgColor && 'argb' in fill.fgColor) {
    return colorFromArgb(fill.fgColor.argb, '#ffffff')
  }
  return '#ffffff'
}

function cellFontColor(cell: ExcelJS.Cell): string {
  const color = cell.font?.color
  if (color && 'argb' in color) {
    return colorFromArgb(color.argb, '#1f2328')
  }
  return '#1f2328'
}

function worksheetDisplayValue(value: unknown): unknown {
  if (value === null || value === undefined) {
    return value
  }
  if (value instanceof Date || typeof value !== 'object') {
    return value
  }

  const record = value as Record<string, unknown>
  if ('result' in record) {
    return worksheetDisplayValue(record.result)
  }
  if ('text' in record) {
    return record.text
  }
  if ('richText' in record && Array.isArray(record.richText)) {
    return record.richText.map((part) => String((part as { text?: unknown }).text ?? '')).join('')
  }
  if ('error' in record || 'formula' in record) {
    return 0
  }
  return ''
}

function formatWorksheetValue(cell: ExcelJS.Cell): string {
  const value = worksheetDisplayValue(cell.value)
  if (value === null || value === undefined) {
    return ''
  }
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10)
  }
  if (typeof value === 'number') {
    if (cell.numFmt?.includes('%')) {
      return `${(value * 100).toFixed(2)}%`
    }
    if (Number.isInteger(value)) {
      return String(value)
    }
    return String(Number(value.toFixed(2)))
  }
  return String(value)
}

function borderColor(cell: ExcelJS.Cell): string {
  const color = cell.border?.bottom?.color ?? cell.border?.top?.color ?? cell.border?.left?.color ?? cell.border?.right?.color
  if (color && 'argb' in color) {
    return colorFromArgb(color.argb, '#d0d7de')
  }
  return '#d0d7de'
}

export function renderWorksheetRangeSvg(sheet: ExcelJS.Worksheet, options: WorksheetRangeRenderOptions): string {
  const rows = options.rowIndices ?? Array.from({ length: options.endRow - options.startRow + 1 }, (_, index) => options.startRow + index)
  const cols = options.colIndices ?? Array.from({ length: options.endCol - options.startCol + 1 }, (_, index) => options.startCol + index)
  const colWidths: number[] = []
  for (const col of cols) {
    const width = options.colWidthOverrides?.get(col) ?? sheet.getColumn(col).width ?? options.defaultColWidth ?? 10
    colWidths.push(Math.max(34, width * 7.2))
  }

  const rowHeights: number[] = []
  for (const row of rows) {
    const rowHeightOverride = options.rowHeightOverrides?.get(row)
    rowHeights.push(rowHeightOverride !== undefined
      ? Math.max(18, rowHeightOverride)
      : Math.max(18, (sheet.getRow(row).height ?? options.defaultRowHeight ?? 20) * 1.35))
  }

  const width = colWidths.reduce((sum, value) => sum + value, 0)
  const height = rowHeights.reduce((sum, value) => sum + value, 0)
  const fontScale = options.fontScale ?? 1
  const rowY = new Map<number, number>()
  const colX = new Map<number, number>()
  let cursorY = 0
  for (const [index, row] of rows.entries()) {
    rowY.set(row, cursorY)
    cursorY += rowHeights[index]
  }
  let cursorX = 0
  for (const [index, col] of cols.entries()) {
    colX.set(col, cursorX)
    cursorX += colWidths[index]
  }
  const hiddenMergeCells = new Set<string>()
  for (const merge of options.verticalMerges ?? []) {
    for (let row = merge.startRow + 1; row <= merge.endRow; row += 1) {
      hiddenMergeCells.add(`${row}:${merge.col}`)
    }
  }

  let body = `<rect width="${width}" height="${height}" fill="#ffffff" />`
  let y = 0
  for (const [rowOffset, row] of rows.entries()) {
    let x = 0
    const rowHeight = rowHeights[rowOffset]
    for (const [colOffset, col] of cols.entries()) {
      const colWidth = colWidths[colOffset]
      const cell = sheet.getCell(row, col)
      const maybeMerged = cell as ExcelJS.Cell & { master?: ExcelJS.Cell; isMerged?: boolean }
      const isMergedChild = Boolean(maybeMerged.isMerged && maybeMerged.master && maybeMerged.master.address !== cell.address)
      const cellFontScale = options.rowFontScaleOverrides?.get(row) ?? options.fontScaleOverrides?.get(col) ?? fontScale
      const fontSize = Math.max(8, (cell.font?.size ?? 10) * cellFontScale)
      const align = cell.alignment?.horizontal === 'right'
        ? 'right'
        : cell.alignment?.horizontal === 'center'
          ? 'center'
          : 'left'
      const anchor = align === 'right' ? 'end' : align === 'center' ? 'middle' : 'start'
      const textX = align === 'right' ? x + colWidth - 5 : align === 'center' ? x + colWidth / 2 : x + 5
      const textY = y + rowHeight / 2 + fontSize * 0.34
      const weight = options.forceBold || cell.font?.bold ? 700 : 400
      const value = isMergedChild || hiddenMergeCells.has(`${row}:${col}`) ? '' : formatWorksheetValue(cell)
      const overlay = options.cellOverlay?.(cell, row, col)
      const background = overlay?.background ?? cellFill(cell)

      body += `<rect x="${x}" y="${y}" width="${colWidth}" height="${rowHeight}" fill="${background}" stroke="${options.gridStrokeColor ?? borderColor(cell)}" stroke-width="${options.gridStrokeWidth ?? 1}" />`
      if (overlay?.dataBarRatio && overlay.dataBarRatio > 0) {
        const barWidth = Math.max(4, (colWidth - 8) * Math.max(0, Math.min(1, overlay.dataBarRatio)))
        body += `<rect x="${x + 4}" y="${y + 5}" width="${barWidth}" height="${Math.max(4, rowHeight - 10)}" fill="${overlay.dataBarColor ?? '#638EC6'}" opacity="0.58" />`
      }
      if (value) {
        body += `<text x="${textX}" y="${textY}" fill="${cellFontColor(cell)}" font-family="Arial, Source Han Sans CN, sans-serif" font-size="${fontSize}" font-weight="${weight}" text-anchor="${anchor}">${escapeXml(value)}</text>`
      }
      x += colWidth
    }
    y += rowHeight
  }

  for (const merge of options.verticalMerges ?? []) {
    const x = colX.get(merge.col)
    const startY = rowY.get(merge.startRow)
    const endY = rowY.get(merge.endRow)
    if (x === undefined || startY === undefined || endY === undefined) {
      continue
    }
    const colWidth = colWidths[cols.indexOf(merge.col)]
    const mergeHeight = endY - startY + rowHeights[rows.indexOf(merge.endRow)]
    const cell = sheet.getCell(merge.startRow, merge.col)
    const fontSize = Math.max(8, (cell.font?.size ?? 10) * fontScale)
    const value = merge.value ?? formatWorksheetValue(cell)
    const overlay = options.cellOverlay?.(cell, merge.startRow, merge.col)
    body += `<rect x="${x}" y="${startY}" width="${colWidth}" height="${mergeHeight}" fill="${overlay?.background ?? cellFill(cell)}" stroke="${options.gridStrokeColor ?? borderColor(cell)}" stroke-width="${options.gridStrokeWidth ?? 1}" />`
    if (value) {
      body += `<text x="${x + colWidth / 2}" y="${startY + mergeHeight / 2 + fontSize * 0.34}" fill="${cellFontColor(cell)}" font-family="Arial, Source Han Sans CN, sans-serif" font-size="${fontSize}" font-weight="${options.forceBold || cell.font?.bold ? 700 : 400}" text-anchor="middle">${escapeXml(value)}</text>`
    }
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${body}</svg>`
}

function wrapText(text: string, width: number, fontSize: number): string[] {
  const roughChars = Math.max(4, Math.floor(width / Math.max(7, fontSize * 0.72)))
  if (text.length <= roughChars) {
    return [text]
  }

  const chunks: string[] = []
  let cursor = text
  while (cursor.length > roughChars) {
    chunks.push(cursor.slice(0, roughChars))
    cursor = cursor.slice(roughChars)
  }
  if (cursor) {
    chunks.push(cursor)
  }
  return chunks
}

export function interpolateColor(value: number, min: number, max: number, start: string, end: string): string {
  if (max <= min) {
    return start
  }

  const ratio = Math.max(0, Math.min(1, (value - min) / (max - min)))
  const parse = (hex: string) => [0, 2, 4].map((offset) => Number.parseInt(hex.slice(offset, offset + 2), 16))
  const [sr, sg, sb] = parse(start)
  const [er, eg, eb] = parse(end)
  const toHex = (component: number) => component.toString(16).padStart(2, '0')
  return [
    sr + (er - sr) * ratio,
    sg + (eg - sg) * ratio,
    sb + (eb - sb) * ratio
  ].map((valuePart) => toHex(Math.round(valuePart))).join('')
}

export function renderTableSvg(options: TableRenderOptions): string {
  const headerHeight = options.headerHeight ?? 36
  const rowHeight = options.rowHeight ?? 28
  const width = options.columns.reduce((sum, column) => sum + column.width, 0)
  const titleBlock = options.title ? 56 : 0
  const height = titleBlock + headerHeight + options.rows.length * rowHeight + 12

  let x = 0
  let header = ''
  for (const column of options.columns) {
    header += `<rect x="${x}" y="${titleBlock}" width="${column.width}" height="${headerHeight}" fill="#0d1117" stroke="#30363d" />`
    header += `<text x="${x + 10}" y="${titleBlock + 23}" fill="#f0f6fc" font-size="14" font-weight="700">${escapeXml(column.label)}</text>`
    x += column.width
  }

  let body = ''
  options.rows.forEach((row, rowIndex) => {
    let cursor = 0
    const y = titleBlock + headerHeight + rowIndex * rowHeight
    body += `<rect x="0" y="${y}" width="${width}" height="${rowHeight}" fill="${rowIndex % 2 === 0 ? '#ffffff' : '#f6f8fa'}" />`

    for (const column of options.columns) {
      const cell = row[column.key] ?? { value: '' }
      const background = cell.background ?? 'transparent'
      body += `<rect x="${cursor}" y="${y}" width="${column.width}" height="${rowHeight}" fill="${background}" stroke="#d0d7de" />`

      if (cell.barRatio && cell.barRatio > 0) {
        const barWidth = Math.max(8, (column.width - 16) * Math.max(0, Math.min(1, cell.barRatio)))
        body += `<rect x="${cursor + 6}" y="${y + 6}" width="${barWidth}" height="${rowHeight - 12}" rx="4" fill="${cell.barColor ?? '#2f81f7'}" opacity="0.26" />`
      }

      const fontSize = cell.fontSize ?? 11
      const lines = wrapText(cell.value, column.width - 16, fontSize)
      const color = cell.color ?? '#1f2328'
      const weight = cell.weight ?? 400
      const anchor = cell.align === 'right' ? 'end' : cell.align === 'center' ? 'middle' : 'start'
      const textX = cell.align === 'right'
        ? cursor + column.width - 8
        : cell.align === 'center'
          ? cursor + column.width / 2
          : cursor + 8
      const baseY = y + rowHeight / 2 - ((lines.length - 1) * (fontSize + 1)) / 2 + 4

      lines.forEach((line, lineIndex) => {
        body += `<text x="${textX}" y="${baseY + lineIndex * (fontSize + 1)}" fill="${color}" font-size="${fontSize}" font-weight="${weight}" text-anchor="${anchor}">${escapeXml(line)}</text>`
      })

      cursor += column.width
    }
  })

  const titleBlockSvg = options.title
    ? `<text x="0" y="20" fill="#0d1117" font-size="22" font-weight="700">${escapeXml(options.title)}</text>${
        options.subtitle
          ? `<text x="0" y="42" fill="#57606a" font-size="11">${escapeXml(options.subtitle)}</text>`
          : ''
      }`
    : ''

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
    <rect width="${width}" height="${height}" fill="#ffffff" />
    ${titleBlockSvg}
    ${header}
    ${body}
  </svg>`
}

export function svgToDataUri(svg: string): string {
  return `data:image/svg+xml;base64,${Buffer.from(svg, 'utf8').toString('base64')}`
}

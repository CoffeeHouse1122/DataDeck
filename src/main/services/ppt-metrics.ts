import { XMLParser } from 'fast-xml-parser'

const textParser = new XMLParser({ parseTagValue: false, trimValues: false, htmlEntities: true })

function cellText(cell: string): string {
  return [...cell.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)]
    .map((match) => String(textParser.parse(`<text>${match[1]}</text>`).text ?? '')).join('')
}

function replaceCellText(cell: string, value: string): string {
  if (!/<a:txBody\b/.test(cell)) throw new Error('PPT 指标单元格缺少文本区域，请检查模板。')
  return cell.replace(/<a:txBody\b[^>]*>[\s\S]*?<\/a:txBody>/, (body) => {
    const paragraphs = [...body.matchAll(/<a:p\b[^>]*(?:\/>|>[\s\S]*?<\/a:p>)/g)].map((match) => match[0])
    const paragraph = paragraphs.find((item) => /<a:t\b/.test(item)) ?? paragraphs[0]
    if (!paragraph) throw new Error('PPT 指标单元格缺少段落，请检查模板。')
    // Keep the paragraph and first text run's formatting, replacing the complete value.
    const properties = paragraph.match(/<a:pPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:pPr>)/)?.[0] ?? ''
    const run = paragraph.match(/<a:(?:r|fld)\b[^>]*>[\s\S]*?<\/a:(?:r|fld)>/)?.[0] ?? ''
    const runProperties = run.match(/<a:rPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/)?.[0] ?? ''
    const endProperties = paragraph.match(/<a:endParaRPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:endParaRPr>)/)?.[0] ?? ''
    const replacement = `<a:p>${properties}<a:r>${runProperties}<a:t>${value}</a:t></a:r>${endProperties}</a:p>`
    let inserted = false
    return body.replace(/<a:p\b[^>]*(?:\/>|>[\s\S]*?<\/a:p>)/g, () => {
      if (inserted) return ''
      inserted = true
      return replacement
    })
  })
}

export function updateMetricSlide(xml: string, metrics: {
  displayName: string
  publicationTcr: number
  revenueTcr: number
  waiverRate2026: number
}): string {
  const definitions = [
    { label: 'Publication TCR', keys: ['publicationtcr'], value: metrics.publicationTcr },
    { label: 'Revenue TCR', keys: ['revenuetcr', 'invoicetcr'], value: metrics.revenueTcr },
    { label: 'Waiver Rate 2026', keys: ['waiverrate2025', 'waiverrate2026'], value: metrics.waiverRate2026 }
  ]
  const counts = definitions.map(() => 0)
  const updated = xml.replace(/<a:tbl\b[^>]*>[\s\S]*?<\/a:tbl>/g, (table) =>
    table.replace(/<a:tr\b[^>]*>[\s\S]*?<\/a:tr>/g, (row) => {
      const cells = [...row.matchAll(/<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g)]
      const replacements = new Map<number, string>()
      cells.forEach((cell, index) => {
        const label = cellText(cell[0]).replace(/\s/g, '').toLowerCase()
        const metricIndex = definitions.findIndex((item) => item.keys.includes(label))
        if (metricIndex < 0) return
        const metric = definitions[metricIndex]
        if (!cells[index + 1] || !Number.isFinite(metric.value)) {
          throw new Error(`PPT 中 ${metrics.displayName} 的 ${metric.label} 缺少有效数值或对应单元格。`)
        }
        counts[metricIndex] += 1
        const value = `${(metric.value * 100).toFixed(2)}%`
        const valueCell = replaceCellText(cells[index + 1][0], value)
        if (cellText(valueCell) !== value) throw new Error(`PPT 指标写入校验失败：${metric.label}`)
        replacements.set(index + 1, valueCell)
        if (label === 'waiverrate2025') replacements.set(index, replaceCellText(cell[0], metric.label))
      })
      let index = 0
      return row.replace(/<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g, (cell) => replacements.get(index++) ?? cell)
    }))
  const invalid = definitions.filter((_item, index) => counts[index] !== 1)
  if (invalid.length) {
    throw new Error(`PPT 模板中 ${metrics.displayName} 的指标缺失或重复：${invalid.map((item) => item.label).join('、')}。请检查指标表格。`)
  }
  return updated
}

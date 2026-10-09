import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import { build } from 'vite'
import JSZip from 'jszip'
import { XMLParser, XMLValidator } from 'fast-xml-parser'

const root = fileURLToPath(new URL('../', import.meta.url))
await fs.mkdir(path.join(root, 'out'), { recursive: true })
const output = await fs.mkdtemp(path.join(root, 'out', 'ppt-metrics-'))
await build({
  configFile: false, root, logLevel: 'error',
  build: {
    ssr: path.join(root, 'src/main/services/ppt-metrics.ts'), outDir: output,
    rollupOptions: { output: { format: 'cjs', entryFileNames: 'metrics.cjs' } }
  }
})
const { updateMetricSlide } = createRequire(import.meta.url)(path.join(output, 'metrics.cjs'))
const values = { displayName: 'Example', publicationTcr: 0.1234, revenueTcr: 0.5678, waiverRate2026: 0.0912 }
const parser = new XMLParser({ ignoreAttributes: false, parseTagValue: false, trimValues: false, htmlEntities: true })
const array = (item) => item === undefined ? [] : Array.isArray(item) ? item : [item]
const text = (cell) => array(parser.parse(cell)['a:tc']['a:txBody']['a:p'])
  .flatMap((p) => array(p['a:r']).map((r) => typeof r['a:t'] === 'object' ? r['a:t']['#text'] ?? '' : r['a:t'] ?? '')).join('')
const run = (value) => `<a:r><a:rPr lang="en-US" sz="2400" b="0"/><a:t xml:space="preserve">${value}</a:t></a:r>`
const cell = (parts) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="l"/>${parts.map(run).join('')}<a:endParaRPr lang="en-US"/></a:p></a:txBody><a:tcPr marL="91440"/></a:tc>`
const row = (label, value) => `<a:tr h="300000">${cell(['Other 88.88%'])}${cell(['Example name'])}${cell(label)}${cell(value)}</a:tr>`
const rows = [
  row(['Publication TCR'], ['20.00%']),
  row(['Revenue TCR'], ['30.00%']),
  row(['Waiver Rate 2026'], ['40.00%'])
]
const slide = (rows) => `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp>${run('99.99%')}</p:sp><a:tbl><a:tblPr/><a:tblGrid/>${rows.join('')}</a:tbl></p:spTree></p:cSld></p:sld>`
function verify(before, expected = values) {
  const after = updateMetricSlide(before, expected)
  assert.equal(XMLValidator.validate(after), true)
  const found = new Map()
  for (const match of after.matchAll(/<a:tr\b[^>]*>[\s\S]*?<\/a:tr>/g)) {
    const cells = [...match[0].matchAll(/<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g)].map((c) => c[0])
    found.set(text(cells[2]).replace(/\s/g, '').toLowerCase(), text(cells[3]))
    assert.equal(text(cells[0]), 'Other 88.88%')
    assert.equal(text(cells[1]), 'Example name')
    assert.ok(cells[3].includes('<a:tcPr marL="91440"/>'))
    if (!before.includes('<a:p/>')) assert.ok(cells[3].includes('<a:pPr algn="l"/>'))
    if (cells[3].includes('<a:rPr')) assert.ok(cells[3].includes('sz="2400" b="0"'))
  }
  assert.equal(found.get('publicationtcr'), `${(expected.publicationTcr * 100).toFixed(2)}%`)
  assert.equal(found.get('revenuetcr'), `${(expected.revenueTcr * 100).toFixed(2)}%`)
  assert.equal(found.get('waiverrate2026'), `${(expected.waiverRate2026 * 100).toFixed(2)}%`)
  assert.ok(after.includes(run('99.99%')), 'Unrelated slide percentages must remain unchanged')
  assert.equal(updateMetricSlide(after, expected), after, 'Repeated updates must be stable')
  return after
}
verify(slide(rows))
verify(slide([row(['Publication TCR'], ['20', '.', '00', '%']), ...rows.slice(1)]))
verify(slide([
  row(['Waiver ', 'Rate ', '2025'], [' 40.00% ']),
  row(['R', 'evenue&#32;TCR'], ['30.00', '%']),
  row(['Publication', ' TCR'], ['20.00', '%'])
]))
verify(slide([row(['Publication TCR'], []), ...rows.slice(1)]))
verify(slide(rows).replace(cell(['20.00%']), cell([]).replace('<a:p><a:pPr algn="l"/><a:endParaRPr lang="en-US"/></a:p>', '<a:p/>')))
verify(slide(rows).replace(cell(['20.00%']), cell(['20']).replace('</a:p>', `<a:br/></a:p><a:p>${run('.00%')}</a:p>`)))
verify(slide(rows), { ...values, publicationTcr: 0, revenueTcr: 1.25, waiverRate2026: 0 })
assert.throws(() => updateMetricSlide(slide(rows.slice(1)), values), /缺失或重复/)
assert.throws(() => updateMetricSlide(slide([...rows, rows[0]]), values), /缺失或重复/)
assert.throws(() => updateMetricSlide(slide(rows).replace(cell(['20.00%']), ''), values), /缺少有效数值或对应单元格/)
assert.throws(() => updateMetricSlide(slide(rows).replace(cell(['20.00%']), '<a:tc><a:tcPr/></a:tc>'), values), /缺少文本区域/)
assert.throws(() => updateMetricSlide(slide(rows), { ...values, revenueTcr: NaN }), /缺少有效数值/)
console.log('PASS: split text/labels, whitespace/entities, empty values, row reordering, unrelated percentages, formatting, repeat updates and validation failures')

const templateIndex = process.argv.indexOf('--template')
if (templateIndex !== -1) {
  assert.ok(process.argv[templateIndex + 1], '--template requires a path')
  const template = await JSZip.loadAsync(await fs.readFile(path.resolve(process.argv[templateIndex + 1])))
  for (const page of [7, 9, 11, 13, 15]) {
    const name = `ppt/slides/slide${page}.xml`
    const original = await template.file(name).async('string')
    const updated = updateMetricSlide(original, values)
    assert.equal(XMLValidator.validate(updated), true)
    const withoutTables = (xml) => xml.replace(/<a:tbl\b[^>]*>[\s\S]*?<\/a:tbl>/g, '')
    assert.equal(withoutTables(updated), withoutTables(original), 'Unrelated shapes/charts must remain byte-identical')
    const updatedRows = [...updated.matchAll(/<a:tr\b[^>]*>[\s\S]*?<\/a:tr>/g)]
    const data = new Map(updatedRows.map((r) => {
      const cells = [...r[0].matchAll(/<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g)].map((c) => text(c[0]))
      return [cells[2].replace(/\s/g, '').toLowerCase(), cells[3]]
    }))
    assert.equal(data.get('publicationtcr'), '12.34%')
    assert.equal(data.get('revenuetcr'), '56.78%')
    assert.equal(data.get('waiverrate2026'), '9.12%')
    template.file(name, updated)
  }
  await fs.writeFile(path.join(output, 'verified-metrics.pptx'), await template.generateAsync({ type: 'nodebuffer' }))
  console.log('PASS: all five template metric tables; text outside tables preserved')
}
console.log(`Artifacts: ${output}`)

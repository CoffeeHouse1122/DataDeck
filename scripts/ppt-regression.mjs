import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { build } from 'vite'

// Exercise the public services with real template parts, without launching Electron.
const root = fileURLToPath(new URL('../', import.meta.url))
const output = await fs.mkdtemp(path.join(await ensureDir(path.join(root, 'out')), 'ppt-regression-'))
const parser = new XMLParser({ ignoreAttributes: false, removeNSPrefix: true, parseTagValue: false, parseAttributeValue: false })
const list = (value) => value === undefined ? [] : Array.isArray(value) ? value : [value]
const parse = (xml) => {
  assert.equal(XMLValidator.validate(xml), true, 'Invalid Office XML')
  return parser.parse(xml)
}
const resolvePart = (owner, target) => path.posix.normalize(target.startsWith('/') ? target.slice(1) : path.posix.join(path.posix.dirname(owner), target))
const relPath = (part) => path.posix.join(path.posix.dirname(part), '_rels', `${path.posix.basename(part)}.rels`)
const read = async (zip, name) => {
  assert.ok(zip.file(name), `Missing package part: ${name}`)
  return zip.file(name).async('string')
}
const relationships = async (zip, owner) => list(parse(await read(zip, relPath(owner))).Relationships.Relationship)

async function ensureDir(directory) {
  await fs.mkdir(directory, { recursive: true })
  return directory
}

const entry = path.join(output, 'services-entry.ts')
await fs.writeFile(entry, `export { buildPresentation } from ${JSON.stringify(path.join(root, 'src/electron/main/services/ppt.ts'))}; export { runPipeline } from ${JSON.stringify(path.join(root, 'src/electron/main/services/pipeline.ts'))};`)
await build({
  configFile: false,
  root,
  logLevel: 'error',
  build: {
    ssr: entry, outDir: output, emptyOutDir: false,
    rollupOptions: { output: { format: 'cjs', entryFileNames: 'services.cjs' } }
  }
})
const { buildPresentation, runPipeline } = createRequire(import.meta.url)(path.join(output, 'services.cjs'))
const templatePath = path.join(root, 'docs/Section Health月会.pptx')
const template = await JSZip.loadAsync(await fs.readFile(templatePath))

// Resolve displayed page order independently of slideN.xml numbering.
async function slidePaths(zip) {
  const ids = list(parse(await read(zip, 'ppt/presentation.xml')).presentation.sldIdLst.sldId)
  const rels = await relationships(zip, 'ppt/presentation.xml')
  return ids.map((item) => {
    const rel = rels.find((rel) => rel['@_Id'] === item['@_id'] && rel['@_Type'].endsWith('/slide'))
    assert.ok(rel, `Slide relationship ${item['@_id']} missing`)
    return resolvePart('ppt/presentation.xml', rel['@_Target'])
  })
}

async function chartPaths(zip, slide) {
  const xml = await read(zip, slide)
  const rels = await relationships(zip, slide)
  return [...xml.matchAll(/<c:chart\b[^>]*r:id="([^"]+)"/g)].map((match) => {
    const rel = rels.find((item) => item['@_Id'] === match[1])
    assert.ok(rel, `${slide}: chart relationship missing`)
    return resolvePart(slide, rel['@_Target'])
  })
}

function cacheValues(cache, expected, label, numeric = false) {
  assert.ok(cache, `${label}: cache missing`)
  assert.equal(Number(cache.ptCount?.['@_val']), expected.length, `${label}: point count`)
  const points = list(cache.pt)
  assert.equal(points.length, expected.length, `${label}: points`)
  points.forEach((point, index) => {
    assert.equal(Number(point['@_idx']), index, `${label}: point index`)
    assert.equal(numeric ? Number(point.v) : point.v, expected[index], `${label}: point ${index}`)
  })
}

function descendants(node, key) {
  if (!node || typeof node !== 'object') return []
  return Object.entries(node).flatMap(([name, value]) => [
    ...(name === key ? list(value) : []),
    ...list(value).flatMap((child) => descendants(child, key))
  ])
}

async function verifyChart(zip, chartPath, expected) {
  const chart = parse(await read(zip, chartPath)).chartSpace
  const series = descendants(chart.chart.plotArea, 'ser')
  const rels = await relationships(zip, chartPath)
  assert.ok(rels.every((rel) => rel['@_TargetMode'] !== 'External'), `${chartPath}: external reference`)
  const rel = rels.find((item) => item['@_Id'] === chart.externalData?.['@_id'])
  assert.ok(rel?.['@_Type'].endsWith('/package'), `${chartPath}: embedded workbook relationship`)
  const workbookPath = resolvePart(chartPath, rel['@_Target'])
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await zip.file(workbookPath).async('nodebuffer'))
  assert.equal(workbook.worksheets.length, 1)
  const sheet = workbook.worksheets[0]
  assert.equal(series.length, sheet.columnCount - 1, `${chartPath}: workbook series count`)
  if (expected) {
    assert.equal(series.length, expected.months.length, `${chartPath}: month count`)
    assert.equal(sheet.name, expected.sheetName)
    assert.equal(sheet.rowCount, expected.metrics.length + 1)
  }
  series.forEach((ser, index) => {
    const label = `${chartPath}, series ${index}`
    assert.equal(Number(ser.idx['@_val']), index, `${label}: idx`)
    assert.equal(Number(ser.order['@_val']), index, `${label}: order`)
    const month = sheet.getCell(1, index + 2).value
    const categories = Array.from({ length: sheet.rowCount - 1 }, (_, row) => sheet.getCell(row + 2, 1).value)
    const values = categories.map((_, row) => sheet.getCell(row + 2, index + 2).value)
    if (expected) {
      assert.equal(month, expected.months[index], `${label}: month`)
      assert.deepEqual(categories, expected.metrics.map((metric) => metric.label), `${label}: categories`)
      assert.deepEqual(values, expected.metrics.map((metric) => metric.values[index]), `${label}: workbook values`)
    }
    cacheValues(ser.tx.strRef.strCache, [month], `${label}: title`)
    cacheValues(ser.cat.strRef.strCache, categories, `${label}: categories`)
    cacheValues(ser.val.numRef.numCache, values, `${label}: numeric values`, true)
    const column = sheet.getColumn(index + 2).letter
    const name = /[ '!]/.test(sheet.name) ? `'${sheet.name.replace(/'/g, "''")}'` : sheet.name
    assert.equal(ser.tx.strRef.f, `${name}!$${column}$1`, `${label}: title formula`)
    assert.equal(ser.cat.strRef.f, `${name}!$A$2:$A$${sheet.rowCount}`, `${label}: category formula`)
    assert.equal(ser.val.numRef.f, `${name}!$${column}$2:$${column}$${sheet.rowCount}`, `${label}: value formula`)
  })
  for (const labels of descendants(chart.chart.plotArea, 'dLbls')) {
    for (const format of descendants(labels, 'numFmt')) {
      assert.equal(format['@_formatCode'], '0', `${chartPath}: integer labels`)
      assert.equal(format['@_sourceLinked'], '0', `${chartPath}: label format must survive Edit Data`)
    }
  }
}

async function verifyDeck(filename, expectations) {
  const zip = await JSZip.loadAsync(await fs.readFile(filename))
  const slides = await slidePaths(zip)
  assert.deepEqual(slides, await slidePaths(template), 'Slide order/count changed')
  const changedTextPages = new Set([1, 6, 7, 9, 11, 13, 15, 19, 20])
  for (const [index, slide] of slides.entries()) {
    if (!changedTextPages.has(index + 1)) assert.equal(await read(zip, slide), await read(template, slide), `Untouched page ${index + 1}`)
  }
  for (const name of Object.keys(template.files).filter((name) => /ppt\/(slideMasters|slideLayouts|theme)\/.+\.xml$/.test(name))) {
    assert.equal(await read(zip, name), await read(template, name), `${name}: template styling changed`)
  }
  let count = 0
  for (const page of [6, 8, 10, 12, 14, 16]) {
    const charts = await chartPaths(zip, slides[page - 1])
    assert.equal(charts.length, page === 6 ? 1 : 2, `Page ${page}: chart count`)
    for (const [index, chart] of charts.entries()) {
      await verifyChart(zip, chart, expectations?.get(page)?.[index])
      count++
    }
  }
  assert.equal(count, 11)
  const report = JSON.parse(await fs.readFile(filename.replace(/\.pptx$/, '-ppt-check.json'), 'utf8'))
  assert.equal(report.totals.warning, 0, 'PPT check warnings')
}

const keys = ['publication', 'submission', 'assignedManuscript', 'siSetUp', 'revenueWCHF', 'waiverRate', 'mpt']
const labels = ['Publication', 'Submission', 'Assigned Manuscript', 'SI Set Up', 'Revenue(WCHF)', 'Waiver Rate(%)', 'MPT']
function fixture(month) {
  const monthNumbers = month < 4 ? Array.from({ length: 4 }, (_, i) => month - 3 + i) : Array.from({ length: month }, (_, i) => i + 1)
  const months = monthNumbers.map((m) => `${m <= 0 ? 2025 : 2026}-${String(m <= 0 ? m + 12 : m).padStart(2, '0')}`)
  // Different fractions, zeros, negative values and >100% expose rounding and transposition.
  const value = (metric, series, journal = 0) => [0, 1.75, 175.25, -2.5, 38.125][(metric + series + journal) % 5] + (series % 2 ? journal : 0)
  const yearlySeries = months.map((label, index) => ({ label, ...Object.fromEntries(keys.map((key, metric) => [key, value(metric, index)])) }))
  const primaryLabels = ['Publ.', 'Sub.', 'Processing', 'MPT', 'TFD']
  const secondaryLabels = ['Revenue (WCHF)', 'Waiver Rate (%)', 'New SI', 'Rejection Rate (%)']
  const focusSeries = ['Foods', 'Nutrients', 'Children', 'Genes', 'BS'].map((sheetName, journal) => ({
    sheetName, displayName: sheetName === 'BS' ? 'Brain Sciences' : sheetName, months,
    primaryMonths: months, secondaryMonths: months,
    publication: [], submission: [], underProcessing: [], mpt: [], tfd: [],
    primaryMetrics: primaryLabels.map((label, metric) => ({ label, values: months.map((_, series) => value(metric, series, journal)) })),
    secondaryMetrics: secondaryLabels.map((label, metric) => ({ label, values: months.map((_, series) => value(metric + 5, series, journal)) }))
  }))
  const expectations = new Map([[6, [{ sheetName: '科室数据', months, metrics: labels.map((label, i) => ({ label, values: yearlySeries.map((row) => row[keys[i]]) })) }]]])
  focusSeries.forEach((series, index) => expectations.set(8 + index * 2, [
    { sheetName: series.displayName, months: series.primaryMonths, metrics: series.primaryMetrics },
    { sheetName: series.displayName, months: series.secondaryMonths, metrics: series.secondaryMetrics }
  ]))
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><text x="1" y="20">Regression</text></svg>'
  return { expectations, input: {
    templatePath, outputPath: path.join(output, `month-${month}.pptx`),
    detected: { reportKey: `2026${String(month).padStart(2, '0')}`, reportMonthLabel: months.at(-1), quarterPubHeader: '', quarterRevenueHeader: '' },
    reportMonthTitle: new Date(Date.UTC(2026, month - 1)).toLocaleDateString('en-US', { month: 'long', year: 'numeric', timeZone: 'UTC' }),
    officeRows: [], completionRows: [],
    departmentMetrics: { ...yearlySeries.at(-1), yearlySeries, publicationDelta: 0, publicationMom: 0, publicationYoy: null, revenueDelta: 0, revenueMom: 0, revenueYoy: null, submissionDelta: 0, submissionMom: 0 },
    snapshots: { officeSvg: svg, completionSvg: svg, ...Object.fromEntries(['staffPiCompletionSvgs', 'staffOwnerSvgs', 'staffSiSetupSvgs', 'staffSiSubSvgs', 'staffAePublSvgs'].map((key) => [key, [svg, svg]])) },
    staffPiCompletion: { currentCount: 1, previousCount: 0, delta: 1 }, focusSeries,
    spotlights: focusSeries.map((series) => ({ displayName: series.displayName, editors: [], publicationTcr: 1.75, revenueTcr: 0, waiverRate2026: 0.38125 }))
  } }
}

// Compare parsed package contents, not ZIP timestamps/compression. This includes values,
// styles, conditional formatting, merges, formulas, drawings and chart relationships.
async function workbookSnapshot(filename) {
  const zip = await JSZip.loadAsync(await fs.readFile(filename))
  const result = {}
  for (const name of Object.keys(zip.files).filter((name) => !zip.files[name].dir && !name.startsWith('docProps/')).sort()) {
    result[name] = /\.(xml|rels)$/.test(name) ? parse(await read(zip, name)) : await zip.file(name).async('base64')
  }
  return result
}

try {
  const pipelineInput = {
    paths: {
      mrWorkbook: path.join(root, 'docs/MR_202603-202603.xlsx'), monthlyTemplate: path.join(root, 'docs/monthly-data-generated-202603.xlsx'),
      staffTemplate: path.join(root, 'docs/staff-data-generated-202603.xlsx'), editorsJournals: path.join(root, 'docs/editors-journals.xlsx'),
      pptTemplate: templatePath, outputDir: path.join(output, 'real-march')
    },
    closeBehavior: 'tray', focusJournalOrder: ['Foods', 'Nutrients', 'Children', 'Genes', 'BS'], forceAeStaff: [],
    summaryOverrides: Object.fromEntries(['reportMonthLabel', ...keys].map((key) => [key, '']))
  }
  const result = await runPipeline({ sender: { send() {} } }, pipelineInput)
  if (process.argv.includes('--record-excel')) {
    console.log(`Excel baseline generated: ${path.dirname(result.outputs.monthlyWorkbook)}`)
  } else {
    const compareIndex = process.argv.indexOf('--compare-excel')
    if (compareIndex !== -1) {
      assert.ok(process.argv[compareIndex + 1], '--compare-excel requires a baseline directory')
      for (const filename of [result.outputs.monthlyWorkbook, result.outputs.staffWorkbook]) {
        const baseline = path.join(path.resolve(process.argv[compareIndex + 1]), path.basename(filename))
        assert.deepEqual(await workbookSnapshot(filename), await workbookSnapshot(baseline), `${path.basename(filename)}: Excel regression`)
      }
      console.log('PASS: both Excel outputs match the baseline package semantics')
    }
    await verifyDeck(result.outputs.presentation)
    console.log('PASS: real March MR, 11 charts and embedded workbooks')
    if (process.argv.includes('--invoice-only')) {
      const { verifyInvoiceSources } = await import('./invoice-regression.mjs')
      await verifyInvoiceSources({ runPipeline, verifyDeck, workbookSnapshot, input: pipelineInput, result, output })
    }
    for (const month of process.argv.includes('--invoice-only') ? [] : [1, 3, 4, 5, 12]) {
      const { input, expectations } = fixture(month)
      await buildPresentation(input)
      await verifyDeck(input.outputPath, expectations)
      console.log(`PASS: month ${month}, exact points/formulas/workbooks, unchanged template pages`)
    }
  }
  console.log(`Artifacts: ${output}`)
} catch (error) {
  console.error(`FAIL (artifacts retained): ${output}`)
  throw error
}

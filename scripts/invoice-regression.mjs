import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { XMLParser } from 'fast-xml-parser'

const currentHeader = '202601-202603 Invoiced(CHF)'
const parser = new XMLParser({ ignoreAttributes: false, removeNSPrefix: true, parseTagValue: false })
const list = (value) => value === undefined ? [] : Array.isArray(value) ? value : [value]
const value = (cell) => cell.value?.result ?? cell.value ?? null
const number = (cell) => Number(value(cell)) || 0
const headers = (sheet) => sheet.getRow(1).values
const load = async (filename) => {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filename)
  return workbook
}
const resolvePart = (owner, target) => path.posix.normalize(path.posix.join(path.posix.dirname(owner), target))

async function chartData(zip, page, index) {
  const owner = `ppt/slides/slide${page}.xml`
  const xml = await zip.file(owner).async('string')
  const ids = [...xml.matchAll(/<c:chart\b[^>]*r:id="([^"]+)"/g)].map((match) => match[1])
  const rels = parser.parse(await zip.file(`ppt/slides/_rels/slide${page}.xml.rels`).async('string'))
  const rel = list(rels.Relationships.Relationship).find((rel) => rel['@_Id'] === ids[index])
  const chartPath = resolvePart(owner, rel['@_Target'])
  const chart = parser.parse(await zip.file(chartPath).async('string')).chartSpace
  const relationships = parser.parse(await zip.file(path.posix.join(path.posix.dirname(chartPath), '_rels', `${path.posix.basename(chartPath)}.rels`)).async('string'))
  const workbookRel = list(relationships.Relationships.Relationship).find((rel) => rel['@_Id'] === chart.externalData['@_id'])
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await zip.file(resolvePart(chartPath, workbookRel['@_Target'])).async('nodebuffer'))
  return workbook.worksheets[0]
}

async function verifyAmounts(result, source) {
  const journals = source.getWorksheet('Journals')
  const sourceHeaders = headers(journals)
  const sourceRows = new Map()
  journals.eachRow((row, index) => {
    if (index > 1) sourceRows.set(String(row.getCell(sourceHeaders.indexOf('Journal')).value), row)
  })
  const monthly = await load(result.outputs.monthlyWorkbook)
  for (const name of ['Office', '完成率']) {
    const sheet = monthly.getWorksheet(name)
    const cols = headers(sheet)
    assert.ok(cols.includes(currentHeader), `${name}: invoice period header`)
    sheet.eachRow((row, index) => {
      if (index === 1) return
      const journal = String(row.getCell(cols.indexOf('Journal')).value)
      const sourceRow = sourceRows.get(journal)
      if (!sourceRow) return
      const included = name !== '完成率' || sourceRow.getCell(sourceHeaders.indexOf('Section')).value === 'Section Health'
      const amount = included ? number(sourceRow.getCell(sourceHeaders.indexOf(currentHeader))) : 0
      const target = included ? number(sourceRow.getCell(sourceHeaders.indexOf('Revenue Target Final'))) : 0
      assert.equal(number(row.getCell(cols.indexOf(currentHeader))), Math.round(amount), `${name}/${journal}: cumulative Invoice`)
      assert.equal(number(row.getCell(cols.indexOf('Revenue TCR'))), target ? amount / target : 0, `${name}/${journal}: Invoice TCR`)
      if (name === 'Office') {
        assert.equal(number(row.getCell(cols.indexOf('Revenue'))), Math.round(number(sourceRow.getCell(sourceHeaders.indexOf('Invoiced sent')))), `${journal}: monthly Invoice`)
      }
    })
  }
  let total = 0
  journals.eachRow((row, index) => {
    if (index > 1 && row.getCell(sourceHeaders.indexOf('Section')).value === 'Section Health') total += number(row.getCell(sourceHeaders.indexOf('Invoiced sent')))
  })
  const department = monthly.getWorksheet('科室数据')
  assert.equal(number(department.getCell('F4')), Number((total / 10000).toFixed(2)), 'Department Invoice in WCHF')
  const zip = await JSZip.loadAsync(await fs.readFile(result.outputs.presentation))
  const departmentChart = await chartData(zip, 6, 0)
  const march = headers(departmentChart).indexOf('Mar.')
  assert.equal(number(departmentChart.getCell(6, march)), number(department.getCell('F4')), 'PPT department Invoice')
  const departmentText = await zip.file('ppt/slides/slide6.xml').async('string')
  assert.ok(departmentText.includes(`${number(department.getCell('F4')).toFixed(2)} WCHF`), 'PPT department amount text')
  assert.ok(departmentText.includes('（比上月--），MoM --，YoY --'), 'Unconfirmed history must not be used for Invoice comparisons')

  for (const [index, name] of ['Foods', 'Nutrients', 'Children', 'Genes', 'BS'].entries()) {
    const journal = name === 'BS' ? 'Brain Sciences' : name
    const row = sourceRows.get(journal)
    const amount = Math.round(number(row.getCell(sourceHeaders.indexOf('Invoiced sent'))) / 10000)
    assert.equal(number(monthly.getWorksheet(name).getCell('E9')), amount, `${name}: focus Invoice`)
    const data = await chartData(zip, 8 + index * 2, 1)
    const column = headers(data).indexOf('Mar.')
    let revenueRow = 0
    data.eachRow((r, i) => { if (r.getCell(1).value === 'Revenue(WCHF)') revenueRow = i })
    assert.ok(revenueRow, 'Focus chart Invoice metric')
    assert.equal(number(data.getCell(revenueRow, column)), amount, `${name}: PPT focus Invoice`)
    const metricXml = await zip.file(`ppt/slides/slide${7 + index * 2}.xml`).async('string')
    const percentages = [...metricXml.matchAll(/<a:t>(\d+(?:\.\d+)?%)<\/a:t>/g)].map((match) => match[1])
    const target = number(row.getCell(sourceHeaders.indexOf('Revenue Target Final')))
    const rate = target ? number(row.getCell(sourceHeaders.indexOf(currentHeader))) / target : 0
    assert.equal(percentages[1], `${(rate * 100).toFixed(2)}%`, `${name}: PPT Invoice TCR`)
  }
  return monthly
}

function verifyUnrelatedCells(before, after) {
  assert.deepEqual(after.worksheets.map((sheet) => sheet.name), before.worksheets.map((sheet) => sheet.name))
  for (const sheet of before.worksheets) {
    const next = after.getWorksheet(sheet.name)
    // Some templates format all 16,384 columns; do not rebuild headers per cell.
    const sheetHeaders = headers(sheet)
    assert.deepEqual(next.model.merges, sheet.model.merges, `${sheet.name}: merges`)
    assert.deepEqual(next.conditionalFormattings, sheet.conditionalFormattings, `${sheet.name}: conditional formats`)
    sheet.eachRow({ includeEmpty: true }, (row, r) => {
      assert.equal(next.getRow(r).height, row.height, `${sheet.name}: row height`)
      row.eachCell({ includeEmpty: true }, (cell, c) => {
        const title = sheetHeaders[c] ?? ''
        const changed = (sheet.name === 'Office' && /Revenue|Invoiced/.test(title))
          || (sheet.name === '完成率' && [7, 8, 9].includes(c))
          || (sheet.name === '科室数据' && r === 4 && c === 6)
          || (['Foods', 'Nutrients', 'Children', 'Genes', 'BS'].includes(sheet.name) && r === 9 && c === 5)
        // TCR-dependent fills may legitimately change; layout/number formats may not.
        const beforeStyle = { ...cell.style }
        const afterStyle = { ...next.getCell(r, c).style }
        if (changed) { delete beforeStyle.fill; delete afterStyle.fill }
        assert.deepEqual(afterStyle, beforeStyle, `${sheet.name}/${cell.address}: style`)
        if (!changed) assert.deepEqual(next.getCell(r, c).value, cell.value, `${sheet.name}/${cell.address}: unrelated data`)
      })
    })
  }
}

export async function verifyInvoiceSources({ runPipeline, verifyDeck, workbookSnapshot, input, result, output }) {
  const original = await load(input.paths.mrWorkbook)
  const before = await verifyAmounts(result, original)
  const fixture = await load(input.paths.mrWorkbook)
  const sheet = fixture.getWorksheet('Journals')
  const cols = headers(sheet)
  sheet.eachRow((row, index) => {
    if (index === 1) return
    // Deliberately different source amounts expose any residual Revenue reads.
    row.getCell(cols.indexOf('Invoiced sent')).value = index % 2 ? 0 : 123456.25
    row.getCell(cols.indexOf(currentHeader)).value = index % 2 ? 0 : 3500
    row.getCell(cols.indexOf('Revenue Target Final')).value = index % 3 ? 2000 : 0
    for (const name of ['Revenue', 'Revenue/LastM', 'Revenue/YoY', '202601-202603 Revenue']) row.getCell(cols.indexOf(name)).value = 987654321
  })
  // Put last year's Invoice column first: choosing the first match is incorrect.
  const current = cols.indexOf(currentHeader)
  const previous = cols.indexOf('202501-202503 Invoiced(CHF)')
  sheet.eachRow((row) => {
    const old = row.getCell(current).value
    row.getCell(current).value = row.getCell(previous).value
    row.getCell(previous).value = old
  })
  const fixturePath = path.join(output, 'MR_202603-202603.xlsx')
  await fixture.xlsx.writeFile(fixturePath)
  const nextInput = { ...input, paths: { ...input.paths, mrWorkbook: fixturePath, outputDir: path.join(output, 'invoice-fixture') } }
  const next = await runPipeline({ sender: { send() {} } }, nextInput)
  const after = await verifyAmounts(next, fixture)
  await verifyDeck(next.outputs.presentation)
  verifyUnrelatedCells(before, after)
  assert.deepEqual(await workbookSnapshot(next.outputs.staffWorkbook), await workbookSnapshot(result.outputs.staffWorkbook), 'Staff workbook unchanged')
  console.log('PASS: Invoice amounts/TCR in Excel and PPT, zero targets, >100%, header order; staff and unrelated cells unchanged')

  sheet.getCell(1, cols.indexOf('Invoiced sent')).value = 'Missing Invoice'
  await fixture.xlsx.writeFile(fixturePath)
  await assert.rejects(runPipeline({ sender: { send() {} } }, { ...nextInput, paths: { ...nextInput.paths, outputDir: path.join(output, 'missing-invoice') } }), /Invoiced sent/)
  console.log('PASS: missing Invoice source fails explicitly instead of falling back to Revenue')
}

import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import ExcelJS from 'exceljs'
import { build } from 'vite'

// Synthetic fixtures only: test the production writers without private input files.
const root = fileURLToPath(new URL('../', import.meta.url))
await fs.mkdir(path.join(root, 'out'), { recursive: true })
const output = await fs.mkdtemp(path.join(root, 'out', 'monthly-layout-'))
const entry = path.join(root, 'src/main/services/pipeline.ts')
await build({
  configFile: false, root, logLevel: 'error',
  plugins: [{
    name: 'expose-monthly-writers-for-regression',
    transform(code, id) {
      if (path.resolve(id) === entry) return `${code}\nexport { populateOfficeSheet, populateCompletionSheet };`
    }
  }],
  build: {
    ssr: entry, outDir: output, emptyOutDir: false,
    rollupOptions: { output: { format: 'cjs', entryFileNames: 'writers.cjs' } }
  }
})
const { populateOfficeSheet, populateCompletionSheet } = createRequire(import.meta.url)(path.join(output, 'writers.cjs'))
const context = { quarterPubHeader: '202601-202609 Publ.', quarterRevenueHeader: '202601-202609 Invoiced(CHF)', timeProgress: 0.75 }
const officeRows = [137, 0, 240].map((sub, index) => ({
  journal: `Example ${index + 1}`, contrib: 123, yearlyPublTarget: 400,
  monthlyPublTarget: 30, publ: 25, publLast: 20, mom: 0.25, pubYoy: 10, yoy: 1.5,
  newSi: 3, avePubSis: 2.5, cnRate: 0.4, mpt: 40, sub, susyCfp: 15,
  mmailerRegularCfp: 20, mmailerFCfp: 10, sendCfp: 45, wrTarget2026: 0.35,
  waiverRate: 0.3, waiverRate2026: 0.32, publicationTcr: 0.6, revenueTcr: 0.5,
  quarterPub: 240, revenue: 2500, quarterRevenue: 20000, revenueTargetFinal: 40000
}))
const completionRows = officeRows.map((row, index) => ({ ...row, group: index < 2 ? 'Group A' : 'Group B' }))

async function fixture(affected) {
  const workbook = new ExcelJS.Workbook()
  for (const name of ['Office', '完成率']) {
    const sheet = workbook.addWorksheet(name)
    for (let col = 1; col <= (name === 'Office' ? 28 : 12); col++) {
      sheet.getColumn(col).width = 18
      sheet.getCell(1, col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF404040' } }
      sheet.getCell(1, col).font = { color: { argb: 'FFFFFFFF' }, bold: true }
      for (let row = 2; row <= 4; row++) {
        sheet.getCell(row, col).numFmt = 'General'
        sheet.getCell(row, col).font = { name: 'Calibri', size: 11 }
      }
    }
    sheet.views = [{ state: 'frozen', xSplit: 2, ySplit: 1 }]
  }
  const office = workbook.getWorksheet('Office')
  office.getCell('N2').value = 137
  if (affected) {
    // Both cell and inherited column formats can carry forward from a prior file.
    office.getColumn(14).numFmt = '0.00%'
    office.getCell('N3').numFmt = '0%'
    for (let col = 3; col <= 12; col++) workbook.getWorksheet('完成率').getColumn(col).hidden = true
  }
  office.getCell('N1').numFmt = 'General'
  const completion = workbook.getWorksheet('完成率')
  completion.getColumn(13).hidden = true // Unrelated helper columns must retain their layout.
  completion.getCell('M2').value = 'Preserve helper'
  workbook.addWorksheet('Unrelated').getCell('A1').value = 'Preserve other sheets'
  // Exercise ExcelJS's real save/read behavior before processing the template.
  const restored = new ExcelJS.Workbook()
  await restored.xlsx.load(await workbook.xlsx.writeBuffer())
  return restored
}

async function generate(workbook) {
  populateOfficeSheet(workbook.getWorksheet('Office'), officeRows, context)
  populateCompletionSheet(workbook.getWorksheet('完成率'), completionRows, context)
  const restored = new ExcelJS.Workbook()
  await restored.xlsx.load(await workbook.xlsx.writeBuffer())
  return restored
}

function verify(workbook) {
  const office = workbook.getWorksheet('Office')
  const completion = workbook.getWorksheet('完成率')
  const problems = []
  officeRows.forEach((row, index) => {
    const cell = office.getCell(index + 2, 14)
    assert.equal(cell.value, row.sub, 'Submission counts must not be scaled')
    if (cell.numFmt !== '0') problems.push(`${cell.address}: Sub. format is ${cell.numFmt}`)
    for (const col of [7, 9, 12, 19, 20, 21, 22, 23]) assert.equal(office.getCell(index + 2, col).numFmt, '0.00%')
  })
  for (let col = 1; col <= 12; col++) {
    if (completion.getColumn(col).hidden) problems.push(`Completion column ${col} is hidden`)
    assert.equal(completion.getColumn(col).width, 18, 'Preserve column widths')
  }
  assert.deepEqual(problems, [], 'Template presentation state must not corrupt generated output')
  completionRows.forEach((row, index) => {
    const top = index + 2, bottom = completionRows.length + 5 + index
    for (const r of [top, bottom]) {
      assert.equal(completion.getCell(r, 2).value, row.journal)
      assert.equal(completion.getCell(r, 3).value, row.contrib)
      assert.equal(completion.getCell(r, 4).value, row.quarterPub)
      assert.equal(completion.getCell(r, 7).value, row.quarterRevenue)
      assert.equal(completion.getCell(r, 8).value, row.revenueTargetFinal)
      for (const c of [6, 9, 10, 11, 12]) assert.equal(completion.getCell(r, c).numFmt, '0.00%')
    }
    assert.deepEqual(completion.getCell(top, 6).value, { formula: `D${top}/E${top}`, result: row.publicationTcr })
    assert.deepEqual(completion.getCell(top, 9).value, { formula: `G${top}/H${top}`, result: row.revenueTcr })
    assert.equal(completion.getCell(bottom, 6).value, row.publicationTcr)
    assert.equal(completion.getCell(bottom, 9).value, row.revenueTcr)
  })
  assert.ok(completion.getCell('A3').isMerged, 'Preserve category groups')
  assert.equal(completion.getColumn(13).hidden, true)
  assert.equal(completion.getCell('M2').value, 'Preserve helper')
  assert.equal(workbook.getWorksheet('Unrelated').getCell('A1').value, 'Preserve other sheets')
  assert.equal(completion.views[0].state, 'frozen')
}

const affected = await generate(await fixture(true))
verify(affected)
const clean = await generate(await fixture(false))
verify(clean)
for (const name of ['Office', '完成率']) {
  const left = affected.getWorksheet(name), right = clean.getWorksheet(name)
  left.eachRow((row, r) => row.eachCell((cell, c) => {
    assert.deepEqual(cell.value, right.getCell(r, c).value, `${name}/${cell.address}: values depend on template formatting`)
    assert.deepEqual(cell.style, right.getCell(r, c).style, `${name}/${cell.address}: unexpected style difference`)
  }))
  assert.deepEqual(left.model.merges, right.model.merges)
  assert.deepEqual(left.conditionalFormattings, right.conditionalFormattings)
}
verify(await generate(affected)) // Reusing generated output as next month's template remains safe.
await affected.xlsx.writeFile(path.join(output, 'verified-monthly.xlsx'))
console.log('PASS: submission counts/formats, visible completion tables, formulas, rates, groups, helper columns, styles and repeat generation')
console.log(`Synthetic artifacts: ${output}`)

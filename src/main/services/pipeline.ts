import fs from 'node:fs/promises'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { XMLBuilder, XMLParser } from 'fast-xml-parser'
import type { IpcMainInvokeEvent } from 'electron'
import type { PipelineInput, PipelineProgressEvent, PipelineResult } from '../../shared/contracts'
import { buildPresentation } from './ppt'
import { interpolateColor, renderTableSvg, renderWorksheetRangeSvg } from './svg'
import {
  buildOutputPath,
  columnLetter,
  compactName,
  deepClone,
  detectReportKey,
  ensureArrayLength,
  monthLabelFromKey,
  normalizeNumber,
  normalizePersonName,
  normalizeText,
  percentageDelta,
  titleMonthFromKey,
  toInteger
} from './utils'

type GenericRow = Record<string, unknown>

type OfficeRow = {
  journal: string
  contrib: number
  yearlyPublTarget: number
  monthlyPublTarget: number
  publ: number
  publLast: number
  mom: number
  pubYoy: number
  yoy: number
  newSi: number
  avePubSis: number
  cnRate: number
  mpt: number
  sub: number
  susyCfp: number
  mmailerRegularCfp: number
  mmailerFCfp: number
  sendCfp: number
  wrTarget2026: number
  waiverRate: number
  waiverRate2026: number
  publicationTcr: number
  revenueTcr: number
  quarterPub: number
  revenue: number
  quarterRevenue: number
  revenueTargetFinal: number
}

type CompletionRow = {
  group: string
  journal: string
  contrib: number
  quarterPub: number
  yearlyPublTarget: number
  publicationTcr: number
  quarterRevenue: number
  revenueTargetFinal: number
  revenueTcr: number
  waiverRate: number
  wrTarget2026: number
  waiverRate2026: number
}

function completionJournalKey(journal: string): string {
  return compactName(journal).toLowerCase()
}

const COMPLETION_JOURNAL_GROUPS: Array<{ group: string; journals: string[] }> = [
  { group: 'SCIE', journals: ['Foods', 'Nutrients', 'Children', 'Genes', 'Brain Sciences'] },
  {
    group: 'ESCI',
    journals: [
      'Antibodies',
      'Beverages',
      'Cardiogenetics',
      'Diabetology',
      'Dietetics',
      'DNA',
      'Endocrines',
      'JVD',
      'Livers',
      'Sports',
      'Surgical Techniques Development'
    ]
  },
  { group: 'Scopus', journals: ['Allergies', 'JMAHP'] },
  { group: 'Others', journals: ['JDAD', 'Sustainable Foods', 'Food Engineering'] }
]
const COMPLETION_GROUP_BY_JOURNAL = new Map(
  COMPLETION_JOURNAL_GROUPS.flatMap(({ group, journals }) =>
    journals.map((journal) => [completionJournalKey(journal), group] as const)
  )
)

type DepartmentMetrics = {
  reportMonthCellLabel: string
  reportMonthTitle: string
  publication: number
  submission: number
  assignedManuscript: number
  siSetUp: number
  revenueWCHF: number
  waiverRate: number
  mpt: number
  publicationDelta: number
  publicationMom: number
  publicationYoy: number | null
  revenueDelta: number
  revenueMom: number
  revenueYoy: number | null
  submissionDelta: number
  submissionMom: number
  yearlySeries: Array<{
    label: string
    publication: number
    submission: number
    assignedManuscript: number
    siSetUp: number
    revenueWCHF: number
    waiverRate: number
    mpt: number
  }>
}

type PptSnapshots = {
  officeSvg: string
  completionSvg: string
  staffPiCompletionSvgs: string[]
  staffOwnerSvgs: string[]
  staffSiSetupSvgs: string[]
  staffSiSubSvgs: string[]
  staffAePublSvgs: string[]
}

type StaffPiCompletionSummary = {
  currentCount: number
  previousCount: number
  delta: number
}

type SnapshotOverlay = {
  background?: string
  dataBarRatio?: number
  dataBarColor?: string
}

type FocusJournalSeries = {
  sheetName: string
  displayName: string
  months: string[]
  publication: number[]
  submission: number[]
  underProcessing: number[]
  mpt: number[]
  tfd: number[]
  primaryMetrics: Array<{ label: string; values: number[] }>
  secondaryMetrics: Array<{ label: string; values: number[] }>
}

type ReportContext = {
  reportKey: string
  reportMonthLabel: string
  reportMonthTitle: string
  comparisonMonthLabels: string[]
  quarterPubHeader: string
  quarterRevenueHeader: string
  quarterPreviousPubHeader?: string
  quarterPreviousRevenueHeader?: string
  timeProgress: number
}

type JournalSpotlight = {
  displayName: string
  editors: string[]
  publicationTcr: number
  revenueTcr: number
  waiverRate2026: number
}

function emitProgress(event: IpcMainInvokeEvent, step: string, message: string, level: PipelineProgressEvent['level'] = 'info'): void {
  event.sender.send('pipeline:progress', { step, message, level } satisfies PipelineProgressEvent)
}

function getWorksheetOrThrow(workbook: ExcelJS.Workbook, name: string): ExcelJS.Worksheet {
  const worksheet = workbook.getWorksheet(name)
  if (!worksheet) {
    throw new Error(`Missing worksheet: ${name}`)
  }
  return worksheet
}

function rowToObject(headers: string[], row: ExcelJS.Row): GenericRow {
  const record: GenericRow = {}
  headers.forEach((header, index) => {
    if (!header) {
      return
    }

    const value = row.getCell(index + 1).value
    record[header] = value
  })
  return record
}

function sheetToObjects(worksheet: ExcelJS.Worksheet): GenericRow[] {
  const headerRow = worksheet.getRow(1)
  const headers = Array.from({ length: headerRow.cellCount }, (_, index) => normalizeText(headerRow.getCell(index + 1).value))
  const rows: GenericRow[] = []
  for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex += 1) {
    const row = worksheet.getRow(rowIndex)
    const firstCell = normalizeText(row.getCell(1).value)
    if (!firstCell && row.actualCellCount === 0) {
      continue
    }
    rows.push(rowToObject(headers, row))
  }
  return rows
}

function copyStyle(source: ExcelJS.Cell, target: ExcelJS.Cell): void {
  target.style = deepClone(source.style)
  if (source.numFmt) {
    target.numFmt = source.numFmt
  }
}

function fillCell(target: ExcelJS.Cell, value: ExcelJS.CellValue, sourceStyle?: ExcelJS.Cell): void {
  target.value = value
  if (sourceStyle) {
    copyStyle(sourceStyle, target)
  }
}

function leftAlignHeaderRow(row: ExcelJS.Row, endCol: number): void {
  for (let col = 1; col <= endCol; col += 1) {
    const cell = row.getCell(col)
    cell.alignment = {
      ...(cell.alignment ?? {}),
      horizontal: 'left'
    }
  }
}

function normalizeRate(value: unknown): number {
  if (typeof value === 'object' && value !== null && 'result' in (value as Record<string, unknown>)) {
    return normalizeRate((value as { result: unknown }).result)
  }
  if (typeof value === 'string' && value.trim().endsWith('%')) {
    return normalizeNumber(value)
  }
  const numericValue = normalizeNumber(value)
  return numericValue > 1 ? numericValue / 100 : numericValue
}

function normalizeStaffPiCompletionRate(value: unknown): number {
  if (typeof value === 'object' && value !== null && 'result' in (value as Record<string, unknown>)) {
    return normalizeStaffPiCompletionRate((value as { result: unknown }).result)
  }
  return normalizeNumber(value)
}

function solidFill(argb: string): ExcelJS.Fill {
  const color = argb.length === 6 ? `FF${argb}` : argb
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: color }, bgColor: { argb: color } }
}

function addThreeColorScale(sheet: ExcelJS.Worksheet, column: string, startRow: number, endRow: number, priority: number): void {
  if (endRow < startRow) {
    return
  }
  sheet.addConditionalFormatting({
    ref: `${column}${startRow}:${column}${endRow}`,
    rules: [{
      type: 'colorScale',
      priority,
      cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
      color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }]
    }]
  })
}

function addDataBar(sheet: ExcelJS.Worksheet, column: string, startRow: number, endRow: number, priority: number): void {
  if (endRow < startRow) {
    return
  }
  sheet.addConditionalFormatting({
    ref: `${column}${startRow}:${column}${endRow}`,
    rules: [{
      type: 'dataBar',
      priority,
      gradient: true,
      minLength: 0,
      maxLength: 100,
      cfvo: [{ type: 'min' }, { type: 'max' }],
      color: { argb: 'FF638EC6' }
    } as ExcelJS.ConditionalFormattingRule]
  })
}

function addCellFillRule(
  sheet: ExcelJS.Worksheet,
  column: string,
  startRow: number,
  endRow: number,
  operator: 'greaterThan' | 'lessThan' | 'equal' | 'between',
  formula: string,
  color: string,
  priority: number
): void {
  if (endRow < startRow) {
    return
  }
  sheet.addConditionalFormatting({
    ref: `${column}${startRow}:${column}${endRow}`,
    rules: [{
      type: 'cellIs',
      operator,
      priority,
      formulae: [formula],
      style: { fill: solidFill(color) }
    }]
  })
}

function addExpressionFillRule(
  sheet: ExcelJS.Worksheet,
  column: string,
  startRow: number,
  endRow: number,
  formula: string,
  color: string,
  priority: number
): void {
  if (endRow < startRow) {
    return
  }
  sheet.addConditionalFormatting({
    ref: `${column}${startRow}:${column}${endRow}`,
    rules: [{
      type: 'expression',
      priority,
      formulae: [formula],
      style: { fill: solidFill(color) }
    }]
  })
}

function resetWorksheet(workbook: ExcelJS.Workbook, name: string): ExcelJS.Worksheet {
  const existing = workbook.getWorksheet(name)
  if (existing) {
    workbook.removeWorksheet(existing.id)
  }
  return workbook.addWorksheet(name)
}

function arrayOf<T>(value: T | T[] | undefined): T[] {
  if (!value) {
    return []
  }
  return Array.isArray(value) ? value : [value]
}

async function zipText(zip: JSZip, pathName: string): Promise<string> {
  const file = zip.file(pathName)
  return file ? file.async('text') : ''
}

async function workbookSheetPaths(zip: JSZip): Promise<Map<string, string>> {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })
  const workbookXml = parser.parse(await zipText(zip, 'xl/workbook.xml'))
  const relsXml = parser.parse(await zipText(zip, 'xl/_rels/workbook.xml.rels'))
  const relTargetById = new Map<string, string>()

  for (const rel of arrayOf<Record<string, string>>(relsXml.Relationships?.Relationship)) {
    if (rel['@_Id'] && rel['@_Target']) {
      relTargetById.set(rel['@_Id'], rel['@_Target'])
    }
  }

  const result = new Map<string, string>()
  for (const sheet of arrayOf<Record<string, string>>(workbookXml.workbook?.sheets?.sheet)) {
    const sheetName = sheet['@_name']
    const relId = sheet['@_r:id']
    const target = relTargetById.get(relId)
    if (sheetName && target) {
      result.set(sheetName, target.startsWith('xl/') ? target : `xl/${target}`)
    }
  }
  return result
}

function nextRelationshipId(relationships: Array<Record<string, string>>): string {
  const used = new Set(relationships.map((rel) => rel['@_Id']))
  for (let index = relationships.length + 1; index < relationships.length + 100; index += 1) {
    const id = `rId${index}`
    if (!used.has(id)) {
      return id
    }
  }
  return `rId${Date.now()}`
}

async function mergeContentTypeOverrides(sourceZip: JSZip, outputZip: JSZip): Promise<void> {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })
  const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', format: false })
  const sourceTypes = parser.parse(await zipText(sourceZip, '[Content_Types].xml'))
  const outputTypes = parser.parse(await zipText(outputZip, '[Content_Types].xml'))
  const outputOverrides = arrayOf<Record<string, string>>(outputTypes.Types?.Override)
  const existingParts = new Set(outputOverrides.map((override) => override['@_PartName']))

  for (const override of arrayOf<Record<string, string>>(sourceTypes.Types?.Override)) {
    const partName = override['@_PartName']
    if (partName && (partName.startsWith('/xl/charts/') || partName.startsWith('/xl/drawings/')) && !existingParts.has(partName)) {
      outputOverrides.push(override)
      existingParts.add(partName)
    }
  }
  outputTypes.Types.Override = outputOverrides
  outputZip.file('[Content_Types].xml', builder.build(outputTypes))
}

async function preserveTemplateCharts(sourcePath: string, outputPath: string, sheetNames: string[]): Promise<void> {
  const [sourceBuffer, outputBuffer] = await Promise.all([fs.readFile(sourcePath), fs.readFile(outputPath)])
  const [sourceZip, outputZip] = await Promise.all([JSZip.loadAsync(sourceBuffer), JSZip.loadAsync(outputBuffer)])
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })
  const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', format: false })
  const sourceSheets = await workbookSheetPaths(sourceZip)
  const outputSheets = await workbookSheetPaths(outputZip)

  for (const file of Object.values(sourceZip.files)) {
    if (!file.dir && (file.name.startsWith('xl/charts/') || file.name.startsWith('xl/drawings/'))) {
      outputZip.file(file.name, await file.async('nodebuffer'))
    }
  }
  await mergeContentTypeOverrides(sourceZip, outputZip)

  for (const sheetName of sheetNames) {
    const sourceSheetPath = sourceSheets.get(sheetName)
    const outputSheetPath = outputSheets.get(sheetName)
    if (!sourceSheetPath || !outputSheetPath) {
      continue
    }

    const sourceSheetXml = await zipText(sourceZip, sourceSheetPath)
    const drawingMatch = sourceSheetXml.match(/<drawing\b[^>]*r:id="([^"]+)"[^>]*\/>/)
    if (!drawingMatch) {
      continue
    }

    const sourceRelsPath = sourceSheetPath.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
    const sourceRelsXml = parser.parse(await zipText(sourceZip, sourceRelsPath))
    const sourceDrawingRel = arrayOf<Record<string, string>>(sourceRelsXml.Relationships?.Relationship)
      .find((rel) => rel['@_Id'] === drawingMatch[1])
    if (!sourceDrawingRel) {
      continue
    }

    const outputRelsPath = outputSheetPath.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
    const outputRelsText = await zipText(outputZip, outputRelsPath)
    const outputRelsXml = outputRelsText
      ? parser.parse(outputRelsText)
      : { Relationships: { '@_xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships', Relationship: [] } }
    const outputRelationships = arrayOf<Record<string, string>>(outputRelsXml.Relationships?.Relationship)
      .filter((rel) => rel['@_Type'] !== sourceDrawingRel['@_Type'])
    const drawingRelId = nextRelationshipId(outputRelationships)

    outputRelationships.push({
      ...sourceDrawingRel,
      '@_Id': drawingRelId
    })
    outputRelsXml.Relationships.Relationship = outputRelationships
    outputZip.file(outputRelsPath, builder.build(outputRelsXml))

    const drawingTag = drawingMatch[0].replace(drawingMatch[1], drawingRelId)
    const outputSheetXml = await zipText(outputZip, outputSheetPath)
    const xmlWithNamespace = outputSheetXml.includes('xmlns:r=')
      ? outputSheetXml
      : outputSheetXml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ')
    const xmlWithoutDrawing = xmlWithNamespace.replace(/<drawing\b[^>]*\/>/g, '')
    const anchorMatch = xmlWithoutDrawing.match(/<(pageMargins|legacyDrawing|drawingHF|picture|tableParts|extLst)\b/)
    const xmlWithDrawing = anchorMatch?.index !== undefined
      ? `${xmlWithoutDrawing.slice(0, anchorMatch.index)}${drawingTag}${xmlWithoutDrawing.slice(anchorMatch.index)}`
      : xmlWithoutDrawing.replace('</worksheet>', `${drawingTag}</worksheet>`)
    outputZip.file(outputSheetPath, xmlWithDrawing)
  }

  await fs.writeFile(outputPath, await outputZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

function parseReportContext(mrWorkbookPath: string, journalRows: GenericRow[]): ReportContext {
  const sample = journalRows[0] ?? {}
  const quarterPubHeader = Object.keys(sample).find((key) => /\d{6}-\d{6} Pub$/.test(key)) ?? '202601-202603 Pub'
  const quarterRevenueHeader = Object.keys(sample).find((key) => /\d{6}-\d{6} Revenue$/.test(key)) ?? '202601-202603 Revenue'
  const quarterPreviousPubHeader = Object.keys(sample).find((key) => /\d{6}-\d{6} Pub$/.test(key) && key !== quarterPubHeader)
  const quarterPreviousRevenueHeader = Object.keys(sample).find((key) => /\d{6}-\d{6} Revenue$/.test(key) && key !== quarterRevenueHeader)
  const reportKey = detectReportKey(mrWorkbookPath, quarterPubHeader)
  const reportMonthLabel = monthLabelFromKey(reportKey)
  const reportMonthTitle = titleMonthFromKey(reportKey)
  const monthNumber = Number.parseInt(reportKey.slice(4, 6), 10) || 1

  return {
    reportKey,
    reportMonthLabel,
    reportMonthTitle,
    comparisonMonthLabels: comparisonMonthLabels(reportKey),
    quarterPubHeader,
    quarterRevenueHeader,
    quarterPreviousPubHeader,
    quarterPreviousRevenueHeader,
    timeProgress: monthNumber / 12
  }
}

function buildOfficeRows(journalRows: GenericRow[], context: ReportContext): OfficeRow[] {
  return journalRows
    .filter((row) => ['SCI/SCOPUS', 'EI/SCI/SCOPUS'].includes(normalizeText(row.Index).replace(/\s+/g, '').toUpperCase()))
    .map((row) => {
      const yearlyPublTarget = normalizeNumber(row['Yearly Publ Target'])
      const revenueTargetFinal = normalizeNumber(row['Revenue Target Final'])
      const quarterPub = normalizeNumber(row[context.quarterPubHeader])
      const quarterRevenue = normalizeNumber(row[context.quarterRevenueHeader])
      return {
        journal: normalizeText(row.Journal),
        contrib: toInteger(normalizeNumber(row['Contrib.'])),
        yearlyPublTarget: toInteger(yearlyPublTarget),
        monthlyPublTarget: toInteger(normalizeNumber(row['Monthly Publ Target'])),
        publ: toInteger(normalizeNumber(row['Publ.'])),
        publLast: toInteger(normalizeNumber(row['Pub/lastM'])),
        mom: normalizeNumber(row['Delta01 (%)']),
        pubYoy: toInteger(normalizeNumber(row['Pub/YOY'])),
        yoy: normalizeNumber(row['Delta02 (%)']),
        newSi: toInteger(normalizeNumber(row['New SI'])),
        avePubSis: normalizeNumber(row['Ave. Pub SIs']),
        cnRate: normalizeNumber(row['CN%']),
        mpt: normalizeNumber(row.MPT),
        sub: toInteger(normalizeNumber(row['Sub.'])),
        susyCfp: toInteger(normalizeNumber(row['Susy CFP'])),
        mmailerRegularCfp: toInteger(normalizeNumber(row['Mmailer-Regular CFP'])),
        mmailerFCfp: toInteger(normalizeNumber(row['Mmailer-FCFP'])),
        sendCfp: toInteger(normalizeNumber(row['Send CFP (CFP+Reminder)']) || normalizeNumber(row.CFP)),
        wrTarget2026: normalizeRate(row['Waiver Rate Target2026']),
        waiverRate: normalizeRate(row['Waiver Rate']),
        waiverRate2026: normalizeRate(row['Waiver Rate2026']),
        publicationTcr: yearlyPublTarget ? quarterPub / yearlyPublTarget : 0,
        revenueTcr: revenueTargetFinal ? quarterRevenue / revenueTargetFinal : 0,
        quarterPub: toInteger(quarterPub),
        revenue: toInteger(normalizeNumber(row.Revenue)),
        quarterRevenue: toInteger(quarterRevenue),
        revenueTargetFinal: toInteger(revenueTargetFinal)
      }
    })
    .sort((left, right) => right.publ - left.publ)
}

function buildCompletionRows(journalRows: GenericRow[], context: ReportContext): CompletionRow[] {
  const sourceRowsByJournal = new Map<string, GenericRow>()
  for (const row of journalRows) {
    if (normalizeText(row.Section) !== 'Section Health') {
      continue
    }
    const key = completionJournalKey(normalizeText(row.Journal))
    if (COMPLETION_GROUP_BY_JOURNAL.has(key) && !sourceRowsByJournal.has(key)) {
      sourceRowsByJournal.set(key, row)
    }
  }

  return COMPLETION_JOURNAL_GROUPS.flatMap(({ group, journals }) =>
    journals.map((journal) => {
      const sourceRow = sourceRowsByJournal.get(completionJournalKey(journal))
      const quarterPub = normalizeNumber(sourceRow?.[context.quarterPubHeader])
      const yearlyPublTarget = normalizeNumber(sourceRow?.['Yearly Publ Target'])
      const quarterRevenue = normalizeNumber(sourceRow?.[context.quarterRevenueHeader])
      const revenueTargetFinal = normalizeNumber(sourceRow?.['Revenue Target Final'])
      return {
        group,
        journal,
        contrib: toInteger(normalizeNumber(sourceRow?.['Contrib.'])),
        quarterPub: toInteger(quarterPub),
        yearlyPublTarget: toInteger(yearlyPublTarget),
        publicationTcr: yearlyPublTarget ? quarterPub / yearlyPublTarget : 0,
        quarterRevenue: toInteger(quarterRevenue),
        revenueTargetFinal: toInteger(revenueTargetFinal),
        revenueTcr: revenueTargetFinal ? quarterRevenue / revenueTargetFinal : 0,
        waiverRate: normalizeRate(sourceRow?.['Waiver Rate']),
        wrTarget2026: normalizeRate(sourceRow?.['Waiver Rate Target2026']),
        waiverRate2026: normalizeRate(sourceRow?.['Waiver Rate2026'])
      }
    })
  )
}

function addOfficeConditionalFormatting(sheet: ExcelJS.Worksheet, startRow: number, endRow: number, timeProgress: number): void {
  ;(sheet as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings = []

  addThreeColorScale(sheet, 'G', startRow, endRow, 1)
  addThreeColorScale(sheet, 'I', startRow, endRow, 2)
  addDataBar(sheet, 'K', startRow, endRow, 3)
  for (const [index, column] of ['O', 'P', 'Q', 'R'].entries()) {
    addDataBar(sheet, column, startRow, endRow, 4 + index)
  }
  addCellFillRule(sheet, 'M', startRow, endRow, 'greaterThan', '40', 'FFFF00', 10)
  addCellFillRule(sheet, 'M', startRow, endRow, 'lessThan', '35', '00B050', 11)
  addCellFillRule(sheet, 'T', startRow, endRow, 'greaterThan', '0.4', 'FFFF00', 12)
  addExpressionFillRule(sheet, 'T', startRow, endRow, `T${startRow}<=0.4`, '00B050', 13)
  addCellFillRule(sheet, 'U', startRow, endRow, 'greaterThan', '0.4', 'FFFF00', 14)
  addExpressionFillRule(sheet, 'U', startRow, endRow, `U${startRow}<=0.4`, '00B050', 15)

  for (const [index, column] of ['V', 'W'].entries()) {
    sheet.addConditionalFormatting({
      ref: `${column}${startRow}:${column}${endRow}`,
      rules: [{
        type: 'expression',
        priority: 20 + index,
        formulae: [`${column}${startRow}>${timeProgress}`],
        style: { fill: solidFill('00B050') }
      }]
    })
  }
}

function comparisonMonthLabels(reportKey: string): string[] {
  const year = Number.parseInt(reportKey.slice(0, 4), 10)
  const month = Number.parseInt(reportKey.slice(4, 6), 10)
  const safeYear = Number.isFinite(year) ? year : new Date().getFullYear()
  const safeMonth = Number.isFinite(month) && month >= 1 && month <= 12 ? month : 1
  const currentIndex = safeYear * 12 + safeMonth - 1
  const startIndex = safeMonth <= 3 ? currentIndex - 3 : safeYear * 12
  const boundedStartIndex = Math.max(startIndex, currentIndex - 11)
  const labels: string[] = []
  for (let index = boundedStartIndex; index <= currentIndex; index += 1) {
    const itemYear = Math.floor(index / 12)
    const itemMonth = (index % 12) + 1
    labels.push(monthLabelFromKey(`${itemYear}${String(itemMonth).padStart(2, '0')}`))
  }
  return labels
}

function orderOfficeRowsByTemplate(sheet: ExcelJS.Worksheet, rows: OfficeRow[]): OfficeRow[] {
  const rowByJournal = new Map(rows.map((row) => [row.journal.toLowerCase(), row]))
  const templateJournals: string[] = []
  for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex += 1) {
    const journal = normalizeText(sheet.getCell(rowIndex, 1).value)
    if (journal) {
      templateJournals.push(journal)
    }
  }
  const orderedRows = templateJournals
    .map((journal) => rowByJournal.get(journal.toLowerCase()))
    .filter((row): row is OfficeRow => Boolean(row))
  return orderedRows.length ? orderedRows : rows
}

function populateOfficeSheet(sheet: ExcelJS.Worksheet, rows: OfficeRow[], context: ReportContext): void {
  const headers = [
    'Journal', 'Contrib.', 'Yearly Publ Target', 'Monthly Publ Target', 'Publ.', 'Pub/lastM', 'MoM (%)', 'Pub/YOY', 'YoY (%)',
    'New SI', 'Ave. Pub SIs', 'CN%', 'MPT', 'Sub.', 'Susy CFP', 'Mmailer-Regular CFP', 'Mmailer-FCFP', 'Send CFP (CFP+Reminder)', 'WR Target2026', 'Waiver Rate',
    'Waiver Rate2026', 'Publication TCR', 'Revenue TCR', context.quarterPubHeader, 'Yearly Publ Target', 'Revenue', context.quarterRevenueHeader, 'Revenue Target Final'
  ]

  headers.forEach((header, index) => {
    sheet.getCell(1, index + 1).value = header
  })
  leftAlignHeaderRow(sheet.getRow(1), headers.length)

  const templateRowIndex = 2
  const keyJournalSet = new Set(['foods', 'nutrients', 'children', 'genes', 'brain sciences'])
  const fillRows = rows.length || 30

  for (let rowIndex = 0; rowIndex < fillRows; rowIndex += 1) {
    const targetRow = sheet.getRow(templateRowIndex + rowIndex)
    const sourceRow = rows[rowIndex]
    const templateStyleRow = sheet.getRow(Math.min(templateRowIndex + rowIndex, sheet.rowCount))

    for (let col = 1; col <= headers.length; col += 1) {
      copyStyle(templateStyleRow.getCell(col), targetRow.getCell(col))
      targetRow.getCell(col).value = null
    }

    if (!sourceRow) {
      continue
    }

    const values: ExcelJS.CellValue[] = [
      sourceRow.journal,
      sourceRow.contrib,
      sourceRow.yearlyPublTarget,
      sourceRow.monthlyPublTarget,
      sourceRow.publ,
      sourceRow.publLast,
      sourceRow.mom,
      sourceRow.pubYoy,
      sourceRow.yoy,
      sourceRow.newSi,
      sourceRow.avePubSis,
      sourceRow.cnRate,
      sourceRow.mpt,
      sourceRow.sub,
      sourceRow.susyCfp,
      sourceRow.mmailerRegularCfp,
      sourceRow.mmailerFCfp,
      sourceRow.sendCfp,
      sourceRow.wrTarget2026,
      sourceRow.waiverRate,
      sourceRow.waiverRate2026,
      sourceRow.publicationTcr,
      sourceRow.revenueTcr,
      sourceRow.quarterPub,
      sourceRow.yearlyPublTarget,
      sourceRow.revenue,
      sourceRow.quarterRevenue,
      sourceRow.revenueTargetFinal
    ]

    values.forEach((value, index) => {
      fillCell(targetRow.getCell(index + 1), value, templateStyleRow.getCell(index + 1))
    })
    targetRow.getCell(19).fill = solidFill('FFFFFF')

    for (const col of [7, 9, 12, 19, 20, 21, 22, 23]) {
      targetRow.getCell(col).numFmt = '0.00%'
    }
    for (const col of [11]) {
      targetRow.getCell(col).numFmt = '0.00'
    }
    for (const col of [13]) {
      targetRow.getCell(col).numFmt = '0'
    }

    const isKeyJournal = keyJournalSet.has(sourceRow.journal.toLowerCase())
    targetRow.eachCell((cell) => {
      cell.font = {
        ...(cell.font ?? {}),
        size: isKeyJournal ? 16 : 11,
        bold: isKeyJournal
      }
    })
    if (isKeyJournal) {
      targetRow.eachCell((cell) => {
        cell.font = {
          ...(cell.font ?? {}),
          size: 16,
          bold: true
        }
      })
    }
  }

  addOfficeConditionalFormatting(sheet, 2, 1 + fillRows, context.timeProgress)
}

function addCompletionConditionalFormatting(sheet: ExcelJS.Worksheet, startRow: number, endRow: number, timeProgress: number): void {
  for (const column of ['F', 'I']) {
    sheet.addConditionalFormatting({
      ref: `${column}${startRow}:${column}${endRow}`,
      rules: [{
        type: 'expression',
        priority: 10,
        formulae: [`${column}${startRow}>${timeProgress}`],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFDFF3E4' }, fgColor: { argb: 'FFDFF3E4' } }
        }
      }]
    })
  }

  sheet.addConditionalFormatting({
    ref: `L${startRow}:L${endRow}`,
    rules: [{
      type: 'expression',
      priority: 12,
      formulae: [`L${startRow}-K${startRow}>0.1`],
      style: {
        fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFF1E5' }, fgColor: { argb: 'FFFFF1E5' } }
      }
    }]
  })
}

function unmergeCompletionGroupCells(sheet: ExcelJS.Worksheet, startRow: number, endRow: number): void {
  try {
    sheet.unMergeCells(`A${startRow}:A${endRow}`)
  } catch {
    // Some templates do not have a merge in this range.
  }
}

function mergeCompletionGroupCells(sheet: ExcelJS.Worksheet, rows: CompletionRow[], startRow: number): void {
  if (!rows.length) {
    return
  }

  const endRow = startRow + rows.length - 1
  unmergeCompletionGroupCells(sheet, startRow, endRow)

  let offset = 0
  while (offset < rows.length) {
    const group = rows[offset].group
    let groupEndOffset = offset
    while (groupEndOffset + 1 < rows.length && rows[groupEndOffset + 1].group === group) {
      groupEndOffset += 1
    }

    const groupStartRow = startRow + offset
    const groupEndRow = startRow + groupEndOffset
    const groupCell = sheet.getCell(groupStartRow, 1)
    groupCell.value = group
    groupCell.alignment = { ...(groupCell.alignment ?? {}), horizontal: 'center', vertical: 'middle' }
    for (let rowIndex = groupStartRow + 1; rowIndex <= groupEndRow; rowIndex += 1) {
      sheet.getCell(rowIndex, 1).value = null
    }
    if (groupEndRow > groupStartRow) {
      sheet.mergeCells(groupStartRow, 1, groupEndRow, 1)
    }

    offset = groupEndOffset + 1
  }
}

function clearCompletionDataRange(sheet: ExcelJS.Worksheet, startRow: number, endRow: number, columnCount: number): void {
  if (endRow < startRow) {
    return
  }

  unmergeCompletionGroupCells(sheet, startRow, endRow)
  for (let rowIndex = startRow; rowIndex <= endRow; rowIndex += 1) {
    for (let col = 1; col <= columnCount; col += 1) {
      sheet.getCell(rowIndex, col).value = null
    }
  }
}

function populateCompletionSheet(sheet: ExcelJS.Worksheet, rows: CompletionRow[], context: ReportContext): void {
  const headers = ['', 'Journal', 'Contrib.', context.quarterPubHeader, 'Yearly Publ Target', 'Publication TCR', context.quarterRevenueHeader, 'Revenue Target Final', 'Revenue TCR', 'Waiver Rate', 'WR Target2026', 'Waiver Rate 2026']
  headers.forEach((header, index) => {
    sheet.getCell(1, index + 1).value = header
  })
  leftAlignHeaderRow(sheet.getRow(1), headers.length)

  const topStartRow = 2
  const bottomStartRow = rows.length + 4
  const templateRow = sheet.getRow(2)
  const formulaRows = rows.length
  const clearEndRow = Math.max(sheet.rowCount, bottomStartRow + rows.length)
  clearCompletionDataRange(sheet, topStartRow, clearEndRow, headers.length)

  unmergeCompletionGroupCells(sheet, topStartRow, topStartRow + rows.length - 1)
  for (let offset = 0; offset < formulaRows; offset += 1) {
    const row = rows[offset]
    const currentRow = topStartRow + offset
    const excelRow = sheet.getRow(currentRow)
    for (let col = 1; col <= headers.length; col += 1) {
      copyStyle(templateRow.getCell(Math.min(col, templateRow.cellCount || col)), excelRow.getCell(col))
    }

    excelRow.getCell(1).value = offset === 0 || rows[offset - 1]?.group !== row.group ? row.group : ''
    excelRow.getCell(2).value = row.journal
    excelRow.getCell(3).value = row.contrib
    excelRow.getCell(4).value = row.quarterPub
    excelRow.getCell(5).value = row.yearlyPublTarget
    excelRow.getCell(6).value = row.yearlyPublTarget ? { formula: `D${currentRow}/E${currentRow}`, result: row.publicationTcr } : 0
    excelRow.getCell(7).value = row.quarterRevenue
    excelRow.getCell(8).value = row.revenueTargetFinal
    excelRow.getCell(9).value = row.revenueTargetFinal ? { formula: `G${currentRow}/H${currentRow}`, result: row.revenueTcr } : 0
    excelRow.getCell(10).value = row.waiverRate
    excelRow.getCell(11).value = row.wrTarget2026
    excelRow.getCell(12).value = row.waiverRate2026
    for (const col of [6, 9, 10, 11, 12]) {
      excelRow.getCell(col).numFmt = '0.00%'
    }
  }

  const topHeaderRow = sheet.getRow(1)
  const bottomHeaderRow = sheet.getRow(bottomStartRow)
  bottomHeaderRow.height = topHeaderRow.height
  for (let col = 1; col <= headers.length; col += 1) {
    copyStyle(topHeaderRow.getCell(col), bottomHeaderRow.getCell(col))
    bottomHeaderRow.getCell(col).value = topHeaderRow.getCell(col).value
  }
  leftAlignHeaderRow(topHeaderRow, headers.length)
  leftAlignHeaderRow(bottomHeaderRow, headers.length)
  mergeCompletionGroupCells(sheet, rows, topStartRow)

  unmergeCompletionGroupCells(sheet, bottomStartRow + 1, bottomStartRow + rows.length)
  for (let offset = 0; offset < rows.length; offset += 1) {
    const row = rows[offset]
    const currentRow = bottomStartRow + 1 + offset
    const excelRow = sheet.getRow(currentRow)
    for (let col = 1; col <= headers.length; col += 1) {
      copyStyle(templateRow.getCell(Math.min(col, templateRow.cellCount || col)), excelRow.getCell(col))
    }
    excelRow.getCell(1).value = offset === 0 || rows[offset - 1]?.group !== row.group ? row.group : ''
    excelRow.getCell(2).value = row.journal
    excelRow.getCell(3).value = row.contrib
    excelRow.getCell(4).value = row.quarterPub
    excelRow.getCell(5).value = row.yearlyPublTarget
    excelRow.getCell(6).value = row.publicationTcr
    excelRow.getCell(7).value = row.quarterRevenue
    excelRow.getCell(8).value = row.revenueTargetFinal
    excelRow.getCell(9).value = row.revenueTcr
    excelRow.getCell(10).value = row.waiverRate
    excelRow.getCell(11).value = row.wrTarget2026
    excelRow.getCell(12).value = row.waiverRate2026
    for (const col of [6, 9, 10, 11, 12]) {
      excelRow.getCell(col).numFmt = '0.00%'
    }
  }
  mergeCompletionGroupCells(sheet, rows, bottomStartRow + 1)

  addCompletionConditionalFormatting(sheet, topStartRow, topStartRow + rows.length - 1, context.timeProgress)
  addCompletionConditionalFormatting(sheet, bottomStartRow + 1, bottomStartRow + rows.length, context.timeProgress)
}

function computeDepartmentMetrics(journalRows: GenericRow[], context: ReportContext, overrides: PipelineInput['summaryOverrides']): DepartmentMetrics {
  const healthRows = journalRows.filter((row) => normalizeText(row.Section) === 'Section Health')
  const publication = toInteger(healthRows.reduce((sum, row) => sum + normalizeNumber(row['Publ.']), 0))
  const publicationLast = toInteger(healthRows.reduce((sum, row) => sum + normalizeNumber(row['Pub/lastM']), 0))
  const publicationYoySource = toInteger(healthRows.reduce((sum, row) => sum + normalizeNumber(row['Pub/YOY']), 0))
  const submission = toInteger(healthRows.reduce((sum, row) => sum + normalizeNumber(row['Sub.']), 0))
  const submissionLast = toInteger(healthRows.reduce((sum, row) => sum + normalizeNumber(row['Sub/lastM']), 0))
  const assignedManuscript = toInteger(healthRows.reduce((sum, row) => sum + normalizeNumber(row.Processing), 0))
  const siSetUp = toInteger(healthRows.reduce((sum, row) => sum + normalizeNumber(row['SIs(Open)']), 0))
  const revenue = healthRows.reduce((sum, row) => sum + normalizeNumber(row.Revenue), 0) / 10000
  const revenueLast = healthRows.reduce((sum, row) => sum + normalizeNumber(row['Revenue/LastM']), 0) / 10000
  const revenueYoySource = healthRows.reduce((sum, row) => sum + normalizeNumber(row['Revenue/YoY']), 0) / 10000
  const totalPublWeight = Math.max(1, healthRows.reduce((sum, row) => sum + normalizeNumber(row['Publ.']), 0))
  const waiverRate = healthRows.reduce((sum, row) => sum + normalizeNumber(row['Waiver Rate']) * normalizeNumber(row['Publ.']), 0) / totalPublWeight
  const mpt = healthRows.reduce((sum, row) => sum + normalizeNumber(row.MPT) * normalizeNumber(row['Publ.']), 0) / totalPublWeight

  const currentYearSeries = context.comparisonMonthLabels.map((label) => ({
    label,
    publication: 0,
    submission: 0,
    assignedManuscript: 0,
    siSetUp: 0,
    revenueWCHF: 0,
    waiverRate: 0,
    mpt: 0
  }))

  return {
    reportMonthCellLabel: overrides.reportMonthLabel || context.reportMonthLabel,
    reportMonthTitle: overrides.reportMonthLabel || context.reportMonthTitle,
    publication: overrides.publication ? Number(overrides.publication) : publication,
    submission: overrides.submission ? Number(overrides.submission) : submission,
    assignedManuscript: overrides.assignedManuscript ? Number(overrides.assignedManuscript) : assignedManuscript,
    siSetUp: overrides.siSetUp ? Number(overrides.siSetUp) : siSetUp,
    revenueWCHF: overrides.revenueWCHF ? Number(overrides.revenueWCHF) : Number(revenue.toFixed(2)),
    waiverRate: overrides.waiverRate ? Number(overrides.waiverRate) / 100 : waiverRate,
    mpt: overrides.mpt ? Number(overrides.mpt) : Number(mpt.toFixed(0)),
    publicationDelta: (overrides.publication ? Number(overrides.publication) : publication) - publicationLast,
    publicationMom: percentageDelta(overrides.publication ? Number(overrides.publication) : publication, publicationLast),
    publicationYoy: publicationYoySource ? percentageDelta(overrides.publication ? Number(overrides.publication) : publication, publicationYoySource) : null,
    revenueDelta: Number(((overrides.revenueWCHF ? Number(overrides.revenueWCHF) : revenue) - revenueLast).toFixed(2)),
    revenueMom: percentageDelta(overrides.revenueWCHF ? Number(overrides.revenueWCHF) : revenue, revenueLast),
    revenueYoy: revenueYoySource ? percentageDelta(overrides.revenueWCHF ? Number(overrides.revenueWCHF) : revenue, revenueYoySource) : null,
    submissionDelta: (overrides.submission ? Number(overrides.submission) : submission) - submissionLast,
    submissionMom: percentageDelta(overrides.submission ? Number(overrides.submission) : submission, submissionLast),
    yearlySeries: currentYearSeries
  }
}

function departmentRowHasValues(sheet: ExcelJS.Worksheet, rowIndex: number): boolean {
  for (let col = 2; col <= 8; col += 1) {
    if (normalizeText(sheet.getCell(rowIndex, col).value)) {
      return true
    }
  }
  return false
}

function syncDepartmentMetricsFromSheet(sheet: ExcelJS.Worksheet, rowIndex: number, metrics: DepartmentMetrics): void {
  const currentPublication = toInteger(normalizeNumber(sheet.getCell(rowIndex, 2).value))
  const currentSubmission = toInteger(normalizeNumber(sheet.getCell(rowIndex, 3).value))
  const currentRevenue = normalizeNumber(sheet.getCell(rowIndex, 6).value)
  const previousPublication = rowIndex > 2 ? toInteger(normalizeNumber(sheet.getCell(rowIndex - 1, 2).value)) : 0
  const previousSubmission = rowIndex > 2 ? toInteger(normalizeNumber(sheet.getCell(rowIndex - 1, 3).value)) : 0
  const previousRevenue = rowIndex > 2 ? normalizeNumber(sheet.getCell(rowIndex - 1, 6).value) : 0

  metrics.publication = currentPublication
  metrics.submission = currentSubmission
  metrics.assignedManuscript = toInteger(normalizeNumber(sheet.getCell(rowIndex, 4).value))
  metrics.siSetUp = toInteger(normalizeNumber(sheet.getCell(rowIndex, 5).value))
  metrics.revenueWCHF = currentRevenue
  metrics.waiverRate = normalizeNumber(sheet.getCell(rowIndex, 7).value) / 100
  metrics.mpt = normalizeNumber(sheet.getCell(rowIndex, 8).value)
  metrics.publicationDelta = currentPublication - previousPublication
  metrics.publicationMom = percentageDelta(currentPublication, previousPublication)
  metrics.revenueDelta = Number((currentRevenue - previousRevenue).toFixed(2))
  metrics.revenueMom = percentageDelta(currentRevenue, previousRevenue)
  metrics.submissionDelta = currentSubmission - previousSubmission
  metrics.submissionMom = percentageDelta(currentSubmission, previousSubmission)

  const monthIndex = rowIndex - 1
  const previousYearPublication = normalizeNumber(sheet.getCell(20, monthIndex + 1).value)
  const previousYearRevenue = normalizeNumber(sheet.getCell(21, monthIndex + 1).value)
  metrics.publicationYoy = previousYearPublication ? percentageDelta(currentPublication, previousYearPublication) : metrics.publicationYoy
  metrics.revenueYoy = previousYearRevenue ? percentageDelta(currentRevenue, previousYearRevenue) : metrics.revenueYoy
}

function populateDepartmentSheet(sheet: ExcelJS.Worksheet, metrics: DepartmentMetrics, context: ReportContext): void {
  leftAlignHeaderRow(sheet.getRow(1), Math.max(sheet.getRow(1).cellCount, 8))
  const targetLabel = metrics.reportMonthCellLabel
  let targetRowIndex = -1
  for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex += 1) {
    const label = normalizeText(sheet.getCell(rowIndex, 1).value)
    if (label === targetLabel) {
      targetRowIndex = rowIndex
      break
    }
  }

  if (targetRowIndex === -1) {
    targetRowIndex = sheet.rowCount + 1
    sheet.getCell(targetRowIndex, 1).value = targetLabel
  }

  if (!departmentRowHasValues(sheet, targetRowIndex)) {
    const row = sheet.getRow(targetRowIndex)
    row.getCell(2).value = metrics.publication
    row.getCell(3).value = metrics.submission
    row.getCell(4).value = metrics.assignedManuscript
    row.getCell(5).value = metrics.siSetUp
    row.getCell(6).value = metrics.revenueWCHF
    row.getCell(7).value = Number((metrics.waiverRate * 100).toFixed(2))
    row.getCell(8).value = metrics.mpt
    row.getCell(7).numFmt = '0'
  }

  syncDepartmentMetricsFromSheet(sheet, targetRowIndex, metrics)

  const rowByLabel = new Map<string, number>()
  for (let rowIndex = 2; rowIndex <= 13; rowIndex += 1) {
    const label = normalizeText(sheet.getCell(rowIndex, 1).value)
    if (label) {
      rowByLabel.set(label, rowIndex)
    }
  }

  const yearlySeries: DepartmentMetrics['yearlySeries'] = []
  for (const label of context.comparisonMonthLabels) {
    const rowIndex = rowByLabel.get(label)
    if (!rowIndex) {
      continue
    }
    yearlySeries.push({
      label,
      publication: normalizeNumber(sheet.getCell(rowIndex, 2).value),
      submission: normalizeNumber(sheet.getCell(rowIndex, 3).value),
      assignedManuscript: normalizeNumber(sheet.getCell(rowIndex, 4).value),
      siSetUp: normalizeNumber(sheet.getCell(rowIndex, 5).value),
      revenueWCHF: normalizeNumber(sheet.getCell(rowIndex, 6).value),
      waiverRate: normalizeNumber(sheet.getCell(rowIndex, 7).value),
      mpt: normalizeNumber(sheet.getCell(rowIndex, 8).value)
    })
  }
  metrics.yearlySeries = yearlySeries
}

function lastNonEmptyRowInRange(sheet: ExcelJS.Worksheet, startRow: number, endRow: number, startCol: number, endCol: number): number {
  for (let rowIndex = endRow; rowIndex >= startRow; rowIndex -= 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      if (normalizeText(sheet.getCell(rowIndex, col).value)) {
        return rowIndex
      }
    }
  }
  return startRow
}

function snapshotColumnValues(sheet: ExcelJS.Worksheet, startRow: number, endRow: number, col: number): number[] {
  const values: number[] = []
  for (let rowIndex = startRow; rowIndex <= endRow; rowIndex += 1) {
    const value = normalizeNumber(sheet.getCell(rowIndex, col).value)
    if (Number.isFinite(value)) {
      values.push(value)
    }
  }
  return values
}

function threeColorScale(value: number, min: number, max: number): string {
  if (max <= min) {
    return '#FFEB84'
  }
  const midpoint = min + (max - min) / 2
  if (value <= midpoint) {
    return `#${interpolateColor(value, min, midpoint, 'F8696B', 'FFEB84')}`
  }
  return `#${interpolateColor(value, midpoint, max, 'FFEB84', '63BE7B')}`
}

function officeSnapshotOverlay(sheet: ExcelJS.Worksheet, startRow: number, endRow: number): (cell: ExcelJS.Cell, row: number, col: number) => SnapshotOverlay {
  const colorScaleColumns = new Map([7, 9].map((col) => {
    const values = snapshotColumnValues(sheet, startRow, endRow, col)
    return [col, { min: Math.min(...values), max: Math.max(...values) }] as const
  }))
  const dataBarColumns = new Map([11].map((col) => {
    const values = snapshotColumnValues(sheet, startRow, endRow, col)
    return [col, Math.max(...values, 1)] as const
  }))

  return (cell, row, col) => {
    if (row < startRow || row > endRow) {
      return {}
    }
    const value = normalizeNumber(cell.value)
    const scale = colorScaleColumns.get(col)
    if (scale) {
      return { background: threeColorScale(value, scale.min, scale.max) }
    }
    const dataBarMax = dataBarColumns.get(col)
    if (dataBarMax) {
      return { dataBarRatio: value / dataBarMax, dataBarColor: '#638EC6' }
    }
    if (col === 13) {
      if (value < 35) {
        return { background: '#00B050' }
      }
      if (value > 40) {
        return { background: '#FFFF00' }
      }
    }
    if (col === 20 || col === 21) {
      return { background: value > 0.4 ? '#FFFF00' : '#00B050' }
    }
    return {}
  }
}

function completionSnapshotOverlay(timeProgress: number): (cell: ExcelJS.Cell, row: number, col: number) => SnapshotOverlay {
  return (cell, row, col) => {
    if (row === 1) {
      return {}
    }
    const value = normalizeNumber(cell.value)
    if ((col === 6 || col === 9) && value > timeProgress) {
      return { background: '#DFF3E4' }
    }
    if (col === 12) {
      const wrTarget2026 = normalizeNumber(cell.worksheet.getCell(row, 11).value)
      if (value - wrTarget2026 > 0.1) {
        return { background: '#FFF1E5' }
      }
    }
    return {}
  }
}

function completionSnapshotMerges(rows: CompletionRow[]): Array<{ startRow: number; endRow: number; col: number; value: string }> {
  const merges: Array<{ startRow: number; endRow: number; col: number; value: string }> = []
  let offset = 0
  while (offset < rows.length) {
    const group = rows[offset].group
    let endOffset = offset
    while (endOffset + 1 < rows.length && rows[endOffset + 1].group === group) {
      endOffset += 1
    }
    const startRow = 2 + offset
    const endRow = 2 + endOffset
    if (group !== 'SCIE' && startRow <= 22 && endRow >= 7) {
      merges.push({
        startRow: Math.max(7, startRow),
        endRow: Math.min(22, endRow),
        col: 1,
        value: group
      })
    }
    offset = endOffset + 1
  }
  return merges
}

function buildPptSnapshots(officeSheet: ExcelJS.Worksheet, completionSheet: ExcelJS.Worksheet, completionRows: CompletionRow[], context: ReportContext): PptSnapshots {
  const officeEndRow = lastNonEmptyRowInRange(officeSheet, 1, officeSheet.rowCount, 1, 21)
  const officeSnapshotColumns = Array.from({ length: 21 }, (_, index) => index + 1)
    .filter((col) => ![15, 16, 17, 18].includes(col))
  return {
    officeSvg: renderWorksheetRangeSvg(officeSheet, {
      startRow: 1,
      endRow: officeEndRow,
      startCol: 1,
      endCol: 21,
      colIndices: officeSnapshotColumns,
      defaultColWidth: 9,
      defaultRowHeight: 18,
      fontScale: 0.82,
      cellOverlay: officeSnapshotOverlay(officeSheet, 2, officeEndRow)
    }),
    completionSvg: renderWorksheetRangeSvg(completionSheet, {
      startRow: 1,
      endRow: 22,
      startCol: 1,
      endCol: 12,
      rowIndices: [1, ...Array.from({ length: 16 }, (_, index) => index + 7)],
      defaultColWidth: 12,
      defaultRowHeight: 19,
      fontScale: 0.9,
      cellOverlay: completionSnapshotOverlay(context.timeProgress),
      verticalMerges: completionSnapshotMerges(completionRows)
    }),
    staffPiCompletionSvgs: [],
    staffOwnerSvgs: [],
    staffSiSetupSvgs: [],
    staffSiSubSvgs: [],
    staffAePublSvgs: []
  }
}

function focusMetricRows(sheet: ExcelJS.Worksheet, headerRowIndex: number): number[] {
  const rows: number[] = []
  for (let rowIndex = headerRowIndex + 1; rowIndex <= sheet.rowCount; rowIndex += 1) {
    const label = normalizeText(sheet.getCell(rowIndex, 1).value)
    if (!label) {
      break
    }
    rows.push(rowIndex)
  }
  return rows
}

function compactFocusMetricBlockColumns(sheet: ExcelJS.Worksheet, headerRow: ExcelJS.Row, headerRowIndex: number, monthLabels: string[]): void {
  const metricRows = focusMetricRows(sheet, headerRowIndex)
  if (!metricRows.length || !monthLabels.length) {
    return
  }

  const maxMonthColumns = Math.max(headerRow.cellCount, 13)
  const sourceByLabel = new Map<string, number>()
  for (let col = 2; col <= maxMonthColumns; col += 1) {
    const label = normalizeText(headerRow.getCell(col).value)
    if (label && !sourceByLabel.has(label)) {
      sourceByLabel.set(label, col)
    }
  }

  const selected = monthLabels
    .map((label) => ({ label, sourceCol: sourceByLabel.get(label) }))
    .filter((item): item is { label: string; sourceCol: number } => item.sourceCol !== undefined)
  if (!selected.length) {
    return
  }

  const rowsToMove = [headerRowIndex, ...metricRows]
  const snapshots = selected.map(({ sourceCol }) => rowsToMove.map((rowIndex) => {
    const sourceCell = sheet.getCell(rowIndex, sourceCol)
    return {
      value: sourceCell.value,
      style: deepClone(sourceCell.style),
      numFmt: sourceCell.numFmt
    }
  }))

  for (const rowIndex of rowsToMove) {
    for (let col = 2; col <= maxMonthColumns; col += 1) {
      sheet.getCell(rowIndex, col).value = null
    }
  }

  snapshots.forEach((columnSnapshot, index) => {
    const targetCol = index + 2
    columnSnapshot.forEach((snapshot, rowOffset) => {
      const targetCell = sheet.getCell(rowsToMove[rowOffset], targetCol)
      targetCell.value = snapshot.value as ExcelJS.CellValue
      targetCell.style = snapshot.style
      if (snapshot.numFmt) {
        targetCell.numFmt = snapshot.numFmt
      }
    })
  })
}

function updateFocusJournalSheet(sheet: ExcelJS.Worksheet, context: ReportContext, rowSource: GenericRow | undefined): FocusJournalSeries {
  const displayName = sheet.name === 'BS' ? 'Brain Sciences' : sheet.name
  const resolvedSource = rowSource ?? {}
  const reportMonthLabel = context.reportMonthLabel

  const headerRows: ExcelJS.Row[] = []
  for (let rowIndex = 1; rowIndex <= sheet.rowCount; rowIndex += 1) {
    const row = sheet.getRow(rowIndex)
    const hasMonthHeader = Array.from({ length: row.cellCount }, (_, index) => normalizeText(row.getCell(index + 1).value))
      .some((value) => value === reportMonthLabel || value === 'Jan.' || value === 'Feb.' || value === 'Mar.')
    if (hasMonthHeader) {
      headerRows.push(row)
    }
    if (headerRows.length >= 2) {
      break
    }
  }

  const firstHeaderRow = headerRows[0] ?? sheet.getRow(1)
  const secondHeaderRow = headerRows[1] ?? sheet.getRow(5)
  const firstHeaderRowIndex = firstHeaderRow.number
  const secondHeaderRowIndex = secondHeaderRow.number
  const findReportColumn = (row: ExcelJS.Row) => {
    for (let col = 2; col <= Math.max(row.cellCount, 13); col += 1) {
      if (normalizeText(row.getCell(col).value) === reportMonthLabel) {
        return col
      }
    }
    return -1
  }
  const firstColumn = findReportColumn(firstHeaderRow)
  const secondColumn = findReportColumn(secondHeaderRow)

  if (firstColumn !== -1) {
    sheet.getCell(firstHeaderRowIndex + 1, firstColumn).value = toInteger(normalizeNumber(resolvedSource['Publ.']))
    sheet.getCell(firstHeaderRowIndex + 2, firstColumn).value = toInteger(normalizeNumber(resolvedSource['Sub.']))
    sheet.getCell(firstHeaderRowIndex + 3, firstColumn).value = toInteger(normalizeNumber(resolvedSource.Processing))
  }
  if (secondColumn !== -1) {
    sheet.getCell(secondHeaderRowIndex + 1, secondColumn).value = normalizeNumber(resolvedSource.MPT)
    sheet.getCell(secondHeaderRowIndex + 2, secondColumn).value = normalizeNumber(resolvedSource.TFD)
  }
  leftAlignHeaderRow(firstHeaderRow, firstHeaderRow.cellCount)
  leftAlignHeaderRow(secondHeaderRow, secondHeaderRow.cellCount)

  compactFocusMetricBlockColumns(sheet, firstHeaderRow, firstHeaderRowIndex, context.comparisonMonthLabels)
  compactFocusMetricBlockColumns(sheet, secondHeaderRow, secondHeaderRowIndex, context.comparisonMonthLabels)

  const collectMetricBlock = (headerRow: ExcelJS.Row, headerRowIndex: number) => {
    const metricRows: number[] = []
    for (let rowIndex = headerRowIndex + 1; rowIndex <= sheet.rowCount; rowIndex += 1) {
      const label = normalizeText(sheet.getCell(rowIndex, 1).value)
      if (!label) {
        break
      }
      metricRows.push(rowIndex)
    }

    const columns: number[] = []
    const months: string[] = []
    for (let col = 2; col <= Math.max(headerRow.cellCount, 13); col += 1) {
      const label = normalizeText(headerRow.getCell(col).value)
      if (!label) {
        continue
      }
      const hasData = metricRows.some((rowIndex) => normalizeText(sheet.getCell(rowIndex, col).value) !== '')
      if (hasData) {
        columns.push(col)
        months.push(label)
      }
    }

    return {
      months,
      metrics: metricRows.map((rowIndex) => ({
        label: normalizeText(sheet.getCell(rowIndex, 1).value),
        values: columns.map((col) => normalizeNumber(sheet.getCell(rowIndex, col).value))
      }))
    }
  }

  const primaryBlock = collectMetricBlock(firstHeaderRow, firstHeaderRowIndex)
  const secondaryBlock = collectMetricBlock(secondHeaderRow, secondHeaderRowIndex)
  const months: string[] = []
  const publication: number[] = []
  const submission: number[] = []
  const underProcessing: number[] = []
  const mpt: number[] = []
  const tfd: number[] = []

  for (let col = 2; col <= firstHeaderRow.cellCount; col += 1) {
    const label = normalizeText(firstHeaderRow.getCell(col).value)
    if (!label) {
      continue
    }
    months.push(label)
    publication.push(normalizeNumber(sheet.getCell(firstHeaderRowIndex + 1, col).value))
    submission.push(normalizeNumber(sheet.getCell(firstHeaderRowIndex + 2, col).value))
    underProcessing.push(normalizeNumber(sheet.getCell(firstHeaderRowIndex + 3, col).value))
  }
  for (let col = 2; col <= secondHeaderRow.cellCount; col += 1) {
    const label = normalizeText(secondHeaderRow.getCell(col).value)
    if (!label) {
      continue
    }
    if (!months.includes(label)) {
      months.push(label)
    }
    mpt.push(normalizeNumber(sheet.getCell(secondHeaderRowIndex + 1, col).value))
    tfd.push(normalizeNumber(sheet.getCell(secondHeaderRowIndex + 2, col).value))
  }

  return {
    sheetName: sheet.name,
    displayName,
    months,
    publication,
    submission,
    underProcessing,
    mpt,
    tfd,
    primaryMetrics: primaryBlock.metrics,
    secondaryMetrics: secondaryBlock.metrics
  }
}

function isJournalMapPlaceholder(value: string): boolean {
  if (!value || /[\u4e00-\u9fa5]/.test(value)) {
    return true
  }
  const normalized = value.replace(/[.\s]+/g, '').toUpperCase()
  return ['HD', 'TJ', 'NJ', 'AE', 'ME', 'SME', 'P', 'AP'].includes(normalized)
}

function canonicalPersonName(value: string): string {
  return compactName(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ')
}

function personLookupKeys(value: string): string[] {
  const compact = normalizePersonName(value)
  const canonical = canonicalPersonName(value)
  const tokens = canonical.split(' ').filter(Boolean)
  const sortedTokens = [...tokens].sort().join(' ')
  return Array.from(new Set([
    compact,
    canonical,
    sortedTokens,
    tokens.length >= 2 ? `${tokens[0]} ${tokens[tokens.length - 1]}` : '',
    tokens.length >= 3 ? `${tokens[0]} ${tokens[1]} ${tokens[tokens.length - 1]}` : ''
  ].filter(Boolean)))
}

function addJournalMapEntry(map: Map<string, Set<string>>, name: string, journal: string): void {
  for (const key of personLookupKeys(name)) {
    if (!map.has(key)) {
      map.set(key, new Set())
    }
    map.get(key)?.add(journal)
  }
}

function buildJournalMap(workbook: ExcelJS.Workbook): Map<string, string[]> {
  const sheet = workbook.worksheets[0]
  const map = new Map<string, Set<string>>()
  const row1 = sheet.getRow(1)

  for (let col = 1; col <= row1.cellCount; col += 1) {
    const journal = compactName(normalizeText(row1.getCell(col).value))
    if (!journal || /\d/.test(journal) || journal === 'ESCI') {
      continue
    }
    for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex += 1) {
      const rawName = compactName(normalizeText(sheet.getCell(rowIndex, col).value))
      if (isJournalMapPlaceholder(rawName)) {
        continue
      }
      addJournalMapEntry(map, rawName, journal === 'BS' ? 'Brain Sciences' : journal)
    }
  }

  return new Map(Array.from(map.entries()).map(([key, journals]) => [key, Array.from(journals)]))
}

function mergeStaffTemplateJournalMap(workbook: ExcelJS.Workbook, journalMap: Map<string, string[]>): void {
  const sheet = workbook.worksheets[0]
  if (!sheet) {
    return
  }
  const headers = Array.from({ length: sheet.getRow(1).cellCount }, (_, index) => normalizeText(sheet.getCell(1, index + 1).value))
  const staffColumn = headers.indexOf('Staff') + 1
  const journalColumn = headers.indexOf('Journal') + 1
  if (!staffColumn || !journalColumn) {
    return
  }

  const mergedMap = new Map(Array.from(journalMap.entries()).map(([key, journals]) => [key, new Set(journals)]))
  for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex += 1) {
    const staffName = compactName(normalizeText(sheet.getCell(rowIndex, staffColumn).value))
    const journals = normalizeText(sheet.getCell(rowIndex, journalColumn).value)
      .split(/[,，]/)
      .map((journal) => compactName(journal))
      .filter(Boolean)
    for (const journal of journals) {
      addJournalMapEntry(mergedMap, staffName, journal)
    }
  }

  journalMap.clear()
  for (const [key, journals] of mergedMap.entries()) {
    journalMap.set(key, Array.from(journals))
  }
}

function staffJournalsForName(journalMap: Map<string, string[]>, name: string): string[] {
  for (const key of personLookupKeys(name)) {
    const journals = journalMap.get(key)
    if (journals?.length) {
      return journals
    }
  }
  return [MISSING_JOURNAL_LABEL]
}

const AE_STAFF_HEADERS = [
  'Staff', 'Journal', 'Join Date', 'Contrib.', 'PI (article owner)', 'PI', 'EPI', 'PIT', 'PI Completion Rate', 'Publ',
  'MPT', 'TFD', 'New Reviewer With Report (Score>2)', 'Average APT (80%)', 'Publ Other', 'Acceptance to Publication Time',
  'Under Process', 'Finalizing', 'Resubmitted Articles', 'Editor Role', 'Ed. Group', 'Office', 'Is Active'
]

const SME_STAFF_HEADERS = [
  'Staff', 'Journal', 'Join Date', 'Contrib.', 'PI (article owner)', 'PI', 'EPI', 'PIT', 'PI Completion Rate', 'SI Sub.',
  'SI Publ', 'Regular Publ', 'SI Setup', 'SI Open', 'SI Under Process.', 'Send GE Invites', 'GE Agree Rate',
  'New Accepted SI', 'All Accepted SI', 'Closed SI with 10+ Publication', 'Closed Special Issue Average Publication',
  'Send EBM Invites', 'Under Process(Journals)', 'Susy CFP', 'CFP Team CFP', 'Mmailer-Regular CFP', 'Mmailer-FCFP',
  'Editor Role', 'Ed. Group', 'Office', 'Is Active'
]

const ALL_STAFF_HEADERS = [
  ...AE_STAFF_HEADERS.slice(0, 19),
  ...SME_STAFF_HEADERS.slice(9, 27),
  'Editor Role', 'Ed. Group', 'Office', 'Is Active'
]

const STAFF_SOURCE_HEADERS: Record<string, string> = {
  'Join Date': 'Join Date',
  'Contrib.': 'Contrib.',
  'PI (article owner)': 'Performance Index (article owner)',
  PI: 'Performance Index',
  EPI: 'Extend Performance Index',
  PIT: 'PIT',
  'PI Completion Rate': 'PI Completion Rate',
  Publ: 'Publ',
  MPT: 'Median APT',
  TFD: 'TFD',
  'New Reviewer With Report (Score>2)': 'New Reviewer With Report (Score>2)',
  'Average APT (80%)': 'Average APT (80%)',
  'Publ Other': 'Publ Other',
  'Acceptance to Publication Time': 'Acceptance to Publication Time',
  'Under Process': 'Under Process',
  Finalizing: 'Finalizing',
  'Resubmitted Articles': 'Resubmitted Articles',
  'SI Sub.': 'SI Sub.',
  'SI Publ': 'SI Publ',
  'Regular Publ': 'Regular Publ',
  'SI Setup': 'Special Issues Setup',
  'SI Open': 'Special Issues Open',
  'SI Under Process.': 'SI Under Process.',
  'Send GE Invites': 'Send GE Invites',
  'GE Agree Rate': 'GE Agree Rate',
  'New Accepted SI': 'New Accepted SI',
  'All Accepted SI': 'All Accepted SI',
  'Closed SI with 10+ Publication': 'Closed SI with 10+ Publication',
  'Closed Special Issue Average Publication': 'Closed Special Issue Average Publication',
  'Send EBM Invites': 'Send EBM Invites',
  'Under Process(Journals)': 'Under Process(Journals)',
  'Susy CFP': 'Susy CFP',
  'CFP Team CFP': 'CFP Team CFP',
  'Mmailer-Regular CFP': 'Mmailer-Regular CFP',
  'Mmailer-FCFP': 'Mmailer-FCFP',
  'Editor Role': 'Editor Role',
  'Ed. Group': 'Ed. Group',
  Office: 'Office',
  'Is Active': 'Is Active'
}

const MISSING_JOURNAL_LABEL = 'Missing Journal Mapping'

const STAFF_JOURNAL_COLORS = new Map<string, string>([
  ['nutrients', '12542F'],
  ['foods', '588715'],
  ['genes', '5F487A'],
  ['children', '1D7EA1'],
  ['brain sciences', '005C6F'],
  ['sports', '80503C'],
  ['diabetology', 'D62129'],
  ['beverages', 'C46F0A'],
  ['livers', 'B84329'],
  ['antibodies', '41184A'],
  ['endocrines', 'C3186D'],
  ['dietetics', '759E1F'],
  ['dna', '0F77B4'],
  ['surgical techniques development', '2682BA'],
  ['jvd', '6D1B1E'],
  ['cardiogenetics', '6262DE'],
  ['jmahp', '783CB4'],
  ['allergies', '729C1C'],
  ['jdad', 'BA7B1E'],
  ['sustainable foods', '0D8440'],
  ['food engineering', 'E06920']
])

const SCIE_MANAGER_JOURNALS = new Set(['foods', 'nutrients', 'children', 'genes', 'brain sciences'])

function firstStaffJournalColor(journals: string[]): string | null {
  if (journals[0] === MISSING_JOURNAL_LABEL) {
    return 'FF7D7D'
  }
  for (const journal of journals) {
    const color = STAFF_JOURNAL_COLORS.get(journal.toLowerCase())
    if (color) {
      return color
    }
  }
  return null
}

function staffCellValue(editor: GenericRow, header: string): ExcelJS.CellValue {
  const sourceHeader = STAFF_SOURCE_HEADERS[header]
  if (!sourceHeader) {
    return null
  }
  if (header === 'PI Completion Rate') {
    return normalizeStaffPiCompletionRate(editor[sourceHeader])
  }
  if (header === 'GE Agree Rate') {
    return normalizeRate(editor[sourceHeader])
  }
  return editor[sourceHeader] as ExcelJS.CellValue
}

function staffPiCompletionRate(editor: GenericRow): number {
  return normalizeStaffPiCompletionRate(editor[STAFF_SOURCE_HEADERS['PI Completion Rate']])
}

function sortStaffRowsByPiCompletion(rows: GenericRow[]): GenericRow[] {
  return [...rows].sort((left, right) => staffPiCompletionRate(right) - staffPiCompletionRate(left))
}

function isScieManagingEditor(row: GenericRow, journalMap: Map<string, string[]>): boolean {
  if (normalizeText(row['Editor Role']) !== 'Managing Editor') {
    return false
  }
  const name = compactName(normalizeText(row.Staff))
  return staffJournalsForName(journalMap, name).some((journal) => SCIE_MANAGER_JOURNALS.has(journal.toLowerCase()))
}

function styleStaffSheet(sheet: ExcelJS.Worksheet, headers: string[], rowCount: number): void {
  sheet.views = [{ state: 'frozen', ySplit: 1 }]
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: headers.length }
  }
  const headerRow = sheet.getRow(1)
  headerRow.height = 22
  headerRow.eachCell((cell) => {
    cell.font = { name: 'Source Han Sans CN', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 }
    cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }
    cell.fill = solidFill('24292F')
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFD0D7DE' } } }
  })
  for (let col = 1; col <= headers.length; col += 1) {
    const header = headers[col - 1]
    sheet.getColumn(col).width = header === 'Journal' ? 36 : Math.max(10, Math.min(28, header.length + 4))
  }
  for (let rowIndex = 2; rowIndex <= rowCount + 1; rowIndex += 1) {
    const row = sheet.getRow(rowIndex)
    row.eachCell((cell) => {
      cell.font = { name: 'Source Han Sans CN', size: 10 }
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: false }
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFD8DEE4' } } }
    })
  }
}

function applyStaffJournalBackgrounds(sheet: ExcelJS.Worksheet, rows: GenericRow[], journalMap: Map<string, string[]>): void {
  const journalColumnIndex = 2
  rows.forEach((editor, rowOffset) => {
    const name = compactName(normalizeText(editor.Staff))
    const journals = staffJournalsForName(journalMap, name)
    const fillColor = firstStaffJournalColor(journals)
    if (!fillColor) {
      return
    }
    const cell = sheet.getRow(rowOffset + 2).getCell(journalColumnIndex)
    cell.fill = solidFill(fillColor)
    cell.font = {
      ...(cell.font ?? {}),
      name: 'Source Han Sans CN',
      size: 10,
      bold: true,
      color: { argb: 'FFFFFFFF' }
    }
  })
}

function writeStaffSheet(sheet: ExcelJS.Worksheet, headers: string[], rows: GenericRow[], journalMap: Map<string, string[]>): void {
  headers.forEach((header, index) => {
    sheet.getCell(1, index + 1).value = header
  })

  rows.forEach((editor, rowOffset) => {
    const row = sheet.getRow(rowOffset + 2)
    const name = compactName(normalizeText(editor.Staff))
    headers.forEach((header, index) => {
      const cell = row.getCell(index + 1)
      if (header === 'Staff') {
        cell.value = name
      } else if (header === 'Journal') {
        cell.value = staffJournalsForName(journalMap, name).join(', ')
      } else {
        cell.value = staffCellValue(editor, header)
      }
      if (header === 'PI Completion Rate' || header === 'GE Agree Rate') {
        cell.numFmt = '0.00%'
      }
    })
  })

  styleStaffSheet(sheet, headers, rows.length)
  applyStaffJournalBackgrounds(sheet, rows, journalMap)
}

function verifyStaffPiCompletionValues(sheet: ExcelJS.Worksheet, headers: string[], rows: GenericRow[]): void {
  const piColumn = headers.indexOf('PI Completion Rate') + 1
  if (!piColumn) {
    return
  }
  rows.forEach((editor, rowOffset) => {
    const expected = staffPiCompletionRate(editor)
    const actual = normalizeStaffPiCompletionRate(sheet.getRow(rowOffset + 2).getCell(piColumn).value)
    if (Math.abs(expected - actual) > 0.000001) {
      const staffName = compactName(normalizeText(editor.Staff))
      throw new Error(`PI Completion Rate mismatch for ${staffName || `row ${rowOffset + 2}`}`)
    }
  })
}

function activeHealthEditorPiCompletionSource(editorRows: GenericRow[]): Map<string, number> {
  const source = new Map<string, number>()
  const ambiguousNames = new Set<string>()
  for (const editor of editorRows) {
    if (normalizeText(editor.Section) !== 'Section Health' || normalizeText(editor['Is Active']) !== 'Activate') {
      continue
    }
    const staffName = compactName(normalizeText(editor.Staff))
    if (!staffName) {
      continue
    }
    const piCompletionRate = staffPiCompletionRate(editor)
    const existing = source.get(staffName)
    if (existing !== undefined && Math.abs(existing - piCompletionRate) > 0.000001) {
      ambiguousNames.add(staffName)
      continue
    }
    source.set(staffName, piCompletionRate)
  }
  if (ambiguousNames.size) {
    throw new Error(`Duplicate staff names with different PI Completion Rate in MR: ${Array.from(ambiguousNames).join(', ')}`)
  }
  return source
}

async function verifySavedStaffWorkbookPiCompletion(filePath: string, editorRows: GenericRow[]): Promise<void> {
  const source = activeHealthEditorPiCompletionSource(editorRows)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)
  const mismatches: string[] = []

  for (const sheet of workbook.worksheets) {
    const headers = Array.from({ length: sheet.getRow(1).cellCount }, (_, index) => normalizeText(sheet.getCell(1, index + 1).value))
    const staffColumn = headers.indexOf('Staff') + 1
    const piColumn = headers.indexOf('PI Completion Rate') + 1
    if (!staffColumn || !piColumn) {
      continue
    }
    for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex += 1) {
      const staffName = compactName(normalizeText(sheet.getCell(rowIndex, staffColumn).value))
      if (!staffName) {
        continue
      }
      const expected = source.get(staffName)
      const actual = normalizeStaffPiCompletionRate(sheet.getCell(rowIndex, piColumn).value)
      if (expected === undefined) {
        mismatches.push(`${sheet.name}!${rowIndex} ${staffName}: not found in MR Editors`)
      } else if (Math.abs(expected - actual) > 0.000001) {
        mismatches.push(`${sheet.name}!${rowIndex} ${staffName}: generated ${(actual * 100).toFixed(2)}%, MR ${(expected * 100).toFixed(2)}%`)
      }
      if (mismatches.length >= 20) {
        break
      }
    }
    if (mismatches.length >= 20) {
      break
    }
  }

  if (mismatches.length) {
    throw new Error(`PI Completion Rate does not match MR source:\n${mismatches.join('\n')}`)
  }
}

function staffWorkbookDataSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet | undefined {
  return workbook.worksheets.find((sheet) => {
    const headers = Array.from({ length: sheet.getRow(1).cellCount }, (_, index) => normalizeText(sheet.getCell(1, index + 1).value))
    return headers.includes('Staff') && headers.includes('PI Completion Rate')
  })
}

function staffPiCompletionColumn(sheet: ExcelJS.Worksheet): number {
  const headers = Array.from({ length: sheet.getRow(1).cellCount }, (_, index) => normalizeText(sheet.getCell(1, index + 1).value))
  return headers.indexOf('PI Completion Rate') + 1
}

function staffPiCompletionRowIndices(sheet: ExcelJS.Worksheet): number[] {
  const piColumn = staffPiCompletionColumn(sheet)
  if (!piColumn) {
    return []
  }
  const rows: number[] = []
  for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex += 1) {
    if (normalizeStaffPiCompletionRate(sheet.getCell(rowIndex, piColumn).value) >= 1) {
      rows.push(rowIndex)
    }
  }
  return rows
}

function countStaffPiCompletionRows(workbook: ExcelJS.Workbook): number {
  const sheet = staffWorkbookDataSheet(workbook)
  return sheet ? staffPiCompletionRowIndices(sheet).length : 0
}

function staffPiCompletionSnapshotOverlay(sheet: ExcelJS.Worksheet, rowIndices: number[]): (cell: ExcelJS.Cell, row: number, col: number) => SnapshotOverlay {
  const dataRows = Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) => index + 2)
  const piValues = dataRows.map((rowIndex) => normalizeNumber(sheet.getCell(rowIndex, 6).value)).filter(Number.isFinite)
  const pitValues = dataRows.map((rowIndex) => normalizeNumber(sheet.getCell(rowIndex, 8).value)).filter(Number.isFinite)
  const piCompletionValues = dataRows
    .map((rowIndex) => normalizeStaffPiCompletionRate(sheet.getCell(rowIndex, 9).value))
    .filter(Number.isFinite)
  const piMax = Math.max(...piValues, 1)
  const pitMax = Math.max(...pitValues, 1)
  const piCompletionMin = Math.min(...piCompletionValues, 1)
  const piCompletionMax = Math.max(...piCompletionValues, 1)

  return (cell, row, col) => {
    if (row === 1) {
      return {}
    }
    if (col === 6) {
      return { dataBarRatio: normalizeNumber(cell.value) / piMax, dataBarColor: '#638EC6' }
    }
    if (col === 8) {
      return { dataBarRatio: normalizeNumber(cell.value) / pitMax, dataBarColor: '#638EC6' }
    }
    if (col === 9) {
      return { background: threeColorScale(normalizeStaffPiCompletionRate(cell.value), piCompletionMin, piCompletionMax) }
    }
    return {}
  }
}

function buildStaffPiCompletionSnapshots(sheet: ExcelJS.Worksheet): string[] {
  const rowIndices = staffPiCompletionRowIndices(sheet)
  if (!rowIndices.length) {
    return []
  }

  const chunks = rowIndices.length > 36
    ? [rowIndices.slice(0, Math.ceil(rowIndices.length / 2)), rowIndices.slice(Math.ceil(rowIndices.length / 2))]
    : [rowIndices]

  return chunks.map((chunk) => renderWorksheetRangeSvg(sheet, {
    startRow: 1,
    endRow: Math.max(...chunk),
    startCol: 1,
    endCol: 9,
    rowIndices: [1, ...chunk],
    colIndices: [1, 2, 6, 8, 9],
    colWidthOverrides: new Map([
      [1, 58],
      [2, 58],
      [6, 16],
      [8, 16],
      [9, 32]
    ]),
    defaultColWidth: 15,
    defaultRowHeight: 68,
    fontScale: 4,
    fontScaleOverrides: new Map([[1, 5]]),
    rowFontScaleOverrides: new Map([[1, 5]]),
    rowHeightOverrides: new Map([[1, 122]]),
    forceBold: true,
    gridStrokeColor: '#57606A',
    gridStrokeWidth: 2,
    cellOverlay: staffPiCompletionSnapshotOverlay(sheet, rowIndices)
  }))
}

function worksheetHeaderColumn(sheet: ExcelJS.Worksheet, header: string): number {
  const headers = Array.from({ length: sheet.getRow(1).cellCount }, (_, index) => normalizeText(sheet.getCell(1, index + 1).value))
  return headers.indexOf(header) + 1
}

function buildStaffOwnerSnapshotOverlay(sheet: ExcelJS.Worksheet): (cell: ExcelJS.Cell, row: number, col: number) => SnapshotOverlay {
  const piCompletionCol = worksheetHeaderColumn(sheet, 'PI Completion Rate')
  const siOpenCol = worksheetHeaderColumn(sheet, 'SI Open')
  const dataRows = Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) => index + 2)
  const piCompletionValues = piCompletionCol
    ? dataRows.map((rowIndex) => normalizeStaffPiCompletionRate(sheet.getCell(rowIndex, piCompletionCol).value)).filter(Number.isFinite)
    : []
  const siOpenValues = siOpenCol
    ? dataRows.map((rowIndex) => normalizeNumber(sheet.getCell(rowIndex, siOpenCol).value)).filter(Number.isFinite)
    : []
  const piCompletionMin = Math.min(...piCompletionValues, 1)
  const piCompletionMax = Math.max(...piCompletionValues, 1)
  const siOpenMax = Math.max(...siOpenValues, 1)

  return (cell, row, col) => {
    if (row === 1) {
      return {}
    }
    if (col === piCompletionCol) {
      return { background: threeColorScale(normalizeStaffPiCompletionRate(cell.value), piCompletionMin, piCompletionMax) }
    }
    if (col === siOpenCol) {
      return {
        dataBarRatio: normalizeNumber(cell.value) / siOpenMax,
        dataBarColor: '#638EC6'
      }
    }
    return {}
  }
}

function buildStaffOwnerSnapshot(sheet: ExcelJS.Worksheet): string {
  const ownerCol = worksheetHeaderColumn(sheet, 'PI (article owner)')
  const piCompletionCol = worksheetHeaderColumn(sheet, 'PI Completion Rate')
  const siOpenCol = worksheetHeaderColumn(sheet, 'SI Open')
  if (!ownerCol || !piCompletionCol || !siOpenCol) {
    return ''
  }

  const rowIndices = Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) => index + 2)
    .filter((rowIndex) => normalizeNumber(sheet.getCell(rowIndex, ownerCol).value) > 10)
    .sort((left, right) => normalizeNumber(sheet.getCell(right, ownerCol).value) - normalizeNumber(sheet.getCell(left, ownerCol).value))
  if (!rowIndices.length) {
    return ''
  }

  const dataRows = Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) => index + 2)
  const piCompletionValues = dataRows.map((rowIndex) => normalizeStaffPiCompletionRate(sheet.getCell(rowIndex, piCompletionCol).value)).filter(Number.isFinite)
  const siOpenValues = dataRows.map((rowIndex) => normalizeNumber(sheet.getCell(rowIndex, siOpenCol).value)).filter(Number.isFinite)
  const piCompletionMin = Math.min(...piCompletionValues, 1)
  const piCompletionMax = Math.max(...piCompletionValues, 1)
  const siOpenMax = Math.max(...siOpenValues, 1)
  const headerHeight = 42
  const rowHeight = 44
  const targetWidth = Math.round((headerHeight + rowIndices.length * rowHeight + 12) * 1.317)
  const baseColumns = [
    { key: 'staff', label: 'Staff', width: 172 },
    { key: 'journal', label: 'Journal', width: 252 },
    { key: 'ownerPi', label: 'PI (article owner)', width: 104 },
    { key: 'piCompletion', label: 'PI Completion Rate', width: 118 },
    { key: 'siOpen', label: 'SI Open', width: 57 }
  ]
  const widthScale = Math.max(1, targetWidth / baseColumns.reduce((sum, column) => sum + column.width, 0))

  return renderTableSvg({
    columns: baseColumns.map((column) => ({ ...column, width: Math.round(column.width * widthScale) })),
    rows: rowIndices.map((rowIndex) => {
      const piCompletion = normalizeStaffPiCompletionRate(sheet.getCell(rowIndex, piCompletionCol).value)
      const siOpen = normalizeNumber(sheet.getCell(rowIndex, siOpenCol).value)
      return {
        staff: {
          value: normalizeText(sheet.getCell(rowIndex, 1).value),
          fontSize: 20,
          weight: 700 as const
        },
        journal: {
          value: normalizeText(sheet.getCell(rowIndex, 2).value),
          fontSize: 20,
          weight: 700 as const
        },
        ownerPi: {
          value: String(Math.round(normalizeNumber(sheet.getCell(rowIndex, ownerCol).value))),
          fontSize: 20,
          weight: 700 as const
        },
        piCompletion: {
          value: `${(piCompletion * 100).toFixed(2)}%`,
          background: threeColorScale(piCompletion, piCompletionMin, piCompletionMax),
          fontSize: 20,
          weight: 700 as const
        },
        siOpen: {
          value: String(Math.round(siOpen)),
          fontSize: 20,
          weight: 700 as const,
          barRatio: siOpen / siOpenMax,
          barColor: '#638EC6'
        }
      }
    }),
    headerHeight,
    rowHeight
  })
}

function buildMrSiPublTopSnapshot(editorRows: GenericRow[]): string {
  const rows = [...editorRows]
    .sort((left, right) => normalizeNumber(right['SI Publ']) - normalizeNumber(left['SI Publ']))
    .slice(0, 30)
  const maxSiPubl = Math.max(...rows.map((row) => normalizeNumber(row['SI Publ'])), 1)

  return renderTableSvg({
    columns: [
      { key: 'staff', label: 'Staff', width: 130 },
      { key: 'section', label: 'Section', width: 310 },
      { key: 'siPubl', label: 'SI Publ', width: 80 }
    ],
    rows: rows.map((row) => {
      const siPubl = normalizeNumber(row['SI Publ'])
      const isHealth = normalizeText(row.Section) === 'Section Health'
      const background = isHealth ? '#12542f' : undefined
      const color = isHealth ? '#ffffff' : undefined
      return {
        staff: {
          value: compactName(normalizeText(row.Staff)),
          background,
          color,
          fontSize: 13,
          weight: 700 as const
        },
        section: {
          value: normalizeText(row.Section),
          background,
          color,
          fontSize: 13,
          weight: 700 as const
        },
        siPubl: {
          value: String(Math.round(siPubl)),
          background,
          color,
          fontSize: 13,
          weight: 700 as const,
          barRatio: siPubl / maxSiPubl,
          barColor: isHealth ? '#8fd19e' : '#63BE7B'
        }
      }
    }),
    headerHeight: 36,
    rowHeight: 25
  })
}

function buildStaffOwnerSnapshots(sheet: ExcelJS.Worksheet, editorRows: GenericRow[]): string[] {
  return [buildStaffOwnerSnapshot(sheet), buildMrSiPublTopSnapshot(editorRows)]
}

function buildStaffSmeSiSetupSnapshot(sheet: ExcelJS.Worksheet): string {
  const siSetupCol = worksheetHeaderColumn(sheet, 'SI Setup')
  const siOpenCol = worksheetHeaderColumn(sheet, 'SI Open')
  if (!siSetupCol || !siOpenCol) {
    return ''
  }

  const rowIndices = Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) => index + 2)
    .filter((rowIndex) => normalizeNumber(sheet.getCell(rowIndex, siSetupCol).value) > 3)
    .sort((left, right) => normalizeNumber(sheet.getCell(right, siSetupCol).value) - normalizeNumber(sheet.getCell(left, siSetupCol).value))
  if (!rowIndices.length) {
    return ''
  }

  const siSetupMax = Math.max(...Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) => normalizeNumber(sheet.getCell(index + 2, siSetupCol).value)), 1)
  const siOpenMax = Math.max(...Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) => normalizeNumber(sheet.getCell(index + 2, siOpenCol).value)), 1)
  const headerHeight = 42
  const rowHeight = 44
  const targetWidth = Math.round((headerHeight + rowIndices.length * rowHeight + 12) * 0.773)
  const baseColumns = [
    { key: 'staff', label: 'Staff', width: 170 },
    { key: 'journal', label: 'Journal', width: 270 },
    { key: 'siSetup', label: 'SI Setup', width: 110 },
    { key: 'siOpen', label: 'SI Open', width: 110 }
  ]
  const widthScale = targetWidth / baseColumns.reduce((sum, column) => sum + column.width, 0)

  return renderTableSvg({
    columns: baseColumns.map((column) => ({ ...column, width: Math.round(column.width * widthScale) })),
    rows: rowIndices.map((rowIndex) => {
      const siSetup = normalizeNumber(sheet.getCell(rowIndex, siSetupCol).value)
      const siOpen = normalizeNumber(sheet.getCell(rowIndex, siOpenCol).value)
      return {
        staff: {
          value: normalizeText(sheet.getCell(rowIndex, 1).value),
          fontSize: 8,
          weight: 700 as const
        },
        journal: {
          value: normalizeText(sheet.getCell(rowIndex, 2).value),
          fontSize: 8,
          weight: 700 as const
        },
        siSetup: {
          value: String(Math.round(siSetup)),
          fontSize: 8,
          weight: 700 as const,
          barRatio: siSetup / siSetupMax,
          barColor: '#63BE7B'
        },
        siOpen: {
          value: String(Math.round(siOpen)),
          fontSize: 8,
          weight: 700 as const,
          barRatio: siOpen / siOpenMax,
          barColor: '#638EC6'
        }
      }
    }),
    headerHeight,
    rowHeight
  })
}

function buildMrSiSetupTopSnapshot(editorRows: GenericRow[]): string {
  const rows = [...editorRows]
    .sort((left, right) => normalizeNumber(right['Special Issues Setup']) - normalizeNumber(left['Special Issues Setup']))
    .slice(0, 30)
  const maxSiSetup = Math.max(...rows.map((row) => normalizeNumber(row['Special Issues Setup'])), 1)
  const maxSiOpen = Math.max(...rows.map((row) => normalizeNumber(row['Special Issues Open'])), 1)

  return renderTableSvg({
    columns: [
      { key: 'staff', label: 'Staff', width: 124 },
      { key: 'section', label: 'Section', width: 238 },
      { key: 'siSetup', label: 'SI Setup', width: 78 },
      { key: 'siOpen', label: 'SI Open', width: 78 }
    ],
    rows: rows.map((row) => {
      const siSetup = normalizeNumber(row['Special Issues Setup'])
      const siOpen = normalizeNumber(row['Special Issues Open'])
      const isHealth = normalizeText(row.Section) === 'Section Health'
      const background = isHealth ? '#12542f' : undefined
      const color = isHealth ? '#ffffff' : undefined
      return {
        staff: {
          value: compactName(normalizeText(row.Staff)),
          background,
          color,
          fontSize: 12,
          weight: 700 as const
        },
        section: {
          value: normalizeText(row.Section),
          background,
          color,
          fontSize: 12,
          weight: 700 as const
        },
        siSetup: {
          value: String(Math.round(siSetup)),
          background,
          color,
          fontSize: 12,
          weight: 700 as const,
          barRatio: siSetup / maxSiSetup,
          barColor: isHealth ? '#8fd19e' : '#63BE7B'
        },
        siOpen: {
          value: String(Math.round(siOpen)),
          background,
          color,
          fontSize: 12,
          weight: 700 as const,
          barRatio: siOpen / maxSiOpen,
          barColor: isHealth ? '#8bb8ff' : '#638EC6'
        }
      }
    }),
    headerHeight: 36,
    rowHeight: 25
  })
}

function buildStaffSiSetupSnapshots(smeSheet: ExcelJS.Worksheet, editorRows: GenericRow[]): string[] {
  return [buildStaffSmeSiSetupSnapshot(smeSheet), buildMrSiSetupTopSnapshot(editorRows)]
}

function buildStaffSmeSiSubSnapshot(sheet: ExcelJS.Worksheet): string {
  const siSubCol = worksheetHeaderColumn(sheet, 'SI Sub.')
  const siOpenCol = worksheetHeaderColumn(sheet, 'SI Open')
  if (!siSubCol || !siOpenCol) {
    return ''
  }

  const rowIndices = Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) => index + 2)
    .filter((rowIndex) => normalizeNumber(sheet.getCell(rowIndex, siSubCol).value) > 15)
    .sort((left, right) => normalizeNumber(sheet.getCell(right, siSubCol).value) - normalizeNumber(sheet.getCell(left, siSubCol).value))
  if (!rowIndices.length) {
    return ''
  }

  const dataRows = Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) => index + 2)
  const siSubMax = Math.max(...dataRows.map((rowIndex) => normalizeNumber(sheet.getCell(rowIndex, siSubCol).value)), 1)
  const siOpenMax = Math.max(...dataRows.map((rowIndex) => normalizeNumber(sheet.getCell(rowIndex, siOpenCol).value)), 1)
  return renderTableSvg({
    columns: [
      { key: 'staff', label: 'Staff', width: 148 },
      { key: 'journal', label: 'Journal', width: 220 },
      { key: 'siSub', label: 'SI Sub.', width: 82 },
      { key: 'siOpen', label: 'SI Open', width: 82 }
    ],
    rows: rowIndices.map((rowIndex) => {
      const siSub = normalizeNumber(sheet.getCell(rowIndex, siSubCol).value)
      const siOpen = normalizeNumber(sheet.getCell(rowIndex, siOpenCol).value)
      return {
        staff: {
          value: normalizeText(sheet.getCell(rowIndex, 1).value),
          fontSize: 11,
          weight: 700 as const
        },
        journal: {
          value: normalizeText(sheet.getCell(rowIndex, 2).value),
          fontSize: 11,
          weight: 700 as const
        },
        siSub: {
          value: String(Math.round(siSub)),
          fontSize: 11,
          weight: 700 as const,
          barRatio: siSub / siSubMax,
          barColor: '#63BE7B'
        },
        siOpen: {
          value: String(Math.round(siOpen)),
          fontSize: 11,
          weight: 700 as const,
          barRatio: siOpen / siOpenMax,
          barColor: '#638EC6'
        }
      }
    }),
    headerHeight: 36,
    rowHeight: 26
  })
}

function buildMrSiSubTopSnapshot(editorRows: GenericRow[]): string {
  const rows = [...editorRows]
    .sort((left, right) => normalizeNumber(right['SI Sub.']) - normalizeNumber(left['SI Sub.']))
    .slice(0, 30)
  const maxSiSub = Math.max(...rows.map((row) => normalizeNumber(row['SI Sub.'])), 1)
  const maxSiOpen = Math.max(...rows.map((row) => normalizeNumber(row['Special Issues Open'])), 1)

  return renderTableSvg({
    columns: [
      { key: 'staff', label: 'Staff', width: 140 },
      { key: 'section', label: 'Section', width: 250 },
      { key: 'siSub', label: 'SI Sub.', width: 82 },
      { key: 'siOpen', label: 'SI Open', width: 82 }
    ],
    rows: rows.map((row) => {
      const siSub = normalizeNumber(row['SI Sub.'])
      const siOpen = normalizeNumber(row['Special Issues Open'])
      const isHealth = normalizeText(row.Section) === 'Section Health'
      const background = isHealth ? '#12542f' : undefined
      const color = isHealth ? '#ffffff' : undefined
      return {
        staff: {
          value: compactName(normalizeText(row.Staff)),
          background,
          color,
          fontSize: 12,
          weight: 700 as const
        },
        section: {
          value: normalizeText(row.Section),
          background,
          color,
          fontSize: 12,
          weight: 700 as const
        },
        siSub: {
          value: String(Math.round(siSub)),
          background,
          color,
          fontSize: 12,
          weight: 700 as const,
          barRatio: siSub / maxSiSub,
          barColor: isHealth ? '#8fd19e' : '#63BE7B'
        },
        siOpen: {
          value: String(Math.round(siOpen)),
          background,
          color,
          fontSize: 12,
          weight: 700 as const,
          barRatio: siOpen / maxSiOpen,
          barColor: isHealth ? '#8bb8ff' : '#638EC6'
        }
      }
    }),
    headerHeight: 36,
    rowHeight: 25
  })
}

function buildStaffSiSubSnapshots(smeSheet: ExcelJS.Worksheet, editorRows: GenericRow[]): string[] {
  return [buildStaffSmeSiSubSnapshot(smeSheet), buildMrSiSubTopSnapshot(editorRows)]
}

function buildStaffAePublSnapshot(sheet: ExcelJS.Worksheet): string {
  const publCol = worksheetHeaderColumn(sheet, 'Publ')
  const mptCol = worksheetHeaderColumn(sheet, 'MPT')
  const tfdCol = worksheetHeaderColumn(sheet, 'TFD')
  if (!publCol || !mptCol || !tfdCol) {
    return ''
  }

  const rowIndices = Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) => index + 2)
    .filter((rowIndex) => normalizeNumber(sheet.getCell(rowIndex, publCol).value) > 15)
    .sort((left, right) => normalizeNumber(sheet.getCell(right, publCol).value) - normalizeNumber(sheet.getCell(left, publCol).value))
  if (!rowIndices.length) {
    return ''
  }

  const dataRows = Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) => index + 2)
  const publMax = Math.max(...dataRows.map((rowIndex) => normalizeNumber(sheet.getCell(rowIndex, publCol).value)), 1)
  return renderTableSvg({
    columns: [
      { key: 'staff', label: 'Staff', width: 140 },
      { key: 'journal', label: 'Journal', width: 220 },
      { key: 'publ', label: 'Publ', width: 70 },
      { key: 'mpt', label: 'MPT', width: 60 },
      { key: 'tfd', label: 'TFD', width: 60 }
    ],
    rows: rowIndices.map((rowIndex) => {
      const publ = normalizeNumber(sheet.getCell(rowIndex, publCol).value)
      return {
        staff: {
          value: normalizeText(sheet.getCell(rowIndex, 1).value),
          fontSize: 11,
          weight: 700 as const
        },
        journal: {
          value: normalizeText(sheet.getCell(rowIndex, 2).value),
          fontSize: 11,
          weight: 700 as const
        },
        publ: {
          value: String(Math.round(publ)),
          fontSize: 11,
          weight: 700 as const,
          barRatio: publ / publMax,
          barColor: '#63BE7B'
        },
        mpt: {
          value: String(Math.round(normalizeNumber(sheet.getCell(rowIndex, mptCol).value))),
          fontSize: 11,
          weight: 700 as const
        },
        tfd: {
          value: String(Math.round(normalizeNumber(sheet.getCell(rowIndex, tfdCol).value))),
          fontSize: 11,
          weight: 700 as const
        }
      }
    }),
    headerHeight: 36,
    rowHeight: 25
  })
}

function buildMrPublTopSnapshot(editorRows: GenericRow[]): string {
  const rows = [...editorRows]
    .sort((left, right) => normalizeNumber(right.Publ) - normalizeNumber(left.Publ))
    .slice(0, 30)
  const maxPubl = Math.max(...rows.map((row) => normalizeNumber(row.Publ)), 1)

  return renderTableSvg({
    columns: [
      { key: 'staff', label: 'Staff', width: 132 },
      { key: 'section', label: 'Section', width: 290 },
      { key: 'publ', label: 'Publ', width: 78 }
    ],
    rows: rows.map((row) => {
      const publ = normalizeNumber(row.Publ)
      const isHealth = normalizeText(row.Section) === 'Section Health'
      const background = isHealth ? '#12542f' : undefined
      const color = isHealth ? '#ffffff' : undefined
      return {
        staff: {
          value: compactName(normalizeText(row.Staff)),
          background,
          color,
          fontSize: 12,
          weight: 700 as const
        },
        section: {
          value: normalizeText(row.Section),
          background,
          color,
          fontSize: 12,
          weight: 700 as const
        },
        publ: {
          value: String(Math.round(publ)),
          background,
          color,
          fontSize: 12,
          weight: 700 as const,
          barRatio: publ / maxPubl,
          barColor: isHealth ? '#8fd19e' : '#63BE7B'
        }
      }
    }),
    headerHeight: 36,
    rowHeight: 25
  })
}

function buildStaffAePublSnapshots(aeSheet: ExcelJS.Worksheet, editorRows: GenericRow[]): string[] {
  return [buildStaffAePublSnapshot(aeSheet), buildMrPublTopSnapshot(editorRows)]
}

function addStaffCommonConditionalFormatting(sheet: ExcelJS.Worksheet, startRow: number, endRow: number): void {
  if (endRow < startRow) {
    return
  }
  sheet.addConditionalFormatting({
    ref: `I${startRow}:I${endRow}`,
    rules: [{
      type: 'colorScale',
      priority: 1,
      cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
      color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }]
    }]
  })
}

function addStaffReferenceColorScale(sheet: ExcelJS.Worksheet, column: string, startRow: number, endRow: number, priority: number): void {
  if (endRow < startRow) {
    return
  }
  sheet.addConditionalFormatting({
    ref: `${column}${startRow}:${column}${endRow}`,
    rules: [{
      type: 'colorScale',
      priority,
      cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
      color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }]
    }]
  })
}

function addAeConditionalFormatting(sheet: ExcelJS.Worksheet, startRow: number, endRow: number): void {
  addStaffCommonConditionalFormatting(sheet, startRow, endRow)
  addDataBar(sheet, 'J', startRow, endRow, 3)
  addExpressionFillRule(sheet, 'K', startRow, endRow, `K${startRow}>60`, 'FF7D7D', 4)
  addExpressionFillRule(sheet, 'K', startRow, endRow, `AND(K${startRow}>40,K${startRow}<=60)`, 'FFC000', 5)
  addExpressionFillRule(sheet, 'K', startRow, endRow, `AND(K${startRow}>35,K${startRow}<=40)`, 'D8E4BC', 6)
  addExpressionFillRule(sheet, 'K', startRow, endRow, `K${startRow}<=35`, '00B050', 7)
  addExpressionFillRule(sheet, 'L', startRow, endRow, `L${startRow}>=20`, 'FF7D7D', 8)
  addExpressionFillRule(sheet, 'L', startRow, endRow, `AND(L${startRow}>=15,L${startRow}<20)`, 'FFC000', 9)
  addExpressionFillRule(sheet, 'L', startRow, endRow, `AND(L${startRow}>10,L${startRow}<15)`, 'D8E4BC', 10)
  addExpressionFillRule(sheet, 'L', startRow, endRow, `L${startRow}<=10`, '00B050', 11)
  addExpressionFillRule(sheet, 'M', startRow, endRow, `M${startRow}>=1`, '00B050', 12)
  addCellFillRule(sheet, 'P', startRow, endRow, 'greaterThan', '3', 'FF7D7D', 13)
  addStaffReferenceColorScale(sheet, 'Q', startRow, endRow, 14)
}

function addSmeConditionalFormatting(sheet: ExcelJS.Worksheet, startRow: number, endRow: number): void {
  addStaffCommonConditionalFormatting(sheet, startRow, endRow)
  for (const [index, column] of ['J', 'K', 'L', 'N', 'X', 'Y', 'Z', 'AA'].entries()) {
    addDataBar(sheet, column, startRow, endRow, 20 + index)
  }
  addStaffReferenceColorScale(sheet, 'M', startRow, endRow, 30)
  addExpressionFillRule(sheet, 'O', startRow, endRow, `O${startRow}>=15`, '00B050', 31)
  addStaffReferenceColorScale(sheet, 'Q', startRow, endRow, 32)
  addExpressionFillRule(sheet, 'R', startRow, endRow, `R${startRow}>=2`, '00B050', 33)
  addStaffReferenceColorScale(sheet, 'U', startRow, endRow, 34)
}

function staffColumn(headers: string[], header: string): string | null {
  const index = headers.indexOf(header)
  return index === -1 ? null : columnLetter(index + 1)
}

function addAllStaffConditionalFormatting(sheet: ExcelJS.Worksheet, headers: string[], startRow: number, endRow: number): void {
  if (endRow < startRow) {
    return
  }
  addStaffCommonConditionalFormatting(sheet, startRow, endRow)

  const column = (header: string) => staffColumn(headers, header)
  const addBar = (header: string, priority: number) => {
    const col = column(header)
    if (col) {
      addDataBar(sheet, col, startRow, endRow, priority)
    }
  }
  const addScale = (header: string, priority: number) => {
    const col = column(header)
    if (col) {
      addStaffReferenceColorScale(sheet, col, startRow, endRow, priority)
    }
  }

  const publCol = column('Publ')
  if (publCol) {
    addDataBar(sheet, publCol, startRow, endRow, 40)
  }
  const mptCol = column('MPT')
  if (mptCol) {
    addExpressionFillRule(sheet, mptCol, startRow, endRow, `${mptCol}${startRow}>60`, 'FF7D7D', 41)
    addExpressionFillRule(sheet, mptCol, startRow, endRow, `AND(${mptCol}${startRow}>40,${mptCol}${startRow}<=60)`, 'FFC000', 42)
    addExpressionFillRule(sheet, mptCol, startRow, endRow, `AND(${mptCol}${startRow}>35,${mptCol}${startRow}<=40)`, 'D8E4BC', 43)
    addExpressionFillRule(sheet, mptCol, startRow, endRow, `${mptCol}${startRow}<=35`, '00B050', 44)
  }
  const tfdCol = column('TFD')
  if (tfdCol) {
    addExpressionFillRule(sheet, tfdCol, startRow, endRow, `${tfdCol}${startRow}>=20`, 'FF7D7D', 45)
    addExpressionFillRule(sheet, tfdCol, startRow, endRow, `AND(${tfdCol}${startRow}>=15,${tfdCol}${startRow}<20)`, 'FFC000', 46)
    addExpressionFillRule(sheet, tfdCol, startRow, endRow, `AND(${tfdCol}${startRow}>10,${tfdCol}${startRow}<15)`, 'D8E4BC', 47)
    addExpressionFillRule(sheet, tfdCol, startRow, endRow, `${tfdCol}${startRow}<=10`, '00B050', 48)
  }
  const reviewerCol = column('New Reviewer With Report (Score>2)')
  if (reviewerCol) {
    addExpressionFillRule(sheet, reviewerCol, startRow, endRow, `${reviewerCol}${startRow}>=1`, '00B050', 49)
  }
  const aptCol = column('Acceptance to Publication Time')
  if (aptCol) {
    addCellFillRule(sheet, aptCol, startRow, endRow, 'greaterThan', '3', 'FF7D7D', 50)
  }
  addScale('Under Process', 51)

  for (const [index, header] of ['SI Sub.', 'SI Publ', 'Regular Publ', 'SI Open', 'Susy CFP', 'CFP Team CFP', 'Mmailer-Regular CFP', 'Mmailer-FCFP'].entries()) {
    addBar(header, 60 + index)
  }
  addScale('SI Setup', 70)
  const siUnderCol = column('SI Under Process.')
  if (siUnderCol) {
    addExpressionFillRule(sheet, siUnderCol, startRow, endRow, `${siUnderCol}${startRow}>=15`, '00B050', 71)
  }
  addScale('GE Agree Rate', 72)
  const acceptedSiCol = column('New Accepted SI')
  if (acceptedSiCol) {
    addExpressionFillRule(sheet, acceptedSiCol, startRow, endRow, `${acceptedSiCol}${startRow}>=2`, '00B050', 73)
  }
  addScale('Closed Special Issue Average Publication', 74)
}

function buildForceAeStaffSet(forceAeStaff: string[]): Set<string> {
  return new Set(forceAeStaff.flatMap((name) => personLookupKeys(name)))
}

function isForceAeStaff(row: GenericRow, forceAeStaff: Set<string>): boolean {
  const name = compactName(normalizeText(row.Staff))
  return personLookupKeys(name).some((key) => forceAeStaff.has(key))
}

function populateStaffWorkbook(workbook: ExcelJS.Workbook, editorRows: GenericRow[], journalMap: Map<string, string[]>, forceAeStaffNames: string[]): void {
  for (const worksheet of [...workbook.worksheets]) {
    workbook.removeWorksheet(worksheet.id)
  }
  const allSheet = resetWorksheet(workbook, '人员数据')
  const aeSheet = resetWorksheet(workbook, 'AE')
  const smeSheet = resetWorksheet(workbook, 'SME')
  const forceAeStaff = buildForceAeStaffSet(forceAeStaffNames)
  const activeHealthEditors = editorRows.filter((row) =>
    normalizeText(row.Section) === 'Section Health' && normalizeText(row['Is Active']) === 'Activate'
  )
  const allRows = sortStaffRowsByPiCompletion(activeHealthEditors)

  const aeRows = sortStaffRowsByPiCompletion(activeHealthEditors.filter((row) => {
    return normalizeText(row['Editor Role']) === 'Assistant Editor' || isForceAeStaff(row, forceAeStaff)
  }))
  const smeRows = sortStaffRowsByPiCompletion(activeHealthEditors.filter((row) => {
    const role = normalizeText(row['Editor Role'])
    return !isForceAeStaff(row, forceAeStaff) && !isScieManagingEditor(row, journalMap) && (
      role.includes('Section Managing Editor') ||
      role.includes('Managing Editor') ||
      role.includes('Journal Development Editor')
    )
  }))

  writeStaffSheet(allSheet, ALL_STAFF_HEADERS, allRows, journalMap)
  writeStaffSheet(aeSheet, AE_STAFF_HEADERS, aeRows, journalMap)
  writeStaffSheet(smeSheet, SME_STAFF_HEADERS, smeRows, journalMap)
  verifyStaffPiCompletionValues(allSheet, ALL_STAFF_HEADERS, allRows)
  verifyStaffPiCompletionValues(aeSheet, AE_STAFF_HEADERS, aeRows)
  verifyStaffPiCompletionValues(smeSheet, SME_STAFF_HEADERS, smeRows)
  addAllStaffConditionalFormatting(allSheet, ALL_STAFF_HEADERS, 2, allRows.length + 1)
  addAeConditionalFormatting(aeSheet, 2, aeRows.length + 1)
  addSmeConditionalFormatting(smeSheet, 2, smeRows.length + 1)
}

function buildSpotlights(focusSeries: FocusJournalSeries[], completionRows: CompletionRow[], journalMap: Map<string, string[]>): JournalSpotlight[] {
  return focusSeries.map((series) => {
    const matchName = series.displayName
    const completionRow = completionRows.find((row) => row.journal === matchName)
    const editors = Array.from(
      new Set(
        Array.from(journalMap.entries())
          .filter(([, journals]) => journals.includes(matchName))
          .map(([name]) => name)
      )
    ).slice(0, 6)

    return {
      displayName: matchName,
      editors,
      publicationTcr: completionRow?.publicationTcr ?? 0,
      revenueTcr: completionRow?.revenueTcr ?? 0,
      waiverRate2026: completionRow?.waiverRate2026 ?? 0
    }
  })
}

export async function runPipeline(event: IpcMainInvokeEvent, input: PipelineInput): Promise<PipelineResult> {
  emitProgress(event, 'bootstrap', 'Loading input files')
  await fs.mkdir(input.paths.outputDir, { recursive: true })

  const mrWorkbook = new ExcelJS.Workbook()
  const monthlyWorkbook = new ExcelJS.Workbook()
  const staffWorkbook = new ExcelJS.Workbook()
  const journalWorkbook = new ExcelJS.Workbook()

  await mrWorkbook.xlsx.readFile(input.paths.mrWorkbook)
  await monthlyWorkbook.xlsx.readFile(input.paths.monthlyTemplate)
  await staffWorkbook.xlsx.readFile(input.paths.staffTemplate)
  await journalWorkbook.xlsx.readFile(input.paths.editorsJournals)

  const journalRows = sheetToObjects(getWorksheetOrThrow(mrWorkbook, 'Journals'))
  const editorRows = sheetToObjects(getWorksheetOrThrow(mrWorkbook, 'Editors'))
  const context = parseReportContext(input.paths.mrWorkbook, journalRows)

  emitProgress(event, 'excel', 'Updating Office and completion sheets')
  const officeSheet = getWorksheetOrThrow(monthlyWorkbook, 'Office')
  const officeRows = orderOfficeRowsByTemplate(officeSheet, buildOfficeRows(journalRows, context))
  const completionRows = buildCompletionRows(journalRows, context)
  const completionSheet = getWorksheetOrThrow(monthlyWorkbook, '\u5b8c\u6210\u7387')
  populateOfficeSheet(officeSheet, officeRows, context)
  populateCompletionSheet(completionSheet, completionRows, context)

  emitProgress(event, 'excel', 'Updating department summary and focus journal sheets')
  const departmentMetrics = computeDepartmentMetrics(journalRows, context, input.summaryOverrides)
  const departmentSheet = getWorksheetOrThrow(monthlyWorkbook, '\u79d1\u5ba4\u6570\u636e')
  populateDepartmentSheet(departmentSheet, departmentMetrics, context)

  const focusOrder = ensureArrayLength(input.focusJournalOrder, 5, () => 'Foods')
  const focusSeries: FocusJournalSeries[] = []
  for (const sheetName of focusOrder) {
    const worksheet = getWorksheetOrThrow(monthlyWorkbook, sheetName)
    const matchName = sheetName === 'BS' ? 'Brain Sciences' : sheetName
    const sourceRow = journalRows.find((row) => normalizeText(row.Journal) === matchName)
    focusSeries.push(updateFocusJournalSheet(worksheet, context, sourceRow))
  }

  emitProgress(event, 'staff', 'Updating staff workbook')
  const journalMap = buildJournalMap(journalWorkbook)
  mergeStaffTemplateJournalMap(staffWorkbook, journalMap)
  const previousStaffPiCompletionCount = countStaffPiCompletionRows(staffWorkbook)
  populateStaffWorkbook(staffWorkbook, editorRows, journalMap, input.forceAeStaff ?? [])
  const staffDataSheet = staffWorkbookDataSheet(staffWorkbook)
  const staffAeSheet = staffWorkbook.getWorksheet('AE')
  const staffSmeSheet = staffWorkbook.getWorksheet('SME')
  const staffPiCompletionSvgs = staffDataSheet ? buildStaffPiCompletionSnapshots(staffDataSheet) : []
  const staffOwnerSvgs = staffSmeSheet ? buildStaffOwnerSnapshots(staffSmeSheet, editorRows) : []
  const staffSiSetupSvgs = staffSmeSheet ? buildStaffSiSetupSnapshots(staffSmeSheet, editorRows) : []
  const staffSiSubSvgs = staffSmeSheet ? buildStaffSiSubSnapshots(staffSmeSheet, editorRows) : []
  const staffAePublSvgs = staffAeSheet ? buildStaffAePublSnapshots(staffAeSheet, editorRows) : []
  const currentStaffPiCompletionCount = staffDataSheet ? staffPiCompletionRowIndices(staffDataSheet).length : 0
  const staffPiCompletionSummary: StaffPiCompletionSummary = {
    currentCount: currentStaffPiCompletionCount,
    previousCount: previousStaffPiCompletionCount,
    delta: currentStaffPiCompletionCount - previousStaffPiCompletionCount
  }

  const monthlyOutputPath = buildOutputPath(input.paths.outputDir, 'monthly-data-generated', context.reportKey, 'xlsx')
  const staffOutputPath = buildOutputPath(input.paths.outputDir, 'staff-data-generated', context.reportKey, 'xlsx')
  const pptSnapshots = buildPptSnapshots(officeSheet, completionSheet, completionRows, context)
  pptSnapshots.staffPiCompletionSvgs = staffPiCompletionSvgs
  pptSnapshots.staffOwnerSvgs = staffOwnerSvgs
  pptSnapshots.staffSiSetupSvgs = staffSiSetupSvgs
  pptSnapshots.staffSiSubSvgs = staffSiSubSvgs
  pptSnapshots.staffAePublSvgs = staffAePublSvgs
  await monthlyWorkbook.xlsx.writeFile(monthlyOutputPath)
  await preserveTemplateCharts(input.paths.monthlyTemplate, monthlyOutputPath, ['\u79d1\u5ba4\u6570\u636e', ...focusOrder])
  await staffWorkbook.xlsx.writeFile(staffOutputPath)
  await verifySavedStaffWorkbookPiCompletion(staffOutputPath, editorRows)

  emitProgress(event, 'ppt', 'Copying and updating PPT template')
  const presentationPath = buildOutputPath(input.paths.outputDir, 'section-health-report-generated', context.reportKey, 'pptx')
  await buildPresentation({
    templatePath: input.paths.pptTemplate,
    outputPath: presentationPath,
    detected: {
      reportKey: context.reportKey,
      reportMonthLabel: context.reportMonthLabel,
      quarterPubHeader: context.quarterPubHeader,
      quarterRevenueHeader: context.quarterRevenueHeader
    },
    reportMonthTitle: context.reportMonthTitle,
    officeRows,
    completionRows,
    departmentMetrics,
    snapshots: pptSnapshots,
    staffPiCompletion: staffPiCompletionSummary,
    focusSeries,
    spotlights: buildSpotlights(focusSeries, completionRows, journalMap)
  })

  emitProgress(event, 'done', 'All files generated', 'success')
  return {
    success: true,
    outputs: {
      monthlyWorkbook: monthlyOutputPath,
      staffWorkbook: staffOutputPath,
      presentation: presentationPath
    },
    assumptions: [
      'PPT is copied from the Section Health meeting template and patched in place.',
      'Department summary defaults to a Section Health rollup and can be overridden in settings.',
      'Staff Journal values are mapped from the first-row journal columns in editors-journals.'
    ],
    detected: {
      reportKey: context.reportKey,
      reportMonthLabel: context.reportMonthLabel,
      quarterPubHeader: context.quarterPubHeader,
      quarterRevenueHeader: context.quarterRevenueHeader
    }
  }
}

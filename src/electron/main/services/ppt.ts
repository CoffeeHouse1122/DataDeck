import fs from 'node:fs/promises'
import JSZip from 'jszip'
import type { PipelineResult } from '../../../shared/contracts'
import { renderTableSvg } from './svg'

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
}

type DepartmentMetrics = {
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
  revenueDelta: number | null
  revenueMom: number | null
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

type JournalSpotlight = {
  displayName: string
  editors: string[]
  publicationTcr: number
  revenueTcr: number
  waiverRate2026: number
}

type PptBuildInput = {
  templatePath: string
  outputPath: string
  detected: PipelineResult['detected']
  reportMonthTitle: string
  officeRows: OfficeRow[]
  completionRows: CompletionRow[]
  departmentMetrics: DepartmentMetrics
  snapshots: {
    officeSvg: string
    completionSvg: string
    staffPiCompletionSvgs: string[]
    staffOwnerSvgs: string[]
    staffSiSetupSvgs: string[]
    staffSiSubSvgs: string[]
    staffAePublSvgs: string[]
  }
  staffPiCompletion: {
    currentCount: number
    previousCount: number
    delta: number
  }
  focusSeries: FocusJournalSeries[]
  spotlights: JournalSpotlight[]
}

const JOURNAL_SLIDES = [
  { metricSlide: 7, chartSlide: 8 },
  { metricSlide: 9, chartSlide: 10 },
  { metricSlide: 11, chartSlide: 12 },
  { metricSlide: 13, chartSlide: 14 },
  { metricSlide: 15, chartSlide: 16 }
]

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function percentText(value: number): string {
  return `${(Number.isFinite(value) ? value * 100 : 0).toFixed(2)}%`
}

function numText(value: number): string {
  return `${Math.round(Number.isFinite(value) ? value : 0)}`
}

function signedNumberText(value: number, digits = 0): string {
  const normalized = Number.isFinite(value) ? value : 0
  const sign = normalized >= 0 ? '+' : ''
  return `${sign}${normalized.toFixed(digits)}`
}

function signedPercentText(value: number): string {
  return signedNumberText(value * 100, 1) + '%'
}

function excelColumnLetter(index: number): string {
  let current = index
  let result = ''
  while (current > 0) {
    const remainder = (current - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    current = Math.floor((current - 1) / 26)
  }
  return result
}

function updateMetricSlide(xml: string, spotlight: JournalSpotlight): string {
  const values = [
    percentText(spotlight.publicationTcr),
    percentText(spotlight.revenueTcr),
    percentText(spotlight.waiverRate2026)
  ]
  let index = 0
  return xml
    .replace(/<a:t>Waiver Rate 2025<\/a:t>/g, '<a:t>Waiver Rate 2026</a:t>')
    .replace(/<a:t>(\d+(?:\.\d+)?%)<\/a:t>/g, (match) => {
      if (index >= values.length) {
        return match
      }
      const nextValue = values[index]
      index += 1
      return `<a:t>${nextValue}</a:t>`
    })
}

function updateTitleSlide(xml: string, reportMonthTitle: string): string {
  const match = reportMonthTitle.match(/^([A-Za-z])([A-Za-z]+)\s+(\d{4})$/)
  if (!match) {
    return xml
  }
  const values = [match[1], match[2], ` ${match[3]}`]
  let index = 0
  return xml.replace(/<a:t>[\s\S]*?<\/a:t>/g, (textNode) => {
    if (index >= values.length) {
      return textNode
    }
    const value = values[index]
    index += 1
    return `<a:t>${escapeXml(value)}</a:t>`
  })
}

function cacheXml(values: Array<string | number>, stringValues: boolean): string {
  // Keep the same precision as Edit Data; integer labels are a display format only.
  const points = values.map((value, index) =>
    `<c:pt idx="${index}"><c:v>${stringValues ? escapeXml(String(value)) : Number(value) || 0}</c:v></c:pt>`
  ).join('')
  return `<c:ptCount val="${values.length}"/>${points}`
}

function replaceNthCache(serXml: string, cacheTag: 'strCache' | 'numCache', occurrence: number, values: Array<string | number>, stringValues: boolean): string {
  let seen = 0
  const pattern = new RegExp(`<c:${cacheTag}>[\\s\\S]*?<\\/c:${cacheTag}>`, 'g')
  return serXml.replace(pattern, (match) => {
    if (seen !== occurrence) {
      seen += 1
      return match
    }
    seen += 1
    const format = cacheTag === 'numCache' ? '<c:formatCode>0</c:formatCode>' : ''
    return `<c:${cacheTag}>${format}${cacheXml(values, stringValues)}</c:${cacheTag}>`
  })
}

function setSeriesDataLabelStyle(serXml: string, size: number): string {
  // Do not inherit General from the workbook when PowerPoint refreshes the chart.
  return serXml
    .replace(/<a:defRPr([^>]*)sz="\d+"/g, `<a:defRPr$1sz="${size}"`)
    .replace(/<c:numFmt\b[^>]*\/>/g, '<c:numFmt formatCode="0" sourceLinked="0"/>')
}

function replaceSerFormulae(serXml: string, sheetName: string, monthIndex: number, metricCount: number): string {
  const column = String.fromCharCode(66 + monthIndex)
  const escapedSheet = sheetName.includes(' ') ? `'${sheetName}'` : sheetName
  const formulae = [
    `${escapedSheet}!$${column}$1`,
    `${escapedSheet}!$A$2:$A$${metricCount + 1}`,
    `${escapedSheet}!$${column}$2:$${column}$${metricCount + 1}`
  ]
  let index = 0
  return serXml.replace(/<c:f>[\s\S]*?<\/c:f>/g, (match) => {
    if (index >= formulae.length) {
      return match
    }
    const formula = formulae[index]
    index += 1
    return `<c:f>${escapeXml(formula)}</c:f>`
  })
}

function updateChartXml(xml: string, sheetName: string, months: string[], metrics: Array<{ label: string; values: number[] }>): string {
  const usableMonths = months.length ? months : ['']
  const usableMetrics = metrics.length ? metrics : [{ label: '', values: [0] }]
  const labels = usableMetrics.map((metric) => metric.label)
  const seriesMatches = xml.match(/<c:ser>[\s\S]*?<\/c:ser>/g) ?? []
  if (!seriesMatches.length) {
    return xml
  }

  const templateSeries = seriesMatches[seriesMatches.length - 1]
  const compactLabelSize = usableMonths.length >= 8 ? 850 : usableMonths.length >= 6 ? 1000 : 1200
  const nextSeries = usableMonths.map((month, index) => {
    const source = seriesMatches[index] ?? templateSeries
    const values = usableMetrics.map((metric) => metric.values[index] ?? 0)
    let serXml = source
      .replace(/<c:idx val="\d+"\s*\/>/, `<c:idx val="${index}"/>`)
      .replace(/<c:order val="\d+"\s*\/>/, `<c:order val="${index}"/>`)
    serXml = replaceNthCache(serXml, 'strCache', 0, [month], true)
    serXml = replaceNthCache(serXml, 'strCache', 1, labels, true)
    serXml = replaceNthCache(serXml, 'numCache', 0, values, false)
    serXml = setSeriesDataLabelStyle(serXml, index === usableMonths.length - 1 ? 1500 : compactLabelSize)
    return replaceSerFormulae(serXml, sheetName, index, labels.length)
  }).join('')

  let replaced = false
  return xml.replace(/(?:<c:ser>[\s\S]*?<\/c:ser>)+/, () => {
    if (replaced) {
      return ''
    }
    replaced = true
    return nextSeries
  })
}

function replaceDepartmentSerFormulae(serXml: string, monthIndex: number, metricCount: number): string {
  const row = monthIndex + 2
  const endColumn = excelColumnLetter(metricCount + 1)
  const formulae = [
    `科室数据!$A$${row}`,
    `科室数据!$B$1:$${endColumn}$1`,
    `科室数据!$B$${row}:$${endColumn}$${row}`
  ]
  let index = 0
  return serXml.replace(/<c:f>[\s\S]*?<\/c:f>/g, (match) => {
    if (index >= formulae.length) {
      return match
    }
    const formula = formulae[index]
    index += 1
    return `<c:f>${escapeXml(formula)}</c:f>`
  })
}

function updateDepartmentChartXml(xml: string, months: string[], metrics: Array<{ label: string; values: number[] }>): string {
  const usableMonths = months.length ? months : ['']
  const usableMetrics = metrics.length ? metrics : [{ label: '', values: [0] }]
  const labels = usableMetrics.map((metric) => metric.label)
  const seriesMatches = xml.match(/<c:ser>[\s\S]*?<\/c:ser>/g) ?? []
  if (!seriesMatches.length) {
    return xml
  }

  const templateSeries = seriesMatches[seriesMatches.length - 1]
  const compactLabelSize = usableMonths.length >= 8 ? 850 : usableMonths.length >= 6 ? 1000 : 1200
  const nextSeries = usableMonths.map((month, index) => {
    const source = seriesMatches[index] ?? templateSeries
    const values = usableMetrics.map((metric) => metric.values[index] ?? 0)
    let serXml = source
      .replace(/<c:idx val="\d+"\s*\/>/, `<c:idx val="${index}"/>`)
      .replace(/<c:order val="\d+"\s*\/>/, `<c:order val="${index}"/>`)
    serXml = replaceNthCache(serXml, 'strCache', 0, [month], true)
    serXml = replaceNthCache(serXml, 'strCache', 1, labels, true)
    serXml = replaceNthCache(serXml, 'numCache', 0, values, false)
    serXml = setSeriesDataLabelStyle(serXml, index === usableMonths.length - 1 ? 1500 : compactLabelSize)
    return replaceDepartmentSerFormulae(serXml, index, labels.length)
  }).join('')

  let replaced = false
  return xml.replace(/(?:<c:ser>[\s\S]*?<\/c:ser>)+/, () => {
    if (replaced) {
      return ''
    }
    replaced = true
    return nextSeries
  })
}

async function slideChartTargets(zip: JSZip, slideNumber: number): Promise<string[]> {
  const slideXml = await zip.file(`ppt/slides/slide${slideNumber}.xml`)?.async('string')
  const relsXml = await zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`)?.async('string')
  if (!slideXml || !relsXml) {
    return []
  }

  const chartIds = Array.from(slideXml.matchAll(/<c:chart[^>]+r:id="([^"]+)"/g)).map((match) => match[1])
  return chartIds.map((id) => {
    const relMatch = relsXml.match(new RegExp(`<Relationship[^>]+Id="${id}"[^>]+Target="([^"]+)"`))
    return relMatch ? `ppt/${relMatch[1].replace(/^\.\.\//, '')}` : ''
  }).filter(Boolean)
}

function chartMonths(metrics: Array<{ label: string; values: number[] }>, fallbackMonths: string[]): string[] {
  const maxLength = Math.max(...metrics.map((metric) => metric.values.length), 0)
  return fallbackMonths.slice(0, maxLength || 1)
}

async function updateJournalCharts(zip: JSZip, slideNumber: number, series: FocusJournalSeries): Promise<void> {
  const chartTargets = await slideChartTargets(zip, slideNumber)
  const chartUpdates = [
    { target: chartTargets[0], metrics: series.primaryMetrics, months: chartMonths(series.primaryMetrics, series.months) },
    { target: chartTargets[1], metrics: series.secondaryMetrics, months: chartMonths(series.secondaryMetrics, series.months) }
  ]

  for (const update of chartUpdates) {
    if (!update.target) {
      continue
    }
    const file = zip.file(update.target)
    if (!file) {
      continue
    }
    const xml = await file.async('string')
    zip.file(update.target, updateChartXml(xml, series.sheetName === 'BS' ? 'Brain Sciences' : series.sheetName, update.months, update.metrics))
  }
}

function buildOfficeTableSvg(rows: OfficeRow[]): string {
  const columns = [
    { key: 'journal', label: 'Journal', width: 126 },
    { key: 'contrib', label: 'Contrib.', width: 56 },
    { key: 'yearlyPublTarget', label: 'Yearly Target', width: 76 },
    { key: 'monthlyPublTarget', label: 'Monthly Target', width: 80 },
    { key: 'publ', label: 'Publ.', width: 54 },
    { key: 'momBase', label: 'Pub/lastM', width: 66 },
    { key: 'mom', label: 'MoM (%)', width: 64 },
    { key: 'pubYoy', label: 'Pub/YOY', width: 64 },
    { key: 'yoy', label: 'YoY (%)', width: 64 },
    { key: 'newSi', label: 'New SI', width: 56 },
    { key: 'avePubSis', label: 'Ave. Pub SIs', width: 76 },
    { key: 'cnRate', label: 'CN%', width: 56 },
    { key: 'mpt', label: 'MPT', width: 48 },
    { key: 'sub', label: 'Sub.', width: 58 },
    { key: 'susyCfp', label: 'Susy CFP', width: 68 },
    { key: 'mmailerRegularCfp', label: 'Mmailer-Regular CFP', width: 104 },
    { key: 'mmailerFCfp', label: 'Mmailer-FCFP', width: 88 },
    { key: 'sendCfp', label: 'Send CFP', width: 74 },
    { key: 'wrTarget2026', label: 'WR Target2026', width: 86 },
    { key: 'waiverRate', label: 'Waiver Rate', width: 76 },
    { key: 'waiverRate2026', label: 'Waiver Rate2026', width: 96 }
  ]

  return renderTableSvg({
    rowHeight: 24,
    headerHeight: 32,
    columns,
    rows: rows.map((row) => ({
      journal: { value: row.journal, fontSize: 10, weight: 600 },
      contrib: { value: numText(row.contrib), align: 'right' },
      yearlyPublTarget: { value: numText(row.yearlyPublTarget), align: 'right' },
      monthlyPublTarget: { value: numText(row.monthlyPublTarget), align: 'right' },
      publ: { value: numText(row.publ), align: 'right' },
      momBase: { value: numText(row.publLast), align: 'right' },
      mom: { value: percentText(row.mom), align: 'right' },
      pubYoy: { value: numText(row.pubYoy), align: 'right' },
      yoy: { value: percentText(row.yoy), align: 'right' },
      newSi: { value: numText(row.newSi), align: 'right' },
      avePubSis: { value: row.avePubSis.toFixed(2), align: 'right' },
      cnRate: { value: percentText(row.cnRate), align: 'right' },
      mpt: { value: numText(row.mpt), align: 'right' },
      sub: { value: numText(row.sub), align: 'right' },
      susyCfp: { value: numText(row.susyCfp), align: 'right' },
      mmailerRegularCfp: { value: numText(row.mmailerRegularCfp), align: 'right' },
      mmailerFCfp: { value: numText(row.mmailerFCfp), align: 'right' },
      sendCfp: { value: numText(row.sendCfp), align: 'right' },
      wrTarget2026: { value: percentText(row.wrTarget2026), align: 'right' },
      waiverRate: { value: percentText(row.waiverRate), align: 'right' },
      waiverRate2026: { value: percentText(row.waiverRate2026), align: 'right' }
    }))
  })
}

async function ensureSvgContentType(zip: JSZip): Promise<void> {
  const contentTypesFile = zip.file('[Content_Types].xml')
  if (!contentTypesFile) {
    return
  }
  const contentTypes = await contentTypesFile.async('string')
  if (!contentTypes.includes('Extension="svg"')) {
    zip.file('[Content_Types].xml', contentTypes.replace(
      '</Types>',
      '<Default Extension="svg" ContentType="image/svg+xml"/></Types>'
    ))
  }
}

async function replaceSlideImageWithSvg(zip: JSZip, slideNumber: number, relationId: string, targetName: string, svg: string): Promise<void> {
  const relPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`
  const relFile = zip.file(relPath)
  if (!relFile) {
    return
  }
  const relsXml = await relFile.async('string')
  zip.file(relPath, relsXml.replace(
    new RegExp(`(<Relationship[^>]+Id="${relationId}"[^>]+Type="http://schemas\\.openxmlformats\\.org/officeDocument/2006/relationships/image"[^>]+Target=")[^"]+("[^>]*\\/>)`),
    `$1../media/${targetName}$2`
  ))
  zip.file(`ppt/media/${targetName}`, svg)
  await ensureSvgContentType(zip)
}

async function replaceSlide5OfficeImage(zip: JSZip, officeSvg: string): Promise<void> {
  await replaceSlideImageWithSvg(zip, 5, 'rId2', 'office-sheet-a-u.svg', officeSvg)
}

function departmentChartMetrics(metrics: DepartmentMetrics): Array<{ label: string; values: number[] }> {
  return [
    { label: 'Publication', values: metrics.yearlySeries.map((item) => item.publication) },
    { label: 'Submission', values: metrics.yearlySeries.map((item) => item.submission) },
    { label: 'Assigned Manuscript', values: metrics.yearlySeries.map((item) => item.assignedManuscript) },
    { label: 'SI Set Up', values: metrics.yearlySeries.map((item) => item.siSetUp) },
    { label: 'Revenue(WCHF)', values: metrics.yearlySeries.map((item) => item.revenueWCHF) },
    { label: 'Waiver Rate(%)', values: metrics.yearlySeries.map((item) => item.waiverRate) },
    { label: 'MPT', values: metrics.yearlySeries.map((item) => item.mpt) }
  ]
}

async function updateDepartmentChart(zip: JSZip, metrics: DepartmentMetrics): Promise<void> {
  const [chartTarget] = await slideChartTargets(zip, 6)
  if (!chartTarget) {
    return
  }
  const file = zip.file(chartTarget)
  if (!file) {
    return
  }
  const months = metrics.yearlySeries.map((item) => item.label)
  zip.file(chartTarget, updateDepartmentChartXml(await file.async('string'), months, departmentChartMetrics(metrics)))
}

function updateDepartmentTextSlide(xml: string, metrics: DepartmentMetrics): string {
  const replacements = [
    `发文：${numText(metrics.publication)}篇（比上月${signedNumberText(metrics.publicationDelta)}），MoM ${signedPercentText(metrics.publicationMom)}，YoY ${metrics.publicationYoy === null ? '--' : signedPercentText(metrics.publicationYoy)}`,
    `营收：${metrics.revenueWCHF.toFixed(2)} WCHF（比上月${metrics.revenueDelta === null ? '--' : `${signedNumberText(metrics.revenueDelta, 2)}W`}），MoM ${metrics.revenueMom === null ? '--' : signedPercentText(metrics.revenueMom)}，YoY ${metrics.revenueYoy === null ? '--' : signedPercentText(metrics.revenueYoy)}`,
    `投稿：${numText(metrics.submission)}篇（比上月${signedNumberText(metrics.submissionDelta)}），MoM ${signedPercentText(metrics.submissionMom)}`,
    `特刊上线：${numText(metrics.siSetUp)}`,
    `免费比例：${(metrics.waiverRate * 100).toFixed(2)}%`,
    `MPT：${numText(metrics.mpt)}天`
  ]
  const patterns = [
    /<a:t>发文：[\s\S]*?<\/a:t>/,
    /<a:t>营收：[\s\S]*?<\/a:t>/,
    /<a:t>投稿：[\s\S]*?<\/a:t>/,
    /<a:t>特刊上线：[\s\S]*?<\/a:t>/,
    /<a:t>免费比例：[\s\S]*?<\/a:t>/,
    /<a:t>MPT：[\s\S]*?<\/a:t>/
  ]
  return patterns.reduce((nextXml, pattern, index) => nextXml.replace(pattern, `<a:t>${escapeXml(replacements[index])}</a:t>`), xml)
}

async function updateDepartmentSlide(zip: JSZip, metrics: DepartmentMetrics): Promise<void> {
  const slidePath = 'ppt/slides/slide6.xml'
  const slide = zip.file(slidePath)
  if (slide) {
    zip.file(slidePath, updateDepartmentTextSlide(await slide.async('string'), metrics))
  }
  await updateDepartmentChart(zip, metrics)
}

function buildCompletionTableSvg(rows: CompletionRow[]): string {
  const filtered = rows.filter((row) => row.group !== 'SCIE')
  const columns = [
    { key: 'group', label: '', width: 70 },
    { key: 'journal', label: 'Journal', width: 190 },
    { key: 'contrib', label: 'Contrib.', width: 64 },
    { key: 'quarterPub', label: 'QTD Publ', width: 70 },
    { key: 'yearlyPublTarget', label: 'Yearly Target', width: 86 },
    { key: 'publicationTcr', label: 'Publication TCR', width: 96 },
    { key: 'quarterRevenue', label: 'QTD Revenue', width: 92 },
    { key: 'revenueTargetFinal', label: 'Revenue Target', width: 98 },
    { key: 'revenueTcr', label: 'Revenue TCR', width: 88 },
    { key: 'waiverRate', label: 'Waiver Rate', width: 82 },
    { key: 'wrTarget2026', label: 'WR Target2026', width: 92 },
    { key: 'waiverRate2026', label: 'Waiver Rate 2026', width: 106 }
  ]

  return renderTableSvg({
    rowHeight: 26,
    headerHeight: 34,
    columns,
    rows: filtered.map((row, index) => ({
      group: { value: index === 0 || filtered[index - 1]?.group !== row.group ? row.group : '', weight: 700, align: 'center' },
      journal: { value: row.journal },
      contrib: { value: numText(row.contrib), align: 'right' },
      quarterPub: { value: numText(row.quarterPub), align: 'right' },
      yearlyPublTarget: { value: numText(row.yearlyPublTarget), align: 'right' },
      publicationTcr: { value: percentText(row.publicationTcr), align: 'right' },
      quarterRevenue: { value: numText(row.quarterRevenue), align: 'right' },
      revenueTargetFinal: { value: numText(row.revenueTargetFinal), align: 'right' },
      revenueTcr: { value: percentText(row.revenueTcr), align: 'right' },
      waiverRate: { value: percentText(row.waiverRate), align: 'right' },
      wrTarget2026: { value: percentText(row.wrTarget2026), align: 'right' },
      waiverRate2026: { value: percentText(row.waiverRate2026), align: 'right' }
    }))
  })
}

async function replaceSlide17Image(zip: JSZip, completionSvg: string): Promise<void> {
  const relPath = 'ppt/slides/_rels/slide17.xml.rels'
  const relFile = zip.file(relPath)
  if (!relFile) {
    return
  }

  const svgTarget = '../media/completion-rate-esci-scopus-others.svg'
  const relsXml = await relFile.async('string')
  const nextRels = relsXml.replace(
    /(<Relationship[^>]+Id="rId2"[^>]+Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/image"[^>]+Target=")[^"]+("[^>]*\/>)/,
    `$1${svgTarget}$2`
  )
  zip.file(relPath, nextRels)
  zip.file('ppt/media/completion-rate-esci-scopus-others.svg', completionSvg)
  await ensureSvgContentType(zip)
}

function blankSvg(): string {
  return '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1" viewBox="0 0 1 1"><rect width="1" height="1" fill="#ffffff" opacity="0"/></svg>'
}

function signedCountText(value: number): string {
  return value >= 0 ? `+${value}` : String(value)
}

function updateStaffPiCompletionText(xml: string, currentCount: number, delta: number): string {
  let markerSeen = false
  let countDone = false
  let deltaDone = false

  return xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (node, text) => {
    if (String(text).includes('100%')) {
      markerSeen = true
      return node
    }
    if (!markerSeen) {
      return node
    }
    if (!countDone && /^\d+$/.test(String(text).trim())) {
      countDone = true
      return `<a:t>${currentCount}</a:t>`
    }
    if (countDone && !deltaDone && /^[+-]?\d+$/.test(String(text).trim())) {
      deltaDone = true
      return `<a:t>${signedCountText(delta)}</a:t>`
    }
    return node
  })
}

async function updateStaffPiCompletionSlide(zip: JSZip, summary: PptBuildInput['staffPiCompletion'], tableSvgs: string[]): Promise<void> {
  const slidePath = 'ppt/slides/slide19.xml'
  const slide = zip.file(slidePath)
  if (slide) {
    zip.file(slidePath, updateStaffPiCompletionText(await slide.async('string'), summary.currentCount, summary.delta))
  }

  await replaceSlideImageWithSvg(zip, 19, 'rId1', 'staff-pi-completion-1.svg', tableSvgs[0] ?? blankSvg())
  await replaceSlideImageWithSvg(zip, 19, 'rId2', 'staff-pi-completion-2.svg', tableSvgs[1] ?? blankSvg())
}

async function updateStaffOwnerPublicationSlide(zip: JSZip, tableSvgs: string[]): Promise<void> {
  const slidePath = 'ppt/slides/slide20.xml'
  const slide = zip.file(slidePath)
  if (slide) {
    const xml = await slide.async('string')
    zip.file(slidePath, xml
      .replace(/<a:t>特刊发文<\/a:t>/g, '<a:t>owner发文</a:t>')
      .replace(/<a:t>鐗瑰垔鍙戞枃<\/a:t>/g, '<a:t>owner发文</a:t>'))
  }

  await replaceSlideImageWithSvg(zip, 20, 'rId1', 'staff-owner-publication.svg', tableSvgs[0] ?? blankSvg())
  await replaceSlideImageWithSvg(zip, 20, 'rId2', 'mr-si-publ-top30.svg', tableSvgs[1] ?? blankSvg())
}

async function updateStaffSiSetupSlide(zip: JSZip, tableSvgs: string[]): Promise<void> {
  await replaceSlideImageWithSvg(zip, 21, 'rId1', 'staff-sme-si-setup.svg', tableSvgs[0] ?? blankSvg())
  await replaceSlideImageWithSvg(zip, 21, 'rId2', 'mr-si-setup-top30.svg', tableSvgs[1] ?? blankSvg())
}

async function updateStaffSiSubSlide(zip: JSZip, tableSvgs: string[]): Promise<void> {
  await replaceSlideImageWithSvg(zip, 22, 'rId1', 'staff-sme-si-sub.svg', tableSvgs[0] ?? blankSvg())
  await replaceSlideImageWithSvg(zip, 22, 'rId2', 'mr-si-sub-top30.svg', tableSvgs[1] ?? blankSvg())
}

async function updateStaffAePublSlide(zip: JSZip, tableSvgs: string[]): Promise<void> {
  await replaceSlideImageWithSvg(zip, 23, 'rId1', 'staff-ae-publ.svg', tableSvgs[0] ?? blankSvg())
  await replaceSlideImageWithSvg(zip, 23, 'rId2', 'mr-publ-top30.svg', tableSvgs[1] ?? blankSvg())
}

function xmlAttr(tag: string, name: string): string | null {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return tag.match(new RegExp(`${escaped}="([^"]+)"`))?.[1] ?? null
}

function nextSlideNumber(zip: JSZip): number {
  const slideNumbers = Object.keys(zip.files)
    .map((fileName) => fileName.match(/^ppt\/slides\/slide(\d+)\.xml$/)?.[1])
    .filter((value): value is string => Boolean(value))
    .map((value) => Number.parseInt(value, 10))
  return Math.max(0, ...slideNumbers) + 1
}

function nextPresentationRelationshipId(relsXml: string): string {
  const ids = Array.from(relsXml.matchAll(/Id="rId(\d+)"/g)).map((match) => Number.parseInt(match[1], 10))
  return `rId${Math.max(0, ...ids) + 1}`
}

function nextPresentationSlideId(presentationXml: string): number {
  const ids = Array.from(presentationXml.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"/g)).map((match) => Number.parseInt(match[1], 10))
  return Math.max(255, ...ids) + 1
}

function presentationSlideRelationshipId(relsXml: string, slideNumber: number): string | null {
  const relationshipTags = relsXml.match(/<Relationship\b[^>]*\/>/g) ?? []
  const target = `slides/slide${slideNumber}.xml`
  const relationship = relationshipTags.find((tag) =>
    tag.includes('/relationships/slide') && xmlAttr(tag, 'Target') === target
  )
  return relationship ? xmlAttr(relationship, 'Id') : null
}

async function slideLayoutTarget(zip: JSZip, slideNumber: number): Promise<string> {
  const relsXml = await zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`)?.async('string')
  const relationshipTags = relsXml?.match(/<Relationship\b[^>]*\/>/g) ?? []
  const layoutRelationship = relationshipTags.find((tag) => tag.includes('/relationships/slideLayout'))
  return layoutRelationship ? xmlAttr(layoutRelationship, 'Target') ?? '../slideLayouts/slideLayout1.xml' : '../slideLayouts/slideLayout1.xml'
}

function blankSlideXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`
}

function blankSlideRelsXml(layoutTarget: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${layoutTarget}"/>
</Relationships>`
}

async function ensureSlideContentType(zip: JSZip, slideNumber: number): Promise<void> {
  const contentTypesFile = zip.file('[Content_Types].xml')
  if (!contentTypesFile) {
    return
  }
  const partName = `/ppt/slides/slide${slideNumber}.xml`
  const contentTypes = await contentTypesFile.async('string')
  if (contentTypes.includes(`PartName="${partName}"`)) {
    return
  }
  zip.file('[Content_Types].xml', contentTypes.replace(
    '</Types>',
    `<Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`
  ))
}

async function updateAppSlideCount(zip: JSZip): Promise<void> {
  const appFile = zip.file('docProps/app.xml')
  if (!appFile) {
    return
  }
  const appXml = await appFile.async('string')
  zip.file('docProps/app.xml', appXml.replace(/<Slides>(\d+)<\/Slides>/, (_match, count) => `<Slides>${Number.parseInt(count, 10) + 1}</Slides>`))
}

export async function buildPresentation(input: PptBuildInput): Promise<void> {
  const zip = await JSZip.loadAsync(await fs.readFile(input.templatePath))

  const titleSlidePath = 'ppt/slides/slide1.xml'
  const titleSlide = zip.file(titleSlidePath)
  if (titleSlide) {
    zip.file(titleSlidePath, updateTitleSlide(await titleSlide.async('string'), input.reportMonthTitle))
  }

  await replaceSlide5OfficeImage(zip, input.snapshots.officeSvg)
  await updateDepartmentSlide(zip, input.departmentMetrics)

  for (const [index, slideMap] of JOURNAL_SLIDES.entries()) {
    const spotlight = input.spotlights[index]
    const series = input.focusSeries[index]
    if (!spotlight || !series) {
      continue
    }

    const metricSlidePath = `ppt/slides/slide${slideMap.metricSlide}.xml`
    const metricSlide = zip.file(metricSlidePath)
    if (metricSlide) {
      zip.file(metricSlidePath, updateMetricSlide(await metricSlide.async('string'), spotlight))
    }

    await updateJournalCharts(zip, slideMap.chartSlide, series)
  }

  await replaceSlide17Image(zip, input.snapshots.completionSvg)
  await updateStaffPiCompletionSlide(zip, input.staffPiCompletion, input.snapshots.staffPiCompletionSvgs)
  await updateStaffOwnerPublicationSlide(zip, input.snapshots.staffOwnerSvgs)
  await updateStaffSiSetupSlide(zip, input.snapshots.staffSiSetupSvgs)
  await updateStaffSiSubSlide(zip, input.snapshots.staffSiSubSvgs)
  await updateStaffAePublSlide(zip, input.snapshots.staffAePublSvgs)
  await fs.writeFile(input.outputPath, await zip.generateAsync({ type: 'nodebuffer' }))
}

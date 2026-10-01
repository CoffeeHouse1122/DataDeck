import { XMLBuilder, XMLParser } from 'fast-xml-parser'

// One color per calendar month, including comparison windows crossing a year boundary.
const MONTH_COLORS = [
  '156082', 'E97132', '196B24', '0F9ED5', '8064A2', 'C49A00',
  'C0504D', '00857D', 'A64D79', '596579', '8A6B3F', '729B37'
]
const MONTH_NAMES = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
const xmlOptions = { ignoreAttributes: false, preserveOrder: true, parseTagValue: false, parseAttributeValue: false }
const parser = new XMLParser(xmlOptions)
const builder = new XMLBuilder({ ...xmlOptions, suppressEmptyNode: true })

export function monthSeriesColor(label: string, fallbackIndex: number): string {
  const normalized = label.trim().toLowerCase()
  const numeric = normalized.match(/^\d{4}[-/](\d{1,2})(?:$|[-/])/) ?? normalized.match(/^(\d{1,2})月$/)
  const monthIndex = numeric ? Number(numeric[1]) - 1 : MONTH_NAMES.indexOf(normalized.slice(0, 3))
  return MONTH_COLORS[monthIndex >= 0 && monthIndex < 12 ? monthIndex : fallbackIndex % 12]
}

function colorShapeProperties(xml: string, color: string): string {
  const nodes = parser.parse(xml || '<c:spPr/>')
  const properties = nodes[0]['c:spPr']
  const fillNames = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill']
  const retained = properties.filter((node: Record<string, unknown>) => !fillNames.some((name) => name in node))
  // DrawingML puts fills after transforms/geometry and before lines/effects.
  const afterGeometry = retained.findIndex((node: Record<string, unknown>) =>
    !['a:xfrm', 'a:prstGeom', 'a:custGeom'].some((name) => name in node))
  retained.splice(afterGeometry < 0 ? retained.length : afterGeometry, 0,
    { 'a:solidFill': [{ 'a:srgbClr': [], ':@': { '@_val': color } }] })
  nodes[0]['c:spPr'] = retained
  return builder.build(nodes)
}

export function colorMonthSeries(xml: string, label: string, index: number): string {
  const color = monthSeriesColor(label, index)
  // Only change shape properties, leaving formulas, cached values and label styling intact.
  const colored = xml.replace(
    /^(<c:ser\b[^>]*>[\s\S]*?<\/c:tx>)(\s*<c:spPr(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/c:spPr>))?/,
    (_match, prefix: string, properties: string | undefined) => prefix + colorShapeProperties(properties?.trim() ?? '', color)
  )
  // Point-level fills would otherwise override the month color for individual bars.
  return colored.replace(/<c:dPt\b[^>]*>[\s\S]*?<\/c:dPt>/g, (point) =>
    point.replace(/<c:spPr(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/c:spPr>)/, (properties) => colorShapeProperties(properties, color)))
}

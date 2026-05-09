export const FOCUS_JOURNAL_OPTIONS = [
  { label: 'Foods', value: 'Foods' },
  { label: 'Nutrients', value: 'Nutrients' },
  { label: 'Children', value: 'Children' },
  { label: 'Genes', value: 'Genes' },
  { label: 'BS / Brain Sciences', value: 'BS' }
]

export const PATH_FIELD_META = [
  {
    key: 'mrWorkbook',
    label: 'MR 数据源',
    description: '选择每月 MR_xxxxxx-xxxxxx 工作簿',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  },
  {
    key: 'monthlyTemplate',
    label: '月会数据模板',
    description: '默认使用 monthly-data-generated-202603.xlsx',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  },
  {
    key: 'staffTemplate',
    label: '人员数据模板',
    description: '默认使用 staff-data-generated-202603.xlsx',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  },
  {
    key: 'editorsJournals',
    label: '人员刊物映射',
    description: '默认使用内置 editors-journals.xlsx',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  },
  {
    key: 'pptTemplate',
    label: 'PPT 模板',
    description: '默认使用内置 Section Health月会.pptx',
    filters: [{ name: 'PowerPoint', extensions: ['pptx'] }]
  },
  {
    key: 'outputDir',
    label: '输出目录',
    description: '生成结果会写入该目录',
    filters: []
  }
] as const

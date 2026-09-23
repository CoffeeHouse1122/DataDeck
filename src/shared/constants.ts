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
    description: '选择手动接收的月会数据模板',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  },
  {
    key: 'staffTemplate',
    label: '人员数据模板',
    description: '选择手动接收的人员数据模板',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  },
  {
    key: 'editorsJournals',
    label: '人员刊物映射',
    description: '选择手动接收的 editors-journals.xlsx',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  },
  {
    key: 'pptTemplate',
    label: 'PPT 模板',
    description: '选择手动接收的 PPT 模板',
    filters: [{ name: 'PowerPoint', extensions: ['pptx'] }]
  },
  {
    key: 'outputDir',
    label: '输出目录',
    description: '生成结果会写入该目录',
    filters: []
  }
] as const

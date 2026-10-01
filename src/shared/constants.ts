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
    description: '每月选择上月生成并核对过的月会 Excel，用于保留历史月份数据。',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  },
  {
    key: 'staffTemplate',
    label: '人员数据模板',
    description: '每月选择上月生成并核对过的人员 Excel，部分人数变化以上月文件为比较基准。',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  },
  {
    key: 'editorsJournals',
    label: '人员刊物映射',
    description: '人员与期刊的对应关系发生变化时更新，不必每月更换。',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  },
  {
    key: 'pptTemplate',
    label: 'PPT 模板',
    description: '使用已验证的固定配套模板即可，不必每月更换。不要自行增删或调整页面顺序。',
    filters: [{ name: 'PowerPoint', extensions: ['pptx'] }]
  },
  {
    key: 'outputDir',
    label: '输出目录',
    description: '生成结果会写入该目录',
    filters: []
  }
] as const

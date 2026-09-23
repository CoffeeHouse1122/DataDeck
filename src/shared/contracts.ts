export type CloseBehavior = 'quit' | 'tray'

export interface SelectedPaths {
  mrWorkbook: string
  monthlyTemplate: string
  staffTemplate: string
  editorsJournals: string
  pptTemplate: string
  outputDir: string
}

export interface SummaryOverrides {
  reportMonthLabel: string
  publication: string
  submission: string
  assignedManuscript: string
  siSetUp: string
  revenueWCHF: string
  waiverRate: string
  mpt: string
}

export interface AppPreferences {
  paths: Partial<SelectedPaths>
  closeBehavior: CloseBehavior
  focusJournalOrder: string[]
  forceAeStaff: string[]
  summaryOverrides: SummaryOverrides
}

export interface PipelineInput {
  paths: SelectedPaths
  closeBehavior: CloseBehavior
  focusJournalOrder: string[]
  forceAeStaff: string[]
  summaryOverrides: SummaryOverrides
}

export type ProgressLevel = 'info' | 'success' | 'warning' | 'error'

export interface PipelineProgressEvent {
  level: ProgressLevel
  message: string
  step: string
}

export interface PipelineResult {
  success: boolean
  outputs: {
    monthlyWorkbook: string
    staffWorkbook: string
    presentation: string
  }
  assumptions: string[]
  detected: {
    reportKey: string
    reportMonthLabel: string
    quarterPubHeader: string
    quarterRevenueHeader: string
  }
}

export interface WindowState {
  isAlwaysOnTop: boolean
  isMaximized: boolean
}

export interface FilePickerOptions {
  title: string
  filters: Array<{
    extensions: string[]
    name: string
  }>
}

export interface ElectronApi {
  getUpdateState(): Promise<AppUpdateState>
  checkForUpdates(): Promise<AppUpdateState>
  downloadUpdate(): Promise<AppUpdateState>
  installUpdate(): Promise<AppUpdateState>
  onUpdateState(listener: (state: AppUpdateState) => void): () => void
  pickFile(options: FilePickerOptions): Promise<string | null>
  pickDirectory(title: string): Promise<string | null>
  getPreferences(): Promise<AppPreferences>
  getDefaultPaths(): Promise<Partial<SelectedPaths>>
  savePreferences(preferences: AppPreferences): Promise<void>
  runPipeline(input: PipelineInput): Promise<PipelineResult>
  revealPath(targetPath: string): Promise<void>
  refreshWindow(): Promise<void>
  toggleAlwaysOnTop(): Promise<WindowState>
  minimizeWindow(): Promise<void>
  toggleMaximizeWindow(): Promise<WindowState>
  closeWindow(): Promise<void>
  getWindowState(): Promise<WindowState>
  onWindowStateChange(listener: (state: WindowState) => void): () => void
  onPipelineProgress(listener: (event: PipelineProgressEvent) => void): () => void
}

export interface AppUpdateState {
  phase: 'disabled' | 'idle' | 'checking' | 'available' | 'downloading' | 'downloaded' | 'installing' | 'error'
  currentVersion: string
  availableVersion?: string
  percent?: number
  message: string
  generationRunning: boolean
}

declare global {
  interface Window {
    electronApi: ElectronApi
  }
}

export const DEFAULT_FOCUS_JOURNAL_ORDER = ['Foods', 'Nutrients', 'Children', 'Genes', 'BS']
export const DEFAULT_FORCE_AE_STAFF: string[] = []

export const DEFAULT_SUMMARY_OVERRIDES: SummaryOverrides = {
  reportMonthLabel: '',
  publication: '',
  submission: '',
  assignedManuscript: '',
  siSetUp: '',
  revenueWCHF: '',
  waiverRate: '',
  mpt: ''
}

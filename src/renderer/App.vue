<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, reactive, ref, watch } from 'vue'
import CustomSelect from './components/CustomSelect.vue'
import PathField from './components/PathField.vue'
import SimpleBarScroll from './components/SimpleBarScroll.vue'
import UpdatePanel from './components/UpdatePanel.vue'
import brandIconUrl from '../../resources/icons/favicon-256x256.png'
import { FOCUS_JOURNAL_OPTIONS, PATH_FIELD_META } from '../shared/constants'
import {
  DEFAULT_FOCUS_JOURNAL_ORDER,
  DEFAULT_FORCE_AE_STAFF,
  DEFAULT_SUMMARY_OVERRIDES,
  type AppPreferences,
  type AppUpdateState,
  type PipelineProgressEvent,
  type PipelineResult,
  type SelectedPaths,
  type SummaryOverrides,
  type WindowState
} from '../shared/contracts'

type Toast = {
  id: number
  tone: 'success' | 'error' | 'info'
  text: string
}

const paths = reactive<SelectedPaths>({
  mrWorkbook: '',
  monthlyTemplate: '',
  staffTemplate: '',
  editorsJournals: '',
  pptTemplate: '',
  outputDir: ''
})
const defaultPaths = reactive<Partial<SelectedPaths>>({})

const closeBehavior = ref<AppPreferences['closeBehavior']>('tray')
const focusJournalOrder = ref<string[]>([...DEFAULT_FOCUS_JOURNAL_ORDER])
const forceAeStaffInput = ref('')
const forceAeStaffTextarea = ref<HTMLTextAreaElement | null>(null)
const summaryOverrides = reactive<SummaryOverrides>({ ...DEFAULT_SUMMARY_OVERRIDES })
const running = ref(false)
const updateState = ref<AppUpdateState>({ phase: 'disabled', currentVersion: '', message: '正在读取版本信息…', generationRunning: false })
const updateActionPending = ref(false)
const settingsOpen = ref(false)
const logs = ref<PipelineProgressEvent[]>([])
const result = ref<PipelineResult | null>(null)
const toasts = ref<Toast[]>([])
const windowState = reactive<WindowState>({
  isAlwaysOnTop: false,
  isMaximized: false
})
const sourceField = PATH_FIELD_META.find((item) => item.key === 'mrWorkbook')!
const outputField = PATH_FIELD_META.find((item) => item.key === 'outputDir')!
const templateFields = PATH_FIELD_META.filter((item) => !['mrWorkbook', 'outputDir'].includes(item.key))
const templatesOpen = ref(false)
const logsOpen = ref(false)
const initialized = ref(false)
const runError = ref('')
const elapsedSeconds = ref(0)
const settingsDialog = ref<HTMLElement | null>(null)
const settingsTrigger = ref<HTMLButtonElement | null>(null)
const busy = computed(() => running.value || updateState.value.generationRunning || updateState.value.phase === 'installing')
const templateCount = computed(() => templateFields.filter((item) => paths[item.key]).length)
const latestLog = computed(() => logs.value.at(-1))
const reportMonth = computed(() => {
  const name = paths.mrWorkbook.split(/[\\/]/).at(-1) ?? ''
  const match = name.match(/^MR_\d{6}-(\d{4})(\d{2})\.xlsx$/i)
  return match && Number(match[2]) >= 1 && Number(match[2]) <= 12 ? `${match[1]}年${match[2]}月` : '选择 MR 后识别'
})
const outputFiles = computed(() => result.value ? [
  { label: '月会数据', path: result.value.outputs.monthlyWorkbook, icon: 'ri-file-excel-2-line' },
  { label: '人员数据', path: result.value.outputs.staffWorkbook, icon: 'ri-file-excel-2-line' },
  { label: 'PPT 成品', path: result.value.outputs.presentation, icon: 'ri-file-ppt-2-line' }
] : [])
const generatedDirectory = computed(() => result.value?.outputs.monthlyWorkbook.replace(/[/\\][^/\\]*$/, '') ?? '')
const statusText = computed(() => !initialized.value ? '正在加载设置' : updateState.value.phase === 'installing' ? '正在安装更新' : running.value ? '正在生成' : runError.value ? '生成失败' : missingFields.value.length ? '待补齐文件' : '就绪')
watch(templateCount, (count) => { if (count < templateFields.length) templatesOpen.value = true })
watch(settingsOpen, async (open) => {
  await nextTick()
  if (open) settingsDialog.value?.querySelector<HTMLButtonElement>('button')?.focus()
  else settingsTrigger.value?.focus()
})

function handleDialogKey(event: KeyboardEvent): void {
  if (event.key === 'Escape') { settingsOpen.value = false; return }
  if (event.key !== 'Tab') return
  const controls = Array.from(settingsDialog.value?.querySelectorAll<HTMLElement>('button:not(:disabled), input, textarea, [tabindex="0"]') ?? [])
  const first = controls[0], last = controls.at(-1)
  if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus() }
  else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus() }
}
const closeBehaviorOptions = [
  { label: '关闭到托盘', value: 'tray' },
  { label: '直接退出', value: 'quit' }
] as const

let toastSeed = 1
let unlisten: (() => void) | null = null
let unlistenWindowState: (() => void) | null = null
let unlistenUpdates: (() => void) | null = null

const missingFields = computed(() =>
  Object.entries(paths)
    .filter(([, value]) => !value)
    .map(([key]) => key)
)

function snapshotPaths(): SelectedPaths {
  return {
    mrWorkbook: paths.mrWorkbook,
    monthlyTemplate: paths.monthlyTemplate,
    staffTemplate: paths.staffTemplate,
    editorsJournals: paths.editorsJournals,
    pptTemplate: paths.pptTemplate,
    outputDir: paths.outputDir
  }
}

function snapshotSummaryOverrides(): SummaryOverrides {
  return {
    reportMonthLabel: summaryOverrides.reportMonthLabel,
    publication: summaryOverrides.publication,
    submission: summaryOverrides.submission,
    assignedManuscript: summaryOverrides.assignedManuscript,
    siSetUp: summaryOverrides.siSetUp,
    revenueWCHF: summaryOverrides.revenueWCHF,
    waiverRate: summaryOverrides.waiverRate,
    mpt: summaryOverrides.mpt
  }
}

function parseForceAeStaff(value: string): string[] {
  const seen = new Set<string>()
  return value
    .split(/[\n,，;；]+/)
    .map((item) => item.replace(/\s+/g, ' ').trim())
    .filter((item) => {
      const key = item.toLowerCase()
      if (!item || seen.has(key)) {
        return false
      }
      seen.add(key)
      return true
    })
}

function formatForceAeStaff(): string[] {
  const names = parseForceAeStaff(forceAeStaffInput.value)
  forceAeStaffInput.value = names.join('\n')
  resizeForceAeStaffTextarea()
  return names
}

function resizeForceAeStaffTextarea(): void {
  window.requestAnimationFrame(() => {
    const element = forceAeStaffTextarea.value
    if (!element) {
      return
    }
    element.style.height = 'auto'
    element.style.height = `${Math.max(112, element.scrollHeight)}px`
  })
}

function pushToast(text: string, tone: Toast['tone'] = 'info'): void {
  const id = toastSeed++
  toasts.value.push({ id, tone, text })
  window.setTimeout(() => {
    toasts.value = toasts.value.filter((toast) => toast.id !== id)
  }, 3200)
}

function applyPreferences(preferences: AppPreferences): void {
  Object.assign(paths, {
    mrWorkbook: preferences.paths.mrWorkbook ?? '',
    monthlyTemplate: preferences.paths.monthlyTemplate ?? '',
    staffTemplate: preferences.paths.staffTemplate ?? '',
    editorsJournals: preferences.paths.editorsJournals ?? '',
    pptTemplate: preferences.paths.pptTemplate ?? '',
    outputDir: preferences.paths.outputDir ?? ''
  })
  closeBehavior.value = preferences.closeBehavior
  focusJournalOrder.value = [...(preferences.focusJournalOrder?.length ? preferences.focusJournalOrder : DEFAULT_FOCUS_JOURNAL_ORDER)]
  forceAeStaffInput.value = [...(preferences.forceAeStaff?.length ? preferences.forceAeStaff : DEFAULT_FORCE_AE_STAFF)].join('\n')
  Object.assign(summaryOverrides, {
    ...DEFAULT_SUMMARY_OVERRIDES,
    ...(preferences.summaryOverrides ?? {})
  })
  void nextTick(resizeForceAeStaffTextarea)
}

async function pickPath(key: keyof SelectedPaths, title: string, filters: Array<{ name: string; extensions: string[] }>): Promise<void> {
  const chosen = key === 'outputDir'
    ? await window.electronApi.pickDirectory(title)
    : await window.electronApi.pickFile({ title, filters })

  if (chosen) {
    paths[key] = chosen
    try { await savePreferences(false) } catch { pushToast('路径已选择，但保存失败，请重试保存设置。', 'error') }
  }
}

async function handlePick(item: (typeof PATH_FIELD_META)[number]): Promise<void> {
  if (busy.value) return
  await pickPath(
    item.key,
    item.label,
    item.filters.map((filter) => ({
      name: filter.name,
      extensions: [...filter.extensions]
    }))
  )
}

async function handleReveal(key: keyof SelectedPaths): Promise<void> {
  if (paths[key]) {
    await window.electronApi.revealPath(paths[key])
  }
}

async function handleReset(key: keyof SelectedPaths): Promise<void> {
  if (busy.value) return
  paths[key] = defaultPaths[key] ?? ''
  try { await savePreferences(false) } catch { pushToast('路径保存失败，请重试保存设置。', 'error') }
}

async function revealOutput(targetPath: string): Promise<void> {
  await window.electronApi.revealPath(targetPath)
}

async function savePreferences(showToast = true): Promise<void> {
  const forceAeStaff = formatForceAeStaff()
  await window.electronApi.savePreferences({
    paths: snapshotPaths(),
    closeBehavior: closeBehavior.value,
    focusJournalOrder: [...focusJournalOrder.value],
    forceAeStaff,
    summaryOverrides: snapshotSummaryOverrides()
  })
  if (showToast) {
    pushToast('偏好设置已保存。', 'success')
  }
}

async function saveAndCloseSettings(): Promise<void> {
  try {
    await savePreferences()
    settingsOpen.value = false
  } catch { pushToast('设置保存失败，请重试。', 'error') }
}

async function refreshWindow(): Promise<void> {
  await window.electronApi.refreshWindow()
}

async function toggleAlwaysOnTop(): Promise<void> {
  Object.assign(windowState, await window.electronApi.toggleAlwaysOnTop())
}

async function minimizeWindow(): Promise<void> {
  await window.electronApi.minimizeWindow()
}

async function toggleMaximizeWindow(): Promise<void> {
  Object.assign(windowState, await window.electronApi.toggleMaximizeWindow())
}

async function closeWindow(): Promise<void> {
  await window.electronApi.closeWindow()
}

async function run(): Promise<void> {
  if (busy.value || !initialized.value) return
  if (missingFields.value.length) {
    if (templateCount.value < templateFields.length) templatesOpen.value = true
    pushToast('还有输入文件没选完，先把路径补齐。', 'error')
    return
  }

  running.value = true
  result.value = null
  logs.value = []
  logsOpen.value = false
  runError.value = ''
  const started = Date.now()

  try {
    await savePreferences(false)
    result.value = await window.electronApi.runPipeline({
      paths: snapshotPaths(),
      closeBehavior: closeBehavior.value,
      focusJournalOrder: [...focusJournalOrder.value],
      forceAeStaff: formatForceAeStaff(),
      summaryOverrides: snapshotSummaryOverrides()
    })
    pushToast('月会文件已经生成完成。', 'success')
  } catch (error) {
    runError.value = error instanceof Error ? error.message : '生成失败，请查看日志。'
    logsOpen.value = true
    pushToast(runError.value, 'error')
  } finally {
    elapsedSeconds.value = Math.max(1, Math.round((Date.now() - started) / 1000))
    running.value = false
  }
}

async function handleUpdate(action: 'check' | 'download' | 'install'): Promise<void> {
  if (updateActionPending.value) return
  updateActionPending.value = true
  try {
    // Save all visible preferences before restart, including edits outside the settings drawer.
    if (action === 'install') await savePreferences(false)
    updateState.value = await (action === 'check' ? window.electronApi.checkForUpdates()
      : action === 'download' ? window.electronApi.downloadUpdate() : window.electronApi.installUpdate())
  } catch (error) {
    pushToast(error instanceof Error ? error.message : '更新操作失败，请重试。', 'error')
  } finally {
    updateActionPending.value = false
  }
}

onMounted(async () => {
  try {
    unlistenUpdates = window.electronApi.onUpdateState((state) => {
      updateState.value = state
    })
    updateState.value = await window.electronApi.getUpdateState()
    Object.assign(defaultPaths, await window.electronApi.getDefaultPaths())
    Object.assign(windowState, await window.electronApi.getWindowState())
    const preferences = await window.electronApi.getPreferences()
    applyPreferences(preferences)
    templatesOpen.value = templateCount.value < templateFields.length
    initialized.value = true
    await nextTick()
    resizeForceAeStaffTextarea()
    unlisten = window.electronApi.onPipelineProgress((event) => {
      logs.value = [...logs.value, event]
      if (event.level === 'error' || event.level === 'warning') logsOpen.value = true
    })
    unlistenWindowState = window.electronApi.onWindowStateChange((state) => {
      Object.assign(windowState, state)
    })
  } catch {
    pushToast('初始化失败，请使用标题栏刷新重试。', 'error')
  }
})

onBeforeUnmount(() => {
  unlistenUpdates?.()
  unlisten?.()
  unlistenWindowState?.()
})
</script>

<template>
  <div class="app-shell">
    <header class="topbar">
      <div class="window-brand"><img :src="brandIconUrl" alt="" /><strong>DataDeck</strong></div>
      <div class="window-controls">
        <button class="titlebar-button" title="刷新" aria-label="刷新" :disabled="busy" @click="refreshWindow"><i aria-hidden="true" class="ri-refresh-line" /></button>
        <button class="titlebar-button" :class="{ active: windowState.isAlwaysOnTop }" :title="windowState.isAlwaysOnTop ? '取消置顶' : '窗口置顶'" aria-label="窗口置顶" :aria-pressed="windowState.isAlwaysOnTop" @click="toggleAlwaysOnTop"><i aria-hidden="true" class="ri-pushpin-line" /></button>
        <button class="titlebar-button" title="最小化" aria-label="最小化" @click="minimizeWindow"><i aria-hidden="true" class="ri-subtract-line" /></button>
        <button class="titlebar-button" :title="windowState.isMaximized ? '还原' : '最大化'" aria-label="最大化或还原" @click="toggleMaximizeWindow"><i aria-hidden="true" :class="windowState.isMaximized ? 'ri-checkbox-multiple-blank-line' : 'ri-checkbox-blank-line'" /></button>
        <button class="titlebar-button titlebar-button--close" title="关闭" aria-label="关闭" @click="closeWindow"><i aria-hidden="true" class="ri-close-line" /></button>
      </div>
    </header>

    <SimpleBarScroll class="workspace-scroll" :inert="settingsOpen">
      <main class="workspace">
        <header class="workspace-heading">
          <h1>月会文件</h1>
          <button ref="settingsTrigger" class="button button--subtle" @click="settingsOpen = true">
            <i aria-hidden="true" class="ri-settings-3-line" /><span>偏好设置</span>
            <span v-if="['available', 'downloaded'].includes(updateState.phase)" class="update-dot" aria-label="有可用更新" />
          </button>
        </header>

        <section class="panel source-panel" aria-label="MR 数据源">
          <PathField :label="sourceField.label" :description="sourceField.description" :value="paths.mrWorkbook" :default-value="defaultPaths.mrWorkbook" icon="ri-file-excel-2-line" :disabled="busy || !initialized" @pick="handlePick(sourceField)" @reset="handleReset('mrWorkbook')" @reveal="handleReveal('mrWorkbook')" />
          <p class="report-month"><span>报告月份</span><strong>{{ reportMonth }}</strong></p>
        </section>

        <section class="panel template-panel">
          <button class="disclosure template-toggle" :aria-expanded="templatesOpen" aria-controls="template-fields" @click="templatesOpen = !templatesOpen">
            <span><i aria-hidden="true" class="ri-stack-line" />模板与映射</span>
            <small :class="{ ready: templateCount === 4 }"><i aria-hidden="true" :class="templateCount === 4 ? 'ri-checkbox-circle-fill' : 'ri-information-line'" /> {{ templateCount }}/4 {{ templateCount === 4 ? '已配置' : '待补齐' }}</small>
            <i aria-hidden="true" :class="templatesOpen ? 'ri-arrow-up-s-line' : 'ri-arrow-down-s-line'" />
          </button>
          <div v-show="templatesOpen" id="template-fields" class="template-fields">
            <PathField v-for="item in templateFields" :key="item.key" compact :label="item.label" :description="item.description" :value="paths[item.key]" :default-value="defaultPaths[item.key]" :icon="item.key === 'pptTemplate' ? 'ri-file-ppt-2-line' : 'ri-file-excel-2-line'" :disabled="busy || !initialized" @pick="handlePick(item)" @reset="handleReset(item.key)" @reveal="handleReveal(item.key)" />
            <p class="hint">手动选择，路径自动记住</p>
          </div>
        </section>

        <section class="panel" aria-label="输出目录">
          <PathField :label="outputField.label" :description="outputField.description" :value="paths.outputDir" :default-value="defaultPaths.outputDir" icon="ri-folder-3-line" full-path :disabled="busy || !initialized" @pick="handlePick(outputField)" @reset="handleReset('outputDir')" @reveal="handleReveal('outputDir')" />
        </section>

        <div class="generate-actions">
          <button class="button button--primary generate-button" :disabled="busy || !initialized" @click="run">
            <i aria-hidden="true" :class="running ? 'ri-loader-4-line spin' : 'ri-play-fill'" />
            {{ running ? '正在生成…' : updateState.phase === 'installing' ? '正在安装更新…' : result ? '重新生成文件' : '生成月会文件' }}
          </button>
          <p class="hint">生成月会 Excel、人员 Excel 和 PPT</p>
        </div>

        <section class="panel progress-panel" aria-label="生成状态">
          <header v-if="running || result || runError" class="panel__head"><strong>本次生成</strong><span class="state-pill" :class="{ error: runError, pending: running }">{{ running ? '处理中' : runError ? '失败' : '已完成' }}</span></header>
          <div class="run-status" role="status" aria-live="polite">
            <i aria-hidden="true" :class="runError ? 'ri-error-warning-fill error-text' : running ? 'ri-loader-4-line spin' : result || !missingFields.length ? 'ri-checkbox-circle-fill ready' : 'ri-information-line'" />
            <div>
              <strong>{{ runError ? '生成失败，请检查后重试' : running ? latestLog?.step || '准备生成文件' : result ? '3 个文件已生成' : missingFields.length ? `还需选择 ${missingFields.length} 项文件或目录` : '已就绪，可以生成' }}</strong>
              <p v-if="runError" class="error-text" role="alert">{{ runError }}</p>
              <p v-else-if="running">{{ latestLog?.message || '正在读取输入文件，请稍候…' }}</p>
              <p v-else-if="result">用时 {{ elapsedSeconds }} 秒 · {{ result.detected.reportMonthLabel }}</p>
              <p v-else>生成后在此查看结果</p>
            </div>
          </div>
          <progress v-if="running" class="generation-progress" aria-label="文件生成进度" />
          <template v-if="logs.length || runError">
            <button class="disclosure log-toggle" :aria-expanded="logsOpen" aria-controls="process-log" @click="logsOpen = !logsOpen"><span>{{ logsOpen ? '收起处理日志' : '查看处理日志' }}</span><i aria-hidden="true" :class="logsOpen ? 'ri-arrow-up-s-line' : 'ri-arrow-down-s-line'" /></button>
            <SimpleBarScroll v-if="logsOpen" id="process-log" class="log-list" content-class="simplebar-stack">
              <p v-if="!logs.length" class="hint">{{ runError }}</p>
              <div v-for="(entry, index) in logs" :key="index" class="log-item" :class="entry.level"><span class="log-item__dot" /><div><strong>{{ entry.step }}</strong><p>{{ entry.message }}</p></div></div>
            </SimpleBarScroll>
          </template>
        </section>

        <section v-if="result" class="panel results-panel" aria-label="生成结果">
          <header class="panel__head"><strong>生成结果</strong></header>
          <div v-for="file in outputFiles" :key="file.label" class="result-row">
            <i aria-hidden="true" :class="file.icon" />
            <div :title="file.path"><strong>{{ file.label }}</strong><span>{{ file.icon === 'ri-file-ppt-2-line' ? 'PPT' : 'Excel' }}</span></div>
            <button class="icon-button" :aria-label="`打开${file.label}所在目录`" :title="file.path" @click="revealOutput(file.path)"><i aria-hidden="true" class="ri-folder-open-line" /></button>
          </div>
          <button class="button button--subtle open-output" @click="revealOutput(generatedDirectory)"><i aria-hidden="true" class="ri-folder-open-line" />打开输出目录</button>
          <details v-if="result.assumptions.length" class="result-notes"><summary>生成说明 · {{ result.assumptions.length }} 项</summary><ul><li v-for="item in result.assumptions" :key="item">{{ item }}</li></ul></details>
        </section>
      </main>
    </SimpleBarScroll>

    <footer class="statusbar"><span><span class="status-dot" :class="{ warning: missingFields.length || runError }" />{{ statusText }}</span><span>v{{ updateState.currentVersion || '—' }}</span></footer>

    <transition name="drawer-fade">
      <div v-if="settingsOpen" class="drawer-backdrop" @click="settingsOpen = false" />
    </transition>
    <transition name="drawer-slide">
      <aside v-if="settingsOpen" ref="settingsDialog" class="drawer" role="dialog" aria-modal="true" aria-label="偏好设置" @keydown="handleDialogKey">
        <header class="drawer__head">
          <div>
            <strong>偏好设置</strong>
            <p>关闭行为、重点刊映射和汇总覆盖值</p>
          </div>
          <button type="button" class="icon-button" aria-label="关闭偏好设置" @click="settingsOpen = false">
            <i aria-hidden="true" class="ri-close-line" />
          </button>
        </header>

        <SimpleBarScroll class="drawer__body">
          <section class="drawer-section">
            <UpdatePanel :state="updateState" :pending="updateActionPending" :running="running" @action="handleUpdate" />
          </section>
          <section class="drawer-section">
            <label class="field">
              <span>关闭行为</span>
              <CustomSelect v-model="closeBehavior" :options="[...closeBehaviorOptions]" />
            </label>
          </section>

          <section class="drawer-section">
            <header class="section-head">
              <span>重点刊映射</span>
              <small>用于图表页 8 / 10 / 12 / 14 / 16</small>
            </header>
            <div class="drawer-grid drawer-grid--single">
              <div v-for="(value, index) in focusJournalOrder" :key="index" class="field">
                <label>图表页 {{ index + 1 }}</label>
                <CustomSelect v-model="focusJournalOrder[index]" :options="FOCUS_JOURNAL_OPTIONS" />
              </div>
            </div>
          </section>

          <section class="drawer-section">
            <header class="section-head">
              <span>强制 AE 名单</span>
              <small>一行一个姓名；粘贴逗号分隔名单会自动整理</small>
            </header>
            <label class="field field--textarea">
              <span>Section Managing Editor 转入 AE</span>
              <textarea
                ref="forceAeStaffTextarea"
                v-model="forceAeStaffInput"
                placeholder="例如&#10;Dawn Shao&#10;Mike Liu"
                spellcheck="false"
                @input="resizeForceAeStaffTextarea"
                @blur="formatForceAeStaff"
              />
            </label>
          </section>

          <section class="drawer-section">
            <header class="section-head">
              <span>科室概览覆盖值</span>
              <small>留空则按 MR 自动汇总</small>
            </header>
            <div class="drawer-grid">
              <label class="field">
                <span>报告月份标题</span>
                <input v-model="summaryOverrides.reportMonthLabel" placeholder="例如 May 2026" />
              </label>
              <label class="field">
                <span>发文</span>
                <input v-model="summaryOverrides.publication" placeholder="篇数" />
              </label>
              <label class="field">
                <span>投稿</span>
                <input v-model="summaryOverrides.submission" placeholder="篇数" />
              </label>
              <label class="field">
                <span>Assigned</span>
                <input v-model="summaryOverrides.assignedManuscript" placeholder="篇数" />
              </label>
              <label class="field">
                <span>特刊上线</span>
                <input v-model="summaryOverrides.siSetUp" placeholder="数量" />
              </label>
              <label class="field">
                <span>营收 WCHF</span>
                <input v-model="summaryOverrides.revenueWCHF" placeholder="例如 127.77" />
              </label>
              <label class="field">
                <span>免费比例 %</span>
                <input v-model="summaryOverrides.waiverRate" placeholder="例如 39.82" />
              </label>
              <label class="field">
                <span>MPT</span>
                <input v-model="summaryOverrides.mpt" placeholder="天数" />
              </label>
            </div>
          </section>
        </SimpleBarScroll>

        <footer class="drawer__foot">
          <button type="button" class="button button--subtle" @click="settingsOpen = false">
            取消
          </button>
          <button type="button" class="button button--primary" @click="saveAndCloseSettings">
            保存设置
          </button>
        </footer>
      </aside>
    </transition>

    <div class="toast-wrap">
      <div v-for="toast in toasts" :key="toast.id" class="toast" :class="toast.tone">
        <i aria-hidden="true" :class="toast.tone === 'success' ? 'ri-checkbox-circle-line' : toast.tone === 'error' ? 'ri-error-warning-line' : 'ri-information-line'" />
        <span>{{ toast.text }}</span>
      </div>
    </div>
  </div>
</template>

<style scoped src="./styles/app.css"></style>

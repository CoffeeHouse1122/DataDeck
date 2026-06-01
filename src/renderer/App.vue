<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, reactive, ref } from 'vue'
import CustomSelect from './components/CustomSelect.vue'
import PathField from './components/PathField.vue'
import brandIconUrl from '../../build/icons/favicon-256x256.png'
import { FOCUS_JOURNAL_OPTIONS, PATH_FIELD_META } from '../shared/constants'
import {
  DEFAULT_FOCUS_JOURNAL_ORDER,
  DEFAULT_FORCE_AE_STAFF,
  DEFAULT_SUMMARY_OVERRIDES,
  type AppPreferences,
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
const summaryOverrides = reactive<SummaryOverrides>({ ...DEFAULT_SUMMARY_OVERRIDES })
const running = ref(false)
const settingsOpen = ref(false)
const logs = ref<PipelineProgressEvent[]>([])
const result = ref<PipelineResult | null>(null)
const toasts = ref<Toast[]>([])
const windowState = reactive<WindowState>({
  isAlwaysOnTop: false,
  isMaximized: false
})
const pathFields = PATH_FIELD_META
const closeBehaviorOptions = [
  { label: '关闭到托盘', value: 'tray' },
  { label: '直接退出', value: 'quit' }
] as const

let toastSeed = 1
let unlisten: (() => void) | null = null
let unlistenWindowState: (() => void) | null = null

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
  return names
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
}

async function pickPath(key: keyof SelectedPaths, title: string, filters: Array<{ name: string; extensions: string[] }>): Promise<void> {
  const chosen = key === 'outputDir'
    ? await window.electronApi.pickDirectory(title)
    : await window.electronApi.pickFile({ title, filters })

  if (chosen) {
    paths[key] = chosen
  }
}

async function handlePick(item: (typeof PATH_FIELD_META)[number]): Promise<void> {
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

function handleReset(key: keyof SelectedPaths): void {
  paths[key] = defaultPaths[key] ?? ''
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
  await savePreferences()
  settingsOpen.value = false
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
  if (missingFields.value.length) {
    pushToast('还有输入文件没选完，先把路径补齐。', 'error')
    return
  }

  running.value = true
  result.value = null
  logs.value = []

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
    pushToast(error instanceof Error ? error.message : '生成失败，请查看日志。', 'error')
  } finally {
    running.value = false
  }
}

onMounted(async () => {
  Object.assign(defaultPaths, await window.electronApi.getDefaultPaths())
  Object.assign(windowState, await window.electronApi.getWindowState())
  const preferences = await window.electronApi.getPreferences()
  applyPreferences(preferences)
  unlisten = window.electronApi.onPipelineProgress((event) => {
    logs.value = [...logs.value, event]
  })
  unlistenWindowState = window.electronApi.onWindowStateChange((state) => {
    Object.assign(windowState, state)
  })
})

onBeforeUnmount(() => {
  unlisten?.()
  unlistenWindowState?.()
})
</script>

<template>
  <div class="app-shell">
    <header class="topbar">
      <div class="brand">
        <div class="brand__mark">
          <img :src="brandIconUrl" alt="" />
        </div>
        <div>
          <strong>DataDeck</strong>
          <p>数据整理与月会文件生成</p>
        </div>
      </div>

      <div class="topbar__actions">
        <button type="button" class="button button--subtle" @click="settingsOpen = true">
          <i class="ri-settings-3-line" />
          <span>偏好设置</span>
        </button>
        <button type="button" class="button button--primary" :disabled="running" @click="run">
          <i :class="running ? 'ri-loader-4-line spin' : 'ri-play-circle-line'" />
          <span>{{ running ? '处理中…' : '生成月会文件' }}</span>
        </button>
      </div>

      <div class="window-controls">
        <button type="button" class="titlebar-button" title="刷新" @click="refreshWindow">
          <i class="ri-refresh-line" />
        </button>
        <button
          type="button"
          class="titlebar-button"
          :class="{ active: windowState.isAlwaysOnTop }"
          :title="windowState.isAlwaysOnTop ? '取消置顶' : '窗口置顶'"
          @click="toggleAlwaysOnTop"
        >
          <i :class="windowState.isAlwaysOnTop ? 'ri-pushpin-2-fill' : 'ri-pushpin-line'" />
        </button>
        <button type="button" class="titlebar-button" title="最小化" @click="minimizeWindow">
          <i class="ri-subtract-line" />
        </button>
        <button
          type="button"
          class="titlebar-button"
          :title="windowState.isMaximized ? '还原' : '最大化'"
          @click="toggleMaximizeWindow"
        >
          <i :class="windowState.isMaximized ? 'ri-checkbox-multiple-blank-line' : 'ri-checkbox-blank-line'" />
        </button>
        <button type="button" class="titlebar-button titlebar-button--close" title="关闭" @click="closeWindow">
          <i class="ri-close-line" />
        </button>
      </div>
    </header>

    <main class="workspace">
      <section class="hero panel">
        <div class="hero__copy">
          <!-- <span class="eyebrow">GitHub-style desktop workflow</span> -->
          <h1>把 MR、模板和输出步骤，收成一个能直接交付的桌面流程。</h1>
          <p>程序会整理 Excel、计算完成率、更新重点刊趋势，并复制 PPT 模板生成结果文件。</p>
        </div>
        <div class="hero__meta">
          <div>
            <strong>{{ result?.detected.reportKey ?? '----' }}</strong>
            <span>Report Key</span>
          </div>
          <div>
            <strong>{{ result?.detected.reportMonthLabel ?? '--' }}</strong>
            <span>月份标识</span>
          </div>
          <div>
            <strong>{{ missingFields.length === 0 ? 'Ready' : `${missingFields.length} Missing` }}</strong>
            <span>输入状态</span>
          </div>
        </div>
      </section>

      <section class="panel">
        <header class="panel__head">
          <span>输入文件</span>
          <small>路径会自动记住</small>
        </header>
        <div class="path-grid">
          <PathField
            v-for="item in pathFields"
            :key="item.key"
            :label="item.label"
            :description="item.description"
            :value="paths[item.key]"
            :default-value="defaultPaths[item.key] ?? ''"
            icon="ri-file-list-3-line"
            @pick="handlePick(item)"
            @reveal="handleReveal(item.key)"
            @reset="handleReset(item.key)"
          />
        </div>
      </section>

      <section class="content-grid">
        <section class="panel">
          <header class="panel__head">
            <span>处理日志</span>
            <small>逐步反馈执行进度</small>
          </header>
          <div class="log-list">
            <div v-if="logs.length === 0" class="empty">还没开始运行，先把输入文件选好。</div>
            <div v-for="(entry, index) in logs" :key="index" class="log-item" :class="entry.level">
              <div class="log-item__dot" />
              <div class="log-item__content">
                <strong>{{ entry.step }}</strong>
                <p>{{ entry.message }}</p>
              </div>
            </div>
          </div>
        </section>

        <section class="panel">
          <header class="panel__head">
            <span>输出结果</span>
            <small>生成后可直接定位</small>
          </header>
          <div v-if="!result" class="empty">这里会显示生成后的工作簿和 PPT。</div>
          <div v-else class="result-stack">
            <div class="result-card">
              <div class="result-card__body">
                <strong>月会数据</strong>
                <p>{{ result.outputs.monthlyWorkbook }}</p>
              </div>
              <button type="button" class="icon-button" @click="revealOutput(result.outputs.monthlyWorkbook)">
                <i class="ri-folder-open-line" />
              </button>
            </div>
            <div class="result-card">
              <div class="result-card__body">
                <strong>人员数据</strong>
                <p>{{ result.outputs.staffWorkbook }}</p>
              </div>
              <button type="button" class="icon-button" @click="revealOutput(result.outputs.staffWorkbook)">
                <i class="ri-folder-open-line" />
              </button>
            </div>
            <div class="result-card">
              <div class="result-card__body">
                <strong>PPT 成品</strong>
                <p>{{ result.outputs.presentation }}</p>
              </div>
              <button type="button" class="icon-button" @click="revealOutput(result.outputs.presentation)">
                <i class="ri-folder-open-line" />
              </button>
            </div>
            <div class="assumptions">
              <strong>当前实现假设</strong>
              <ul>
                <li v-for="item in result.assumptions" :key="item">{{ item }}</li>
              </ul>
            </div>
          </div>
        </section>
      </section>
    </main>

    <transition name="drawer-fade">
      <div v-if="settingsOpen" class="drawer-backdrop" @click="settingsOpen = false" />
    </transition>
    <transition name="drawer-slide">
      <aside v-if="settingsOpen" class="drawer">
        <header class="drawer__head">
          <div>
            <strong>偏好设置</strong>
            <p>关闭行为、重点刊映射和汇总覆盖值</p>
          </div>
          <button type="button" class="icon-button" @click="settingsOpen = false">
            <i class="ri-close-line" />
          </button>
        </header>

        <div class="drawer__body">
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
                v-model="forceAeStaffInput"
                placeholder="例如&#10;Dawn Shao&#10;Mike Liu"
                spellcheck="false"
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
        </div>

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
        <i :class="toast.tone === 'success' ? 'ri-checkbox-circle-line' : toast.tone === 'error' ? 'ri-error-warning-line' : 'ri-information-line'" />
        <span>{{ toast.text }}</span>
      </div>
    </div>
  </div>
</template>

<style scoped>
@font-face {
  font-family: NotoSansSC;
  src: url('./assets/NotoSansSC-Regular.woff2') format('woff2');
  font-weight: 400;
}

:global(*) {
  box-sizing: border-box;
}

:global(body) {
  margin: 0;
  font-family: NotoSansSC, 'Segoe UI', sans-serif;
  line-height: 1.2;
  background: #0d1117;
  color: #1f2328;
}

:global(button),
:global(input),
:global(textarea) {
  font: inherit;
}

:global(*::-webkit-scrollbar) {
  width: 10px;
  height: 10px;
}

:global(*::-webkit-scrollbar-track) {
  background: transparent;
}

:global(*::-webkit-scrollbar-thumb) {
  border: 2px solid transparent;
  border-radius: 999px;
  background: #c1c7d0;
  background-clip: padding-box;
}

:global(*::-webkit-scrollbar-thumb:hover) {
  background: #98a2ad;
  background-clip: padding-box;
}

:global(*) {
  scrollbar-width: thin;
  scrollbar-color: #c1c7d0 transparent;
}

.app-shell {
  --accent: #0aa19e;
  --accent-strong: hsla(179, 88%, 34%, 0.86);
  min-height: 100vh;
  background: #f6f8fa;
}

.topbar {
  height: 56px;
  padding: 0 14px;
  border-bottom: 1px solid #d0d7de;
  background: #0d1117;
  color: #f0f6fc;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  user-select: none;
  -webkit-app-region: drag;
}

.brand {
  display: flex;
  align-items: center;
  gap: 9px;
}

.brand__mark {
  width: 30px;
  height: 30px;
  border-radius: 7px;
  overflow: hidden;
  background: #0d1117;
  box-shadow: 0 0 0 1px rgba(240, 246, 252, 0.12);
}

.brand__mark img {
  display: block;
  width: 100%;
  height: 100%;
  object-fit: cover;
}

.brand p {
  margin: 2px 0 0;
  font-size: 10px;
  color: #8b949e;
}

.topbar__actions {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-left: auto;
  -webkit-app-region: no-drag;
}

.window-controls {
  display: flex;
  align-items: center;
  gap: 2px;
  height: 100%;
  -webkit-app-region: no-drag;
}

.titlebar-button {
  width: 34px;
  height: 32px;
  border: 0;
  border-radius: 6px;
  background: transparent;
  color: #c9d1d9;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  -webkit-app-region: no-drag;
}

.titlebar-button i {
  font-size: 16px;
}

.titlebar-button:hover,
.titlebar-button.active {
  background: #21262d;
  color: #ffffff;
}

.titlebar-button.active {
  color: #7ee7e4;
}

.titlebar-button--close:hover {
  background: #da3633;
  color: #ffffff;
}

.workspace {
  height: calc(100vh - 56px);
  padding: 12px;
  display: grid;
  grid-template-rows: auto auto minmax(0, 1fr);
  gap: 10px;
}

.panel {
  border: 1px solid #d0d7de;
  border-radius: 8px;
  background: #fff;
  padding: 12px;
  min-height: 0;
  min-width: 0;
  overflow: hidden;
}

.hero {
  display: flex;
  justify-content: space-between;
  gap: 14px;
  align-items: flex-start;
}

.hero__copy {
  min-width: 0;
}

.hero h1 {
  margin: 5px 0 6px;
  max-width: 580px;
  font-size: 21px;
  line-height: 1.18;
}

.hero p,
.eyebrow,
.panel__head small,
.empty,
.result-card p,
.assumptions li,
.section-head small,
.drawer__head p {
  color: #57606a;
}

.hero p {
  font-size: 12px;
}

.eyebrow {
  font-size: 10px;
  font-weight: 700;
}

.hero__meta {
  min-width: 230px;
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 8px;
}

.hero__meta div {
  min-width: 0;
  border: 1px solid #d0d7de;
  border-radius: 8px;
  background: #f6f8fa;
  padding: 10px;
}

.hero__meta strong {
  display: block;
  font-size: 15px;
  color: #1f2328;
}

.hero__meta span {
  display: block;
  margin-top: 4px;
  font-size: 10px;
  line-height: 1.2;
  color: #57606a;
}

.panel__head,
.section-head,
.drawer__head {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 8px;
}

.panel__head {
  margin-bottom: 10px;
}

.panel__head span,
.section-head span,
.drawer__head strong {
  font-size: 13px;
  font-weight: 700;
}

.panel__head small,
.section-head small {
  font-size: 10px;
}

.path-grid {
  display: grid;
  grid-template-columns: repeat(6, minmax(0, 1fr));
  gap: 6px;
}

@media (max-width: 1180px) {
  .path-grid {
    grid-template-columns: repeat(3, minmax(0, 1fr));
  }
}

.content-grid {
  min-height: 0;
  display: grid;
  grid-template-columns: minmax(0, 1fr) 360px;
  gap: 10px;
}

.content-grid > .panel {
  display: flex;
  flex-direction: column;
  min-width: 0;
  min-height: 0;
  overflow: hidden;
}

.log-list,
.result-stack {
  flex: 1;
  min-height: 0;
  min-width: 0;
  max-height: 100%;
  overflow: auto;
  display: grid;
  gap: 8px;
  align-content: start;
}

.log-item {
  display: grid;
  grid-template-columns: 8px minmax(0, 1fr);
  gap: 9px;
  padding: 9px 10px;
  border: 1px solid #d0d7de;
  border-radius: 8px;
  background: #f6f8fa;
  min-width: 0;
  max-width: 100%;
}

.log-item__dot {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  margin-top: 4px;
  background: #2f81f7;
}

.log-item.success .log-item__dot {
  background: #1a7f37;
}

.log-item.warning .log-item__dot {
  background: #9a6700;
}

.log-item.error .log-item__dot {
  background: #cf222e;
}

.log-item__content {
  min-width: 0;
}

.log-item strong {
  display: block;
  font-size: 11px;
  line-height: 1.3;
  overflow-wrap: anywhere;
  word-break: break-word;
}

.log-item p {
  margin: 4px 0 0;
  font-size: 11px;
  line-height: 1.3;
  overflow-wrap: anywhere;
  word-break: break-word;
  white-space: normal;
}

.result-card {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  align-items: flex-start;
  gap: 10px;
  padding: 10px;
  border: 1px solid #d0d7de;
  border-radius: 8px;
  background: #f6f8fa;
  min-width: 0;
  max-width: 100%;
}

.result-card__body {
  min-width: 0;
  flex: 1;
}

.result-card strong {
  display: block;
  margin-bottom: 4px;
  font-size: 12px;
  line-height: 1.3;
  overflow-wrap: anywhere;
  word-break: break-word;
}

.result-card p {
  margin: 0;
  font-size: 11px;
  line-height: 1.3;
  overflow-wrap: anywhere;
  word-break: break-word;
  white-space: normal;
}

.assumptions {
  padding: 10px;
  border: 1px solid #d0d7de;
  border-radius: 8px;
  background: #fff;
  min-width: 0;
  max-width: 100%;
}

.assumptions strong {
  display: block;
  margin-bottom: 8px;
  font-size: 12px;
}

.assumptions ul {
  margin: 0;
  padding-left: 18px;
  min-width: 0;
}

.assumptions li {
  font-size: 11px;
  line-height: 1.35;
  overflow-wrap: anywhere;
  word-break: break-word;
}

.button,
.icon-button {
  border-radius: 6px;
  border: 1px solid transparent;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 7px;
  -webkit-app-region: no-drag;
}

.button {
  padding: 8px 11px;
  font-size: 11px;
  font-weight: 700;
  white-space: nowrap;
}

.button--primary {
  background: var(--accent);
  border-color: var(--accent-strong);
  color: #fff;
}

.button--primary:hover:not(:disabled) {
  background: var(--accent-strong);
}

.button--subtle {
  background: #161b22;
  color: #c9d1d9;
  border-color: #30363d;
}

.button:disabled {
  opacity: 0.72;
  cursor: wait;
}

.icon-button {
  width: 28px;
  height: 28px;
  flex: none;
  border: 1px solid #d0d7de;
  background: #fff;
  color: #57606a;
}

.empty {
  min-height: 108px;
  display: grid;
  place-items: center;
  text-align: center;
  border: 1px dashed #d0d7de;
  border-radius: 8px;
  font-size: 11px;
  line-height: 1.35;
  padding: 12px;
}

.drawer-backdrop {
  position: fixed;
  inset: 0;
  background: rgba(13, 17, 23, 0.34);
  z-index: 50;
}

.drawer {
  position: fixed;
  top: 0;
  right: 0;
  width: min(396px, 100vw);
  height: 100vh;
  background: #ffffff;
  border-left: 1px solid #d0d7de;
  z-index: 60;
  display: grid;
  grid-template-rows: auto minmax(0, 1fr) auto;
}

.drawer__head {
  padding: 14px 14px 12px;
  border-bottom: 1px solid #d0d7de;
}

.drawer__head p {
  margin: 5px 0 0;
  font-size: 11px;
}

.drawer__body {
  min-height: 0;
  overflow: auto;
  padding-bottom: 10px;
}

.drawer-section {
  padding: 12px 14px 0;
}

.drawer-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 10px;
  margin-top: 10px;
}

.drawer-grid--single {
  grid-template-columns: 1fr;
}

.field {
  display: grid;
  gap: 6px;
}

.field label,
.field span {
  font-size: 11px;
  font-weight: 700;
  color: #1f2328;
}

.field input,
.field textarea {
  width: 100%;
  border: 1px solid #d0d7de;
  background: #fff;
  color: #1f2328;
  border-radius: 6px;
  padding: 9px 10px;
  line-height: 1.2;
}

.field textarea {
  min-height: 96px;
  max-height: 160px;
  resize: vertical;
  overflow: auto;
  white-space: pre-wrap;
}

.field--textarea {
  margin-top: 10px;
}

.drawer__foot {
  padding: 12px 14px;
  border-top: 1px solid #d0d7de;
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  background: #fff;
}

.toast-wrap {
  position: fixed;
  right: 18px;
  bottom: 18px;
  display: grid;
  gap: 10px;
  z-index: 80;
}

.toast {
  position: relative;
  min-width: 280px;
  max-width: min(420px, calc(100vw - 36px));
  padding: 12px 14px 12px 16px;
  border-radius: 8px;
  border: 1px solid #8c959f;
  background: #ffffff;
  color: #1f2328;
  display: flex;
  align-items: center;
  gap: 10px;
  font-size: 12px;
  line-height: 1.35;
  font-weight: 700;
  box-shadow: 0 18px 44px rgba(31, 35, 40, 0.22), 0 0 0 1px rgba(31, 35, 40, 0.04);
}

.toast::before {
  content: '';
  position: absolute;
  inset: 0 auto 0 0;
  width: 5px;
  border-radius: 8px 0 0 8px;
  background: #57606a;
}

.toast i {
  width: 22px;
  height: 22px;
  border-radius: 999px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  flex: none;
  font-size: 17px;
  background: #f6f8fa;
  color: #57606a;
}

.toast span {
  overflow-wrap: anywhere;
  word-break: break-word;
}

.toast.success {
  border-color: #0aa19e;
  background: #f0fffd;
}

.toast.success::before {
  background: var(--accent);
}

.toast.success i {
  background: rgba(10, 161, 158, 0.12);
  color: #087f7c;
}

.toast.error {
  border-color: #cf222e;
  background: #fff5f5;
}

.toast.error::before {
  background: #cf222e;
}

.toast.error i {
  background: rgba(207, 34, 46, 0.1);
  color: #cf222e;
}

.toast.info {
  border-color: #2f81f7;
  background: #f1f8ff;
}

.toast.info::before {
  background: #2f81f7;
}

.toast.info i {
  background: rgba(47, 129, 247, 0.12);
  color: #2f81f7;
}

.spin {
  animation: spin 1s linear infinite;
}

.drawer-fade-enter-active,
.drawer-fade-leave-active {
  transition: opacity 0.18s ease;
}

.drawer-fade-enter-from,
.drawer-fade-leave-to {
  opacity: 0;
}

.drawer-slide-enter-active,
.drawer-slide-leave-active {
  transition: transform 0.2s ease;
}

.drawer-slide-enter-from,
.drawer-slide-leave-to {
  transform: translateX(100%);
}

@keyframes spin {
  from {
    transform: rotate(0deg);
  }
  to {
    transform: rotate(360deg);
  }
}
</style>

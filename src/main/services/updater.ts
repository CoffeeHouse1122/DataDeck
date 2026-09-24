import type { AppUpdater } from 'electron-updater'
import type { AppUpdateState } from '../../shared/contracts'

type UpdateClient = Pick<AppUpdater, 'on' | 'autoDownload' | 'autoInstallOnAppQuit' | 'allowPrerelease' | 'allowDowngrade' | 'checkForUpdates' | 'downloadUpdate' | 'quitAndInstall'>

/** Main-process authority for updates and generation/install mutual exclusion. */
export class UpdateController {
  private state: AppUpdateState
  private timer: ReturnType<typeof setTimeout> | undefined

  constructor(
    private readonly client: UpdateClient,
    enabled: boolean,
    version: string,
    private readonly publish: (state: AppUpdateState) => void,
    private readonly setQuitting: (value: boolean) => void
  ) {
    this.state = {
      phase: enabled ? 'idle' : 'disabled', currentVersion: version, generationRunning: false,
      message: enabled ? '尚未检查更新' : '开发模式或当前平台不检查更新'
    }
    client.autoDownload = false
    client.autoInstallOnAppQuit = false
    client.allowPrerelease = false
    client.allowDowngrade = false
    if (!enabled) return
    client.on('update-available', (info) => this.change({ phase: 'available', availableVersion: info.version, message: `发现新版本 ${info.version}，可下载更新` }))
    client.on('update-not-available', () => this.change({ phase: 'idle', availableVersion: undefined, message: '当前已是最新版本' }))
    client.on('download-progress', (progress) => {
      const percent = Math.max(0, Math.min(100, Number.isFinite(progress.percent) ? progress.percent : 0))
      this.change({ phase: 'downloading', percent, message: `正在下载 ${percent.toFixed(1)}%` })
    })
    client.on('update-downloaded', (info) => this.change({ phase: 'downloaded', availableVersion: info.version, percent: 100, message: `${info.version} 已下载，重启后安装` }))
    // Do not expose raw provider responses/URLs; retry starts with a fresh version check.
    client.on('error', () => this.fail())
  }

  getState(): AppUpdateState { return { ...this.state } }

  private change(next: Partial<AppUpdateState>): void {
    this.state = { ...this.state, ...next }
    this.publish(this.getState())
  }

  private fail(): void {
    if (this.state.phase === 'installing') this.setQuitting(false)
    this.change({ phase: 'error', percent: undefined, message: '更新失败，请检查网络或稍后重新检查更新。' })
  }

  start(): void {
    if (this.state.phase === 'disabled' || this.timer) return
    this.timer = setTimeout(() => { this.timer = undefined; void this.check() }, 12_000)
    this.timer.unref()
  }

  stop(): void {
    if (this.timer) clearTimeout(this.timer)
    this.timer = undefined
  }

  async check(): Promise<AppUpdateState> {
    // Ignore duplicate checks and preserve a downloaded installer until explicitly installed.
    if (!['idle', 'error', 'available'].includes(this.state.phase)) return this.getState()
    this.change({ phase: 'checking', percent: undefined, availableVersion: undefined, message: '正在检查更新…' })
    try { await this.client.checkForUpdates() } catch { this.fail() }
    return this.getState()
  }

  async download(): Promise<AppUpdateState> {
    if (this.state.phase !== 'available') return this.getState()
    this.change({ phase: 'downloading', percent: 0, message: '正在下载 0.0%' })
    try { await this.client.downloadUpdate() } catch { this.fail() }
    return this.getState()
  }

  install(): AppUpdateState {
    if (this.state.generationRunning) throw new Error('报表正在生成，请完成后再安装更新。')
    if (this.state.phase !== 'downloaded') throw new Error('请先下载更新。')
    // Set this before quitAndInstall so the close-to-tray handler cannot swallow exit.
    this.change({ phase: 'installing', message: '正在退出并安装更新…' })
    this.setQuitting(true)
    try { this.client.quitAndInstall(false, true) } catch { this.fail() }
    return this.getState()
  }

  async runGeneration<T>(operation: () => Promise<T>): Promise<T> {
    if (this.state.phase === 'installing') throw new Error('正在安装更新，请重启后生成报表。')
    if (this.state.generationRunning) throw new Error('已有报表正在生成。')
    // Acquire before the first await; release even when preference saving or generation fails.
    this.change({ generationRunning: true })
    try { return await operation() } finally { this.change({ generationRunning: false }) }
  }
}

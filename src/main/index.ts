import { app, BrowserWindow, dialog, globalShortcut, ipcMain, Menu, nativeImage, screen, shell, Tray } from 'electron'
import fs from 'node:fs'
import path from 'node:path'
import electronUpdater from 'electron-updater'
import type { AppPreferences, FilePickerOptions, PipelineInput, SelectedPaths, WindowState } from '../shared/contracts'
import { loadPreferences, savePreferences } from './services/settings'
import { runPipeline } from './services/pipeline'
import { UpdateController } from './services/updater'

let mainWindow: BrowserWindow | null = null
let tray: Tray | null = null
let quitting = false
let cachedPreferences: AppPreferences | null = null
let cachedAppIcon: Electron.NativeImage | null = null
let updates: UpdateController

const APP_ICON_SIZES = [16, 32, 48, 64, 128, 192, 256, 512]

function stringifyError(error: unknown): string {
  if (error instanceof Error) {
    return error.message
  }
  if (typeof error === 'string') {
    return error
  }
  try {
    return JSON.stringify(error)
  } catch {
    return 'Unknown error'
  }
}

function createFallbackIcon(): Electron.NativeImage {
  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32">
      <rect width="32" height="32" rx="7" fill="#0d1117"/>
      <path d="M8 9h16v3H8zm0 5h11v3H8zm0 5h16v3H8z" fill="#f0f6fc"/>
      <circle cx="23.5" cy="15.5" r="2.5" fill="#2f81f7"/>
    </svg>
  `
  return nativeImage.createFromDataURL(`data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`)
}

function addIconRepresentation(icon: Electron.NativeImage, fileName: string, size: number): void {
  const iconPath = appIconPath(fileName)
  if (!fs.existsSync(iconPath)) {
    return
  }

  const dataUrl = `data:image/png;base64,${fs.readFileSync(iconPath).toString('base64')}`
  icon.addRepresentation({
    dataURL: dataUrl,
    height: size,
    width: size
  })
}

function createAppIcon(): Electron.NativeImage {
  if (cachedAppIcon && !cachedAppIcon.isEmpty()) {
    return cachedAppIcon
  }

  const icon = nativeImage.createFromPath(appIconPath('favicon.ico'))
  for (const size of APP_ICON_SIZES) {
    addIconRepresentation(icon, `favicon-${size}x${size}.png`, size)
  }

  cachedAppIcon = icon.isEmpty() ? createFallbackIcon() : icon
  return cachedAppIcon
}

function appIconPath(fileName: string): string {
  return app.isPackaged
    ? path.join(process.resourcesPath, 'icons', fileName)
    : path.join(process.cwd(), 'resources', 'icons', fileName)
}

function windowIconPath(): string {
  return appIconPath('favicon.ico')
}

function retiredTemplatePath(fileName: string): string {
  return app.isPackaged
    ? path.join(process.resourcesPath, 'templates', fileName)
    : path.join(process.cwd(), 'docs', fileName)
}

// These locations are used only to clear obsolete bundled paths on upgrade.
// Templates are now selected by the user and never loaded automatically.
function retiredTemplatePaths(): Partial<SelectedPaths> {
  return {
    monthlyTemplate: retiredTemplatePath('monthly-data-generated-202603.xlsx'),
    staffTemplate: retiredTemplatePath('staff-data-generated-202603.xlsx'),
    editorsJournals: retiredTemplatePath('editors-journals.xlsx'),
    pptTemplate: retiredTemplatePath('Section Health月会.pptx')
  }
}

async function currentPreferences(): Promise<AppPreferences> {
  if (!cachedPreferences) {
    cachedPreferences = await loadPreferences(app.getPath('userData'), retiredTemplatePaths())
  }
  return cachedPreferences
}

function getWindowState(): WindowState {
  return {
    isAlwaysOnTop: Boolean(mainWindow?.isAlwaysOnTop()),
    isMaximized: Boolean(mainWindow?.isMaximized())
  }
}

function sendWindowState(): void {
  mainWindow?.webContents.send('window:state-changed', getWindowState())
}

async function createWindow(): Promise<void> {
  const { width, height } = screen.getPrimaryDisplay().workAreaSize
  mainWindow = new BrowserWindow({
    width: Math.min(430, width),
    height: Math.min(760, height),
    resizable: false,
    maximizable: false,
    fullscreenable: false,
    title: 'DataDeck',
    frame: false,
    icon: windowIconPath(),
    backgroundColor: '#f4f7fa',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, '../preload/index.mjs'),
      sandbox: false
    }
  })
  mainWindow.setIcon(windowIconPath())

  mainWindow.on('ready-to-show', () => {
    mainWindow?.show()
    sendWindowState()
  })

  mainWindow.on('maximize', sendWindowState)
  mainWindow.on('unmaximize', sendWindowState)
  mainWindow.on('enter-full-screen', sendWindowState)
  mainWindow.on('leave-full-screen', sendWindowState)

  mainWindow.on('close', async (event) => {
    if (quitting) {
      return
    }
    const preferences = await currentPreferences()
    if (preferences.closeBehavior === 'tray') {
      event.preventDefault()
      mainWindow?.hide()
    }
  })

  mainWindow.webContents.setWindowOpenHandler((details) => {
    shell.openExternal(details.url)
    return { action: 'deny' }
  })

  if (!app.isPackaged && process.env.ELECTRON_RENDERER_URL) {
    await mainWindow.loadURL(process.env.ELECTRON_RENDERER_URL)
  } else {
    await mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'))
  }
}

function createTray(): void {
  tray = new Tray(createAppIcon())
  tray.setToolTip('DataDeck')
  tray.on('double-click', () => {
    mainWindow?.show()
    mainWindow?.focus()
  })

  const menu = Menu.buildFromTemplate([
    {
      label: '显示窗口',
      click: () => {
        mainWindow?.show()
        mainWindow?.focus()
      }
    },
    {
      label: '退出',
      click: () => {
        quitting = true
        app.quit()
      }
    }
  ])

  tray.setContextMenu(menu)
}

function registerIpc(): void {
  ipcMain.handle('update:state', () => updates.getState())
  ipcMain.handle('update:check', () => updates.check())
  ipcMain.handle('update:download', () => updates.download())
  ipcMain.handle('update:install', () => updates.install())
  ipcMain.handle('pick:file', async (_event, options: FilePickerOptions) => {
    const result = await dialog.showOpenDialog(mainWindow!, {
      title: options.title,
      properties: ['openFile'],
      filters: options.filters
    })
    return result.canceled ? null : result.filePaths[0]
  })

  ipcMain.handle('pick:directory', async (_event, title: string) => {
    const result = await dialog.showOpenDialog(mainWindow!, {
      title,
      properties: ['openDirectory', 'createDirectory']
    })
    return result.canceled ? null : result.filePaths[0]
  })

  ipcMain.handle('preferences:get', async () => currentPreferences())

  ipcMain.handle('paths:defaults', async () => ({}))

  ipcMain.handle('preferences:save', async (_event, preferences: AppPreferences) => {
    cachedPreferences = preferences
    await savePreferences(app.getPath('userData'), preferences)
  })

  ipcMain.handle('pipeline:run', async (event, input: PipelineInput) => updates.runGeneration(async () => {
    cachedPreferences = {
      paths: input.paths,
      closeBehavior: input.closeBehavior,
      focusJournalOrder: input.focusJournalOrder,
      forceAeStaff: input.forceAeStaff ?? [],
      summaryOverrides: input.summaryOverrides
    }
    await savePreferences(app.getPath('userData'), cachedPreferences)
    try {
      return await runPipeline(event, input)
    } catch (error) {
      throw new Error(stringifyError(error))
    }
  }))

  ipcMain.handle('path:reveal', async (_event, targetPath: string) => {
    shell.showItemInFolder(targetPath)
  })

  ipcMain.handle('window:refresh', () => {
    mainWindow?.reload()
  })

  ipcMain.handle('window:toggle-always-on-top', () => {
    const nextValue = !mainWindow?.isAlwaysOnTop()
    mainWindow?.setAlwaysOnTop(nextValue)
    sendWindowState()
    return getWindowState()
  })

  ipcMain.handle('window:minimize', () => {
    mainWindow?.minimize()
  })

  ipcMain.handle('window:toggle-maximize', () => {
    // Keep the existing IPC contract without allowing the fixed-size window to maximize.
    return getWindowState()
  })

  ipcMain.handle('window:close', () => {
    mainWindow?.close()
  })

  ipcMain.handle('window:state', () => getWindowState())
}

const lock = app.requestSingleInstanceLock()
if (!lock) {
  app.quit()
} else {
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) {
        mainWindow.restore()
      }
      mainWindow.show()
      mainWindow.focus()
    }
  })

  app.whenReady().then(async () => {
    app.setAppUserModelId('com.zhero.datadeck')

    updates = new UpdateController(electronUpdater.autoUpdater, app.isPackaged && process.platform === 'win32', app.getVersion(), (state) => {
      for (const window of BrowserWindow.getAllWindows()) {
        if (!window.isDestroyed()) window.webContents.send('update:state-changed', state)
      }
    }, (value) => { quitting = value })
    registerIpc()
    await createWindow()
    createTray()
    updates.start()
    globalShortcut.register('CommandOrControl+R', () => {
      if (!app.isPackaged) {
        mainWindow?.reload()
      }
    })
  })

  app.on('before-quit', () => {
    quitting = true
  })

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
      app.quit()
    }
  })

  app.on('will-quit', () => {
    updates?.stop()
    globalShortcut.unregisterAll()
  })

  app.on('activate', async () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      await createWindow()
    } else {
      mainWindow?.show()
    }
  })
}

import { contextBridge, ipcRenderer } from 'electron'
import type { AppPreferences, ElectronApi, FilePickerOptions, PipelineInput, PipelineProgressEvent, PipelineResult, WindowState } from '../../shared/contracts'

const electronApi: ElectronApi = {
  pickFile: (options: FilePickerOptions) => ipcRenderer.invoke('pick:file', options),
  pickDirectory: (title: string) => ipcRenderer.invoke('pick:directory', title),
  getPreferences: () => ipcRenderer.invoke('preferences:get') as Promise<AppPreferences>,
  getDefaultPaths: () => ipcRenderer.invoke('paths:defaults'),
  savePreferences: (preferences: AppPreferences) => ipcRenderer.invoke('preferences:save', preferences),
  runPipeline: (input: PipelineInput) => ipcRenderer.invoke('pipeline:run', input) as Promise<PipelineResult>,
  revealPath: (targetPath: string) => ipcRenderer.invoke('path:reveal', targetPath),
  refreshWindow: () => ipcRenderer.invoke('window:refresh'),
  toggleAlwaysOnTop: () => ipcRenderer.invoke('window:toggle-always-on-top') as Promise<WindowState>,
  minimizeWindow: () => ipcRenderer.invoke('window:minimize'),
  toggleMaximizeWindow: () => ipcRenderer.invoke('window:toggle-maximize') as Promise<WindowState>,
  closeWindow: () => ipcRenderer.invoke('window:close'),
  getWindowState: () => ipcRenderer.invoke('window:state') as Promise<WindowState>,
  onWindowStateChange: (listener: (state: WindowState) => void) => {
    const handler = (_event: Electron.IpcRendererEvent, payload: WindowState) => listener(payload)
    ipcRenderer.on('window:state-changed', handler)
    return () => {
      ipcRenderer.removeListener('window:state-changed', handler)
    }
  },
  onPipelineProgress: (listener: (event: PipelineProgressEvent) => void) => {
    const handler = (_event: Electron.IpcRendererEvent, payload: PipelineProgressEvent) => listener(payload)
    ipcRenderer.on('pipeline:progress', handler)
    return () => {
      ipcRenderer.removeListener('pipeline:progress', handler)
    }
  }
}

contextBridge.exposeInMainWorld('electronApi', electronApi)

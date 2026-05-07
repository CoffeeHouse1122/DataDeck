import { contextBridge, ipcRenderer } from 'electron'
import type { AppPreferences, ElectronApi, FilePickerOptions, PipelineInput, PipelineProgressEvent, PipelineResult } from '../../shared/contracts'

const electronApi: ElectronApi = {
  pickFile: (options: FilePickerOptions) => ipcRenderer.invoke('pick:file', options),
  pickDirectory: (title: string) => ipcRenderer.invoke('pick:directory', title),
  getPreferences: () => ipcRenderer.invoke('preferences:get') as Promise<AppPreferences>,
  savePreferences: (preferences: AppPreferences) => ipcRenderer.invoke('preferences:save', preferences),
  runPipeline: (input: PipelineInput) => ipcRenderer.invoke('pipeline:run', input) as Promise<PipelineResult>,
  revealPath: (targetPath: string) => ipcRenderer.invoke('path:reveal', targetPath),
  onPipelineProgress: (listener: (event: PipelineProgressEvent) => void) => {
    const handler = (_event: Electron.IpcRendererEvent, payload: PipelineProgressEvent) => listener(payload)
    ipcRenderer.on('pipeline:progress', handler)
    return () => {
      ipcRenderer.removeListener('pipeline:progress', handler)
    }
  }
}

contextBridge.exposeInMainWorld('electronApi', electronApi)

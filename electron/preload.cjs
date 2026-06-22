const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  request: (config) => ipcRenderer.invoke('http-request', config),

  downloadFile: (config) =>
      ipcRenderer.invoke('download-file', config),

  onDownloadStatus: (callback) => {
    ipcRenderer.on(
      'download-status',
      (_, data) => callback(data)
    )
  },

  removeDownloadStatus: () => {
    ipcRenderer.removeAllListeners(
      'download-status'
    )
  }
})
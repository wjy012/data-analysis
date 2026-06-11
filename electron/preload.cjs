const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  send: (channel, data) => ipcRenderer.send(channel, data),
  request: (config) => ipcRenderer.invoke('http-request', config),
  downloadFile: (config) =>
      ipcRenderer.invoke('download-file', config)
})
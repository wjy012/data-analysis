import { app, BrowserWindow, ipcMain, dialog } from 'electron'
import path from 'path'
import axios from 'axios'
import fs from 'fs'
import { fileURLToPath } from 'url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

function createWindow() {
  const win = new BrowserWindow({
    width: 1000,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs')
    }
  })

  if (app.isPackaged) {
    win.loadFile(
      path.join(__dirname, '../dist/index.html')
    )
  } else {
    win.loadURL('http://localhost:5173')
    win.webContents.openDevTools()
  }
}

ipcMain.handle('http-request', async (_, config) => {
  try {
    const response = await axios({
      headers: {
        Accept: '*/*',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Content-Type': 'application/json',
        noRedirect: '1',
        zjarkLang: 'zh_CN',
        token: config.token || '' // 从请求配置中获取 token，默认为空字符串 
      },
      ...config
    })

    return {
      success: true,
      data: response.data,
      status: response.status
    }
  } catch (error) {
    return {
      success: false,
      message: error.message,
      status: error.response?.status,
      data: error.response?.data
    }
  }
})

ipcMain.handle('download-file', async (_, config) => {
  try {
    const { url, fileName } = config

    const savePath = dialog.showSaveDialogSync({
      defaultPath: fileName
    })

    if (!savePath) {
      return {
        success: false,
        message: '用户取消下载'
      }
    }

    const response = await axios({
      url,
      method: 'GET',
      responseType: 'arraybuffer'
    })

    fs.writeFileSync(savePath, response.data)

    return {
      success: true,
      path: savePath
    }
  } catch (error) {
    return {
      success: false,
      message: error.message
    }
  }
})

app.whenReady().then(() => {
  createWindow()
})
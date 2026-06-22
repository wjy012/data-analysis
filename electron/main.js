import { app, BrowserWindow, ipcMain, dialog } from 'electron'
import path from 'path'
import axios from 'axios'
import fs from 'fs'
import { fileURLToPath } from 'url'
import { processSrmData, processPurchasePlan, processOrderFrameData } from '../src/utils/processExcel.js'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const bufferCache = new Map()

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
    win.webContents.openDevTools()
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

ipcMain.handle('download-file', async (event, config) => {
  try {
    const { url, fileName, taskType, method = 'GET', data } = config

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
      headers: {
        Accept: '*/*',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Content-Type': 'application/json',
        noRedirect: '1',
        zjarkLang: 'zh_CN',
        token: config.token || '' // 从请求配置中获取 token，默认为空字符串 
      },
      url,
      method,
      data,
      responseType: 'arraybuffer'
    })

    //
    // 通知Vue下载完成
    //
    event.sender.send(
      'download-status',
      {
        type: taskType,
        stage: 'processing',
        savePath
      }
    )

    //
    // 后台处理
    //
    setTimeout(async () => {
      try {
        let processedBuffer = response.data
        if (taskType === 'total') {
          processedBuffer = 
            await processSrmData(
              response.data
            )
        } else if (taskType === 'orderFrameInfo') {
          processedBuffer = await processOrderFrameData(
            response.data
          )
        } else if (taskType === 'fbk45Section') {
          processedBuffer = await processPurchasePlan(
            response.data
          )
        }

        fs.writeFileSync(
          savePath,
          Buffer.from(
            processedBuffer
          )
        )

        //
        // 通知Vue处理完成
        //
        event.sender.send(
          'download-status',
          {
            type: taskType,
            stage: 'success',
          }
        )

      } catch (err) {
        event.sender.send(
          'download-status',
          {
            type: taskType,
            stage: 'error',
            message: err.message
          }
        )
      }

    }, 0)

    return {
      success: true
    }
  } catch (error) {
    console.error('下载文件失败', error)
    return {
      success: false,
      message: error.message
    }
  }
})


app.whenReady().then(() => {
  createWindow()
})
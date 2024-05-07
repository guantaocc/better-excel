import { app, BrowserWindow, shell, ipcMain, dialog, Menu } from 'electron'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import fs from 'node:fs/promises'
import path from 'node:path'
import os from 'node:os'

const require = createRequire(import.meta.url)
const __dirname = path.dirname(fileURLToPath(import.meta.url))

import { expandTemplate } from './expand-template'

const sepTemplate = expandTemplate({ sep: '{{}}' })

const Excel = require('exceljs')



// The built directory structure
//
// ├─┬ dist-electron
// │ ├─┬ main
// │ │ └── index.js    > Electron-Main
// │ └─┬ preload
// │   └── index.mjs   > Preload-Scripts
// ├─┬ dist
// │ └── index.html    > Electron-Renderer
//
process.env.APP_ROOT = path.join(__dirname, '../..')

export const MAIN_DIST = path.join(process.env.APP_ROOT, 'dist-electron')
export const RENDERER_DIST = path.join(process.env.APP_ROOT, 'dist')
export const VITE_DEV_SERVER_URL = process.env.VITE_DEV_SERVER_URL

process.env.VITE_PUBLIC = VITE_DEV_SERVER_URL
  ? path.join(process.env.APP_ROOT, 'public')
  : RENDERER_DIST

// Disable GPU Acceleration for Windows 7
if (os.release().startsWith('6.1')) app.disableHardwareAcceleration()

// Set application name for Windows 10+ notifications
if (process.platform === 'win32') app.setAppUserModelId(app.getName())

if (!app.requestSingleInstanceLock()) {
  app.quit()
  process.exit(0)
}

let win: BrowserWindow | null = null
const preload = path.join(__dirname, '../preload/index.mjs')
const indexHtml = path.join(RENDERER_DIST, 'index.html')

async function createWindow() {
  win = new BrowserWindow({
    title: 'Main window',
    icon: path.join(process.env.VITE_PUBLIC, 'favicon.ico'),
    webPreferences: {
      preload,
      // Warning: Enable nodeIntegration and disable contextIsolation is not secure in production
      // nodeIntegration: true,

      // Consider using contextBridge.exposeInMainWorld
      // Read more on https://www.electronjs.org/docs/latest/tutorial/context-isolation
      // contextIsolation: true,
    },
  })

  if (VITE_DEV_SERVER_URL) { // #298
    win.loadURL(VITE_DEV_SERVER_URL)
    // Open devTool if the app is not packaged
    win.webContents.openDevTools()
  } else {
    // 隐藏菜单
    Menu.setApplicationMenu(null)
    win.loadFile(indexHtml)
  }

  // Test actively push message to the Electron-Renderer
  win.webContents.on('did-finish-load', () => {
    win?.webContents.send('main-process-message', new Date().toLocaleString())
  })

  // Make all links open with the browser, not with the application
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('https:')) shell.openExternal(url)
    return { action: 'deny' }
  })
 
  // win.webContents.on('will-navigate', (event, url) => { }) #344
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  win = null
  if (process.platform !== 'darwin') app.quit()
})

app.on('second-instance', () => {
  if (win) {
    // Focus on the main window if the user tried to open another
    if (win.isMinimized()) win.restore()
    win.focus()
  }
})

app.on('activate', () => {
  const allWindows = BrowserWindow.getAllWindows()
  if (allWindows.length) {
    allWindows[0].focus()
  } else {
    createWindow()
  }
})

// New window example arg: new windows url
ipcMain.handle('open-win', (_, arg) => {
  const childWindow = new BrowserWindow({
    width: 1200,
    height: 900,
    webPreferences: {
      preload,
      nodeIntegration: true,
      contextIsolation: false,
    },
  })

  if (VITE_DEV_SERVER_URL) {
    childWindow.loadURL(`${VITE_DEV_SERVER_URL}#${arg}`)
  } else {
    childWindow.loadFile(indexHtml, { hash: arg })
  }
})

const openFileDialog = async () => {
  const res = await dialog.showOpenDialog({
    title: '请选择文件',
    // 默认打开的路径，比如这里默认打开下载文件夹
    defaultPath: app.getPath('desktop'), 
    buttonLabel: '确定',
    // 限制能够选择的文件类型
    filters: [
      { name: 'excel', extensions: ['xls', 'xlsx', 'csv'] },
    ],
    properties: [ 'openFile','showHiddenFiles' ],
  })
  console.log('res', res)
  return res.canceled ? '' : res.filePaths[0]
}

const usedTemplateExcel = async (filename: string, template: string, sheetIndex: number) => {
  try {
    let errorMessage = ''
    let result = ''
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(filename);
    const worksheet = workbook.worksheets[sheetIndex - 1]
    if(!worksheet){
      errorMessage = '工作表不存在'
      return {
        errorMessage,
        result
      }
    }
    let headers = []
    worksheet.eachRow((row, index) => {
      if(index === 1){
      console.log(JSON.stringify(row.values), index);
        row.eachCell((cell, cellIndex) => {
          headers[cellIndex] = cell.value
        })
      } else {
        // 填充template
        let obj = {}
        row.eachCell({ includeEmpty: true },(cell, cellIndex) => {
          obj[headers[cellIndex]] = cell.value
        })
        console.log(obj);
        let rowsTemplate = sepTemplate(template, obj)
        result += `${rowsTemplate}\n\n`
      }
    })
    // 获取表头属性
    return {
      errorMessage,
      result
    }
  } catch (error) {
    return {
      errorMessage: '解析异常',
      result: ''
    }
  }
}

/**
 * 打开文件选择框
 * @param oldPath - 上一次打开的路径
 */
const downloadFileDialog = async (template: string) => {
  let oldPath =  app.getPath('downloads')
  if (!win) return oldPath

  const { canceled, filePaths } = await dialog.showOpenDialog(win, {
    title: '选择保存位置',
    properties: ['openDirectory', 'createDirectory'],
    defaultPath: oldPath,
  })

  const file_path = filePaths[0]

  const downloadPath = path.join(file_path, `${Date.now()}-指令.txt`)

  if(!canceled){
    await fs.writeFile(downloadPath, template, 'utf-8')
    oldPath = file_path
    return {
      success: true,
      filePath: downloadPath
    }
  } else {
    return {
      success: false
    }
  }
}

ipcMain.handle('openFileDialog', (event) => openFileDialog())
ipcMain.handle('genTplFromExcel', (event, filename: string, template: string, sheetIndex: number) => usedTemplateExcel(filename, template, sheetIndex))
ipcMain.handle('downloadFileDialog', (event, template: string) => downloadFileDialog(template))
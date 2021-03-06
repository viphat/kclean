import { app, BrowserWindow, Menu, dialog } from 'electron'
import { mainMenuTemplate } from './mainMenu'
import { clearCustomerData } from './clearCustomerData'
import { importData } from './importData'
import { generateReport } from './generateReport'
import _ from 'lodash'

// Keep a global reference of the window object, if you don't, the window will be closed automatically when the JavaScript object is garbage collected.
let mainWindow

function createWindow () {
  // Create the browser window.
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true
    }
  })

  // and load the index.html of the app.
  mainWindow.loadURL(`file://${__dirname}/index.html`)

  // Emitted when the window is closed.
  mainWindow.on('closed', () => {
    // Dereference the window object, usually you would store windows in an array if your app supports multi windows, this is the time when you should delete the corresponding element.
    mainWindow = null
  })
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on('ready', () => {
  createWindow()
  setApplicationMenu()
})

// Quit when all windows are closed.
app.on('window-all-closed', () => {
  app.quit()
})

app.on('activate', () => {
  if (mainWindow === null) {
    createWindow()
  }
})

// In this file you can include the rest of your app's specific main process code. You can also put them in separate files and require them here.
const setApplicationMenu = () => {
  const menus = [mainMenuTemplate(mainWindow)];
  Menu.setApplicationMenu(Menu.buildFromTemplate(menus));
};

var ipc = require('electron').ipcMain;

ipc.on('clearCustomerData', (event, batch) => {
  clearCustomerData(batch).then((response) => {
    event.sender.send('clearCustomerDataSuccessful', { success: true })
  }, (errRes) => {
    event.sender.send('clearCustomerDataFailed', { success: true })
  })
})

ipc.on('importData', (event, data) => {
  let { inputFile, outputDirectory, batch } = data;
  importData(inputFile, batch, outputDirectory).then((res) => {
    event.sender.send('importDataSuccessful', { success: true })
  })
})

ipc.on('generateReport', (event, data) => {
  let { outputDirectory, batch } = data;
  generateReport(batch, outputDirectory).then((reportFilePath) => {
    event.sender.send('generateReportSuccessful', reportFilePath)
  })
})


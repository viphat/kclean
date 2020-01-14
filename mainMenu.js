import { app, BrowserWindow, dialog } from 'electron'
import { setupDatabase, importData } from './database'

export const mainMenuTemplate = (mainWindow) => {
  return {
    label: 'Main Menu',
    submenu: [
      {
        label: 'Setup Database',
        click: () => {
          var message = 'Thao tác này sẽ khởi tạo và xóa hoàn toàn dữ liệu từ trước đến nay. Bạn có chắc không?';
          dialog.showMessageBox({
            message: message,
            buttons: ['OK', 'Cancel']
          }).then((obj) => {
            if (obj.response === 0) {
              setupDatabase()
            }
          });
        }
      },
      {
        label: 'Import Data',
        click: () => {
          var message = 'Thao tác này sẽ thêm một số dữ liệu ban đầu (Khu vực, thành phố) vào Database. Bạn có chắc không?';
          dialog.showMessageBox({
            message: message,
            buttons: ['OK', 'Cancel']
          }).then((obj) => {
            if (obj.response === 0) {
              importData()
            }
          });
        }
      },
      {
        label: 'Toggle DevTools',
        accelerator: 'Alt+CmdOrCtrl+I',
        click: () => {
          BrowserWindow.getFocusedWindow().toggleDevTools();
        },
      },
      {
        label: 'Quit',
        accelerator: 'CmdOrCtrl+Q',
        click: () => {
          app.quit();
        },
      }
    ]
  }
}

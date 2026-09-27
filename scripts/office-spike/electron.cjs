const { app, BrowserWindow, session } = require('electron')
const path = require('node:path')
app.setPath('userData', path.join(__dirname, 'electron-data'))
global.networkAttempts = []
app.whenReady().then(() => {
  session.defaultSession.webRequest.onBeforeRequest(
    { urls: ['http://*/*', 'https://*/*'] },
    (details, cb) => {
      global.networkAttempts.push(details.url)
      cb({ cancel: true })
    },
  )
  const window = new BrowserWindow({
    width: 1300,
    height: 900,
    show: false,
    webPreferences: { contextIsolation: true, sandbox: true, nodeIntegration: false },
  })
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  window.loadFile(path.join(__dirname, 'dist', 'index.html'))
})
app.on('window-all-closed', () => app.quit())

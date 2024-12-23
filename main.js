const { app, BrowserWindow } = require('electron');
const path = require('path');
const expressApp = require('./index.js'); 
let mainWindow;
const PORT = 8000;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1024,
        height: 768,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        }
    });

    mainWindow.loadURL(`http://localhost:${PORT}`);
    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

// Start Express server
app.on("ready", () => {
    // Start Express server
    expressApp.listen(PORT, () => {
      console.log(`Express server listening on port ${PORT}`);
    });
  
    createWindow();
  });
  
  app.on("window-all-closed", () => {
    if (process.platform !== "darwin") app.quit();
  });
  
  app.on("activate", () => {
    if (mainWindow === null) createWindow();
  });

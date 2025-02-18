const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  submitButton: (blNo, filePath) => ipcRenderer.invoke('submit-button', blNo, filePath),
});

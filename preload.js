const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  sendVGM: (blNo, filePath) => ipcRenderer.invoke('send-vgm', blNo, filePath),
  submitSI: (blNo, filePath) => ipcRenderer.invoke('submit-si', blNo, filePath),
});

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  submitBlNo: (blNo) => ipcRenderer.invoke('submit-bl-no', blNo),
  uploadXlsx: (filePath) => ipcRenderer.invoke('upload-xlsx', filePath)
});

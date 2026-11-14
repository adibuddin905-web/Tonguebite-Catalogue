const { contextBridge, ipcRenderer: r } = require('electron');
contextBridge.exposeInMainWorld('api', Object.fromEntries(
  ['fonts', 'load', 'save', 'saveImgs', 'loadImgs', 'img', 'sheet', 'saveAs', 'open', 'sample', 'saveFile', 'pickFolder', 'readImage', 'exportSheet', 'pdf', 'update'].map(n => [n, (...a) => r.invoke(n, ...a)])
));

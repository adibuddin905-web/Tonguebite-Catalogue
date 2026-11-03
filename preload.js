const { contextBridge, ipcRenderer: r } = require('electron');
contextBridge.exposeInMainWorld('api', Object.fromEntries(
  ['fonts', 'load', 'save', 'img', 'sheet', 'saveAs', 'open', 'pdf', 'update'].map(n => [n, (...a) => r.invoke(n, ...a)])
));

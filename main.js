const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const { autoUpdater } = require('electron-updater');
const fs = require('fs'), path = require('path');
let win;
const store = () => path.join(app.getPath('userData'), 'project.json');

app.whenReady().then(() => {
  win = new BrowserWindow({ width: 1440, height: 900, webPreferences: { preload: path.join(__dirname, 'preload.js') } });
  win.setMenuBarVisibility(false);
  win.loadFile('index.html');
});
app.on('window-all-closed', () => app.quit());

ipcMain.handle('fonts', async () => { try { return await require('font-list').getFonts({ disableQuoting: true }); } catch { return ['Arial']; } });
ipcMain.handle('load', () => fs.existsSync(store()) ? fs.readFileSync(store(), 'utf8') : null);
ipcMain.handle('save', (e, j) => fs.writeFileSync(store(), j));
ipcMain.handle('img', async () => {
  const r = await dialog.showOpenDialog(win, { filters: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'webp', 'svg'] }] });
  if (r.canceled) return null;
  const f = r.filePaths[0], x = path.extname(f).slice(1).toLowerCase();
  return `data:image/${x === 'jpg' ? 'jpeg' : x === 'svg' ? 'svg+xml' : x};base64,` + fs.readFileSync(f).toString('base64');
});
ipcMain.handle('sheet', async () => {
  const r = await dialog.showOpenDialog(win, { filters: [{ name: 'Excel / CSV', extensions: ['xlsx', 'xls', 'csv'] }] });
  if (r.canceled) return null;
  const X = require('xlsx'), wb = X.readFile(r.filePaths[0]);
  return X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
});
ipcMain.handle('saveAs', async (e, j) => {
  const r = await dialog.showSaveDialog(win, { defaultPath: 'catalogue.json' });
  if (!r.canceled) fs.writeFileSync(r.filePath, j);
});
ipcMain.handle('open', async () => {
  const r = await dialog.showOpenDialog(win, { filters: [{ name: 'Catalogue project', extensions: ['json'] }] });
  return r.canceled ? null : fs.readFileSync(r.filePaths[0], 'utf8');
});
ipcMain.handle('pdf', async () => {
  const r = await dialog.showSaveDialog(win, { defaultPath: 'catalogue.pdf' });
  if (r.canceled) return false;
  fs.writeFileSync(r.filePath, await win.webContents.printToPDF({ printBackground: true, preferCSSPageSize: true }));
  return r.filePath;
});

autoUpdater.autoDownload = false;
ipcMain.handle('update', async () => {
  if (!app.isPackaged) return dialog.showMessageBox(win, { message: 'Updates work only in the installed app.' });
  try {
    const r = await autoUpdater.checkForUpdates();
    if (!r || !r.updateInfo || r.updateInfo.version === app.getVersion())
      return dialog.showMessageBox(win, { message: 'You are on the latest version (' + app.getVersion() + ').' });
    const a = await dialog.showMessageBox(win, { message: 'Version ' + r.updateInfo.version + ' is available. Download and install now?', buttons: ['Update', 'Later'] });
    if (a.response) return;
    await autoUpdater.downloadUpdate();
    autoUpdater.quitAndInstall();
  } catch (e) { dialog.showErrorBox('Update failed', String(e)); }
});

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const { autoUpdater } = require('electron-updater');
const fs = require('fs'), path = require('path');
let win;
const store = () => path.join(app.getPath('userData'), 'project.json');

app.whenReady().then(() => {
  win = new BrowserWindow({ width: 1440, height: 900, webPreferences: { preload: path.join(__dirname, 'preload.js'), spellcheck: false } });
  win.setMenuBarVisibility(false);
  win.loadFile('index.html');
});
app.on('window-all-closed', () => app.quit());

ipcMain.handle('fonts', async () => { try { return await require('font-list').getFonts({ disableQuoting: true }); } catch { return ['Arial']; } });
ipcMain.handle('load', () => fs.existsSync(store()) ? fs.readFileSync(store(), 'utf8') : null);
ipcMain.handle('save', (e, j) => fs.writeFileSync(store(), j));
const imgFile = () => path.join(app.getPath('userData'), 'images.json');
ipcMain.handle('saveImgs', (e, j) => fs.writeFileSync(imgFile(), j));
ipcMain.handle('loadImgs', () => fs.existsSync(imgFile()) ? fs.readFileSync(imgFile(), 'utf8') : null);
ipcMain.handle('img', async () => {
  const r = await dialog.showOpenDialog(win, { filters: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'webp', 'svg'] }] });
  if (r.canceled) return null;
  const f = r.filePaths[0], x = path.extname(f).slice(1).toLowerCase();
  return `data:image/${x === 'jpg' ? 'jpeg' : x === 'svg' ? 'svg+xml' : x};base64,` + fs.readFileSync(f).toString('base64');
});
const IMG_EXT = ['png', 'jpg', 'jpeg', 'webp', 'svg'];
const nzName = t => String(t).toLowerCase().replace(/[^a-z0-9]+/g, '');
const isImgCol = k => /^(image|img|photo|picture)\s*\d*$/i.test(String(k).trim());
const toDataUrl = f => { const x = path.extname(f).slice(1).toLowerCase(); return `data:image/${x === 'jpg' ? 'jpeg' : x === 'svg' ? 'svg+xml' : x};base64,` + fs.readFileSync(f).toString('base64'); };
function cleanTarget(t) {
  t = String(t).trim().replace(/^file:\/\//i, '');
  try { t = decodeURIComponent(t); } catch (e) {}
  t = t.replace(/^"|"$/g, '');
  return /^\/[A-Za-z]:/.test(t) ? t.slice(1) : t;
}
// Reads a worksheet cell by cell so hyperlinks (Ctrl+K picture links) and blank rows stay aligned.
function readRows(X, ws) {
  if (!ws || !ws['!ref']) return [];
  const R = X.utils.decode_range(ws['!ref']), heads = [];
  for (let c = R.s.c; c <= R.e.c; c++) {
    const h = ws[X.utils.encode_cell({ r: R.s.r, c })];
    heads[c] = h && h.v != null ? String(h.v).trim() : '';
  }
  const rows = [];
  for (let r = R.s.r + 1; r <= R.e.r; r++) {
    const o = {}; let any = false;
    for (let c = R.s.c; c <= R.e.c; c++) {
      const k = heads[c]; if (!k) continue;
      const cell = ws[X.utils.encode_cell({ r, c })];
      let v = cell && cell.v != null ? String(cell.v) : '';
      if (cell && cell.l && cell.l.Target && isImgCol(k) && !/^https?:/i.test(String(cell.l.Target))) v = cleanTarget(cell.l.Target);
      o[k] = v; if (v.trim()) any = true;
    }
    if (any) rows.push(o);
  }
  return rows;
}
function attachPictures(rows, dir) {
  let files = null;
  const lazy = rows.length > 25;   // big sheets: keep picture paths only, load pictures when needed
  return rows.map(o => {
    const imgs = [], paths = []; let auto = false;
    const add = f => { paths.push(f); if (!lazy) imgs.push(toDataUrl(f)); };
    Object.keys(o).forEach(k => {
      if (!isImgCol(k) || !String(o[k]).trim()) return;
      let f = cleanTarget(o[k]);
      if (!path.isAbsolute(f)) f = path.join(dir, f);
      const x = path.extname(f).slice(1).toLowerCase();
      try { if (IMG_EXT.includes(x) && fs.existsSync(f)) add(f); } catch (e) {}
    });
    const nk = Object.keys(o).find(k => ['name', 'product', 'product name'].includes(k.trim().toLowerCase()));
    const name = nk ? String(o[nk]).trim() : '';
    if (!paths.length && name) {
      try {
        files = files || fs.readdirSync(dir).filter(f => IMG_EXT.includes(path.extname(f).slice(1).toLowerCase()));
        const key = nzName(name);
        files.filter(f => { const b = nzName(path.basename(f, path.extname(f))); return b === key || (b.startsWith(key) && /^[0-9]+$/.test(b.slice(key.length))); })
          .sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).slice(0, 6)
          .forEach(f => { add(path.join(dir, f)); auto = true; });
      } catch (e) {}
    }
    return { ...o, __imgs: imgs, __files: paths, __auto: auto };
  });
}
ipcMain.handle('sheet', async () => {
  const r = await dialog.showOpenDialog(win, { filters: [{ name: 'Excel / CSV', extensions: ['xlsx', 'xlsm', 'xls', 'csv'] }] });
  if (r.canceled) return null;
  const f0 = r.filePaths[0], dir = path.dirname(f0), X = require('xlsx');
  const wb = /\.csv$/i.test(f0) ? X.read(fs.readFileSync(f0, 'utf8'), { type: 'string' }) : X.readFile(f0);
  const sn = wb.SheetNames.find(n => /^products$/i.test(n)) || wb.SheetNames[0];
  return attachPictures(readRows(X, wb.Sheets[sn]), dir);
});
ipcMain.handle('sample', async () => {
  const src = path.join(__dirname, 'products-template.xlsx'), has = fs.existsSync(src);
  const r = await dialog.showSaveDialog(win, { defaultPath: has ? 'products-template.xlsx' : 'products-sample.csv' });
  if (r.canceled) return;
  if (has) { fs.writeFileSync(r.filePath, fs.readFileSync(src)); return; }
  fs.writeFileSync(r.filePath, '\ufeffname,category,description,variants,image,image2\n' +
    'Turmeric Powder,Powder,Pure ground turmeric,100g - Rs 60 | 250g - Rs 140,turmeric.jpg,turmeric-pack.jpg\n' +
    'Cumin Seeds,Seeds,Whole cumin seeds,100g - Rs 80 | 500g - Rs 360,cumin.jpg,\n');
});
const FOLDER_EXT = ['jpg', 'jpeg', 'png', 'webp', 'gif', 'bmp'];
// Choose a folder; returns every supported picture in it (paths only, nothing is loaded into memory)
ipcMain.handle('pickFolder', async () => {
  const r = await dialog.showOpenDialog(win, { properties: ['openDirectory'] });
  if (r.canceled || !r.filePaths[0]) return null;
  const dir = r.filePaths[0];
  let names = [];
  try { names = fs.readdirSync(dir); } catch (e) { return { dir, files: [], error: String(e.message || e) }; }
  const files = names
    .filter(n => !n.startsWith('.') && !n.startsWith('~$') && FOLDER_EXT.includes(path.extname(n).slice(1).toLowerCase()))
    .map(n => path.join(dir, n))
    .filter(p => { try { return fs.statSync(p).isFile(); } catch (e) { return false; } })
    .sort((a, b) => path.basename(a).localeCompare(path.basename(b), undefined, { numeric: true, sensitivity: 'base' }))
    .map(p => ({ path: p, name: path.basename(p), size: fs.statSync(p).size }));
  return { dir, files };
});
// Reads ONE picture file when the catalogue needs it (original quality)
ipcMain.handle('readImage', async (e, p) => {
  try {
    const x = path.extname(String(p)).slice(1).toLowerCase();
    if (!FOLDER_EXT.concat(['svg']).includes(x) || !fs.existsSync(p)) return null;
    return toDataUrl(p);
  } catch (err) { return null; }
});
// Writes the product list to a new Excel (.xlsx) or CSV file with the same columns the importer reads
ipcMain.handle('exportSheet', async (e, rows) => {
  const r = await dialog.showSaveDialog(win, { defaultPath: 'products.xlsx', filters: [{ name: 'Excel', extensions: ['xlsx'] }, { name: 'CSV (UTF-8)', extensions: ['csv'] }] });
  if (r.canceled) return null;
  const X = require('xlsx'), n = Math.max(4, ...rows.map(o => (o.pics || []).length));
  const head = ['name', 'category', 'description', 'variants', 'image'].concat(Array.from({ length: n - 1 }, (_, i) => 'image' + (i + 2)));
  const aoa = [head].concat(rows.map(o => [o.name || '', o.cat || '', o.desc || '', String(o.variants || '').split(/\r?\n/).filter(Boolean).join(' | ')]
    .concat(Array.from({ length: n }, (_, i) => (o.pics || [])[i] || ''))));
  const ws = X.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 26 }, { wch: 16 }, { wch: 38 }, { wch: 34 }].concat(Array.from({ length: n }, () => ({ wch: 40 })));
  if (/\.csv$/i.test(r.filePath)) { fs.writeFileSync(r.filePath, '\ufeff' + X.utils.sheet_to_csv(ws)); return r.filePath; }
  const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, 'Products'); X.writeFile(wb, r.filePath);
  return r.filePath;
});
ipcMain.handle('saveFile', async (e, name, data) => {
  const r = await dialog.showSaveDialog(win, { defaultPath: name });
  if (r.canceled) return null;
  fs.writeFileSync(r.filePath, typeof data === 'string' ? data : Buffer.from(data));
  return r.filePath;
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

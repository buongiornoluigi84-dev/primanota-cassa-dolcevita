/**
 * PrimaNota Cassa Dolce Vita — Backend (backend.gs)
 * --------------------------------------------------
 * VERSIONE AGGIORNATA — aggiunge il registro degli import CSV.
 *
 * COME AGGIORNARE:
 * 1. Apri il progetto Apps Script, seleziona il file "backend.gs"
 * 2. Seleziona TUTTO il contenuto (Ctrl+A) e incolla questo al suo posto
 * 3. Salva (dischetto)
 * 4. Deploy > Gestisci deployment > matita ✏️ > Versione: "Nuova versione"
 *    > Esegui deployment
 *
 * Il tab "Imports" viene creato da solo al primo caricamento di un CSV:
 * NON serve rieseguire setupPrimaNota (che cancellerebbe i dati esistenti).
 */

// ============ CONFIG ============
const TIMEZONE = 'Europe/London';
const PHOTOS_ROOT_FOLDER_NAME = 'PrimaNota Cassa Dolce Vita - Ricevute';
const RECORD_HEADERS = ['ID', 'Data', 'Fornitore', 'Importo', 'Descrizione', 'Categoria', 'Metodo Pagamento', 'Foto', 'Riconciliato'];

// ============ ENTRY POINTS ============

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'records') {
      return jsonOutput({ ok: true, data: readRecords() });
    }
    if (action === 'settings') {
      return jsonOutput({ ok: true, data: readSettings() });
    }
    if (action === 'imports') {
      return jsonOutput({ ok: true, data: readImports() });
    }
    if (action === 'bankrows') {
      return jsonOutput({ ok: true, data: readBankRows(e.parameter.importId) });
    }
    if (action === 'photolist') {
      return jsonOutput({ ok: true, data: readPhotoList(e.parameter.id) });
    }
    if (action === 'photo') {
      return jsonOutput({ ok: true, data: readPhoto(e.parameter.fileId) });
    }
    return jsonOutput({ ok: false, error: 'Azione GET non riconosciuta: ' + action });
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'addRecord') {
      return jsonOutput({ ok: true, data: addRecord(body) });
    }
    if (action === 'updateRecord') {
      return jsonOutput({ ok: true, data: updateRecord(body.id, body) });
    }
    if (action === 'deleteRecord') {
      deleteRecord(body.id);
      return jsonOutput({ ok: true });
    }
    if (action === 'addSetting') {
      addSettingValue(body.type, body.value);
      return jsonOutput({ ok: true });
    }
    if (action === 'removeSetting') {
      removeSettingValue(body.type, body.value);
      return jsonOutput({ ok: true });
    }
    if (action === 'renameSetting') {
      const count = renameSettingValue(body.type, body.oldValue, body.newValue);
      return jsonOutput({ ok: true, recordsUpdated: count });
    }
    if (action === 'addImport') {
      return jsonOutput({ ok: true, data: addImport(body) });
    }
    if (action === 'removeImport') {
      removeImport(body.id);
      return jsonOutput({ ok: true });
    }
    if (action === 'addBankRows') {
      addBankRows(body.importId, body.rows);
      return jsonOutput({ ok: true });
    }
    if (action === 'updateBankRow') {
      updateBankRow(body.id, body.stato, body.recordId);
      return jsonOutput({ ok: true });
    }

    return jsonOutput({ ok: false, error: 'Azione POST non riconosciuta: ' + action });
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message });
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============ SHEET HELPERS ============

function getRecordsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Records');
}

function getSettingsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
}

function readRecords() {
  const sheet = getRecordsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, RECORD_HEADERS.length).getValues();
  return values
    .filter(row => row[0] !== '') // salta righe vuote
    .map(row => {
      const obj = {};
      RECORD_HEADERS.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function readSettings() {
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { fornitori: [], categorie: [], metodiPagamento: [] };
  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return {
    fornitori: values.map(r => r[0]).filter(v => v !== ''),
    categorie: values.map(r => r[1]).filter(v => v !== ''),
    metodiPagamento: values.map(r => r[2]).filter(v => v !== '')
  };
}

// ============ RECORDS ============

function addRecord(body) {
  const sheet = getRecordsSheet();
  const now = new Date();
  const id = Utilities.formatDate(now, TIMEZONE, 'yyyyMMdd-HHmmss');

  let fotoLink = '';
  if (body.photos && body.photos.length > 0) {
    fotoLink = savePhotosForRecord(id, body.data, body.fornitore, body.photos);
  }

  const row = [
    id,
    body.data,
    body.fornitore,
    body.importo,
    body.descrizione || '',
    body.categoria,
    body.metodoPagamento,
    fotoLink,
    body.riconciliato || 'No'
  ];
  sheet.appendRow(row);

  const obj = {};
  RECORD_HEADERS.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

function updateRecord(id, body) {
  const sheet = getRecordsSheet();
  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let rowIndex = -1;
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) { rowIndex = i + 2; break; }
  }
  if (rowIndex === -1) throw new Error('Record non trovato: ' + id);

  // Campi aggiornabili
  const fieldToCol = {
    data: 2, fornitore: 3, importo: 4, descrizione: 5,
    categoria: 6, metodoPagamento: 7, riconciliato: 9
  };
  Object.keys(fieldToCol).forEach(key => {
    if (body[key] !== undefined) {
      sheet.getRange(rowIndex, fieldToCol[key]).setValue(body[key]);
    }
  });

  // Foto aggiuntive (si aggiungono alla cartella esistente, o se ne crea una nuova)
  if (body.photos && body.photos.length > 0) {
    const currentRow = sheet.getRange(rowIndex, 1, 1, RECORD_HEADERS.length).getValues()[0];
    const fornitore = body.fornitore || currentRow[2];
    const data = body.data || currentRow[1];
    const fotoLink = savePhotosForRecord(id, data, fornitore, body.photos, currentRow[7]);
    sheet.getRange(rowIndex, 8).setValue(fotoLink);
  }

  const updatedRow = sheet.getRange(rowIndex, 1, 1, RECORD_HEADERS.length).getValues()[0];
  const obj = {};
  RECORD_HEADERS.forEach((h, i) => obj[h] = updatedRow[i]);
  return obj;
}

function deleteRecord(id) {
  const sheet = getRecordsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('Nessun record da eliminare');
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      unlinkBankRowsForRecord(id);
      return;
    }
  }
  throw new Error('Record non trovato: ' + id);
}

/**
 * Se il record eliminato era collegato a una riga dell'estratto conto,
 * quella riga torna "open" così ricompare tra i movimenti da controllare.
 */
function unlinkBankRowsForRecord(recordId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BankRows');
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const values = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(recordId)) {
      sheet.getRange(i + 2, 8).setValue('open');
      sheet.getRange(i + 2, 9).setValue('');
    }
  }
}

// ============ FOTO / DRIVE ============

function getOrCreateFolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function sanitizeForFolderName(text) {
  return String(text).replace(/[\\/:*?"<>|]/g, '-').trim();
}

/**
 * Salva le foto (base64) in una sottocartella dedicata al record,
 * organizzata per anno/mese. Se existingFolderUrl è passato, riusa
 * quella cartella invece di crearne una nuova (caso: si aggiungono
 * foto a un record già esistente).
 */
function savePhotosForRecord(id, dataStr, fornitore, photosBase64, existingFolderUrl) {
  const rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), PHOTOS_ROOT_FOLDER_NAME);

  let recordFolder;
  if (existingFolderUrl) {
    const folderId = existingFolderUrl.match(/[-\w]{25,}/)[0];
    recordFolder = DriveApp.getFolderById(folderId);
  } else {
    const date = new Date(dataStr);
    const year = Utilities.formatDate(date, TIMEZONE, 'yyyy');
    const month = Utilities.formatDate(date, TIMEZONE, 'MM');
    const yearFolder = getOrCreateFolder(rootFolder, year);
    const monthFolder = getOrCreateFolder(yearFolder, month);
    const folderName = id + '_' + sanitizeForFolderName(fornitore || 'Senza-fornitore');
    recordFolder = getOrCreateFolder(monthFolder, folderName);
  }

  const existingCount = recordFolder.getFiles();
  let n = 0;
  while (existingCount.hasNext()) { existingCount.next(); n++; }

  photosBase64.forEach((base64, i) => {
    const match = base64.match(/^data:(image\/\w+);base64,(.+)$/);
    const mimeType = match ? match[1] : 'image/jpeg';
    const data = match ? match[2] : base64;
    const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, 'pagina' + (n + i + 1) + '.jpg');
    recordFolder.createFile(blob);
  });

  return recordFolder.getUrl();
}

// ============ IMPORTS (registro estratti conto) ============
const IMPORT_HEADERS = ['ID', 'Caricato il', 'Periodo Da', 'Periodo A', 'Movimenti', 'File'];

function getImportsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Imports');
  if (!sheet) {
    sheet = ss.insertSheet('Imports');
    sheet.getRange(1, 1, 1, IMPORT_HEADERS.length).setValues([IMPORT_HEADERS]);
    sheet.getRange(1, 1, 1, IMPORT_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readImports() {
  const sheet = getImportsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, IMPORT_HEADERS.length).getValues();
  return values
    .filter(row => row[0] !== '')
    .map(row => {
      const obj = {};
      IMPORT_HEADERS.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function addImport(body) {
  const sheet = getImportsSheet();
  const now = new Date();
  const id = Utilities.formatDate(now, TIMEZONE, 'yyyyMMdd-HHmmss');
  const row = [
    id,
    Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm'),
    body.from || '',
    body.to || '',
    body.count || 0,
    body.filename || ''
  ];
  sheet.appendRow(row);
  const obj = {};
  IMPORT_HEADERS.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

function removeImport(id) {
  const sheet = getImportsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      break;
    }
  }
  deleteBankRowsForImport(id);
}

// ============ BANK ROWS (movimenti estratto conto salvati) ============
const BANKROW_HEADERS = ['ID', 'ImportID', 'Data', 'Data Contabile', 'Importo', 'Descrizione', 'Tipo', 'Stato', 'Record ID'];

function getBankRowsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BankRows');
  if (!sheet) {
    sheet = ss.insertSheet('BankRows');
    sheet.getRange(1, 1, 1, BANKROW_HEADERS.length).setValues([BANKROW_HEADERS]);
    sheet.getRange(1, 1, 1, BANKROW_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readBankRows(importId) {
  const sheet = getBankRowsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, BANKROW_HEADERS.length).getValues();
  return values
    .filter(row => row[0] !== '')
    .filter(row => !importId || String(row[1]) === String(importId))
    .map(row => {
      const obj = {};
      BANKROW_HEADERS.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function addBankRows(importId, rows) {
  if (!rows || rows.length === 0) return;
  const sheet = getBankRowsSheet();
  const values = rows.map((r, i) => ([
    importId + '-' + String(i).padStart(3, '0'),
    importId,
    r.data || '',
    r.dataContabile || '',
    r.importo || 0,
    r.descrizione || '',
    r.tipo || '',
    r.stato || 'open',
    r.recordId || ''
  ]));
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, BANKROW_HEADERS.length).setValues(values);
}

function updateBankRow(id, stato, recordId) {
  const sheet = getBankRowsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      if (stato !== undefined) sheet.getRange(i + 2, 8).setValue(stato);
      if (recordId !== undefined) sheet.getRange(i + 2, 9).setValue(recordId);
      return;
    }
  }
}

function deleteBankRowsForImport(importId) {
  const sheet = getBankRowsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]) === String(importId)) {
      sheet.deleteRow(i + 2);
    }
  }
}

/**
 * Elenca le foto collegate a un record, senza scaricarle.
 * Ritorna [{ id, name }] — le immagini vere si leggono poi una alla volta
 * con readPhoto(), così l'app resta veloce anche con più pagine.
 */
function readPhotoList(recordId) {
  const sheet = getRecordsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, RECORD_HEADERS.length).getValues();
  let folderUrl = '';
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(recordId)) { folderUrl = String(values[i][7] || ''); break; }
  }
  if (!folderUrl) return [];

  const match = folderUrl.match(/[-\w]{25,}/);
  if (!match) return [];
  const folder = DriveApp.getFolderById(match[0]);
  const files = folder.getFiles();
  const out = [];
  while (files.hasNext()) {
    const f = files.next();
    out.push({ id: f.getId(), name: f.getName() });
  }
  out.sort((a, b) => a.name.localeCompare(b.name));
  return out;
}

/** Restituisce una singola foto come immagine incorporabile nell'app. */
function readPhoto(fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  return {
    name: file.getName(),
    dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes())
  };
}

// ============ SETTINGS ============

function addSettingValue(type, value) {
  const col = settingsColumnForType(type);
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  const colValues = sheet.getRange(2, col, Math.max(lastRow - 1, 0), 1).getValues().flat();
  const nextEmptyRow = colValues.findIndex(v => v === '') ;
  const targetRow = nextEmptyRow === -1 ? lastRow + 1 : nextEmptyRow + 2;
  sheet.getRange(targetRow, col).setValue(value);
}

function removeSettingValue(type, value) {
  const col = settingsColumnForType(type);
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  const colValues = sheet.getRange(2, col, Math.max(lastRow - 1, 0), 1).getValues();
  for (let i = 0; i < colValues.length; i++) {
    if (colValues[i][0] === value) {
      sheet.getRange(i + 2, col).setValue('');
      return;
    }
  }
}

/**
 * Rinomina un valore in Settings E aggiorna tutti i record storici
 * che lo usano. Ritorna il numero di record aggiornati.
 */
function renameSettingValue(type, oldValue, newValue) {
  const col = settingsColumnForType(type);
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  const colValues = sheet.getRange(2, col, Math.max(lastRow - 1, 0), 1).getValues();
  for (let i = 0; i < colValues.length; i++) {
    if (colValues[i][0] === oldValue) {
      sheet.getRange(i + 2, col).setValue(newValue);
      break;
    }
  }

  // Aggiorna anche i record storici
  const recordCol = recordColumnForType(type);
  const recordsSheet = getRecordsSheet();
  const lastRecordRow = recordsSheet.getLastRow();
  if (lastRecordRow < 2) return 0;
  const recordValues = recordsSheet.getRange(2, recordCol, lastRecordRow - 1, 1).getValues();
  let count = 0;
  for (let i = 0; i < recordValues.length; i++) {
    if (recordValues[i][0] === oldValue) {
      recordsSheet.getRange(i + 2, recordCol).setValue(newValue);
      count++;
    }
  }
  return count;
}

function settingsColumnForType(type) {
  if (type === 'fornitore') return 1;
  if (type === 'categoria') return 2;
  if (type === 'metodoPagamento') return 3;
  throw new Error('Tipo settings non valido: ' + type);
}

function recordColumnForType(type) {
  if (type === 'fornitore') return 3;
  if (type === 'categoria') return 6;
  if (type === 'metodoPagamento') return 7;
  throw new Error('Tipo record non valido: ' + type);
}

// ============================================================
// VISTA PORTÁL — Google Apps Script Backend
// Zkopíruj celý tento kód do Google Apps Script (script.google.com)
// ============================================================

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const DRIVE_FOLDER_NAME = 'Vista Portal Dokumenty';

// ---- SHEET NAMES ----
const SHEETS = {
  EMPLOYEES: 'Zaměstnanci',
  DOCUMENTS: 'Dokumenty',
  CONFIRMATIONS: 'Potvrzení'
};

// ============================================================
// MAIN ENTRY POINT
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    let result;
    switch (action) {
      case 'getEmployees':    result = getEmployees(); break;
      case 'getDocuments':    result = getDocuments(); break;
      case 'getAll':          result = getAll(); break;
      case 'addEmployee':     result = addEmployee(data); break;
      case 'deleteEmployee':  result = deleteEmployee(data.id); break;
      case 'uploadDocument':  result = uploadDocument(data); break;
      case 'deleteDocument':  result = deleteDocument(data.id); break;
      case 'addConfirmation': result = addConfirmation(data); break;
      default: result = { error: 'Unknown action: ' + action };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doGet(e) {
  // Allow GET for simple testing
  const action = e.parameter.action || 'getAll';
  try {
    let result;
    switch (action) {
      case 'getAll':       result = getAll(); break;
      case 'getEmployees': result = getEmployees(); break;
      case 'getDocuments': result = getDocuments(); break;
      default: result = { error: 'Unknown action' };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// INIT — vytvoří sheety a složku pokud neexistují
// ============================================================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Employees sheet
  let empSheet = ss.getSheetByName(SHEETS.EMPLOYEES);
  if (!empSheet) {
    empSheet = ss.insertSheet(SHEETS.EMPLOYEES);
    empSheet.appendRow(['id', 'name', 'role', 'pin', 'createdAt']);
    empSheet.setFrozenRows(1);
    empSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
  }

  // Documents sheet
  let docSheet = ss.getSheetByName(SHEETS.DOCUMENTS);
  if (!docSheet) {
    docSheet = ss.insertSheet(SHEETS.DOCUMENTS);
    docSheet.appendRow(['id', 'title', 'category', 'desc', 'url', 'fileName', 'uploadedAt']);
    docSheet.setFrozenRows(1);
    docSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
  }

  // Confirmations sheet
  let confSheet = ss.getSheetByName(SHEETS.CONFIRMATIONS);
  if (!confSheet) {
    confSheet = ss.insertSheet(SHEETS.CONFIRMATIONS);
    confSheet.appendRow(['docId', 'docTitle', 'employeeId', 'employeeName', 'confirmedAt']);
    confSheet.setFrozenRows(1);
    confSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
  }

  return { success: true, message: 'Sheets initialized' };
}

// ============================================================
// EMPLOYEES
// ============================================================
function getEmployees() {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { employees: [] };

  const employees = rows.slice(1).map(r => ({
    id: r[0], name: r[1], role: r[2], pin: r[3], createdAt: r[4]
  })).filter(e => e.id);

  return { employees };
}

function addEmployee(data) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const id = data.id || 'emp_' + Date.now();
  sheet.appendRow([id, data.name, data.role || '', data.pin, new Date().toISOString()]);
  return { success: true, id };
}

function deleteEmployee(id) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

// ============================================================
// DOCUMENTS
// ============================================================
function getDocuments() {
  const sheet = getSheet(SHEETS.DOCUMENTS);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { documents: [] };

  const documents = rows.slice(1).map(r => ({
    id: r[0], title: r[1], category: r[2], desc: r[3], url: r[4], fileName: r[5], uploadedAt: r[6]
  })).filter(d => d.id);

  return { documents };
}

function uploadDocument(data) {
  // Save file to Google Drive
  let fileUrl = '';
  let docId = data.id || 'doc_' + Date.now();

  if (data.fileData && data.fileName) {
    const folder = getDriveFolder();
    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.fileData),
      data.fileType || 'application/octet-stream',
      data.fileName
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileUrl = file.getDownloadUrl();
  }

  const sheet = getSheet(SHEETS.DOCUMENTS);
  sheet.appendRow([
    docId,
    data.title,
    data.category || '',
    data.desc || '',
    fileUrl,
    data.fileName || '',
    data.uploadedAt || new Date().toISOString()
  ]);

  return { success: true, docId, url: fileUrl };
}

function deleteDocument(id) {
  const sheet = getSheet(SHEETS.DOCUMENTS);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

// ============================================================
// CONFIRMATIONS
// ============================================================
function addConfirmation(data) {
  const sheet = getSheet(SHEETS.CONFIRMATIONS);

  // Check duplicate
  const rows = sheet.getDataRange().getValues();
  const exists = rows.slice(1).some(r => r[0] === data.docId && r[2] === data.employeeId);
  if (exists) return { success: true, duplicate: true };

  // Get doc title
  const docSheet = getSheet(SHEETS.DOCUMENTS);
  const docRows = docSheet.getDataRange().getValues();
  const docRow = docRows.slice(1).find(r => r[0] === data.docId);
  const docTitle = docRow ? docRow[1] : data.docId;

  sheet.appendRow([
    data.docId,
    docTitle,
    data.employeeId,
    data.employeeName,
    data.confirmedAt || new Date().toISOString()
  ]);

  return { success: true };
}

function getConfirmations() {
  const sheet = getSheet(SHEETS.CONFIRMATIONS);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { confirmations: [] };

  const confirmations = rows.slice(1).map(r => ({
    docId: r[0], docTitle: r[1], employeeId: r[2], employeeName: r[3], confirmedAt: r[4]
  })).filter(c => c.docId);

  return { confirmations };
}

function getAll() {
  const emp = getEmployees();
  const doc = getDocuments();
  const conf = getConfirmations();
  return {
    employees: emp.employees,
    documents: doc.documents,
    confirmations: conf.confirmations
  };
}

// ============================================================
// HELPERS
// ============================================================
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    initSheets();
    sheet = ss.getSheetByName(name);
  }
  return sheet;
}

function getDriveFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

// ============================================================
// RUN THIS MANUALLY FIRST: Inicializace (spusť jednou ručně)
// ============================================================
function setup() {
  initSheets();
  Logger.log('Vista Portál setup dokončen.');
  Logger.log('Nyní nasaď web app: Nasadit > Nové nasazení > Web app');
}

// ============================================================
// VISTA PORTÁL — Google Apps Script Backend v2.1
// ============================================================

const DRIVE_FOLDER_NAME = 'Vista Portal Dokumenty';
const PORTAL_URL = 'https://davidsiwy.github.io/vista-portal/';

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
      case 'updateEmployee':  result = updateEmployee(data); break;
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
// INIT
// ============================================================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let empSheet = ss.getSheetByName(SHEETS.EMPLOYEES);
  if (!empSheet) {
    empSheet = ss.insertSheet(SHEETS.EMPLOYEES);
    empSheet.appendRow(['id', 'name', 'role', 'dept', 'pin', 'email', 'createdAt']);
    empSheet.setFrozenRows(1);
    empSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
  } else {
    // Upgrade: add email column if missing
    const headers = empSheet.getRange(1, 1, 1, empSheet.getLastColumn()).getValues()[0];
    if (!headers.includes('email')) {
      const emailCol = empSheet.getLastColumn() + 1;
      empSheet.getRange(1, emailCol).setValue('email').setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
    }
    if (!headers.includes('dept')) {
      const deptCol = headers.indexOf('pin') + 1; // insert before pin
      empSheet.insertColumnBefore(deptCol);
      empSheet.getRange(1, deptCol).setValue('dept').setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
    }
  }

  let docSheet = ss.getSheetByName(SHEETS.DOCUMENTS);
  if (!docSheet) {
    docSheet = ss.insertSheet(SHEETS.DOCUMENTS);
    docSheet.appendRow(['id', 'title', 'category', 'desc', 'url', 'fileName', 'uploadedAt']);
    docSheet.setFrozenRows(1);
    docSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
  }

  let confSheet = ss.getSheetByName(SHEETS.CONFIRMATIONS);
  if (!confSheet) {
    confSheet = ss.insertSheet(SHEETS.CONFIRMATIONS);
    confSheet.appendRow(['docId', 'docTitle', 'employeeId', 'employeeName', 'confirmedAt']);
    confSheet.setFrozenRows(1);
    confSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
  }

  return { success: true };
}

// ============================================================
// EMPLOYEES
// ============================================================
function getEmployees() {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { employees: [] };

  // Support both old (no dept/email) and new schema
  const headers = rows[0];
  const idx = {
    id:    headers.indexOf('id'),
    name:  headers.indexOf('name'),
    role:  headers.indexOf('role'),
    dept:  headers.indexOf('dept'),
    pin:   headers.indexOf('pin'),
    email: headers.indexOf('email')
  };

  const employees = rows.slice(1).map(r => ({
    id:    r[idx.id]    || '',
    name:  r[idx.name]  || '',
    role:  r[idx.role]  || '',
    dept:  idx.dept  >= 0 ? r[idx.dept]  : '',
    pin:   idx.pin   >= 0 ? r[idx.pin]   : '',
    email: idx.email >= 0 ? r[idx.email] : ''
  })).filter(e => e.id);

  return { employees };
}

function addEmployee(data) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const id = data.id || 'emp_' + Date.now();
  sheet.appendRow([id, data.name, data.role || '', data.dept || '', data.pin, data.email || '', new Date().toISOString()]);
  return { success: true, id };
}

function updateEmployee(data) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idx = {
    id:    headers.indexOf('id'),
    name:  headers.indexOf('name'),
    role:  headers.indexOf('role'),
    dept:  headers.indexOf('dept'),
    pin:   headers.indexOf('pin'),
    email: headers.indexOf('email')
  };
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idx.id] === data.id) {
      if (idx.name  >= 0) sheet.getRange(i+1, idx.name+1).setValue(data.name || '');
      if (idx.role  >= 0) sheet.getRange(i+1, idx.role+1).setValue(data.role || '');
      if (idx.dept  >= 0) sheet.getRange(i+1, idx.dept+1).setValue(data.dept || '');
      if (idx.pin   >= 0) sheet.getRange(i+1, idx.pin+1).setValue(data.pin || '');
      if (idx.email >= 0) sheet.getRange(i+1, idx.email+1).setValue(data.email || '');
      break;
    }
  }
  return { success: true };
}

function deleteEmployee(id) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === id) { sheet.deleteRow(i + 1); break; }
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
  let fileUrl = '';
  const docId = data.id || 'doc_' + Date.now();

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
  sheet.appendRow([docId, data.title, data.category || '', data.desc || '', fileUrl, data.fileName || '', data.uploadedAt || new Date().toISOString()]);

  // Send notification emails to all employees with email
  sendDocumentNotification(data.title, data.category || '', data.desc || '', fileUrl);

  return { success: true, docId, url: fileUrl };
}

function deleteDocument(id) {
  const sheet = getSheet(SHEETS.DOCUMENTS);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === id) { sheet.deleteRow(i + 1); break; }
  }
  return { success: true };
}

// ============================================================
// CONFIRMATIONS
// ============================================================
function addConfirmation(data) {
  const sheet = getSheet(SHEETS.CONFIRMATIONS);
  const rows = sheet.getDataRange().getValues();
  const exists = rows.slice(1).some(r => r[0] === data.docId && r[2] === data.employeeId);
  if (exists) return { success: true, duplicate: true };

  const docSheet = getSheet(SHEETS.DOCUMENTS);
  const docRows = docSheet.getDataRange().getValues();
  const docRow = docRows.slice(1).find(r => r[0] === data.docId);
  const docTitle = docRow ? docRow[1] : data.docId;

  sheet.appendRow([data.docId, docTitle, data.employeeId, data.employeeName, data.confirmedAt || new Date().toISOString()]);

  // Send confirmation receipt to admin
  sendConfirmationNotification(data.employeeName, docTitle);

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
  return {
    employees:     getEmployees().employees,
    documents:     getDocuments().documents,
    confirmations: getConfirmations().confirmations
  };
}

// ============================================================
// EMAIL NOTIFICATIONS
// ============================================================
function sendDocumentNotification(title, category, desc, fileUrl) {
  try {
    const empData = getEmployees();
    const emails = empData.employees
      .map(e => e.email)
      .filter(email => email && email.includes('@'));

    if (emails.length === 0) return;

    const subject = 'Vista Portál: Nový dokument vyžaduje vaše potvrzení';
    const catText = category ? ` [${category}]` : '';

    const htmlBody = `
      <div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;">
        <div style="background:#1a3a2a;padding:24px 32px;border-radius:8px 8px 0 0;">
          <h1 style="color:#c9a84c;margin:0;font-size:22px;letter-spacing:-0.5px;">Vista Portál</h1>
          <p style="color:rgba(255,255,255,0.7);margin:4px 0 0;font-size:13px;">Interní dokumentový systém</p>
        </div>
        <div style="background:#ffffff;padding:32px;border:1px solid #e8e8e4;border-top:none;">
          <p style="color:#1a1a18;font-size:15px;margin:0 0 20px;">Dobrý den,</p>
          <p style="color:#1a1a18;font-size:15px;margin:0 0 20px;">
            byl nahrán nový dokument, který je třeba potvrdit:
          </p>
          <div style="background:#f5f5f3;border-left:4px solid #1a3a2a;padding:16px 20px;border-radius:0 6px 6px 0;margin-bottom:24px;">
            <div style="font-size:16px;font-weight:700;color:#1a3a2a;margin-bottom:4px;">${title}${catText}</div>
            ${desc ? `<div style="font-size:13px;color:#6b6b64;margin-top:4px;">${desc}</div>` : ''}
          </div>
          <p style="color:#6b6b64;font-size:14px;margin:0 0 24px;">
            Přihlaste se do portálu, dokument si přečtěte a klikněte na tlačítko <strong>"Potvrdit přečtení"</strong>.
          </p>
          <a href="${PORTAL_URL}" style="display:inline-block;background:#1a3a2a;color:#ffffff;padding:13px 28px;border-radius:7px;text-decoration:none;font-size:15px;font-weight:600;">
            Otevřít Vista Portál →
          </a>
          ${fileUrl ? `<p style="margin-top:20px;font-size:13px;color:#a0a09a;">Přímý odkaz na soubor: <a href="${fileUrl}" style="color:#2d5a3d;">Stáhnout dokument</a></p>` : ''}
        </div>
        <div style="background:#f5f5f3;padding:16px 32px;border-radius:0 0 8px 8px;border:1px solid #e8e8e4;border-top:none;">
          <p style="color:#a0a09a;font-size:12px;margin:0;">
            Vista Resort &amp; Club · Interní portál · Tento email byl odeslán automaticky.
          </p>
        </div>
      </div>`;

    const plainBody = `Vista Portál — Nový dokument\n\n${title}${catText}\n${desc || ''}\n\nPřihlaste se a potvrďte přečtení: ${PORTAL_URL}`;

    // Send to each employee individually (BCC option)
    emails.forEach(email => {
      GmailApp.sendEmail(email, subject, plainBody, { htmlBody: htmlBody, name: 'Vista Portál' });
    });

  } catch(err) {
    Logger.log('Email error: ' + err.message);
  }
}

function sendConfirmationNotification(employeeName, docTitle) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    if (!adminEmail) return;

    const subject = `✓ Vista Portál: ${employeeName} potvrdil/a dokument`;
    const now = new Date().toLocaleString('cs-CZ', { timeZone: 'Europe/Prague' });

    const htmlBody = `
      <div style="font-family:Arial,sans-serif;max-width:480px;">
        <div style="background:#1a3a2a;padding:20px 28px;border-radius:8px 8px 0 0;">
          <h2 style="color:#c9a84c;margin:0;font-size:18px;">Vista Portál — Potvrzení</h2>
        </div>
        <div style="background:#ffffff;padding:28px;border:1px solid #e8e8e4;border-top:none;">
          <div style="background:#eafaf1;border:1.5px solid #a9dfbf;border-radius:6px;padding:16px 20px;margin-bottom:20px;">
            <div style="font-size:15px;font-weight:700;color:#27ae60;">✓ Dokument potvrzen</div>
          </div>
          <table style="font-size:14px;color:#1a1a18;border-collapse:collapse;width:100%">
            <tr><td style="padding:6px 0;color:#6b6b64;width:120px;">Zaměstnanec</td><td style="padding:6px 0;font-weight:600;">${employeeName}</td></tr>
            <tr><td style="padding:6px 0;color:#6b6b64;">Dokument</td><td style="padding:6px 0;font-weight:600;">${docTitle}</td></tr>
            <tr><td style="padding:6px 0;color:#6b6b64;">Datum a čas</td><td style="padding:6px 0;">${now}</td></tr>
          </table>
          <p style="margin-top:20px;font-size:13px;color:#6b6b64;">
            Celkový přehled potvrzení: <a href="${PORTAL_URL}" style="color:#2d5a3d;">Vista Portál → Admin</a>
          </p>
        </div>
      </div>`;

    GmailApp.sendEmail(adminEmail, subject, `${employeeName} potvrdil/a dokument: ${docTitle} (${now})`, { htmlBody, name: 'Vista Portál' });

  } catch(err) {
    Logger.log('Confirmation email error: ' + err.message);
  }
}

// ============================================================
// HELPERS
// ============================================================
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) { initSheets(); sheet = ss.getSheetByName(name); }
  return sheet;
}

function getDriveFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

// ============================================================
// SETUP — spusť jednou ručně
// ============================================================
function setup() {
  initSheets();
  Logger.log('Vista Portál setup dokončen.');
  Logger.log('Nyní nasaď web app: Nasadit > Nové nasazení > Web app');
}

// ============================================================
// VISTA PORTÁL — Google Apps Script Backend v3.0
// ============================================================

const DRIVE_FOLDER_NAME = 'Vista Portal Dokumenty';
const PORTAL_URL = 'https://davidsiwy.github.io/vista-portal/';
const ADMIN_EMAIL = 'dsiwy2000@gmail.com';  // kam chodí souhrny adminovi

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
  }

  let docSheet = ss.getSheetByName(SHEETS.DOCUMENTS);
  if (!docSheet) {
    docSheet = ss.insertSheet(SHEETS.DOCUMENTS);
    docSheet.appendRow(['id', 'title', 'category', 'desc', 'url', 'fileName', 'uploadedAt', 'deadline', 'audienceType', 'audienceList', 'remindedAt']);
    docSheet.setFrozenRows(1);
    docSheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
  } else {
    // Upgrade: add new columns if missing
    const h = docSheet.getRange(1, 1, 1, docSheet.getLastColumn()).getValues()[0];
    ['deadline', 'audienceType', 'audienceList', 'remindedAt'].forEach(col => {
      if (!h.includes(col)) {
        const c = docSheet.getLastColumn() + 1;
        docSheet.getRange(1, c).setValue(col).setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
      }
    });
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

  const headers = rows[0];
  const idx = {
    id: headers.indexOf('id'), name: headers.indexOf('name'), role: headers.indexOf('role'),
    dept: headers.indexOf('dept'), pin: headers.indexOf('pin'), email: headers.indexOf('email')
  };

  const employees = rows.slice(1).map(r => ({
    id:    r[idx.id]    || '',
    name:  r[idx.name]  || '',
    role:  r[idx.role]  || '',
    dept:  idx.dept  >= 0 ? r[idx.dept]  : '',
    pin:   idx.pin   >= 0 ? String(r[idx.pin]) : '',
    email: idx.email >= 0 ? r[idx.email] : ''
  })).filter(e => e.id);

  return { employees };
}

function addEmployee(data) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const id = data.id || 'emp_' + Date.now();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill('');
  const set = (col, val) => { const i = headers.indexOf(col); if (i >= 0) row[i] = val; };
  set('id', id); set('name', data.name || ''); set('role', data.role || '');
  set('dept', data.dept || ''); set('pin', data.pin || ''); set('email', data.email || '');
  set('createdAt', new Date().toISOString());
  sheet.appendRow(row);
  return { success: true, id };
}

function updateEmployee(data) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idx = {
    id: headers.indexOf('id'), name: headers.indexOf('name'), role: headers.indexOf('role'),
    dept: headers.indexOf('dept'), pin: headers.indexOf('pin'), email: headers.indexOf('email')
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
  const h = rows[0];
  const idx = {
    id: h.indexOf('id'), title: h.indexOf('title'), category: h.indexOf('category'),
    desc: h.indexOf('desc'), url: h.indexOf('url'), fileName: h.indexOf('fileName'),
    uploadedAt: h.indexOf('uploadedAt'), deadline: h.indexOf('deadline'),
    audienceType: h.indexOf('audienceType'), audienceList: h.indexOf('audienceList')
  };
  const documents = rows.slice(1).map(r => ({
    id: r[idx.id], title: r[idx.title], category: r[idx.category], desc: r[idx.desc],
    url: r[idx.url], fileName: r[idx.fileName], uploadedAt: r[idx.uploadedAt],
    deadline: idx.deadline >= 0 ? r[idx.deadline] : '',
    audienceType: idx.audienceType >= 0 ? (r[idx.audienceType] || 'all') : 'all',
    audienceList: idx.audienceList >= 0 ? parseList(r[idx.audienceList]) : []
  })).filter(d => d.id);
  return { documents };
}

function parseList(val) {
  if (!val) return [];
  try { return JSON.parse(val); } catch(e) { return String(val).split(',').map(s => s.trim()).filter(Boolean); }
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
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill('');
  const set = (col, val) => { const i = headers.indexOf(col); if (i >= 0) row[i] = val; };
  set('id', docId);
  set('title', data.title);
  set('category', data.category || '');
  set('desc', data.desc || '');
  set('url', fileUrl);
  set('fileName', data.fileName || '');
  set('uploadedAt', data.uploadedAt || new Date().toISOString());
  set('deadline', data.deadline || '');
  set('audienceType', data.audienceType || 'all');
  set('audienceList', JSON.stringify(data.audienceList || []));
  set('remindedAt', '');
  sheet.appendRow(row);

  // Send notification to targeted employees
  const targets = resolveAudience(data.audienceType || 'all', data.audienceList || []);
  sendDocumentNotification(data.title, data.category || '', data.desc || '', fileUrl, data.deadline || '', targets);

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

// Resolve who should receive a document based on audience settings
function resolveAudience(audienceType, audienceList) {
  const all = getEmployees().employees;
  if (audienceType === 'all' || !audienceType) return all;
  if (audienceType === 'depts') {
    return all.filter(e => audienceList.indexOf(e.dept) >= 0);
  }
  if (audienceType === 'specific') {
    return all.filter(e => audienceList.indexOf(e.id) >= 0);
  }
  return all;
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
// EMAIL: nový dokument
// ============================================================
function sendDocumentNotification(title, category, desc, fileUrl, deadline, targets) {
  try {
    const emails = (targets || getEmployees().employees)
      .map(e => e.email).filter(em => em && em.includes('@'));
    if (emails.length === 0) return;

    const subject = 'Vista Portál: Nový dokument vyžaduje vaše potvrzení';
    const catText = category ? ` [${category}]` : '';
    const deadlineText = deadline ? fmtDateCz(deadline) : '';

    const htmlBody = `
      <div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;">
        <div style="background:#1a3a2a;padding:24px 32px;border-radius:8px 8px 0 0;">
          <h1 style="color:#c9a84c;margin:0;font-size:22px;">Vista Portál</h1>
          <p style="color:rgba(255,255,255,0.7);margin:4px 0 0;font-size:13px;">Interní dokumentový systém</p>
        </div>
        <div style="background:#ffffff;padding:32px;border:1px solid #e8e8e4;border-top:none;">
          <p style="color:#1a1a18;font-size:15px;margin:0 0 20px;">Dobrý den,</p>
          <p style="color:#1a1a18;font-size:15px;margin:0 0 20px;">byl nahrán nový dokument, který je třeba potvrdit:</p>
          <div style="background:#f5f5f3;border-left:4px solid #1a3a2a;padding:16px 20px;border-radius:0 6px 6px 0;margin-bottom:24px;">
            <div style="font-size:16px;font-weight:700;color:#1a3a2a;">${title}${catText}</div>
            ${desc ? `<div style="font-size:13px;color:#6b6b64;margin-top:4px;">${desc}</div>` : ''}
            ${deadlineText ? `<div style="font-size:13px;color:#c0392b;margin-top:8px;font-weight:600;">⏰ Potvrďte do: ${deadlineText}</div>` : ''}
          </div>
          <p style="color:#6b6b64;font-size:14px;margin:0 0 24px;">Přihlaste se do portálu a klikněte na <strong>"Potvrdit přečtení"</strong>.</p>
          <a href="${PORTAL_URL}" style="display:inline-block;background:#1a3a2a;color:#ffffff;padding:13px 28px;border-radius:7px;text-decoration:none;font-size:15px;font-weight:600;">Otevřít Vista Portál →</a>
          ${fileUrl ? `<p style="margin-top:20px;font-size:13px;color:#a0a09a;">Přímý odkaz: <a href="${fileUrl}" style="color:#2d5a3d;">Stáhnout dokument</a></p>` : ''}
        </div>
        <div style="background:#f5f5f3;padding:16px 32px;border-radius:0 0 8px 8px;border:1px solid #e8e8e4;border-top:none;">
          <p style="color:#a0a09a;font-size:12px;margin:0;">Vista Resort &amp; Club · Interní portál · Odesláno automaticky.</p>
        </div>
      </div>`;

    const plainBody = `Vista Portál — Nový dokument\n\n${title}${catText}\n${desc || ''}\n${deadlineText ? 'Potvrďte do: ' + deadlineText + '\n' : ''}\nPotvrďte přečtení: ${PORTAL_URL}`;

    emails.forEach(email => {
      GmailApp.sendEmail(email, subject, plainBody, { htmlBody, name: 'Vista Portál' });
    });
  } catch(err) {
    Logger.log('Email error: ' + err.message);
  }
}

// ============================================================
// EMAIL: potvrzení adminovi
// ============================================================
function sendConfirmationNotification(employeeName, docTitle) {
  try {
    const adminEmail = ADMIN_EMAIL || Session.getActiveUser().getEmail();
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
          <p style="margin-top:20px;font-size:13px;color:#6b6b64;">Přehled: <a href="${PORTAL_URL}" style="color:#2d5a3d;">Vista Portál → Admin</a></p>
        </div>
      </div>`;

    GmailApp.sendEmail(adminEmail, subject, `${employeeName} potvrdil/a: ${docTitle} (${now})`, { htmlBody, name: 'Vista Portál' });
  } catch(err) {
    Logger.log('Confirmation email error: ' + err.message);
  }
}

// ============================================================
// DENNÍ KONTROLA TERMÍNŮ — připomínky
// Spouští se automaticky (time-driven trigger). Nastav přes setupTrigger().
// ============================================================
function checkDeadlines() {
  const docs  = getDocuments().documents;
  const emps  = getEmployees().employees;
  const confs = getConfirmations().confirmations;
  const now   = new Date();

  let adminSummary = [];

  docs.forEach(doc => {
    if (!doc.deadline) return;
    const deadline = new Date(doc.deadline);

    // Komu byl dokument určen
    const targets = resolveAudience(doc.audienceType, doc.audienceList);

    // Kdo nepotvrdil
    const notConfirmed = targets.filter(emp =>
      !confs.some(c => c.docId === doc.id && c.employeeId === emp.id)
    );

    if (notConfirmed.length === 0) return;

    const afterDeadline = now > deadline;

    notConfirmed.forEach(emp => {
      if (!emp.email || !emp.email.includes('@')) return;
      sendReminder(emp, doc, afterDeadline);
    });

    // Souhrn pro admina
    adminSummary.push({
      doc: doc.title,
      deadline: fmtDateCz(doc.deadline),
      afterDeadline: afterDeadline,
      notConfirmed: notConfirmed.map(e => e.name + (e.dept ? ' (' + e.dept + ')' : ''))
    });
  });

  if (adminSummary.length > 0) {
    sendAdminReminderSummary(adminSummary);
  }
}

function sendReminder(emp, doc, afterDeadline) {
  try {
    const subject = afterDeadline
      ? `⚠️ Vista Portál: Prošlý termín potvrzení — ${doc.title}`
      : `Připomínka Vista Portál: Potvrďte dokument ${doc.title}`;
    const deadlineText = doc.deadline ? fmtDateCz(doc.deadline) : '';
    const intro = afterDeadline
      ? `Termín pro potvrzení dokumentu <strong>již vypršel</strong>. Potvrďte ho prosím co nejdříve:`
      : `Připomínáme, že máte k potvrzení dokument:`;

    const htmlBody = `
      <div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;">
        <div style="background:${afterDeadline ? '#c0392b' : '#1a3a2a'};padding:24px 32px;border-radius:8px 8px 0 0;">
          <h1 style="color:${afterDeadline ? '#ffffff' : '#c9a84c'};margin:0;font-size:22px;">Vista Portál</h1>
          <p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px;">${afterDeadline ? 'Prošlý termín potvrzení' : 'Připomínka'}</p>
        </div>
        <div style="background:#ffffff;padding:32px;border:1px solid #e8e8e4;border-top:none;">
          <p style="color:#1a1a18;font-size:15px;margin:0 0 20px;">Dobrý den ${emp.name},</p>
          <p style="color:#1a1a18;font-size:15px;margin:0 0 20px;">${intro}</p>
          <div style="background:#f5f5f3;border-left:4px solid ${afterDeadline ? '#c0392b' : '#1a3a2a'};padding:16px 20px;border-radius:0 6px 6px 0;margin-bottom:24px;">
            <div style="font-size:16px;font-weight:700;color:#1a3a2a;">${doc.title}${doc.category ? ' [' + doc.category + ']' : ''}</div>
            ${deadlineText ? `<div style="font-size:13px;color:#c0392b;margin-top:8px;font-weight:600;">⏰ Termín: ${deadlineText}</div>` : ''}
          </div>
          <a href="${PORTAL_URL}" style="display:inline-block;background:#1a3a2a;color:#ffffff;padding:13px 28px;border-radius:7px;text-decoration:none;font-size:15px;font-weight:600;">Potvrdit nyní →</a>
        </div>
        <div style="background:#f5f5f3;padding:16px 32px;border-radius:0 0 8px 8px;border:1px solid #e8e8e4;border-top:none;">
          <p style="color:#a0a09a;font-size:12px;margin:0;">Vista Resort &amp; Club · Interní portál</p>
        </div>
      </div>`;

    GmailApp.sendEmail(emp.email, subject, `${intro.replace(/<[^>]+>/g,'')} ${doc.title}. Potvrďte: ${PORTAL_URL}`, { htmlBody, name: 'Vista Portál' });
  } catch(err) {
    Logger.log('Reminder error: ' + err.message);
  }
}

function sendAdminReminderSummary(summary) {
  try {
    const adminEmail = ADMIN_EMAIL || Session.getActiveUser().getEmail();
    if (!adminEmail) return;

    let rows = '';
    summary.forEach(s => {
      const statusColor = s.afterDeadline ? '#c0392b' : '#e67e22';
      const statusText = s.afterDeadline ? 'PO TERMÍNU' : 'Čeká';
      rows += `
        <div style="margin-bottom:18px;padding:16px 20px;background:#f5f5f3;border-left:4px solid ${statusColor};border-radius:0 6px 6px 0;">
          <div style="font-size:15px;font-weight:700;color:#1a3a2a;">${s.doc}</div>
          <div style="font-size:12px;color:${statusColor};font-weight:600;margin:4px 0;">${statusText} · termín ${s.deadline}</div>
          <div style="font-size:13px;color:#6b6b64;margin-top:6px;">Nepotvrdili (${s.notConfirmed.length}): ${s.notConfirmed.join(', ')}</div>
        </div>`;
    });

    const htmlBody = `
      <div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;">
        <div style="background:#1a3a2a;padding:24px 32px;border-radius:8px 8px 0 0;">
          <h1 style="color:#c9a84c;margin:0;font-size:22px;">Vista Portál — Denní přehled</h1>
          <p style="color:rgba(255,255,255,0.7);margin:4px 0 0;font-size:13px;">Nepotvrzené dokumenty</p>
        </div>
        <div style="background:#ffffff;padding:32px;border:1px solid #e8e8e4;border-top:none;">
          <p style="color:#1a1a18;font-size:15px;margin:0 0 20px;">Přehled dokumentů které ještě nebyly potvrzeny:</p>
          ${rows}
          <a href="${PORTAL_URL}" style="display:inline-block;background:#1a3a2a;color:#ffffff;padding:13px 28px;border-radius:7px;text-decoration:none;font-size:15px;font-weight:600;margin-top:8px;">Otevřít admin panel →</a>
        </div>
      </div>`;

    GmailApp.sendEmail(adminEmail, 'Vista Portál: Denní přehled nepotvrzených dokumentů', 'Přehled nepotvrzených dokumentů. Otevřete: ' + PORTAL_URL, { htmlBody, name: 'Vista Portál' });
  } catch(err) {
    Logger.log('Admin summary error: ' + err.message);
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

function fmtDateCz(iso) {
  if (!iso) return '';
  try {
    return new Date(iso).toLocaleDateString('cs-CZ', { day: 'numeric', month: 'long', year: 'numeric' });
  } catch(e) { return String(iso); }
}

// ============================================================
// SETUP — spusť jednou ručně
// ============================================================
function setup() {
  initSheets();
  Logger.log('Vista Portál setup dokončen.');
}

// Nastaví denní automatickou kontrolu termínů (připomínky). Spusť jednou ručně.
function setupTrigger() {
  // Smaž staré triggery pro checkDeadlines
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkDeadlines') ScriptApp.deleteTrigger(t);
  });
  // Vytvoř nový denní trigger (každý den v 8:00)
  ScriptApp.newTrigger('checkDeadlines')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('Europe/Prague')
    .create();
  Logger.log('Denní kontrola termínů nastavena na 8:00 (Europe/Prague).');
}

// Smaže rozbitý list Zaměstnanci a vytvoří čistý. POZOR: smaže všechny zaměstnance.
function resetEmployees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(SHEETS.EMPLOYEES);
  if (old) ss.deleteSheet(old);
  const s = ss.insertSheet(SHEETS.EMPLOYEES);
  s.appendRow(['id', 'name', 'role', 'dept', 'pin', 'email', 'createdAt']);
  s.setFrozenRows(1);
  s.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
  Logger.log('List Zaměstnanci resetován.');
}

// Test připomínek — spusť ručně pro otestování emailů bez čekání na trigger
function testCheckDeadlines() {
  checkDeadlines();
  Logger.log('Kontrola termínů spuštěna ručně. Zkontroluj emaily.');
}

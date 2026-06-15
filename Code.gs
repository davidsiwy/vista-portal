// ============================================================
// VISTA PORTÁL — Google Apps Script Backend v4.0
// Dokumenty + Úkoly + Role
// ============================================================

const DRIVE_FOLDER_NAME = 'Vista Portal Dokumenty';
const PORTAL_URL = 'https://davidsiwy.github.io/vista-portal/';
const ADMIN_EMAIL = 'dsiwy2000@gmail.com';

const SHEETS = {
  EMPLOYEES: 'Zaměstnanci',
  DOCUMENTS: 'Dokumenty',
  CONFIRMATIONS: 'Potvrzení',
  TASKS: 'Úkoly',
  COMMENTS: 'Komentáře'
};

// ============================================================
// ENTRY POINTS
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;
    switch (action) {
      // employees
      case 'getEmployees':    result = getEmployees(); break;
      case 'addEmployee':     result = addEmployee(data); break;
      case 'updateEmployee':  result = updateEmployee(data); break;
      case 'deleteEmployee':  result = deleteEmployee(data.id); break;
      // documents
      case 'getDocuments':    result = getDocuments(); break;
      case 'uploadDocument':  result = uploadDocument(data); break;
      case 'deleteDocument':  result = deleteDocument(data.id); break;
      case 'addConfirmation': result = addConfirmation(data); break;
      // tasks
      case 'getTasks':        result = getTasks(); break;
      case 'addTask':         result = addTask(data); break;
      case 'updateTask':      result = updateTask(data); break;
      case 'updateTaskStatus':result = updateTaskStatus(data); break;
      case 'deleteTask':      result = deleteTask(data.id); break;
      case 'addComment':      result = addComment(data); break;
      case 'addTaskAttachment': result = addTaskAttachment(data); break;
      // all
      case 'getAll':          result = getAll(); break;
      default: result = { error: 'Unknown action: ' + action };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message, stack: err.stack });
  }
}

function doGet(e) {
  try {
    return jsonResponse(getAll());
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// INIT
// ============================================================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(ss, SHEETS.EMPLOYEES, ['id','name','role','dept','pin','email','createdAt']);
  ensureSheet(ss, SHEETS.DOCUMENTS, ['id','title','category','desc','url','fileName','uploadedAt','deadline','audienceType','audienceList','remindedAt']);
  ensureSheet(ss, SHEETS.CONFIRMATIONS, ['docId','docTitle','employeeId','employeeName','confirmedAt']);
  ensureSheet(ss, SHEETS.TASKS, ['id','title','desc','segment','priority','status','assignees','dependsOn','deadline','createdBy','createdByName','createdAt','attachments','remindedAt']);
  ensureSheet(ss, SHEETS.COMMENTS, ['id','taskId','authorId','authorName','text','createdAt']);
  return { success: true };
}

function ensureSheet(ss, name, headers) {
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(headers);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#13302e').setFontColor('#dfc196');
  } else {
    // add any missing columns
    const have = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    headers.forEach(col => {
      if (have.indexOf(col) < 0) {
        const c = s.getLastColumn() + 1;
        s.getRange(1, c).setValue(col).setFontWeight('bold').setBackground('#13302e').setFontColor('#dfc196');
      }
    });
  }
  return s;
}

// ============================================================
// GENERIC ROW HELPERS
// ============================================================
function readSheet(name) {
  const sheet = getSheet(name);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { headers: rows[0] || [], items: [] };
  const headers = rows[0];
  const items = rows.slice(1).map((r, i) => {
    const o = { _row: i + 2 };
    headers.forEach((h, ci) => o[h] = r[ci]);
    return o;
  }).filter(o => o[headers[0]]);
  return { headers, items };
}

function appendByHeaders(name, obj) {
  const sheet = getSheet(name);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => (obj[h] !== undefined && obj[h] !== null) ? obj[h] : '');
  sheet.appendRow(row);
}

function updateRowByHeaders(name, matchCol, matchVal, obj) {
  const sheet = getSheet(name);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const mi = headers.indexOf(matchCol);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][mi]) === String(matchVal)) {
      headers.forEach((h, ci) => {
        if (obj[h] !== undefined) sheet.getRange(i + 1, ci + 1).setValue(obj[h]);
      });
      return true;
    }
  }
  return false;
}

function deleteRowByCol(name, col, val) {
  const sheet = getSheet(name);
  const data = sheet.getDataRange().getValues();
  const ci = data[0].indexOf(col);
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][ci]) === String(val)) { sheet.deleteRow(i + 1); }
  }
  return true;
}

// ============================================================
// EMPLOYEES
// ============================================================
function getEmployees() {
  const { items } = readSheet(SHEETS.EMPLOYEES);
  const employees = items.map(r => ({
    id: r.id || '', name: r.name || '', role: r.role || '',
    dept: r.dept || '', pin: r.pin !== undefined ? String(r.pin) : '',
    email: r.email || ''
  })).filter(e => e.id);
  return { employees };
}

function addEmployee(data) {
  const id = data.id || 'emp_' + Date.now();
  appendByHeaders(SHEETS.EMPLOYEES, {
    id, name: data.name || '', role: data.role || '', dept: data.dept || '',
    pin: data.pin || '', email: data.email || '', createdAt: new Date().toISOString()
  });
  return { success: true, id };
}

function updateEmployee(data) {
  updateRowByHeaders(SHEETS.EMPLOYEES, 'id', data.id, {
    name: data.name || '', role: data.role || '', dept: data.dept || '',
    pin: data.pin || '', email: data.email || ''
  });
  return { success: true };
}

function deleteEmployee(id) { deleteRowByCol(SHEETS.EMPLOYEES, 'id', id); return { success: true }; }

// ============================================================
// DOCUMENTS
// ============================================================
function getDocuments() {
  const { items } = readSheet(SHEETS.DOCUMENTS);
  const documents = items.map(r => ({
    id: r.id, title: r.title, category: r.category, desc: r.desc, url: r.url,
    fileName: r.fileName, uploadedAt: r.uploadedAt, deadline: r.deadline || '',
    audienceType: r.audienceType || 'all', audienceList: parseList(r.audienceList)
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
    fileUrl = saveFileToDrive(data.fileData, data.fileType, data.fileName);
  }
  appendByHeaders(SHEETS.DOCUMENTS, {
    id: docId, title: data.title, category: data.category || '', desc: data.desc || '',
    url: fileUrl, fileName: data.fileName || '', uploadedAt: data.uploadedAt || new Date().toISOString(),
    deadline: data.deadline || '', audienceType: data.audienceType || 'all',
    audienceList: JSON.stringify(data.audienceList || []), remindedAt: ''
  });
  const targets = resolveAudience(data.audienceType || 'all', data.audienceList || []);
  sendDocumentNotification(data.title, data.category || '', data.desc || '', fileUrl, data.deadline || '', targets);
  return { success: true, docId, url: fileUrl };
}

function deleteDocument(id) { deleteRowByCol(SHEETS.DOCUMENTS, 'id', id); return { success: true }; }

function resolveAudience(audienceType, audienceList) {
  const all = getEmployees().employees;
  if (audienceType === 'all' || !audienceType) return all;
  if (audienceType === 'depts') return all.filter(e => audienceList.indexOf(e.dept) >= 0);
  if (audienceType === 'specific') return all.filter(e => audienceList.indexOf(e.id) >= 0);
  return all;
}

// ============================================================
// CONFIRMATIONS
// ============================================================
function addConfirmation(data) {
  const { items } = readSheet(SHEETS.CONFIRMATIONS);
  if (items.some(r => r.docId === data.docId && r.employeeId === data.employeeId)) return { success: true, duplicate: true };
  const docs = readSheet(SHEETS.DOCUMENTS).items;
  const doc = docs.find(d => d.id === data.docId);
  appendByHeaders(SHEETS.CONFIRMATIONS, {
    docId: data.docId, docTitle: doc ? doc.title : data.docId,
    employeeId: data.employeeId, employeeName: data.employeeName,
    confirmedAt: data.confirmedAt || new Date().toISOString()
  });
  sendConfirmationNotification(data.employeeName, doc ? doc.title : data.docId);
  return { success: true };
}

function getConfirmations() {
  const { items } = readSheet(SHEETS.CONFIRMATIONS);
  return { confirmations: items.map(r => ({
    docId: r.docId, docTitle: r.docTitle, employeeId: r.employeeId,
    employeeName: r.employeeName, confirmedAt: r.confirmedAt
  })).filter(c => c.docId) };
}

// ============================================================
// TASKS
// ============================================================
function getTasks() {
  const { items } = readSheet(SHEETS.TASKS);
  const tasks = items.map(r => ({
    id: r.id, title: r.title, desc: r.desc, segment: r.segment, priority: r.priority || 'medium',
    status: r.status || 'todo', assignees: parseList(r.assignees), dependsOn: parseList(r.dependsOn),
    deadline: r.deadline || '', createdBy: r.createdBy, createdByName: r.createdByName,
    createdAt: r.createdAt, attachments: parseAttachments(r.attachments)
  })).filter(t => t.id);
  const comments = getComments().comments;
  return { tasks, comments };
}

function parseAttachments(val) {
  if (!val) return [];
  try { return JSON.parse(val); } catch(e) { return []; }
}

function addTask(data) {
  const id = data.id || 'task_' + Date.now();
  appendByHeaders(SHEETS.TASKS, {
    id, title: data.title, desc: data.desc || '', segment: data.segment || '',
    priority: data.priority || 'medium', status: data.status || 'todo',
    assignees: JSON.stringify(data.assignees || []), dependsOn: JSON.stringify(data.dependsOn || []),
    deadline: data.deadline || '', createdBy: data.createdBy || '', createdByName: data.createdByName || '',
    createdAt: new Date().toISOString(), attachments: JSON.stringify(data.attachments || []), remindedAt: ''
  });
  // Notify assignees
  notifyTaskAssigned(id, data.title, data.desc || '', data.priority || 'medium', data.deadline || '', data.assignees || [], data.createdByName || '');
  return { success: true, id };
}

function updateTask(data) {
  const obj = {};
  ['title','desc','segment','priority','status','deadline'].forEach(k => { if (data[k] !== undefined) obj[k] = data[k]; });
  if (data.assignees !== undefined) obj.assignees = JSON.stringify(data.assignees);
  if (data.dependsOn !== undefined) obj.dependsOn = JSON.stringify(data.dependsOn);
  updateRowByHeaders(SHEETS.TASKS, 'id', data.id, obj);
  return { success: true };
}

function updateTaskStatus(data) {
  updateRowByHeaders(SHEETS.TASKS, 'id', data.id, { status: data.status });
  // Notify creator + assignees about status change
  const tasks = readSheet(SHEETS.TASKS).items;
  const task = tasks.find(t => t.id === data.id);
  if (task) {
    notifyStatusChange(task, data.status, data.changedByName || '');
  }
  return { success: true };
}

function deleteTask(id) {
  deleteRowByCol(SHEETS.TASKS, 'id', id);
  deleteRowByCol(SHEETS.COMMENTS, 'taskId', id);
  return { success: true };
}

function addTaskAttachment(data) {
  const tasks = readSheet(SHEETS.TASKS).items;
  const task = tasks.find(t => t.id === data.taskId);
  if (!task) return { error: 'Task not found' };
  let url = '';
  if (data.fileData && data.fileName) {
    url = saveFileToDrive(data.fileData, data.fileType, data.fileName);
  }
  const atts = parseAttachments(task.attachments);
  atts.push({ name: data.fileName, url: url, uploadedAt: new Date().toISOString() });
  updateRowByHeaders(SHEETS.TASKS, 'id', data.taskId, { attachments: JSON.stringify(atts) });
  return { success: true, url, attachments: atts };
}

// ============================================================
// COMMENTS
// ============================================================
function addComment(data) {
  const id = 'cmt_' + Date.now();
  appendByHeaders(SHEETS.COMMENTS, {
    id, taskId: data.taskId, authorId: data.authorId || '', authorName: data.authorName || '',
    text: data.text || '', createdAt: new Date().toISOString()
  });
  // Notify task assignees + creator about new comment
  const tasks = readSheet(SHEETS.TASKS).items;
  const task = tasks.find(t => t.id === data.taskId);
  if (task) notifyNewComment(task, data.authorName || '', data.text || '');
  return { success: true, id };
}

function getComments() {
  const { items } = readSheet(SHEETS.COMMENTS);
  return { comments: items.map(r => ({
    id: r.id, taskId: r.taskId, authorId: r.authorId, authorName: r.authorName,
    text: r.text, createdAt: r.createdAt
  })).filter(c => c.id) };
}

// ============================================================
// GET ALL
// ============================================================
function getAll() {
  const t = getTasks();
  return {
    employees:     getEmployees().employees,
    documents:     getDocuments().documents,
    confirmations: getConfirmations().confirmations,
    tasks:         t.tasks,
    comments:      t.comments
  };
}

// ============================================================
// DRIVE
// ============================================================
function saveFileToDrive(fileData, fileType, fileName) {
  const folder = getDriveFolder();
  const blob = Utilities.newBlob(Utilities.base64Decode(fileData), fileType || 'application/octet-stream', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getDownloadUrl();
}

function getDriveFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

// ============================================================
// EMAIL — shared shell
// ============================================================
function emailShell(headerTitle, headerSub, bodyHtml, accent) {
  accent = accent || '#dfc196';
  return `
    <div style="font-family:Georgia,'Times New Roman',serif;max-width:560px;margin:0 auto;background:#ffffff;">
      <div style="background:#13302e;padding:28px 36px;">
        <div style="color:${accent};font-size:24px;font-weight:bold;letter-spacing:1px;">VISTA RESORT</div>
        <div style="color:rgba(255,255,255,0.6);margin:6px 0 0;font-size:12px;letter-spacing:2px;text-transform:uppercase;font-family:Arial,sans-serif;">${headerSub}</div>
      </div>
      <div style="padding:36px;border:1px solid #e8e5dd;border-top:none;font-family:Arial,sans-serif;">
        ${bodyHtml}
      </div>
      <div style="background:#f5f3ee;padding:18px 36px;border:1px solid #e8e5dd;border-top:none;">
        <p style="color:#9c7852;font-size:12px;margin:0;font-family:Arial,sans-serif;">Vista Resort &amp; Club · Interní portál · Odesláno automaticky.</p>
      </div>
    </div>`;
}

function btnHtml(text, url) {
  return `<a href="${url}" style="display:inline-block;background:#13302e;color:#dfc196;padding:13px 28px;border-radius:4px;text-decoration:none;font-size:15px;font-weight:bold;font-family:Arial,sans-serif;">${text}</a>`;
}

// ============================================================
// EMAIL — documents
// ============================================================
function sendDocumentNotification(title, category, desc, fileUrl, deadline, targets) {
  try {
    const emails = (targets || getEmployees().employees).map(e => e.email).filter(em => em && em.includes('@'));
    if (emails.length === 0) return;
    const catText = category ? ` [${category}]` : '';
    const deadlineText = deadline ? fmtDateCz(deadline) : '';
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">Dobrý den,</p>
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">byl nahrán nový dokument, který je třeba potvrdit:</p>
      <div style="background:#f5f3ee;border-left:4px solid #13302e;padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${title}${catText}</div>
        ${desc ? `<div style="font-size:13px;color:#9c7852;margin-top:4px;">${desc}</div>` : ''}
        ${deadlineText ? `<div style="font-size:13px;color:#9c3a2e;margin-top:8px;font-weight:bold;">Potvrďte do: ${deadlineText}</div>` : ''}
      </div>
      <p style="color:#9c7852;font-size:14px;margin:0 0 24px;">Přihlaste se do portálu a klikněte na „Potvrdit přečtení".</p>
      ${btnHtml('Otevřít Vista Portál', PORTAL_URL)}`;
    const html = emailShell('', 'Interní dokumentový systém', body);
    const plain = `Vista Portál — Nový dokument\n\n${title}${catText}\n${desc||''}\n${deadlineText?'Potvrďte do: '+deadlineText+'\n':''}\n${PORTAL_URL}`;
    emails.forEach(em => GmailApp.sendEmail(em, 'Vista Portál: Nový dokument k potvrzení', plain, { htmlBody: html, name: 'Vista Resort' }));
  } catch(err) { Logger.log('doc email: ' + err.message); }
}

function sendConfirmationNotification(employeeName, docTitle) {
  try {
    const adminEmail = ADMIN_EMAIL;
    if (!adminEmail) return;
    const now = new Date().toLocaleString('cs-CZ', { timeZone: 'Europe/Prague' });
    const body = `
      <div style="background:#eef5ee;border:1px solid #b8d4b0;padding:16px 20px;margin-bottom:20px;">
        <div style="font-size:15px;font-weight:bold;color:#586845;">Dokument potvrzen</div>
      </div>
      <table style="font-size:14px;color:#3d301f;width:100%">
        <tr><td style="padding:6px 0;color:#9c7852;width:120px;">Zaměstnanec</td><td style="padding:6px 0;font-weight:bold;">${employeeName}</td></tr>
        <tr><td style="padding:6px 0;color:#9c7852;">Dokument</td><td style="padding:6px 0;font-weight:bold;">${docTitle}</td></tr>
        <tr><td style="padding:6px 0;color:#9c7852;">Datum</td><td style="padding:6px 0;">${now}</td></tr>
      </table>`;
    GmailApp.sendEmail(adminEmail, `Vista Portál: ${employeeName} potvrdil/a dokument`, `${employeeName} potvrdil/a: ${docTitle} (${now})`, { htmlBody: emailShell('', 'Potvrzení dokumentu', body), name: 'Vista Resort' });
  } catch(err) { Logger.log('conf email: ' + err.message); }
}

// ============================================================
// EMAIL — tasks
// ============================================================
function empById(id) { return getEmployees().employees.find(e => e.id === id); }
function emailsForIds(ids) {
  const emps = getEmployees().employees;
  return ids.map(id => { const e = emps.find(x => x.id === id); return e ? e.email : null; }).filter(em => em && em.includes('@'));
}

function priorityCz(p) { return p === 'high' ? 'Vysoká' : p === 'low' ? 'Nízká' : 'Střední'; }
function statusCz(s) {
  return s === 'todo' ? 'Zadáno' : s === 'inprogress' ? 'Probíhá' : s === 'review' ? 'Ke kontrole' : s === 'done' ? 'Hotovo' : s;
}

function notifyTaskAssigned(taskId, title, desc, priority, deadline, assigneeIds, createdByName) {
  try {
    const emails = emailsForIds(assigneeIds || []);
    if (emails.length === 0) return;
    const deadlineText = deadline ? fmtDateCz(deadline) : '';
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">Dobrý den,</p>
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">byl vám přiřazen nový úkol${createdByName ? ' (zadal: ' + createdByName + ')' : ''}:</p>
      <div style="background:#f5f3ee;border-left:4px solid #13302e;padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${title}</div>
        ${desc ? `<div style="font-size:13px;color:#9c7852;margin-top:6px;">${desc}</div>` : ''}
        <div style="font-size:13px;color:#3d301f;margin-top:8px;">Priorita: <strong>${priorityCz(priority)}</strong>${deadlineText ? ` · Termín: <strong style="color:#9c3a2e">${deadlineText}</strong>` : ''}</div>
      </div>
      ${btnHtml('Otevřít úkol', PORTAL_URL)}`;
    const html = emailShell('', 'Nový úkol', body);
    emails.forEach(em => GmailApp.sendEmail(em, 'Vista Portál: Nový úkol — ' + title, `Nový úkol: ${title}. ${PORTAL_URL}`, { htmlBody: html, name: 'Vista Resort' }));
  } catch(err) { Logger.log('task assign email: ' + err.message); }
}

function notifyStatusChange(task, newStatus, changedByName) {
  try {
    // notify creator + assignees (except the one who changed, but we don't know their email reliably; send to all involved)
    const ids = parseList(task.assignees).slice();
    if (task.createdBy) ids.push(task.createdBy);
    const emails = [...new Set(emailsForIds(ids))];
    if (emails.length === 0) return;
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">Změna stavu úkolu:</p>
      <div style="background:#f5f3ee;border-left:4px solid #13302e;padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${task.title}</div>
        <div style="font-size:14px;color:#3d301f;margin-top:8px;">Nový stav: <strong>${statusCz(newStatus)}</strong>${changedByName ? ` · změnil/a ${changedByName}` : ''}</div>
      </div>
      ${btnHtml('Otevřít úkol', PORTAL_URL)}`;
    const html = emailShell('', 'Změna stavu úkolu', body);
    emails.forEach(em => GmailApp.sendEmail(em, `Vista Portál: ${task.title} → ${statusCz(newStatus)}`, `${task.title} změněn na ${statusCz(newStatus)}. ${PORTAL_URL}`, { htmlBody: html, name: 'Vista Resort' }));
  } catch(err) { Logger.log('status email: ' + err.message); }
}

function notifyNewComment(task, authorName, text) {
  try {
    const ids = parseList(task.assignees).slice();
    if (task.createdBy) ids.push(task.createdBy);
    const emails = [...new Set(emailsForIds(ids))];
    if (emails.length === 0) return;
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 16px;">Nový komentář u úkolu:</p>
      <div style="background:#f5f3ee;border-left:4px solid #13302e;padding:16px 20px;margin-bottom:20px;">
        <div style="font-size:15px;font-weight:bold;color:#13302e;">${task.title}</div>
        <div style="font-size:14px;color:#3d301f;margin-top:10px;padding:10px 14px;background:#ffffff;border-radius:4px;"><strong>${authorName}:</strong> ${text}</div>
      </div>
      ${btnHtml('Odpovědět v portálu', PORTAL_URL)}`;
    const html = emailShell('', 'Nový komentář', body);
    emails.forEach(em => GmailApp.sendEmail(em, `Vista Portál: Komentář — ${task.title}`, `${authorName}: ${text} (${PORTAL_URL})`, { htmlBody: html, name: 'Vista Resort' }));
  } catch(err) { Logger.log('comment email: ' + err.message); }
}

// ============================================================
// DAILY CHECK — deadlines (docs + tasks)
// ============================================================
function checkDeadlines() {
  checkDocumentDeadlines();
  checkTaskDeadlines();
}

function checkDocumentDeadlines() {
  const docs = getDocuments().documents;
  const confs = getConfirmations().confirmations;
  const now = new Date();
  let summary = [];
  docs.forEach(doc => {
    if (!doc.deadline) return;
    const deadline = new Date(doc.deadline);
    const targets = resolveAudience(doc.audienceType, doc.audienceList);
    const notConfirmed = targets.filter(emp => !confs.some(c => c.docId === doc.id && c.employeeId === emp.id));
    if (notConfirmed.length === 0) return;
    const afterDeadline = now > deadline;
    const dayBefore = !afterDeadline && (deadline - now) <= 24*60*60*1000;
    if (!afterDeadline && !dayBefore) return; // only remind day-before or after
    notConfirmed.forEach(emp => { if (emp.email && emp.email.includes('@')) sendDocReminder(emp, doc, afterDeadline); });
    summary.push({ kind: 'Dokument', title: doc.title, deadline: fmtDateCz(doc.deadline), afterDeadline, names: notConfirmed.map(e => e.name + (e.dept ? ' ('+e.dept+')' : '')) });
  });
  if (summary.length) sendAdminSummary(summary, 'Nepotvrzené dokumenty');
}

function checkTaskDeadlines() {
  const { tasks } = getTasks();
  const now = new Date();
  let summary = [];
  tasks.forEach(task => {
    if (!task.deadline || task.status === 'done') return;
    const deadline = new Date(task.deadline);
    const afterDeadline = now > deadline;
    const dayBefore = !afterDeadline && (deadline - now) <= 24*60*60*1000;
    if (!afterDeadline && !dayBefore) return;
    const emails = emailsForIds(task.assignees);
    const emps = getEmployees().employees;
    task.assignees.forEach(id => {
      const e = emps.find(x => x.id === id);
      if (e && e.email && e.email.includes('@')) sendTaskReminder(e, task, afterDeadline);
    });
    const names = task.assignees.map(id => { const e = emps.find(x => x.id === id); return e ? e.name : id; });
    summary.push({ kind: 'Úkol', title: task.title, deadline: fmtDateCz(task.deadline), afterDeadline, status: statusCz(task.status), names });
  });
  if (summary.length) sendAdminSummary(summary, 'Úkoly po termínu / blížící se termín');
}

function sendDocReminder(emp, doc, afterDeadline) {
  try {
    const subject = afterDeadline ? `Vista Portál: Prošlý termín — ${doc.title}` : `Vista Portál: Připomínka — ${doc.title}`;
    const deadlineText = doc.deadline ? fmtDateCz(doc.deadline) : '';
    const intro = afterDeadline ? 'Termín pro potvrzení dokumentu <strong>již vypršel</strong>. Potvrďte ho prosím:' : 'Připomínáme dokument k potvrzení:';
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">Dobrý den ${emp.name},</p>
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">${intro}</p>
      <div style="background:#f5f3ee;border-left:4px solid ${afterDeadline ? '#9c3a2e' : '#13302e'};padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${doc.title}${doc.category ? ' ['+doc.category+']' : ''}</div>
        ${deadlineText ? `<div style="font-size:13px;color:#9c3a2e;margin-top:8px;font-weight:bold;">Termín: ${deadlineText}</div>` : ''}
      </div>
      ${btnHtml('Potvrdit nyní', PORTAL_URL)}`;
    GmailApp.sendEmail(emp.email, subject, `${doc.title} — potvrďte: ${PORTAL_URL}`, { htmlBody: emailShell('', afterDeadline ? 'Prošlý termín' : 'Připomínka', body), name: 'Vista Resort' });
  } catch(err) { Logger.log('doc reminder: ' + err.message); }
}

function sendTaskReminder(emp, task, afterDeadline) {
  try {
    const subject = afterDeadline ? `Vista Portál: Úkol po termínu — ${task.title}` : `Vista Portál: Blíží se termín — ${task.title}`;
    const deadlineText = task.deadline ? fmtDateCz(task.deadline) : '';
    const intro = afterDeadline ? 'Termín úkolu <strong>již vypršel</strong> a úkol stále není hotový:' : 'Blíží se termín úkolu:';
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">Dobrý den ${emp.name},</p>
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">${intro}</p>
      <div style="background:#f5f3ee;border-left:4px solid ${afterDeadline ? '#9c3a2e' : '#13302e'};padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${task.title}</div>
        <div style="font-size:13px;color:#3d301f;margin-top:6px;">Stav: <strong>${statusCz(task.status)}</strong>${deadlineText ? ` · Termín: <strong style="color:#9c3a2e">${deadlineText}</strong>` : ''}</div>
      </div>
      ${btnHtml('Otevřít úkol', PORTAL_URL)}`;
    GmailApp.sendEmail(emp.email, subject, `${task.title} — ${PORTAL_URL}`, { htmlBody: emailShell('', afterDeadline ? 'Úkol po termínu' : 'Blížící se termín', body), name: 'Vista Resort' });
  } catch(err) { Logger.log('task reminder: ' + err.message); }
}

function sendAdminSummary(summary, heading) {
  try {
    const adminEmail = ADMIN_EMAIL;
    if (!adminEmail) return;
    let rows = '';
    summary.forEach(s => {
      const color = s.afterDeadline ? '#9c3a2e' : '#9c7852';
      const status = s.afterDeadline ? 'PO TERMÍNU' : 'Blíží se termín';
      rows += `
        <div style="margin-bottom:16px;padding:16px 20px;background:#f5f3ee;border-left:4px solid ${color};">
          <div style="font-size:15px;font-weight:bold;color:#13302e;">${s.kind}: ${s.title}</div>
          <div style="font-size:12px;color:${color};font-weight:bold;margin:4px 0;">${status} · termín ${s.deadline}${s.status ? ' · stav '+s.status : ''}</div>
          <div style="font-size:13px;color:#9c7852;margin-top:6px;">Dotčení (${s.names.length}): ${s.names.join(', ')}</div>
        </div>`;
    });
    const body = `<p style="color:#3d301f;font-size:15px;margin:0 0 20px;">${heading}:</p>${rows}${btnHtml('Otevřít admin panel', PORTAL_URL)}`;
    GmailApp.sendEmail(adminEmail, 'Vista Portál: Denní přehled — ' + heading, heading + ': ' + PORTAL_URL, { htmlBody: emailShell('', 'Denní přehled', body), name: 'Vista Resort' });
  } catch(err) { Logger.log('admin summary: ' + err.message); }
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

function fmtDateCz(iso) {
  if (!iso) return '';
  try { return new Date(iso).toLocaleDateString('cs-CZ', { day: 'numeric', month: 'long', year: 'numeric' }); }
  catch(e) { return String(iso); }
}

// ============================================================
// SETUP
// ============================================================
function setup() {
  initSheets();
  Logger.log('Vista Portál v4 setup dokončen. Sheety: Zaměstnanci, Dokumenty, Potvrzení, Úkoly, Komentáře.');
}

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'checkDeadlines') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('checkDeadlines').timeBased().atHour(8).everyDays(1).inTimezone('Europe/Prague').create();
  Logger.log('Denní kontrola termínů nastavena na 8:00.');
}

function resetEmployees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(SHEETS.EMPLOYEES);
  if (old) ss.deleteSheet(old);
  ensureSheet(ss, SHEETS.EMPLOYEES, ['id','name','role','dept','pin','email','createdAt']);
  Logger.log('List Zaměstnanci resetován.');
}

function testCheckDeadlines() { checkDeadlines(); Logger.log('Kontrola spuštěna ručně.'); }

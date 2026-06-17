// ============================================================
// VISTA PORTÁL — Google Apps Script Backend v5.0
// Dokumenty + Úkoly + Faktury + Deník + Reporty + Role
// ============================================================

const DRIVE_FOLDER_NAME = 'Vista Portal Dokumenty';
const PORTAL_URL = 'https://internal.vistaresort.cz/';
const ADMIN_EMAIL = 'gabriela.vachova@vistaresort.cz';  // Administrátor — dostává souhrny potvrzení a přehled termínů
const GM_EMAIL    = 'katerina.novakova@vistaresort.cz';  // General Manager — kopie důležitých notifikací
const OWNER_EMAIL = '';   // PRÁZDNÉ = majitel nedostává emaily (vše vidí v portálu). Pro zapnutí: 'siwy@rscredit.cz'


// Odešle email na admina i GM (bez duplicit)
function sendToManagement(subject, plainText, htmlBody) {
  const recipients = [...new Set([ADMIN_EMAIL, GM_EMAIL].filter(e => e && e.includes('@')))];
  recipients.forEach(em => {
    try { GmailApp.sendEmail(em, subject, plainText, { htmlBody, name: 'Vista Resort' }); }
    catch(err) { Logger.log('mgmt email to ' + em + ': ' + err.message); }
  });
}

const SHEETS = {
  EMPLOYEES: 'Zaměstnanci',
  DOCUMENTS: 'Dokumenty',
  CONFIRMATIONS: 'Potvrzení',
  TASKS: 'Úkoly',
  COMMENTS: 'Komentáře',
  INVOICES: 'Faktury',
  WORKLOG: 'Deník',
  TRZBY: 'Tržby',
  REPORTS: 'Reporty',
  NOTIFICATIONS: 'Notifikace',
  TEMPLATES: 'Šablony'
};

// ============================================================
// AUTENTIZACE — PINy jsou TADY na serveru (ne v prohlížeči).
// Endpoint vyžaduje platný token (kromě akce 'login').
// ============================================================
const AUTH = {
  pins: {
    owner: 'VistaMajitel2050',
    gm:    'VistaGM2050',
    admin: 'VistaResort2050'
  },
  // Podpisový klíč tokenů. KLIDNĚ ZMĚŇ na vlastní náhodný řetězec (a pak znovu Deploy).
  secret: 'Vista-Token-Secret-2026-eY7pQ2mZ',
  ttlDays: 30,
  // Když true, endpoint odmítne požadavky bez platného tokenu.
  // Pokud by se cokoli pokazilo a nešlo se přihlásit, dočasně dej false a Deploy.
  enforce: true
};

function _hmacHex(message) {
  const raw = Utilities.computeHmacSha256Signature(message, AUTH.secret);
  return raw.map(function(b){ return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function makeToken(role, uid, name) {
  const payload = { role: role, uid: uid, name: name, exp: Date.now() + AUTH.ttlDays * 86400000 };
  const pstr = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=+$/, '');
  return pstr + '.' + _hmacHex(pstr);
}
function validateToken(token) {
  if (!token || String(token).indexOf('.') < 0) return null;
  const parts = String(token).split('.');
  if (parts.length !== 2) return null;
  if (_hmacHex(parts[0]) !== parts[1]) return null;
  try {
    const json = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    const payload = JSON.parse(json);
    if (!payload.exp || payload.exp < Date.now()) return null;
    return payload;
  } catch (e) { return null; }
}
function doLoginAction(data) {
  const pin = String(data.pin || '').trim();
  if (!pin) return { error: 'Zadejte PIN', code: 401 };
  if (pin === AUTH.pins.owner) return { success: true, role: 'owner', user: { id: '__owner__', name: 'Majitel', role: 'Majitel', dept: '' }, token: makeToken('owner', '__owner__', 'Majitel') };
  if (pin === AUTH.pins.gm)    return { success: true, role: 'gm',    user: { id: '__gm__', name: 'General Manager', role: 'GM', dept: '' }, token: makeToken('gm', '__gm__', 'General Manager') };
  if (pin === AUTH.pins.admin) return { success: true, role: 'admin', user: { id: '__admin__', name: 'Administrátor', role: 'Admin', dept: '' }, token: makeToken('admin', '__admin__', 'Administrátor') };
  const emp = readSheet(SHEETS.EMPLOYEES).items.find(function(e){ return String(e.pin).trim() === pin; });
  if (emp) {
    var role = emp.accessLevel === 'segment_manager' ? 'segment_manager' : 'employee';
    return { success: true, role: role, user: { id: emp.id, name: emp.name, role: emp.role, dept: emp.dept, accessLevel: emp.accessLevel||'standard' }, token: makeToken(role, emp.id, emp.name) };
  }
  return { error: 'Nesprávný PIN', code: 401 };
}

// ============================================================
// ENTRY POINTS
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    // 'login' je veřejná akce; vše ostatní vyžaduje platný token (když AUTH.enforce)
    if (action !== 'login' && AUTH.enforce) {
      if (!validateToken(data.token)) {
        return jsonResponse({ error: 'Neautorizováno. Přihlaste se prosím znovu.', code: 401 });
      }
    }

    let result;
    switch (action) {
      // auth
      case 'login':           result = doLoginAction(data); break;
      case 'getEssential':    result = getEssential(); break;
      // employees
      case 'getEmployees':    result = getEmployees(); break;
      case 'addEmployee':     result = addEmployee(data); break;
      case 'updateEmployee':  result = updateEmployee(data); break;
      case 'deleteEmployee':  result = deleteEmployee(data.id); break;
      // documents
      case 'getDocuments':    result = getDocuments(); break;
      case 'uploadDocument':  result = uploadDocument(data); break;
      case 'updateDocument':  result = updateDocument(data); break;
      case 'deleteDocument':  result = deleteDocument(data.id); break;
      case 'addConfirmation': result = addConfirmation(data); break;
      case 'urgeDocument':    result = urgeDocument(data); break;
      case 'urgeTask':        result = urgeTask(data); break;
      // tasks
      case 'getTasks':        result = getTasks(); break;
      case 'addTask':         result = addTask(data); break;
      case 'updateTask':      result = updateTask(data); break;
      case 'updateTaskStatus':result = updateTaskStatus(data); break;
      case 'deleteTask':      result = deleteTask(data.id); break;
      case 'addComment':      result = addComment(data); break;
      case 'addTaskAttachment': result = addTaskAttachment(data); break;
      // invoices
      case 'getInvoices':     result = getInvoices(); break;
      case 'addInvoice':      result = addInvoice(data); break;
      case 'decideInvoice':   result = decideInvoice(data); break;
      case 'deleteInvoice':   result = deleteInvoice(data.id); break;
      // worklog
      case 'getWorklog':      result = getWorklog(); break;
      case 'saveTrzby':       result = saveTrzby(data); break;
      case 'getTrzby':        result = getTrzby(); break;
      case 'uploadReport':    result = uploadReport(data); break;
      case 'deleteReport':    result = deleteReport(data.id); break;
      // notifications
      case 'getNotifications':  result = getNotifications(data); break;
      case 'markNotifRead':     result = markNotifRead(data); break;
      case 'pushNotification':  result = pushNotification(data); break;
      // task templates
      case 'saveTemplate':           result = saveTemplate(data); break;
      case 'deleteTemplate':         result = deleteTemplate(data.id); break;
      case 'updateTemplateLastRun':  result = updateTemplateLastRun(data); break;
      case 'addWorklog':      result = addWorklog(data); break;
      case 'updateWorklog':   result = updateWorklog(data); break;
      case 'deleteWorklog':   result = deleteWorklog(data.id); break;
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
  // Data se vrací jen přes autentizovaný POST. GET nevrací žádná citlivá data.
  return jsonResponse({ status: 'Vista Portál API', message: 'Použijte aplikaci pro přístup.' });
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// INIT
// ============================================================
function initSheets() {
  const ss = getSS();
  ensureSheet(ss, SHEETS.EMPLOYEES, ['id','name','role','dept','accessLevel','pin','email','createdAt']);
  ensureSheet(ss, SHEETS.DOCUMENTS, ['id','title','category','desc','url','fileId','fileName','uploadedAt','deadline','audienceType','audienceList','version','requiresSignature','readOnly','remindedAt']);
  ensureSheet(ss, SHEETS.CONFIRMATIONS, ['docId','docTitle','employeeId','employeeName','confirmedAt','signedUrl','signedFileId','signedFileName']);
  ensureSheet(ss, SHEETS.TASKS, ['id','title','desc','segment','priority','status','assignees','dependsOn','deadline','createdBy','createdByName','createdAt','completedAt','attachments','remindedAt']);
  ensureSheet(ss, SHEETS.COMMENTS, ['id','taskId','authorId','authorName','text','createdAt','entityType']);
  ensureSheet(ss, SHEETS.INVOICES, ['id','title','supplier','amount','currency','url','fileId','fileName','dueDate','uploadedBy','uploadedByName','uploadedAt','status','decidedBy','decidedByName','decidedAt']);
  ensureSheet(ss, SHEETS.WORKLOG, ['id','employeeId','employeeName','segment','date','text','createdAt']);
  ensureSheet(ss, SHEETS.TRZBY, ['key','json','uploadedBy','uploadedAt']);
  ensureSheet(ss, SHEETS.REPORTS, ['id','title','category','desc','url','fileId','fileName','uploadedBy','uploadedByName','uploadedAt']);
  ensureSheet(ss, SHEETS.NOTIFICATIONS, ['id','recipientId','type','icon','message','link','createdAt','readBy']);
  ensureSheet(ss, SHEETS.TEMPLATES, ['id','title','desc','segment','priority','assignees','freq','days','monthDay','deadlineOffset','active','createdAt','lastRun']);
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
      headers.forEach((h, ci) => { if (obj[h] !== undefined) sheet.getRange(i + 1, ci + 1).setValue(obj[h]); });
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
    if (String(data[i][ci]) === String(val)) sheet.deleteRow(i + 1);
  }
  return true;
}

// ============================================================
// EMPLOYEES
// ============================================================
function getEmployees() {
  const { items } = readSheet(SHEETS.EMPLOYEES);
  const employees = items.map(r => ({
    id: r.id || '', name: r.name || '', role: r.role || '', dept: r.dept || '',
    pin: r.pin !== undefined ? String(r.pin) : '', email: r.email || ''
  })).filter(e => e.id);
  return { employees };
}
function addEmployee(data) {
  const id = data.id || 'emp_' + Date.now();
  appendByHeaders(SHEETS.EMPLOYEES, { id, name: data.name||'', role: data.role||'', dept: data.dept||'', accessLevel: data.accessLevel||'standard', pin: data.pin||'', email: data.email||'', createdAt: new Date().toISOString() });
  return { success: true, id };
}
function updateEmployee(data) {
  updateRowByHeaders(SHEETS.EMPLOYEES, 'id', data.id, { name: data.name||'', role: data.role||'', dept: data.dept||'', accessLevel: data.accessLevel||'standard', pin: data.pin||'', email: data.email||'' });
  return { success: true };
}
function deleteEmployee(id) { deleteRowByCol(SHEETS.EMPLOYEES, 'id', id); return { success: true }; }

// ============================================================
// DOCUMENTS
// ============================================================
function getDocuments() {
  const { items } = readSheet(SHEETS.DOCUMENTS);
  const documents = items.map(r => ({
    id: r.id, title: r.title, category: r.category, desc: r.desc, url: r.url, fileId: r.fileId || '',
    fileName: r.fileName, uploadedAt: r.uploadedAt, deadline: r.deadline || '', version: r.version || 1,
    requiresSignature: r.requiresSignature === true || r.requiresSignature === 'true' || r.requiresSignature === 1 || r.requiresSignature === 'ANO',
    readOnly: r.readOnly === true || r.readOnly === 'true' || r.readOnly === 1 || r.readOnly === 'ANO',
    audienceType: r.audienceType || 'all', audienceList: parseList(r.audienceList)
  })).filter(d => d.id);
  return { documents };
}
function parseList(val) {
  if (!val) return [];
  try { return JSON.parse(val); } catch(e) { return String(val).split(',').map(s => s.trim()).filter(Boolean); }
}
function uploadDocument(data) {
  let fileUrl = '', fileId = '';
  const docId = data.id || 'doc_' + Date.now();
  if (data.fileData && data.fileName) {
    const saved = saveFileToDrive(data.fileData, data.fileType, data.fileName);
    fileUrl = saved.url; fileId = saved.id;
  }
  appendByHeaders(SHEETS.DOCUMENTS, {
    id: docId, title: data.title, category: data.category||'', desc: data.desc||'', url: fileUrl, fileId,
    fileName: data.fileName||'', uploadedAt: data.uploadedAt||new Date().toISOString(), deadline: data.deadline||'',
    audienceType: data.audienceType||'all', audienceList: JSON.stringify(data.audienceList||[]), version: 1,
    requiresSignature: data.requiresSignature ? true : false, readOnly: data.readOnly ? true : false, remindedAt: ''
  });
  const targets = resolveAudience(data.audienceType||'all', data.audienceList||[]);
  sendDocumentNotification(data.title, data.category||'', data.desc||'', fileUrl, data.deadline||'', targets, data.requiresSignature);
  // push notif na každého cílového zaměstnance
  const docIcon = data.requiresSignature ? '✍️' : (data.readOnly ? '👁' : '📄');
  targets.forEach(function(emp) { notifyUser(emp.id, 'doc', docIcon, 'Nový dokument k ' + (data.requiresSignature?'podpisu':'přečtení') + ': ' + data.title, ''); });
  return { success: true, docId, url: fileUrl, fileId };
}
function updateDocument(data) {
  const doc = readSheet(SHEETS.DOCUMENTS).items.find(d => d.id === data.id);
  if (!doc) return { error: 'Document not found' };
  const obj = {};
  ['title','category','desc','deadline','audienceType'].forEach(k => { if (data[k] !== undefined) obj[k] = data[k]; });
  if (data.audienceList !== undefined) obj.audienceList = JSON.stringify(data.audienceList);
  if (data.requiresSignature !== undefined) obj.requiresSignature = data.requiresSignature ? true : false;
  if (data.readOnly !== undefined) obj.readOnly = data.readOnly ? true : false;
  let fileUrl = doc.url, fileId = doc.fileId;
  if (data.fileData && data.fileName) {
    const saved = saveFileToDrive(data.fileData, data.fileType, data.fileName);
    fileUrl = saved.url; fileId = saved.id;
    obj.url = fileUrl; obj.fileId = fileId; obj.fileName = data.fileName;
  }
  const ver = (parseInt(doc.version) || 1) + 1;
  obj.version = ver;
  obj.uploadedAt = new Date().toISOString();
  updateRowByHeaders(SHEETS.DOCUMENTS, 'id', data.id, obj);
  deleteRowByCol(SHEETS.CONFIRMATIONS, 'docId', data.id);
  const aType = obj.audienceType || doc.audienceType;
  const aList = data.audienceList !== undefined ? data.audienceList : parseList(doc.audienceList);
  const targets = resolveAudience(aType, aList);
  const reqSig = obj.requiresSignature !== undefined ? obj.requiresSignature : (doc.requiresSignature === true || doc.requiresSignature === 'true');
  sendDocumentUpdateNotification(obj.title || doc.title, obj.category !== undefined ? obj.category : doc.category, obj.desc !== undefined ? obj.desc : doc.desc, fileUrl, obj.deadline !== undefined ? obj.deadline : doc.deadline, targets, reqSig);
  return { success: true, version: ver, url: fileUrl, fileId };
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
  const doc = readSheet(SHEETS.DOCUMENTS).items.find(d => d.id === data.docId);
  let signedUrl = '', signedFileId = '', signedFileName = '';
  if (data.fileData && data.fileName) {
    const saved = saveFileToDrive(data.fileData, data.fileType, data.fileName);
    signedUrl = saved.url; signedFileId = saved.id; signedFileName = data.fileName;
  }
  appendByHeaders(SHEETS.CONFIRMATIONS, { docId: data.docId, docTitle: doc?doc.title:data.docId, employeeId: data.employeeId, employeeName: data.employeeName, confirmedAt: data.confirmedAt||new Date().toISOString(), signedUrl, signedFileId, signedFileName });
  sendConfirmationNotification(data.employeeName, doc?doc.title:data.docId, !!signedUrl);
  return { success: true, signedUrl, signedFileId };
}
function getConfirmations() {
  const { items } = readSheet(SHEETS.CONFIRMATIONS);
  return { confirmations: items.map(r => ({ docId: r.docId, docTitle: r.docTitle, employeeId: r.employeeId, employeeName: r.employeeName, confirmedAt: r.confirmedAt, signedUrl: r.signedUrl||'', signedFileId: r.signedFileId||'', signedFileName: r.signedFileName||'' })).filter(c => c.docId) };
}

// ============================================================
// URGE (zaurgovat) — owner/admin nudges
// ============================================================
function urgeDocument(data) {
  const doc = readSheet(SHEETS.DOCUMENTS).items.find(d => d.id === data.docId);
  if (!doc) return { error: 'Document not found' };
  const targets = resolveAudience(doc.audienceType, parseList(doc.audienceList));
  const confs = getConfirmations().confirmations.filter(c => c.docId === data.docId);
  const confirmedIds = new Set(confs.map(c => c.employeeId));
  const pending = targets.filter(e => !confirmedIds.has(e.id));
  let sent = 0;
  pending.forEach(e => { if (e.email && e.email.includes('@')) { sendDocUrge(e, doc, data.message||'', data.byName||''); sent++; } });
  return { success: true, urged: sent, pending: pending.length };
}
function urgeTask(data) {
  const task = readSheet(SHEETS.TASKS).items.find(t => t.id === data.taskId);
  if (!task) return { error: 'Task not found' };
  const ids = parseList(task.assignees);
  let sent = 0;
  ids.forEach(id => { const e = empById(id); if (e && e.email && e.email.includes('@')) { sendTaskUrge(e, task, data.message||'', data.byName||''); sent++; } });
  return { success: true, urged: sent };
}
function sendDocUrge(emp, doc, message, byName) {
  try {
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 18px;">Dobrý den ${emp.name},</p>
      <div style="background:#f7eed8;border:1px solid #ecd6a3;border-radius:8px;padding:14px 18px;margin-bottom:18px;"><div style="font-size:14px;font-weight:bold;color:#a6781f;">Připomínka: čeká se na vás</div></div>
      <p style="color:#3d301f;font-size:15px;margin:0 0 16px;">Je třeba ${doc.requiresSignature ? 'podepsat a nahrát' : 'potvrdit'} dokument:</p>
      <div style="background:#f5f3ee;border-left:4px solid #13302e;padding:16px 20px;margin-bottom:20px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${doc.title}${doc.category?' ['+doc.category+']':''}</div>
        ${doc.deadline?`<div style="font-size:13px;color:#9c3a2e;margin-top:8px;font-weight:bold;">Termín: ${fmtDateCz(doc.deadline)}</div>`:''}
      </div>
      ${message?`<div style="background:#ffffff;border:1px solid #e8e5dd;border-radius:8px;padding:14px 16px;margin-bottom:20px;"><div style="font-size:12px;color:#9c7852;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">Vzkaz${byName?' od '+byName:''}</div><div style="font-size:14px;color:#3d301f;">${message}</div></div>`:''}
      ${btnHtml('Otevřít Vista Portál', PORTAL_URL)}`;
    GmailApp.sendEmail(emp.email, 'Vista Portál: Připomínka — ' + doc.title, `Připomínka: ${doc.title}. ${message||''} ${PORTAL_URL}`, { htmlBody: emailShell('Připomínka', body), name: 'Vista Resort' });
  } catch(err) { Logger.log('doc urge: ' + err.message); }
}
function sendTaskUrge(emp, task, message, byName) {
  try {
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 18px;">Dobrý den ${emp.name},</p>
      <div style="background:#f7eed8;border:1px solid #ecd6a3;border-radius:8px;padding:14px 18px;margin-bottom:18px;"><div style="font-size:14px;font-weight:bold;color:#a6781f;">Připomínka úkolu</div></div>
      <div style="background:#f5f3ee;border-left:4px solid #13302e;padding:16px 20px;margin-bottom:20px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${task.title}</div>
        <div style="font-size:13px;color:#3d301f;margin-top:6px;">Stav: <strong>${statusCz(task.status)}</strong>${task.deadline?` · Termín: <strong style="color:#9c3a2e">${fmtDateCz(task.deadline)}</strong>`:''}</div>
      </div>
      ${message?`<div style="background:#ffffff;border:1px solid #e8e5dd;border-radius:8px;padding:14px 16px;margin-bottom:20px;"><div style="font-size:12px;color:#9c7852;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">Vzkaz${byName?' od '+byName:''}</div><div style="font-size:14px;color:#3d301f;">${message}</div></div>`:''}
      ${btnHtml('Otevřít úkol', PORTAL_URL)}`;
    GmailApp.sendEmail(emp.email, 'Vista Portál: Připomínka úkolu — ' + task.title, `Připomínka: ${task.title}. ${message||''} ${PORTAL_URL}`, { htmlBody: emailShell('Připomínka úkolu', body), name: 'Vista Resort' });
  } catch(err) { Logger.log('task urge: ' + err.message); }
}

// ============================================================
// TASKS
// ============================================================
function getTasks() {
  const { items } = readSheet(SHEETS.TASKS);
  const tasks = items.map(r => ({
    id: r.id, title: r.title, desc: r.desc, segment: r.segment, priority: r.priority||'medium',
    status: r.status||'todo', assignees: parseList(r.assignees), dependsOn: parseList(r.dependsOn),
    deadline: r.deadline||'', createdBy: r.createdBy, createdByName: r.createdByName, createdAt: r.createdAt,
    completedAt: r.completedAt||'', attachments: parseAttachments(r.attachments)
  })).filter(t => t.id);
  const comments = getComments().comments;
  return { tasks, comments };
}
function parseAttachments(val) { if (!val) return []; try { return JSON.parse(val); } catch(e) { return []; } }
function addTask(data) {
  const id = data.id || 'task_' + Date.now();
  appendByHeaders(SHEETS.TASKS, {
    id, title: data.title, desc: data.desc||'', segment: data.segment||'', priority: data.priority||'medium',
    status: data.status||'todo', assignees: JSON.stringify(data.assignees||[]), dependsOn: JSON.stringify(data.dependsOn||[]),
    deadline: data.deadline||'', createdBy: data.createdBy||'', createdByName: data.createdByName||'',
    createdAt: new Date().toISOString(), completedAt: '', attachments: JSON.stringify(data.attachments||[]), remindedAt: ''
  });
  notifyTaskAssigned(id, data.title, data.desc||'', data.priority||'medium', data.deadline||'', data.assignees||[], data.createdByName||'');
  // push notif na každého řešitele
  (data.assignees||[]).forEach(function(uid) {
    notifyUser(uid, 'task', '📌', 'Byl vám přiřazen úkol: ' + data.title, '');
  });
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
  const obj = { status: data.status, statusChangedAt: new Date().toISOString() };
  if (data.status === 'done') obj.completedAt = new Date().toISOString();
  updateRowByHeaders(SHEETS.TASKS, 'id', data.id, obj);
  const task = readSheet(SHEETS.TASKS).items.find(t => t.id === data.id);
  if (task) notifyStatusChange(task, data.status, data.changedByName||'');
  return { success: true };
}
function deleteTask(id) { deleteRowByCol(SHEETS.TASKS, 'id', id); deleteRowByCol(SHEETS.COMMENTS, 'taskId', id); return { success: true }; }
function addTaskAttachment(data) {
  const task = readSheet(SHEETS.TASKS).items.find(t => t.id === data.taskId);
  if (!task) return { error: 'Task not found' };
  let url = '';
  if (data.fileData && data.fileName) { url = saveFileToDrive(data.fileData, data.fileType, data.fileName).url; }
  const atts = parseAttachments(task.attachments);
  atts.push({ name: data.fileName, url, uploadedAt: new Date().toISOString() });
  updateRowByHeaders(SHEETS.TASKS, 'id', data.taskId, { attachments: JSON.stringify(atts) });
  return { success: true, url, attachments: atts };
}

// ============================================================
// COMMENTS (tasks + invoices)
// ============================================================
function addComment(data) {
  const id = 'cmt_' + Date.now();
  const entityType = data.entityType || 'task';
  appendByHeaders(SHEETS.COMMENTS, { id, taskId: data.taskId, authorId: data.authorId||'', authorName: data.authorName||'', text: data.text||'', createdAt: new Date().toISOString(), entityType });
  if (entityType === 'task') {
    const task = readSheet(SHEETS.TASKS).items.find(t => t.id === data.taskId);
    if (task) notifyNewComment(task, data.authorName||'', data.text||'');
  } else if (entityType === 'invoice') {
    const inv = readSheet(SHEETS.INVOICES).items.find(i => i.id === data.taskId);
    if (inv) notifyInvoiceComment(inv, data.authorName||'', data.text||'');
  }
  return { success: true, id };
}
function getComments() {
  const { items } = readSheet(SHEETS.COMMENTS);
  return { comments: items.map(r => ({ id: r.id, taskId: r.taskId, authorId: r.authorId, authorName: r.authorName, text: r.text, createdAt: r.createdAt, entityType: r.entityType||'task' })).filter(c => c.id) };
}

// ============================================================
// INVOICES (Faktury)
// ============================================================
function getInvoices() {
  const { items } = readSheet(SHEETS.INVOICES);
  const invoices = items.map(r => ({
    id: r.id, title: r.title, supplier: r.supplier||'', amount: r.amount||'', currency: r.currency||'CZK',
    url: r.url||'', fileId: r.fileId||'', fileName: r.fileName||'', dueDate: r.dueDate||'',
    uploadedBy: r.uploadedBy||'', uploadedByName: r.uploadedByName||'', uploadedAt: r.uploadedAt||'',
    status: r.status||'pending', decidedBy: r.decidedBy||'', decidedByName: r.decidedByName||'', decidedAt: r.decidedAt||''
  })).filter(i => i.id);
  return { invoices };
}
function addInvoice(data) {
  let url = '', fileId = '';
  const id = data.id || 'inv_' + Date.now();
  if (data.fileData && data.fileName) {
    const saved = saveFileToDrive(data.fileData, data.fileType, data.fileName);
    url = saved.url; fileId = saved.id;
  }
  appendByHeaders(SHEETS.INVOICES, {
    id, title: data.title||data.fileName||'Faktura', supplier: data.supplier||'', amount: data.amount||'', currency: data.currency||'CZK',
    url, fileId, fileName: data.fileName||'', dueDate: data.dueDate||'',
    uploadedBy: data.uploadedBy||'', uploadedByName: data.uploadedByName||'', uploadedAt: new Date().toISOString(),
    status: 'pending', decidedBy: '', decidedByName: '', decidedAt: ''
  });
  notifyInvoiceUploaded({ id, title: data.title||data.fileName, supplier: data.supplier||'', amount: data.amount||'', currency: data.currency||'CZK', dueDate: data.dueDate||'', uploadedByName: data.uploadedByName||'' });
  notifyManagement('invoice', '🧾', 'Nová faktura k schválení: ' + (data.title||data.fileName||'faktura'), '');
  return { success: true, id, url, fileId };
}
function decideInvoice(data) {
  updateRowByHeaders(SHEETS.INVOICES, 'id', data.id, { status: data.status, decidedBy: data.decidedBy||'', decidedByName: data.decidedByName||'', decidedAt: new Date().toISOString() });
  const inv = readSheet(SHEETS.INVOICES).items.find(i => i.id === data.id);
  if (inv) {
    notifyInvoiceDecision(inv, data.status, data.decidedByName||'');
    const icon = data.status === 'approved' ? '✅' : '❌';
    const label = data.status === 'approved' ? 'schválena k proplacení' : 'zamítnuta';
    notifyManagement('invoice', icon, 'Faktura ' + label + ': ' + (inv.title||''), '');
  }
  return { success: true };
}
function deleteInvoice(id) { deleteRowByCol(SHEETS.INVOICES, 'id', id); deleteRowByCol(SHEETS.COMMENTS, 'taskId', id); return { success: true }; }

// ============================================================
// WORKLOG (Deník)
// ============================================================
function getWorklog() {
  const { items } = readSheet(SHEETS.WORKLOG);
  return { worklog: items.map(r => ({ id: r.id, employeeId: r.employeeId, employeeName: r.employeeName, segment: r.segment||'', date: r.date||'', text: r.text||'', createdAt: r.createdAt })).filter(w => w.id) };
}
function addWorklog(data) {
  const id = 'wl_' + Date.now();
  appendByHeaders(SHEETS.WORKLOG, { id, employeeId: data.employeeId||'', employeeName: data.employeeName||'', segment: data.segment||'', date: data.date||new Date().toISOString().slice(0,10), text: data.text||'', createdAt: new Date().toISOString() });
  return { success: true, id };
}
function updateWorklog(data) {
  updateRowByHeaders(SHEETS.WORKLOG, 'id', data.id, { text: data.text || '' });
  return { success: true };
}

// ============================================================
// TRŽBY — uložení reportu na server (jen vedení k nim má přístup ve frontendu)
// ============================================================
function saveTrzby(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.TRZBY);
  if (!sheet) { sheet = ss.insertSheet(SHEETS.TRZBY); sheet.appendRow(['key','json','uploadedBy','uploadedAt']); }
  const now = new Date().toISOString();
  const chunkSize = 45000;
  // helper: remove existing chunks for a prefix, then write new
  const writeChunks = (prefix, json) => {
    if (json === undefined || json === null) return;
    // read all, filter out this prefix, rewrite
    const last = sheet.getLastRow();
    let kept = [];
    if (last > 1) {
      const vals = sheet.getRange(2,1,last-1,4).getValues();
      kept = vals.filter(r => !String(r[0]).startsWith(prefix + ':'));
    }
    const chunks = [];
    for (let i = 0; i < json.length; i += chunkSize) chunks.push(json.slice(i, i+chunkSize));
    chunks.forEach((c,i) => kept.push([prefix + ':' + i, c, data.uploadedBy||'', now]));
    if (last > 1) sheet.deleteRows(2, last-1);
    if (kept.length) sheet.getRange(2,1,kept.length,4).setValues(kept);
  };
  if (data.json !== undefined)  writeChunks('vynosy', data.json);
  if (data.json2 !== undefined) writeChunks('prijmy', data.json2);
  return { success: true, uploadedAt: now };
}
function getTrzby() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TRZBY);
  if (!sheet || sheet.getLastRow() < 2) return { json: '', json2: '', uploadedAt: '', uploadedBy: '' };
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues();
  const collect = prefix => {
    const part = rows.filter(r => String(r[0]).startsWith(prefix + ':'));
    part.sort((a,b) => parseInt(String(a[0]).split(':')[1]) - parseInt(String(b[0]).split(':')[1]));
    return part.map(r => r[1]).join('');
  };
  // legacy support: old chunks named 'chunk0' => treat as vynosy
  const legacy = rows.filter(r => String(r[0]).startsWith('chunk'));
  let json = collect('vynosy');
  if (!json && legacy.length) { legacy.sort((a,b)=>parseInt(String(a[0]).replace('chunk',''))-parseInt(String(b[0]).replace('chunk',''))); json = legacy.map(r=>r[1]).join(''); }
  const uploadedAt = rows[0] ? rows[0][3] : '';
  return { json, json2: collect('prijmy'), uploadedAt, uploadedBy: rows[0] ? rows[0][2] : '' };
}
function deleteWorklog(id) { deleteRowByCol(SHEETS.WORKLOG, 'id', id); return { success: true }; }

// ============================================================
// REPORTY
// ============================================================
function getReports() {
  const { items } = readSheet(SHEETS.REPORTS);
  return { reports: items.map(r => ({ id:r.id, title:r.title, category:r.category||'Ostatní', desc:r.desc||'', url:r.url||'', fileId:r.fileId||'', fileName:r.fileName||'', uploadedBy:r.uploadedBy||'', uploadedByName:r.uploadedByName||'', uploadedAt:r.uploadedAt||'' })).filter(r=>r.id).sort((a,b)=>new Date(b.uploadedAt)-new Date(a.uploadedAt)) };
}
function uploadReport(data) {
  let url = '', fileId = '';
  const id = 'rep_' + Date.now();
  if (data.fileData && data.fileName) {
    const saved = saveFileToDrive(data.fileData, data.fileType, data.fileName);
    url = saved.url; fileId = saved.id;
  }
  appendByHeaders(SHEETS.REPORTS, { id, title:data.title||'', category:data.category||'Ostatní', desc:data.desc||'', url, fileId, fileName:data.fileName||'', uploadedBy:data.uploadedBy||'', uploadedByName:data.uploadedByName||'', uploadedAt:data.uploadedAt||new Date().toISOString() });
  return { success:true, id, url, fileId };
}
function deleteReport(id) { deleteRowByCol(SHEETS.REPORTS, 'id', id); return { success:true }; }
function getAll() {
  // Batch read — otevře spreadsheet JEDNOU a načte všechny listy najednou
  // Namísto 9 samostatných I/O volání = ~4–8× rychlejší getAll
  const cache = batchReadAll();

  // employees
  const empItems = readSheetFromCache(cache, SHEETS.EMPLOYEES).items;
  const employees = empItems.map(function(r) {
    return { id:r.id, name:r.name||'', role:r.role||'', dept:r.dept||'', accessLevel:r.accessLevel||'standard', pin:r.pin||'', email:r.email||'', createdAt:r.createdAt||'' };
  }).filter(function(r){ return r.id; });

  // documents
  const docItems = readSheetFromCache(cache, SHEETS.DOCUMENTS).items;
  const documents = docItems.map(function(r) {
    return { id:r.id, title:r.title||'', category:r.category||'', desc:r.desc||'', url:r.url||'', fileId:r.fileId||'',
      fileName:r.fileName||'', uploadedAt:r.uploadedAt||'', deadline:r.deadline||'',
      audienceType:r.audienceType||'all', audienceList: tryParseJSON(r.audienceList, []),
      version:parseInt(r.version)||1,
      requiresSignature: isTruthy(r.requiresSignature),
      readOnly: isTruthy(r.readOnly) };
  }).filter(function(r){ return r.id; });

  // confirmations
  const confItems = readSheetFromCache(cache, SHEETS.CONFIRMATIONS).items;
  const confirmations = confItems.map(function(r) {
    return { id:r.id, docId:r.docId||'', employeeId:r.employeeId||'', employeeName:r.employeeName||'',
      confirmedAt:r.confirmedAt||'', signedUrl:r.signedUrl||'', signedFileId:r.signedFileId||'', signedFileName:r.signedFileName||'' };
  }).filter(function(r){ return r.id; });

  // tasks + comments
  const taskItems = readSheetFromCache(cache, SHEETS.TASKS).items;
  const tasks = taskItems.map(function(r) {
    return { id:r.id, title:r.title||'', desc:r.desc||'', segment:r.segment||'', priority:r.priority||'medium',
      status:r.status||'todo', assignees:tryParseJSON(r.assignees,[]), dependsOn:tryParseJSON(r.dependsOn,[]),
      deadline:r.deadline||'', createdBy:r.createdBy||'', createdByName:r.createdByName||'',
      createdAt:r.createdAt||'', completedAt:r.completedAt||'', statusChangedAt:r.statusChangedAt||'',
      attachments:tryParseJSON(r.attachments,[]), remindedAt:r.remindedAt||'' };
  }).filter(function(r){ return r.id; });

  const cmtItems = readSheetFromCache(cache, SHEETS.COMMENTS).items;
  const comments = cmtItems.map(function(r) {
    return { id:r.id, taskId:r.taskId||r.entityId||'', entityType:r.entityType||'task',
      authorId:r.authorId||'', authorName:r.authorName||'', text:r.text||'', createdAt:r.createdAt||'' };
  }).filter(function(r){ return r.id; });

  // invoices
  const invItems = readSheetFromCache(cache, SHEETS.INVOICES).items;
  const invoices = invItems.map(function(r) {
    return { id:r.id, title:r.title||'', supplier:r.supplier||'', amount:r.amount||'', currency:r.currency||'CZK',
      dueDate:r.dueDate||'', url:r.url||'', fileId:r.fileId||'', fileName:r.fileName||'',
      uploadedBy:r.uploadedBy||'', uploadedByName:r.uploadedByName||'', uploadedAt:r.uploadedAt||'',
      status:r.status||'pending', decidedBy:r.decidedBy||'', decidedByName:r.decidedByName||'',
      decidedAt:r.decidedAt||'', comment:r.comment||'' };
  }).filter(function(r){ return r.id; });

  // worklog
  const wlItems = readSheetFromCache(cache, SHEETS.WORKLOG).items;
  const worklog = wlItems.map(function(r) {
    return { id:r.id, employeeId:r.employeeId||'', employeeName:r.employeeName||'', segment:r.segment||'',
      date:r.date||'', text:r.text||'', createdAt:r.createdAt||'' };
  }).filter(function(r){ return r.id; });

  // reports
  const repItems = readSheetFromCache(cache, SHEETS.REPORTS).items;
  const reports = repItems.map(function(r) {
    return { id:r.id, title:r.title||'', category:r.category||'Ostatní', desc:r.desc||'', url:r.url||'',
      fileId:r.fileId||'', fileName:r.fileName||'', uploadedBy:r.uploadedBy||'',
      uploadedByName:r.uploadedByName||'', uploadedAt:r.uploadedAt||'' };
  }).filter(function(r){ return r.id; }).sort(function(a,b){ return new Date(b.uploadedAt)-new Date(a.uploadedAt); });

  // task templates
  const tmplItems = readSheetFromCache(cache, SHEETS.TEMPLATES).items;
  const taskTemplates = tmplItems.map(function(r) {
    return { id:r.id, title:r.title||'', desc:r.desc||'', segment:r.segment||'', priority:r.priority||'medium',
      assignees:tryParseJSON(r.assignees,[]), freq:r.freq||'weekly', days:tryParseJSON(r.days,[]).map(Number),
      monthDay:r.monthDay||'', deadlineOffset:parseInt(r.deadlineOffset)||0,
      active: isTruthy(r.active), createdAt:r.createdAt||'', lastRun:r.lastRun||'' };
  }).filter(function(r){ return r.id; });

  return { employees, documents, confirmations, tasks, comments, invoices, worklog, reports, taskTemplates };
}

// getEssential — rychlý subset pro první zobrazení Přehledu
// Vrátí jen faktury, úkoly a zaměstnance (bez dokumentů, worklog, reportů atd.)
// Typicky 2–3× rychlejší než getAll
function getEssential() {
  const cache = batchReadAll();
  const empItems = readSheetFromCache(cache, SHEETS.EMPLOYEES).items;
  const employees = empItems.filter(function(r){ return r.id; }).map(function(r){
    return { id:r.id, name:r.name||'', role:r.role||'', dept:r.dept||'', accessLevel:r.accessLevel||'standard', pin:r.pin||'', email:r.email||'', createdAt:r.createdAt||'' };
  });
  const taskItems = readSheetFromCache(cache, SHEETS.TASKS).items;
  const tasks = taskItems.filter(function(r){ return r.id; }).map(function(r){
    return { id:r.id, title:r.title||'', desc:r.desc||'', segment:r.segment||'', priority:r.priority||'medium',
      status:r.status||'todo', assignees:tryParseJSON(r.assignees,[]), dependsOn:tryParseJSON(r.dependsOn,[]),
      deadline:r.deadline||'', createdBy:r.createdBy||'', createdByName:r.createdByName||'',
      createdAt:r.createdAt||'', completedAt:r.completedAt||'', statusChangedAt:r.statusChangedAt||'',
      attachments:tryParseJSON(r.attachments,[]), remindedAt:r.remindedAt||'' };
  });
  const invItems = readSheetFromCache(cache, SHEETS.INVOICES).items;
  const invoices = invItems.filter(function(r){ return r.id; }).map(function(r){
    return { id:r.id, title:r.title||'', supplier:r.supplier||'', amount:r.amount||'', currency:r.currency||'CZK',
      dueDate:r.dueDate||'', url:r.url||'', fileId:r.fileId||'', fileName:r.fileName||'',
      uploadedBy:r.uploadedBy||'', uploadedByName:r.uploadedByName||'', uploadedAt:r.uploadedAt||'',
      status:r.status||'pending', decidedBy:r.decidedBy||'', decidedByName:r.decidedByName||'',
      decidedAt:r.decidedAt||'', comment:r.comment||'' };
  });
  const cmtItems = readSheetFromCache(cache, SHEETS.COMMENTS).items;
  const comments = cmtItems.filter(function(r){ return r.id; }).map(function(r){
    return { id:r.id, taskId:r.taskId||r.entityId||'', entityType:r.entityType||'task',
      authorId:r.authorId||'', authorName:r.authorName||'', text:r.text||'', createdAt:r.createdAt||'' };
  });
  return { employees, tasks, comments, invoices,
    documents:[], confirmations:[], worklog:[], reports:[], taskTemplates:[] };
}

function isTruthy(v) { return v===true||v==='true'||v==='ANO'||v===1||v==='1'; }
function tryParseJSON(s, fallback) { try { return s ? JSON.parse(s) : fallback; } catch(e) { return fallback; } }

// ============================================================
// DRIVE
// ============================================================
function saveFileToDrive(fileData, fileType, fileName) {
  const folder = getDriveFolder();
  const blob = Utilities.newBlob(Utilities.base64Decode(fileData), fileType || 'application/octet-stream', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { id: file.getId(), url: file.getUrl() };
}
function getDriveFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

// ============================================================
// EMAIL — shell
// ============================================================
function emailShell(headerSub, bodyHtml, accent) {
  accent = accent || '#dfc196';
  return `
    <div style="font-family:Georgia,'Times New Roman',serif;max-width:560px;margin:0 auto;background:#ffffff;">
      <div style="background:#13302e;padding:28px 36px;">
        <div style="color:${accent};font-size:24px;font-weight:bold;letter-spacing:1px;">VISTA RESORT</div>
        <div style="color:rgba(255,255,255,0.6);margin:6px 0 0;font-size:12px;letter-spacing:2px;text-transform:uppercase;font-family:Arial,sans-serif;">${headerSub}</div>
      </div>
      <div style="padding:36px;border:1px solid #e8e5dd;border-top:none;font-family:Arial,sans-serif;">${bodyHtml}</div>
      <div style="background:#f5f3ee;padding:18px 36px;border:1px solid #e8e5dd;border-top:none;"><p style="color:#9c7852;font-size:12px;margin:0;font-family:Arial,sans-serif;">Vista Resort &amp; Club · Interní portál · Odesláno automaticky.</p></div>
    </div>`;
}
function btnHtml(text, url) {
  return `<a href="${url}" style="display:inline-block;background:#13302e;color:#dfc196;padding:13px 28px;border-radius:4px;text-decoration:none;font-size:15px;font-weight:bold;font-family:Arial,sans-serif;">${text}</a>`;
}

// ============================================================
// EMAIL — documents
// ============================================================
function sendDocumentNotification(title, category, desc, fileUrl, deadline, targets, requiresSignature) {
  try {
    const emails = (targets || getEmployees().employees).map(e => e.email).filter(em => em && em.includes('@'));
    if (emails.length === 0) return;
    const catText = category ? ` [${category}]` : '';
    const deadlineText = deadline ? fmtDateCz(deadline) : '';
    const action = requiresSignature ? 'podepsat a nahrát zpět' : 'potvrdit';
    const cta = requiresSignature ? 'Nahrát podepsaný dokument' : 'Potvrdit přečtení';
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">Dobrý den,</p>
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">byl nahrán nový dokument, který je třeba ${action}:</p>
      <div style="background:#f5f3ee;border-left:4px solid #13302e;padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${title}${catText}</div>
        ${desc ? `<div style="font-size:13px;color:#9c7852;margin-top:4px;">${desc}</div>` : ''}
        ${requiresSignature ? `<div style="font-size:13px;color:#a6781f;margin-top:8px;font-weight:bold;">Vyžaduje podpis</div>` : ''}
        ${deadlineText ? `<div style="font-size:13px;color:#9c3a2e;margin-top:8px;font-weight:bold;">${requiresSignature?'Nahrajte do':'Potvrďte do'}: ${deadlineText}</div>` : ''}
      </div>
      <p style="color:#9c7852;font-size:14px;margin:0 0 24px;">Přihlaste se do portálu a klikněte na „${cta}".</p>
      ${btnHtml('Otevřít Vista Portál', PORTAL_URL)}`;
    const html = emailShell('Interní dokumentový systém', body);
    emails.forEach(em => GmailApp.sendEmail(em, 'Vista Portál: Nový dokument k ' + (requiresSignature?'podpisu':'potvrzení'), `${title}${catText} — ${PORTAL_URL}`, { htmlBody: html, name: 'Vista Resort' }));
  } catch(err) { Logger.log('doc email: ' + err.message); }
}
function sendDocumentUpdateNotification(title, category, desc, fileUrl, deadline, targets, requiresSignature) {
  try {
    const emails = (targets || getEmployees().employees).map(e => e.email).filter(em => em && em.includes('@'));
    if (emails.length === 0) return;
    const catText = category ? ` [${category}]` : '';
    const deadlineText = deadline ? fmtDateCz(deadline) : '';
    const action = requiresSignature ? 'znovu podepsat a nahrát' : 'znovu potvrdit';
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">Dobrý den,</p>
      <div style="background:#f7eed8;border:1px solid #ecd6a3;border-radius:8px;padding:12px 16px;margin-bottom:20px;"><div style="font-size:14px;font-weight:bold;color:#a6781f;">Dokument byl aktualizován</div><div style="font-size:13px;color:#876012;margin-top:3px;">Je třeba ho ${action}.</div></div>
      <div style="background:#f5f3ee;border-left:4px solid #13302e;padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${title}${catText}</div>
        ${desc ? `<div style="font-size:13px;color:#9c7852;margin-top:4px;">${desc}</div>` : ''}
        ${requiresSignature ? `<div style="font-size:13px;color:#a6781f;margin-top:8px;font-weight:bold;">Vyžaduje podpis</div>` : ''}
        ${deadlineText ? `<div style="font-size:13px;color:#9c3a2e;margin-top:8px;font-weight:bold;">Termín: ${deadlineText}</div>` : ''}
      </div>
      ${btnHtml('Otevřít Vista Portál', PORTAL_URL)}`;
    emails.forEach(em => GmailApp.sendEmail(em, 'Vista Portál: Aktualizovaný dokument k ' + (requiresSignature?'podpisu':'potvrzení'), `${title}${catText} byl aktualizován — ${PORTAL_URL}`, { htmlBody: emailShell('Aktualizace dokumentu', body), name: 'Vista Resort' }));
  } catch(err) { Logger.log('doc update email: ' + err.message); }
}
function sendConfirmationNotification(employeeName, docTitle, signed) {
  try {
    if (!ADMIN_EMAIL && !GM_EMAIL) return;
    const now = new Date().toLocaleString('cs-CZ', { timeZone: 'Europe/Prague' });
    const label = signed ? 'Dokument podepsán a nahrán' : 'Dokument potvrzen';
    const body = `
      <div style="background:#eef5ee;border:1px solid #b8d4b0;padding:16px 20px;margin-bottom:20px;"><div style="font-size:15px;font-weight:bold;color:#586845;">${label}</div></div>
      <table style="font-size:14px;color:#3d301f;width:100%">
        <tr><td style="padding:6px 0;color:#9c7852;width:120px;">Zaměstnanec</td><td style="padding:6px 0;font-weight:bold;">${employeeName}</td></tr>
        <tr><td style="padding:6px 0;color:#9c7852;">Dokument</td><td style="padding:6px 0;font-weight:bold;">${docTitle}</td></tr>
        <tr><td style="padding:6px 0;color:#9c7852;">Datum</td><td style="padding:6px 0;">${now}</td></tr>
      </table>`;
    sendToManagement(`Vista Portál: ${employeeName} ${signed?'podepsal/a':'potvrdil/a'} dokument`, `${employeeName}: ${docTitle} (${now})`, emailShell(label, body));
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
function statusCz(s) { return s === 'todo' ? 'Zadáno' : s === 'inprogress' ? 'Probíhá' : s === 'review' ? 'Ke kontrole' : s === 'done' ? 'Hotovo' : s; }

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
      </div>${btnHtml('Otevřít úkol', PORTAL_URL)}`;
    emails.forEach(em => GmailApp.sendEmail(em, 'Vista Portál: Nový úkol — ' + title, `Nový úkol: ${title}. ${PORTAL_URL}`, { htmlBody: emailShell('Nový úkol', body), name: 'Vista Resort' }));
  } catch(err) { Logger.log('task assign: ' + err.message); }
}
function notifyStatusChange(task, newStatus, changedByName) {
  try {
    const ids = parseList(task.assignees).slice();
    if (task.createdBy) ids.push(task.createdBy);
    const emails = [...new Set(emailsForIds(ids))];
    if (emails.length === 0) return;
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">Změna stavu úkolu:</p>
      <div style="background:#f5f3ee;border-left:4px solid #13302e;padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${task.title}</div>
        <div style="font-size:14px;color:#3d301f;margin-top:8px;">Nový stav: <strong>${statusCz(newStatus)}</strong>${changedByName ? ` · změnil/a ${changedByName}` : ''}</div>
      </div>${btnHtml('Otevřít úkol', PORTAL_URL)}`;
    emails.forEach(em => GmailApp.sendEmail(em, `Vista Portál: ${task.title} → ${statusCz(newStatus)}`, `${task.title} → ${statusCz(newStatus)}. ${PORTAL_URL}`, { htmlBody: emailShell('Změna stavu úkolu', body), name: 'Vista Resort' }));
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
      </div>${btnHtml('Odpovědět v portálu', PORTAL_URL)}`;
    emails.forEach(em => GmailApp.sendEmail(em, `Vista Portál: Komentář — ${task.title}`, `${authorName}: ${text}`, { htmlBody: emailShell('Nový komentář', body), name: 'Vista Resort' }));
  } catch(err) { Logger.log('comment email: ' + err.message); }
}

// ============================================================
// EMAIL — invoices
// ============================================================
function fmtAmount(amount, currency) {
  if (amount === '' || amount === null || amount === undefined) return '';
  const n = Number(amount);
  if (isNaN(n)) return String(amount) + (currency ? ' ' + currency : '');
  return n.toLocaleString('cs-CZ') + ' ' + (currency || 'CZK');
}
function notifyInvoiceUploaded(inv) {
  try {
    if (!OWNER_EMAIL) return;
    const amt = fmtAmount(inv.amount, inv.currency);
    const due = inv.dueDate ? fmtDateCz(inv.dueDate) : '';
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">Dobrý den,</p>
      <p style="color:#3d301f;font-size:15px;margin:0 0 20px;">byla nahrána nová faktura ke schválení:</p>
      <div style="background:#f5f3ee;border-left:4px solid #13302e;padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${inv.title}</div>
        ${inv.supplier ? `<div style="font-size:13px;color:#9c7852;margin-top:4px;">Dodavatel: ${inv.supplier}</div>` : ''}
        ${amt ? `<div style="font-size:14px;color:#3d301f;margin-top:8px;font-weight:bold;">Částka: ${amt}</div>` : ''}
        ${due ? `<div style="font-size:13px;color:#9c3a2e;margin-top:6px;">Splatnost: ${due}</div>` : ''}
        ${inv.uploadedByName ? `<div style="font-size:12px;color:#9c7852;margin-top:6px;">Nahrál: ${inv.uploadedByName}</div>` : ''}
      </div>
      <p style="color:#9c7852;font-size:14px;margin:0 0 24px;">Fakturu si prohlédnete a schválíte přímo v portálu.</p>
      ${btnHtml('Zobrazit a schválit', PORTAL_URL)}`;
    GmailApp.sendEmail(OWNER_EMAIL, 'Vista Portál: Nová faktura ke schválení — ' + inv.title, `Nová faktura: ${inv.title} ${amt}. Schvalte v portálu: ${PORTAL_URL}`, { htmlBody: emailShell('Faktura ke schválení', body), name: 'Vista Resort' });
  } catch(err) { Logger.log('invoice upload email: ' + err.message); }
}
function notifyInvoiceDecision(inv, status, decidedByName) {
  try {
    const emps = getEmployees().employees;
    const uploader = emps.find(e => e.id === inv.uploadedBy);
    const to = uploader && uploader.email && uploader.email.includes('@') ? uploader.email : ADMIN_EMAIL;
    if (!to) return;
    const approved = status === 'approved';
    const label = approved ? 'schválena' : 'zamítnuta';
    const color = approved ? '#586845' : '#9c3a2e';
    const body = `
      <div style="background:${approved?'#eef5ee':'#f6ece9'};border:1px solid ${approved?'#b8d4b0':'#e3c4bd'};padding:16px 20px;margin-bottom:20px;">
        <div style="font-size:15px;font-weight:bold;color:${color};">Faktura ${label}</div>
      </div>
      <table style="font-size:14px;color:#3d301f;width:100%">
        <tr><td style="padding:6px 0;color:#9c7852;width:120px;">Faktura</td><td style="padding:6px 0;font-weight:bold;">${inv.title}</td></tr>
        ${inv.supplier?`<tr><td style="padding:6px 0;color:#9c7852;">Dodavatel</td><td style="padding:6px 0;">${inv.supplier}</td></tr>`:''}
        ${fmtAmount(inv.amount,inv.currency)?`<tr><td style="padding:6px 0;color:#9c7852;">Částka</td><td style="padding:6px 0;font-weight:bold;">${fmtAmount(inv.amount,inv.currency)}</td></tr>`:''}
        <tr><td style="padding:6px 0;color:#9c7852;">Rozhodl/a</td><td style="padding:6px 0;">${decidedByName||'—'}</td></tr>
      </table>`;
    GmailApp.sendEmail(to, `Vista Portál: Faktura ${label} — ${inv.title}`, `Faktura ${inv.title} byla ${label}.`, { htmlBody: emailShell('Rozhodnutí o faktuře', body), name: 'Vista Resort' });
  } catch(err) { Logger.log('invoice decision email: ' + err.message); }
}
function notifyInvoiceComment(inv, authorName, text) {
  try {
    const emps = getEmployees().employees;
    const uploader = emps.find(e => e.id === inv.uploadedBy);
    const recipients = new Set();
    if (OWNER_EMAIL) recipients.add(OWNER_EMAIL);
    if (uploader && uploader.email && uploader.email.includes('@')) recipients.add(uploader.email);
    if (recipients.size === 0) return;
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 16px;">Nový komentář u faktury:</p>
      <div style="background:#f5f3ee;border-left:4px solid #13302e;padding:16px 20px;margin-bottom:20px;">
        <div style="font-size:15px;font-weight:bold;color:#13302e;">${inv.title}</div>
        <div style="font-size:14px;color:#3d301f;margin-top:10px;padding:10px 14px;background:#ffffff;border-radius:4px;"><strong>${authorName}:</strong> ${text}</div>
      </div>${btnHtml('Otevřít v portálu', PORTAL_URL)}`;
    [...recipients].forEach(em => GmailApp.sendEmail(em, `Vista Portál: Komentář k faktuře — ${inv.title}`, `${authorName}: ${text}`, { htmlBody: emailShell('Komentář k faktuře', body), name: 'Vista Resort' }));
  } catch(err) { Logger.log('invoice comment email: ' + err.message); }
}

// ============================================================
// DAILY CHECK — deadlines (docs + tasks)
// ============================================================
function checkDeadlines() { checkDocumentDeadlines(); checkTaskDeadlines(); }

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
    if (!afterDeadline && !dayBefore) return;
    notConfirmed.forEach(emp => { if (emp.email && emp.email.includes('@')) sendDocReminder(emp, doc, afterDeadline); });
    summary.push({ kind:'Dokument', title:doc.title, deadline:fmtDateCz(doc.deadline), afterDeadline, names:notConfirmed.map(e=>e.name+(e.dept?' ('+e.dept+')':'')) });
  });
  if (summary.length) sendAdminSummary(summary, 'Nepotvrzené dokumenty');
}
function checkTaskDeadlines() {
  const { tasks } = getTasks();
  const now = new Date();
  let summary = [];
  const emps = getEmployees().employees;
  tasks.forEach(task => {
    if (!task.deadline || task.status === 'done') return;
    const deadline = new Date(task.deadline);
    const afterDeadline = now > deadline;
    const dayBefore = !afterDeadline && (deadline - now) <= 24*60*60*1000;
    if (!afterDeadline && !dayBefore) return;
    task.assignees.forEach(id => { const e = emps.find(x => x.id === id); if (e && e.email && e.email.includes('@')) sendTaskReminder(e, task, afterDeadline); });
    const names = task.assignees.map(id => { const e = emps.find(x => x.id === id); return e ? e.name : id; });
    summary.push({ kind:'Úkol', title:task.title, deadline:fmtDateCz(task.deadline), afterDeadline, status:statusCz(task.status), names });
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
      <div style="background:#f5f3ee;border-left:4px solid ${afterDeadline?'#9c3a2e':'#13302e'};padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${doc.title}${doc.category?' ['+doc.category+']':''}</div>
        ${deadlineText?`<div style="font-size:13px;color:#9c3a2e;margin-top:8px;font-weight:bold;">Termín: ${deadlineText}</div>`:''}
      </div>${btnHtml('Potvrdit nyní', PORTAL_URL)}`;
    GmailApp.sendEmail(emp.email, subject, `${doc.title} — ${PORTAL_URL}`, { htmlBody: emailShell(afterDeadline?'Prošlý termín':'Připomínka', body), name: 'Vista Resort' });
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
      <div style="background:#f5f3ee;border-left:4px solid ${afterDeadline?'#9c3a2e':'#13302e'};padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:16px;font-weight:bold;color:#13302e;">${task.title}</div>
        <div style="font-size:13px;color:#3d301f;margin-top:6px;">Stav: <strong>${statusCz(task.status)}</strong>${deadlineText?` · Termín: <strong style="color:#9c3a2e">${deadlineText}</strong>`:''}</div>
      </div>${btnHtml('Otevřít úkol', PORTAL_URL)}`;
    GmailApp.sendEmail(emp.email, subject, `${task.title} — ${PORTAL_URL}`, { htmlBody: emailShell(afterDeadline?'Úkol po termínu':'Blížící se termín', body), name: 'Vista Resort' });
  } catch(err) { Logger.log('task reminder: ' + err.message); }
}
function sendAdminSummary(summary, heading) {
  try {
    if (!ADMIN_EMAIL && !GM_EMAIL) return;
    let rows = '';
    summary.forEach(s => {
      const color = s.afterDeadline ? '#9c3a2e' : '#9c7852';
      const status = s.afterDeadline ? 'PO TERMÍNU' : 'Blíží se termín';
      rows += `<div style="margin-bottom:16px;padding:16px 20px;background:#f5f3ee;border-left:4px solid ${color};">
        <div style="font-size:15px;font-weight:bold;color:#13302e;">${s.kind}: ${s.title}</div>
        <div style="font-size:12px;color:${color};font-weight:bold;margin:4px 0;">${status} · termín ${s.deadline}${s.status?' · stav '+s.status:''}</div>
        <div style="font-size:13px;color:#9c7852;margin-top:6px;">Dotčení (${s.names.length}): ${s.names.join(', ')}</div></div>`;
    });
    const body = `<p style="color:#3d301f;font-size:15px;margin:0 0 20px;">${heading}:</p>${rows}${btnHtml('Otevřít portál', PORTAL_URL)}`;
    sendToManagement('Vista Portál: Denní přehled — ' + heading, heading, emailShell('Denní přehled', body));
  } catch(err) { Logger.log('admin summary: ' + err.message); }
}

// ============================================================
// BI-WEEKLY REPORT (pro majitele)
// ============================================================
function generateReportData(days) {
  days = days || 14;
  const since = new Date(Date.now() - days*24*60*60*1000);
  const { tasks } = getTasks();
  const worklog = getWorklog().worklog;
  const emps = getEmployees().employees;

  // segment -> { doneTasks:[], logs:{ empName: [texts] } }
  const segments = {};
  function seg(name) { name = name || 'Bez segmentu'; if (!segments[name]) segments[name] = { doneTasks: [], logs: {} }; return segments[name]; }

  tasks.forEach(t => {
    if (t.status === 'done' && t.completedAt && new Date(t.completedAt) >= since) {
      const names = (t.assignees||[]).map(id => { const e = emps.find(x=>x.id===id); return e?e.name:''; }).filter(Boolean);
      seg(t.segment).doneTasks.push({ title: t.title, who: names.join(', ') });
    }
  });
  worklog.forEach(w => {
    const d = w.date ? new Date(w.date) : (w.createdAt ? new Date(w.createdAt) : null);
    if (d && d >= since) {
      const s = seg(w.segment);
      if (!s.logs[w.employeeName]) s.logs[w.employeeName] = [];
      s.logs[w.employeeName].push({ date: w.date, text: w.text });
    }
  });
  return { since, days, segments };
}

function sendBiweeklyReport() {
  try {
    if (!OWNER_EMAIL) return;
    const { since, days, segments } = generateReportData(14);
    const segNames = Object.keys(segments).sort();
    let inner = '';
    if (segNames.length === 0) {
      inner = '<p style="color:#9c7852;font-size:14px;">Za uplynulé období nejsou žádné dokončené úkoly ani zápisy v deníku.</p>';
    } else {
      segNames.forEach(sn => {
        const s = segments[sn];
        let block = `<div style="margin-bottom:26px;"><div style="font-family:Georgia,serif;font-size:19px;font-weight:bold;color:#13302e;border-bottom:2px solid #dfc196;padding-bottom:6px;margin-bottom:12px;">${sn}</div>`;
        if (s.doneTasks.length) {
          block += `<div style="font-size:12px;font-weight:bold;color:#586845;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">Dokončené úkoly (${s.doneTasks.length})</div>`;
          s.doneTasks.forEach(t => { block += `<div style="font-size:14px;color:#3d301f;margin-bottom:5px;">• ${t.title}${t.who?` <span style="color:#9c7852;font-size:12px;">— ${t.who}</span>`:''}</div>`; });
        }
        const logEmps = Object.keys(s.logs);
        if (logEmps.length) {
          block += `<div style="font-size:12px;font-weight:bold;color:#586845;text-transform:uppercase;letter-spacing:0.5px;margin:12px 0 8px;">Zápisy z deníku</div>`;
          logEmps.forEach(en => {
            block += `<div style="margin-bottom:8px;"><div style="font-size:13px;font-weight:bold;color:#3d301f;">${en}</div>`;
            s.logs[en].forEach(l => { block += `<div style="font-size:13px;color:#6f6757;margin-left:10px;">${l.date?fmtDateShort(l.date)+': ':''}${l.text}</div>`; });
            block += `</div>`;
          });
        }
        if (!s.doneTasks.length && !logEmps.length) block += '<div style="font-size:13px;color:#9c7852;">Žádná aktivita.</div>';
        block += '</div>';
        inner += block;
      });
    }
    const period = fmtDateCz(since.toISOString()) + ' – ' + fmtDateCz(new Date().toISOString());
    const body = `
      <p style="color:#3d301f;font-size:15px;margin:0 0 6px;">Dobrý den,</p>
      <p style="color:#3d301f;font-size:15px;margin:0 0 24px;">přehled činnosti za uplynulé období (<strong>${period}</strong>):</p>
      ${inner}${btnHtml('Otevřít portál', PORTAL_URL)}`;
    GmailApp.sendEmail(OWNER_EMAIL, 'Vista Portál: Report za poslední 2 týdny', 'Přehled činnosti. ' + PORTAL_URL, { htmlBody: emailShell('Report činnosti', body), name: 'Vista Resort' });
  } catch(err) { Logger.log('biweekly report: ' + err.message); }
}

// ============================================================
// HELPERS
// ============================================================
// ── Spreadsheet cache (platí po dobu jednoho requestu) ───────
let _ss = null;
let _sheetCache = {};
function getSS() {
  if (!_ss) _ss = SpreadsheetApp.getActiveSpreadsheet();
  return _ss;
}
function getSheet(name) {
  if (!_sheetCache[name]) {
    const ss = getSS();
    let sheet = ss.getSheetByName(name);
    if (!sheet) { initSheets(); sheet = ss.getSheetByName(name); }
    _sheetCache[name] = sheet;
  }
  return _sheetCache[name];
}

// Batch read — načte všechny listy najednou do paměti, pak getAll čte z cache
function batchReadAll() {
  const ss = getSS();
  const allSheets = ss.getSheets();
  const result = {};
  allSheets.forEach(function(sheet) {
    const name = sheet.getName();
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 1) { result[name] = { headers: [], items: [] }; return; }
    const headers = rows[0];
    const items = rows.slice(1).map(function(r, i) {
      const o = { _row: i + 2 };
      headers.forEach(function(h, ci) { o[h] = r[ci]; });
      return o;
    }).filter(function(o) { return o[headers[0]]; });
    result[name] = { headers: headers, items: items };
  });
  return result;
}

function readSheetFromCache(cache, name) {
  return cache[name] || { headers: [], items: [] };
}
function fmtDateCz(iso) {
  if (!iso) return '';
  try { return new Date(iso).toLocaleDateString('cs-CZ', { day: 'numeric', month: 'long', year: 'numeric' }); } catch(e) { return String(iso); }
}
function fmtDateShort(iso) {
  if (!iso) return '';
  try { return new Date(iso).toLocaleDateString('cs-CZ', { day: 'numeric', month: 'numeric' }); } catch(e) { return String(iso); }
}

// ============================================================
// SETUP
// ============================================================
function setup() {
  initSheets();
  Logger.log('Vista Portál v5 setup hotov. Listy: Zaměstnanci, Dokumenty, Potvrzení, Úkoly, Komentáře, Faktury, Deník.');
}

// Nastaví automatický spouštěč: denní kontrola termínů (8:00). Majiteli se NEodesílají žádné emaily.
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'checkDeadlines' || fn === 'sendBiweeklyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkDeadlines').timeBased().atHour(8).everyDays(1).inTimezone('Europe/Prague').create();
  ScriptApp.newTrigger('runDailyTemplates').timeBased().atHour(8).everyDays(1).inTimezone('Europe/Prague').create();
  Logger.log('Spouštěče nastaveny: denní kontrola termínů + generování šablon úkolů v 8:00.');
}

function resetEmployees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(SHEETS.EMPLOYEES);
  if (old) ss.deleteSheet(old);
  ensureSheet(ss, SHEETS.EMPLOYEES, ['id','name','role','dept','accessLevel','pin','email','createdAt']);
  Logger.log('List Zaměstnanci resetován.');
}

function testCheckDeadlines() { checkDeadlines(); Logger.log('Kontrola termínů spuštěna ručně.'); }
function testReport() { sendBiweeklyReport(); Logger.log('Report spuštěn ručně. Zkontroluj email majitele.'); }

// ============================================================
// NOTIFIKACE — server-side, viditelné napříč zařízeními
// ============================================================
// recipientId: konkrétní user ID, nebo '__admin__', '__gm__', '__owner__', nebo 'all'
function pushNotification(data) {
  const id = 'notif_' + Date.now() + '_' + Math.random().toString(36).slice(2,6);
  appendByHeaders(SHEETS.NOTIFICATIONS, {
    id, recipientId: data.recipientId || 'all', type: data.type || 'info',
    icon: data.icon || '🔔', message: data.message || '', link: data.link || '',
    createdAt: new Date().toISOString(), readBy: ''
  });
  return { success: true, id };
}

function getNotifications(data) {
  const uid = data.uid || '';
  const since = data.since ? new Date(data.since) : new Date(Date.now() - 30*86400000);
  const { items } = readSheet(SHEETS.NOTIFICATIONS);
  const notifs = items.filter(function(n) {
    if (!n.id) return false;
    if (new Date(n.createdAt) < since) return false;
    // Recipient matches: this user, their role, or 'all'
    const r = String(n.recipientId);
    if (r !== 'all' && r !== uid && r !== data.role) return false;
    return true;
  }).map(function(n) {
    const readBy = String(n.readBy || '');
    return { id:n.id, type:n.type||'info', icon:n.icon||'🔔', message:n.message||'', link:n.link||'', createdAt:n.createdAt, read: readBy.split(',').filter(Boolean).indexOf(uid) >= 0 };
  }).sort(function(a,b){ return new Date(b.createdAt)-new Date(a.createdAt); }).slice(0,60);
  return { notifications: notifs };
}

function markNotifRead(data) {
  const uid = data.uid || '';
  const ids = Array.isArray(data.ids) ? data.ids : [data.id];
  const sheet = getSheet(SHEETS.NOTIFICATIONS);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('id'), readCol = headers.indexOf('readBy');
  if (idCol < 0 || readCol < 0) return { success: false };
  for (let i = 1; i < rows.length; i++) {
    if (ids.indexOf(rows[i][idCol]) >= 0) {
      const existing = String(rows[i][readCol]);
      const readers = existing.split(',').filter(Boolean);
      if (readers.indexOf(uid) < 0) { readers.push(uid); sheet.getRange(i+1, readCol+1).setValue(readers.join(',')); }
    }
  }
  return { success: true };
}

// Pomocná funkce pro ostatní endpointy — pošle notif na vedení
function notifyManagement(type, icon, message, link) {
  ['__owner__','__admin__','__gm__'].forEach(function(r) { pushNotification({ recipientId:r, type, icon, message, link }); });
}
// Pošle notif konkrétnímu zaměstnanci
function notifyUser(uid, type, icon, message, link) {
  pushNotification({ recipientId:uid, type, icon, message, link:link||'' });
}

// ============================================================
// ŠABLONY ÚKOLŮ
// ============================================================
function getTemplates() {
  const { items } = readSheet(SHEETS.TEMPLATES);
  return { templates: items.filter(function(r){ return r.id; }).map(function(r) {
    var assignees=[],days=[];
    try{ assignees=JSON.parse(r.assignees||'[]'); }catch(e){}
    try{ days=JSON.parse(r.days||'[]').map(Number); }catch(e){}
    return { id:r.id, title:r.title||'', desc:r.desc||'', segment:r.segment||'', priority:r.priority||'medium',
      assignees, freq:r.freq||'weekly', days, monthDay:r.monthDay||'', deadlineOffset:parseInt(r.deadlineOffset)||0,
      active: r.active===true||r.active==='true'||r.active==='ANO'||r.active===1,
      createdAt:r.createdAt||'', lastRun:r.lastRun||'' };
  }) };
}
function saveTemplate(data) {
  var existing = readSheet(SHEETS.TEMPLATES).items.find(function(r){ return r.id===data.id; });
  var row = { id:data.id, title:data.title||'', desc:data.desc||'', segment:data.segment||'',
    priority:data.priority||'medium', assignees:JSON.stringify(data.assignees||[]),
    freq:data.freq||'weekly', days:JSON.stringify(data.days||[]), monthDay:data.monthDay||'',
    deadlineOffset:data.deadlineOffset||0, active:data.active?true:false,
    createdAt:data.createdAt||new Date().toISOString(), lastRun:data.lastRun||'' };
  if (existing) { updateRowByHeaders(SHEETS.TEMPLATES,'id',data.id,row); }
  else { appendByHeaders(SHEETS.TEMPLATES, row); }
  return { success:true };
}
function deleteTemplate(id) { deleteRowByCol(SHEETS.TEMPLATES,'id',id); return { success:true }; }
function updateTemplateLastRun(data) { updateRowByHeaders(SHEETS.TEMPLATES,'id',data.id,{lastRun:data.lastRun||new Date().toISOString()}); return { success:true }; }

// Denní trigger spouští šablony (běží v 8:00 přes setupTrigger)
function runDailyTemplates() {
  var today = new Date();
  var dow = today.getDay();
  var dom = today.getDate();
  var lastDay = new Date(today.getFullYear(),today.getMonth()+1,0).getDate();
  var todayKey = today.toISOString().slice(0,10);
  var templates = getTemplates().templates;
  templates.forEach(function(t) {
    if (!t.active) return;
    if (t.lastRun && t.lastRun.slice(0,10)===todayKey) return;
    var shouldRun=false;
    if (t.freq==='daily') shouldRun=true;
    else if (t.freq==='weekly') shouldRun=t.days.indexOf(dow)>=0;
    else if (t.freq==='monthly') { if (t.monthDay==='last') shouldRun=(dom===lastDay); else shouldRun=(dom===parseInt(t.monthDay)); }
    if (!shouldRun) return;
    var deadline='';
    if (t.deadlineOffset>0) { var d=new Date(today.getTime()+t.deadlineOffset*86400000); deadline=d.toISOString().slice(0,10); }
    else deadline=todayKey;
    var taskId='task_'+Date.now()+'_'+Math.random().toString(36).slice(2,5);
    appendByHeaders(SHEETS.TASKS,{id:taskId,title:t.title,desc:t.desc||'',segment:t.segment||'',priority:t.priority||'medium',
      status:'todo',assignees:JSON.stringify(t.assignees||[]),dependsOn:'[]',deadline:deadline,
      createdBy:'__system__',createdByName:'Systém (šablona)',createdAt:new Date().toISOString(),
      completedAt:'',attachments:'[]',remindedAt:''});
    updateRowByHeaders(SHEETS.TEMPLATES,'id',t.id,{lastRun:new Date().toISOString()});
    // notif na přiřazené zaměstnance
    (t.assignees||[]).forEach(function(uid){ notifyUser(uid,'task','📌','Nový opakující se úkol: '+t.title,''); });
    Logger.log('Vygenerován úkol z šablony: '+t.title+' ('+taskId+')');
  });
}

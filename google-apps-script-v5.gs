// TEAM ANDENG KPI - Google Apps Script backend (v5)
// Paste this entire file into Extensions > Apps Script in the connected Sheet.

const SHEET_NAME = 'Records';
const AGENTS_SHEET = 'Agents';
const CONFIG_SHEET = 'Config';

function doGet(e) {
  if (!isAuthorized_(e && e.parameter && e.parameter.accessCode)) {
    return json_({ ok: false, error: 'unauthorized' });
  }

  const sh = getSheet_();
  const rows = sh.getDataRange().getValues();
  rows.shift();
  const data = rows.map(function(r) {
    return {
      id: String(r[0]),
      date: fmtDate_(r[1]),
      agent: String(r[2]),
      status: String(r[3] || ''),
      activity: String(r[4]),
      count: Number(r[5] || 0),
      points: Number(r[6] || 0),
      notes: String(r[7] || '')
    };
  }).filter(function(r) {
    return r.id && r.agent && r.id !== 'id';
  });

  const ash = getAgentsSheet_();
  const arows = ash.getDataRange().getValues();
  arows.shift();
  const agents = arows.map(function(r) {
    return String(r[0]).trim();
  }).filter(function(n) {
    return n && n !== 'name';
  });

  const csh = getConfigSheet_();
  const cvals = csh.getDataRange().getValues();
  cvals.shift();
  const config = cvals.map(function(r) {
    return String(r[0] || '').replace(/^#/, '');
  }).join('');

  return json_({
    ok: true,
    securityVersion: 5,
    records: data,
    agents: agents,
    config: config || null
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (!isAuthorized_(body.accessCode)) {
      return json_({ ok: false, error: 'unauthorized' });
    }

    if (body.action === 'add') {
      const sh = getSheet_();
      (body.records || []).forEach(function(r) {
        sh.appendRow([
          String(r.id), String(r.date), String(r.agent),
          String(r.status || ''), String(r.activity),
          Number(r.count || 0), Number(r.points || 0),
          String(r.notes || '')
        ]);
      });
      return json_({ ok: true });
    }

    if (body.action === 'delete') {
      const sh = getSheet_();
      const data = sh.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(body.id)) sh.deleteRow(i + 1);
      }
      return json_({ ok: true });
    }

    if (body.action === 'addAgent') {
      const ash = getAgentsSheet_();
      const name = String(body.name || '').trim();
      if (!name) return json_({ ok: false, error: 'blank name' });
      const existing = ash.getDataRange().getValues().slice(1).map(function(r) {
        return String(r[0]).trim().toLowerCase();
      });
      if (existing.indexOf(name.toLowerCase()) === -1) ash.appendRow([name]);
      return json_({ ok: true });
    }

    if (body.action === 'deleteAgent') {
      const ash = getAgentsSheet_();
      const name = String(body.name || '').trim().toLowerCase();
      const data = ash.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]).trim().toLowerCase() === name) ash.deleteRow(i + 1);
      }
      return json_({ ok: true });
    }

    if (body.action === 'saveConfig') {
      const csh = getConfigSheet_();
      const str = String(body.config || '');
      if (csh.getLastRow() > 1) csh.deleteRows(2, csh.getLastRow() - 1);
      const chunkSize = 45000;
      for (let i = 0; i < str.length; i += chunkSize) {
        csh.appendRow(['#' + str.slice(i, i + chunkSize)]);
      }
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function isAuthorized_(providedCode) {
  const expectedCode = PropertiesService.getScriptProperties()
    .getProperty('TEAM_ACCESS_CODE');
  return Boolean(expectedCode) && String(providedCode || '') === expectedCode;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'id', 'date', 'agent', 'status',
      'activity', 'count', 'points', 'notes'
    ]);
  }
  return sh;
}

function getAgentsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AGENTS_SHEET);
  if (!sh) sh = ss.insertSheet(AGENTS_SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(['name']);
  return sh;
}

function getConfigSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG_SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(['config']);
  return sh;
}

function fmtDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Manila', 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
  const date = new Date(text);
  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(date, 'Asia/Manila', 'yyyy-MM-dd');
  }
  return text.slice(0, 10);
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

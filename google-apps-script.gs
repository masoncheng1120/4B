/*
  Google Apps Script for Student Schedule Tracker

  Supports:
  - action=logEntry with op=upsert|delete
  - action=adminData with password check
  - JSON and JSONP responses
*/

const SHEET_NAME = 'DailyLog';
const ADMIN_PASSWORD = 'mason';
const HEADER = [
  'Date',
  'Name',
  'Study Hours',
  'Sleep Hours',
  'Last Update Type',
  'Saved At (Client)',
  'Received At (Server)',
];

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = String(params.action || '').trim();
    const callback = String(params.callback || '').trim();

    if (action === 'logEntry') {
      const payload = parsePayloadFromParams_(params);
      const result = payload.op === 'delete' ? deleteEntry_(payload) : upsertEntry_(payload);
      return jsonOrJsonp_({ ok: true, result: result }, callback);
    }

    if (action === 'adminData') {
      const password = String(params.password || '');
      if (password !== ADMIN_PASSWORD) {
        return jsonOrJsonp_({ ok: false, error: 'Invalid admin password.' }, callback);
      }
      return jsonOrJsonp_(getAdminData_(), callback);
    }

    return jsonOrJsonp_({ ok: true, message: 'Schedule Tracker endpoint is running.' }, callback);
  } catch (error) {
    const callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
    return jsonOrJsonp_({ ok: false, error: String(error) }, callback);
  }
}

function doPost(e) {
  try {
    const payload = parsePayloadFromBody_(e);
    const result = payload.op === 'delete' ? deleteEntry_(payload) : upsertEntry_(payload);
    return jsonOutput_({ ok: true, result: result });
  } catch (error) {
    return jsonOutput_({ ok: false, error: String(error) });
  }
}

function parsePayloadFromBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Missing POST body');
  }
  const data = JSON.parse(e.postData.contents);
  return parsePayloadObject_(data);
}

function parsePayloadFromParams_(params) {
  return parsePayloadObject_({
    op: params.op,
    type: params.type,
    date: params.date,
    name: params.name,
    hours: params.hours,
    savedAt: params.savedAt,
  });
}

function parsePayloadObject_(data) {
  const op = String(data.op || 'upsert').trim().toLowerCase();
  const type = String(data.type || '').trim().toLowerCase();
  const date = String(data.date || '').trim();
  const name = String(data.name || '').trim();
  const hours = Number(data.hours);
  const savedAt = String(data.savedAt || '').trim();

  if (op !== 'upsert' && op !== 'delete') {
    throw new Error('Invalid op. Expected upsert or delete.');
  }
  if (type !== 'study' && type !== 'sleep') {
    throw new Error('Invalid type. Expected study or sleep.');
  }
  if (!date) throw new Error('Missing date');
  if (!name) throw new Error('Missing name');
  if (op === 'upsert' && !Number.isFinite(hours)) {
    throw new Error('Invalid hours');
  }

  return {
    op: op,
    type: type,
    date: date,
    name: name,
    hours: hours,
    savedAt: savedAt,
  };
}

function upsertEntry_(payload) {
  const sheet = getOrCreateSheet_();
  const rowIndex = findRow_(sheet, payload.date, payload.name);

  const studyValue = payload.type === 'study' ? payload.hours : '';
  const sleepValue = payload.type === 'sleep' ? payload.hours : '';

  if (rowIndex > 0) {
    const existing = sheet.getRange(rowIndex, 1, 1, HEADER.length).getValues()[0];
    const mergedStudy = payload.type === 'study' ? payload.hours : existing[2];
    const mergedSleep = payload.type === 'sleep' ? payload.hours : existing[3];

    sheet.getRange(rowIndex, 1, 1, HEADER.length).setValues([[
      payload.date,
      payload.name,
      mergedStudy,
      mergedSleep,
      payload.type,
      payload.savedAt,
      new Date(),
    ]]);

    return { action: 'updated', row: rowIndex };
  }

  sheet.appendRow([
    payload.date,
    payload.name,
    studyValue,
    sleepValue,
    payload.type,
    payload.savedAt,
    new Date(),
  ]);

  return { action: 'inserted', row: sheet.getLastRow() };
}

function deleteEntry_(payload) {
  const sheet = getOrCreateSheet_();
  const rowIndex = findRow_(sheet, payload.date, payload.name);
  if (rowIndex <= 0) {
    return { action: 'not-found' };
  }

  const existing = sheet.getRange(rowIndex, 1, 1, HEADER.length).getValues()[0];
  const mergedStudy = payload.type === 'study' ? '' : existing[2];
  const mergedSleep = payload.type === 'sleep' ? '' : existing[3];

  if (String(mergedStudy || '') === '' && String(mergedSleep || '') === '') {
    sheet.deleteRow(rowIndex);
    return { action: 'deleted-row', row: rowIndex };
  }

  sheet.getRange(rowIndex, 1, 1, HEADER.length).setValues([[
    payload.date,
    payload.name,
    mergedStudy,
    mergedSleep,
    'delete',
    payload.savedAt,
    new Date(),
  ]]);

  return { action: 'cleared', row: rowIndex };
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  ensureHeader_(sheet);
  return sheet;
}

function ensureHeader_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, HEADER.length).getValues()[0];
  const hasHeader = firstRow.join('|') === HEADER.join('|');
  if (!hasHeader) {
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  }
}

function findRow_(sheet, date, name) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const targetDate = normalizeDate_(date);
  const targetName = String(name || '').trim();

  for (var i = 0; i < values.length; i++) {
    const rowDate = normalizeDate_(values[i][0]);
    const rowName = String(values[i][1] || '').trim();
    if (rowDate === targetDate && rowName === targetName) {
      return i + 2;
    }
  }

  return -1;
}

function getAdminData_() {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { ok: true, rows: [], students: [] };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
  const byKey = {};
  const students = {};

  values.forEach(function (row) {
    const date = normalizeDate_(row[0]);
    const name = String(row[1] || '').trim();
    if (!date || !name) return;

    const study = toNumberOrBlank_(row[2]);
    const sleep = toNumberOrBlank_(row[3]);

    students[name] = true;

    const key = date + '|' + name;
    if (!byKey[key]) {
      byKey[key] = {
        date: date,
        name: name,
        studyHours: '',
        sleepHours: '',
      };
    }

    if (study !== '') byKey[key].studyHours = study;
    if (sleep !== '') byKey[key].sleepHours = sleep;
  });

  const rows = Object.keys(byKey).map(function (key) {
    return byKey[key];
  });

  rows.sort(function (a, b) {
    if (a.date === b.date) return a.name.localeCompare(b.name);
    return a.date < b.date ? -1 : 1;
  });

  const studentList = Object.keys(students).sort();
  return { ok: true, rows: rows, students: studentList };
}

function normalizeDate_(value) {
  if (value === null || value === undefined || value === '') return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const asText = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(asText)) {
    return asText;
  }

  const parsed = new Date(asText);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return '';
}

function toNumberOrBlank_(value) {
  if (value === null || value === undefined || value === '') return '';
  const n = Number(value);
  return Number.isFinite(n) ? n : '';
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonOrJsonp_(obj, callback) {
  const cb = String(callback || '').trim();
  if (cb) {
    const safeCallback = cb.replace(/[^a-zA-Z0-9_.$]/g, '');
    return ContentService
      .createTextOutput(safeCallback + '(' + JSON.stringify(obj) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOutput_(obj);
}

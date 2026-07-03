/*
  Google Apps Script for Student Schedule Tracker

  What it does:
  - Accepts POST JSON from the website.
  - Upserts rows into one sheet named DailyLog.
  - Keeps study and sleep hours in separate columns.

  Expected payload:
  {
    "type": "study" | "sleep",
    "date": "YYYY-MM-DD",
    "name": "Student Name",
    "hours": 7.5,
    "savedAt": "2026-03-18T12:34:56.000Z"
  }
*/

const SHEET_NAME = 'DailyLog';
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
      const result = upsertEntry_(payload);
      return jsonOrJsonp_({ ok: true, result: result }, callback);
    }

    if (action === 'adminData') {
      const password = String(params.password || '');
      if (password !== 'mason') {
        return jsonOrJsonp_({ ok: false, error: 'Invalid admin password.' }, callback);
      }
      return jsonOrJsonp_(getAdminData_(), callback);
    }

    return jsonOrJsonp_({ ok: true, message: 'Schedule Tracker endpoint is running.' }, callback);
  } catch (error) {
    return jsonOrJsonp_({ ok: false, error: String(error) }, (e && e.parameter && e.parameter.callback) || null);
  }
}

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    const result = upsertEntry_(payload);
    return jsonOutput_({ ok: true, result: result });
  } catch (error) {
    return jsonOutput_({ ok: false, error: String(error) });
  }
}

function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Missing POST body');
  }

  const data = JSON.parse(e.postData.contents);
  return parsePayloadObject_(data);
}

function parsePayloadFromParams_(params) {
  const data = {
    type: params.type,
    date: params.date,
    name: params.name,
    hours: params.hours,
    savedAt: params.savedAt,
  };
  return parsePayloadObject_(data);
}

function parsePayloadObject_(data) {
  const type = String(data.type || '').toLowerCase();
  const date = String(data.date || '').trim();
  const name = String(data.name || '').trim();
  const hours = Number(data.hours);
  const savedAt = String(data.savedAt || '').trim();

  if (type !== 'study' && type !== 'sleep') {
    throw new Error('Invalid type. Expected study or sleep.');
  }
  if (!date) throw new Error('Missing date');
  if (!name) throw new Error('Missing name');
  if (!Number.isFinite(hours)) throw new Error('Invalid hours');

  return { type, date, name, hours, savedAt };
}

function upsertEntry_(payload) {
  const sheet = getOrCreateSheet_();
  const rowIndex = findRow_(sheet, payload.date, payload.name);

  const studyValue = payload.type === 'study' ? payload.hours : '';
  const sleepValue = payload.type === 'sleep' ? payload.hours : '';

  if (rowIndex > 0) {
    const existing = sheet.getRange(rowIndex, 1, 1, HEADER.length).getValues()[0];
    const updatedStudy = payload.type === 'study' ? payload.hours : existing[2];
    const updatedSleep = payload.type === 'sleep' ? payload.hours : existing[3];

    sheet.getRange(rowIndex, 1, 1, HEADER.length).setValues([[
      payload.date,
      payload.name,
      updatedStudy,
      updatedSleep,
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

  for (let i = 0; i < values.length; i++) {
    const rowDate = normalizeDate_(values[i][0]);
    const rowName = String(values[i][1]).trim();
    if (rowDate === targetDate && rowName === targetName) {
      return i + 2;
    }
  }
  return -1;
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

function getAdminData_() {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { ok: true, rows: [], students: [] };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
  const rowsByKey = {};
  const studentSet = {};

  values.forEach(function (r) {
    const date = normalizeDate_(r[0]);
    const name = String(r[1] || '').trim();
    const studyHours = toNumberOrBlank_(r[2]);
    const sleepHours = toNumberOrBlank_(r[3]);
    if (!date || !name) return;

    studentSet[name] = true;

    const key = date + '|' + name;
    if (!rowsByKey[key]) {
      rowsByKey[key] = {
        date: date,
        name: name,
        studyHours: '',
        sleepHours: '',
      };
    }

    /* Merge duplicate rows so study/sleep values can come from different updates */
    if (studyHours !== '') rowsByKey[key].studyHours = studyHours;
    if (sleepHours !== '') rowsByKey[key].sleepHours = sleepHours;
  });

  const rows = Object.keys(rowsByKey).map(function (k) {
    return rowsByKey[k];
  });

  const students = Object.keys(studentSet).sort();
  return { ok: true, rows: rows, students: students };
}

function normalizeDate_(value) {
  if (value === null || value === undefined || value === '') return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const str = String(value).trim();
  const ymdMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (ymdMatch) return str;

  const parsed = new Date(str);
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

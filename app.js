/* ============================================================
   app.js  — Shared logic for Schedule Tracker
   All data is stored in localStorage, keyed by student name.
   No external dependencies.
   ============================================================ */

// ── KEY HELPERS ─────────────────────────────────────────────

const NAME_KEY = 'scheduleTracker_studentName';
/* Paste your deployed Apps Script Web App URL here */
const SHEETS_WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbzt40ULW2vebkbU7QI0qYNa0li_c2xcLjQw5Ssc8CEGjKGgfEN92CyZmDRKVSo5BlL0gw/exec';
const SHEETS_WEBHOOK_KEY = 'scheduleTracker_sheetsWebhookUrl_v2';
const SYNC_QUEUE_KEY = 'scheduleTracker_syncQueue_v2';
const SYNC_RETRY_BASE_MS = 2000;
const SYNC_RETRY_MAX_MS = 300000;
const _memoryStorageFallback = {};
let _syncFlushTimer = null;
let _syncInFlight = false;

function readStorage(key) {
  try {
    return localStorage.getItem(key);
  } catch {
    return Object.prototype.hasOwnProperty.call(_memoryStorageFallback, key)
      ? _memoryStorageFallback[key]
      : null;
  }
}

function writeStorage(key, value) {
  try {
    localStorage.setItem(key, value);
  } catch {
    _memoryStorageFallback[key] = String(value);
  }
}

function removeStorage(key) {
  try {
    localStorage.removeItem(key);
  } catch {
    delete _memoryStorageFallback[key];
  }
}

function getStudentName() {
  return (readStorage(NAME_KEY) || '').trim();
}

function setStudentName(name) {
  writeStorage(NAME_KEY, name.trim());
}

function clearStudentName() {
  removeStorage(NAME_KEY);
}

function getSheetsWebhookUrl() {
  const value = readStorage(SHEETS_WEBHOOK_KEY) || SHEETS_WEBHOOK_URL;
  return String(value || '').trim();
}

function setSheetsWebhookUrl(url) {
  writeStorage(SHEETS_WEBHOOK_KEY, String(url || '').trim());
}

function getAppsScriptBaseUrl() {
  const raw = getSheetsWebhookUrl();
  if (!raw || raw.indexOf('PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE') !== -1) {
    return '';
  }
  return raw;
}

function buildAppsScriptUrl(params) {
  const base = getAppsScriptBaseUrl();
  if (!base) return '';
  try {
    const url = new URL(base);
    Object.keys(params || {}).forEach(key => {
      if (params[key] !== undefined && params[key] !== null) {
        url.searchParams.set(key, String(params[key]));
      }
    });
    return url.toString();
  } catch {
    return '';
  }
}

function fetchAdminDataFromSheet(password) {
  const jsonEndpoint = buildAppsScriptUrl({ action: 'adminData', password: password });

  if (!jsonEndpoint) {
    return Promise.reject(new Error('Apps Script URL is not configured.'));
  }

  const controller = new AbortController();
  const fetchTimeout = setTimeout(() => controller.abort(), 15000);

  return fetch(jsonEndpoint, { signal: controller.signal })
    .then(res => {
      clearTimeout(fetchTimeout);
      if (!res.ok) throw new Error('Server error: ' + res.status);
      return res.json();
    })
    .then(data => {
      if (!data || data.ok !== true) {
        throw new Error(data && data.error ? data.error : 'Failed to load admin data.');
      }
      return data;
    })
    .catch(err => {
      clearTimeout(fetchTimeout);
      /* fetch blocked by CORS on the GAS redirect — fall back to JSONP */
      if (err.name === 'TypeError') {
        return fetchAdminDataJsonp_(password);
      }
      if (err.name === 'AbortError') {
        throw new Error('Request timed out while loading admin data.');
      }
      throw err;
    });
}

function fetchAdminDataJsonp_(password) {
  const callbackName = '__adminCb_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
  const endpoint = buildAppsScriptUrl({
    action: 'adminData',
    password: password,
    callback: callbackName,
  });

  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    let settled = false;

    const cleanup = () => {
      try { delete window[callbackName]; } catch { window[callbackName] = undefined; }
      if (script.parentNode) script.parentNode.removeChild(script);
    };

    const timeout = setTimeout(() => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(new Error('Request timed out. Make sure the Apps Script is deployed as "Anyone can access".'));
    }, 20000);

    window[callbackName] = (data) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      cleanup();
      if (!data || data.ok !== true) {
        reject(new Error(data && data.error ? data.error : 'Failed to load admin data.'));
        return;
      }
      resolve(data);
    };

    script.onerror = () => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      cleanup();
      reject(new Error('Cannot reach the Apps Script. Check it is deployed as "Anyone can access".'));
    };

    script.src = endpoint;
    document.head.appendChild(script);
  });
}

/**
 * If no student name is saved, redirect to index.
 * Otherwise show the name in the nav and return it.
 */
function guardName() {
  const name = getStudentName();
  if (!name) {
    window.location.href = 'index.html';
    return '';
  }
  showNavStudent(name);
  return name;
}

function showNavStudent(name) {
  const el = document.getElementById('navStudent');
  const nm = document.getElementById('navName');
  if (el) el.style.display = 'flex';
  if (nm) nm.textContent   = name;
}

// ── ENTRY STORAGE ────────────────────────────────────────────
// Storage schema:
//   localStorage key: `scheduleTracker_<type>`   (e.g. scheduleTracker_study)
//   value: JSON array of Entry objects
//
// Entry shape:
//   { id, student, date, hours, savedAt }

function storageKey(type) {
  return 'scheduleTracker_' + type;
}

function getEntries(type) {
  try {
    return JSON.parse(readStorage(storageKey(type))) || [];
  } catch {
    return [];
  }
}

function saveEntry(type, date, hours) {
  const name    = getStudentName();
  const entries = getEntries(type);
  let savedEntry;

  /* If an entry already exists for this student + date, update it */
  const existing = entries.findIndex(e => e.student === name && e.date === date);
  if (existing !== -1) {
    entries[existing].hours   = hours;
    entries[existing].savedAt = new Date().toISOString();
    savedEntry = entries[existing];
  } else {
    const newEntry = {
      id:      crypto.randomUUID(),
      student: name,
      date:    date,
      hours:   hours,
      savedAt: new Date().toISOString(),
    };
    entries.push(newEntry);
    savedEntry = newEntry;
  }

  writeStorage(storageKey(type), JSON.stringify(entries));
  syncEntryToGoogleSheet(type, savedEntry);
}

function removeEntry(type, id) {
  const entries = getEntries(type).filter(e => e.id !== id);
  writeStorage(storageKey(type), JSON.stringify(entries));
}

function clearEntries(type) {
  removeStorage(storageKey(type));
}

function syncEntryToGoogleSheet(type, entry) {
  if (!entry) return;

  const queue = getSyncQueue();
  queue.push({
    id: String(entry.id || (Date.now() + '_' + Math.floor(Math.random() * 100000))),
    payload: {
      type: type,
      date: entry.date,
      name: entry.student,
      hours: entry.hours,
      savedAt: entry.savedAt,
    },
    retryCount: 0,
    nextAttemptAt: Date.now(),
  });

  if (queue.length > 300) {
    queue.splice(0, queue.length - 300);
  }

  setSyncQueue(queue);
  scheduleSyncFlush(0);
}

function getSyncQueue() {
  try {
    const parsed = JSON.parse(readStorage(SYNC_QUEUE_KEY) || '[]');
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function setSyncQueue(queue) {
  writeStorage(SYNC_QUEUE_KEY, JSON.stringify(Array.isArray(queue) ? queue : []));
}

function getRetryDelayMs(retryCount) {
  const step = Math.max(0, Number(retryCount) || 0);
  return Math.min(SYNC_RETRY_MAX_MS, SYNC_RETRY_BASE_MS * Math.pow(2, step));
}

function scheduleSyncFlush(delayMs) {
  if (_syncFlushTimer) clearTimeout(_syncFlushTimer);
  _syncFlushTimer = setTimeout(() => {
    _syncFlushTimer = null;
    flushSyncQueue();
  }, Math.max(0, Number(delayMs) || 0));
}

function sendSyncPayloadJsonp_(payload) {
  const callbackName = '__syncCb_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
  const endpoint = buildAppsScriptUrl({
    action: 'logEntry',
    type: payload.type,
    date: payload.date,
    name: payload.name,
    hours: payload.hours,
    savedAt: payload.savedAt,
    callback: callbackName,
    _: Date.now(),
  });

  if (!endpoint) {
    return Promise.reject(new Error('Apps Script URL is not configured.'));
  }

  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    let settled = false;

    const cleanup = () => {
      try { delete window[callbackName]; } catch { window[callbackName] = undefined; }
      if (script.parentNode) script.parentNode.removeChild(script);
    };

    const timeout = setTimeout(() => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(new Error('Timed out while syncing entry.'));
    }, 15000);

    window[callbackName] = (data) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      cleanup();
      if (!data || data.ok !== true) {
        reject(new Error(data && data.error ? data.error : 'Sync failed.'));
        return;
      }
      resolve(true);
    };

    script.onerror = () => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      cleanup();
      reject(new Error('Network error while syncing entry.'));
    };

    script.src = endpoint;
    document.head.appendChild(script);
  });
}

function flushSyncQueue() {
  if (_syncInFlight) return;

  const queue = getSyncQueue();
  if (!queue.length) return;

  const now = Date.now();
  const idx = queue.findIndex(item => Number(item.nextAttemptAt || 0) <= now);

  if (idx === -1) {
    const nextAt = Math.min(...queue.map(item => Number(item.nextAttemptAt || now + 5000)));
    scheduleSyncFlush(Math.max(300, nextAt - now));
    return;
  }

  const item = queue[idx];
  _syncInFlight = true;

  sendSyncPayloadJsonp_(item.payload)
    .then(() => {
      _syncInFlight = false;
      const latest = getSyncQueue();
      const removeIdx = latest.findIndex(x => x.id === item.id);
      if (removeIdx !== -1) {
        latest.splice(removeIdx, 1);
        setSyncQueue(latest);
      }
      scheduleSyncFlush(0);
    })
    .catch(() => {
      _syncInFlight = false;
      const latest = getSyncQueue();
      const failIdx = latest.findIndex(x => x.id === item.id);
      if (failIdx !== -1) {
        const cur = latest[failIdx];
        const retries = (Number(cur.retryCount) || 0) + 1;
        cur.retryCount = retries;
        cur.nextAttemptAt = Date.now() + getRetryDelayMs(retries);
        latest[failIdx] = cur;
        setSyncQueue(latest);
        scheduleSyncFlush(getRetryDelayMs(retries));
      }
    });
}

function initSyncQueueEngine() {
  scheduleSyncFlush(1200);
  window.addEventListener('online', () => scheduleSyncFlush(0));
  window.addEventListener('pageshow', () => scheduleSyncFlush(0));
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') {
      scheduleSyncFlush(0);
    }
  });
}

// ── DATE / TIME FORMATTERS ───────────────────────────────────

/**
 * Format a yyyy-mm-dd string to a readable date.
 * Parses the date in local time to avoid off-by-one timezone issues.
 */
function formatDate(dateStr) {
  const [y, m, d] = dateStr.split('-').map(Number);
  const date = new Date(y, m - 1, d);
  return date.toLocaleDateString(undefined, {
    weekday: 'short',
    year:    'numeric',
    month:   'short',
    day:     'numeric',
  });
}

/** Format an ISO timestamp to a short locale string */
function formatTimestamp(iso) {
  if (!iso) return '—';
  return new Date(iso).toLocaleString(undefined, {
    month:  'short',
    day:    'numeric',
    hour:   '2-digit',
    minute: '2-digit',
  });
}

// ── SECURITY: HTML ESCAPING ──────────────────────────────────

/**
 * Escape user-supplied strings before inserting into innerHTML.
 * Prevents XSS when student names or dates are rendered in table rows.
 */
function escHtml(str) {
  return String(str)
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&#039;');
}

// ── TOAST NOTIFICATION ───────────────────────────────────────

let _toastTimer = null;

function showToast(message) {
  const toast = document.getElementById('toast');
  if (!toast) return;

  toast.textContent = message;
  toast.classList.add('show');

  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => toast.classList.remove('show'), 2600);
}

initSyncQueueEngine();

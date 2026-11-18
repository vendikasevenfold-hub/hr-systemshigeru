'use strict';
/**
 * Pengganti layanan Google Apps Script supaya Code.gs & Setup.gs bisa berjalan di Node.js
 * tanpa ditulis ulang: SpreadsheetApp, Utilities, Session, LockService, CacheService,
 * PropertiesService, UrlFetchApp, HtmlService, Logger.
 *
 * File .gs dijalankan di konteks global proses ini (vm.runInThisContext), sama seperti
 * Apps Script menjalankan semua file .gs di satu ruang global. Proses ini khusus untuk
 * aplikasi HR, jadi tidak ada kode lain yang ikut tercampur.
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const { isDate } = require('./workbook');

const GAS_FILES = ['Code.gs', 'Setup.gs'];
const TIME_ZONE = process.env.TZ || 'Asia/Jakarta';

// ---------------- Utilities ----------------

const dtfCache = new Map();
function partsIn(date, tz) {
  let f = dtfCache.get(tz);
  if (!f) {
    f = new Intl.DateTimeFormat('en-GB', {
      timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
    });
    dtfCache.set(tz, f);
  }
  const p = {};
  for (const x of f.formatToParts(date)) p[x.type] = x.value;
  return p;
}

const DAY_ID = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

function formatDate(date, tz, pattern) {
  if (!isDate(date) || isNaN(date.getTime())) throw new Error('Utilities.formatDate: tanggal tidak valid.');
  const p = partsIn(date, tz || TIME_ZONE);
  const year = p.year.padStart(4, '0');
  return String(pattern).replace(/yyyy|yy|MM|dd|HH|mm|ss|EEEE|'[^']*'/g, (tok) => {
    switch (tok) {
      case 'yyyy': return year;
      case 'yy': return year.slice(-2);
      case 'MM': return p.month;
      case 'dd': return p.day;
      case 'HH': return p.hour;
      case 'mm': return p.minute;
      case 'ss': return p.second;
      case 'EEEE': {
        const d = new Date(Date.UTC(+p.year, +p.month - 1, +p.day));
        return DAY_ID[d.getUTCDay()];
      }
      default: return tok.slice(1, -1);
    }
  });
}

const sleepBuf = new Int32Array(new SharedArrayBuffer(4));

const Utilities = {
  formatDate,
  getUuid: () => crypto.randomUUID(),
  sleep: (ms) => { Atomics.wait(sleepBuf, 0, 0, Math.max(0, Number(ms) || 0)); },
  base64Encode: (s) => Buffer.from(String(s), 'utf8').toString('base64'),
  base64Decode: (s) => [...Buffer.from(String(s), 'base64')],
};

const Session = {
  getScriptTimeZone: () => TIME_ZONE,
  getActiveUser: () => ({ getEmail: () => '' }),
  getEffectiveUser: () => ({ getEmail: () => '' }),
};

// ---------------- LockService ----------------
// Node menjalankan setiap panggilan fungsi .gs secara berurutan (sinkron, satu proses),
// jadi dua panggilan tidak pernah berjalan bersamaan. Kunci cukup berupa tiruan.
const scriptLock = {
  waitLock() {},
  tryLock() { return true; },
  releaseLock() {},
  hasLock() { return true; },
};
const LockService = {
  getScriptLock: () => scriptLock,
  getDocumentLock: () => scriptLock,
  getUserLock: () => scriptLock,
};

// ---------------- CacheService ----------------
const cacheStore = new Map();
const scriptCache = {
  get(key) {
    const e = cacheStore.get(key);
    if (!e) return null;
    if (e.exp < Date.now()) { cacheStore.delete(key); return null; }
    return e.v;
  },
  put(key, value, ttlSec) {
    const ttl = Math.min(Math.max(Number(ttlSec) || 600, 1), 21600);
    cacheStore.set(key, { v: String(value), exp: Date.now() + ttl * 1000 });
  },
  remove(key) { cacheStore.delete(key); },
};
// Bersihkan entri kedaluwarsa sesekali supaya memori tidak menumpuk.
setInterval(() => {
  const now = Date.now();
  for (const [k, e] of cacheStore) if (e.exp < now) cacheStore.delete(k);
}, 10 * 60 * 1000).unref();
const CacheService = { getScriptCache: () => scriptCache, getUserCache: () => scriptCache };

// ---------------- PropertiesService ----------------
// Dulu: Project Settings > Script Properties. Sekarang: variabel lingkungan (.env).
const scriptProperties = {
  getProperty: (key) => (process.env[key] === undefined || process.env[key] === '' ? null : process.env[key]),
  getProperties: () => ({ ...process.env }),
};
const PropertiesService = { getScriptProperties: () => scriptProperties };

// ---------------- UrlFetchApp ----------------
// Code.gs memanggil fetch secara sinkron; permintaan HTTP dijalankan di proses anak
// (sync-fetch.js) supaya hasilnya bisa ditunggu tanpa mengubah Code.gs.
const SYNC_FETCH = path.join(__dirname, 'sync-fetch.js');
const UrlFetchApp = {
  fetch(url, opts) {
    const o = opts || {};
    const req = {
      url: String(url),
      method: String(o.method || 'get').toUpperCase(),
      headers: o.headers || {},
      payload: o.payload === undefined ? undefined : (typeof o.payload === 'string' ? o.payload : JSON.stringify(o.payload)),
      contentType: o.contentType,
    };
    const out = execFileSync(process.execPath, [SYNC_FETCH], {
      input: JSON.stringify(req), maxBuffer: 256 * 1024 * 1024, timeout: 120000,
    });
    const res = JSON.parse(out.toString('utf8'));
    if (res.error) throw new Error('Gagal menghubungi ' + req.url.split('?')[0] + ': ' + res.error);
    if (!o.muteHttpExceptions && res.status >= 400) {
      throw new Error('Request ke ' + req.url.split('?')[0] + ' gagal dengan kode ' + res.status + '. ' + res.body.slice(0, 500));
    }
    return {
      getResponseCode: () => res.status,
      getContentText: () => res.body,
      getHeaders: () => res.headers || {},
    };
  },
};

// ---------------- HtmlService (tidak dipakai; halaman disajikan Express) ----------------
const htmlOutput = {
  evaluate() { return htmlOutput; },
  setTitle() { return htmlOutput; },
  setXFrameOptionsMode() { return htmlOutput; },
  addMetaTag() { return htmlOutput; },
};
const HtmlService = {
  createTemplateFromFile: () => htmlOutput,
  createHtmlOutputFromFile: () => htmlOutput,
  XFrameOptionsMode: { ALLOWALL: 'ALLOWALL', DEFAULT: 'DEFAULT' },
};

const Logger = { log: (...a) => console.log('[gas]', ...a) };

/**
 * Pasang layanan pengganti ke global, jalankan Code.gs & Setup.gs, lalu kembalikan
 * daftar fungsi yang boleh dipanggil dari browser (sama seperti Apps Script: fungsi
 * yang namanya diakhiri "_" bersifat privat).
 */
function loadGas(workbook) {
  const spreadsheetApp = {
    getActiveSpreadsheet: () => workbook,
    getActive: () => workbook,
    flush: () => {},
  };
  Object.assign(globalThis, {
    SpreadsheetApp: spreadsheetApp, Utilities, Session, LockService, CacheService,
    PropertiesService, UrlFetchApp, HtmlService, Logger,
  });

  const before = new Set(Object.getOwnPropertyNames(globalThis));
  for (const file of GAS_FILES) {
    const full = path.join(__dirname, 'gas', file);
    vm.runInThisContext(fs.readFileSync(full, 'utf8'), { filename: full });
  }
  const callable = Object.getOwnPropertyNames(globalThis)
    .filter((n) => !before.has(n) && typeof globalThis[n] === 'function' && !n.endsWith('_') && n !== 'doGet');
  return new Set(callable);
}

module.exports = { loadGas, formatDate };

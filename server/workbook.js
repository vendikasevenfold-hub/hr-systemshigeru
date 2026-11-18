'use strict';
/**
 * Workbook di memori yang meniru bagian SpreadsheetApp yang dipakai Code.gs / Setup.gs.
 *
 * Setiap "sheet" disimpan sebagai array baris (array sel). Nilai sel disimpan apa adanya
 * (string, number, boolean, Date) - TIDAK ada konversi otomatis seperti di Google Sheets
 * (misalnya "2025-01-02" tidak berubah jadi Date). Code.gs sudah menangani kedua bentuk
 * (lihat tglKeStr_, jamKeStr_, toClientText_), jadi perilakunya tetap sama.
 *
 * Perubahan ditandai "dirty" per sheet; server menyimpan sheet yang berubah ke PostgreSQL
 * setelah setiap panggilan selesai (lihat store.js).
 */

const isDate = (v) => Object.prototype.toString.call(v) === '[object Date]';
const isEmpty = (v) => v === '' || v === null || v === undefined;

function normalize(v) {
  if (v === null || v === undefined) return '';
  if (isDate(v)) return new Date(v.getTime());
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v;
  return String(v);
}

function clone(v) {
  return isDate(v) ? new Date(v.getTime()) : v;
}

function pad2(n) { return String(n).padStart(2, '0'); }

/** Teks tampilan sel, meniru getDisplayValues() secukupnya. */
function displayOf(v) {
  if (isEmpty(v)) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isDate(v)) {
    if (isNaN(v.getTime())) return '';
    const hh = pad2(v.getHours()), mi = pad2(v.getMinutes()), ss = pad2(v.getSeconds());
    // Sel jam saja (tanpa tanggal) di Google Sheets bertanggal 30-12-1899.
    if (v.getFullYear() <= 1899) return `${hh}:${mi}`;
    const tgl = `${v.getFullYear()}-${pad2(v.getMonth() + 1)}-${pad2(v.getDate())}`;
    if (hh === '00' && mi === '00' && ss === '00') return tgl;
    return `${tgl} ${hh}:${mi}:${ss}`;
  }
  return String(v);
}

function colToNum(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

class Range {
  constructor(sheet, row, col, numRows, numCols) {
    if (!(row >= 1) || !(col >= 1) || !(numRows >= 1) || !(numCols >= 1)) {
      throw new Error(`Koordinat range tidak valid (${row}, ${col}, ${numRows}, ${numCols}).`);
    }
    this._s = sheet; this._r = row; this._c = col; this._nr = numRows; this._nc = numCols;
  }

  getRow() { return this._r; }
  getColumn() { return this._c; }
  getNumRows() { return this._nr; }
  getNumColumns() { return this._nc; }
  getSheet() { return this._s; }

  getValues() {
    const out = [];
    for (let i = 0; i < this._nr; i++) {
      const row = [];
      for (let j = 0; j < this._nc; j++) row.push(clone(this._s._get(this._r + i, this._c + j)));
      out.push(row);
    }
    return out;
  }

  getDisplayValues() {
    return this.getValues().map((r) => r.map(displayOf));
  }

  getValue() { return clone(this._s._get(this._r, this._c)); }
  getDisplayValue() { return displayOf(this._s._get(this._r, this._c)); }

  setValue(v) {
    for (let i = 0; i < this._nr; i++) {
      for (let j = 0; j < this._nc; j++) this._s._set(this._r + i, this._c + j, v);
    }
    return this;
  }

  setValues(values) {
    if (!Array.isArray(values) || values.length !== this._nr) {
      throw new Error(`Jumlah baris data (${Array.isArray(values) ? values.length : 0}) tidak sama dengan jumlah baris range (${this._nr}).`);
    }
    values.forEach((row, i) => {
      if (!Array.isArray(row) || row.length !== this._nc) {
        throw new Error(`Jumlah kolom data (${Array.isArray(row) ? row.length : 0}) tidak sama dengan jumlah kolom range (${this._nc}).`);
      }
      row.forEach((v, j) => this._s._set(this._r + i, this._c + j, v));
    });
    return this;
  }

  clearContent() { return this.setValue(''); }

  // Format/tampilan tidak berpengaruh di server ini; tetap bisa dirantai seperti di Apps Script.
  setNumberFormat() { return this; }
  setFontWeight() { return this; }
  setBackground() { return this; }
  setFontColor() { return this; }
  setHorizontalAlignment() { return this; }
  setWrap() { return this; }
}

class Sheet {
  constructor(workbook, name, rows) {
    this._wb = workbook;
    this._name = name;
    this._rows = rows || [];
  }

  getName() { return this._name; }
  _dirty() { this._wb._markDirty(this._name); }

  _get(r, c) {
    const row = this._rows[r - 1];
    if (!row) return '';
    const v = row[c - 1];
    return v === undefined ? '' : v;
  }

  _set(r, c, v) {
    while (this._rows.length < r) this._rows.push([]);
    const row = this._rows[r - 1];
    while (row.length < c - 1) row.push('');
    row[c - 1] = normalize(v);
    this._dirty();
  }

  getLastRow() {
    for (let i = this._rows.length - 1; i >= 0; i--) {
      if (this._rows[i] && this._rows[i].some((v) => !isEmpty(v))) return i + 1;
    }
    return 0;
  }

  getLastColumn() {
    let max = 0;
    for (const row of this._rows) {
      if (!row) continue;
      for (let j = row.length - 1; j >= max; j--) {
        if (!isEmpty(row[j])) { max = j + 1; break; }
      }
    }
    return max;
  }

  getMaxRows() { return Math.max(this._rows.length, 1000); }
  getMaxColumns() { return Math.max(this.getLastColumn(), 26); }

  getDataRange() {
    return new Range(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1));
  }

  getRange(a, b, c, d) {
    if (typeof a === 'string') return this._rangeA1(a);
    return new Range(this, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
  }

  _rangeA1(a1) {
    const s = a1.replace(/\$/g, '').trim().toUpperCase();
    let m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(s);
    if (m) {
      const c1 = colToNum(m[1]), r1 = Number(m[2]);
      const c2 = m[3] ? colToNum(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
      return new Range(this, Math.min(r1, r2), Math.min(c1, c2), Math.abs(r2 - r1) + 1, Math.abs(c2 - c1) + 1);
    }
    m = /^([A-Z]+):([A-Z]+)$/.exec(s);
    if (m) {
      const c1 = colToNum(m[1]), c2 = colToNum(m[2]);
      return new Range(this, 1, Math.min(c1, c2), Math.max(this.getLastRow(), 1), Math.abs(c2 - c1) + 1);
    }
    throw new Error('Notasi range tidak dikenali: ' + a1);
  }

  appendRow(values) {
    const last = this.getLastRow();
    this._rows.length = last;
    this._rows.push((values || []).map(normalize));
    this._dirty();
    return this;
  }

  deleteRow(r) {
    if (r >= 1 && r <= this._rows.length) { this._rows.splice(r - 1, 1); this._dirty(); }
    return this;
  }

  deleteRows(start, count) {
    if (start >= 1 && count > 0 && start <= this._rows.length) { this._rows.splice(start - 1, count); this._dirty(); }
    return this;
  }

  insertColumnBefore(c) {
    for (const row of this._rows) {
      if (row && row.length >= c - 1) row.splice(c - 1, 0, '');
    }
    this._dirty();
    return this;
  }

  clear() { this._rows = []; this._dirty(); return this; }

  // Tampilan saja: tidak berpengaruh di server ini.
  setFrozenRows() { return this; }
  setColumnWidths() { return this; }
  setColumnWidth() { return this; }
  autoResizeColumns() { return this; }
}

class Workbook {
  constructor() {
    this._sheets = [];
    this._dirty = new Set();
    this._removed = new Set();
  }

  _markDirty(name) { this._dirty.add(name); }

  /** Dipakai saat memuat dari database (tidak menandai dirty). */
  _load(name, rows) { this._sheets.push(new Sheet(this, name, rows)); }

  getSheetByName(name) {
    return this._sheets.find((s) => s._name === name) || null;
  }

  getSheets() { return this._sheets.slice(); }

  insertSheet(name) {
    if (this.getSheetByName(name)) throw new Error(`Sheet dengan nama "${name}" sudah ada.`);
    const sh = new Sheet(this, name, []);
    this._sheets.push(sh);
    this._removed.delete(name);
    this._markDirty(name);
    return sh;
  }

  deleteSheet(sheet) {
    const i = this._sheets.indexOf(sheet);
    if (i > -1) {
      this._sheets.splice(i, 1);
      this._dirty.delete(sheet._name);
      this._removed.add(sheet._name);
    }
  }

  /** Ambil daftar perubahan sejak penyimpanan terakhir, lalu kosongkan penanda. */
  takeChanges() {
    const upserts = [...this._dirty].map((name) => {
      const sh = this.getSheetByName(name);
      return sh ? { name, position: this._sheets.indexOf(sh), rows: sh._rows } : null;
    }).filter(Boolean);
    const removed = [...this._removed];
    this._dirty.clear();
    this._removed.clear();
    return { upserts, removed };
  }
}

module.exports = { Workbook, Sheet, Range, displayOf, isDate };

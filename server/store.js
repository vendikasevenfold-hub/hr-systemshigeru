'use strict';
/**
 * Penyimpanan workbook di PostgreSQL. Satu baris tabel `sheets` = satu sheet lama dari
 * Google Sheets (nama sheet + seluruh barisnya dalam JSONB). Sel bertipe tanggal disimpan
 * sebagai {"$d": <epoch ms>} supaya kembali menjadi Date saat dimuat.
 */
const { Pool } = require('pg');
const { isDate } = require('./workbook');

const encodeCell = (v) => (isDate(v) ? { $d: v.getTime() } : v);
const decodeCell = (v) => (v && typeof v === 'object' && '$d' in v ? new Date(v.$d) : v);

function encodeRows(rows) {
  return JSON.stringify(rows.map((r) => (r || []).map(encodeCell)));
}

function createStore(databaseUrl) {
  const pool = new Pool({ connectionString: databaseUrl, max: 4 });

  async function init() {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS sheets (
        name        text PRIMARY KEY,
        position    integer NOT NULL DEFAULT 0,
        rows        jsonb   NOT NULL DEFAULT '[]'::jsonb,
        updated_at  timestamptz NOT NULL DEFAULT now()
      )`);
  }

  /** Muat semua sheet ke workbook di memori. Kembalikan jumlah sheet. */
  async function loadInto(workbook) {
    const { rows } = await pool.query('SELECT name, rows FROM sheets ORDER BY position, name');
    for (const r of rows) workbook._load(r.name, r.rows.map((row) => row.map(decodeCell)));
    return rows.length;
  }

  /**
   * Simpan perubahan. `changes` harus sudah di-encode secara sinkron (encodeChanges)
   * sebelum ada await, supaya isinya adalah keadaan tepat saat panggilan selesai.
   */
  async function save(encoded) {
    if (!encoded.upserts.length && !encoded.removed.length) return;
    const client = await pool.connect();
    try {
      await client.query('BEGIN');
      for (const name of encoded.removed) await client.query('DELETE FROM sheets WHERE name = $1', [name]);
      for (const s of encoded.upserts) {
        await client.query(
          `INSERT INTO sheets (name, position, rows, updated_at) VALUES ($1, $2, $3::jsonb, now())
           ON CONFLICT (name) DO UPDATE SET position = EXCLUDED.position, rows = EXCLUDED.rows, updated_at = now()`,
          [s.name, s.position, s.json]);
      }
      await client.query('COMMIT');
    } catch (err) {
      await client.query('ROLLBACK').catch(() => {});
      throw err;
    } finally {
      client.release();
    }
  }

  return { pool, init, loadInto, save };
}

function encodeChanges(changes) {
  return {
    removed: changes.removed,
    upserts: changes.upserts.map((s) => ({ name: s.name, position: s.position, json: encodeRows(s.rows) })),
  };
}

module.exports = { createStore, encodeChanges, encodeRows, decodeCell };

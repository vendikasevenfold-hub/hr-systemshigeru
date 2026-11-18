'use strict';
/**
 * Impor data dari Google Sheets lama ke PostgreSQL.
 *
 * 1. Di Google Sheets database HR: File > Download > Microsoft Excel (.xlsx)
 * 2. Simpan file-nya di folder data-import/ (misalnya data-import/hr.xlsx)
 * 3. Hentikan aplikasi, jalankan impor, lalu nyalakan lagi:
 *      docker compose stop app
 *      docker compose run --rm app node scripts/import-xlsx.js /import/hr.xlsx
 *      docker compose up -d app
 *
 * Setiap sheet di file .xlsx MENGGANTI sheet dengan nama yang sama di database.
 * Sheet di database yang tidak ada di file dibiarkan. Pakai --dry-run untuk melihat
 * ringkasannya saja tanpa menyimpan.
 */
const ExcelJS = require('exceljs');
const { createStore, encodeRows } = require('../server/store');

/** Ubah nilai sel ExcelJS menjadi nilai sel workbook (string/number/boolean/Date). */
function cellValue(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    // ExcelJS membaca tanggal Excel sebagai jam dinding dalam UTC; jadikan jam dinding lokal (TZ).
    return new Date(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate(), v.getUTCHours(), v.getUTCMinutes(), v.getUTCSeconds(), v.getUTCMilliseconds());
  }
  if (typeof v === 'object') {
    if ('result' in v) return cellValue(v.result);                       // rumus -> hasil terakhir
    if ('formula' in v || 'sharedFormula' in v) return '';
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('');
    if ('text' in v) return String(v.text);                              // hyperlink
    if ('error' in v) return String(v.error);
    return String(v);
  }
  return v;
}

async function main() {
  const args = process.argv.slice(2);
  const dryRun = args.includes('--dry-run');
  const file = args.find((a) => !a.startsWith('--'));
  if (!file) {
    console.error('Pemakaian: node scripts/import-xlsx.js <file.xlsx> [--dry-run]');
    process.exit(2);
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);

  const sheets = [];
  wb.eachSheet((ws, id) => {
    const rows = [];
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const vals = [];
      for (let c = 1; c <= ws.columnCount; c++) vals.push(cellValue(row.getCell(c).value));
      while (vals.length && vals[vals.length - 1] === '') vals.pop();
      rows.push(vals);
    }
    while (rows.length && rows[rows.length - 1].length === 0) rows.pop();
    sheets.push({ name: ws.name, position: id, rows });
  });

  console.log('Isi file ' + file + ':');
  for (const s of sheets) console.log(`  - ${s.name}: ${Math.max(s.rows.length - 1, 0)} baris data`);
  if (dryRun) { console.log('--dry-run: tidak ada yang disimpan.'); return; }

  if (!process.env.DATABASE_URL) throw new Error('DATABASE_URL belum diisi.');
  const store = createStore(process.env.DATABASE_URL);
  await store.init();
  await store.save({ removed: [], upserts: sheets.map((s) => ({ name: s.name, position: s.position, json: encodeRows(s.rows) })) });
  await store.pool.end();
  console.log(`Selesai: ${sheets.length} sheet diimpor. Nyalakan ulang aplikasi supaya data baru terbaca.`);
}

main().catch((err) => {
  console.error('Impor gagal:', err.message);
  process.exit(1);
});

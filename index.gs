/**
 * Sinkronisasi isian Rekonsiliasi - Dashboard Panjar PSO Per Cabang
 *
 * Dipasang di spreadsheet BARU (mis. "Rekon PSO"):
 *   Extensions > Apps Script > tempel kode ini > Deploy > New deployment
 *   Type: Web app | Execute as: Me | Who has access: Anyone
 * Salin URL Web App (berakhiran /exec) ke konstanta REKON_URL di Panjar_PSO.html.
 *
 * Sheet yang dibuat otomatis:
 *   "Rekon"          : nilai terbaru tiap isian (satu baris per isian)
 *   "Riwayat Rekon"  : log setiap perubahan (waktu, nama, isian, nilai lama, nilai baru)
 */

// ID spreadsheet "Rekon PSO": bagian di antara /d/ dan /edit pada alamat spreadsheet-nya.
// Wajib diisi bila skrip dibuat dari script.google.com (berdiri sendiri).
// Boleh dikosongkan bila skrip dibuat lewat Extensions > Apps Script di dalam spreadsheet itu.
const SPREADSHEET_ID = '';

const SHEET_DATA = 'Rekon';
const SHEET_LOG = 'Riwayat Rekon';
const HDR_DATA = ['Key', 'Cabang', 'Isian', 'Nilai', 'Diubah oleh', 'Waktu (ms)', 'Waktu'];
const HDR_LOG = ['Waktu', 'Nama', 'Cabang', 'Isian', 'Nilai lama', 'Nilai baru'];

function sheet_(name, header) {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet tidak ditemukan: isi SPREADSHEET_ID di bagian atas kode');
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    if (p.action === 'get') out = getAll_();
    else if (p.action === 'set') out = setOne_(p);
    else out = { ok: false, error: 'aksi tidak dikenal' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  const json = JSON.stringify(out);
  const cb = p.callback || '';
  if (cb && /^[A-Za-z0-9_.]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function getAll_() {
  const sh = sheet_(SHEET_DATA, HDR_DATA);
  const n = sh.getLastRow() - 1;
  const data = {};
  if (n > 0) {
    const vals = sh.getRange(2, 1, n, 6).getValues();
    vals.forEach(function (r) {
      if (!r[0]) return;
      data[r[0]] = {
        v: r[3] === '' || r[3] === null ? null : Number(r[3]),
        by: String(r[4] || ''),
        at: Number(r[5]) || 0
      };
    });
  }
  return { ok: true, data: data };
}

function setOne_(p) {
  const key = String(p.key || '');
  if (!/^[A-Za-z0-9 _|.\-]{1,100}$/.test(key)) return { ok: false, error: 'key tidak valid' };
  let val = '';
  if (p.v !== undefined && String(p.v) !== '') {
    val = Number(p.v);
    if (isNaN(val)) return { ok: false, error: 'nilai bukan angka' };
  }
  const by = String(p.by || '').replace(/[^A-Za-z0-9 .'\-]/g, '').slice(0, 60);
  const base = String(p.base || '');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SHEET_DATA, HDR_DATA);
    const n = sh.getLastRow() - 1;
    let row = -1, cur = null;
    if (n > 0) {
      const keys = sh.getRange(2, 1, n, 1).getValues();
      for (let i = 0; i < n; i++) {
        if (keys[i][0] === key) { row = i + 2; break; }
      }
      if (row > 0) {
        const r = sh.getRange(row, 1, 1, 6).getValues()[0];
        cur = { v: r[3] === '' || r[3] === null ? null : Number(r[3]), by: String(r[4] || ''), at: Number(r[5]) || 0 };
      }
    }
    // Ada orang lain yang sudah mengubah isian ini sejak perangkat ini terakhir membacanya
    if (cur && String(cur.at) !== base) return { ok: false, conflict: true, cur: cur };

    const now = Math.max(Date.now(), cur ? cur.at + 1 : 0); // waktu selalu naik, supaya deteksi bentrok tidak meleset
    const parts = key.split('|');
    const cab = parts[0], isian = parts.slice(1).join(' ');
    const rowVals = [key, cab, isian, val, by, now, Utilities.formatDate(new Date(now), 'Asia/Jakarta', 'dd-MMM-yyyy HH:mm:ss')];
    if (row > 0) sh.getRange(row, 1, 1, 7).setValues([rowVals]);
    else sh.appendRow(rowVals);

    const log = sheet_(SHEET_LOG, HDR_LOG);
    log.appendRow([new Date(now), by, cab, isian, cur && cur.v !== null ? cur.v : '', val]);
    return { ok: true, at: now };
  } finally {
    lock.releaseLock();
  }
}

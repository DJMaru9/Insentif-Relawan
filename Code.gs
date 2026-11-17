/***** BACKEND INSENTIF RELAWAN SPPG — Google Apps Script *****
 * 1. Ganti PIN di bawah (dipakai semua perangkat untuk masuk).
 * 2. Deploy > New deployment > Web app
 *    Execute as: Me   |   Who has access: Anyone
 * 3. Salin URL Web App ke aplikasi (API_URL di index.html).
 ****************************************************************/
const PIN = '1234';

const HEAD = {
  Relawan:    ['id', 'nama', 'posisi', 'tarif', 'aktif'],
  Kehadiran:  ['tanggal', 'relawan_id', 'nama', 'nilai', 'diubah'],
  Pengaturan: ['kunci', 'nilai']
};

function doGet(e)  { return out(handle({ action: e.parameter.action || 'get', pin: e.parameter.pin })); }
function doPost(e) {
  let p; try { p = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, err: 'Data tidak valid' }); }
  return out(handle(p));
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle(p) {
  if (String(p.pin) !== String(PIN)) return { ok: false, err: 'PIN salah' };
  if (p.action === 'ping') return { ok: true };
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (p.action === 'ops') applyOps(p.ops || []);
    return { ok: true, data: readAll() };
  } catch (err) {
    return { ok: false, err: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- util sheet ---------- */
function sh(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]).setFontWeight('bold').setBackground('#0b3d91').setFontColor('#ffffff');
    s.setFrozenRows(1);
    s.getRange('A:B').setNumberFormat('@'); // tanggal & id disimpan sebagai teks
  }
  return s;
}
function rows(s) { const n = s.getLastRow(); return n < 2 ? [] : s.getRange(2, 1, n - 1, s.getLastColumn()).getValues(); }
function tgl(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v); }
function setup() { Object.keys(HEAD).forEach(sh); }   // opsional: jalankan sekali untuk membuat sheet

/* ---------- baca semua ---------- */
function readAll() {
  const relawan = rows(sh('Relawan')).filter(r => r[0]).map(r => ({
    id: String(r[0]), nama: r[1], posisi: r[2], tarif: Number(r[3]) || 0, aktif: !(r[4] === false || String(r[4]).toUpperCase() === 'FALSE')
  }));
  const hadir = {};
  rows(sh('Kehadiran')).forEach(r => {
    const v = Number(r[3]); if (!r[0] || !v) return;
    const d = tgl(r[0]); (hadir[d] = hadir[d] || {})[String(r[1])] = v;
  });
  const set = {};
  rows(sh('Pengaturan')).forEach(r => { if (r[0]) set[r[0]] = r[0] === 'tarif' ? Number(r[1]) || 0 : String(r[1]); });
  return { set, relawan, hadir };
}

/* ---------- terapkan perubahan ---------- */
function applyOps(ops) {
  const sR = sh('Relawan'), sK = sh('Kehadiran');
  // indeks relawan
  const rR = rows(sR), idxR = {}; rR.forEach((r, i) => idxR[String(r[0])] = i + 2);
  const nama = {}; rR.forEach(r => nama[String(r[0])] = r[1]);
  // indeks kehadiran
  const rK = rows(sK), idxK = {}; rK.forEach((r, i) => idxK[tgl(r[0]) + '|' + r[1]] = i + 2);
  const now = new Date();
  const appendK = [];
  const setHadir = (d, id, v) => {
    const k = d + '|' + id, row = idxK[k];
    if (row === -1) { const a = appendK.find(x => x[0] === d && x[1] === id); if (a) { a[3] = v || 0; a[4] = now; } return; }
    if (row) { sK.getRange(row, 3, 1, 3).setValues([[nama[id] || '', v || 0, now]]); return; }
    if (v) { appendK.push([d, id, nama[id] || '', v, now]); idxK[k] = -1; }
  };

  ops.forEach(o => {
    if (o.op === 'hadir') setHadir(o.d, String(o.id), Number(o.v) || 0);
    else if (o.op === 'hadirBulk') (o.items || []).forEach(it => setHadir(o.d, String(it.id), Number(it.v) || 0));
    else if (o.op === 'relawan') {
      const r = o.r, vals = [[String(r.id), r.nama, r.posisi || '', Number(r.tarif) || 0, r.aktif !== false]];
      nama[String(r.id)] = r.nama;
      if (idxR[r.id]) sR.getRange(idxR[r.id], 1, 1, 5).setValues(vals);
      else { sR.appendRow(vals[0]); idxR[r.id] = sR.getLastRow(); }
    }
    else if (o.op === 'delRelawan') {
      const row = idxR[o.id];
      if (row) {
        sR.deleteRow(row); delete idxR[o.id];
        Object.keys(idxR).forEach(k => { if (idxR[k] > row) idxR[k]--; });
        Object.keys(idxK).forEach(k => { if (k.split('|')[1] === String(o.id) && idxK[k] > 0) sK.getRange(idxK[k], 4).setValue(0); });
        appendK.forEach(a => { if (a[1] === String(o.id)) a[3] = 0; });
      }
    }
    else if (o.op === 'set') {
      const s = sh('Pengaturan'); if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).clearContent();
      const v = Object.keys(o.set).map(k => [k, o.set[k]]);
      if (v.length) s.getRange(2, 1, v.length, 2).setValues(v);
    }
    else if (o.op === 'restore') {
      flushK(); appendK.length = 0;
      const st = o.state || {};
      [sR, sK].forEach(s => { if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent(); });
      const rl = (st.relawan || []).map(r => [String(r.id), r.nama, r.posisi || '', Number(r.tarif) || 0, r.aktif !== false]);
      if (rl.length) sR.getRange(2, 1, rl.length, 5).setValues(rl);
      Object.keys(idxR).forEach(k => delete idxR[k]); rl.forEach((r, i) => { idxR[r[0]] = i + 2; nama[r[0]] = r[1]; });
      const kh = []; Object.keys(st.hadir || {}).forEach(d => Object.keys(st.hadir[d]).forEach(id => { if (st.hadir[d][id]) kh.push([d, id, nama[id] || '', st.hadir[d][id], now]); }));
      if (kh.length) sK.getRange(2, 1, kh.length, 5).setValues(kh);
      Object.keys(idxK).forEach(k => delete idxK[k]); kh.forEach((r, i) => idxK[r[0] + '|' + r[1]] = i + 2);
      if (st.set) applyOps([{ op: 'set', set: st.set }]);
    }
  });
  function flushK() {
    if (!appendK.length) return;
    const start = sK.getLastRow() + 1;
    sK.getRange(start, 1, appendK.length, 5).setValues(appendK);
    appendK.forEach((r, i) => idxK[r[0] + '|' + r[1]] = start + i);
  }
  flushK();
}

// API Keuangan Pribadi. Tempel di Apps Script, jalankan setup() sekali, lalu Deploy sebagai Web app.
const SS_ID = ''; // kosongkan jika script dibuat dari menu Extensions > Apps Script di dalam spreadsheet
const SHEETS = {'Transaksi':'TRX','Rencana Liburan':'LBR','Rencana Program':'PRG','Rencana Pembelian':'BLI'};
const CALC = ['Selisih (Rp)','Sisa Anggaran (Rp)','Pokok Pinjaman (Rp)','Cicilan per Bulan (Rp)','Total Bayar (Rp)','Total Bunga (Rp)'];
const P = PropertiesService.getScriptProperties();
const ss = () => SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet();
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const hash = (s, salt) => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + s));

function setup() { ss(); DriveApp.getRootFolder(); accSheet(); }  // untuk memberi izin akses
function doGet() { return out({ok: true, msg: 'API Keuangan aktif'}); }

function doPost(e) {
  const l = LockService.getScriptLock(); l.waitLock(20000);
  try { return out(handle(JSON.parse(e.postData.contents))); }
  catch (err) { return out({ok: false, error: String(err.message || err)}); }
  finally { l.releaseLock(); }
}

function handle(d) {
  if (d.action === 'login' || d.action === 'register') return auth(d);
  const t = P.getProperty('t_' + d.token), s = t && JSON.parse(t);
  if (!s || s.e < Date.now()) throw new Error('Sesi habis, silakan masuk lagi');
  switch (d.action) {
    case 'list': return {ok: true, data: listAll()};
    case 'save': return save(d);
    case 'delete': return del(d);
    case 'upload': return upload(d);
    case 'logout': P.deleteProperty('t_' + d.token); return {ok: true};
  }
  throw new Error('Aksi tidak dikenal');
}

function accSheet() {
  let s = ss().getSheetByName('Akun');
  if (!s) { s = ss().insertSheet('Akun'); s.appendRow(['Username', 'Salt', 'Hash', 'Dibuat']); }
  return s;
}

function auth(d) {
  const s = accSheet(), rows = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 3).getValues() : [];
  const u = String(d.username || '').trim().toLowerCase(), pw = String(d.password || '');
  if (!u || pw.length < 6) throw new Error('Isi username dan password (minimal 6 karakter)');
  if (d.action === 'register') {
    // akun pertama bebas; akun berikutnya butuh KODE_DAFTAR (Project Settings > Script properties)
    if (rows.length && d.code !== P.getProperty('KODE_DAFTAR')) throw new Error('Kode daftar salah atau pendaftaran ditutup');
    if (rows.some(r => String(r[0]) === u)) throw new Error('Username sudah dipakai');
    const salt = Utilities.getUuid(); s.appendRow([u, salt, hash(pw, salt), new Date()]);
  } else {
    const r = rows.find(r => String(r[0]) === u);
    if (!r || r[2] !== hash(pw, r[1])) throw new Error('Username atau password salah');
  }
  const all = P.getProperties();
  for (const k in all) if (k.indexOf('t_') === 0 && JSON.parse(all[k]).e < Date.now()) P.deleteProperty(k);
  const t = Utilities.getUuid() + Utilities.getUuid();
  P.setProperty('t_' + t, JSON.stringify({u: u, e: Date.now() + 30 * 864e5}));
  return {ok: true, token: t, user: u};
}

function tbl(n) { const s = ss().getSheetByName(n); if (!s || !SHEETS[n]) throw new Error('Sheet tidak ditemukan: ' + n); return s; }
function ids(s) { return s.getRange(2, 1, s.getMaxRows() - 1, 1).getValues().map(r => String(r[0])); }

function listAll() {
  const tz = ss().getSpreadsheetTimeZone(), o = {};
  for (const n in SHEETS) {
    const v = tbl(n).getDataRange().getValues(), h = v[0];
    o[n] = {rows: v.slice(1).filter(r => r[0] !== '').map(r => {
      const x = {}; h.forEach((k, i) => { let c = r[i]; if (c instanceof Date) c = Utilities.formatDate(c, tz, 'yyyy-MM-dd'); x[k] = c; }); return x; })};
  }
  return o;
}

function save(d) {
  const s = tbl(d.sheet), h = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0], list = ids(s);
  let id = d.row.ID, i;
  if (id) { i = list.indexOf(String(id)); if (i < 0) throw new Error('Data tidak ditemukan'); }
  else {
    const nums = list.filter(x => x).map(x => parseInt(x.split('-')[1]) || 0);
    id = SHEETS[d.sheet] + '-' + ('00' + ((nums.length ? Math.max.apply(null, nums) : 0) + 1)).slice(-3);
    i = list.indexOf(''); if (i < 0) throw new Error('Baris di sheet sudah penuh, tambahkan baris kosong');
  }
  d.row.ID = id;
  h.forEach((k, c) => {
    if (CALC.indexOf(k) >= 0 || d.row[k] === undefined) return;
    let v = d.row[k] === null ? '' : d.row[k];
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) v = Utilities.parseDate(v, ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    s.getRange(i + 2, c + 1).setValue(v);
  });
  return {ok: true, id: id};
}

function del(d) {
  const s = tbl(d.sheet), i = ids(s).indexOf(String(d.id));
  if (i < 0) throw new Error('Data tidak ditemukan');
  s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].forEach((k, c) => { if (CALC.indexOf(k) < 0) s.getRange(i + 2, c + 1).clearContent(); });
  return {ok: true};
}

function upload(d) {
  const it = DriveApp.getFoldersByName('Keuangan Pribadi - Gambar');
  const f = it.hasNext() ? it.next() : DriveApp.createFolder('Keuangan Pribadi - Gambar');
  const file = f.createFile(Utilities.newBlob(Utilities.base64Decode(d.data), 'image/jpeg', 'img_' + Date.now() + '.jpg'));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {ok: true, url: 'https://drive.google.com/file/d/' + file.getId() + '/view'};
}

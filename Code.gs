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
  try {
    const d = JSON.parse(e.postData.contents), w = ['save', 'delete', 'register'].indexOf(d.action) >= 0, l = LockService.getScriptLock();
    if (w && !l.tryLock(25000)) throw new Error('Server sibuk, coba lagi sebentar');
    try { return out(handle(d)); } finally { if (w) l.releaseLock(); }
  } catch (err) { return out({ok: false, error: String(err.message || err)}); }
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

function rowObj(h, r, tz) {
  const x = {};
  h.forEach((k, i) => { let c = r[i]; if (c instanceof Date) c = Utilities.formatDate(c, tz, 'yyyy-MM-dd'); x[k] = c; });
  return x;
}

function listAll() {
  const tz = ss().getSpreadsheetTimeZone(), o = {};
  for (const n in SHEETS) {
    const s = tbl(n), list = ids(s); let last = list.length;
    while (last > 0 && !list[last - 1]) last--;
    const v = s.getRange(1, 1, last + 1, s.getLastColumn()).getValues(), h = v[0];
    o[n] = {rows: v.slice(1).filter(r => r[0] !== '').map(r => rowObj(h, r, tz))};
  }
  return o;
}

function save(d) {
  const cache = CacheService.getScriptCache(), rid = d.rid ? 'r_' + d.rid : '';
  if (rid && cache.get(rid)) return JSON.parse(cache.get(rid));  // permintaan ulang: jangan tulis dua kali
  const s = tbl(d.sheet), list = ids(s), ncol = s.getLastColumn(), tz = ss().getSpreadsheetTimeZone();
  let id = d.row.ID, i;
  if (id) { i = list.indexOf(String(id)); if (i < 0) throw new Error('Data tidak ditemukan'); }
  else {
    const nums = list.filter(x => x).map(x => parseInt(x.split('-')[1]) || 0);
    id = SHEETS[d.sheet] + '-' + ('00' + ((nums.length ? Math.max.apply(null, nums) : 0) + 1)).slice(-3);
    i = list.indexOf(''); if (i < 0) throw new Error('Baris di sheet sudah penuh, tambahkan baris kosong');
  }
  d.row.ID = id;
  const rg = s.getRange(i + 2, 1, 1, ncol), fm = rg.getFormulas()[0], old = rg.getValues()[0];
  const h = s.getRange(1, 1, 1, ncol).getValues()[0];
  rg.setValues([h.map((k, c) => {
    if (fm[c]) return fm[c];  // kolom rumus dibiarkan
    let v = d.row[k]; if (v === undefined) return old[c];
    if (v === null) return '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) v = Utilities.parseDate(v, tz, 'yyyy-MM-dd');
    return v;
  })]);
  SpreadsheetApp.flush();
  const res = {ok: true, id: id, row: rowObj(h, rg.getValues()[0], tz)};
  if (rid) cache.put(rid, JSON.stringify(res), 600);
  return res;
}

function del(d) {
  const s = tbl(d.sheet), i = ids(s).indexOf(String(d.id));
  if (i < 0) throw new Error('Data tidak ditemukan');
  const rg = s.getRange(i + 2, 1, 1, s.getLastColumn());
  rg.setValues([rg.getFormulas()[0].map(f => f || '')]);
  return {ok: true};
}

function upload(d) {
  const it = DriveApp.getFoldersByName('Keuangan Pribadi - Gambar');
  const f = it.hasNext() ? it.next() : DriveApp.createFolder('Keuangan Pribadi - Gambar');
  const file = f.createFile(Utilities.newBlob(Utilities.base64Decode(d.data), 'image/jpeg', 'img_' + Date.now() + '.jpg'));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {ok: true, url: 'https://drive.google.com/file/d/' + file.getId() + '/view'};
}

// Jalankan dari editor Apps Script: menambah / memperbaiki baris "sisa belum lunas" di sheet Ringkasan (aman dijalankan ulang).
function tambahRingkasan() {
  const s = ss().getSheetByName('Ringkasan'), L = "'Rencana Liburan'!", P = "'Rencana Program'!", RP = '"Rp" #,##0;("Rp" #,##0);-';
  const find = t => s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => r[0]).indexOf(t) + 1;
  const rg = (sh, c) => sh + '$' + c + '$2:$' + c + '$200';
  const sisa = (sh, tot, real, st, v) => '=MAX(0,SUMIFS(' + rg(sh, tot) + ',' + rg(sh, st) + ',"<>' + v + '")-SUMIFS(' + rg(sh, real) + ',' + rg(sh, st) + ',"<>' + v + '"))';
  const hitung = (sh, st, v) => '=COUNTIFS(' + rg(sh, 'A') + ',"<>",' + rg(sh, st) + ',"<>' + v + '")';
  [['Total Biaya Aktual', [
      ['Sisa belum lunas (Rp)', sisa(L, 'G', 'H', 'J', 'Lunas'), RP, 'Estimasi - Aktual, hanya item yang statusnya bukan Lunas. Isi Aktual dengan jumlah yang sudah dibayar (mis. DP).'],
      ['Item liburan belum lunas', hitung(L, 'J', 'Lunas'), '0', '']]],
   ['Sisa Anggaran', [
      ['Sisa belum terealisasi (Rp)', sisa(P, 'G', 'H', 'K', 'Selesai'), RP, 'Anggaran - Realisasi, hanya program yang statusnya bukan Selesai.'],
      ['Program belum selesai', hitung(P, 'K', 'Selesai'), '0', '']]]
  ].forEach(([after, rows]) => {
    let prev = find(after);
    rows.forEach(x => {
      let r = find(x[0]);
      if (!r) { s.insertRowAfter(prev); r = prev + 1; s.getRange(r, 1).setValue(x[0]); }
      prev = r;
      s.getRange(r, 1, 1, 3).setFontFamily('Arial').setFontSize(10).setFontWeight('normal').setBackground(null);
      s.getRange(r, 2).setFormula(x[1]).setNumberFormat(x[2]).setBorder(true, true, true, true, false, false, '#BFBFBF', SpreadsheetApp.BorderStyle.SOLID);
      s.getRange(r, 3).setValue(x[3]).setFontStyle('italic').setFontColor('#5E7280');
    });
  });
}

// Jalankan SEKALI: Ringkasan memakai aturan baru. Pengeluaran dengan Metode "Tabungan" mengurangi tabungan, bukan kas.
function perbaruiRingkasan() {
  const s = ss().getSheetByName('Ringkasan'), t = ss().getSheetByName('Transaksi');
  const R = c => 'Transaksi!$' + c + '$2:$' + c + '$1000', mon = R('B') + ',">="&$B$3,' + R('B') + ',"<"&EDATE($B$3,1)', tab = R('G') + ',"Tabungan"';
  const sf = (tipe, ex) => 'SUMIFS(' + R('F') + ',' + R('C') + ',"' + tipe + '"' + (ex ? ',' + ex : '') + ')';
  const row = l => s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => r[0]).indexOf(l) + 1;
  const set = (l, f, note) => {
    const r = row(l); if (!r) throw new Error('Baris tidak ditemukan: ' + l);
    s.getRange(r, 2).setFormula(f); s.getRange(r, 3).setValue(note).setFontStyle('italic').setFontColor('#5E7280').setFontFamily('Arial').setFontSize(10);
  };
  set('Sisa bulan ini (Pemasukan - Pengeluaran - Tabungan)', '=B6-(B7-' + sf('Pengeluaran', tab + ',' + mon) + ')-B8', 'Pengeluaran bermetode Tabungan tidak mengurangi kas.');
  set('Total Tabungan terkumpul', '=' + sf('Tabungan') + '-' + sf('Pengeluaran', tab), 'Setoran dikurangi pengeluaran bermetode Tabungan.');
  set('Saldo kas (Pemasukan - Pengeluaran - Tabungan)', '=' + sf('Pemasukan') + '-(' + sf('Pengeluaran') + '-' + sf('Pengeluaran', tab) + ')-' + sf('Tabungan'), 'Tabungan yang dipakai tidak mengurangi kas.');
  t.getRange('G2:G1000').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Tunai', 'Transfer Bank', 'E-Wallet', 'QRIS', 'Kartu Debit', 'Kartu Kredit', 'Tabungan'], true).build());
}

// Jalankan SEKALI: menambah kategori "Elektronik" dan metode "QRIS" ke pilihan (dropdown) di sheet Transaksi.
function perbaruiDaftar() {
  const t = ss().getSheetByName('Transaksi'), dv = l => SpreadsheetApp.newDataValidation().requireValueInList(l, true).build();
  t.getRange('D2:D1000').setDataValidation(dv(['Gaji', 'Bonus', 'Makan', 'Transport', 'Tagihan', 'Belanja', 'Elektronik', 'Hiburan', 'Kesehatan', 'Pendidikan', 'Tabungan', 'Lainnya']));
  t.getRange('G2:G1000').setDataValidation(dv(['Tunai', 'Transfer Bank', 'E-Wallet', 'QRIS', 'Kartu Debit', 'Kartu Kredit', 'Tabungan']));
}

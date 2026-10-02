/**
 * Starlite Reward Q4 — Google Apps Script backend
 * Menempel di Google Sheet "Reward Sales / IKR" (tab "Data").
 * Web app membaca & menulis langsung ke sheet ini (sync 2 arah).
 *
 *  Kolom tab Data (10 kolom):
 *  A NAMA MITRA | B NAMA IKR | C NAMA SALES | D STASIUN | E NAMA ACCOUNT | F NO HP
 *  G ALAMAT EMAIL | H POSISI | I TANGGAL INPUT | J ID (internal) | K IKR ID / SALES ID
 *  (khusus pendaftaran peserta; tidak ada tier / pencapaian)
 *
 * Tab Mitra & Stasiun: daftar pilihan dropdown (kolom A). Admin boleh edit langsung;
 * mitra juga bisa menambah nama baru lewat form (tombol "+ Tambahkan").
 */

const SHEET_DATA = 'Data';
const SHEET_MITRA = 'Mitra';
const SHEET_STASIUN = 'Stasiun';
const SHEET_FINAL = 'Hasil Final';
const SHEET_ACH = 'Ach sales ikr';   // data pencapaian (1 baris = 1 pelanggan): registration_date, active_date, customer_code, mitra, station, ikr_id, ikr_name, sales, sales_name, ...
const HEADERS = ['NAMA MITRA','NAMA IKR','NAMA SALES','STASIUN','NAMA ACCOUNT','NO HP','ALAMAT EMAIL','POSISI','TANGGAL INPUT','ID','IKR ID / SALES ID'];
const C = {MITRA:0, IKR:1, SALES:2, STASIUN:3, AKUN:4, HP:5, EMAIL:6, POSISI:7, TGL:8, ID:9, UID:10};

/* ---------- Web app ---------- */
function doGet() {
  setup_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Starlite Reward')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** API untuk frontend yang di-host di luar Apps Script (mis. GitHub Pages). Body: {fn, args:[...]} (Content-Type text/plain). */
function doPost(e) {
  const ALLOWED = { getData: getData, getDigest: getDigest, submitPeserta: submitPeserta, updatePeserta: updatePeserta,
    addMitra: addMitra, addStasiun: addStasiun, deletePeserta: deletePeserta, checkAdminPin: checkAdminPin, saveFinal: saveFinal, getFinal: getFinal };
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!ALLOWED.hasOwnProperty(req.fn)) throw new Error('Fungsi tidak dikenal: ' + req.fn);
    const cch = CacheService.getScriptCache();
    if (!cch.get('setup_ok')) { setup_(); cch.put('setup_ok', '1', 21600); }   // setup hanya sesekali (lebih cepat)
    out = { ok: true, result: ALLOWED[req.fn].apply(null, req.args || []) };
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Jalankan sekali dari editor (pilih fungsi setup → Run) untuk menyiapkan kolom & izin. */
function setup() { setup_(); }

function setup_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_DATA);
  if (!sh) sh = ss.insertSheet(SHEET_DATA);
  const cur = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (cur.join('|') !== HEADERS.join('|')) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('F:F').setNumberFormat('@');           // No HP tetap teks (0 di depan tidak hilang)
    sh.getRange('I:I').setNumberFormat('dd/MM/yyyy');
  }
  if (!ss.getSheetByName(SHEET_MITRA)) {
    const m = ss.insertSheet(SHEET_MITRA);
    m.getRange(1, 1).setValue('NAMA MITRA').setFontWeight('bold');
    m.setFrozenRows(1);
  }
  if (!ss.getSheetByName(SHEET_STASIUN)) {
    const s = ss.insertSheet(SHEET_STASIUN);
    s.getRange(1, 1).setValue('NAMA STASIUN').setFontWeight('bold');
    s.setFrozenRows(1);
  }
  seedMitra_();
  seedStasiun_();
  if (!PropertiesService.getScriptProperties().getProperty('ADMIN_PIN')) {
    PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', '2026');
  }
}

/** Isi tab Mitra dengan daftar resmi (sekali saja; setelah itu tab Mitra bebas diedit admin). */
function seedMitra_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('MITRA_SEEDED_V1')) return;
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_MITRA);
  const last = sh.getLastRow();
  const have = new Set((last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : []).map(r => String(r[0]).trim().toUpperCase()));
  const add = MITRA_LIST.filter(n => !have.has(n.toUpperCase())).map(n => [n]);
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 1).setValues(add);
  sh.autoResizeColumn(1);
  props.setProperty('MITRA_SEEDED_V1', new Date().toISOString());
}

/** Isi tab Stasiun dengan daftar resmi + stasiun yang sudah pernah diinput di tab Data (sekali saja). */
function seedStasiun_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('STASIUN_SEEDED_V1')) return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_STASIUN);
  const last = sh.getLastRow();
  const have = new Set((last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : []).map(r => normStasiun_(r[0])));
  const data = ss.getSheetByName(SHEET_DATA);
  const dl = data.getLastRow();
  const existing = dl > 1 ? data.getRange(2, C.STASIUN + 1, dl - 1, 1).getValues().map(r => normStasiun_(r[0])) : [];
  const add = [];
  STASIUN_LIST.concat(existing).forEach(n => { n = normStasiun_(n); if (n && !have.has(n)) { have.add(n); add.push([n]); } });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 1).setValues(add);
  sh.autoResizeColumn(1);
  props.setProperty('STASIUN_SEEDED_V1', new Date().toISOString());
}

/* ---------- Read ---------- */
/** Cache besar (dipecah per 90 KB karena batas CacheService 100 KB per kunci). */
function putBig_(key, obj, ttl) {
  try {
    const cache = CacheService.getScriptCache(), s = JSON.stringify(obj), size = 90000, parts = {};
    let n = 0; for (let i = 0; i < s.length; i += size) parts[key + '_' + (n++)] = s.slice(i, i + size);
    parts[key + '_n'] = String(n);
    cache.putAll(parts, ttl);
  } catch (e) {}
}
function getBig_(key) {
  try {
    const cache = CacheService.getScriptCache(), n = Number(cache.get(key + '_n'));
    if (!(n > 0)) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    const got = cache.getAll(keys); let s = '';
    for (let i = 0; i < n; i++) { if (got[key + '_' + i] == null) return null; s += got[key + '_' + i]; }
    return JSON.parse(s);
  } catch (e) { return null; }
}

function readValues_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);
  const last = sh.getLastRow();
  let values = last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
  // Baris yang diisi manual di sheet belum punya ID → beri ID (kunci hanya dipakai saat menulis)
  if (values.some(r => !String(r[C.ID]).trim() && rowHasData_(r))) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      values = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), HEADERS.length).getValues();
      values.forEach(r => { if (!String(r[C.ID]).trim() && rowHasData_(r)) r[C.ID] = Utilities.getUuid().slice(0, 8); });
      sh.getRange(2, C.ID + 1, values.length, 1).setValues(values.map(r => [r[C.ID]]));
    } finally { lock.releaseLock(); }
  }
  return values;
}

/** Data yang boleh dilihat Mitra (di-cache): daftar mitra/stasiun, pencapaian, dan peserta terdaftar versi ringkas (tanpa HP/email/akun). */
function buildPub_(values) {
  const ss = SpreadsheetApp.getActive();
  const col = name => { const s = ss.getSheetByName(name), n = s.getLastRow(); return n > 1 ? s.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0]).trim()).filter(String) : []; };
  const regs = values.filter(rowHasData_).map(r => { const o = toObj_(r); return { peran: o.peran, uid: o.uid, nama: o.nama, mitra: o.mitra, stasiun: o.stasiun }; }).filter(x => x.uid);
  return { mitra: col(SHEET_MITRA), stasiun: col(SHEET_STASIUN), ach: achData_(), regs: regs, at: new Date().toISOString() };
}

/** Mitra (tanpa password) menerima data publik dari cache (cepat); Admin (password benar) juga menerima seluruh data peserta. */
function getData(pin) {
  const isAdmin = checkPin_(pin);
  let pub = getBig_('pub'), values = null;
  if (!pub || isAdmin) values = readValues_();
  if (!pub) { pub = buildPub_(values); storePub_(pub); }
  const payload = { rows: isAdmin ? values.filter(rowHasData_).map(toObj_) : [], mitra: pub.mitra, stasiun: pub.stasiun, admin: isAdmin, ach: pub.ach, regs: pub.regs, pubAt: pub.at || '' };
  payload.digest = getDigest();   // harus sama dengan nilai yang dipakai pengecekan berkala, kalau tidak web memuat ulang terus
  payload.at = new Date().toISOString();
  return payload;
}

/** Dipanggil tiap beberapa detik oleh web: hanya mengembalikan sidik data untuk deteksi perubahan. */
function getDigest() {
  const ss = SpreadsheetApp.getActive();
  const lr = n => { const s = ss.getSheetByName(n); return s ? s.getLastRow() : 0; };
  const a = ss.getSheetByName(SHEET_ACH);
  const v = PropertiesService.getScriptProperties().getProperty('DATA_V') || '';
  // Ringan: tanpa membaca isi sheet. PUB_H = sidik isi data publik, diperbarui oleh refreshCache_ (tiap menit) sehingga web hanya memuat ulang kalau data benar-benar berubah.
  return digest_(JSON.stringify([lr(SHEET_DATA), lr(SHEET_MITRA), lr(SHEET_STASIUN), a ? [a.getLastRow(), a.getLastColumn()] : [0, 0], v, PropertiesService.getScriptProperties().getProperty('PUB_H') || '']));
}

/* ---------- Pencapaian (tab "Ach sales ikr") ---------- */
/** Bulan Q4 2026: Okt=1, Nov=2, Des=3; di luar itu 0. Menerima Date atau teks 'YYYY-MM-DD' / 'DD/MM/YYYY'. */
function achMonth_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  let y, m;
  if (v instanceof Date) { y = v.getFullYear(); m = v.getMonth() + 1; }
  else {
    const s = String(v).trim();
    let t = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (t) { y = +t[1]; m = +t[2]; }
    else { t = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (!t) return 0; y = +t[3]; m = +t[2]; }
  }
  return (y === 2026 && m >= 10 && m <= 12) ? m - 9 : 0;
}
/** Tanggal Q4 2026 sebagai 'MM-DD' (mis. '10-02'); selain itu ''. */
function achDay_(v) {
  if (!achMonth_(v)) return '';
  let m, d;
  if (v instanceof Date) { m = v.getMonth() + 1; d = v.getDate(); }
  else {
    const s = String(v).trim();
    let t = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (t) { m = +t[2]; d = +t[3]; } else { t = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); m = +t[2]; d = +t[1]; }
  }
  return (m < 10 ? '0' : '') + m + '-' + (d < 10 ? '0' : '') + d;
}
function achId_(v) { return String(v === null || v === undefined ? '' : v).trim().replace(/\.0+$/, ''); }
function achKey_(v) { return achId_(v).toLowerCase().replace(/\s+/g, ''); }

/**
 * Rekap per IKR ID dan per Sales ID (hanya tanggal di Q4 2026 yang dihitung; 1 customer_code dihitung sekali per ID).
 * IKR: aktivasi = ada active_date. Sales: registrasi = ada registration_date; aktivasi = ada registration_date dan active_date.
 */
function achData_() {
  const hit = getBig_('ach');
  if (hit) return hit;
  const res = achCompute_();
  putBig_('ach', res, 120);   // cache singkat 30 detik; juga dihapus otomatis tiap ada perubahan di sheet
  return res;
}
function achCompute_() {
  const out = { ikr: [], sales: [], found: false, rows: 0 };
  const warn = { noIdIkr: 0, noIdSales: 0, noCust: 0, dup: [] };
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ACH);
  if (!sh || sh.getLastRow() < 2) return out;
  out.found = true;
  const lc = sh.getLastColumn();
  const rawHead = sh.getRange(1, 1, 1, lc).getValues()[0].map(x => String(x).trim());
  const head = rawHead.map(x => x.toLowerCase().replace(/[\s\-]+/g, '_'));
  // cari kolom berdasarkan nama header (beberapa nama alternatif diterima)
  const ix = names => { for (const n of names) { const i = head.indexOf(n); if (i >= 0) return i; } return -1; };
  const I = {
    reg: ix(['registration_date', 'registrasi_date', 'tgl_registrasi', 'registration']),
    act: ix(['active_date', 'activation_date', 'aktif_date', 'tgl_aktivasi']),
    cust: ix(['customer_code', 'customer_id', 'cust_code']),
    mitra: ix(['mitra']), st: ix(['station', 'stasiun']), reg2: ix(['region']),
    ikrId: ix(['ikr_number', 'ikr_no', 'ikr_id', 'id_ikr', 'ikrid']), ikrName: ix(['ikr_name', 'nama_ikr']), ikrMail: ix(['ikr_email']),
    sal: ix(['sales_number', 'sales_no', 'sales_id', 'sales', 'id_sales', 'salesid', 'sales_code', 'id_sales_kit']), salName: ix(['sales_name', 'nama_sales']), salMail: ix(['sales_email'])
  };
  if (I.sal < 0) I.sal = head.findIndex(h => /sales/.test(h) && /id|code|kode|number|no/.test(h) && !/name|nama|email|mail/.test(h));
  if (I.ikrId < 0) I.ikrId = head.findIndex(h => /ikr/.test(h) && /id|code|kode|number|no/.test(h) && !/name|nama|email|mail/.test(h));
  out.headers = rawHead;
  out.missing = Object.keys(I).filter(k => I[k] < 0 && ['reg', 'act', 'ikrId', 'sal'].indexOf(k) >= 0).map(k => ({ reg: 'registration_date', act: 'active_date', ikrId: 'ikr_number', sal: 'sales_number' }[k]));
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, lc).getValues();
  out.rows = vals.length;
  const at = (r, i) => i < 0 ? '' : r[i];
  const maps = { ikr: {}, sales: {} };
  const bump = (o, k) => { if (k) o[k] = (o[k] || 0) + 1; };
  const top = o => Object.keys(o).sort((a, b) => o[b] - o[a])[0] || '';
  vals.forEach((r, n) => {
    const cust = String(at(r, I.cust)).trim() || ('row' + n);
    const rm = achMonth_(at(r, I.reg)), am = achMonth_(at(r, I.act));
    const hasReg = String(at(r, I.reg)).trim() !== '';
    if (!String(at(r, I.cust)).trim() && (am || rm)) warn.noCust++;
    const mitra = String(at(r, I.mitra)).trim(), st = String(at(r, I.st)).trim(), rgn = String(at(r, I.reg2)).trim();
    [['ikr', at(r, I.ikrId), at(r, I.ikrName), at(r, I.ikrMail)], ['sales', at(r, I.sal), at(r, I.salName), at(r, I.salMail)]].forEach(([kind, rawId, nm, mail]) => {
      const key = achKey_(rawId); if (!key || /^(-+|0|null|undefined|n\/a|na|#n\/a|none|tidak ada|ny defined)$/.test(key)) { if (kind === 'ikr' ? am : (am || rm)) warn[kind === 'ikr' ? 'noIdIkr' : 'noIdSales']++; return; }   // ID kosong/placeholder → tidak dihitung
      const M = maps[kind];
      const e = M[key] || (M[key] = { id: achId_(rawId), key: key, name: '', email: '', mit: {}, stn: {}, rgn: {}, act: [0, 0, 0], reg: [0, 0, 0], dy: {}, dr: {}, nms: {}, sa: {}, sr: {} });
      if (!e.name) e.name = String(nm).trim();
      { const nv = String(nm).trim(); if (nv) e.nms[nv] = 1; }
      if (!e.email) e.email = String(mail).trim();
      bump(e.mit, mitra); bump(e.rgn, rgn); if (st && st.toUpperCase() !== 'NY DEFINED') bump(e.stn, st);
      // IKR  : aktivasi = ada active_date (bulan mengikuti active_date).
      // Sales: registrasi = ada registration_date; aktivasi = ada registration_date DAN active_date (dihitung 1, bulan mengikuti active_date).
      const okAct = am && (kind === 'ikr' || hasReg);
      if (okAct && !e.sa[cust]) { e.sa[cust] = 1; e.act[am - 1]++; const dd = achDay_(at(r, I.act)); if (dd) e.dy[dd] = (e.dy[dd] || 0) + 1; }
      if (kind === 'sales' && rm && !e.sr[cust]) { e.sr[cust] = 1; e.reg[rm - 1]++; const dr = achDay_(at(r, I.reg)); if (dr) e.dr[dr] = (e.dr[dr] || 0) + 1; }
    });
  });
  ['ikr', 'sales'].forEach(kind => {
    out[kind] = Object.keys(maps[kind]).map(k => { const e = maps[kind][k];
      return { id: e.id, key: e.key, name: e.name, mitra: top(e.mit), stasiun: top(e.stn), region: top(e.rgn), act: e.act, reg: e.reg, dy: e.dy, dr: e.dr }; });
  });
  ['ikr', 'sales'].forEach(kind => Object.keys(maps[kind]).forEach(k => { const nn = Object.keys(maps[kind][k].nms); if (nn.length > 1 && warn.dup.length < 30) warn.dup.push({ kind: kind, id: maps[kind][k].id, names: nn.slice(0, 5) }); }));
  out.warn = warn;
  const sum = (arr, f) => arr.reduce((s, e) => s + f(e).reduce((a, b) => a + b, 0), 0);
  out.stats = { ikrIds: out.ikr.length, ikrAct: sum(out.ikr, e => e.act), salesIds: out.sales.length, salesAct: sum(out.sales, e => e.act), salesReg: sum(out.sales, e => e.reg) };
  return out;
}

/* ---------- Write ---------- */
function submitPeserta_(d) {
  d = d || {};
  const peran = d.peran === 'Sales' ? 'Sales' : (d.peran === 'IKR' ? 'IKR' : '');
  const posisi = d.posisi === 'Lead' ? 'Lead' : (d.posisi === 'Tim' ? 'Tim' : '');
  const nama = clean_(d.nama), mitra = clean_(d.mitra), stasiun = clean_(d.stasiun), akun = clean_(d.akun), uid = clean_(d.uid);
  const hp = String(d.hp || '').replace(/[\s-]/g, '');
  const email = String(d.email || '').trim().toLowerCase();
  if (!mitra || !peran || !posisi || !nama || !stasiun || !uid || !akun) return { error: 'Semua kolom wajib diisi.' };
  if (!/^(\+62|62|0)8\d{7,12}$/.test(hp)) return { error: 'Format No. HP belum benar. Contoh: 081234567890.' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return { error: 'Format email belum benar.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);
    const last = sh.getLastRow();
    if (last > 1) {
      const v = sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
      const dup = v.find(r => String(r[C.HP]).replace(/[\s-]/g, '') === hp || String(r[C.EMAIL]).trim().toLowerCase() === email);
      if (dup) return { error: 'Sudah terdaftar: ' + (dup[C.IKR] || dup[C.SALES]) + ' (No. HP atau email sama).' };
    }
    const row = last + 1;
    const stasiunFix = ensureStasiun_(stasiun);   // nama stasiun dibakukan (huruf besar) & otomatis masuk daftar bila baru
    const newId = Utilities.getUuid().slice(0, 8);
    const rec = [mitra, peran === 'IKR' ? nama : '', peran === 'Sales' ? nama : '', stasiunFix, akun, hp, email, posisi,
      new Date(), newId, uid];
    sh.getRange(row, C.HP + 1).setNumberFormat('@');
    sh.getRange(row, C.TGL + 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(row, 1, 1, HEADERS.length).setValues([rec]);
    ensureMitra_(mitra);
    return { ok: true, id: newId, nama: nama, peran: peran, posisi: posisi };
  } finally { lock.releaseLock(); }
}

/** Mitra mengoreksi data yang baru dikirim (ID dari hasil submit dipakai sebagai kunci). Tanggal input & ID tidak berubah. */
function updatePeserta_(id, d) {
  d = d || {};
  const peran = d.peran === 'Sales' ? 'Sales' : (d.peran === 'IKR' ? 'IKR' : '');
  const posisi = d.posisi === 'Lead' ? 'Lead' : (d.posisi === 'Tim' ? 'Tim' : '');
  const nama = clean_(d.nama), mitra = clean_(d.mitra), stasiun = clean_(d.stasiun), akun = clean_(d.akun), uid = clean_(d.uid);
  const hp = String(d.hp || '').replace(/[\s-]/g, '');
  const email = String(d.email || '').trim().toLowerCase();
  if (!id) return { error: 'Data tidak ditemukan.' };
  if (!mitra || !peran || !posisi || !nama || !stasiun || !uid || !akun) return { error: 'Semua kolom wajib diisi.' };
  if (!/^(\+62|62|0)8\d{7,12}$/.test(hp)) return { error: 'Format No. HP belum benar. Contoh: 081234567890.' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return { error: 'Format email belum benar.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);
    const row = findRow_(sh, id);
    if (!row) return { error: 'Data tidak ditemukan (mungkin sudah dihapus admin).' };
    const last = sh.getLastRow();
    const v = sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
    const dup = v.find((r, i) => (i + 2) !== row && (String(r[C.HP]).replace(/[\s-]/g, '') === hp || String(r[C.EMAIL]).trim().toLowerCase() === email));
    if (dup) return { error: 'Sudah terdaftar: ' + (dup[C.IKR] || dup[C.SALES]) + ' (No. HP atau email sama).' };
    const stasiunFix = ensureStasiun_(stasiun);
    sh.getRange(row, C.HP + 1).setNumberFormat('@');
    sh.getRange(row, 1, 1, C.POSISI + 1).setValues([[mitra, peran === 'IKR' ? nama : '', peran === 'Sales' ? nama : '', stasiunFix, akun, hp, email, posisi]]);
    sh.getRange(row, C.UID + 1).setNumberFormat('@').setValue(uid);
    ensureMitra_(mitra);
    return { ok: true, id: id, nama: nama, peran: peran, posisi: posisi };
  } finally { lock.releaseLock(); }
}

function addMitra_(name) {
  name = clean_(name);
  if (name.length < 3) return { error: 'Nama mitra minimal 3 huruf.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return { ok: true, nama: ensureMitra_(name) }; } finally { lock.releaseLock(); }
}

function addStasiun_(name) {
  name = normStasiun_(name);
  if (name.length < 3) return { error: 'Nama stasiun minimal 3 huruf.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return { ok: true, nama: ensureStasiun_(name) }; } finally { lock.releaseLock(); }
}

function deletePeserta_(pin, id) {
  if (!checkPin_(pin)) return { error: 'Password admin salah.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);
    const row = findRow_(sh, id);
    if (!row) return { error: 'Peserta tidak ditemukan.' };
    sh.deleteRow(row);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function checkAdminPin(pin) { return { ok: checkPin_(pin) }; }

/* ---------- Helpers ---------- */
function rowHasData_(r) { return [C.MITRA, C.IKR, C.SALES, C.HP, C.EMAIL].some(i => String(r[i]).trim() !== ''); }
function clean_(s) { return String(s || '').trim().replace(/\s+/g, ' ').slice(0, 120); }
function digest_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function checkPin_(pin) { return !!pin && String(pin) === String(PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || ''); }
function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  const ids = sh.getRange(2, C.ID + 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
function ensureMitra_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_MITRA);
  const last = sh.getLastRow();
  const list = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
  const hit = list.find(x => x.toLowerCase() === name.toLowerCase());
  if (hit) return hit;
  sh.appendRow([name]);
  return name;
}
function normStasiun_(s) { return clean_(s).toUpperCase(); }
function ensureStasiun_(name) {
  name = normStasiun_(name);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STASIUN);
  const last = sh.getLastRow();
  const list = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => normStasiun_(r[0])) : [];
  if (list.indexOf(name) >= 0) return name;
  sh.appendRow([name]);
  return name;
}
function toObj_(r) {
  const ikr = String(r[C.IKR]).trim(), sales = String(r[C.SALES]).trim();
  const tgl = r[C.TGL] instanceof Date ? r[C.TGL].toISOString() : String(r[C.TGL] || '');
  return {
    id: String(r[C.ID]),
    mitra: String(r[C.MITRA]).trim(),
    peran: ikr ? 'IKR' : 'Sales',
    nama: ikr || sales,
    stasiun: String(r[C.STASIUN]).trim(),
    uid: String(r[C.UID]).trim(),
    akun: String(r[C.AKUN]).trim(),
    hp: String(r[C.HP]).trim(),
    email: String(r[C.EMAIL]).trim(),
    posisi: /lead/i.test(String(r[C.POSISI])) ? 'Lead' : 'Tim',
    createdAt: tgl
  };
}

/* ---------- Daftar resmi stasiun (379 nama) ---------- */
const STASIUN_LIST = [
  "CIGADING",
  "MERAK",
  "KRENCENG",
  "CILEGON",
  "KARANGANTU",
  "SERANG",
  "WALANTAKA",
  "CIKEUSAL",
  "CATANG",
  "JAMBU BARU",
  "RANGKASBITUNG",
  "CITERAS",
  "MAJA",
  "CIKOYA",
  "TIGARAKSA",
  "TENJO",
  "DARU",
  "CILEJIT",
  "PARUNG PANJANG",
  "CICAYUR",
  "CISAUK",
  "SERPONG",
  "RAWA BUNTU",
  "SUDIMARA",
  "JURANGMANGU",
  "PONDOK RANJI",
  "KEBAYORAN",
  "PALMERAH",
  "TANGERANG",
  "BATUCEPER",
  "PORIS",
  "KALIDERES",
  "TAMAN KOTA",
  "PESING",
  "GROGOL",
  "TANAH ABANG",
  "KARET",
  "JAYAKARTA",
  "MANGGA BESAR",
  "RAJAWALI",
  "MANGGARAI",
  "CIPINANG",
  "KLENDER",
  "BUARAN",
  "KLENDER BARU",
  "CAKUNG",
  "KRANJI",
  "BEKASI",
  "BEKASI TIMUR",
  "TAMBUN",
  "LEMAH ABANG",
  "KEDUNGGEDEH",
  "KOSAMBI",
  "CIKAMPEK",
  "CIBUNGUR",
  "SADANG",
  "PURWAKARTA",
  "SUKATANI",
  "PLERED",
  "CISOMANG",
  "CIKADONGDONG",
  "RENDEH",
  "SASAKSAAT",
  "CILAME",
  "PADALARANG",
  "GADOBANGKONG",
  "CIMAHI",
  "ANDIR",
  "TAGOGAPU",
  "CIPATAT",
  "RAJAMANDALA",
  "CIPEUYEUM",
  "SELAJAMBE",
  "MALEBER",
  "CILAKU",
  "CIBEBER",
  "LAMPEGAN",
  "CIREUNGAS",
  "GANDASOLI",
  "CISAAT",
  "KARANG TENGAH",
  "CIBADAK",
  "PARUNG KUDA",
  "CICURUG",
  "CIGOMBONG",
  "MASENG",
  "BOGOR",
  "UNIVERSITAS PANCASILA",
  "LENTENG AGUNG",
  "TANJUNG BARAT",
  "PASAR MINGGU",
  "DUREN KALIBATA",
  "CAWANG",
  "CIBINONG",
  "BOGASARI",
  "NAMBO",
  "TANJUNGRASA",
  "PABUARAN",
  "PRINGKASAP",
  "PASIRBUNGUR",
  "CIKAUM",
  "PEGADEN BARU",
  "HAURGEULIS",
  "CILEGEH",
  "TERISI",
  "TELAGASARI",
  "KERTASEMAYA",
  "CIREBON",
  "WARUDUWUR",
  "TANJUNG",
  "BULAKAMBA",
  "BREBES",
  "TEGAL",
  "BANJARAN",
  "SLAWI",
  "BALAPULANG",
  "MARGASARI",
  "PRUPUK",
  "LINGGAPURA",
  "BUMIAYU",
  "KRETEK",
  "PATUGURAN",
  "LEGOK",
  "KARANGSARI JAWA TENGAH",
  "KARANG GANDUL",
  "PURWOKERTO",
  "NOTOG",
  "KEBASEN",
  "RANDEGAN",
  "KROYA",
  "SIKAMPUH",
  "MAOS",
  "KASUGIHAN",
  "LEBENG",
  "JERUKLEGI",
  "KAWUNGANTEN",
  "GANDRUNGMANGUN",
  "SIDAREJA",
  "CIPARI",
  "MELUWUNG",
  "LANGEN",
  "BANJAR",
  "KARANGPUCUNG",
  "BOJONG",
  "CIAMIS",
  "MANONJAYA",
  "TASIKMALAYA",
  "INDIHIANG",
  "RAJAPOLAH",
  "CIAWI",
  "CIPEUNDEUY",
  "WARUNG BANDREK",
  "CIBATU",
  "LEUWIGOONG",
  "LELES",
  "LEBAKJERO",
  "NAGREG",
  "CICALENGKA",
  "HAURPUGUR",
  "RANCAEKEK",
  "GEDEBAGE",
  "LUWUNG",
  "SINDANGLAUT",
  "KARANGSUWUNG",
  "CILEDUG",
  "KETANGGUNGAN",
  "LARANGAN BREBES",
  "SONGGOM",
  "KARANGKANDRI",
  "CILACAP",
  "PASIRJENGKOL",
  "GARUT",
  "SURADADI",
  "PEMALANG",
  "PETARUKAN",
  "COMAL",
  "SRAGI",
  "PEKALONGAN",
  "BATANG",
  "UJUNGNEGORO",
  "KRENGSENG",
  "WELERI",
  "KALIBODRI",
  "KALIWUNGU",
  "MANGKANG",
  "JERAKAH",
  "SEMARANG PONCOL",
  "SEMARANG TAWANG",
  "ALASTUA",
  "BRUMBUNG",
  "TEGOWANU",
  "GUBUG",
  "KARANGJATI",
  "SEDADI",
  "NGROMBO",
  "GAMBRINGAN",
  "GUNDIH",
  "GOPRAK",
  "SUMBER LAWANG",
  "SALEM",
  "KALIOSO",
  "KADIPIRO",
  "SOLO BALAPAN",
  "PURWOSARI",
  "GAWOK",
  "CEPER",
  "KLATEN",
  "SROWOT",
  "BRAMBANAN",
  "MAGUWO",
  "PATUKAN",
  "REWULU",
  "SENTOLO",
  "WATES",
  "KEDUNDANG",
  "WOJO",
  "JENAR",
  "KUTOARJO",
  "BUTUH",
  "PREMBUN",
  "KUTOWINANGUN",
  "WONOSARI",
  "KEBUMEN",
  "SOKA",
  "SRUWENG",
  "KARANGANYAR",
  "GOMBONG",
  "IJO",
  "TAMBAK",
  "SUMPIUH",
  "KEMRANJEN",
  "TANGGUNG",
  "KEDUNGJATI",
  "TELAWA",
  "KARANGSONO",
  "SOLO KOTA",
  "SUKOHARJO",
  "KEPUH",
  "PASARNGUTER",
  "WONOGIRI",
  "JAMBON",
  "PANUNGGALAN",
  "KRADENAN",
  "SULUR",
  "DOPLANG",
  "RANDUBLATUNG",
  "WADU",
  "KAPUAN",
  "CEPU",
  "KALITIDU",
  "BOJONEGORO",
  "KAPAS",
  "SUMBERREJO",
  "SROYO",
  "BOWERNO",
  "GEMBONG",
  "LAMONGAN",
  "CERME",
  "BENOWO",
  "TANDES",
  "SURABAYA PASAR TURI",
  "SURABAYA KOTA",
  "SURABAYA GUBENG",
  "NGAGEL",
  "WONOKROMO",
  "WARU",
  "GEDANGAN",
  "BUDURAN",
  "SIDOARJO",
  "TANGGULANGIN",
  "PORONG",
  "BANGIL",
  "SENGON",
  "LAWANG",
  "SINGOSARI",
  "BLIMBING",
  "MALANG",
  "MALANG KOTA LAMA",
  "PAKISAJI",
  "KEPANJEN",
  "NGEBRUK",
  "SUMBERPUCUNG",
  "KESAMBEN",
  "WLINGI",
  "TALUN",
  "GARUM",
  "BLITAR",
  "REJOTANGAN",
  "NGUNUT",
  "SUMBERGEMPOL",
  "TULUNGAGUNG",
  "NGUJANG",
  "KRAS",
  "NGADILUWIH",
  "KEDIRI",
  "SUSUHAN",
  "MINGGIRAN",
  "PAPAR",
  "PURWOASRI",
  "KERTOSONO",
  "BARON",
  "SUKOMORO",
  "NGANJUK",
  "BAGOR",
  "WILANGAN",
  "SARADAN",
  "CARUBAN",
  "BABADAN",
  "MADIUN",
  "MAGETAN",
  "GENENG",
  "NGAWI",
  "KEDUNGGALAR",
  "WALIKUKUN",
  "KEDUNGBANTENG",
  "KEBONROMO",
  "SRAGEN",
  "MASARAN",
  "KEMIRI",
  "PALUR",
  "SOLO JEBRES",
  "SEPANJANG",
  "BOHARAN",
  "KRIAN",
  "TARIK",
  "MOJOKERTO",
  "CURAHMALANG",
  "SUMOBITO",
  "PETERONGAN",
  "JOMBANG",
  "SEMBUNG",
  "TULANGAN",
  "PASURUAN",
  "REJOSO",
  "BAYEMAN",
  "PROBOLINGGO",
  "LECES",
  "KLAKAH",
  "RANDUAGUNG",
  "JATIROTO",
  "TANGGUL",
  "BANGSALSARI",
  "ROGOJAMPI",
  "RAMBIPUJI",
  "MANGLI",
  "KOTOK",
  "KALISAT",
  "GARAHAN",
  "MRAWAN",
  "KALIBARU",
  "GLENMORE",
  "SUMBERWADUNG",
  "KALISETAIL",
  "TEMUGURUH",
  "SINGOJURUH",
  "BANYUWANGI KOTA",
  "ARGOPURO",
  "KETAPANG",
  "SURABAYAN",
  "DAWUAN",
  "LEMPUYANGAN",
  "KANDANGAN",
  "PASAR MINGGU BARU",
  "BUMI WALUYA",
  "AWIPARI",
  "WANARAJA",
  "BABAT",
  "TEBET",
  "CIMINDI",
  "DELANGGU",
  "LOSARI",
  "KERTOMENANGGAL",
  "PURWOREJO",
  "GUMILIR",
  "CIKARANG",
  "BUMIWALUYA",
  "CILEBUT",
  "PAGERWOJO",
  "MARGOREJO"
];

/* ---------- Daftar resmi mitra (405 nama) ---------- */
const MITRA_LIST = [
  "PT ABIMANYU BANGKIT MANDIRI",
  "PT ABIMANYU PUTRA PRATAMA",
  "PT ABYUDAYA BERKAH SOLUSITAMA",
  "PT ACCESS PRIMA NUSANTARA",
  "ADE PRIYADI",
  "PT ADHI BIMA PARAMITHA",
  "PT ADHIGUNA DIRGANTARA GEMILANG",
  "PT ADHITAMA GANIA SEJAHTERA",
  "PT ADIANA EXPRESS INDONESIA",
  "PT ADZKA MEDIA INDOPERKASA",
  "PT AFHAM SOLUSI KOMUNIKASI",
  "AFNAN WICAKSANA MANDIRI",
  "PT AGUS JAYA TEKNIK",
  "PT AKSES ARTHA MEDIA",
  "PT AKSES BERSAMA SEDAYA",
  "ALAN NAURIE SYAFEI",
  "PT ALFARY MODULA",
  "PT ALIGATOR FIBER OPTIC",
  "PT ALL NETWORK DIGITAL INDONESIA",
  "PT ALPHA GLOBAL TEKNOLOGI",
  "CV ANANTA TIGA BINTANG",
  "PT ANDIKA KOMUNIKA TAMA",
  "PT ANDOR GLOBAL PERKASA",
  "PT ANJANI DIGITAL NUSANTARA",
  "PT ANTA SOLUSI ADIKARYA",
  "PT ANTERO TECHNOLOGY",
  "PT ANUGERAH SELARAS INDONESIA",
  "PT ANUGRAH BORNEO NUSANTARA",
  "PT ANUGRAH SOLUSI INFRASTRUKTUR KOMUNIKASI",
  "CV ANUGRAH SUKSES ABADI COMMUNICATION",
  "PT APRILLIA PROFESSIONAL TEKNOLOGI",
  "PT ARBYAN ELEKTRINDO PERSADA",
  "PT ARIOSY SADIA RESWARA",
  "PT ARIS INDOGAS",
  "PT ARJI TECH INDONESIA",
  "PT ARTHA BERKAH EXPONENSIA",
  "PT ARTHA GLOBAL SOLUSI",
  "PT ARTHADIRO PUTRA MAHESWARA",
  "PT ARUNIKA JAYA PERSADA",
  "PT ASIA GLOBAL SURYA",
  "PT ASSET RECOVERY INDONESIA",
  "PT AWAN PUTRA KUSUMA",
  "PT AYUBEE LOGAM PERKASA",
  "PT AZZAHRA RAJAWALI MAKMUR",
  "PT BAHANA INFRA KOMUNIKA",
  "PT BAHTERA PINTAS UTAMA",
  "PT BANGTELINDO",
  "PT BANGUN JARINGAN BERSAMA",
  "PT BARAYA UTAMA TEKNIK",
  "CV BAROKAH UTAMA",
  "PT BATAVIA INDO GLOBAL MANDIRI",
  "PT BERATHI",
  "CV BERKAH INTER GENERATION",
  "CV BERKAH JAGO PERKASA",
  "PT BERKAH SERIKAT MANDIRI",
  "CV BERKAH SHOLAWAT INDONESIA",
  "CV BERKAH TEKNIK MITRA ABADI",
  "PT BERKAT BERSAMA TEKNIK",
  "PT BESTCAMP PRIMA DATA",
  "PT BIMASAKTI SELULAR INDONESIA",
  "CV BINA TARUNA MANDIRI",
  "PT BINA TECHINDO SOLUTION",
  "PT BINA UPAYA SEMESTA",
  "PT BINTANG MULTIMEDIA PRAKASA",
  "PT BINTANG TERANG DELAPAN SEMBILAN",
  "PT BINTANG TRANS KHATULISTIWA",
  "PT BLACKFAOS DATA NUSANTARA",
  "PT BLUE CONNECT INDONESIA",
  "CV BNET KARYA BERSAMA",
  "PT BRAINWORX SOLUSI INTEGRASI",
  "CV BUANA SEJAHTERA",
  "PT CADASPANGERAN DATA SOLUSION",
  "PT CAHAYA AGUNG OETAMA",
  "CV CAHAYA DUA PUTRA",
  "PT CAHAYA TEKNO UNIVERSA",
  "PT CAKRAWALA SOLUSI NUSANTARA",
  "PT CAKRAWALA STRATA OPTIMA",
  "CV CANDI SEJAHTERA",
  "PT CEMERLANG MEDIA VISION",
  "PT CHARA ABADI SEJAHTERA",
  "PT CIPTA AKSES INDOTAMA",
  "PT CIPTA OPTIMA SOLUSI",
  "PT CIPTA UTAMA KARRYA",
  "PT CIPTAJAYA SEJAHTERA ABADI",
  "PT CITRA DAYA MAXIMA",
  "PT CITRA META DATA",
  "CV CITRA TUNGGAL ABADI",
  "PT CLOVER TECH DIGITAL",
  "PT COMMSEC SOLUTION PROVIDER",
  "PT CYNET INDONESIA NETWORK",
  "CV DAKSA KARYA TEKNOLOGI",
  "PT DAPOER POESAT NOESANTARA GROUP",
  "PT DARIA PRATAMA MANDIRI",
  "DARWIN",
  "PT DATANET SOLUSI MAKMUR",
  "PT DAVON MEDIA TEKNOLOGI",
  "PT DAYA GUNA KARSA",
  "CV DAYA KARYA MANDIRI",
  "PT DECON SUKSES INDONESIA",
  "PT DELAPAN UNSUR INTI TAMA",
  "PT DESNARUM JAYA AKASHA",
  "PT DEXA ENERGY",
  "PT DEXRADO SINERGI SELARAS",
  "CV DIGITAL ANDALAN SOLUSI",
  "PT DIMENSI MULYA TALENTA",
  "DIO PRATAMA",
  "CV DUA PUTRA AKAS SEJAHTERA",
  "CV DUA PUTRA MILIARTA",
  "PT DUTA ANUGRAH DAMAI SEJAHTERA",
  "PT DUTACOM INTERNUSA PERSADA",
  "PT DWI PILAR PRATAMA",
  "DWI WAHYUDI",
  "PT EKSPANINDO PRIMA MULTIMEDIA",
  "PT EL-KOKAR TIMUR",
  "PT ERA BANGUN INDONESIA",
  "PT ERAJAYA MAKMUR PERKASA",
  "PT ERZAR INTI SOLUSI",
  "CV ESKA GLOBAL PERSADA",
  "PT FAJAR MITRA KRIDA ABADI",
  "PT FANAUVI INFOTECH GEMILANG",
  "PT FANIZ AMANAT BERSAMA",
  "PT FIBEART TRANS NETWORK",
  "PT FIBER DATA NUSANTARA",
  "PT FIBER PULSE NETWORK",
  "PT FIBERHOME TECHNOLOGIES INDONESIA",
  "CV FLASHNET SINERGI BERJAYA",
  "PT FORTY INTEGRASI OPTIMAL",
  "PT FOTON INFO SISTEM",
  "FREDY NUGROHO",
  "PT GAISAR TEKNOLOGI BERSAMA",
  "PT GANDA MADY INDOTAMA",
  "PT GARDA UTAMA PRIMA",
  "PT GARUDA MITRA SOLUSI",
  "PT GARUDA PRIMA SEKAWAN",
  "PT GENESIS BERKAT USAHA",
  "PT GENPOWER TOTAL ENERGI",
  "PT GENZ PERSADA INDONESIA",
  "PT GERBANG KERUMAH INDONESIA",
  "PT GERBANG NUSANTARA SAKTI",
  "PT GIANDRA SAKA MEDIA",
  "PT GIGA DATACOM",
  "PT GILANG SAKTI KANAKA",
  "PT GLOBAL ACCESS SERVICE NETWORK",
  "PT GLOBAL AKSES INTERNUSA",
  "PT GLOBAL SARANA MEDIAKOM",
  "PT GOLDEN NETWORK NUSANTARA",
  "PT GUMILANG NUSANTARA ABADI",
  "PT HANDAYANI FIBER MEDIA",
  "PT HARAPAN USAHA BHAKTI",
  "CV HARSA SANJAYA SOLUTION",
  "PT HASIAN PRIMA TELINDO",
  "CV HASTALOKA MITRA UTAMA",
  "PT HAYOOKERJA ABHINAYA INDONESIA",
  "CV HEGA SOLUSINDO",
  "HENDRIK",
  "PT IDKO ECO JAYA ENERGI",
  "CV INDOMITRA TEKNOLOGI GROUP",
  "PT INDOTECH DIGITAMA SUPERLINK",
  "PT INFRA SEKAWAN TELEKOMUNIKASI",
  "PT INTERINDO OETAMA NEOTEKNOLOGI",
  "PT INTERMEDIA LINTAS NUSA",
  "PT INTERNET ANAK BANGSA ( ASKO )",
  "PT INTERNET ANAK BANGSA ( BSI )",
  "PT INTERNET ANAK BANGSA ( ERZAR )",
  "PT INTERNET ANAK BANGSA ( OLT )",
  "PT INTERNET ANAK BANGSA ( TBM )",
  "PT INTERNET RAKYAT INDONESIA TERCEPAT",
  "PT INTERNET RAKYAT SOLUSI INDONESIA",
  "PT INTERNUSA DUTA MAKMUR",
  "ISDARYANTO",
  "PT J & JC CONSULTING",
  "PT JALATIKA AGUNG PERKASA",
  "PT JAMBON ELEKTRIK UTAMA",
  "PT JANUR KUNING SENTOSA",
  "PT JARING INFRASTRUKTUR NASIONAL",
  "PT JARING SOLUSI PERSADA",
  "PT JARINGAN CYBER EVO",
  "PT JARINGAN INTERNET BANTEN",
  "PT MITRA JASIKOM INDONESIA",
  "PT JAYAHANA MUNAURA MEKANIKA SELARAS",
  "PT JENDELA WAHANA SOLUSI",
  "PT KAISAR UTAMA MANDIRI",
  "PT KALYANA MITRA LESTARI",
  "PT KARYA LANGIT MANDIRI",
  "PT KARYA MANUNGGAL TEKNOLOGI",
  "PT KARYA TARUNA TEKNIK",
  "PT KASIH RAJAWALI SEJAHTERA",
  "PT KAWAN INFORMATIKA TEKNOLOGI ABADI",
  "PT KEDIREN NET MEDIA BLORA",
  "KELLYK SUDARAONO",
  "PT KENTJANG INTERNET AKSES",
  "PT KESHAVA MULTI INTER MEDIA",
  "PT KHANCA SINERGI SOLUTIONS",
  "PT KHARISMA ELSYADAI SUKSES ABADI",
  "PT KHATULISTIWA JARINGAN NUSATECH",
  "PT KONNEK JAYA BERSAMA",
  "PT KONSULTASI INDONESIA KANADA",
  "PT KOPERASI PEGAWAI INDOSAT (KOPINDOSAT)",
  "PT KOPERASI PEGAWAI TELKOM MALANG (KOPEGTEL)",
  "PT KOPNATEL JAYA",
  "PT KRAKATOA PRADASWARA",
  "PT KREASI INOVASI ONLINE SEJAHTERA NETWORK",
  "PT KRESA UTAMA MANDIRI",
  "KUSNADI",
  "PT LENTERA ABADI SOLUSINET",
  "PT LENTERA DIGITAL NUSANTARA",
  "PT LIM CIPTA UTAMA",
  "PT LIMA SINERGI UTAMA",
  "PT LINGKAR KABEL TELEKOMUNIKASI",
  "PT LINGKUP TOTAL TECHNOLOGY",
  "PT LINK NET",
  "CV LINTANG UNGGUL NUSANTARA",
  "PT LINTAS TEKNOLOGI SOLUSINDO",
  "PT LUMINTU",
  "PT MADISON GLOBAL SOLUSINDO",
  "PT MADIUN AKSES TELEKOMUNIKASI",
  "PT MAGNA CAHAYA UTAMA",
  "PT MAHAGUNA KARYA MANDIRI",
  "PT MAJU JAYA TEKNOLOGI",
  "PT MANDIRI DAYA UTAMA NUSANTARA",
  "PT MATERIAL SUPPLY INDONESIA",
  "PT MAWAH MITRA MANDIRI",
  "PT MEDIA AKSES INDONESIA",
  "PT MEDIA CIPTA TRIMETIKA",
  "PT MEGA ARTHA CARAKA",
  "PT MEGA LESTARI JARINGAN TELEKOMUNIKASI",
  "PT MEGA NETWORKING ABADI",
  "PT MELESAT PRIMA NUSANTARA",
  "PT MENARA INDRA UTAMA",
  "PT METRO AKSES PRATAMA",
  "PT METRO DUTA SELARAS",
  "MITRA MIFTAHUDDIN",
  "PT MIGE TAMA SINERGI",
  "PT MILIARTO PRADANA PROYEK",
  "PT MITRA AKSES SOLUSINDO",
  "PT MITRA DIGITAL GLOBALINDO",
  "PT MITRA JAVA TECHNOLOGY",
  "PT MITRA MAS EKA",
  "PT MITRA METASTRUKTUR GLOBAL",
  "PT MITRA SISTEMATIKA GLOBAL",
  "PT WANWAN DIKDIK",
  "MOCH SYAIFUL ULUM",
  "MUHAMMAD MIFTAKHUDDIN HANIF",
  "PT MULTIPOLAR TECHNOLOGY",
  "PT NABAWI NETWORK INDONESIA",
  "NAIM ROFIQ",
  "NANDA AGUNG GUMILANG",
  "CV NATORI DEMAH AMANAH",
  "CV NAVIS WIRA MANDIRI",
  "PT NEIMA NETWORKS ACCESS",
  "PT NETIN DATA PRATAMA",
  "PT NETZEN MEDIA AKSES",
  "CV NIXBIAN SOLUSI",
  "PT NOKTURA OPTIMA NALATIKA",
  "PT NUGRAHA NETWORK CENTER",
  "PT NUGROHO INFRASTRUKTUR TEKNOLOGY",
  "NUR DAWAN",
  "PT NURIZ BERKAH ABADI",
  "PT NUSA ADHANU SABIATAMA",
  "PT NUSA BAKTI KOMUNIKA",
  "PT OLEAN PERMATA TELEMATIKA",
  "PT OPMC INDONESIA",
  "PT OPTICAL LINTAS TEKNOLOGI",
  "PT ORBIT EKA SEMESTA",
  "PT PANCA BASWARA SAKTI",
  "CV PANDU SATRIA WICAKSANA",
  "PT PANGKALAN LINTAS DATA",
  "PT PANGKALAN LINTAS DATA PALINDO",
  "PT PASS INTERNET INDONESIA",
  "PT PEMERSATU WIFI NUSANTARA",
  "PT PERSONEL ALIH DAYA",
  "PT PHAINAN JAYA UTAMA",
  "PT PHATRIA INTI PERSADA",
  "PT PILAR GAPURA NUSA",
  "PT PLANET HARAPAN PENUH ENERGY",
  "PT POLARISNET",
  "PT POWER TELCO INDUSTRI",
  "PT PRADIKTA UNGGUL PERKASA",
  "PT PRAGATA MAKMUR PERSADA",
  "PT PRIMA AKSES SOLUSI GLOBAL",
  "PT PULINTA KARYA UTAMA",
  "PT PUNCAK TIMUR PARAHYANGAN",
  "PT PUNDI MAS TIGAPUTRI",
  "PT PUTRA LEBAK BANTEN",
  "PT PUTRA PERSADA TAMA",
  "PT PUTRA TEKNOLOGI SOLUSINDO",
  "PT PUTRA TELEKOMUNIKASI INDONESIA",
  "PT PUTRI NAULI MANDIRI",
  "CV QIANA NURUL HIKMAH",
  "PT QUANTUM NUSATAMA",
  "PT QUEEN TECHNOLOGY GLOBAL",
  "PT QUEENSHA MIKAILA RAHAYU",
  "PT RADHIKA JAYA PRATAMA",
  "PT RAHAYU KARYA MESARI",
  "PT RAIHAN TEKNOLOGI PRATAMA",
  "CV RAJASA EMILY",
  "PT RAJASTAR MEDIA NUSANTARA",
  "PT RAJEG MEDIA TELEKOMUNIKASI",
  "PT RAMA MULTIGUNA ADIKARYA",
  "RENALDI",
  "PT RIA KUSUMAH BERSAMA",
  "CV RIZKY DIGITAL SOLUSINDO",
  "PT RIZQI KARYA BERSAMA",
  "PT RNJ NETWORK JARINGAN",
  "PT ROMA UNGGUL TELEKOMINDO",
  "PT SALWA CITRA MANDIRI",
  "PT SAMARRA ALAMSAH BERKAH",
  "PT SAMUDERA EMAS NUSANTARA",
  "PT SAPUTRA GLOBAL NUSANTARA",
  "PT SARANA MANDIRI USAHA UNGGUL",
  "PT SARANA MITRA POWERINDO",
  "PT SARVA SOLUTION INDONESIA",
  "PT SATYA JALA MANDIRI",
  "SBAGJA TECH",
  "PT SECURINDO MITRA SEJATI",
  "PT SEDAYU CAHAYA PERKASA",
  "SEGALABISAKONSEP",
  "PT SEJAHTERA TRI MULYA INDONESIA",
  "PT SEMANGAT MEDIATAMA",
  "PT SEMBADA MAJU BERSAMA",
  "PT SEMERU AGUNG MANDIRI",
  "PT SEMESTA ENERGI SEVICES",
  "PT SEMESTA PUSAT KREASI",
  "CV SEMPUR JAYA",
  "PT SENJU METADATA INDONESIA",
  "PT SENTRA ENERGI TEKNIK KOMUNIKASI",
  "PT SIMBIKA TEKNOLOGI SOLUSI",
  "PT SINAR DATA BERSAMA",
  "CV SINAR MODERN",
  "PT SINAR PALASARI INDONESIA",
  "PT SINERGI MULTILINK INDONESIA",
  "PT SINERGI SELARAS TEKNOLOGI",
  "PT SKYNET LINTAS NUSANTARA",
  "SOLIHUDIN",
  "SONY PRAWIRA",
  "PT SRINARENDRA SEJAHTERA ABADI",
  "PT STAR HOME INDONESIA",
  "PT STAR TEKNOLOGI DEVELOPMENT",
  "PT STEP POINT INDONESIA",
  "PT SUMBER CEMERLANG KENCANA PERMAI",
  "PT SUMBER HARUMAN NUSANTARA",
  "PT SUNRISE INTERNUSA",
  "PT SURYASANTIKA INFRASTRUKTUR MEDIA SELARAS",
  "PT SYAMINDO TEKNIK MANDIRI",
  "PT TALANG AMANAH MAKMUR",
  "CV TANGGAP WASKITA",
  "PT TANGGUH KAPITAL INFRASARANA",
  "PT TEKKOMINDO KARYA WIJAYA",
  "PT TEKLING MEDIA TELEMATIKA",
  "PT TEKNO INDO JAYA",
  "PT TEKNOLOGI ANUGERAH ABADI",
  "PT TELEMEDIA NETWORK CAKRAWALA",
  "PT TELEMITRA GLOBAL SOLUSINDO",
  "PT TELINCO NETWORKS INDONESIA",
  "PT TELINDO FLASH MEDIATAMA",
  "PT TELIO INTI NUSA",
  "PT TELKOM AKSES",
  "PT TERUS RAIH PRESTASI",
  "PT TIGA NOVA SENTOSA",
  "PT TIRTAMAS BERKAH MULIA",
  "TOKO NETWORK",
  "PT TRANS ARTHA TEKNOLOGI NUSANTARA",
  "PT TRANS HYBRID COMMUNICATION",
  "PT TRESNO SUMBER REJEKI",
  "PT TRI AKSES NUSANTARA",
  "PT TRI BUMI ASIH",
  "PT TRI SUKHA PRATAMA",
  "PT TRIJAYA DIGITAL TEKNOLOGI",
  "PT TRIJAYA MAJU TOTALINDO",
  "CV TRIJAYA STEEL",
  "PT TRIK MEDIA DATA",
  "PT TRIPOLA PANATA",
  "PT TRISARI DATA INDONESIA",
  "PT TRISULA MEGAH JAYA",
  "PT TRITUNAS PUTRA NUSANTARA",
  "PT TRIVORA NUSANTARA TEKNOLOGI",
  "CV TRIZA PRATAMA",
  "PT TUJUH SINAR ABADI",
  "TULUNGAGUNG NET",
  "TUMINO",
  "PT TURANDI MITRA HANDAL",
  "UJANG SUPRIATNA",
  "PT ULTRA MANDIRI TEKNOLOGI",
  "CV VANTARA KARYA JAYA",
  "PT VINMORZA JAYA NUSANTARA",
  "PT WAHANA ELEKSIA TECHNOLOGY",
  "PT WAHANA KARYA NETWORK",
  "PT WARNINDO DIGITAL INDONESIA",
  "PT WIDIA KARYA ERA MANDIRI UTAMA",
  "PT WIJAYA KARYA ARTA",
  "CV WIJAYA PRIMA TELEINDO",
  "PT WIRASATYA WALI NASSER",
  "PT WIYASA PERENCANA INDONESIA",
  "PT XAPIENS TEKNOLOGI INDONESIA",
  "YOGA INDRIAS MAYA",
  "ZAENAL A.",
  "PT ZONA OPTIC NUSANTARA",
  "PT MEDIA ARTOPALA NUSANTARA",
  "PT CYBER NETWORK SOLUSINDO",
  "PT FASILITAS TELEKOM NUSANTARA",
  "PT DUKODU DIGITAL SOLUTION",
  "PT ERA MEDIA UTAMA",
  "PT INTEGRASI JARINGAN EKOSISTEM",
  "PT THREE MUSKETEERS TEKNOLOGI",
  "PT SATRIA SAKTI MANDIRI"
];
function submitPeserta(d) { const r = submitPeserta_(d); bumpV_(); return r; }

function updatePeserta(id, d) { const r = updatePeserta_(id, d); bumpV_(); return r; }

function addMitra(name) { const r = addMitra_(name); bumpV_(); return r; }

function addStasiun(name) { const r = addStasiun_(name); bumpV_(); return r; }

function deletePeserta(pin, id) { const r = deletePeserta_(pin, id); bumpV_(); return r; }

/** Penanda versi data (naik tiap ada tulis dari web) — dipakai getDigest tanpa harus membaca seluruh sheet. */
/* ---------- Hasil final (dibekukan admin) ---------- */
const FINAL_HEAD = ['peran', 'level', 'region', 'id', 'nama', 'mitra', 'stasiun', 'total_aktivasi', 'total_registrasi', 'rata2_per_hari'];
/** Admin membekukan hasil: seluruh isi tab "Hasil Final" diganti snapshot terbaru. rows = [{peran,level,region,id,nama,mitra,stasiun,total,reg,rate}] */
function saveFinal(pin, rows) {
  if (!checkPin_(pin)) throw new Error('Password admin salah.');
  rows = (rows || []).slice(0, 5000);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(SHEET_FINAL); if (!sh) sh = ss.insertSheet(SHEET_FINAL);
    sh.clearContents();
    const at = new Date();
    const out = [FINAL_HEAD].concat(rows.map(r => [r.peran, r.level, r.region, r.id, r.nama, r.mitra, r.stasiun, r.total, r.reg, r.rate]));
    sh.getRange(1, 1, out.length, FINAL_HEAD.length).setNumberFormat('@').setValues(out.map(a => a.map(x => String(x === undefined || x === null ? '' : x))));
    sh.getRange(1, FINAL_HEAD.length + 2).setValue('dibekukan');
    sh.getRange(2, FINAL_HEAD.length + 2).setValue(at.toISOString());
    PropertiesService.getScriptProperties().setProperty('FINAL_AT', at.toISOString());
    return { ok: true, at: at.toISOString(), n: rows.length };
  } finally { lock.releaseLock(); }
}
function getFinal(pin) {
  if (!checkPin_(pin)) throw new Error('Password admin salah.');
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_FINAL);
  const at = PropertiesService.getScriptProperties().getProperty('FINAL_AT') || '';
  if (!sh || sh.getLastRow() < 2 || !at) return { at: '', rows: [] };
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, FINAL_HEAD.length).getValues();
  return { at: at, rows: v.map(r => ({ peran: r[0], level: r[1], region: r[2], id: r[3], nama: r[4], mitra: r[5], stasiun: r[6], total: Number(r[7]) || 0, reg: Number(r[8]) || 0, rate: Number(r[9]) || 0 })) };
}

/** Simpan data publik ke cache (2 menit) dan perbarui sidik isi bila berubah. */
function storePub_(pub) {
  putBig_('pub', pub, 120);
  try {
    const h = digest_(JSON.stringify([pub.mitra, pub.stasiun, pub.ach, pub.regs])), P = PropertiesService.getScriptProperties();
    if (P.getProperty('PUB_H') !== h) P.setProperty('PUB_H', h);
  } catch (e) {}
}
/** Dijalankan trigger waktu tiap 1 menit: menghitung ulang data di latar belakang agar pengunjung selalu mendapat cache hangat (cepat). */
function refreshCache_() {
  try { CacheService.getScriptCache().remove('ach_n'); storePub_(buildPub_(readValues_())); } catch (e) {}
}
function bumpV_() { try { PropertiesService.getScriptProperties().setProperty('DATA_V', String(Date.now())); const cc = CacheService.getScriptCache(); cc.remove('pub_n'); cc.remove('ach_n'); } catch (e) {} }

/* ---------- Realtime: perubahan manual di Google Sheet langsung terdeteksi ---------- */
/** Trigger sederhana: aktif otomatis saat sel diedit manual. */
function onEdit(e) { bumpV_(); }
/** Trigger terpasang (onChange): juga menangkap paste/import/hapus baris. Jalankan fungsi setupTriggers sekali dari editor. */
function onSheetChange_(e) { bumpV_(); }
function setupTriggers() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onSheetChange_').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onSheetChange_').forSpreadsheet(ss).onChange().create();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'refreshCache_').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshCache_').timeBased().everyMinutes(1).create();
  refreshCache_();
}

/**
 * SISTEM KHS - SMK Muhammadiyah Bojong
 * Script ini TERIKAT (bound) ke Google Spreadsheet.
 * Sheet: Pengaturan | Mapel | Kelas | Leger  (dibuat otomatis)
 */
const SH = { SET: 'Pengaturan', MAP: 'Mapel', KLS: 'Kelas', LEG: 'Leger' };

const DEF_SET = [
  ['NAMA_SEKOLAH', 'SMK MUHAMMADIYAH BOJONG'],
  ['BARIS_1', 'MAJELIS PENDIDIKAN DASAR DAN MENENGAH'],
  ['BARIS_2', 'DAERAH MUHAMMADIYAH KABUPATEN PEKALONGAN'],
  ['PROGRAM_STUDI', 'Program Studi: Teknik Sepeda Motor Terakreditasi ( B ) & Teknik Komputer Multimedia Terakreditasi ( B )'],
  ['ALAMAT', 'Alamat : Jl. Raya Sembung Jambu Kec. Bojong, Kab. Pekalongan 51156 (0285) 7830017'],
  ['KONTAK', 'Email : smkmuhamjong@gmail.com   Website : www.smkmuhamjong.sch.id'],
  ['JUDUL', 'KARTU HASIL BELAJAR TENGAH SEMESTER GANJIL'],
  ['TAHUN_PELAJARAN', '2026 / 2027'],
  ['SEMESTER', 'Ganjil'],
  ['TEMPAT', 'Bojong'],
  ['TANGGAL', '17 Oktober 2026'],
  ['KEPSEK', 'Moh.Andi Nugroho, M.Pd'],
  ['NBM_KEPSEK', '1204 862'],
  ['TANDAI_REMIDI', 'Ya'],
  ['LOGO_KIRI', ''],
  ['LOGO_KANAN', ''],
  ['TTD_KEPSEK', '']
];

const DEF_MAPEL = [
  ['A', 'Pendidikan Agama Islam dan Budi Pekerti', 70],
  ['A', 'Pendidikan Pancasila', 75],
  ['A', 'Bahasa Indonesia', 75],
  ['A', 'Pendidikan Jasmani, Olahraga, dan Kesehatan', 75],
  ['A', 'Sejarah', 75],
  ['A', 'Seni Budaya', 75],
  ['A', 'Bahasa Jawa', 75],
  ['B', 'Matematika', 75],
  ['B', 'Bahasa Inggris', 75],
  ['B', 'Informatika', 75],
  ['B', 'Projek Ilmu Pengetahuan Alam dan Sosial', 75],
  ['B', 'Dasar-dasar Program Keahlian', 75],
  ['C', 'Al Islam', 70],
  ['C', 'Pendidikan Kemuhammadiyahan', 70],
  ['C', 'Bahasa Arab', 70]
];

/* ---------- util ---------- */
const norm_ = s => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
const kkey_ = s => String(s == null ? '' : s).toUpperCase().replace(/[\s\-]/g, '').replace(/([A-Z])I$/, (m, a) => a + '1');
const isNum_ = v => v !== '' && v !== null && v !== undefined && !isNaN(v);

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(n) { return ss_().getSheetByName(n) || ss_().insertSheet(n); }

function onOpen() {
  try { ensureSheets(); } catch (e) {}
  SpreadsheetApp.getUi().createMenu('KHS')
    .addItem('Buka Sistem KHS (link)', 'showLink')
    .addItem('Hitung ulang jumlah & peringkat', 'recalcAll')
    .addItem('Siapkan sheet', 'ensureSheets')
    .addToUi();
}

/** Jalankan SEKALI dari editor (pilih fungsi setup > Run) agar sheet langsung muncul */
function setup() {
  ensureSheets();
  SpreadsheetApp.getActiveSpreadsheet().toast('Sheet Pengaturan, Mapel, Kelas, Leger sudah siap.', 'Sistem KHS', 5);
}

function showLink() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(url
    ? '<p>Buka sistem:</p><a href="' + url + '" target="_blank">' + url + '</a>'
    : '<p>Aplikasi web belum di-deploy. Terapkan &gt; Deployment baru &gt; Aplikasi web.</p>').setWidth(420).setHeight(110);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sistem KHS');
}

function doGet() {
  ensureSheets();
  const baca = n => HtmlService.createHtmlOutputFromFile(n).getContent();
  // fungsi pengganti (bukan string) agar karakter "$" di kode JS tidak ditafsirkan
  const html = baca('index')
    .replace('<link rel="stylesheet" href="style.css">', () => '<style>\n' + baca('style') + '\n</style>')
    .replace('<script src="script.js"></script>', () => '<script>\n' + baca('script') + '\n</script>');
  return HtmlService.createHtmlOutput(html)
    .setTitle('Sistem KHS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ensureSheets() {
  const s = ss_();
  const order = [SH.LEG, SH.MAP, SH.KLS, SH.SET];
  if (!s.getSheetByName(SH.SET)) {
    const sh = s.insertSheet(SH.SET);
    sh.getRange(1, 1, 1, 2).setValues([['Kunci', 'Nilai']]).setFontWeight('bold');
    sh.getRange(2, 1, DEF_SET.length, 2).setValues(DEF_SET);
    sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 500);
  }
  if (!s.getSheetByName(SH.MAP)) {
    const sh = s.insertSheet(SH.MAP);
    sh.getRange(1, 1, 1, 3).setValues([['Kelompok', 'Mata Pelajaran', 'KKTP']]).setFontWeight('bold');
    sh.getRange(2, 1, DEF_MAPEL.length, 3).setValues(DEF_MAPEL);
    sh.setColumnWidth(2, 340);
  }
  if (!s.getSheetByName(SH.KLS)) {
    const sh = s.insertSheet(SH.KLS);
    sh.getRange(1, 1, 1, 7).setValues([['Kelas', 'Wali Kelas', 'NBM Wali', 'Bidang Keahlian', 'Program Keahlian', 'Kosentrasi', 'Fase']]).setFontWeight('bold');
    sh.setColumnWidth(2, 240);
  }
  if (!s.getSheetByName(SH.LEG)) {
    const sh = s.insertSheet(SH.LEG);
    writeLeger_(DEF_MAPEL.map(m => m[1]), []);
  }
  // hapus sheet bawaan kosong ("Sheet1"/"Sheet1") bila ada
  s.getSheets().forEach(sh => {
    if (/^(Sheet|Lembar)\s?1$/i.test(sh.getName()) && sh.getLastRow() === 0 && s.getSheets().length > 1) s.deleteSheet(sh);
  });
  // urutan tab: Leger, Mapel, Kelas, Pengaturan
  order.forEach((n, i) => {
    const sh = s.getSheetByName(n);
    if (sh) { s.setActiveSheet(sh); s.moveActiveSheet(i + 1); }
  });
  s.setActiveSheet(s.getSheetByName(SH.LEG));
}

/* ---------- pengaturan ---------- */
function getSettings_() {
  const v = sheet_(SH.SET).getDataRange().getValues().slice(1);
  const o = {};
  DEF_SET.forEach(d => o[d[0]] = d[1]);
  v.forEach(r => { if (r[0]) o[r[0]] = String(r[1] == null ? '' : r[1]); });
  return o;
}
function getMapel_() {
  return sheet_(SH.MAP).getDataRange().getValues().slice(1)
    .filter(r => r[1] !== '')
    .map(r => ({ kelompok: String(r[0]).toUpperCase(), nama: String(r[1]).trim(), kktp: isNum_(r[2]) ? Number(r[2]) : 75 }));
}
function getKelas_() {
  return sheet_(SH.KLS).getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '')
    .map(r => ({ kelas: String(r[0]), wali: r[1], nbm: String(r[2] || ''), bidang: r[3], program: r[4], kosentrasi: r[5], fase: r[6] }));
}

function getState() {
  ensureSheets();
  const L = readLeger_();
  return {
    settings: getSettings_(),
    mapel: getMapel_(),
    kelas: getKelas_(),
    jumlahSiswa: L.rows.length,
    url: ss_().getUrl()
  };
}

function saveSettings(obj) {
  const keys = DEF_SET.map(d => d[0]);
  Object.keys(obj).forEach(k => { if (keys.indexOf(k) < 0) keys.push(k); });
  const out = keys.map(k => [k, obj[k] == null ? '' : String(obj[k])]);
  const sh = sheet_(SH.SET);
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([['Kunci', 'Nilai']]).setFontWeight('bold');
  sh.getRange(2, 1, out.length, 2).setValues(out);
  return true;
}

function saveKelas(arr) {
  const sh = sheet_(SH.KLS);
  sh.clear();
  sh.getRange(1, 1, 1, 7).setValues([['Kelas', 'Wali Kelas', 'NBM Wali', 'Bidang Keahlian', 'Program Keahlian', 'Kosentrasi', 'Fase']]).setFontWeight('bold');
  const out = arr.filter(k => k.kelas).map(k => [k.kelas, k.wali || '', k.nbm || '', k.bidang || '', k.program || '', k.kosentrasi || '', k.fase || '']);
  if (out.length) {
    sh.getRange(2, 3, out.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, out.length, 7).setValues(out);
  }
  return true;
}

function saveMapel(arr) {
  const clean = arr.filter(m => m.nama && String(m.nama).trim());
  const old = readLeger_();
  const rows = old.rows.map(r => {
    const nilai = clean.map(m => {
      const i = old.mapel.findIndex(n => norm_(n) === norm_(m.nama));
      return i >= 0 ? r.nilai[i] : '';
    });
    return Object.assign({}, r, { nilai });
  });
  const sh = sheet_(SH.MAP);
  sh.clear();
  sh.getRange(1, 1, 1, 3).setValues([['Kelompok', 'Mata Pelajaran', 'KKTP']]).setFontWeight('bold');
  if (clean.length) sh.getRange(2, 1, clean.length, 3).setValues(clean.map(m => [m.kelompok || 'A', String(m.nama).trim(), Number(m.kktp) || 75]));
  writeLeger_(clean.map(m => String(m.nama).trim()), rows);
  recalcAll();
  return true;
}

/* ---------- leger ---------- */
function readLeger_() {
  const sh = sheet_(SH.LEG);
  const v = sh.getDataRange().getValues();
  const head = v[0] || [];
  const iJ = head.indexOf('Jumlah');
  if (iJ < 0) return { mapel: getMapel_().map(m => m.nama), rows: [], iJ: 3 + getMapel_().length };
  const mapel = head.slice(3, iJ).map(String);
  const rows = v.slice(1).filter(r => String(r[1]).trim() !== '').map(r => ({
    nis: String(r[0]), nama: String(r[1]), kelas: String(r[2]),
    nilai: r.slice(3, iJ), sakit: r[iJ + 2], izin: r[iJ + 3], alpa: r[iJ + 4]
  }));
  return { mapel, rows, iJ };
}

function writeLeger_(mapel, rows) {
  const sh = sheet_(SH.LEG);
  sh.clear();
  const head = ['NIS', 'Nama', 'Kelas'].concat(mapel, ['Jumlah', 'Peringkat', 'Sakit', 'Izin', 'Alpa']);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setWrap(true).setBackground('#e8f0fe');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
    const data = rows.map(r => [r.nis, r.nama, r.kelas].concat(r.nilai, ['', '', r.sakit === undefined ? '' : r.sakit, r.izin === undefined ? '' : r.izin, r.alpa === undefined ? '' : r.alpa]));
    sh.getRange(2, 1, data.length, head.length).setValues(data);
  }
  sh.setFrozenRows(1); sh.setFrozenColumns(3);
  sh.setColumnWidth(2, 220);
}

function recalcAll() {
  const L = readLeger_();
  if (!L.rows.length) return;
  const jum = L.rows.map(r => {
    let s = 0, n = 0;
    r.nilai.forEach(x => { if (isNum_(x)) { s += Number(x); n++; } });
    return n ? Math.round(s * 100) / 100 : '';
  });
  const groups = {};
  L.rows.forEach((r, i) => { if (jum[i] !== '') (groups[kkey_(r.kelas)] = groups[kkey_(r.kelas)] || []).push(jum[i]); });
  const out = L.rows.map((r, i) => {
    if (jum[i] === '') return ['', ''];
    const rank = 1 + groups[kkey_(r.kelas)].filter(x => x > jum[i]).length;
    return [jum[i], rank];
  });
  sheet_(SH.LEG).getRange(2, L.iJ + 1, out.length, 2).setValues(out);
}

function mergeKelas_(list) {
  const cur = getKelas_();
  list.forEach(k => {
    const key = kkey_(k.kelas);
    const i = cur.findIndex(c => kkey_(c.kelas) === key);
    const fase = /^X{1}\s/i.test(k.kelas) ? 'E' : 'F';
    if (i < 0) cur.push({ kelas: k.kelas, wali: k.wali || '', nbm: '', bidang: k.bidang || '', program: k.program || '', kosentrasi: k.kosentrasi || '', fase });
    else ['wali', 'bidang', 'program', 'kosentrasi'].forEach(f => { if (k[f]) cur[i][f] = k[f]; });
  });
  saveKelas(cur);
}

/** Impor leger lama (hasil parsing di browser) -> sheet Leger */
function importLeger(payload) {
  ensureSheets();
  const mapel = getMapel_().map(m => m.nama);
  const idx = {};
  mapel.forEach((n, i) => idx[norm_(n)] = i);
  let skip = {};
  const rows = payload.siswa.map(s => {
    const nilai = mapel.map(() => '');
    Object.keys(s.nilai || {}).forEach(k => {
      const i = idx[norm_(k)];
      if (i === undefined) { skip[k] = 1; return; }
      if (isNum_(s.nilai[k])) nilai[i] = Number(s.nilai[k]);
    });
    return { nis: String(s.nis || ''), nama: s.nama, kelas: s.kelas, nilai, sakit: s.sakit, izin: s.izin, alpa: s.alpa };
  });
  writeLeger_(mapel, rows);
  mergeKelas_(payload.kelas || []);
  recalcAll();
  return { siswa: rows.length, mapelTidakDikenal: Object.keys(skip) };
}

/** Masukkan nilai dari template nilai (satu mapel) ke Leger */
function importNilai(p) {
  ensureSheets();
  const L = readLeger_();
  const col = L.mapel.findIndex(n => norm_(n) === norm_(p.mapel));
  if (col < 0) throw new Error('Mata pelajaran "' + p.mapel + '" tidak ada di sheet Mapel.');
  const byName = {};
  L.rows.forEach((r, i) => (byName[norm_(r.nama)] = byName[norm_(r.nama)] || []).push(i));
  let updated = 0, added = 0;
  const unmatched = [];
  const newKelas = {};
  p.rows.forEach(t => {
    if (!isNum_(t.nilai)) return;
    let c = byName[norm_(t.nama)] || [];
    if (c.length > 1) {
      const f = c.filter(i => kkey_(L.rows[i].kelas) === kkey_(t.kelas));
      if (f.length) c = f;
    }
    if (c.length) {
      L.rows[c[0]].nilai[col] = Number(t.nilai);
      updated++;
    } else if (p.tambahBaru) {
      const nilai = L.mapel.map(() => '');
      nilai[col] = Number(t.nilai);
      const kel = String(t.kelas || '').replace(/\s+/g, ' ').trim().replace(/ I$/, ' 1');
      L.rows.push({ nis: String(t.nis || ''), nama: t.nama, kelas: kel, nilai, sakit: '', izin: '', alpa: '' });
      newKelas[kel] = 1;
      added++;
    } else unmatched.push(t.nama + ' (' + t.kelas + ')');
  });
  writeLeger_(L.mapel, L.rows);
  if (added) mergeKelas_(Object.keys(newKelas).map(k => ({ kelas: k })));
  recalcAll();
  return { updated, added, unmatched };
}

function getLeger(kelas) {
  const L = readLeger_();
  const kmap = getMapel_();
  const rows = L.rows.filter(r => !kelas || kkey_(r.kelas) === kkey_(kelas));
  // jumlah & peringkat dibaca dari sheet
  const v = sheet_(SH.LEG).getDataRange().getValues().slice(1).filter(r => String(r[1]).trim() !== '');
  const extra = {};
  v.forEach(r => extra[String(r[0]) + '|' + r[1]] = { jumlah: r[L.iJ], peringkat: r[L.iJ + 1] });
  const wali = (getKelas_().find(k => kkey_(k.kelas) === kkey_(kelas)) || {}).wali || '';
  return {
    mapel: L.mapel, kktp: kmap.map(m => m.kktp), wali,
    rows: rows.map(r => Object.assign({}, r, extra[r.nis + '|' + r.nama] || {}))
  };
}

function getKelasList() {
  const L = readLeger_();
  const seen = {}, out = [];
  L.rows.forEach(r => { const k = kkey_(r.kelas); if (!seen[k]) { seen[k] = 1; out.push(r.kelas); } });
  return out.sort();
}

/** Data cetak KHS: seluruh siswa satu kelas (klien memilih yang dicetak) */
function getKHS(kelas) {
  recalcAll();
  const g = getLeger(kelas);
  const info = getKelas_().find(k => kkey_(k.kelas) === kkey_(kelas)) || {};
  return { settings: getSettings_(), mapel: getMapel_(), kelas: info, leger: g };
}

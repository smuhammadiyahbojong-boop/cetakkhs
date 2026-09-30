/* === KONFIGURASI ===
   Tempel URL aplikasi web Apps Script (berakhiran /exec) di bawah ini bila
   index.html / script.js / style.css di-host terpisah (GitHub Pages, dll).
   Biarkan kosong bila dibuka langsung dari URL Apps Script. */
const API_URL = 'https://script.google.com/macros/s/AKfycbz3duPipMZfdPX63CLmIsxcuVv7h9afHcsLhRbbRABO5aN46I2TzqEIHchsd3riNdy1/exec';

// <PURE>
const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const norm = s => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
const tingkatOf = s => {
  const m = String(s == null ? '' : s).trim().toUpperCase().match(/^[XIVL1]+/);
  if (!m) return '';
  const t = m[0].replace(/[L1]/g, 'I');
  return t === 'X' || t === 'XI' || t === 'XII' ? t : '';
};
const cleanKelas = s => {
  let v = String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  const m = v.toUpperCase().match(/^[XIVL1]+/), t = tingkatOf(v);
  if (m && t) v = t + v.slice(m[0].length);
  return v.replace(/ I$/i, ' 1');
};
const kkey = s => cleanKelas(s).toUpperCase().replace(/[\s\-]/g, '');
const cocok = (m, t) => {
  const x = String(m.tingkat || 'Semua').toUpperCase().replace(/\s/g, '');
  return !t || x === 'SEMUA' || x === '' || x.split(',').indexOf(t) >= 0;
};
const isNum = v => v !== '' && v !== null && v !== undefined && !isNaN(v) && typeof v !== 'boolean';
const sheetRows = (wb, sn) => XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null, raw: true });
const txt = v => (v == null ? '' : String(v)).trim();

/** Parse template nilai: kembalikan array {sheet, mapel, rows[{nis,nama,kelas,murni,rata}]} */
function parseTemplate(wb) {
  const out = [];
  wb.SheetNames.forEach(sn => {
    const A = sheetRows(wb, sn);
    let hr = A.findIndex(r => r.some(c => /^nama$/i.test(txt(c))));
    if (hr < 0) return;
    let mapel = '';
    for (let i = 0; i < hr; i++) {
      const r = A[i];
      if (r[0] && /^mapel/i.test(txt(r[0]))) {
        const inline = txt(r[0]).split(':')[1];
        mapel = txt(inline) || txt(r.slice(1).find(v => txt(v) && txt(v) !== ':'));
        break;
      }
    }
    const H = A[hr];
    const cNama = H.findIndex(c => /^nama$/i.test(txt(c)));
    const cNis = H.findIndex(c => /^nis/i.test(txt(c)));
    const cKelas = H.findIndex(c => /^kelas$/i.test(txt(c)));
    let cMurni = -1, cRata = -1;
    for (let i = hr; i <= hr + 2 && i < A.length; i++) A[i].forEach((c, j) => {
      if (/nilai\s*murni/i.test(txt(c))) cMurni = j;
      if (/rata/i.test(txt(c))) cRata = j;
    });
    if (cMurni < 0) cMurni = Math.max(cNama, cKelas) + 1;
    const rows = [];
    for (let i = hr + 1; i < A.length; i++) {
      const r = A[i];
      const nama = txt(r[cNama]);
      if (!nama || /^(nama|cp|nilai)/i.test(nama)) continue;
      if (cNis >= 0 && !txt(r[cNis]) && !txt(r[cKelas])) continue;
      const m = r[cMurni], ra = cRata >= 0 ? r[cRata] : null;
      rows.push({ nis: txt(r[cNis]), nama, kelas: txt(r[cKelas]),
        murni: isNum(m) ? Number(m) : null, rata: isNum(ra) ? Number(ra) : null });
    }
    if (rows.some(r => r.murni != null || r.rata != null)) out.push({ sheet: sn, mapel, rows });
  });
  return out;
}

/** Parse leger lama: {siswa[], kelas[]} */
function parseLeger(wb) {
  const info = {}, siswa = [], kelasUsed = {}, peringatan = [];
  const cariInfo = k => info[kkey(k)] ? Object.assign({ exact: true }, info[kkey(k)]) : (info[kkey(k).replace(/\d+$/, '')] ? Object.assign({ exact: false }, info[kkey(k).replace(/\d+$/, '')]) : {});
  wb.SheetNames.forEach(sn => {
    const A = sheetRows(wb, sn);
    const h = A[0] || [];
    const col = re => h.findIndex(v => re.test(txt(v)));
    if (col(/^kls$/i) >= 0 && col(/nama\s*siswa/i) >= 0) {
      const iK = col(/^kls$/i), iW = col(/wali/i), iB = col(/bidang/i), iP = col(/^program/i), iS = col(/kosentrasi|konsentrasi/i);
      A.slice(1).forEach(r => {
        const k = txt(r[iK]);
        if (!k || /^kls$/i.test(k)) return;
        info[kkey(k)] = { kelas: cleanKelas(k), wali: txt(r[iW]), bidang: txt(r[iB]), program: txt(r[iP]), kosentrasi: txt(r[iS]) };
      });
    }
  });
  wb.SheetNames.forEach(sn => {
    const A = sheetRows(wb, sn);
    if (!A.length || !/leger/i.test(txt(A[0][0]))) return;
    let kelas = '', wali = '';
    A.slice(0, 8).forEach(r => {
      if (/^kelas$/i.test(txt(r[2]))) kelas = txt(r[4]);
      if (/^wali/i.test(txt(r[2]))) wali = txt(r[4]);
    });
    const hr = A.findIndex(r => txt(r[0]).toUpperCase() === 'NO' && /nis/i.test(txt(r[1])));
    if (hr < 0 || !kelas) return;
    const nr = A.findIndex((r, i) => i > hr && r.slice(4).filter(v => txt(v)).length >= 5);
    if (nr < 0) return;
    const mCols = []; let cS = -1, cI = -1, cA = -1;
    A[nr].forEach((v, c) => {
      const n = txt(v); if (c < 4 || !n) return;
      if (/sakit/i.test(n)) cS = c; else if (/izin|ijin/i.test(n)) cI = c; else if (/tanpa|alpa/i.test(n)) cA = c; else mCols.push([c, n]);
    });
    // nama kelas: utamakan nama sheet (sel KELAS sering salah salin), cadangan sel KELAS
    const kelasName = cleanKelas(tingkatOf(sn) && /\s/.test(sn.trim()) ? sn : kelas);
    const inf = cariInfo(kelasName);
    const waliPakai = inf.wali || wali;
    const fb = inf.exact === false;
    if (kelasUsed[kkey(kelasName)]) peringatan.push(`Kelas "${kelasName}" muncul di lebih dari satu sheet (${sn}); siswa digabung.`);
    if (inf.exact !== false && inf.wali && wali && norm(inf.wali) !== norm(wali)) peringatan.push(`${kelasName}: wali di sheet kelas "${wali}" berbeda dengan sheet rekap "${inf.wali}" (dipakai: ${inf.wali}).`);
    kelasUsed[kkey(kelasName)] = { kelas: kelasName, wali: waliPakai, waliSheet: wali, fb, bidang: inf.bidang || '', program: inf.program || '', kosentrasi: inf.kosentrasi || '' };
    for (let i = nr + 1; i < A.length; i++) {
      const r = A[i];
      if (!isNum(r[0]) || !txt(r[2])) continue;
      const nilai = {};
      mCols.forEach(([c, n]) => { if (isNum(r[c]) && Number(r[c]) > 0) nilai[n] = Number(r[c]); });
      const ab = c => (c >= 0 && isNum(r[c]) ? Number(r[c]) : '');
      siswa.push({ nis: txt(r[1]), nama: txt(r[2]), kelas: kelasName, nilai, sakit: ab(cS), izin: ab(cI), alpa: ab(cA) });
    }
  });
  // rekap hanya menulis "XII TSM" (tanpa nomor) untuk >1 kelas: pakai wali dari sheet kelas masing-masing
  const list = Object.values(kelasUsed);
  list.forEach(k => {
    if (!k.fb) return;
    const dasar = kkey(k.kelas).replace(/\d+$/, '');
    if (list.filter(x => kkey(x.kelas).replace(/\d+$/, '') === dasar).length > 1 && k.waliSheet) {
      if (norm(k.wali) !== norm(k.waliSheet)) peringatan.push(`${k.kelas}: rekap hanya menulis satu wali untuk beberapa kelas; dipakai wali dari sheet kelas: ${k.waliSheet}.`);
      k.wali = k.waliSheet;
    }
  });
  list.forEach(k => { delete k.waliSheet; delete k.fb; });
  return { siswa, kelas: list, peringatan };
}

/** Cari mapel target dari nama di template */
function guessMapel(name, list) {
  const n = norm(name); if (!n) return '';
  let m = list.find(x => norm(x.nama) === n) || list.find(x => norm(x.nama).includes(n) || n.includes(norm(x.nama)));
  return m ? m.nama : '';
}

/** Bangun HTML satu halaman KHS */
function khsHTML(d, s, idx) {
  const st = d.settings, k = d.kelas || {};
  const tandai = String(st.TANDAI_REMIDI || 'Ya').toLowerCase() === 'ya';
  const groups = [['A', 'KELOMPOK MATA PELAJARAN UMUM'], ['B', 'KELOMPOK MATA PELAJARAN KEJURUAN'], ['C', 'KELOMPOK MATA PELAJARAN CIRI KHUSUS']];
  let no = 0, body = '';
  groups.forEach(([g, title], gi) => {
    const items = d.mapel.map((m, i) => [m, i]).filter(x => x[0].kelompok === g);
    if (!items.length) return;
    body += `<tr class="gr"><td class="c">${g}</td><td colspan="3">${title}</td></tr>`;
    items.forEach(([m, i]) => {
      no++;
      const v = s.nilai[i];
      const has = isNum(v);
      const low = has && tandai && Number(v) < m.kktp;
      body += `<tr><td class="c">${no}.</td><td>${esc(m.nama)}</td><td class="c">${has ? Math.round(Number(v)) + (low ? ' *' : '') : '-'}</td><td class="c">${m.kktp}</td></tr>`;
    });
    if (gi < 2) body += `<tr class="sp"><td colspan="4"></td></tr>`;
  });
  const img = (u, h) => u ? `<img src="${u}" style="max-height:${h}">` : '';
  const ab = v => (isNum(v) ? v : '');
  return `<div class="khs-page">
  <div class="kop">
    <div class="lg">${img(st.LOGO_KIRI, '2.2cm')}</div>
    <div class="tx"><div class="b1">${esc(st.BARIS_1)}</div><div class="b1">${esc(st.BARIS_2)}</div>
      <div class="nm">${esc(st.NAMA_SEKOLAH)}</div>
      <div class="sm">${esc(st.PROGRAM_STUDI)}</div><div class="sm">${esc(st.ALAMAT)}</div><div class="sm">${esc(st.KONTAK)}</div></div>
    <div class="lg">${img(st.LOGO_KANAN, '2.2cm')}</div>
  </div>
  <div class="jd">${esc(st.JUDUL)}</div>
  <div class="ta">Tahun Pelajaran ${esc(st.TAHUN_PELAJARAN)}</div>
  <table class="inf">
    <tr><td width="17%">Nama Siswa</td><td width="33%">: <b>${esc(s.nama)}</b></td><td width="17%">Semester</td><td>: ${esc(st.SEMESTER)}</td></tr>
    <tr><td>Nomor Induk</td><td>: ${esc(s.nis)}</td><td>Bidang Keahlian</td><td>: ${esc(k.bidang)}</td></tr>
    <tr><td>Kelas</td><td>: ${esc(s.kelas)}</td><td>Program Keahlian</td><td>: ${esc(k.program)}</td></tr>
    <tr><td>Fase</td><td>: ${esc(k.fase)}</td><td>Kosentrasi</td><td>: ${esc(k.kosentrasi)}</td></tr>
  </table>
  <table class="kh">
    <tr><th rowspan="2" width="6%">NO</th><th rowspan="2">MATA PELAJARAN</th><th colspan="2">Rincian Nilai Sumatif</th></tr>
    <tr><th width="16%">Nilai Akhir</th><th width="12%">KKTP **)</th></tr>
    ${body}
  </table>
  <table class="kh" style="margin-top:8px">
    <tr><th style="text-align:left">ABSENSI KEHADIRAN ***)</th><th width="14%">SAKIT</th><th width="14%">IJIN</th><th width="14%">ALPA</th></tr>
    <tr><td><b>Jumlah Kehadiran</b></td><td class="c">${ab(s.sakit)}</td><td class="c">${ab(s.izin)}</td><td class="c">${ab(s.alpa)}</td></tr>
    <tr><td><b>PERINGKAT KELAS</b></td><td colspan="3" class="c"><b>${ab(s.peringkat)}</b></td></tr>
  </table>
  <div style="margin-top:14px;text-align:right;padding-right:1cm">${esc(st.TEMPAT)}, ${esc(st.TANGGAL)}</div>
  <table class="sg">
    <tr><td>Mengetahui,<br>Kepala Sekolah</td><td><br>Orang Tua / Wali Siswa</td><td><br>Wali Kelas</td></tr>
    <tr><td style="height:2cm;vertical-align:middle">${img(st.TTD_KEPSEK, '1.8cm')}</td><td></td><td></td></tr>
    <tr><td><b><u>${esc(st.KEPSEK)}</u></b><br>NBM. ${esc(st.NBM_KEPSEK)}</td><td>( ............................. )</td><td><b><u>${esc(k.wali)}</u></b><br>NBM. ${esc(k.nbm)}</td></tr>
  </table>
  <div class="fn">*) Nilai di bawah KKM untuk segera melaksanakan REMIDI<br>**) KKTP adalah singkatan dari kriteria ketercapaian tujuan pembelajaran<br>***) Jumlah Kehadiran didapat dari Rekap Absensi selama 3 bulan proses pembelajaran</div>
</div>`;
}

/** Bangun HTML leger satu kelas */
function legerHTML(st, kelas, g) {
  const head = g.mapel.map(m => `<th style="min-width:1.6cm">${esc(m)}</th>`).join('');
  const rows = g.rows.map((r, i) => `<tr><td>${i + 1}</td><td>${esc(r.nis)}</td><td class="n">${esc(r.nama)}</td><td>${esc(r.kelas)}</td>` +
    r.nilai.map(v => `<td>${isNum(v) ? Math.round(v * 100) / 100 : ''}</td>`).join('') +
    `<td>${isNum(r.jumlah) ? r.jumlah : ''}</td><td>${isNum(r.peringkat) ? r.peringkat : ''}</td><td>${isNum(r.sakit) ? r.sakit : ''}</td><td>${isNum(r.izin) ? r.izin : ''}</td><td>${isNum(r.alpa) ? r.alpa : ''}</td></tr>`).join('');
  return `<div class="leg-page"><div style="text-align:center;font-weight:bold;font-size:11pt">DAFTAR KUMPULAN NILAI (LEGER) ${esc((st.JUDUL || '').replace(/^KARTU HASIL BELAJAR\s*/i, ''))}</div>
  <div style="text-align:center;font-weight:bold">${esc(st.NAMA_SEKOLAH)} TAHUN PELAJARAN ${esc(st.TAHUN_PELAJARAN)}</div>
  <div style="margin:6px 0"><b>KELAS</b> : ${esc(kelas)} &nbsp;&nbsp; <b>SEMESTER</b> : ${esc(st.SEMESTER)} &nbsp;&nbsp; <b>WALI KELAS</b> : ${esc(g.wali)}</div>
  <table class="lg"><tr><th>NO</th><th>NIS</th><th>Nama Siswa</th><th>Kelas</th>${head}<th>Jumlah</th><th>Peringkat</th><th>S</th><th>I</th><th>A</th></tr>${rows}</table></div>`;
}
// </PURE>

/* ================= UI ================= */
const $ = id => document.getElementById(id);

const apiCall = async (fn, args, retry) => {
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // tanpa preflight CORS
    body: JSON.stringify({ fn, args, pin: localStorage.getItem('khs_pin') || '' })
  });
  const j = await res.json();
  if (j.auth && !retry) {
    const p = prompt('Masukkan PIN akses sistem KHS:');
    if (p === null) throw new Error('PIN diperlukan.');
    localStorage.setItem('khs_pin', p);
    return apiCall(fn, args, true);
  }
  if (!j.ok) { if (j.auth) localStorage.removeItem('khs_pin'); throw new Error(j.error || 'Gagal memanggil server'); }
  return j.data;
};
const gs = (fn, ...a) => API_URL ? apiCall(fn, a) :
  new Promise((ok, er) => google.script.run.withSuccessHandler(ok).withFailureHandler(er)[fn](...a));
let S = null, KHS = null, legerParsed = null, tplParsed = [];
const msg = (el, t, c) => { $(el).innerHTML = `<div class="msg ${c || 'in'}">${esc(t)}</div>`; };
const busy = (b, on) => { b.disabled = on; };

const TABS = [['mulai', '1. Impor Leger'], ['nilai', '2. Unggah Nilai'], ['leger', '3. Leger'], ['khs', '4. Cetak KHS'], ['atur', 'Pengaturan']];
function tab(t) {
  document.querySelectorAll('.sec').forEach(s => s.classList.toggle('on', s.id === 't-' + t));
  document.querySelectorAll('#nav button').forEach(b => b.classList.toggle('on', b.dataset.t === t));
  if (t === 'leger' || t === 'khs') fillKelas();
}
$('nav').innerHTML = TABS.map(t => `<button data-t="${t[0]}">${t[1]}</button>`).join('');
$('nav').onclick = e => { if (e.target.dataset.t) tab(e.target.dataset.t); };

async function load() {
  S = await gs('getState');
  $('lgSheet').href = S.url;
  renderAtur();
}
async function fillKelas() {
  const list = await gs('getKelasList');
  ['lgKelas', 'khKelas'].forEach(id => {
    const cur = $(id).value;
    $(id).innerHTML = list.map(k => `<option>${esc(k)}</option>`).join('');
    if (cur) $(id).value = cur;
  });
}

/* --- 1. impor leger --- */
$('fLeger').onchange = async e => {
  const f = e.target.files[0]; if (!f) return;
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
    legerParsed = parseLeger(wb);
    $('legerInfo').innerHTML = `<div class="msg in">Terbaca: ${legerParsed.siswa.length} siswa di ${legerParsed.kelas.length} kelas\n` +
      legerParsed.kelas.map(k => `${esc(k.kelas)} \u2014 wali: ${esc(k.wali || '(kosong)')}`).join('\n') + '</div>' +
      (legerParsed.peringatan.length ? `<div class="msg er">Periksa:\n${esc(legerParsed.peringatan.join('\n'))}</div>` : '');
    $('btnLeger').disabled = !legerParsed.siswa.length;
  } catch (er) { msg('legerInfo', 'Gagal membaca file: ' + er.message, 'er'); }
};
$('btnLeger').onclick = async function () {
  if (!confirm('Data siswa untuk kelas yang ada di file ini akan diganti. Kelas lain tidak tersentuh. Lanjutkan?')) return;
  busy(this, true);
  try {
    const pay = JSON.parse(JSON.stringify(legerParsed));
    if (!$('cNilaiLama').checked) pay.siswa.forEach(x => x.nilai = {});
    const r = await gs('importLeger', pay);
    msg('legerMsg', `Berhasil: ${r.siswa} siswa diimpor (${r.dipertahankan} siswa kelas lain dipertahankan).` + (r.mapelTidakDikenal.length ? `\nMapel tidak cocok (dilewati): ${r.mapelTidakDikenal.join(', ')}` : '') + `\nCek wali kelas di tab Pengaturan:\n${r.kelas.join('\n')}`, 'ok');
    await load();
  } catch (er) { msg('legerMsg', er.message || String(er), 'er'); }
  busy(this, false);
};

/* --- 2. upload nilai --- */
$('fNilai').onchange = async e => {
  tplParsed = [];
  for (const f of e.target.files) {
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
      parseTemplate(wb).forEach(p => tplParsed.push(Object.assign(p, { file: f.name })));
    } catch (er) { tplParsed.push({ file: f.name, error: er.message, rows: [] }); }
  }
  renderTpl();
};
function renderTpl() {
  if (!S) return;
  $('nilaiList').innerHTML = tplParsed.map((p, i) => {
    if (p.error || !p.rows.length) return `<div class="card"><b>${esc(p.file)}</b><div class="msg er">${esc(p.error || 'Tidak ada data nilai terbaca')}</div></div>`;
    const tk = [...new Set(p.rows.map(r => tingkatOf(r.kelas)).filter(Boolean))];
    const opsi = tk.length === 1 ? S.mapel.filter(m => cocok(m, tk[0])) : S.mapel;
    const g = guessMapel(p.mapel, opsi);
    const kelasList = [...new Set(p.rows.map(r => r.kelas))].join(', ');
    return `<div class="card" data-i="${i}"><b>${esc(p.file)}</b> <small>(${esc(p.sheet)})</small>
      <div>Mapel terbaca: <b>${esc(p.mapel || '-')}</b> &middot; ${p.rows.length} siswa &middot; ${p.rows.filter(r => r.murni == null).length} belum ada nilai</div>
      <div>Kelas: ${esc(kelasList)}${tk.length === 1 ? ' &middot; daftar mapel disesuaikan untuk kelas ' + tk[0] : ''}</div>
      <div class="row"><label style="margin:0">Masuk ke mapel:</label>
        <select class="tMapel"><option value="">-- pilih --</option>${opsi.map(m => `<option ${m.nama === g ? 'selected' : ''}>${esc(m.nama)}</option>`).join('')}</select>
        <label style="margin:0">Sumber nilai:</label>
        <select class="tSumber"><option value="murni">Nilai Murni</option><option value="rata">Rata-rata</option></select>
        <label style="margin:0"><input type="checkbox" class="tBaru" ${S.jumlahSiswa ? '' : 'checked'}> tambahkan siswa yang belum ada di Leger</label>
        <button class="p tGo">Masukkan ke Leger</button></div>
      <div class="tMsg"></div></div>`;
  }).join('');
}
$('nilaiList').onclick = async e => {
  if (!e.target.classList.contains('tGo')) return;
  const card = e.target.closest('.card'), p = tplParsed[card.dataset.i];
  const mapel = card.querySelector('.tMapel').value, sumber = card.querySelector('.tSumber').value;
  const out = card.querySelector('.tMsg');
  if (!mapel) { out.innerHTML = '<div class="msg er">Pilih mata pelajaran tujuan.</div>'; return; }
  const rows = p.rows.map(r => ({ nis: r.nis, nama: r.nama, kelas: r.kelas, nilai: sumber === 'rata' ? r.rata : r.murni }));
  busy(e.target, true);
  try {
    const r = await gs('importNilai', { mapel, rows, tambahBaru: card.querySelector('.tBaru').checked });
    out.innerHTML = `<div class="msg ok">Masuk: ${r.updated} nilai diperbarui, ${r.added} siswa baru ditambahkan.</div>` +
      (r.salahTingkat ? `<div class="msg er">${r.salahTingkat} siswa dilewati karena mapel ini tidak berlaku untuk tingkat kelasnya.</div>` : '') +
      (r.unmatched.length ? `<div class="msg er">Nama tidak ditemukan di Leger (${r.unmatched.length}):\n${esc(r.unmatched.join('\n'))}</div>` : '');
    await load();
  } catch (er) { out.innerHTML = `<div class="msg er">${esc(er.message || er)}</div>`; }
  busy(e.target, false);
};

/* --- 3. leger --- */
async function muatLeger() {
  const k = $('lgKelas').value; if (!k) return;
  const g = await gs('getLeger', k);
  window._lg = g;
  $('lgTabel').innerHTML = `<tr><th>No</th><th>NIS</th><th>Nama</th>${g.mapel.map(m => `<th>${esc(m)}</th>`).join('')}<th>Jumlah</th><th>Peringkat</th><th>S</th><th>I</th><th>A</th></tr>` +
    g.rows.map((r, i) => `<tr><td>${i + 1}</td><td>${esc(r.nis)}</td><td>${esc(r.nama)}</td>` +
      r.nilai.map((v, j) => `<td class="${isNum(v) && v < g.kktp[j] ? 'low' : ''}">${isNum(v) ? v : ''}</td>`).join('') +
      `<td>${esc(r.jumlah)}</td><td>${esc(r.peringkat)}</td><td>${esc(r.sakit)}</td><td>${esc(r.izin)}</td><td>${esc(r.alpa)}</td></tr>`).join('');
}
$('lgMuat').onclick = muatLeger;
$('lgHitung').onclick = async function () { busy(this, true); await gs('recalcAll'); await muatLeger(); busy(this, false); };
$('lgCetak').onclick = async () => {
  await muatLeger();
  $('pages').innerHTML = legerHTML(S.settings, $('lgKelas').value, window._lg);
  openPrev('landscape', 'Leger ' + $('lgKelas').value);
};

/* --- 4. KHS --- */
$('khMuat').onclick = async function () {
  const k = $('khKelas').value; if (!k) return;
  busy(this, true);
  KHS = await gs('getKHS', k);
  drawKh();
  busy(this, false);
};
function drawKh() {
  if (!KHS) return;
  const q = norm($('khCari').value);
  $('khTabel').innerHTML = '<tr><th></th><th>No</th><th>NIS</th><th>Nama</th><th>Jumlah</th><th>Peringkat</th></tr>' +
    KHS.leger.rows.map((r, i) => (!q || norm(r.nama).includes(q)) ? `<tr><td><input type="checkbox" class="khc" value="${i}" checked></td><td>${i + 1}</td><td>${esc(r.nis)}</td><td>${esc(r.nama)}</td><td>${esc(r.jumlah)}</td><td>${esc(r.peringkat)}</td></tr>` : '').join('');
}
$('khCari').oninput = drawKh;
$('khAll').onclick = () => document.querySelectorAll('.khc').forEach(c => c.checked = true);
$('khNone').onclick = () => document.querySelectorAll('.khc').forEach(c => c.checked = false);
$('khCetak').onclick = () => {
  const ids = [...document.querySelectorAll('.khc:checked')].map(c => Number(c.value));
  if (!ids.length) return alert('Pilih minimal satu siswa.');
  const d = { settings: KHS.settings, mapel: KHS.mapel, kelas: KHS.kelas };
  $('pages').innerHTML = ids.map(i => {
    const r = KHS.leger.rows[i];
    return khsHTML(d, { nis: r.nis, nama: r.nama, kelas: r.kelas, nilai: r.nilai, sakit: r.sakit, izin: r.izin, alpa: r.alpa, peringkat: r.peringkat });
  }).join('');
  openPrev('portrait', ids.length + ' KHS');
};

function openPrev(mode, info) {
  $('pageStyle').textContent = mode === 'landscape'
    ? '@page{size:14in 8.5in;margin:1cm}'
    : '@page{size:8.5in 14in;margin:1.2cm 1cm 1.2cm 1.6cm}';
  $('pInfo').textContent = info + ' \u2014 ukuran kertas F4/Legal';
  document.body.classList.add('previewing');
  window.scrollTo(0, 0);
}
$('pClose').onclick = () => document.body.classList.remove('previewing');

/* --- pengaturan --- */
const LBL = { NAMA_SEKOLAH: 'Nama sekolah', BARIS_1: 'Kop baris 1', BARIS_2: 'Kop baris 2', PROGRAM_STUDI: 'Program studi (kop)', ALAMAT: 'Alamat', KONTAK: 'Email / website',
  JUDUL: 'Judul KHS', TAHUN_PELAJARAN: 'Tahun pelajaran', SEMESTER: 'Semester', TEMPAT: 'Tempat', TANGGAL: 'Tanggal cetak', KEPSEK: 'Kepala sekolah', NBM_KEPSEK: 'NBM kepala sekolah', TANDAI_REMIDI: 'Tandai nilai di bawah KKTP dengan * (Ya/Tidak)' };
function renderAtur() {
  $('fSet').innerHTML = Object.keys(LBL).map(k => `<div><label>${LBL[k]}</label><input type="text" data-k="${k}" value="${esc(S.settings[k])}"></div>`).join('');
  $('logoPrev').innerHTML = ['LOGO_KIRI', 'LOGO_KANAN', 'TTD_KEPSEK'].map(k => S.settings[k] ? `<div><small>${k}</small><br><img src="${S.settings[k]}" style="height:50px;border:1px solid #ccc"></div>` : '').join('');
  $('tMapel').innerHTML = '<tr><th>Kelompok</th><th>Mata Pelajaran</th><th>KKTP</th><th>Tingkat</th><th></th></tr>' + S.mapel.map(mapelRow).join('');
  $('tKelas').innerHTML = '<tr><th>Kelas</th><th>Wali Kelas</th><th>NBM Wali</th><th>Bidang Keahlian</th><th>Program Keahlian</th><th>Kosentrasi</th><th>Fase</th><th></th></tr>' + S.kelas.map(kelasRow).join('');
  renderTpl();
}
const mapelRow = m => `<tr><td><select class="mK">${['A', 'B', 'C'].map(g => `<option ${g === m.kelompok ? 'selected' : ''}>${g}</option>`).join('')}</select></td><td><input type="text" class="mN" value="${esc(m.nama)}" style="min-width:280px"></td><td><input type="number" class="mT" value="${m.kktp}" style="width:70px"></td><td><input type="text" class="mG" value="${esc(m.tingkat || 'Semua')}" style="width:110px" placeholder="X,XI,XII"></td><td><button class="s" onclick="this.closest('tr').remove()">Hapus</button></td></tr>`;
const kelasRow = k => `<tr>${['kelas', 'wali', 'nbm', 'bidang', 'program', 'kosentrasi', 'fase'].map(f => `<td><input type="text" value="${esc(k[f])}"></td>`).join('')}<td><button class="s" onclick="this.closest('tr').remove()">Hapus</button></td></tr>`;
$('addMapel').onclick = () => $('tMapel').insertAdjacentHTML('beforeend', mapelRow({ kelompok: 'A', nama: '', kktp: 75, tingkat: 'Semua' }));
$('addKelas').onclick = () => $('tKelas').insertAdjacentHTML('beforeend', kelasRow({ kelas: '', wali: '', nbm: '', bidang: '', program: '', kosentrasi: '', fase: 'E' }));

document.querySelectorAll('[data-logo]').forEach(inp => inp.onchange = e => {
  const f = e.target.files[0]; if (!f) return;
  const key = inp.dataset.logo, img = new Image();
  img.onload = () => {
    const max = key === 'TTD_KEPSEK' ? 260 : 150, sc = Math.min(1, max / Math.max(img.width, img.height));
    const c = document.createElement('canvas'); c.width = Math.round(img.width * sc); c.height = Math.round(img.height * sc);
    c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
    let u = c.toDataURL('image/png');
    if (u.length > 45000) u = c.toDataURL('image/jpeg', 0.7);
    if (u.length > 48000) return alert('Gambar terlalu besar, gunakan gambar yang lebih sederhana.');
    S.settings[key] = u; renderAtur();
  };
  img.src = URL.createObjectURL(f);
});

$('simpanAtur').onclick = async function () {
  busy(this, true);
  try {
    document.querySelectorAll('#fSet input').forEach(i => S.settings[i.dataset.k] = i.value);
    const mapel = [...$('tMapel').querySelectorAll('tr')].slice(1).map(tr => ({ kelompok: tr.querySelector('.mK').value, nama: tr.querySelector('.mN').value, kktp: tr.querySelector('.mT').value, tingkat: tr.querySelector('.mG').value }));
    const kelas = [...$('tKelas').querySelectorAll('tr')].slice(1).map(tr => { const i = tr.querySelectorAll('input'); return { kelas: i[0].value, wali: i[1].value, nbm: i[2].value, bidang: i[3].value, program: i[4].value, kosentrasi: i[5].value, fase: i[6].value }; });
    await gs('saveSettings', S.settings); await gs('saveKelas', kelas); await gs('saveMapel', mapel);
    await load(); msg('aturMsg', 'Pengaturan tersimpan.', 'ok');
  } catch (er) { msg('aturMsg', er.message || String(er), 'er'); }
  busy(this, false);
};

load().then(() => tab(S.jumlahSiswa ? 'nilai' : 'mulai')).catch(er => document.querySelector('main').insertAdjacentHTML('afterbegin', `<div class="msg er">${esc(er.message || er)}</div>`));

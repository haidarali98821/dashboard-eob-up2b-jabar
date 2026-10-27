/* eob-reader: membaca file Lampiran EOB UP2B Jawa Barat menjadi data dashboard.
   Semua data dicari berdasarkan label (judul tabel, nama bulan, judul kolom), bukan nomor baris,
   supaya tahan terhadap baris yang bertambah atau bergeser antar bulan.
   Setiap modul dibaca terpisah: bila satu modul gagal, modul lain tetap jalan, dan alasannya dicatat. */
(function (global) {
  'use strict';

  const BULAN = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const BLN = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  const TOKEN = {JAN:0,JANUARI:0,FEB:1,FEBRUARI:1,MAR:2,MARET:2,APR:3,APRIL:3,MEI:4,MAY:4,JUN:5,JUNI:5,JUL:6,JULI:6,
    AGU:7,AGS:7,AUG:7,AGT:7,AGUSTUS:7,SEP:8,SEPT:8,SEPTEMBER:8,OKT:9,OCT:9,OKTOBER:9,OKOTBER:9,NOV:10,NOP:10,NOVEMBER:10,DES:11,DEC:11,DESEMBER:11};
  const MON_EN = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Mei:4,Jun:5,Jul:6,Aug:7,Agu:7,Agt:7,Ags:7,Sep:8,Oct:9,Okt:9,Nov:10,Dec:11,Des:11};

  const norm = s => String(s ?? '').toLowerCase().replace(/\s+/g, ' ').trim();
  const squash = s => String(s ?? '').toLowerCase().replace(/[\s_]+/g, '');
  function monthOf(v) {
    if (typeof v !== 'string') return -1;
    const t = v.trim().toUpperCase();
    if (!/^[A-Z]{3,9}$/.test(t)) return -1;
    return TOKEN[t] ?? -1;
  }
  const r2 = (x, d = 2) => (typeof x === 'number' && isFinite(x)) ? Math.round(x * 10 ** d) / 10 ** d : null;
  const isoDate = v => v instanceof Date ? v.toISOString().slice(0, 10) : null;
  const timeStr = v => v && v.time ? v.time : (v instanceof Date ? v.toISOString().slice(11, 16) : (typeof v === 'string' ? v.slice(0, 5) : null));

  function fail(msg) { const e = new Error(msg); e.userFacing = true; throw e; }

  // Mencari sheet berdasarkan nama tanpa memedulikan spasi dan huruf besar/kecil.
  function findSheet(wb, test, label) {
    const s = wb.sheets.find(sh => test(squash(sh.name)));
    if (!s) fail(`Sheet "${label}" tidak ditemukan.`);
    return s;
  }
  // Mencari sel pertama yang memenuhi syarat.
  function findCell(sh, test, opt = {}) {
    const r0 = opt.r0 || 1, r1 = Math.min(opt.r1 || sh.maxRow, sh.maxRow), c0 = opt.c0 || 1, c1 = Math.min(opt.c1 || sh.maxCol, sh.maxCol);
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) { const v = sh.get(r, c); if (v != null && test(v, r, c)) return { r, c }; }
    return null;
  }
  function findAll(sh, test, opt = {}) {
    const out = []; const r0 = opt.r0 || 1, r1 = Math.min(opt.r1 || sh.maxRow, sh.maxRow), c0 = opt.c0 || 1, c1 = Math.min(opt.c1 || sh.maxCol, sh.maxCol);
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) { const v = sh.get(r, c); if (v != null && test(v, r, c)) out.push({ r, c }); }
    return out;
  }
  const txtIs = want => v => typeof v === 'string' && norm(v) === norm(want);
  const txtStarts = want => v => typeof v === 'string' && norm(v).startsWith(norm(want));
  const txtHas = want => v => typeof v === 'string' && norm(v).includes(norm(want));

  /* ---------------- Periode ---------------- */
  function readPeriode(wb) {
    for (const sh of wb.sheets) {
      if (!/pengesahan|wajibisi/.test(squash(sh.name))) continue;
      const hit = findCell(sh, v => typeof v === 'string' && /^[A-Za-z]+\s+\d{4}$/.test(v.trim()) && monthOf(v.trim().split(/\s+/)[0]) >= 0);
      if (hit) { const [m, y] = sh.text(hit.r, hit.c).split(/\s+/); return { bulan: monthOf(m), tahun: +y }; }
    }
    fail('Periode laporan (contoh: "SEPTEMBER 2026") tidak ditemukan di Lembar Pengesahan.');
  }

  /* ---------------- 1. Ringkasan ---------------- */
  function readRingkasan(wb) {
    const sh = findSheet(wb, n => n === 'rangkuman', 'Rangkuman');
    const rows = [];
    for (let r = 1; r <= sh.maxRow; r++) rows.push({ r, label: norm(sh.text(r, 2)), v: sh.get(r, 3) });
    const after = (anchor) => { const i = rows.findIndex(x => x.label.startsWith(anchor)); return i < 0 ? rows : rows.slice(i); };
    const val = (label, from) => { const x = (from || rows).find(x => x.label.startsWith(label)); return x ? x.v : null; };
    const bp = val('beban puncak (mw)');
    if (typeof bp !== 'number') fail('Baris "Beban Puncak (MW)" tidak ditemukan di sheet Rangkuman.');
    const eks = after('ekskursi');
    const g = v => typeof v === 'number' ? v : null;
    return {
      bp_tanggal: isoDate(val('tanggal')), bp_jam: timeStr(val('jam')), bp: r2(bp), rekor: r2(g(val('rekor'))),
      growth_rekor: g(val('growth')) != null ? r2(val('growth') * 100) : null,
      lf: r2(g(val('faktor beban (%)'))), lf_kum: r2(g(val('faktor beban kumulatif'))),
      ens: r2(g(val('total ens'))), ens_ufls: r2(g(val('- ufls'))), ens_ols: r2(g(val('- ols'))), ens_mls: r2(g(val('- mls'))),
      ekskursi: g(val('jumlah', eks)), energi_gwh: r2(g(val('energi (gwh)'))), htgp: g(val('htgp')),
      pht60: g((rows.find(x => x.label.includes('>60%')) || {}).v), trf80: g((rows.find(x => x.label.includes('trafo distribusi')) || {}).v),
      ibt80: g((rows.find(x => x.label.includes('ibt berbeban')) || {}).v),
      scadatel: g(val('jumlah gangguan scadatel (kali)')), scadatel_kum: g(val('jumlah gangguan scadatel kumulatif'))
    };
  }

  /* ---------------- 2. Tren beban puncak ---------------- */
  function readTrenBP(wb, P) {
    const sh = findSheet(wb, n => n === 'tabel1', 'Tabel 1');
    const heads = findAll(sh, v => monthOf(v) >= 0, { c0: 2, c1: 2 });
    if (heads.length < P.bulan + 1) fail('Blok bulan di Tabel 1 tidak lengkap sampai bulan laporan.');
    const out = [];
    for (const h of heads) {
      const m = monthOf(sh.get(h.r, h.c)); if (m > P.bulan) continue;
      const rm = findCell(sh, txtIs('Malam'), { r0: h.r + 1, r1: h.r + 5, c0: 1, c1: 1 });
      const rs = findCell(sh, txtIs('Siang'), { r0: h.r + 1, r1: h.r + 6, c0: 1, c1: 1 });
      if (!rm || !rs) fail(`Baris Malam/Siang untuk ${BULAN[m]} tidak ditemukan di Tabel 1.`);
      out[m] = { bulan: BLN[m], malam_prev: r2(sh.num(rm.r, 2)), malam_cur: r2(sh.num(rm.r, 3)), siang_prev: r2(sh.num(rs.r, 2)), siang_cur: r2(sh.num(rs.r, 3)),
        tgl_malam: isoDate(sh.get(rm.r, 5)), jam_malam: timeStr(sh.get(rm.r, 6)), lf: null };
    }
    for (let m = 0; m <= P.bulan; m++) if (!out[m]) fail(`Data beban puncak ${BULAN[m]} tidak ditemukan di Tabel 1.`);
    // Faktor beban dari sheet Load Factor (opsional)
    try {
      const lf = findSheet(wb, n => n === 'loadfactor', 'Load Factor');
      const hd = findCell(lf, txtIs('LF (%)'));
      if (hd) for (let r = hd.r + 1; r <= lf.maxRow; r++) {
        const m = monthOf(lf.get(r, 1)); if (m < 0 || m > P.bulan) continue;
        const v = lf.num(r, hd.c); if (v != null) out[m].lf = r2(v);
      }
    } catch (e) { /* faktor beban tidak wajib */ }
    return out;
  }

  /* ---------------- 3. Sub sistem ---------------- */
  function readSubsistem(wb, P, notes) {
    const sh = findSheet(wb, n => n.includes('bpkoinsiden'), 'Tabel BP Koinsiden APB Per SS');
    const heads = findAll(sh, (v, r, c) => monthOf(v) >= 0 && findCell(sh, txtIs('Sub-sistem'), { r0: r, r1: r, c0: c + 1, c1: c + 20 }) != null);
    const res = {};
    for (const h of heads) {
      const m = monthOf(sh.get(h.r, h.c)); if (m > P.bulan) continue;
      const cols = findAll(sh, txtIs('Sub-sistem'), { r0: h.r, r1: h.r, c0: h.c + 1 }).map(x => x.c);
      let best = null;
      for (const nc of cols) {
        const hdr = k => { const f = findCell(sh, txtHas(k), { r0: h.r, r1: h.r, c0: nc + 1, c1: nc + 6 }); return f ? f.c : null; };
        const cD = hdr('DMN'), cB = hdr('Beban'), cC = hdr('Cadangan'), cO = hdr('Outage');
        if (!cD || !cB) continue;
        const rows = [];
        for (let r = h.r + 1; r <= sh.maxRow; r++) {
          const nm = sh.get(r, nc); const no = sh.num(r, nc - 1);
          if (nm == null || no == null) break;
          rows.push({ nama: String(nm).replace(/\s+/g, ' ').trim(), dmn: r2(sh.num(r, cD)), beban: r2(sh.num(r, cB)), outage: r2(sh.num(r, cO) || 0), cadangan_excel: r2(sh.num(r, cC)) });
        }
        if (!best || rows.length >= best.length) best = rows;
      }
      if (best && best.length) res[BLN[m]] = best;
    }
    if (!res[BLN[P.bulan]]) fail(`Blok ${BULAN[P.bulan]} tidak ditemukan di Tabel BP Koinsiden APB Per SS.`);
    // Catatan di atas tabel yang menyebut bulan lain
    const note = sh.text(1, 2);
    if (note) {
      const mm = note.match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
      if (mm && monthOf(mm[2]) >= 0 && (monthOf(mm[2]) !== P.bulan || +mm[3] !== P.tahun))
        notes.push({ level: 'warn', modul: 'Pasokan', teks: `Catatan di atas Tabel BP Koinsiden (sel B1) menyebut koinsiden ${mm[1]} ${BULAN[monthOf(mm[2])]} ${mm[3]}, padahal periode laporan ${BULAN[P.bulan]} ${P.tahun}. Mohon dicek apakah labelnya belum diperbarui.` });
    }
    const cur = res[BLN[P.bulan]];
    const bad = cur.filter(s => s.cadangan_excel != null && Math.abs((s.dmn - s.outage - s.beban) - s.cadangan_excel) > 1);
    if (bad.length) notes.push({ level: 'warn', modul: 'Pasokan', teks: `${bad.length} sub sistem punya kolom Cadangan Pasokan di Excel yang berbeda lebih dari 1 MW dari hasil DMN − Outage − Beban.` });
    else notes.push({ level: 'info', modul: 'Pasokan', teks: 'Daya Mampu Pasok dihitung sebagai DMN IBT + KIT dikurangi Outage & Derating. Hasil DMP − Beban sama dengan kolom Cadangan Pasokan di Excel untuk semua sub sistem.' });
    const cnt = {}; cur.forEach(s => { const k = s.dmn; cnt[k] = (cnt[k] || 0) + 1; });
    Object.entries(cnt).filter(([, n]) => n >= 3).forEach(([v, n]) => notes.push({ level: 'info', modul: 'Pasokan', teks: `${n} sub sistem punya DMN identik ${Number(v).toLocaleString('id-ID', { maximumFractionDigits: 2 })} MW. Wajar bila konfigurasi IBT-nya sama, tapi perlu dikonfirmasi.` }));
    return res;
  }

  /* ---------------- 4. Keandalan ---------------- */
  function readAF(wb, P) {
    const sh = findSheet(wb, n => n === 'tabel3', 'Tabel 3');
    const blocks = (kind) => {
      const res = {};
      findAll(sh, (v, r, c) => monthOf(v) >= 0 && norm(sh.text(r, c + 1)) === kind).forEach(h => {
        const m = monthOf(sh.get(h.r, h.c)); if (m > P.bulan || res[m]) return;
        const hi = findCell(sh, txtIs('Instalasi'), { r0: h.r, r1: h.r + 3, c0: h.c + 2, c1: h.c + 2 });
        if (!hi) return;
        const tot = findCell(sh, txtIs('Total'), { r0: hi.r + 1, r1: hi.r + 25, c0: h.c + 1, c1: h.c + 1 });
        if (!tot) return;
        res[m] = { h, hi, tot };
      });
      return res;
    };
    const P_ = blocks('penghantar'), T_ = blocks('trafo');
    for (let m = 0; m <= P.bulan; m++) {
      if (!P_[m]) fail(`Blok penghantar ${BULAN[m]} tidak ditemukan di Tabel 3.`);
      if (!T_[m]) fail(`Blok trafo ${BULAN[m]} tidak ditemukan di Tabel 3.`);
    }
    const n = (r, c) => sh.num(r, c) || 0;
    const bulanan = [];
    for (let m = 0; m <= P.bulan; m++) {
      const p = P_[m], t = T_[m], pc = p.h.c, tc = t.h.c;
      bulanan.push({ bulan: BLN[m], pht_ggn: n(p.tot.r, pc + 5), pht_ggn_jam: r2(n(p.tot.r, pc + 6)), pht_har: n(p.tot.r, pc + 7), pht_har_jam: r2(n(p.tot.r, pc + 8)),
        trf_ggn: n(t.tot.r, tc + 5), trf_ggn_jam: r2(n(t.tot.r, tc + 6)), trf_har: n(t.tot.r, tc + 8), trf_har_jam: r2(n(t.tot.r, tc + 9)) });
    }
    const detail = (b, cols) => { const out = []; for (let r = b.hi.r + 1; r < b.tot.r; r++) { const nm = sh.get(r, b.h.c + 2); if (typeof nm !== 'string' || /^(sirkit|unit|kms|mva)$/i.test(nm.trim())) continue; const o = { instalasi: nm.trim() }; for (const [k, off] of Object.entries(cols)) o[k] = r2(sh.num(r, b.h.c + off), 4); out.push(o); } return out; };
    return { bulanan,
      pht: detail(P_[P.bulan], { sirkit: 3, kms: 4, ggn_kali: 5, ggn_jam: 6, har_kali: 7, har_jam: 8, af: 9 }),
      trf: detail(T_[P.bulan], { unit: 3, mva: 4, ggn_kali: 5, ggn_jam: 6, har_kali: 8, har_jam: 9, af: 11 }) };
  }

  function readPenyebab(wb, P) {
    const sh = findSheet(wb, n => n === 'tabel4', 'Tabel 4');
    const hd = findCell(sh, (v, r, c) => monthOf(v) === P.bulan && typeof v === 'string' && v.trim().length > 3 && txtHas('kumulatif')(sh.get(r, c + 1)), { c0: 3, c1: 3 });
    if (!hd) fail(`Blok ${BULAN[P.bulan]} tidak ditemukan di Tabel 4.`);
    const out = [];
    for (let r = hd.r + 1; r <= sh.maxRow; r++) {
      if (norm(sh.text(r, 1)) === 'jumlah') break;
      const k = sh.get(r, 2); if (typeof k !== 'string') continue;
      out.push({ kategori: k.trim(), tp_bln: sh.num(r, 3) || 0, tp_kum: sh.num(r, 4) || 0, p_bln: sh.num(r, 6) || 0, p_kum: sh.num(r, 7) || 0 });
    }
    if (!out.length) fail('Kategori penyebab gangguan kosong di Tabel 4.');
    return out;
  }

  function parseWaktu(s) {
    if (s instanceof Date) return s;
    const m = String(s || '').match(/(\d{1,2})-([A-Za-z]{3})-(\d{2,4})\s*,?\s*(\d{1,2}):(\d{2})/);
    if (!m || MON_EN[m[2]] == null) return null;
    let y = +m[3]; if (y < 100) y += 2000;
    return new Date(Date.UTC(y, MON_EN[m[2]], +m[1], +m[4], +m[5]));
  }

  function readGangguan(wb, P, notes) {
    const sh = findSheet(wb, n => n.startsWith('tabel7.1'), 'Tabel 7.1 & 7.2');
    const st = findCell(sh, txtStarts('Tabel 7.1'), { c0: 1, c1: 1 });
    const en = findCell(sh, txtStarts('Tabel 7.2'), { c0: 1, c1: 1 });
    if (!st || !en) fail('Judul "Tabel 7.1" atau "Tabel 7.2" tidak ditemukan.');
    const hd = findCell(sh, txtIs('Keluar'), { r0: st.r, r1: st.r + 5 });
    if (!hd) fail('Judul kolom "Keluar" tidak ditemukan di Tabel 7.1.');
    const col = (k, rr = hd.r) => { const f = findCell(sh, txtHas(k), { r0: rr, r1: rr + 1 }); return f ? f.c : null; };
    const C = { keluar: hd.c, masuk: col('Masuk'), lokasi: col('Lokasi'), bay: col('Bay'), sebab: col('Penyebab'), kat: col('Kategori'), padam: col('Padam', hd.r + 1) || col('Padam'), ens: col('ENS'), srt: col('SRT') };
    // kolom Padam/ENS/SRT memakai satuan di baris di bawah judul; cari judul utamanya di baris judul
    C.padam = findCell(sh, txtIs('Padam'), { r0: hd.r, r1: hd.r, c0: C.kat + 1 })?.c ?? C.padam;
    let upt = '-'; const out = [];
    for (let r = hd.r + 1; r < en.r; r++) {
      const a = sh.get(r, 1);
      if (typeof a === 'string' && /^upt/i.test(a.trim())) { upt = a.trim().replace(/^upt\s*/i, ''); continue; }
      const kv = sh.get(r, C.keluar); const tk = parseWaktu(kv);
      if (!tk) continue;
      const mv = sh.get(r, C.masuk); const tm = parseWaktu(mv);
      const ev = { baris: r, upt: upt.charAt(0) + upt.slice(1).toLowerCase(), keluar: typeof kv === 'string' ? kv.trim() : kv.toISOString(), masuk: typeof mv === 'string' ? mv.trim() : (mv ? String(mv) : '-'),
        bulan: tk.getUTCMonth(), tahun: tk.getUTCFullYear(), lokasi: sh.text(r, C.lokasi), bay: sh.text(r, C.bay).replace(/\s+/g, ' '), penyebab: sh.text(r, C.sebab), kategori: sh.text(r, C.kat).toUpperCase(),
        padam: r2(sh.num(r, C.padam)), ens: r2(sh.num(r, C.ens)), srt: r2(sh.num(r, C.srt)) };
      out.push(ev);
      if (tm && tm < tk) notes.push({ level: 'warn', modul: 'Keandalan', teks: `Tabel 7.1 baris ${r}: waktu masuk (${ev.masuk}) lebih awal dari waktu keluar (${ev.keluar}).` });
      if (typeof mv === 'string' && /belum normal/i.test(mv)) notes.push({ level: 'info', modul: 'Keandalan', teks: `Tabel 7.1 baris ${r}: instalasi tercatat "Belum Normal" sejak ${ev.keluar}.` });
      if (ev.tahun !== P.tahun || ev.bulan > P.bulan) notes.push({ level: 'warn', modul: 'Keandalan', teks: `Tabel 7.1 baris ${r}: tanggal kejadian ${ev.keluar} berada di luar periode Januari–${BULAN[P.bulan]} ${P.tahun}.` });
    }
    return out;
  }

  function readSRT(wb, P, notes) {
    const sh = findSheet(wb, n => n === 'tabel8', 'Tabel 8');
    const out = [];
    for (let r = 1; r <= sh.maxRow; r++) {
      const m = monthOf(sh.get(r, 1)); if (m < 0 || m > P.bulan || out[m]) continue;
      const t = findCell(sh, txtIs('Total'), { r0: r + 1, r1: r + 6, c0: 1, c1: 1 });
      if (!t) continue;
      out[m] = { bulan: BLN[m], padam: r2(sh.num(t.r, 3) || 0), ens: r2(sh.num(t.r, 4) || 0), srt: r2(sh.num(t.r, 5) || 0), srt_kum: r2(sh.num(t.r, 9) || 0) };
    }
    for (let m = 0; m <= P.bulan; m++) if (!out[m]) fail(`Rekap SRT ${BULAN[m]} tidak ditemukan di Tabel 8.`);
    const vals = out.map(x => x.srt).filter(v => v > 0).sort((a, b) => a - b);
    const med = vals.length ? vals[Math.floor(vals.length / 2)] : 0;
    out.forEach(x => { if (med && x.srt > 3 * med) notes.push({ level: 'info', modul: 'Keandalan', teks: `SRT ${BULAN[BLN.indexOf(x.bulan)]} tercatat ${x.srt.toLocaleString('id-ID')} menit, lebih dari 3 kali median bulanan (${med.toLocaleString('id-ID')} menit). Mohon dipastikan.` }); });
    return out;
  }

  function readHTGP(wb, P, notes) {
    const sh = findSheet(wb, n => n === 'htgp', 'HTGP');
    const titles = findAll(sh, txtHas('hari tanpa gi padam'), { c0: 1, c1: 1 });
    const part = (t) => {
      const bh = findCell(sh, txtIs('BULAN'), { r0: t.r, r1: t.r + 4, c0: 2, c1: 2 });
      if (!bh) return null;
      const jc = findCell(sh, txtHas('jumlah hari tanpa'), { r0: t.r, r1: t.r, c0: 3 });
      const res = [];
      for (let r = bh.r + 1; r <= bh.r + 12; r++) {
        const m = monthOf(sh.get(r, 2)); if (m < 0) continue;
        const hari = []; for (let d = 1; d <= 31; d++) if (sh.get(r, 2 + d) != null) hari.push(d);
        res[m] = { jumlah: jc ? sh.num(r, jc.c) : null, hari };
      }
      return res;
    };
    const tr = titles.find(t => /transmisi/i.test(sh.text(t.r, 1))), ds = titles.find(t => /distribu/i.test(sh.text(t.r, 1)));
    if (!tr) fail('Tabel HTGP transmisi tidak ditemukan.');
    const T = part(tr), Dd = ds ? part(ds) : [];
    if (!T || !T[P.bulan]) fail(`Baris ${BULAN[P.bulan]} tidak ditemukan di tabel HTGP.`);
    const wrongY = titles.map(t => ({ r: t.r, y: (sh.text(t.r, 1).match(/(20\d{2})/) || [])[1] })).filter(x => x.y && +x.y !== P.tahun);
    if (wrongY.length) notes.push({ level: 'info', modul: 'Keandalan', teks: `Judul tabel di sheet HTGP (baris ${wrongY.map(x => x.r).join(', ')}) tertulis tahun ${[...new Set(wrongY.map(x => x.y))].join('/')}, sementara periode laporan ${P.tahun}.` });
    const out = [];
    for (let m = 0; m <= P.bulan; m++) out.push({ bulan: BLN[m], htgp_trans: T[m]?.jumlah ?? null, htgp_dist: Dd?.[m]?.jumlah ?? null, hari_trans: T[m]?.hari || [], hari_dist: Dd?.[m]?.hari || [] });
    return out;
  }

  /* ---------------- 5. Mutu tegangan ---------------- */
  function readTegangan(wb, P, notes) {
    const sh = findSheet(wb, n => n === 'tabel10', 'Tabel 10');
    const out = []; const futur = [];
    for (let r = 1; r <= sh.maxRow; r++) {
      const v = sh.get(r, 1); const m = monthOf(v); if (m < 0 || norm(sh.text(r, 2)) !== 'uraian') continue;
      const lv = {};
      for (const k of ['66', '70', '150', '500']) {
        const f = findCell(sh, txtIs(k + ' kV'), { r0: r, r1: r });
        if (!f) continue;
        const rmin = findCell(sh, txtStarts('Teg. Min'), { r0: r + 1, r1: r + 4, c0: 2, c1: 2 }), rmax = findCell(sh, txtStarts('Teg. Max'), { r0: r + 1, r1: r + 5, c0: 2, c1: 2 });
        if (!rmin || !rmax) continue;
        lv[k] = { min_r: r2(sh.num(rmin.r, f.c)), min_a: r2(sh.num(rmin.r, f.c + 1)), max_r: r2(sh.num(rmax.r, f.c)), max_a: r2(sh.num(rmax.r, f.c + 1)) };
      }
      if (!Object.keys(lv).length) continue;
      if (m > P.bulan) { if (Object.values(lv).some(x => x.min_a != null || x.max_a != null)) futur.push(BULAN[m]); continue; }
      out[m] = { bulan: BULAN[m], ...lv };
    }
    for (let m = 0; m <= P.bulan; m++) if (!out[m]) fail(`Data tegangan ${BULAN[m]} tidak ditemukan di Tabel 10.`);
    if (futur.length) notes.push({ level: 'warn', modul: 'Mutu Tegangan', teks: `Tabel 10 berisi nilai realisasi tegangan untuk ${futur.join(', ')}, padahal periode laporan baru sampai ${BULAN[P.bulan]} ${P.tahun}. Dashboard hanya memakai Januari–${BULAN[P.bulan]}.` });
    return out;
  }

  function readEkskursi(wb, P, notes) {
    const sh = findSheet(wb, n => n === 'tabel9', 'Tabel 9');
    const re = findCell(sh, txtIs('Ekskursi GITET-GI')); const rt = findCell(sh, txtIs('Total GITET-GI'));
    if (!re || !rt) fail('Baris "Ekskursi GITET-GI" tidak ditemukan di Tabel 9.');
    const hf = findCell(sh, txtStarts('Jumlah'), { r0: 1, r1: re.r, c0: re.c + 1, c1: re.c + 8 }); const hdrRow = hf ? hf.r : 1;
    const out = []; const seen = new Set();
    for (let c = re.c + 1; c <= re.c + 8; c++) {
      const h = sh.text(hdrRow, c); const mm = h.match(/jumlah\s+([A-Za-z]+)/i); const m = mm ? monthOf(mm[1]) : -1;
      if (m < 0 || m > P.bulan || seen.has(m)) continue; seen.add(m);
      out.push({ bulan: BLN[m], total: sh.num(rt.r, c), ekskursi: sh.num(re.r, c) });
    }
    if (!seen.has(P.bulan)) notes.push({ level: 'info', modul: 'Mutu Tegangan', teks: `Tabel 9 Ekskursi Tegangan tidak punya kolom ${BULAN[P.bulan]} (hanya ada ${out.length ? out.map(x => x.bulan).join(', ') : 'kolom tanpa nama bulan'}). Angka ekskursi bulan ini diambil dari Rangkuman.` });
    return out;
  }

  /* ---------------- 6. Pembebanan ---------------- */
  function readPembebanan(wb, P) {
    const sh = findSheet(wb, n => n.startsWith('gambar2.1'), 'Gambar 2.1 & 2.2');
    const heads = findAll(sh, (v, r, c) => typeof v === 'string' && /^(penghantar|trafo)$/i.test(v.trim()) && monthOf(sh.get(r, c + 1)) === 0);
    const read = (kind) => {
      const h = heads.find(x => norm(sh.text(x.r, x.c)) === kind); if (!h) fail(`Tabel bulanan ${kind} (Gambar 2.x) tidak ditemukan.`);
      const res = {};
      for (let r = h.r + 1; r <= sh.maxRow; r++) { const k = sh.get(r, h.c); if (k == null || norm(k) === 'kategori') break; const arr = []; for (let m = 0; m <= P.bulan; m++) arr.push(sh.num(r, h.c + 1 + m)); res[String(k).trim()] = arr; }
      return res;
    };
    const band = (test, label) => { const s = findSheet(wb, test, label); const hd = findCell(s, txtIs('Label')); if (!hd) fail(`Judul "Label" tidak ditemukan di ${label}.`); const out = []; for (let r = hd.r + 1; r <= s.maxRow; r++) { const nm = s.get(r, hd.c); if (nm == null) continue; out.push({ nama: String(nm).replace(/\s+/g, ' ').trim(), b0: r2(s.num(r, hd.c + 1)), b40: r2(s.num(r, hd.c + 2)), b60: r2(s.num(r, hd.c + 3)), b80: r2(s.num(r, hd.c + 4)) }); } return out; };
    return { pht60: read('penghantar'), trf80: read('trafo'), band_ibt: band(n => n === 'datapythonpembebananibt', 'Data Python Pembebanan IBT'), band_trf: band(n => n === 'datapythonpembebanantrafo', 'Data Python Pembebanan Trafo') };
  }

  /* ---------------- 7. ROB ---------------- */
  function readROB(wb, P, notes) {
    const sh = findSheet(wb, n => n.includes('pekerjaanrob'), 'Evaluasi Pekerjaan ROB');
    const hd = findCell(sh, v => typeof v === 'string' && /^no\.?$/i.test(v.trim()), { c0: 1, c1: 1 });
    if (!hd) fail('Judul kolom "NO." tidak ditemukan di Evaluasi Pekerjaan ROB.');
    const col = k => findCell(sh, txtStarts(k), { r0: hd.r, r1: hd.r })?.c;
    const cLok = col('LOKASI'), cTeg = col('TEG'), cRen = col('RENC'), cSt = col('Terlaksana');
    const sum = findCell(sh, txtIs('RENCANA'));
    if (!cLok || !cRen || !cSt || !sum) fail('Kolom LOKASI/RENC/Terlaksana atau tabel ringkasan "RENCANA" tidak ditemukan di Evaluasi Pekerjaan ROB.');
    const per = {}, teg = {}, bad = [];
    for (let r = hd.r + 2; r < sum.r; r++) {
      if (sh.get(r, cLok) == null) continue;
      const jv = sh.get(r, cRen); let jn = typeof jv === 'string' ? jv.trim() : '';
      if (!/^[A-Za-z][A-Za-z .]*$/.test(jn)) { bad.push(r); jn = 'Jenis tidak valid'; }
      const st0 = sh.text(r, cSt); const st = /^terlaksana$/i.test(st0) ? 'Terlaksana' : /^abk$/i.test(st0) ? 'ABK' : 'Belum diisi';
      per[jn] = per[jn] || { jenis: jn, jumlah: 0, terlaksana: 0, abk: 0, kosong: 0 };
      per[jn].jumlah++; per[jn][st === 'Terlaksana' ? 'terlaksana' : st === 'ABK' ? 'abk' : 'kosong']++;
      const t = (sh.text(r, cTeg) || '-').replace(/\s+/g, ''); teg[t] = teg[t] || {}; teg[t][st] = (teg[t][st] || 0) + 1;
    }
    const rows = [];
    for (let r = sum.r + 1; r <= sum.r + 12; r++) { const j = sh.text(r, sum.c); if (!j) break; rows.push({ jenis: j, jumlah: sh.num(r, sum.c + 1) || 0, terlaksana: sh.num(r, sum.c + 2) || 0, abk: sh.num(r, sum.c + 3) || 0, kosong: sh.num(r, sum.c + 4) || 0 }); if (/^total/i.test(j)) break; }
    const tot = rows.find(x => /^total/i.test(x.jenis));
    const cnt = Object.values(per).reduce((a, x) => ({ jumlah: a.jumlah + x.jumlah, terlaksana: a.terlaksana + x.terlaksana, abk: a.abk + x.abk, kosong: a.kosong + x.kosong }), { jumlah: 0, terlaksana: 0, abk: 0, kosong: 0 });
    if (tot && (tot.jumlah !== cnt.jumlah || tot.terlaksana !== cnt.terlaksana)) notes.push({ level: 'warn', modul: 'ROB', teks: `Ringkasan ROB di Excel (${tot.jumlah} pekerjaan, ${tot.terlaksana} terlaksana) berbeda dengan hitungan dari daftar pekerjaan (${cnt.jumlah} pekerjaan, ${cnt.terlaksana} terlaksana).` });
    if (bad.length) notes.push({ level: 'warn', modul: 'ROB', teks: `Evaluasi Pekerjaan ROB baris ${bad.join(', ')}: kolom jenis rencana (RENC) tidak berisi kode ROB/ROT/ROM/ROH/FORM A, sehingga pekerjaan ini tidak ikut terhitung di ringkasan Excel.` });
    if (cnt.kosong > 0) notes.push({ level: 'warn', modul: 'ROB', teks: `${cnt.kosong} dari ${cnt.jumlah} pekerjaan ROB belum diisi statusnya (Terlaksana/ABK).` });
    const list = Object.values(per).sort((a, b) => b.jumlah - a.jumlah);
    return { per_jenis: list, total: { jenis: 'Total Pekerjaan', ...cnt }, per_teg: teg };
  }

  function checkGangguanMeluas(wb, P, notes) {
    try {
      const sh = findSheet(wb, n => n.startsWith('tabel6.1') && n.includes('baru'), 'Tabel 6.1 (Baru)');
      const st = findCell(sh, txtStarts('Tabel 6.1'), { c0: 1, c1: 1 }), en = findCell(sh, txtStarts('Tabel 6.2'), { c0: 1, c1: 1 });
      if (!st) return;
      const yrs = {};
      for (let r = st.r + 1; r < (en ? en.r : sh.maxRow); r++) { const t = parseWaktu(sh.get(r, 2)); if (t && t.getUTCFullYear() !== P.tahun) yrs[t.getUTCFullYear()] = (yrs[t.getUTCFullYear()] || 0) + 1; }
      Object.entries(yrs).forEach(([y, n]) => notes.push({ level: 'warn', modul: 'Keandalan', teks: `Tabel 6.1 Gangguan Meluas berisi ${n} kejadian bertahun ${y}, sementara periode laporan ${P.tahun}.` }));
    } catch (e) { /* tabel 6 tidak dipakai di dashboard */ }
  }

  /* ---------------- Utama ---------------- */
  function readEOB(wb, fileName) {
    const notes = [], status = {};
    const P = readPeriode(wb);
    const D = { meta: { periode: `${BULAN[P.bulan]} ${P.tahun}`, bulan: P.bulan, tahun: P.tahun, sumber: fileName || 'File EOB', ambang_cadangan: 400, samaran: false } };
    const run = (key, label, fn) => { try { const v = fn(); status[key] = { ok: true }; return v; } catch (e) { status[key] = { ok: false, label, alasan: e.userFacing ? e.message : 'Terjadi kesalahan saat membaca: ' + e.message }; return null; } };
    D.ringkasan = run('ringkasan', 'Ringkasan', () => readRingkasan(wb));
    D.tren_bp = run('beban', 'Tren Beban Puncak', () => readTrenBP(wb, P));
    D.subsistem = run('pasokan', 'Pasokan Sub Sistem', () => readSubsistem(wb, P, notes));
    const k = run('keandalan', 'Keandalan', () => ({ af: readAF(wb, P), penyebab: readPenyebab(wb, P), gangguan: readGangguan(wb, P, notes), srt: readSRT(wb, P, notes), htgp: readHTGP(wb, P, notes) }));
    if (k) Object.assign(D, { af_bulanan: k.af.bulanan, af_pht: k.af.pht, af_trf: k.af.trf, penyebab: k.penyebab, gangguan: k.gangguan, srt: k.srt, htgp: k.htgp });
    D.tegangan = run('tegangan', 'Mutu Tegangan', () => readTegangan(wb, P, notes));
    D.ekskursi = (() => { try { return readEkskursi(wb, P, notes); } catch (e) { return []; } })();
    const pb = run('pembebanan', 'Pembebanan', () => readPembebanan(wb, P));
    if (pb) Object.assign(D, pb);
    D.rob = run('rob', 'Pekerjaan ROB', () => readROB(wb, P, notes));
    checkGangguanMeluas(wb, P, notes);
    // Pemeriksaan silang antarmodul
    if (D.ringkasan && D.tren_bp) {
      const t = D.tren_bp[P.bulan]; const mx = Math.max(t.malam_cur || 0, t.siang_cur || 0);
      if (Math.abs(mx - D.ringkasan.bp) > 1) notes.push({ level: 'warn', modul: 'Ringkasan', teks: `Beban puncak di Rangkuman (${D.ringkasan.bp} MW) berbeda dengan Tabel 1 untuk ${BULAN[P.bulan]} (${mx} MW).` });
    }
    if (D.ringkasan && D.srt) {
      const s = D.srt[P.bulan];
      if (D.ringkasan.ens === 0 && s.ens > 0) notes.push({ level: 'info', modul: 'Ringkasan', teks: `ENS 0 MWh di Rangkuman adalah ENS pelepasan beban (UFLS, OLS, MLS). ENS akibat gangguan di Tabel 8 untuk ${BULAN[P.bulan]} adalah ${s.ens.toLocaleString('id-ID')} MWh (padam ${s.padam.toLocaleString('id-ID')} MW).` });
    }
    D.catatan = notes;
    D.status = status;
    return D;
  }

  global.EOBReader = { readEOB, BULAN, BLN };
})(typeof window !== 'undefined' ? window : globalThis);

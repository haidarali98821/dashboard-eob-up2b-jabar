/* xlsx-lite: pembaca file .xlsx minimal yang berjalan di browser.
   Hanya membaca nilai sel (termasuk hasil rumus yang tersimpan), tanpa mengirim file ke mana pun.
   Memakai DecompressionStream bawaan browser untuk membuka arsip ZIP. */
(function (global) {
  'use strict';

  async function inflateRaw(bytes) {
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  // Membaca daftar file di ZIP lewat central directory.
  function zipEntries(buf) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let eocd = -1;
    for (let i = buf.length - 22; i >= Math.max(0, buf.length - 66000); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('File bukan arsip .xlsx yang valid.');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const out = {};
    const td = new TextDecoder();
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Struktur ZIP rusak.');
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = td.decode(buf.subarray(p + 46, p + 46 + nlen));
      const lnlen = dv.getUint16(lho + 26, true), lelen = dv.getUint16(lho + 28, true);
      const start = lho + 30 + lnlen + lelen;
      out[name] = { method, data: buf.subarray(start, start + csize) };
      p += 46 + nlen + elen + clen;
    }
    return out;
  }

  async function readEntry(e) {
    if (!e) return null;
    const raw = e.method === 0 ? e.data : await inflateRaw(e.data);
    return new TextDecoder().decode(raw);
  }

  const xmlUnescape = s => s.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (m, g) => {
    if (g === 'lt') return '<'; if (g === 'gt') return '>'; if (g === 'amp') return '&';
    if (g === 'quot') return '"'; if (g === 'apos') return "'";
    if (g[1] === 'x') return String.fromCodePoint(parseInt(g.slice(2), 16));
    return String.fromCodePoint(parseInt(g.slice(1), 10));
  });
  const attr = (tag, name) => { const m = tag.match(new RegExp('\\s' + name + '="([^"]*)"')); return m ? xmlUnescape(m[1]) : null; };
  const textOf = frag => { let s = ''; frag.replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, (m, t) => { s += t; return m; }); return xmlUnescape(s); };

  function colNum(letters) { let n = 0; for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }

  const DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57]);

  function parseStyles(xml) {
    const custom = {};
    if (!xml) return [];
    xml.replace(/<numFmt\s[^>]*>/g, t => { custom[+attr(t, 'numFmtId')] = attr(t, 'formatCode') || ''; return t; });
    const cx = xml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
    const res = [];
    if (cx) cx[1].replace(/<xf\s[^>]*?\/?>/g, t => {
      const id = +attr(t, 'numFmtId');
      let isDate = DATE_IDS.has(id);
      if (!isDate && custom[id] != null) {
        const f = custom[id].replace(/\[[^\]]*\]|"[^"]*"|\\./g, '');
        isDate = /[dmyhs]/i.test(f) && !/^[#0.,%\s]*$/.test(f);
      }
      res.push(isDate);
      return t;
    });
    return res;
  }

  class Sheet {
    constructor(name) { this.name = name; this.cells = new Map(); this.maxRow = 0; this.maxCol = 0; }
    get(r, c) { const v = this.cells.get(r * 20000 + c); return v === undefined ? null : v; }
    set(r, c, v) { this.cells.set(r * 20000 + c, v); if (r > this.maxRow) this.maxRow = r; if (c > this.maxCol) this.maxCol = c; }
    text(r, c) { const v = this.get(r, c); return v == null ? '' : (v instanceof Date ? v.toISOString() : String(v)).trim(); }
    num(r, c) { const v = this.get(r, c); return typeof v === 'number' && isFinite(v) ? v : null; }
  }

  // Nilai tanggal Excel → Date (UTC). Nilai jam saja (< 1) dikembalikan sebagai {time:"HH:MM"}.
  function excelDate(serial) {
    if (serial < 1) {
      const mins = Math.round(serial * 1440);
      return { time: String(Math.floor(mins / 60) % 24).padStart(2, '0') + ':' + String(mins % 60).padStart(2, '0') };
    }
    return new Date(Math.round((serial - 25569) * 86400000));
  }

  async function readWorkbook(arrayBuffer) {
    const buf = new Uint8Array(arrayBuffer);
    const files = zipEntries(buf);
    const wbXml = await readEntry(files['xl/workbook.xml']);
    if (!wbXml) throw new Error('File tidak berisi workbook Excel.');
    const relXml = await readEntry(files['xl/_rels/workbook.xml.rels']) || '';
    const rels = {};
    relXml.replace(/<Relationship\s[^>]*>/g, t => { rels[attr(t, 'Id')] = attr(t, 'Target'); return t; });
    const ssXml = await readEntry(files['xl/sharedStrings.xml']);
    const shared = [];
    if (ssXml) ssXml.replace(/<si>([\s\S]*?)<\/si>/g, (m, inner) => { shared.push(textOf(inner)); return m; });
    const dateStyles = parseStyles(await readEntry(files['xl/styles.xml']));

    const sheets = [];
    const sheetTags = wbXml.match(/<sheet\s[^>]*>/g) || [];
    for (const t of sheetTags) {
      const name = attr(t, 'name');
      const rid = attr(t, 'r:id');
      let target = rels[rid] || '';
      target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
      const path = 'xl/' + target;
      const sh = new Sheet(name);
      sh.state = attr(t, 'state') || 'visible';
      const xml = await readEntry(files[path]);
      if (xml) {
        const re = /<c\s([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
        let m;
        while ((m = re.exec(xml))) {
          const a = ' ' + m[1];
          const ref = attr(a, 'r'); if (!ref) continue;
          const mm = ref.match(/^([A-Z]+)(\d+)$/); if (!mm) continue;
          const r = +mm[2], c = colNum(mm[1]);
          const type = attr(a, 't'), s = attr(a, 's');
          const inner = m[3] || '';
          let val = null;
          if (type === 'inlineStr') val = textOf(inner);
          else {
            const vm = inner.match(/<v>([\s\S]*?)<\/v>/);
            if (!vm) continue;
            const raw = xmlUnescape(vm[1]);
            if (type === 's') val = shared[+raw] ?? '';
            else if (type === 'str') val = raw;
            else if (type === 'b') val = raw === '1';
            else if (type === 'e') val = { error: raw };
            else {
              const n = parseFloat(raw);
              val = (s != null && dateStyles[+s] && isFinite(n)) ? excelDate(n) : n;
            }
          }
          if (val === '' || val == null) continue;
          sh.set(r, c, val);
        }
      }
      sheets.push(sh);
    }
    return { sheets, sheet: name => sheets.find(s => s.name === name) };
  }

  global.XlsxLite = { readWorkbook };
})(typeof window !== 'undefined' ? window : globalThis);

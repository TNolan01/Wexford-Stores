/* Minimal spreadsheet reader: .xlsx/.xlsm (zip + XML) and .csv/.tsv/.txt. No dependencies, works offline. */
(function () {
  'use strict';

  async function inflateRaw(bytes) {
    var ds = new DecompressionStream('deflate-raw');
    var stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function unzip(buf) {
    var u8 = new Uint8Array(buf), dv = new DataView(buf), e = -1;
    for (var i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { e = i; break; }
    }
    if (e < 0) throw new Error('notzip');
    var count = dv.getUint16(e + 10, true), p = dv.getUint32(e + 16, true), entries = {};
    for (var n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      var method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true),
          nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true),
          clen = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true),
          name = new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nlen));
      entries[name] = { method: method, csize: csize, off: off };
      p += 46 + nlen + elen + clen;
    }
    return { entries: entries, u8: u8, dv: dv };
  }

  async function readEntry(z, name) {
    var en = z.entries[name];
    if (!en) return null;
    var nlen = z.dv.getUint16(en.off + 26, true), elen = z.dv.getUint16(en.off + 28, true);
    var start = en.off + 30 + nlen + elen;
    var data = z.u8.subarray(start, start + en.csize);
    var out = en.method === 0 ? data : await inflateRaw(data);
    return new TextDecoder('utf-8').decode(out);
  }

  function xml(text) {
    var doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('badxml');
    return doc;
  }
  function els(node, tag) { return Array.prototype.slice.call(node.getElementsByTagName(tag)); }
  var REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  function colIndex(ref) {
    var m = /^([A-Z]+)/i.exec(ref || '');
    if (!m) return -1;
    var s = m[1].toUpperCase(), n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }

  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function serialToText(serial, timeOnly) {
    if (timeOnly || serial < 1) {
      var secs = Math.round((serial % 1) * 86400);
      return pad(Math.floor(secs / 3600) % 24) + ':' + pad(Math.floor(secs / 60) % 60);
    }
    var d = new Date(Math.round((serial - 25569) * 86400000));
    var out = pad(d.getUTCDate()) + '/' + pad(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
    if (serial % 1) out += ' ' + pad(d.getUTCHours()) + ':' + pad(d.getUTCMinutes());
    return out;
  }

  function dateStyles(stylesText) {
    var flags = [];
    if (!stylesText) return flags;
    var doc = xml(stylesText), custom = {};
    els(doc, 'numFmt').forEach(function (n) { custom[n.getAttribute('numFmtId')] = n.getAttribute('formatCode') || ''; });
    var xfs = doc.getElementsByTagName('cellXfs')[0];
    if (!xfs) return flags;
    els(xfs, 'xf').forEach(function (xf, i) {
      var id = parseInt(xf.getAttribute('numFmtId'), 10), kind = null;
      if ((id >= 14 && id <= 17) || id === 22) kind = 'date';
      else if (id >= 18 && id <= 21) kind = 'time';
      else if (id >= 45 && id <= 47) kind = 'time';
      else if (custom[id] !== undefined) {
        var code = custom[id].replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
        if (/[dy]/i.test(code)) kind = 'date';
        else if (/h|s/i.test(code) && !/[#0]/.test(code)) kind = 'time';
      }
      flags[i] = kind;
    });
    return flags;
  }

  function cleanNumber(v) {
    var n = Number(v);
    return isFinite(n) ? String(parseFloat(n.toPrecision(15))) : v;
  }

  async function readXlsx(buf) {
    var z = unzip(buf);
    var wbText = await readEntry(z, 'xl/workbook.xml');
    if (!wbText) throw new Error('notxlsx');
    var relText = await readEntry(z, 'xl/_rels/workbook.xml.rels');
    var rels = {};
    if (relText) els(xml(relText), 'Relationship').forEach(function (r) { rels[r.getAttribute('Id')] = r.getAttribute('Target'); });

    var shared = [];
    var ssText = await readEntry(z, 'xl/sharedStrings.xml');
    if (ssText) {
      els(xml(ssText), 'si').forEach(function (si) {
        var s = '';
        Array.prototype.forEach.call(si.childNodes, function (ch) {
          if (ch.nodeType !== 1) return;
          if (ch.localName === 't') s += ch.textContent;
          else if (ch.localName === 'r') els(ch, 't').forEach(function (t) { s += t.textContent; });
        });
        shared.push(s);
      });
    }
    var styleFlags = dateStyles(await readEntry(z, 'xl/styles.xml'));

    var result = {};
    var sheetNodes = els(xml(wbText), 'sheet');
    for (var k = 0; k < sheetNodes.length; k++) {
      var sn = sheetNodes[k];
      if (sn.getAttribute('state') === 'hidden' || sn.getAttribute('state') === 'veryHidden') continue;
      var rid = sn.getAttribute('r:id') || sn.getAttributeNS(REL_NS, 'id');
      var target = rels[rid];
      if (!target) continue;
      var path = target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      var text = await readEntry(z, path);
      if (!text) continue;
      var doc = xml(text), aoa = [], rowSeq = 0;
      els(doc, 'row').forEach(function (row) {
        var ri = row.getAttribute('r') ? parseInt(row.getAttribute('r'), 10) - 1 : rowSeq;
        rowSeq = ri + 1;
        var arr = [], seq = 0;
        els(row, 'c').forEach(function (c) {
          var ci = c.getAttribute('r') ? colIndex(c.getAttribute('r')) : seq;
          seq = ci + 1;
          var t = c.getAttribute('t'), vEl = c.getElementsByTagName('v')[0], val = '';
          if (t === 'inlineStr') {
            var is = c.getElementsByTagName('is')[0];
            val = is ? els(is, 't').map(function (x) { return x.textContent; }).join('') : '';
          } else if (vEl) {
            var raw = vEl.textContent;
            if (t === 's') val = shared[parseInt(raw, 10)] || '';
            else if (t === 'b') val = raw === '1' ? 'TRUE' : 'FALSE';
            else if (t === 'str' || t === 'e') val = raw;
            else {
              var kind = styleFlags[parseInt(c.getAttribute('s') || '0', 10)];
              var num = Number(raw);
              if (kind && isFinite(num)) val = serialToText(num, kind === 'time');
              else val = cleanNumber(raw);
            }
          }
          arr[ci] = val;
        });
        for (var i = 0; i < arr.length; i++) if (arr[i] === undefined) arr[i] = '';
        aoa[ri] = arr;
      });
      var compact = [];
      aoa.forEach(function (r) { if (r) compact.push(r); });
      result[sn.getAttribute('name')] = compact;
    }
    return result;
  }

  function parseCsv(text) {
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    var first = text.split(/\r?\n/, 1)[0] || '';
    var counts = { ',': 0, ';': 0, '\t': 0 };
    first.replace(/"[^"]*"/g, '').split('').forEach(function (ch) { if (counts[ch] !== undefined) counts[ch]++; });
    var delim = ',';
    if (counts[';'] > counts[delim]) delim = ';';
    if (counts['\t'] > counts[delim]) delim = '\t';
    var rows = [], row = [], cell = '', q = false;
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === delim) { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }

  window.readSpreadsheet = async function (file) {
    var name = (file.name || '').toLowerCase();
    var buf = await file.arrayBuffer();
    var head = new Uint8Array(buf.slice(0, 4));
    if (head[0] === 0xD0 && head[1] === 0xCF && head[2] === 0x11 && head[3] === 0xE0) throw new Error('xls');
    if (head[0] === 0x50 && head[1] === 0x4B) return readXlsx(buf);
    if (/\.(csv|tsv|txt)$/.test(name)) {
      var out = {};
      out[file.name.replace(/\.[^.]+$/, '') || 'Sheet1'] = parseCsv(new TextDecoder('utf-8').decode(buf));
      return out;
    }
    throw new Error('unsupported');
  };
})();

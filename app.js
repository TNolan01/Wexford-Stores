(function () {
  'use strict';
  var PAGE = 50;
  var $ = function (id) { return document.getElementById(id); };
  var state = { name: '', sample: true, sheets: {}, sheet: '', col: '-1', shown: PAGE };

  var SAMPLE = {
    'Inventory': {
      headers: ['Item', 'SKU', 'Location', 'Qty', 'Supplier'],
      rows: [
        ['M6 hex bolt, zinc', 'BLT-0601', 'Shelf A2', '420', 'Fixwell Ltd'],
        ['M8 hex bolt, zinc', 'BLT-0801', 'Shelf A2', '310', 'Fixwell Ltd'],
        ['Cable tie 200mm black', 'CBL-0200', 'Drawer C1', '1500', 'Strapco'],
        ['Cable tie 300mm white', 'CBL-0301', 'Drawer C1', '640', 'Strapco'],
        ['USB-C cable 1m', 'USB-0100', 'Shelf D4', '58', 'Linkpoint'],
        ['USB-C cable 2m', 'USB-0200', 'Shelf D4', '34', 'Linkpoint'],
        ['LED bulb E27 9W', 'LGT-0927', 'Shelf B1', '96', 'Brightside'],
        ['LED bulb B22 9W', 'LGT-0922', 'Shelf B1', '44', 'Brightside'],
        ['Duct tape silver 50m', 'TPE-0050', 'Shelf E3', '27', 'Strapco']
      ]
    }
  };

  /* ---------- storage (IndexedDB, so large files persist) ---------- */
  function idb() {
    return new Promise(function (res, rej) {
      var r = indexedDB.open('wexford-stores', 1);
      r.onupgradeneeded = function () { r.result.createObjectStore('kv'); };
      r.onsuccess = function () { res(r.result); };
      r.onerror = function () { rej(r.error); };
    });
  }
  function dbGet() {
    return idb().then(function (db) {
      return new Promise(function (res, rej) {
        var q = db.transaction('kv').objectStore('kv').get('data');
        q.onsuccess = function () { res(q.result); };
        q.onerror = function () { rej(q.error); };
      });
    });
  }
  function dbSet(val) {
    return idb().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction('kv', 'readwrite');
        tx.objectStore('kv').put(val, 'data');
        tx.oncomplete = res; tx.onerror = function () { rej(tx.error); };
      });
    });
  }

  /* ---------- helpers ---------- */
  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function reEsc(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  function str(v) { return v === null || v === undefined ? '' : String(v); }

  function buildSheet(aoa) {
    var rows = aoa.filter(function (r) { return r.some(function (c) { return str(c).trim() !== ''; }); });
    if (!rows.length) return { headers: [], rows: [] };
    var width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    var headers = [];
    for (var i = 0; i < width; i++) headers.push(str(rows[0][i]).trim() || ('Column ' + (i + 1)));
    var body = rows.slice(1).map(function (r) {
      var out = [];
      for (var i = 0; i < width; i++) out.push(str(r[i]));
      return out;
    });
    return { headers: headers, rows: body };
  }

  function prep(sheet) {
    sheet.hay = sheet.rows.map(function (r) { return r.join(' \u0001 ').toLowerCase(); });
    sheet.colHay = sheet.headers.map(function (_, i) { return sheet.rows.map(function (r) { return r[i].toLowerCase(); }); });
  }

  function loadData(name, sheets, sample) {
    Object.keys(sheets).forEach(function (k) { prep(sheets[k]); });
    state.name = name; state.sheets = sheets; state.sample = !!sample;
    state.sheet = Object.keys(sheets)[0] || '';
    state.col = '-1'; state.shown = PAGE;
    fillSheetSelect(); fillColSelect(); render();
  }

  function fillSheetSelect() {
    var sel = $('sheet'); sel.innerHTML = '';
    Object.keys(state.sheets).forEach(function (k) {
      var o = document.createElement('option'); o.value = k; o.textContent = k; sel.appendChild(o);
    });
    sel.value = state.sheet;
    sel.hidden = Object.keys(state.sheets).length < 2;
  }
  function fillColSelect() {
    var sel = $('col'); sel.innerHTML = '';
    var all = document.createElement('option'); all.value = '-1'; all.textContent = 'All columns'; sel.appendChild(all);
    var s = state.sheets[state.sheet];
    if (s) s.headers.forEach(function (h, i) {
      var o = document.createElement('option'); o.value = String(i); o.textContent = h; sel.appendChild(o);
    });
    sel.value = state.col;
  }

  function highlight(text, terms) {
    if (!terms.length) return esc(text);
    var re = new RegExp('(' + terms.map(reEsc).join('|') + ')', 'ig');
    return text.split(re).map(function (part, i) {
      return i % 2 ? '<mark>' + esc(part) + '</mark>' : esc(part);
    }).join('');
  }

  function showNotice(html) { $('notice').hidden = false; $('notice').innerHTML = html; }

  function render() {
    var s = state.sheets[state.sheet];
    var q = $('q').value.trim().toLowerCase();
    var terms = q ? q.split(/\s+/) : [];
    $('clear').hidden = !q;
    $('fname').textContent = state.name;
    if (state.sample) showNotice('<strong>Example data.</strong> Tap Open file and choose your own .xlsx or .csv. The file is read on this phone and is not uploaded anywhere.');
    else if (!$('notice').dataset.error) $('notice').hidden = true;
    var list = $('results'); list.innerHTML = '';
    if (!s || !s.rows.length) {
      $('count').textContent = '0 rows';
      list.innerHTML = '<li class="empty">This sheet has no data rows.</li>';
      $('more').hidden = true; return;
    }
    var ci = parseInt(state.col, 10), matches = [];
    for (var i = 0; i < s.rows.length; i++) {
      var hay = ci >= 0 ? s.colHay[ci][i] : s.hay[i], ok = true;
      for (var t = 0; t < terms.length; t++) { if (hay.indexOf(terms[t]) === -1) { ok = false; break; } }
      if (ok) matches.push(i);
    }
    $('count').textContent = (terms.length ? matches.length + ' of ' + s.rows.length : s.rows.length) + ' rows';
    if (!matches.length) {
      list.innerHTML = '<li class="empty">No rows match. Try fewer words or switch to All columns.</li>';
      $('more').hidden = true; return;
    }
    var slice = matches.slice(0, state.shown);
    list.innerHTML = slice.map(function (ri) {
      var r = s.rows[ri], dl = '';
      for (var c = 1; c < s.headers.length; c++) {
        if (r[c] === '') continue;
        dl += '<dt>' + esc(s.headers[c]) + '</dt><dd>' + highlight(r[c], terms) + '</dd>';
      }
      return '<li class="card"><h2>' + highlight(r[0] || '(blank)', terms) + '</h2><dl>' + dl + '</dl></li>';
    }).join('');
    var left = matches.length - slice.length;
    $('more').hidden = left <= 0;
    $('more').textContent = 'Show ' + Math.min(PAGE, left) + ' more (' + left + ' left)';
  }

  function handleFile(file) {
    if (!file) return;
    window.readSpreadsheet(file).then(function (raw) {
      var sheets = {};
      Object.keys(raw).forEach(function (n) { sheets[n] = buildSheet(raw[n]); });
      if (!Object.keys(sheets).length) throw new Error('empty');
      $('q').value = ''; $('notice').dataset.error = '';
      loadData(file.name, sheets, false);
      var plain = {};
      Object.keys(sheets).forEach(function (k) { plain[k] = { headers: sheets[k].headers, rows: sheets[k].rows }; });
      dbSet({ name: file.name, sheets: plain }).catch(function () { /* storage blocked: reopen the file next time */ });
    }).catch(function (e) {
      $('notice').dataset.error = '1';
      var msg = e && e.message === 'xls'
        ? 'That is an old .xls file. In Excel choose Save As and pick .xlsx or .csv, then open it again.'
        : 'Check that it is an .xlsx or .csv file and is not password protected.';
      showNotice('<strong>Could not read that file.</strong> ' + msg);
    });
  }

  /* ---------- events ---------- */
  $('file').addEventListener('change', function (e) { handleFile(e.target.files[0]); e.target.value = ''; });
  $('pick').addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file').click(); } });
  $('q').addEventListener('input', function () { state.shown = PAGE; render(); });
  $('clear').addEventListener('click', function () { $('q').value = ''; state.shown = PAGE; render(); $('q').focus(); });
  $('sheet').addEventListener('change', function (e) { state.sheet = e.target.value; state.col = '-1'; state.shown = PAGE; fillColSelect(); render(); });
  $('col').addEventListener('change', function (e) { state.col = e.target.value; state.shown = PAGE; render(); });
  $('more').addEventListener('click', function () { state.shown += PAGE; render(); });

  /* ---------- install button ---------- */
  var deferred = null;
  window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); deferred = e; $('install').hidden = false; });
  $('install').addEventListener('click', function () {
    if (!deferred) return;
    deferred.prompt();
    deferred.userChoice.then(function () { deferred = null; $('install').hidden = true; });
  });
  window.addEventListener('appinstalled', function () { $('install').hidden = true; });

  /* ---------- service worker ---------- */
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () { navigator.serviceWorker.register('sw.js').catch(function () {}); });
  }

  /* ---------- start ---------- */
  loadData('example.xlsx', SAMPLE, true);
  dbGet().then(function (d) {
    if (d && d.sheets) loadData(d.name, d.sheets, false);
  }).catch(function () {});
})();

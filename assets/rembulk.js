/* khata.niveshsahayak.in: "Remind from your outstanding list" (byof2-khata, 28 Nov 2026).
   The visitor drops the outstanding report they already have (Tally Bills Receivable / Group Outstandings / Ledger Outstandings as Excel
   or XML, Busy's outstanding report, any Excel or CSV) and sends one payment reminder per party through THEIR OWN WhatsApp, one after
   another: NSBiz opens WhatsApp with the number and the message typed in (NSMsg.ownLink: whatsapp:// in NSBiz for Windows, WhatsApp Web
   on a computer browser, wa.me on a phone; SMS app as an option on phones) and the visitor presses Send there. Never sent automatically.
   Read in the browser by NSRemind (remcore.js); nothing leaves the device and nothing is saved to the NSBiz books.
   Kept in this browser: localStorage ns-rem (settings, template, do-not-send list), ns-rem-list (the list read, the queue and what happened
   to each party, so a reload continues where it stopped), ns-rem-lay (column choices by heading signature).
   English only for now (owner rule 28 Nov). Designed & Developed by: Nivesh Sahayak (https://www.niveshsahayak.in/) */
(function () {
  'use strict';
  var NS = window.NS, R = window.NSRemind, M = window.NSMsg; if (!NS || !R || !M) return;
  var $ = NS.$, esc = NS.esc;
  var KEY = 'ns-rem', LKEY = 'ns-rem-list', LAY = 'ns-rem-lay', XLJS = '/assets/vendor/xlsx-0.20.3.full.min.js', QRJS = '/assets/vendor/qrcode-generator-1.4.4.min.js';
  var PROG = !!(window.chrome && window.chrome.webview);
  var DEF = { tpl: R.DEFTPL, business: '', phone: '', upi: '', link: '', ch: 'mywa', open: PROG ? 'app' : 'web', min: 1, minDays: 0, stop: {} };
  var S = load(), L = loadList(), root, q = '', shown = 100, colsOpen = false, F = null, busyMsg = '';
  /* L = { file, asOf, kind, parties, none, skipped, status: { k: { what, at } }, hand: { ids, i } | null, mob: { k: mobile typed } } ; F = raw sheet for the Columns screen */

  function load() { var s = {}; try { s = JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (e) {} s = Object.assign({}, DEF, s); s.stop = s.stop || {}; return s; }
  function save() { var o = Object.assign({}, S); if (o.tpl === R.DEFTPL) delete o.tpl; try { localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) {} }   /* the ready message is not stored, so a better ready message reaches everyone who did not change it */
  function loadList() { try { var x = JSON.parse(localStorage.getItem(LKEY) || 'null'); return x && x.parties ? x : null; } catch (e) { return null; } }
  function saveList() { try { if (L) localStorage.setItem(LKEY, JSON.stringify(L)); else localStorage.removeItem(LKEY); } catch (e) {} }
  function layouts() { try { return JSON.parse(localStorage.getItem(LAY) || '{}'); } catch (e) { return {}; } }
  function phoneDev() { return /Android|iPhone|iPad|iPod/i.test(navigator.userAgent); }
  function pl(n, one, many) { return n + ' ' + (n === 1 ? one : many || one + 's'); }
  function money(n) { return '₹' + R.inr(n); }
  function biz() { var b = window.NSApp && NSApp.biz && NSApp.biz(); return b || {}; }
  function ctx() { var b = biz(); return { business: S.business || b.name || '', phone: S.phone || b.phone || '', upi: S.upi || b.upi || '', link: S.link }; }
  function mobOf(p) { return (L.mob && L.mob[p.k] != null ? L.mob[p.k] : p.mobile) || ''; }
  function list() { return L ? L.parties.map(function (p) { var m = R.mobile(mobOf(p)); return Object.assign({}, p, { mobile: m }); }) : []; }
  function stopAll() { return S.stop; }

  /* NSBiz's own saved parties give a mobile number to a party whose row has none (same name) */
  function fillFromBooks() {
    if (!L || !window.NSApp || !NSApp.phoneOf || !window.NSBooks) return 0;
    var n = 0; L.mob = L.mob || {};
    L.parties.forEach(function (p) { if (R.mobile(mobOf(p))) return; var ph = R.mobile(NSApp.phoneOf(NSBooks.key(p.name))); if (ph) { L.mob[p.k] = ph; n++; } });
    if (n) saveList(); return n;
  }

  /* ---------- reading the file ---------- */
  function decode(buf) {
    var b = new Uint8Array(buf);
    if (b[0] === 0xFF && b[1] === 0xFE) return new TextDecoder('utf-16le').decode(buf);
    if (b[0] === 0xFE && b[1] === 0xFF) return new TextDecoder('utf-16be').decode(buf);
    if (b.length > 3 && b[1] === 0 && b[3] === 0) return new TextDecoder('utf-16le').decode(buf);   /* Tally writes UTF-16 without a mark sometimes */
    return new TextDecoder().decode(buf);
  }
  function sheets(f, buf) {
    var b = new Uint8Array(buf);
    return NS.loadJs(XLJS, '').then(function () {
      var X = window.XLSX, csv = !(b[0] === 0x50 && b[1] === 0x4B) && !(b[0] === 0xD0 && b[1] === 0xCF);
      var wb = csv ? X.read(decode(buf), { type: 'string', raw: true }) : X.read(buf, { type: 'array', cellDates: false });
      return wb.SheetNames.map(function (n) { return { name: n, rows: X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }).map(function (r) { return r.map(function (c) { return c == null ? '' : String(c).trim(); }); }) }; });
    }, function () {
      if (b[0] === 0xD0 && b[1] === 0xCF) throw new Error('This old .xls file needs the internet once to be read. Or save it in Excel as .xlsx or CSV.');
      return NS.xlsx.read(new File([buf], f.name)).then(function (rows) { return [{ name: f.name, rows: rows }]; });
    });
  }
  function readFile(f) {
    busy('Reading ' + f.name + '…');
    return f.arrayBuffer().then(function (buf) {
      var head = decode(buf.slice(0, 4000)).replace(/^﻿/, '').trim();
      if (/\.xml$/i.test(f.name) || /^<(\?xml|ENVELOPE)/i.test(head)) {
        var X = R.readTallyXml(decode(buf), { asOf: asOf() });
        if (X.err) throw new Error(X.err);
        return took(f.name, X.kind, X.lines, X.skipped, null);
      }
      return sheets(f, buf).then(function (book) {
        var best = null, why = '';
        book.forEach(function (sh) { if (best) return; var r = R.readRows(sh.rows, { asOf: asOf(), name: f.name, sheet: sh.name, layouts: layouts() }); if (r.err) { why = why || r.err; if (!F) F = { name: f.name, sheet: sh.name, rows: sh.rows, map: r.map || {}, row: r.row || 0 }; return; } if (r.lines.length) best = { sh: sh, r: r }; });
        if (!best) { colsOpen = true; if (!F) F = { name: f.name, sheet: book[0] && book[0].name, rows: (book[0] && book[0].rows) || [], map: {}, row: 0 }; throw new Error(why || 'No rows with a party and an amount due were found.'); }
        F = { name: f.name, sheet: best.sh.name, rows: best.sh.rows, map: best.r.map, row: best.r.row };
        return took(f.name, best.r.single ? 'Ledger outstandings of one party' : 'Excel / CSV list' + (book.length > 1 ? ' (sheet ' + best.sh.name + ')' : ''), best.r.lines, best.r.skipped, best.r);
      });
    }).then(function () { busy(''); }, function (e) { busy(''); NS.toast(e.message || 'The file could not be read.', 'bad'); draw(); });
  }
  function took(name, kind, lines, skipped, r) {
    var P = R.parties(lines);
    L = { file: name, asOf: asOf(), kind: kind, parties: P.due, none: P.none, skipped: skipped, rows: lines.length, status: {}, hand: null, mob: {}, sample: /sample/i.test(name) };
    var n = fillFromBooks(); L.fromBooks = n; colsOpen = false; q = ''; shown = 100;
    saveList(); draw();
    var t0 = $('[data-rb-step="2"]', root); if (t0) t0.scrollIntoView({ block: 'start', behavior: 'smooth' });
    return L;
  }
  function asOf() { var i = root && $('[data-rb=asof]', root); return (i && i.value) || R.today(); }
  /* the Columns screen: the visitor names the columns of a sheet the reader could not place */
  function rereadCols() {
    var map = {}; [].forEach.call(root.querySelectorAll('[data-rbcol]'), function (s) { if (s.value !== '') map[s.getAttribute('data-rbcol')] = +s.value; });
    var row = +(($('[data-rb=headrow]', root) || {}).value || 1) - 1;
    if (!('party' in map) || !('amt' in map || 'dr' in map)) { NS.toast('Choose at least the party name and the amount due (or Debit).', 'bad'); return; }
    var heads = F.rows[row] || [], lay = layouts(); lay[NSTpl.sig(heads)] = map; try { localStorage.setItem(LAY, JSON.stringify(lay)); } catch (e) {}
    var r = R.readRows(F.rows, { asOf: asOf(), name: F.name, layouts: lay });
    if (r.err || !r.lines.length) { NS.toast(r.err || 'No rows were read with these columns.', 'bad'); return; }
    took(F.name, 'Excel / CSV list (your columns)', r.lines, r.skipped, r);
  }

  /* ---------- the sample ---------- */
  function sampleBook() { return [['Bills Receivable', R.sampleRows(R.today()), [12, 14, 28, 16, 16, 12, 14]]]; }
  function trySample() { var rows = R.sampleRows(asOf()), r = R.readRows(rows, { asOf: asOf(), name: 'sample' }); F = { name: 'Sample outstanding list', rows: rows, map: r.map, row: r.row }; took('Sample outstanding list (made-up)', 'Tally Bills Receivable layout (sample)', r.lines, r.skipped, r); if (!S.business) { S.business = 'Sample Traders'; S.upi = S.upi || 'sampletraders@okaxis'; save(); draw(); } }
  function download(name, blob) { var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000); }
  function xlsxBlob(sheetsList) { var z = NS.xlsx.write(sheetsList); return z instanceof Blob ? z : new Blob([z], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); }

  /* ---------- sending: one party after another through the visitor's own WhatsApp (or SMS app) ---------- */
  function howOpen() { if (S.ch === 'mysms') return /iPhone|iPad|iPod/i.test(navigator.userAgent) ? 'ios' : 'phone'; if (phoneDev()) return 'phone'; return S.open || (PROG ? 'app' : 'web'); }
  function openOwn(to, text) {
    var link = M.ownLink(S.ch, to, text, howOpen());
    if (PROG || /^https:/.test(link)) window.open(link, /^https:/.test(link) ? 'ns-wa' : '_blank'); else location.href = link;
    return link;
  }
  function Q() { return R.queue(list(), { stop: stopAll(), min: +S.min || 0, minDays: +S.minDays || 0 }); }
  function start() {
    var G = Q();
    if (!G.go.length) { NS.toast('Nobody to remind: check the mobile numbers and the do-not-send list.', 'bad'); return; }
    var bad = R.unknownPlaces(S.tpl); if (bad.length) { NS.toast('Your message has places NSBiz does not know: ' + bad.join(' ') + '. Use only the ones listed.', 'bad'); return; }
    var todo = G.go.filter(function (p) { return !(L.status[p.k] && L.status[p.k].what === 'opened in ' + chName()); });
    L.hand = { ids: (todo.length ? todo : G.go).map(function (p) { return p.k; }), i: 0, opened: 0, skipped: 0 }; saveList(); draw();
    var c = $('[data-rb-hand]', root); if (c) { c.scrollIntoView({ block: 'center', behavior: 'smooth' }); var b = $('[data-rbdo=open]', c); if (b) b.focus(); }
  }
  function chName() { return S.ch === 'mysms' ? 'SMS app' : 'WhatsApp'; }
  function cur() { if (!L || !L.hand) return null; var k = L.hand.ids[L.hand.i], p = list().filter(function (x) { return x.k === k; })[0]; return p || null; }
  function step(what) {
    var p = cur(); if (!p) return;
    if (what === 'open') { openOwn(p.mobile, R.message(S.tpl, p, ctx())); L.status[p.k] = { what: 'opened in ' + chName(), at: Date.now() }; L.hand.opened++; }
    else { L.status[p.k] = { what: 'skipped', at: Date.now() }; L.hand.skipped++; }
    L.hand.i++;
    if (L.hand.i >= L.hand.ids.length) { var h = L.hand; L.hand = null; NS.toast('Done: ' + h.opened + ' opened in ' + chName() + ', ' + h.skipped + ' skipped. Check in ' + chName() + ' that each one was sent.'); }
    saveList(); draw();
    var b = $('[data-rbdo=open]', root); if (b) b.focus();
  }
  function stopHand() { if (!L || !L.hand) return; var h = L.hand; L.hand = null; saveList(); NS.toast('Stopped after ' + h.opened + ' opened. Press Start again to go on with the ones not done.'); draw(); }

  /* the UPI QR for the amount (offline: the library is kept on this site) */
  function qrSvg(text) { try { var qr = window.qrcode(0, 'M'); qr.addData(text); qr.make(); return qr.createSvgTag({ cellSize: 4, margin: 2, scalable: true }); } catch (e) { return ''; } }
  function qrPng(text, label) {
    return new Promise(function (res) {
      var qr = window.qrcode(0, 'M'); qr.addData(text); qr.make();
      var n = qr.getModuleCount(), cell = 8, pad = 24, w = n * cell + pad * 2, c = document.createElement('canvas'); c.width = w; c.height = w + 60;
      var g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.fillStyle = '#000';
      for (var r = 0; r < n; r++) for (var k = 0; k < n; k++) if (qr.isDark(r, k)) g.fillRect(pad + k * cell, pad + r * cell, cell, cell);
      g.font = 'bold 18px sans-serif'; g.textAlign = 'center'; g.fillText(label.slice(0, 40), w / 2, w + 30);
      c.toBlob(res, 'image/png');
    });
  }

  /* ---------- the screen ---------- */
  function busy(m) { busyMsg = m || ''; var b = root && $('[data-rb-busy]', root); if (b) { b.hidden = !busyMsg; b.textContent = busyMsg; } }
  function sel(attr, opts, curv) { return '<select ' + attr + '>' + opts.map(function (o) { return '<option value="' + esc(o[0]) + '"' + (String(o[0]) === String(curv) ? ' selected' : '') + '>' + esc(o[1]) + '</option>'; }).join('') + '</select>'; }
  function step1() {
    return '<div class="card rb-card" data-rb-step="1"><h3>1. Your outstanding list</h3>' +
      '<div class="rb-drop" data-rb-drop><b>Drop the file here, or <label class="rb-pick" for="rb-file">choose it</label></b><span>Excel (.xlsx, .xls), CSV or Tally XML: Bills Receivable, Group Outstandings, Ledger Outstandings, Busy outstanding report, or your own sheet with party, mobile and amount due.</span>' +
      '<input type="file" id="rb-file" class="sr" accept=".xlsx,.xls,.csv,.txt,.xml"></div>' +
      '<p class="rb-row"><label>Count ages up to <input type="date" data-rb="asof" value="' + esc((L && L.asOf) || R.today()) + '"></label> ' +
      '<button type="button" class="btn" data-rbdo="sample">Try with the sample list</button> <button type="button" class="btn" data-rbdo="sampledl">Download the sample file</button></p>' +
      '<p class="soft" data-rb-busy role="status"' + (busyMsg ? '' : ' hidden') + '>' + esc(busyMsg) + '</p>' +
      '<p class="soft">The file is read here in your browser. Nothing is uploaded, and nothing is added to your khata or books.</p>' + (colsOpen && F ? colsHtml() : '') + '</div>';
  }
  function colsHtml() {
    var rows = F.rows || [], hr = F.row || 0, heads = rows[hr] || [], opts = [['', 'not in the file']].concat(heads.map(function (h, i) { return [String(i), (h || 'Column ' + (i + 1))]; }));
    return '<div class="rb-cols"><h4>Which column is which?</h4><p class="soft">' + esc(F.name) + (F.sheet ? ', sheet ' + esc(F.sheet) : '') + '. NSBiz remembers your choice for files with the same headings.</p>' +
      '<p><label>The headings are in row <input type="number" min="1" max="40" data-rb="headrow" value="' + (hr + 1) + '" style="width:5em"></label></p><div class="rb-grid">' +
      R.COLS.map(function (c) { return '<label>' + esc(c.h) + (c.need ? ' *' : '') + ' ' + sel('data-rbcol="' + c.k + '"', opts, c.k in (F.map || {}) ? F.map[c.k] : '') + '</label>'; }).join('') +
      '</div><div class="btns"><button type="button" class="btn primary" data-rbdo="cols">Read with these columns</button> <button type="button" class="btn" data-rbdo="colsoff">Close</button></div></div>';
  }
  function step2() {
    if (!L) return '';
    var all = list(), T = R.totals(all), names = R.bucketNames(), G = Q(), noMob = all.filter(function (p) { return !p.mobile; }).length;
    var f = q ? all.filter(function (p) { return p.name.toLowerCase().indexOf(q.toLowerCase()) >= 0 || (p.mobile || '').indexOf(q) >= 0; }) : all;
    var h = '<div class="card rb-card" data-rb-step="2"><h3>2. What was read</h3>' + (L.sample ? '<p class="b1-warn">This is made-up sample data: the mobile numbers are not real people\'s. Try the steps, but do not press Send in WhatsApp.</p>' : '') +
      '<p><b>' + esc(L.file) + '</b> <span class="soft">(' + esc(L.kind) + ', ' + L.rows + ' rows, ages counted to ' + esc(L.asOf.split('-').reverse().join('-')) + ')</span></p>' +
      '<p><b>' + pl(T.n, 'party', 'parties') + '</b> ' + (T.n === 1 ? 'owes' : 'owe') + ' you <b>' + money(T.due) + '</b>.' + (noMob ? ' ' + pl(noMob, 'party has', 'parties have') + ' no mobile number: type it below.' : '') + (L.fromBooks ? ' ' + pl(L.fromBooks, 'mobile number') + ' came from your saved parties.' : '') + '</p>' +
      '<div class="tbl-wrap"><table class="tbl rb-age"><thead><tr>' + names.map(function (n) { return '<th class="n">' + esc(n) + '</th>'; }).join('') + (T.nodate ? '<th class="n">Age not known</th>' : '') + '</tr></thead><tbody><tr>' +
      T.ages.map(function (v) { return '<td class="n">' + money(v) + '</td>'; }).join('') + (T.nodate ? '<td class="n">' + money(T.nodate) + '</td>' : '') + '</tr></tbody></table></div>';
    var skip = (L.skipped || []).filter(function (s) { return !s.keep; }), warn = (L.skipped || []).filter(function (s) { return s.keep; });
    if (skip.length || warn.length || L.none.length) h += '<details class="rb-skip"><summary>' + (skip.length ? pl(skip.length, 'row') + ' left out' : 'Rows left out') + (L.none.length ? ', ' + pl(L.none.length, 'party', 'parties') + ' with nothing due' : '') + (warn.length ? ', ' + pl(warn.length, 'note') : '') + '</summary><ul>' +
      skip.concat(warn).slice(0, 300).map(function (s) { return '<li>Row ' + s.row + ': ' + esc(s.why) + '</li>'; }).join('') + L.none.map(function (p) { return '<li>' + esc(p.name) + ': ' + esc(p.why) + (p.due ? ' (' + money(Math.abs(p.due)) + ')' : '') + '</li>'; }).join('') + '</ul></details>';
    h += '<p class="rb-row"><label>Find <input type="search" data-rb="q" value="' + esc(q) + '" placeholder="party or mobile"></label> <button type="button" class="btn sm" data-rbdo="colson">Change columns</button> <button type="button" class="btn sm" data-rbdo="clear">Start with another file</button></p>' +
      '<div class="tbl-wrap"><table class="tbl rb-list"><thead><tr><th>Party</th><th>Mobile</th><th class="n">Due</th><th class="n">Oldest (days)</th><th>Send?</th><th>Status</th></tr></thead><tbody>' +
      f.slice(0, shown).map(function (p) {
        var stp = S.stop['n:' + p.k] || (p.mobile && S.stop[p.mobile]), st = L.status[p.k];
        return '<tr><th scope="row">' + esc(p.name) + (p.bills.length ? '<br><small class="soft">' + esc(R.billText(p, 3)) + '</small>' : '') + '</th>' +
          '<td><input type="tel" inputmode="tel" class="rb-mob" data-rbmob="' + esc(p.k) + '" value="' + esc(mobOf(p)) + '" aria-label="Mobile of ' + esc(p.name) + '" placeholder="mobile"' + (mobOf(p) && !p.mobile ? ' aria-invalid="true"' : '') + '></td>' +
          '<td class="n">' + money(p.due) + '</td><td class="n">' + (p.days == null ? '—' : p.days) + '</td>' +
          '<td><label class="rb-ok"><input type="checkbox" data-rbstop="' + esc(p.k) + '"' + (stp ? '' : ' checked') + '> ' + (stp ? 'do not send' : 'send') + '</label></td>' +
          '<td>' + (st ? esc(st.what) : '') + '</td></tr>';
      }).join('') + '</tbody></table></div>' +
      (f.length > shown ? '<p><button type="button" class="btn sm" data-rbdo="more">Show more (' + (f.length - shown) + ' left)</button></p>' : '') + '</div>';
    return h;
  }
  function step3() {
    if (!L) return '';
    var c = ctx(), G = Q(), first = G.go[0] || list()[0], done = Object.keys(L.status).length;
    var h = '<div class="card rb-card" data-rb-step="3"><h3>3. The message and sending</h3><div class="rb-grid">' +
      '<label>Your business name <input data-rbs="business" value="' + esc(S.business) + '" placeholder="' + esc(biz().name || 'as on your bills') + '"></label>' +
      '<label>Your phone <input data-rbs="phone" value="' + esc(S.phone) + '" placeholder="' + esc(biz().phone || '') + '"></label>' +
      '<label>Your UPI ID <input data-rbs="upi" value="' + esc(S.upi) + '" placeholder="' + esc(biz().upi || 'name@bank') + '"></label>' +
      '<label>Payment link (optional) <input data-rbs="link" value="' + esc(S.link) + '" placeholder="https://… from your bank or payment app"></label>' +
      '<label>Only parties owing at least ₹ <input type="number" min="0" data-rbs="min" value="' + esc(S.min) + '"></label>' +
      '<label>Only dues at least this many days old <input type="number" min="0" data-rbs="minDays" value="' + esc(S.minDays) + '"></label>' +
      '<label>Send by ' + sel('data-rbs="ch"', [['mywa', 'My own WhatsApp (free)']].concat(phoneDev() ? [['mysms', 'My phone\'s SMS (free)']] : []), S.ch) + '</label>' +
      (phoneDev() ? '' : '<label>Open WhatsApp in ' + sel('data-rbs="open"', [['web', 'WhatsApp Web (browser)'], ['app', 'the WhatsApp app on this computer']], S.open) + '</label>') +
      '</div><label class="rb-tpl">Message (you can change it) <textarea data-rbs="tpl" rows="9">' + esc(S.tpl) + '</textarea></label>' +
      '<p class="soft">Places: ' + R.PLACES.map(function (p) { return '<code>{' + p + '}</code>'; }).join(' ') + '. {days} = days since the oldest bill (or overdue days, when your file has them); {upilink} = a UPI payment address with the amount. A line whose place is empty (no UPI ID, no link) is left out. <button type="button" class="btn sm" data-rbdo="deftpl">Back to the ready message</button></p>' +
      (first ? '<div class="rb-prev"><b>How it will look for ' + esc(first.name) + ':</b><p class="msg-text">' + esc(R.message(S.tpl, first, c)) + '</p></div>' : '') +
      '<p><b>' + pl(G.go.length, 'reminder') + '</b> ready' + (G.left.length ? ', ' + G.left.length + ' left out' : '') + (done ? ', ' + done + ' done already' : '') + '.</p>' +
      (G.left.length ? '<details class="rb-skip"><summary>Who is left out and why</summary><ul>' + G.left.map(function (x) { return '<li>' + esc(x.name) + ': ' + esc(x.why) + '</li>'; }).join('') + '</ul></details>' : '') +
      '<div class="btns"><button type="button" class="btn primary" data-rbdo="start"' + (G.go.length ? '' : ' disabled') + '>Start sending, one by one</button> <button type="button" class="btn" data-rbdo="report">Download the report (Excel)</button> <button type="button" class="btn" data-rbdo="reportcsv">Report as CSV</button></div>' +
      '<p class="soft">NSBiz opens ' + (S.ch === 'mysms' ? 'your SMS app' : 'WhatsApp') + ' with the number and the message typed in; you press Send there, then come back for the next party. Nothing is sent by itself: WhatsApp does not allow automatic sending from a normal number, and many messages at once to people who did not save your number can get it blocked. Send to your own customers only.</p></div>';
    return h;
  }
  function handHtml() {
    var p = cur(); if (!p) return '';
    var c = ctx(), text = R.message(S.tpl, p, c), up = R.upiLink(c.upi, c.business, p.due, 'Dues ' + p.name);
    return '<div class="card rb-card rb-hand" data-rb-hand><h3>' + (L.hand.i + 1) + ' of ' + L.hand.ids.length + ': ' + esc(p.name) + '</h3>' +
      '<p><b>' + money(p.due) + '</b> due · <span class="soft">' + esc(p.mobile) + (p.days != null ? ' · oldest ' + p.days + ' days' : '') + '</span></p>' +
      '<div class="rb-handbody"><p class="msg-text">' + esc(text) + '</p>' + (up && window.qrcode ? '<figure class="rb-qr">' + qrSvg(up) + '<figcaption>UPI ' + money(p.due) + '</figcaption></figure>' : '') + '</div>' +
      '<div class="btns"><button type="button" class="btn primary" data-rbdo="open">Open in ' + (S.ch === 'mysms' ? 'the SMS app' : 'WhatsApp') + '</button> <button type="button" class="btn" data-rbdo="skip">Skip</button> <button type="button" class="btn" data-rbdo="stop">Stop</button>' +
      (up ? ' <button type="button" class="btn" data-rbdo="qr">Save the QR picture</button>' : '') + '</div>' +
      '<p class="soft">Press Send in ' + chName() + ', then come back here: the next party is ready. ' + (up ? 'To send the QR too, save the picture and attach it in the chat.' : 'Type your UPI ID in step 3 to show a payment QR here.') + '</p></div>';
  }
  function stopHtml() {
    var k = Object.keys(S.stop);
    return '<details class="card rb-card rb-stop"><summary><b>Do-not-send list</b> <span class="soft">(' + k.length + ')</span></summary><p class="soft">Numbers or parties that asked not to get reminders, or that you do not want to message. They are left out every time, in every file.</p>' +
      '<p class="rb-row"><input data-rb="stopadd" placeholder="mobile numbers, comma between"> <button type="button" class="btn sm" data-rbdo="stopadd">Add</button></p>' +
      (k.length ? '<ul>' + k.sort().map(function (x) { return '<li>' + esc(x.indexOf('n:') === 0 ? x.slice(2) + ' (party)' : x) + ' <button type="button" class="btn sm" data-rbunstop="' + esc(x) + '">Remove</button></li>'; }).join('') + '</ul>' : '<p class="soft">The list is empty.</p>') + '</details>';
  }
  function draw() {
    if (!root) return;
    var a = document.activeElement, keep = a && root.contains(a) && (a.getAttribute('data-rbs') || a.getAttribute('data-rb') || a.getAttribute('data-rbmob')), pos = keep && a.selectionStart;
    var keepSel = keep ? '[' + (a.hasAttribute('data-rbs') ? 'data-rbs="' + a.getAttribute('data-rbs') : a.hasAttribute('data-rbmob') ? 'data-rbmob="' + a.getAttribute('data-rbmob') : 'data-rb="' + a.getAttribute('data-rb')) + '"]' : '';
    $('[data-rb-body]', root).innerHTML = (L && L.hand ? handHtml() : '') + step1() + step2() + step3() + stopHtml();
    if (keepSel) { var e = $(keepSel, root); if (e) { e.focus(); try { if (pos != null) e.setSelectionRange(pos, pos); } catch (x) {} } }
  }
  var redrawT = 0; function later() { clearTimeout(redrawT); redrawT = setTimeout(draw, 350); }

  function wire() {
    root.addEventListener('click', function (e) {
      var b = e.target.closest('[data-rbdo]'), u = e.target.closest('[data-rbunstop]');
      if (u) { delete S.stop[u.getAttribute('data-rbunstop')]; save(); draw(); return; }
      if (!b) return; var op = b.getAttribute('data-rbdo');
      if (op === 'sample') trySample();
      if (op === 'sampledl') download('Sample outstanding list (made-up).xlsx', xlsxBlob(sampleBook()));
      if (op === 'cols') rereadCols();
      if (op === 'colson') { colsOpen = true; draw(); var c = $('.rb-cols', root); if (c) c.scrollIntoView({ block: 'center' }); }
      if (op === 'colsoff') { colsOpen = false; draw(); }
      if (op === 'clear') { if (!L || confirm('Clear this list and what was sent? Download the report first if you need it.')) { L = null; F = null; saveList(); draw(); } }
      if (op === 'more') { shown += 200; draw(); }
      if (op === 'deftpl') { S.tpl = R.DEFTPL; save(); draw(); }
      if (op === 'start') start();
      if (op === 'open') step('open');
      if (op === 'skip') step('skip');
      if (op === 'stop') stopHand();
      if (op === 'qr') { var p = cur(), c2 = ctx(); if (p) qrPng(R.upiLink(c2.upi, c2.business, p.due, 'Dues ' + p.name), 'Pay ' + money(p.due) + ' to ' + (c2.business || c2.upi)).then(function (bl) { download('UPI QR ' + p.name.replace(/[\\/:*?"<>|]+/g, ' ') + '.png', bl); }); }
      if (op === 'report' || op === 'reportcsv') {
        if (!L) return; var rows = R.reportRows(list(), L.status), nm = 'Payment reminders ' + R.today();
        if (op === 'report') download(nm + '.xlsx', xlsxBlob([['Reminders', rows, [30, 13, 14, 12, 12, 12, 12, 12, 50, 22, 20]]]));
        else download(nm + '.csv', new Blob(['﻿' + rows.map(function (r) { return r.map(function (v) { v = String(v == null ? '' : v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(','); }).join('\r\n')], { type: 'text/csv' }));
      }
      if (op === 'stopadd') { var inp = $('[data-rb=stopadd]', root); String(inp && inp.value || '').split(/[\n,;]+/).forEach(function (x) { var m = R.mobile(x); if (m) S.stop[m] = 1; }); save(); draw(); }
    });
    root.addEventListener('change', function (e) {
      var el = e.target;
      if (el.id === 'rb-file' && el.files && el.files[0]) { readFile(el.files[0]); el.value = ''; return; }
      if (el.getAttribute('data-rbstop') != null) { var k = el.getAttribute('data-rbstop'), p = list().filter(function (x) { return x.k === k; })[0]; if (el.checked) { delete S.stop['n:' + k]; if (p && p.mobile) delete S.stop[p.mobile]; } else S.stop['n:' + k] = 1; save(); draw(); return; }
      if (el.getAttribute('data-rb') === 'asof' && L && F) { var r = R.readRows(F.rows, { asOf: asOf(), name: F.name, layouts: layouts() }); if (!r.err) { var st = L.status, mob = L.mob; took(L.file, L.kind, r.lines, r.skipped, r); L.status = st; L.mob = mob; saveList(); draw(); } }
      if (el.getAttribute('data-rbs') === 'ch' || el.getAttribute('data-rbs') === 'open') { S[el.getAttribute('data-rbs')] = el.value; save(); draw(); }
    });
    root.addEventListener('input', function (e) {
      var el = e.target, s = el.getAttribute('data-rbs');
      if (s && s !== 'ch' && s !== 'open') { S[s] = el.value; save(); later(); return; }
      if (el.getAttribute('data-rb') === 'q') { q = el.value; shown = 100; later(); return; }
      var m = el.getAttribute('data-rbmob'); if (m != null && L) { L.mob = L.mob || {}; L.mob[m] = el.value; saveList(); later(); }
    });
    var d = $('[data-rb-drop]', root);
    root.addEventListener('dragover', function (e) { if (e.target.closest('[data-rb-drop]')) { e.preventDefault(); e.target.closest('[data-rb-drop]').classList.add('on'); } });
    root.addEventListener('dragleave', function (e) { var z = e.target.closest('[data-rb-drop]'); if (z) z.classList.remove('on'); });
    root.addEventListener('drop', function (e) { var z = e.target.closest('[data-rb-drop]'); if (!z) return; e.preventDefault(); z.classList.remove('on'); var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; if (f) readFile(f); });
    void d;
  }

  function init() {
    root = document.querySelector('[data-rembulk]');
    if (!root) {   /* translated home pages: the section is added after the khata itself */
      var after = document.querySelector('.tool-app'); if (!after) return;
      root = document.createElement('section'); root.className = 'wrap rembulk'; root.setAttribute('data-rembulk', ''); root.id = 'remind';
      root.innerHTML = '<h2>Remind from your Tally or Excel outstanding list</h2><p class="soft">Drop the outstanding report you already have and send each party a WhatsApp reminder with your UPI QR, one after another.</p><div data-rb-body></div>';
      after.parentNode.insertBefore(root, after.nextSibling);
    }
    if (!$('[data-rb-body]', root)) { var bdy = document.createElement('div'); bdy.setAttribute('data-rb-body', ''); root.appendChild(bdy); }
    wire(); draw();
    NS.loadJs(QRJS, '').then(function () { if (L && L.hand) draw(); }, function () {});
    /* NSBiz's parties arrive later than the page: fill mobiles then, and show the business's own name, phone and UPI ID */
    document.addEventListener('ns-data', function () { if (fillFromBooks()) draw(); else if (L && !S.business) later(); });
    setTimeout(function () { if (L && fillFromBooks()) draw(); else if (L) draw(); }, 2500);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
  window.NSRemBulk = { draw: draw, read: readFile, sample: trySample };
})();

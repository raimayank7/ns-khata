/* khata.niveshsahayak.in: "Balance confirmation letters for every party" (byof3-confirm, 28 Nov 2026).
   The visitor drops the outstanding list they already have (Tally Bills Receivable / Payable, Group Outstandings, Group Summary
   of Sundry Debtors or Creditors as Excel or XML, Busy's outstanding report, or any Excel / CSV with party, balance, Dr / Cr and
   optional bills, address, mobile and email) and gets one confirmation letter per party with a reply slip: all letters in one PDF,
   one PDF each (zip), Print all, a register (Excel) to tick off the replies, and a WhatsApp and an email text for every party,
   opened one after another in their own WhatsApp (NSMsg.ownLink, as khata's reminders) or email program. Nothing is sent by itself.
   Read in this browser by NSConfirm (bconfcore.js, on top of NSRemind's outstanding reader); nothing is uploaded and nothing is
   saved to the NSBiz books. Kept in this browser: localStorage ns-bconf (letter settings, texts if changed), ns-bconf-list (the
   list read, contact details typed, what was sent), ns-bconf-lay (column choices by heading signature).
   Screen words: engine/i18n.js EN keys bcf.* (English only for now, owner rule 28 Nov). The letter is English, like every printed
   document of NSBiz. Designed & Developed by: Nivesh Sahayak (https://www.niveshsahayak.in/) */
(function () {
  'use strict';
  var NS = window.NS, C = window.NSConfirm, R = window.NSRemind, M = window.NSMsg; if (!NS || !C || !R || !M) return;
  var $ = NS.$, esc = NS.esc;
  function L(k, v) { return NS.t('bcf.' + k, v); }
  var KEY = 'ns-bconf', LKEY = 'ns-bconf-list', LAY = 'ns-bconf-lay', XLJS = '/assets/vendor/xlsx-0.20.3.full.min.js';
  var PROG = !!(window.chrome && window.chrome.webview);
  var DEF = { asOn: '', date: '', days: 15, plain: 'dr', nil: 0, bills: 1, noReply: 0, prefix: 'BC', start: 1, name: '', addr: '', city: '', pin: '', gstin: '', phone: '', email: '', signer: '',
    wa: C.WATPL, subj: C.SUBJ, mail: C.MAILTPL, open: PROG ? 'app' : 'web' };
  var S = load(), D = loadList(), root, q = '', shown = 100, colsOpen = false, F = null, busyMsg = '', pick = '';
  /* D = { file, kind, list: [party], nil, skipped, rows, plainRows, off: { k: 1 }, mob: { k }, mail: { k }, status: { k: { what, at } }, hand: { via, ids, i, done, skipped } | null, sample } */

  function load() { var s = {}; try { s = JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (e) {} return Object.assign({}, DEF, s); }
  function save() { var o = Object.assign({}, S); ['wa', 'subj', 'mail'].forEach(function (k) { if (o[k] === DEF[k]) delete o[k]; }); try { localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) {} }
  function loadList() { try { var x = JSON.parse(localStorage.getItem(LKEY) || 'null'); return x && x.list ? x : null; } catch (e) { return null; } }
  function saveList() { try { if (D) localStorage.setItem(LKEY, JSON.stringify(D)); else localStorage.removeItem(LKEY); } catch (e) {} }
  function layouts() { try { return JSON.parse(localStorage.getItem(LAY) || '{}'); } catch (e) { return {}; } }
  function phoneDev() { return /Android|iPhone|iPad|iPod/i.test(navigator.userAgent); }
  function money(v) { return '₹' + C.inr(Math.abs(v)) + (C.side(v) ? ' ' + C.side(v) : ''); }
  function today() { return NS.today ? NS.today() : R.today(); }
  function asOn() { return S.asOn || C.defAsOn(today()); }
  function ldate() { return S.date || today(); }

  /* the business on the letter: what is typed here, else the business saved in NSBiz (khata's own) */
  function bizRec() { var b = window.NSApp && NSApp.biz && NSApp.biz(); return b || {}; }
  function biz() {
    var b = bizRec(), pick2 = function (k, alt) { return String(S[k] || b[alt || k] || '').trim(); };
    return { name: pick2('name'), addr: pick2('addr', b.addr ? 'addr' : 'address'), city: pick2('city'), pin: pick2('pin'), gstin: pick2('gstin'), phone: pick2('phone'), email: pick2('email') };
  }
  function ctx(i) { return { biz: biz(), asOn: asOn(), date: ldate(), days: +S.days || 15, signer: S.signer, noReply: !!+S.noReply, bills: !!+S.bills, prefix: S.prefix, start: +S.start || 1, ref: i == null ? '' : C.refNo(S.prefix, asOn(), (+S.start || 1) + i), credit: NS.creditLine ? NS.creditLine('st-ft') : '' }; }

  /* the list with what the visitor typed (mobile, email) and only the parties ticked */
  function full() { return D ? D.list.map(function (p) { var m = D.mob && D.mob[p.k] != null ? D.mob[p.k] : p.mobile, e = D.mail && D.mail[p.k] != null ? D.mail[p.k] : p.email; return Object.assign({}, p, { mobile: R.mobile(m) || '', mobTyped: m || '', email: String(e || '').trim() }); }) : []; }
  function chosen() { return full().filter(function (p) { return !(D.off && D.off[p.k]); }); }
  function indexOf(p) { var L0 = chosen(); for (var i = 0; i < L0.length; i++) if (L0[i].k === p.k) return i; return -1; }

  /* NSBiz's saved parties give a mobile, email and address to a party whose row has none (same name) */
  function fillFromBooks() {
    if (!D || !window.NSApp || !NSApp.data || !window.NSBooks) return 0;
    var n = 0, all = (NSApp.data().party || []), by = {};
    all.forEach(function (x) { by[NSBooks.key(x.name)] = x; });
    D.list.forEach(function (p) {
      var r = by[NSBooks.key(p.name)]; if (!r) return; var got = 0;
      if (!p.mobile && R.mobile(r.phone || r.mobile)) { p.mobile = R.mobile(r.phone || r.mobile); got = 1; }
      if (!p.email && r.email) { p.email = String(r.email); got = 1; }
      if (!p.addr && (r.addr || r.address)) { p.addr = String(r.addr || r.address); p.city = p.city || r.city || ''; p.pin = p.pin || r.pin || ''; got = 1; }
      if (!p.gstin && r.gstin) { p.gstin = String(r.gstin); got = 1; }
      n += got;
    });
    if (n) saveList(); return n;
  }

  /* ---------- reading the file ---------- */
  function decode(buf) {
    var b = new Uint8Array(buf);
    if (b[0] === 0xFF && b[1] === 0xFE) return new TextDecoder('utf-16le').decode(buf);
    if (b[0] === 0xFE && b[1] === 0xFF) return new TextDecoder('utf-16be').decode(buf);
    if (b.length > 3 && b[1] === 0 && b[3] === 0) return new TextDecoder('utf-16le').decode(buf);
    return new TextDecoder().decode(buf);
  }
  function sheets(f, buf) {
    var b = new Uint8Array(buf);
    return NS.loadJs(XLJS, '').then(function () {
      var X = window.XLSX, csv = !(b[0] === 0x50 && b[1] === 0x4B) && !(b[0] === 0xD0 && b[1] === 0xCF);
      var wb = csv ? X.read(decode(buf), { type: 'string', raw: true }) : X.read(buf, { type: 'array', cellDates: false });
      return wb.SheetNames.map(function (n) { return { name: n, rows: X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }).map(function (r) { return r.map(function (c) { return c == null ? '' : String(c).trim(); }); }) }; });
    }, function () {
      if (b[0] === 0xD0 && b[1] === 0xCF) throw new Error(L('oldxls'));
      return NS.xlsx.read(new File([buf], f.name)).then(function (rows) { return [{ name: f.name, rows: rows }]; });
    });
  }
  function plainOf(name, rows) { var top = (rows || []).slice(0, 8).map(function (r) { return (r || []).join(' '); }).join(' ') + ' ' + name; return /payable|creditor|supplier/i.test(top) && !/receivable|debtor/i.test(top) ? 'cr' : S.plain; }
  function readFile(f) {
    busy(L('reading', { f: f.name }));
    return f.arrayBuffer().then(function (buf) {
      var head = decode(buf.slice(0, 4000)).replace(/^﻿/, '').trim();
      if (/\.xml$/i.test(f.name) || /^<(\?xml|ENVELOPE)/i.test(head)) {
        var X = C.readTallyXml(decode(buf), { asOf: asOn() });
        if (X.err) throw new Error(X.err);
        F = null; return took(f.name, X.kind, X.lines, X.skipped, { plainRows: 0 });
      }
      return sheets(f, buf).then(function (book) {
        var best = null, why = '';
        book.forEach(function (sh) {
          if (best) return;
          var r = C.readRows(sh.rows, { asOf: asOn(), name: f.name, sheet: sh.name, layouts: layouts(), plain: plainOf(f.name, sh.rows) });
          if (r.err) { why = why || r.err; if (!F) F = { name: f.name, sheet: sh.name, rows: C.joinHeads(sh.rows).rows, map: r.map || {}, row: r.row || 0 }; return; }
          if (r.lines.length) best = { sh: sh, r: r };
        });
        if (!best) { colsOpen = true; if (!F) F = { name: f.name, sheet: book[0] && book[0].name, rows: (book[0] && book[0].rows) || [], map: {}, row: 0 }; throw new Error(why || L('nothing')); }
        F = { name: f.name, sheet: best.sh.name, rows: C.joinHeads(best.sh.rows).rows, map: best.r.map, row: best.r.row, plain: plainOf(f.name, best.sh.rows) };
        return took(f.name, best.r.single ? L('k_one') : L('k_xl') + (book.length > 1 ? ' (' + L('sheet') + ' ' + best.sh.name + ')' : ''), best.r.lines, best.r.skipped, best.r);
      });
    }).then(function () { busy(''); }, function (e) { busy(''); NS.toast(e.message || L('notread'), 'bad'); draw(); });
  }
  function took(name, kind, lines, skipped, r) {
    var P = C.parties(lines, { nil: !!+S.nil });
    D = { file: name, kind: kind, list: P.list, nil: P.nil, skipped: skipped || [], rows: lines.length, plainRows: (r && r.plainRows) || 0, plain: F && F.plain || S.plain, off: {}, mob: {}, mail: {}, status: {}, hand: null, sample: /sample/i.test(name) };
    D.fromBooks = fillFromBooks(); colsOpen = false; q = ''; shown = 100; pick = '';
    saveList(); draw();
    var s2 = $('[data-bc-step="2"]', root); if (s2 && s2.scrollIntoView) s2.scrollIntoView({ block: 'start', behavior: 'smooth' });
    return D;
  }
  function reread() {
    if (!D || !F || !F.rows) return;
    var keep = { off: D.off, mob: D.mob, mail: D.mail, status: D.status }, r = C.readRows(F.rows, { asOf: asOn(), name: F.name, layouts: layouts(), plain: D.plain });
    if (r.err) return; took(D.file, D.kind, r.lines, r.skipped, r); Object.assign(D, keep); D.plain = F.plain = D.plain; saveList(); draw();
  }
  function rereadCols() {
    var map = {}; [].forEach.call(root.querySelectorAll('[data-bccol]'), function (s) { if (s.value !== '') map[s.getAttribute('data-bccol')] = +s.value; });
    var row = +(($('[data-bc=headrow]', root) || {}).value || 1) - 1;
    if (!('party' in map) || !('amt' in map || 'dr' in map || 'cr' in map)) { NS.toast(L('needcols'), 'bad'); return; }
    var heads = F.rows[row] || [], lay = layouts(); lay[NSTpl.sig(heads)] = map; try { localStorage.setItem(LAY, JSON.stringify(lay)); } catch (e) {}
    var r = C.readRows(F.rows, { asOf: asOn(), name: F.name, layouts: lay, plain: F.plain || S.plain });
    if (r.err || !r.lines.length) { NS.toast(r.err || L('nocolrows'), 'bad'); return; }
    took(F.name, L('k_yours'), r.lines, r.skipped, r);
  }

  /* ---------- the sample ---------- */
  function trySample() {
    var rows = C.sampleRows(), r = C.readRows(rows, { asOf: asOn(), name: 'sample' });
    F = { name: 'Sample list', rows: C.joinHeads(rows).rows, map: r.map, row: r.row, plain: 'dr' };
    if (!S.name && !bizRec().name) { S.name = 'Sample Traders'; S.addr = '22 Main Bazar'; S.city = 'Ludhiana'; S.pin = '141008'; S.phone = '98140 00000'; S.signer = 'R. Kumar (Partner)'; save(); }
    took(L('samplename'), L('k_sample'), r.lines, r.skipped, r);
  }
  function download(name, blob) { var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000); }
  function xlsxBlob(list) { var z = NS.xlsx.write(list); return z instanceof Blob ? z : new Blob([z], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); }
  function csvBlob(rows) { return new Blob(['﻿' + rows.map(function (r) { return r.map(function (v) { v = String(v == null ? '' : v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(','); }).join('\r\n')], { type: 'text/csv' }); }

  /* ---------- the letters: PDF, zip, print, register ---------- */
  function letter(p) { return C.letterHtml(p, ctx(indexOf(p))); }
  function stage(list) {
    var hold = document.createElement('div'); hold.style.cssText = 'position:fixed;left:-12000px;top:0;width:794px'; hold.setAttribute('aria-hidden', 'true');
    hold.innerHTML = list.map(letter).join(''); document.body.appendChild(hold);
    return hold;
  }
  function needBiz() { if (biz().name) return true; NS.toast(L('needbiz'), 'bad'); var i = $('[data-bcs=name]', root); if (i) i.focus(); return false; }
  function base() { return L('fbase') + ' ' + asOn() + (D && D.sample ? ' sample' : ''); }
  function progBar(n) {
    var b = document.createElement('div'), st = { yes: false };
    b.className = 'ns-prog'; b.setAttribute('role', 'status');
    b.innerHTML = '<span class="ns-prog-t"></span><span class="ns-prog-b"><i></i></span><button type="button" class="btn sm">' + esc(NS.t('prog.cancel')) + '</button>';
    b.querySelector('button').addEventListener('click', function () { st.yes = true; this.disabled = true; });
    document.body.appendChild(b);
    st.step = function (i) { if (st.yes) throw new Error('cancel'); b.querySelector('.ns-prog-t').textContent = NS.t('prog.of', { l: L('making'), a: i, n: n }); b.querySelector('i').style.width = Math.round(i / n * 100) + '%'; return new Promise(function (r) { setTimeout(r, 0); }); };
    st.close = function () { b.remove(); };
    return st;
  }
  function onePdf() {
    if (!needBiz()) return; var list = chosen(); if (!list.length) { NS.toast(L('none'), 'bad'); return; }
    var hold = stage(list), b = progBar(list.length);
    NS.makePdf(hold.querySelectorAll('.sh'), 'a4', { step: b.step }).then(function (blob) { download(base() + '.pdf', blob); }, function (e) { if (!/cancel/.test(e && e.message)) NS.toast(L('pdffail'), 'bad'); })
      .then(function () { hold.remove(); b.close(); });
  }
  function zipPdfs() {
    if (!needBiz()) return; var list = chosen(); if (!list.length) { NS.toast(L('none'), 'bad'); return; }
    var files = {}, used = {};
    NS.eachSeq(list, function (p) {
      var hold = stage([p]);
      return NS.makePdf(hold.querySelectorAll('.sh'), 'a4').then(function (blob) { return blob.arrayBuffer(); }).then(function (buf) {
        var nm = C.fileName(p, ctx(indexOf(p)).ref), k = nm, i = 2; while (used[k]) k = nm + ' (' + (i++) + ')'; used[k] = 1;
        files[k + '.pdf'] = new Uint8Array(buf); hold.remove();
      }, function (e) { hold.remove(); throw e; });
    }, { label: L('making') }).then(function (r) {
      if (!Object.keys(files).length) return;
      download(base() + (r && r.cancelled ? ' (part)' : '') + '.zip', NS.xlsx.zip(files));
    }, function () { NS.toast(L('pdffail'), 'bad'); });
  }
  function onePdfOf(p) {
    if (!needBiz()) return;
    var hold = stage([p]), nm = C.fileName(p, ctx(indexOf(p)).ref);
    if (phoneDev() && NS.sharePdf) { NS.sharePdf(hold.querySelectorAll('.sh'), nm, 'a4', 'Balance confirmation', C.fill(S.wa, p, ctx(indexOf(p))), { phone: p.mobile, email: p.email }).then(function () { hold.remove(); }, function () { hold.remove(); }); return; }
    NS.makePdf(hold.querySelectorAll('.sh'), 'a4').then(function (blob) { download(nm + '.pdf', blob); NS.toast(L('saved', { f: nm + '.pdf' })); }, function () { NS.toast(L('pdffail'), 'bad'); }).then(function () { hold.remove(); });
  }
  function printAll() {
    if (!needBiz()) return; var list = chosen(); if (!list.length) { NS.toast(L('none'), 'bad'); return; }
    var box = document.createElement('div'); box.className = 'lg-print bc-printbox'; box.innerHTML = list.map(letter).join('');
    document.body.appendChild(box); document.documentElement.classList.add('bc-printing');
    var old = document.title; document.title = base();
    var done = function () { window.removeEventListener('afterprint', done); box.remove(); document.documentElement.classList.remove('bc-printing'); document.title = old; };
    window.addEventListener('afterprint', done);
    setTimeout(function () { window.print(); setTimeout(function () { if (box.isConnected && !matchMedia('print').matches) done(); }, 1500); }, 60);
  }
  function register(fmt) {
    if (!D) return; var c = ctx(); c.wa = S.wa; c.subj = S.subj; c.mail = S.mail;
    var rows = C.registerRows(chosen(), c, D.status), nm = L('regname') + ' ' + asOn();
    if (fmt === 'csv') download(nm + '.csv', csvBlob(rows));
    else download(nm + '.xlsx', xlsxBlob([['Confirmations', rows, [14, 30, 34, 17, 12, 26, 16, 7, 9, 18, 18, 16, 16, 24, 50, 40, 60]]]));
  }

  /* ---------- sending: one party after another through the visitor's own WhatsApp or email ---------- */
  function howOpen() { return phoneDev() ? 'phone' : (S.open || (PROG ? 'app' : 'web')); }
  function openWa(to, text) { var link = M.ownLink('mywa', to, text, howOpen()); if (PROG || /^https:/.test(link)) window.open(link, /^https:/.test(link) ? 'ns-wa' : '_blank'); else location.href = link; }
  function start(via) {
    var bad = C.unknownPlaces(via === 'mail' ? S.subj + '\n' + S.mail : S.wa); if (bad.length) { NS.toast(L('badplaces', { p: bad.join(' ') }), 'bad'); return; }
    var G = C.queue(chosen(), via); if (!G.go.length) { NS.toast(via === 'mail' ? L('nomails') : L('nomobs'), 'bad'); return; }
    var tag = via === 'mail' ? 'email' : 'wa', todo = G.go.filter(function (p) { var s = D.status[p.k]; return !(s && s[tag]); });
    D.hand = { via: via, ids: (todo.length ? todo : G.go).map(function (p) { return p.k; }), i: 0, done: 0, skipped: 0 }; saveList(); draw();
    var c = $('[data-bc-hand]', root); if (c) { c.scrollIntoView({ block: 'center', behavior: 'smooth' }); var b = $('[data-bcdo=open]', c); if (b) b.focus(); }
  }
  function cur() { if (!D || !D.hand) return null; var k = D.hand.ids[D.hand.i]; return full().filter(function (x) { return x.k === k; })[0] || null; }
  function texts(p) { var c = ctx(indexOf(p)); return { wa: C.fill(S.wa, p, c), subj: C.fill(S.subj, p, c), mail: C.fill(S.mail, p, c) }; }
  function mark(p, what) { var s = D.status[p.k] = D.status[p.k] || {}; s[what === 'email' ? 'email' : 'wa'] = 1; s.what = [s.wa && 'WhatsApp opened', s.email && 'email opened'].filter(Boolean).join(', '); s.at = Date.now(); }
  function step(what) {
    var p = cur(); if (!p) return; var T = texts(p), via = D.hand.via;
    if (what === 'open') {
      if (via === 'mail') location.href = C.mailto(p.email, T.subj, T.mail); else openWa(p.mobile, T.wa);
      mark(p, via === 'mail' ? 'email' : 'wa'); D.hand.done++;
    } else D.hand.skipped++;
    D.hand.i++;
    if (D.hand.i >= D.hand.ids.length) { var h = D.hand; D.hand = null; NS.toast(L('handdone', { a: h.done, b: h.skipped })); }
    saveList(); draw();
    var b = $('[data-bcdo=open]', root); if (b) b.focus();
  }
  function stopHand() { if (!D || !D.hand) return; var h = D.hand; D.hand = null; saveList(); NS.toast(L('stopped', { a: h.done })); draw(); }
  function copyText(s) { if (navigator.clipboard) navigator.clipboard.writeText(s).then(function () { NS.toast(L('copied')); }, function () {}); }

  /* ---------- the screen ---------- */
  function busy(m) { busyMsg = m || ''; var b = root && $('[data-bc-busy]', root); if (b) { b.hidden = !busyMsg; b.textContent = busyMsg; } }
  function sel(attr, opts, curv) { return '<select ' + attr + '>' + opts.map(function (o) { return '<option value="' + esc(o[0]) + '"' + (String(o[0]) === String(curv) ? ' selected' : '') + '>' + esc(o[1]) + '</option>'; }).join('') + '</select>'; }
  function inp(k, label, o) { o = o || {}; var b = bizRec(), ph = o.ph != null ? o.ph : (b[k] || ''); return '<label' + (o.wide ? ' class="bc-wide"' : '') + '>' + esc(label) + ' <input' + (o.type ? ' type="' + o.type + '"' : '') + (o.min != null ? ' min="' + o.min + '"' : '') + ' data-bcs="' + k + '" value="' + esc(S[k]) + '"' + (ph ? ' placeholder="' + esc(ph) + '"' : '') + '></label>'; }
  function step1() {
    return '<div class="card bc-card" data-bc-step="1"><h3>' + esc(L('s1')) + '</h3>' +
      '<div class="rb-drop bc-drop" data-bc-drop><b>' + esc(L('drop')) + ' <label class="rb-pick" for="bc-file">' + esc(L('choose')) + '</label></b><span>' + esc(L('dropwhat')) + '</span>' +
      '<input type="file" id="bc-file" class="sr" accept=".xlsx,.xls,.csv,.txt,.xml"></div>' +
      '<p class="rb-row"><button type="button" class="btn" data-bcdo="sample">' + esc(L('trysample')) + '</button> <button type="button" class="btn" data-bcdo="sampledl">' + esc(L('dlsample')) + '</button></p>' +
      '<p class="soft" data-bc-busy role="status"' + (busyMsg ? '' : ' hidden') + '>' + esc(busyMsg) + '</p>' +
      '<p class="soft">' + esc(L('private')) + '</p>' + (colsOpen && F ? colsHtml() : '') +
      '<h4>' + esc(L('lethd')) + '</h4><div class="rb-grid">' +
      '<label>' + esc(L('ason')) + ' <input type="date" data-bcs="asOn" value="' + esc(asOn()) + '"></label>' +
      '<label>' + esc(L('ldate')) + ' <input type="date" data-bcs="date" value="' + esc(ldate()) + '"></label>' +
      inp('days', L('days'), { type: 'number', min: 1, ph: '15' }) +
      '<label>' + esc(L('plain')) + ' ' + sel('data-bcs="plain"', [['dr', L('plain_dr')], ['cr', L('plain_cr')]], D ? D.plain : S.plain) + '</label>' +
      inp('prefix', L('prefix'), { ph: 'BC' }) + inp('start', L('start'), { type: 'number', min: 1, ph: '1' }) +
      '<label class="chk"><input type="checkbox" data-bcs="bills"' + (+S.bills ? ' checked' : '') + '> ' + esc(L('withbills')) + '</label>' +
      '<label class="chk"><input type="checkbox" data-bcs="noReply"' + (+S.noReply ? ' checked' : '') + '> ' + esc(L('noreply')) + '</label>' +
      '<label class="chk"><input type="checkbox" data-bcs="nil"' + (+S.nil ? ' checked' : '') + '> ' + esc(L('withnil')) + '</label>' +
      '</div><details class="bc-biz"' + (biz().name ? '' : ' open') + '><summary><b>' + esc(L('bizhd')) + '</b> <span class="soft">' + esc(biz().name || L('bizneed')) + '</span></summary><div class="rb-grid">' +
      inp('name', L('b_name')) + inp('addr', L('b_addr'), { wide: 1 }) + inp('city', L('b_city')) + inp('pin', L('b_pin')) + inp('gstin', 'GSTIN') + inp('phone', L('b_phone')) + inp('email', L('b_email')) + inp('signer', L('b_signer'), { ph: L('b_signer_ph') }) +
      '</div><p class="soft">' + esc(L('bizfrom')) + '</p></details></div>';
  }
  function colsHtml() {
    var rows = F.rows || [], hr = F.row || 0, heads = rows[hr] || [], opts = [['', L('notin')]].concat(heads.map(function (h, i) { return [String(i), (h || L('col') + ' ' + (i + 1))]; }));
    return '<div class="rb-cols"><h4>' + esc(L('colshd')) + '</h4><p class="soft">' + esc(F.name) + (F.sheet ? ', ' + esc(L('sheet')) + ' ' + esc(F.sheet) : '') + '. ' + esc(L('colsmem')) + '</p>' +
      '<p><label>' + esc(L('headrow')) + ' <input type="number" min="1" max="40" data-bc="headrow" value="' + (hr + 1) + '" style="width:5em"></label></p><div class="rb-grid">' +
      C.COLS.filter(function (c) { return !/^(days|due|billamt)$/.test(c.k); }).map(function (c) { return '<label>' + esc(c.k === 'amt' ? L('c_amt') : c.h) + (c.need ? ' *' : '') + ' ' + sel('data-bccol="' + c.k + '"', opts, c.k in (F.map || {}) ? F.map[c.k] : '') + '</label>'; }).join('') +
      '</div><div class="btns"><button type="button" class="btn primary" data-bcdo="cols">' + esc(L('readcols')) + '</button> <button type="button" class="btn" data-bcdo="colsoff">' + esc(L('close')) + '</button></div></div>';
  }
  function step2() {
    if (!D) return '';
    var all = full(), T = C.totals(all.filter(function (p) { return !(D.off && D.off[p.k]); })), nOff = all.length - chosen().length;
    var f = q ? all.filter(function (p) { var s = q.toLowerCase(); return p.name.toLowerCase().indexOf(s) >= 0 || (p.mobile || '').indexOf(q) >= 0 || (p.city || '').toLowerCase().indexOf(s) >= 0; }) : all;
    var h = '<div class="card bc-card" data-bc-step="2"><h3>' + esc(L('s2')) + '</h3>' + (D.sample ? '<p class="b1-warn">' + esc(L('samplewarn')) + '</p>' : '') +
      '<p><b>' + esc(D.file) + '</b> <span class="soft">(' + esc(D.kind) + ', ' + D.rows + ' ' + esc(L('rows')) + ')</span></p>' +
      '<div class="bc-tiles"><div><span>' + esc(L('t_letters')) + '</span><b>' + T.n + '</b></div><div><span>' + esc(L('t_dr', { n: T.ndr })) + '</span><b>₹' + C.inr(T.dr) + '</b></div><div><span>' + esc(L('t_cr', { n: T.ncr })) + '</span><b>₹' + C.inr(T.cr) + '</b></div></div>' +
      (nOff ? '<p class="soft">' + esc(L('offn', { n: nOff })) + '</p>' : '') + (D.fromBooks ? '<p class="soft">' + esc(L('frombooks', { n: D.fromBooks })) + '</p>' : '') +
      (D.plainRows ? '<p class="bc-note">' + esc(L('plainnote', { n: D.plainRows, s: D.plain === 'cr' ? L('plain_cr') : L('plain_dr') })) + '</p>' : '');
    var skip = (D.skipped || []).filter(function (s) { return !s.keep; }), warn = (D.skipped || []).filter(function (s) { return s.keep; });
    if (skip.length || warn.length || D.nil.length) h += '<details class="rb-skip"><summary>' + esc(L('leftout', { a: skip.length, b: D.nil.length, c: warn.length })) + '</summary><ul>' +
      skip.concat(warn).slice(0, 300).map(function (s) { return '<li>' + esc(L('row')) + ' ' + s.row + ': ' + esc(s.why) + '</li>'; }).join('') + D.nil.map(function (p) { return '<li>' + esc(p.name) + ': ' + esc(p.why) + '</li>'; }).join('') + '</ul></details>';
    h += '<p class="rb-row"><label>' + esc(L('find')) + ' <input type="search" data-bc="q" value="' + esc(q) + '" placeholder="' + esc(L('findph')) + '"></label> <button type="button" class="btn sm" data-bcdo="all">' + esc(L('tickall')) + '</button> <button type="button" class="btn sm" data-bcdo="nonew">' + esc(L('ticknone')) + '</button> <button type="button" class="btn sm" data-bcdo="colson">' + esc(L('chcols')) + '</button> <button type="button" class="btn sm" data-bcdo="clear">' + esc(L('another')) + '</button></p>' +
      '<div class="tbl-wrap"><table class="tbl bc-list"><thead><tr><th>' + esc(L('h_letter')) + '</th><th>' + esc(L('h_party')) + '</th><th class="n">' + esc(L('h_bal')) + '</th><th>' + esc(L('h_mob')) + '</th><th>' + esc(L('h_mail')) + '</th><th>' + esc(L('h_status')) + '</th><th></th></tr></thead><tbody>' +
      f.slice(0, shown).map(function (p) {
        var on = !(D.off && D.off[p.k]), st = D.status[p.k];
        return '<tr' + (on ? '' : ' class="bc-off"') + '><td><label class="rb-ok"><input type="checkbox" data-bcon="' + esc(p.k) + '"' + (on ? ' checked' : '') + ' aria-label="' + esc(L('make_for', { p: p.name })) + '"></label></td>' +
          '<th scope="row">' + esc(p.name) + ((p.city || p.addr) ? '<br><small class="soft">' + esc([p.addr, p.city].filter(Boolean).join(', ')) + '</small>' : '') + (p.bills.length ? '<br><small class="soft">' + esc(L('nbills', { n: p.bills.length })) + '</small>' : '') + '</th>' +
          '<td class="n">' + money(p.bal) + '</td>' +
          '<td><input type="tel" inputmode="tel" class="rb-mob" data-bcmob="' + esc(p.k) + '" value="' + esc(p.mobTyped) + '" aria-label="' + esc(L('mob_of', { p: p.name })) + '" placeholder="' + esc(L('h_mob')) + '"' + (p.mobTyped && !p.mobile ? ' aria-invalid="true"' : '') + '></td>' +
          '<td><input type="email" class="bc-mail" data-bcmail="' + esc(p.k) + '" value="' + esc(p.email) + '" aria-label="' + esc(L('mail_of', { p: p.name })) + '" placeholder="' + esc(L('h_mail')) + '"></td>' +
          '<td>' + (st && st.what ? esc(st.what) : '') + '</td><td><button type="button" class="btn sm" data-bcprev="' + esc(p.k) + '">' + esc(L('see')) + '</button></td></tr>';
      }).join('') + '</tbody></table></div>' +
      (f.length > shown ? '<p><button type="button" class="btn sm" data-bcdo="more">' + esc(L('more', { n: f.length - shown })) + '</button></p>' : '') + '</div>';
    return h;
  }
  function step3() {
    if (!D) return '';
    var list = chosen(), p = (pick && list.filter(function (x) { return x.k === pick; })[0]) || list[0];
    return '<div class="card bc-card" data-bc-step="3"><h3>' + esc(L('s3')) + '</h3>' +
      '<div class="btns"><button type="button" class="btn primary" data-bcdo="pdf"' + (list.length ? '' : ' disabled') + '>' + esc(L('b_pdf', { n: list.length })) + '</button> <button type="button" class="btn" data-bcdo="zip"' + (list.length ? '' : ' disabled') + '>' + esc(L('b_zip')) + '</button> <button type="button" class="btn" data-bcdo="print"' + (list.length ? '' : ' disabled') + '>' + esc(L('b_print')) + '</button> <button type="button" class="btn" data-bcdo="reg">' + esc(L('b_reg')) + '</button> <button type="button" class="btn" data-bcdo="regcsv">' + esc(L('b_regcsv')) + '</button></div>' +
      '<p class="soft">' + esc(L('regtip')) + '</p>' +
      (p ? '<p><b>' + esc(L('prevof', { p: p.name })) + '</b> <button type="button" class="btn sm" data-bcone="' + esc(p.k) + '">' + esc(L('thispdf')) + '</button></p><div class="tbl-wrap bc-prev">' + letter(p) + '</div>' : '') + '</div>';
  }
  function step4() {
    if (!D) return '';
    var list = chosen(), p = list[0], Gw = C.queue(list, 'wa'), Gm = C.queue(list, 'mail'), T = p ? texts(p) : null;
    return '<div class="card bc-card" data-bc-step="4"><h3>' + esc(L('s4')) + '</h3><div class="bc-two">' +
      '<div><label class="rb-tpl">' + esc(L('watxt')) + ' <textarea data-bcs="wa" rows="8">' + esc(S.wa) + '</textarea></label>' +
      (T ? '<div class="rb-prev"><b>' + esc(L('looks', { p: p.name })) + '</b><p class="msg-text">' + esc(T.wa) + '</p></div>' : '') + '</div>' +
      '<div><label class="rb-tpl">' + esc(L('subj')) + ' <input data-bcs="subj" value="' + esc(S.subj) + '"></label><label class="rb-tpl">' + esc(L('mailtxt')) + ' <textarea data-bcs="mail" rows="8">' + esc(S.mail) + '</textarea></label>' +
      (T ? '<div class="rb-prev"><b>' + esc(T.subj) + '</b><p class="msg-text">' + esc(T.mail) + '</p></div>' : '') + '</div></div>' +
      '<p class="soft">' + esc(L('places')) + ' ' + C.PLACES.map(function (x) { return '<code>{' + x + '}</code>'; }).join(' ') + ' <button type="button" class="btn sm" data-bcdo="deftxt">' + esc(L('deftxt')) + '</button></p>' +
      (phoneDev() ? '' : '<p><label>' + esc(L('openin')) + ' ' + sel('data-bcs="open"', [['web', L('o_web')], ['app', L('o_app')]], S.open) + '</label></p>') +
      '<div class="btns"><button type="button" class="btn primary" data-bcdo="startwa"' + (Gw.go.length ? '' : ' disabled') + '>' + esc(L('startwa', { n: Gw.go.length })) + '</button> <button type="button" class="btn" data-bcdo="startmail"' + (Gm.go.length ? '' : ' disabled') + '>' + esc(L('startmail', { n: Gm.go.length })) + '</button></div>' +
      (Gw.left.length || Gm.left.length ? '<details class="rb-skip"><summary>' + esc(L('nocontact', { a: Gw.left.length, b: Gm.left.length })) + '</summary><ul>' + Gw.left.map(function (x) { return '<li>' + esc(x.name) + ': ' + esc(x.why) + '</li>'; }).join('') + Gm.left.map(function (x) { return '<li>' + esc(x.name) + ': ' + esc(x.why) + '</li>'; }).join('') + '</ul></details>' : '') +
      '<p class="soft">' + esc(L('sendtip')) + '</p></div>';
  }
  function handHtml() {
    var p = cur(); if (!p) return '';
    var T = texts(p), mail = D.hand.via === 'mail';
    return '<div class="card bc-card rb-hand" data-bc-hand><h3>' + (D.hand.i + 1) + ' / ' + D.hand.ids.length + ': ' + esc(p.name) + '</h3>' +
      '<p><b>' + money(p.bal) + '</b> · <span class="soft">' + esc(mail ? p.email : p.mobile) + ' · ' + esc(ctx(indexOf(p)).ref) + '</span></p>' +
      (mail ? '<p><b>' + esc(T.subj) + '</b></p>' : '') + '<p class="msg-text bc-handtxt">' + esc(mail ? T.mail : T.wa) + '</p>' +
      '<div class="btns"><button type="button" class="btn primary" data-bcdo="open">' + esc(mail ? L('openmail') : L('openwa')) + '</button> <button type="button" class="btn" data-bcone="' + esc(p.k) + '">' + esc(L('thispdf')) + '</button> <button type="button" class="btn" data-bcdo="copy">' + esc(L('copy')) + '</button> <button type="button" class="btn" data-bcdo="skip">' + esc(L('skip')) + '</button> <button type="button" class="btn" data-bcdo="stop">' + esc(L('stop')) + '</button></div>' +
      '<p class="soft">' + esc(mail ? L('handmail') : L('handwa')) + '</p></div>';
  }
  function draw() {
    if (!root) return;
    var a = document.activeElement, kk = a && root.contains(a) && ['data-bcs', 'data-bc', 'data-bcmob', 'data-bcmail'].filter(function (x) { return a.hasAttribute(x); })[0], pos = kk && a.selectionStart;
    var keepSel = kk ? '[' + kk + '="' + a.getAttribute(kk) + '"]' : '';
    $('[data-bc-body]', root).innerHTML = (D && D.hand ? handHtml() : '') + step1() + step2() + step3() + step4();
    if (keepSel) { var e = $(keepSel, root); if (e) { e.focus(); try { if (pos != null) e.setSelectionRange(pos, pos); } catch (x) {} } }
  }
  var redrawT = 0; function later() { clearTimeout(redrawT); redrawT = setTimeout(draw, 400); }

  function wire() {
    root.addEventListener('click', function (e) {
      var pv = e.target.closest('[data-bcprev]'), one = e.target.closest('[data-bcone]'), b = e.target.closest('[data-bcdo]');
      if (pv) { pick = pv.getAttribute('data-bcprev'); draw(); var s3 = $('.bc-prev', root); if (s3) s3.scrollIntoView({ block: 'start', behavior: 'smooth' }); return; }
      if (one) { var p1 = full().filter(function (x) { return x.k === one.getAttribute('data-bcone'); })[0]; if (p1) onePdfOf(p1); return; }
      if (!b) return; var op = b.getAttribute('data-bcdo');
      if (op === 'sample') trySample();
      if (op === 'sampledl') download(L('samplefile') + '.xlsx', xlsxBlob([['Group Summary', C.sampleRows(), [30, 22, 16, 18, 14, 28, 14, 14]]]));
      if (op === 'cols') rereadCols();
      if (op === 'colson') { colsOpen = true; draw(); var c = $('.rb-cols', root); if (c) c.scrollIntoView({ block: 'center' }); }
      if (op === 'colsoff') { colsOpen = false; draw(); }
      if (op === 'clear') { if (!D || confirm(L('clearask'))) { D = null; F = null; saveList(); draw(); } }
      if (op === 'more') { shown += 200; draw(); }
      if (op === 'all') { D.off = {}; saveList(); draw(); }
      if (op === 'nonew') { D.off = {}; full().forEach(function (p) { D.off[p.k] = 1; }); saveList(); draw(); }
      if (op === 'pdf') onePdf();
      if (op === 'zip') zipPdfs();
      if (op === 'print') printAll();
      if (op === 'reg') register('xlsx');
      if (op === 'regcsv') register('csv');
      if (op === 'deftxt') { S.wa = C.WATPL; S.subj = C.SUBJ; S.mail = C.MAILTPL; save(); draw(); }
      if (op === 'startwa') start('wa');
      if (op === 'startmail') start('mail');
      if (op === 'open') step('open');
      if (op === 'skip') step('skip');
      if (op === 'stop') stopHand();
      if (op === 'copy') { var p2 = cur(); if (p2) { var T = texts(p2); copyText(D.hand.via === 'mail' ? T.subj + '\n\n' + T.mail : T.wa); } }
    });
    root.addEventListener('change', function (e) {
      var el = e.target, s = el.getAttribute('data-bcs');
      if (el.id === 'bc-file' && el.files && el.files[0]) { readFile(el.files[0]); el.value = ''; return; }
      if (el.hasAttribute('data-bcon')) { var k = el.getAttribute('data-bcon'); D.off = D.off || {}; if (el.checked) delete D.off[k]; else D.off[k] = 1; saveList(); draw(); return; }
      if (s === 'bills' || s === 'noReply' || s === 'nil') { S[s] = el.checked ? 1 : 0; save(); if (s === 'nil') reread(); else draw(); return; }
      if (s === 'plain') { S.plain = el.value; save(); if (D) { D.plain = el.value; if (F) F.plain = el.value; reread(); } else draw(); return; }
      if (s === 'asOn' || s === 'date' || s === 'open') { S[s] = el.value; save(); draw(); }
    });
    root.addEventListener('input', function (e) {
      var el = e.target, s = el.getAttribute('data-bcs');
      if (s && !/^(asOn|date|open|plain|bills|noReply|nil)$/.test(s)) { S[s] = el.value; save(); later(); return; }
      if (el.getAttribute('data-bc') === 'q') { q = el.value; shown = 100; later(); return; }
      var m = el.getAttribute('data-bcmob'); if (m != null && D) { D.mob = D.mob || {}; D.mob[m] = el.value; saveList(); later(); return; }
      var em = el.getAttribute('data-bcmail'); if (em != null && D) { D.mail = D.mail || {}; D.mail[em] = el.value; saveList(); later(); }
    });
    root.addEventListener('dragover', function (e) { var z = e.target.closest('[data-bc-drop]'); if (z) { e.preventDefault(); z.classList.add('on'); } });
    root.addEventListener('dragleave', function (e) { var z = e.target.closest('[data-bc-drop]'); if (z) z.classList.remove('on'); });
    root.addEventListener('drop', function (e) { var z = e.target.closest('[data-bc-drop]'); if (!z) return; e.preventDefault(); z.classList.remove('on'); var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; if (f) readFile(f); });
  }

  function init() {
    root = document.querySelector('[data-bconf]');
    if (!root) {   /* translated home pages: the section goes after the reminders (or after the khata itself) */
      var after = document.querySelector('[data-rembulk]') || document.querySelector('.tool-app'); if (!after) return;
      root = document.createElement('section'); root.className = 'wrap bconf'; root.setAttribute('data-bconf', ''); root.id = 'confirm';
      root.innerHTML = '<h2>' + esc(L('h2')) + '</h2><p class="soft">' + esc(L('lead')) + '</p><div data-bc-body></div>';
      after.parentNode.insertBefore(root, after.nextSibling);
    }
    if (!$('[data-bc-body]', root)) { var bdy = document.createElement('div'); bdy.setAttribute('data-bc-body', ''); root.appendChild(bdy); }
    wire(); draw();
    document.addEventListener('ns-data', function () { if (fillFromBooks()) draw(); else later(); });
    document.addEventListener('ns-lang', function () { draw(); });
    setTimeout(function () { if (D && fillFromBooks()) draw(); else draw(); }, 2500);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
  window.NSBConf = { draw: draw, read: readFile, sample: trySample, list: chosen };
})();

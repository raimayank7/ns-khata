/* NSRemind: payment reminders from an outstanding list the visitor already has (byof2-khata, 28 Nov 2026). Pure, no screen; khata.'s
   "Remind from your outstanding list" section (rembulk.js) and scripts/test_remcore.js use it.
   Reads: Tally "Bills Receivable", "Group Outstandings" and "Ledger Outstandings" (Excel export or XML export), Busy's outstanding
   report, and any Excel / CSV sheet that has a party name and an amount due (mobile, bill no., bill date, due date, overdue days are
   used when the sheet has them). One party = one reminder: bill-wise rows of the same party are added up, with their ageing.
   Columns are found with NSTpl.matchCols (shared heading matching); dates and amounts with NSTpl.date / NSTpl.amt.
   Designed & Developed by: Nivesh Sahayak (https://www.niveshsahayak.in/) */
(function (G) {
  'use strict';
  var Tpl = G.NSTpl || (typeof require === 'function' ? require('./tplcore.js') : null);
  function t(s) { return String(s == null ? '' : s).trim(); }
  function key(s) { return t(s).toLowerCase().replace(/[^a-z0-9ऀ-෿]+/g, ' ').trim(); }
  function r2(n) { return Math.round((+n + Number.EPSILON) * 100) / 100; }

  /* the columns an outstanding list may have (heading words of Tally, Busy, Marg, Vyapar and hand-made sheets) */
  var COLS0 = [
    { k: 'party', h: 'Party name', need: 1, a: ['party', 'partys name', "party's name", 'name', 'customer', 'customer name', 'debtor', 'debtor name', 'ledger', 'ledger name', 'account', 'account name', 'particulars', 'name of party', 'name of the party', 'buyer', 'client', 'client name'] },
    { k: 'mobile', h: 'Mobile', a: ['mobile no', 'mobile number', 'phone', 'phone no', 'phone number', 'contact', 'contact no', 'contact number', 'whatsapp', 'whatsapp no', 'whatsapp number', 'cell', 'mob', 'mob no', 'telephone', 'mobile/phone'] },
    { k: 'amt', h: 'Amount due', a: ['pending amount', 'pending', 'outstanding', 'outstanding amount', 'balance', 'balance amount', 'closing balance', 'due amount', 'amount', 'net balance', 'pending amt', 'bal amount', 'bal amt', 'os amount', 'amount pending', 'total due', 'due', 'receivable', 'amount receivable', 'pending bill amount', 'closing'] },
    { k: 'dr', h: 'Debit', a: ['dr', 'debit amount', 'debit balance', 'dr amount', 'closing debit'] },
    { k: 'cr', h: 'Credit', a: ['cr', 'credit amount', 'credit balance', 'cr amount', 'closing credit'] },
    { k: 'side', h: 'Dr or Cr', a: ['drcr', 'dr/cr', 'dr / cr', 'type', 'balance type'] },
    { k: 'bill', h: 'Bill no', a: ['ref no', 'ref. no.', 'ref', 'reference', 'bill number', 'bill ref', 'invoice no', 'invoice number', 'vch no', 'voucher no', 'bill', 'bill/ref', 'doc no', 'name of bill'] },
    { k: 'date', h: 'Bill date', a: ['date', 'invoice date', 'vch date', 'voucher date', 'bill dt', 'doc date', 'txn date'] },
    { k: 'billamt', h: 'Bill amount', a: ['opening amount', 'original amount', 'invoice amount', 'bill value', 'invoice value', 'opening amt'] },
    { k: 'due', h: 'Due date', a: ['due on', 'due dt', 'due by'] },
    { k: 'days', h: 'Overdue days', a: ['overdue by days', 'overdue by', 'overdue', 'days', 'age', 'age days', 'ageing', 'aging', 'days overdue', 'no of days', 'overdue days', 'pending days', 'due days', 'days pending'] },
    { k: 'email', h: 'Email', a: ['e-mail', 'email id', 'mail'] }
  ];
  var COLS = COLS0, TOTAL = /^(grand\s*)?(sub[\s-]*)?total\b|^carried over|^c\/o$/i;

  function days(a, b) { var x = Date.parse(a + 'T00:00:00Z'), y = Date.parse(b + 'T00:00:00Z'); return isFinite(x) && isFinite(y) ? Math.round((y - x) / 864e5) : null; }
  function today() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }

  /* one row's amount due: Debit - Credit columns, or the Amount column with Dr / Cr beside it (Cr = advance) */
  function rowAmt(o) {
    if (/\d/.test(o.dr || '') || /\d/.test(o.cr || '')) return { ok: true, v: r2(Math.abs(Tpl.amt(o.dr).v) - Math.abs(Tpl.amt(o.cr).v)) };
    if (!/\d/.test(o.amt || '')) return { ok: false, v: 0 };
    var a = Tpl.amt(o.amt), side = a.side || (/^\s*cr/i.test(t(o.side)) ? 'cr' : /^\s*dr/i.test(t(o.side)) ? 'dr' : '');
    return { ok: true, v: r2(side === 'cr' ? -Math.abs(a.v) : side === 'dr' ? Math.abs(a.v) : a.v) };
  }

  /* rows of one sheet -> { lines, skipped, map, row, single } ; opts { asOf, name (file name), layouts (remembered maps by heading signature) } */
  function readRows(rows, opts) {
    opts = opts || {}; rows = rows || [];
    var COLS = opts.cols || COLS0, L = null, s;   /* opts.cols (byof3-confirm): a longer column list; every line then also carries x = all its cells */
    for (var r = 0; r < Math.min(rows.length, 40) && opts.layouts; r++) { s = Tpl.sig(rows[r] || []); if (s && opts.layouts[s]) { L = { row: r, map: opts.layouts[s], heads: rows[r], remembered: true }; break; } }
    if (!L) L = Tpl.matchCols(rows.slice(0, 40), COLS);
    var noParty = false;
    if (!L || !('party' in L.map)) {
      /* Tally "Ledger Outstandings" of one party: no party column; the party is named above the headings ("Ledger: ABC Traders") */
      var L2 = Tpl.matchCols(rows.slice(0, 40), COLS.map(function (c) { return c.k === 'party' ? { k: 'party', h: 'Party name', a: c.a } : c.k === 'amt' ? Object.assign({ need: 1 }, c) : c; }));
      if (L2 && ('amt' in L2.map || 'dr' in L2.map) && ('bill' in L2.map || 'date' in L2.map)) { L = L2; noParty = true; }
      else return { err: 'No column with the party name and the amount due was found in the first 40 rows. Choose the columns below.', rows: rows };
    }
    if (!('amt' in L.map) && !('dr' in L.map) && !('cr' in L.map)) return { err: 'No column with the amount due was found. Choose the columns below.', rows: rows, map: L.map, row: L.row };
    var single = '';
    if (noParty) {
      for (var i = 0; i < L.row; i++) { var line = (rows[i] || []).filter(Boolean).join(' '), m = line.match(/^\s*ledger\s*:?\s*(.+)$/i); if (m) { single = t(m[1]); break; } }
      if (!single) single = t(String(opts.name || '').replace(/\.[a-z0-9]+$/i, '').replace(/[_-]+/g, ' ')) || 'Party';
    }
    var lines = [], skipped = [], asOf = opts.asOf || today(), last = '';
    for (var r2i = L.row + 1; r2i < rows.length; r2i++) {
      var row = rows[r2i] || [], o = { _r: r2i + 1 }, any = false;
      Object.keys(L.map).forEach(function (f) { var v = t(row[L.map[f]]); o[f] = v; if (v) any = true; });
      if (!any) continue;
      var party = noParty ? single : t(o.party);
      if (!noParty && !party && (o.bill || o.date) && last) party = last;   /* Busy / Tally bill-wise: the party name only on its first row */
      var first = row.map(t).filter(Boolean)[0] || '';
      if (TOTAL.test(party) || TOTAL.test(first)) { skipped.push({ row: o._r, why: 'total line' }); continue; }
      if (!party) { skipped.push({ row: o._r, why: 'no party name' }); continue; }
      var a = rowAmt(o);
      if (!a.ok) {
        /* a party heading row of a bill-wise report (name only, bills below) */
        if (!noParty && !o.bill && !o.date) { last = party; continue; }
        skipped.push({ row: o._r, why: 'no amount' }); continue;
      }
      if (!noParty && t(o.party)) last = party;
      var d = o.date ? Tpl.date(o.date) : '', du = o.due ? Tpl.date(o.due) : '', dy = o.days ? parseInt(String(o.days).replace(/[^\d-]/g, ''), 10) : NaN;
      if (o.date && !d) skipped.push({ row: o._r, why: 'bill date not read: ' + o.date, keep: 1 });
      var age = isFinite(dy) ? dy : d ? days(d, asOf) : du ? days(du, asOf) : null;
      lines.push({ row: o._r, party: party, mobile: t(o.mobile), email: t(o.email), bill: t(o.bill), date: d, due: du, amt: a.v, days: age == null ? null : Math.max(0, age), sheet: opts.sheet || '', x: opts.cols ? o : undefined });
    }
    return { lines: lines, skipped: skipped, map: L.map, row: L.row, heads: L.heads || rows[L.row], sig: Tpl.sig(rows[L.row] || []), single: single, remembered: !!L.remembered };
  }

  /* ---------- Tally XML exports of outstanding reports ----------
     Bills Receivable / Payable: repeated <BILLFIXED><BILLDATE/><BILLREF/><BILLPARTY/></BILLFIXED><BILLCL/><BILLDUE/><BILLOVERDUE/>
     Group Outstandings / Sundry Debtors: repeated <DSPACCNAME><DSPDISPNAME/></DSPACCNAME><DSPACCINFO>…<DSPCLDRAMTA/>…<DSPCLCRAMTA/>…</DSPACCINFO>
     Tally writes debit amounts as negative numbers in its XML: a receivable is negative there. */
  function untag(s) { return String(s || '').replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, function (a, n) { return String.fromCharCode(+n); }).replace(/&#x([0-9a-f]+);/gi, function (a, n) { return String.fromCharCode(parseInt(n, 16)); }).replace(/\r?\n/g, ' ').trim(); }
  function tag(s, name) { var m = new RegExp('<' + name + '\\b[^>]*>([\\s\\S]*?)</' + name + '>', 'i').exec(s); return m ? untag(m[1]) : ''; }
  function tallyDate(s) { s = t(s); if (/^\d{8}$/.test(s)) return s.slice(0, 4) + '-' + s.slice(4, 6) + '-' + s.slice(6, 8); return Tpl.date(s.replace(/-/g, ' ').replace(/\s+/g, '-')) || Tpl.date(s); }
  function readTallyXml(text, opts) {
    opts = opts || {}; var asOf = opts.asOf || today(), lines = [], skipped = [], s = String(text || '');
    if (/<BILLFIXED/i.test(s)) {
      var re = /<BILLFIXED\b[^>]*>([\s\S]*?)<\/BILLFIXED>([\s\S]*?)(?=<BILLFIXED\b|<\/ENVELOPE>|$)/gi, m, n = 0;
      while ((m = re.exec(s))) {
        n++;
        var party = tag(m[1], 'BILLPARTY'), cl = tag(m[2], 'BILLCL'), d = tallyDate(tag(m[1], 'BILLDATE')), du = tallyDate(tag(m[2], 'BILLDUE')), od = parseInt(tag(m[2], 'BILLOVERDUE'), 10);
        if (!party) { skipped.push({ row: n, why: 'no party name' }); continue; }
        var v = Tpl.amt(cl).v; if (!cl) { skipped.push({ row: n, why: 'no amount' }); continue; }
        var age = isFinite(od) ? od : d ? days(d, asOf) : null;
        lines.push({ row: n, party: party, mobile: '', bill: tag(m[1], 'BILLREF'), date: d, due: du, amt: r2(-v), days: age == null ? null : Math.max(0, age) });
      }
      return { lines: lines, skipped: skipped, kind: 'Tally bills (XML)', xmlSign: 1 };
    }
    if (/<DSPACCNAME/i.test(s)) {
      var re2 = /<DSPACCNAME\b[^>]*>([\s\S]*?)<\/DSPACCNAME>\s*<DSPACCINFO\b[^>]*>([\s\S]*?)<\/DSPACCINFO>/gi, m2, k = 0;
      while ((m2 = re2.exec(s))) {
        k++;
        var nm = tag(m2[1], 'DSPDISPNAME'), dr = tag(m2[2], 'DSPCLDRAMTA') || tag(m2[2], 'DSPCLDRAMT'), cr = tag(m2[2], 'DSPCLCRAMTA') || tag(m2[2], 'DSPCLCRAMT');
        if (!nm) { skipped.push({ row: k, why: 'no party name' }); continue; }
        if (TOTAL.test(nm)) { skipped.push({ row: k, why: 'total line' }); continue; }
        var due = r2(Math.abs(Tpl.amt(dr).v) - Math.abs(Tpl.amt(cr).v));
        lines.push({ row: k, party: nm, mobile: '', bill: '', date: '', due: '', amt: due, days: null });
      }
      return { lines: lines, skipped: skipped, kind: 'Tally outstandings (XML)' };
    }
    return { err: 'This XML is not a Tally outstanding report. In Tally open Display More Reports > Statements of Accounts > Outstandings > Receivables (or Group Outstandings), then Export > XML, or export to Excel.' };
  }

  /* ---------- mobile numbers of the parties: from NSBiz (lookup by name) or a second list ---------- */
  function mobile(s) { var d = String(s || '').replace(/\D/g, ''); if (d.length === 12 && d.slice(0, 2) === '91') d = d.slice(2); else if (d.length === 11 && d[0] === '0') d = d.slice(1); else if (d.length === 13 && d.slice(0, 3) === '091') d = d.slice(3); return /^[6-9]\d{9}$/.test(d) ? d : ''; }
  /* first Indian mobile found in a cell like "0161-2345678, 98140 12345" */
  function firstMobile(s) { var parts = String(s || '').split(/[,;\/|]+|\s{2,}/); for (var i = 0; i < parts.length; i++) { var m = mobile(parts[i]); if (m) return m; } return mobile(s); }

  /* lines -> parties (one reminder each) with ageing buckets; parties with nothing due are listed apart */
  var BUCKETS = [[0, 30, '0-30 days'], [31, 60, '31-60 days'], [61, 90, '61-90 days'], [91, 1e9, 'over 90 days']];
  function bucketOf(d) { if (d == null) return -1; for (var i = 0; i < BUCKETS.length; i++) if (d >= BUCKETS[i][0] && d <= BUCKETS[i][1]) return i; return 3; }
  function parties(lines) {
    var by = {}, order = [];
    (lines || []).forEach(function (l) {
      var k = key(l.party); if (!k) return;
      var p = by[k]; if (!p) { p = by[k] = { k: k, name: l.party, mobile: '', email: '', due: 0, bills: [], ages: [0, 0, 0, 0], nodate: 0, days: null, rows: [] }; order.push(k); }
      if (!p.mobile) p.mobile = firstMobile(l.mobile); if (!p.email && l.email) p.email = l.email;
      p.due = r2(p.due + l.amt); p.rows.push(l.row);
      if (l.amt > 0) { var b = bucketOf(l.days); if (b < 0) p.nodate = r2(p.nodate + l.amt); else p.ages[b] = r2(p.ages[b] + l.amt); if (l.days != null && (p.days == null || l.days > p.days)) p.days = l.days; }
      if (l.bill || l.date) p.bills.push({ no: l.bill, date: l.date, amt: l.amt, days: l.days });
    });
    var due = [], none = [];
    order.forEach(function (k) {
      var p = by[k];
      if (p.due <= 0.5) { none.push({ name: p.name, due: p.due, why: p.due < 0 ? 'advance (they have paid more)' : 'nothing due' }); return; }
      p.bills = p.bills.filter(function (b) { return b.amt > 0; }).sort(function (a, b) { return (b.days || 0) - (a.days || 0); });
      due.push(p);
    });
    return { due: due, none: none };
  }
  function bucketNames() { return BUCKETS.map(function (b) { return b[2]; }); }
  function totals(list) { var t0 = { n: list.length, due: 0, ages: [0, 0, 0, 0], nodate: 0 }; list.forEach(function (p) { t0.due = r2(t0.due + p.due); p.ages.forEach(function (v, i) { t0.ages[i] = r2(t0.ages[i] + v); }); t0.nodate = r2(t0.nodate + p.nodate); }); return t0; }

  /* ---------- the message ---------- */
  function inr(n) { var v = r2(n), neg = v < 0; v = Math.abs(v); var s = v.toFixed(2), i = s.split('.')[0], f = s.split('.')[1], last3 = i.slice(-3), rest = i.slice(0, -3); if (rest) last3 = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + last3; return (neg ? '-' : '') + last3 + (f === '00' ? '' : '.' + f); }
  function ddmmyyyy(d) { return d ? d.slice(8, 10) + '-' + d.slice(5, 7) + '-' + d.slice(0, 4) : ''; }
  var DEFTPL = 'Dear {name},\nThis is a friendly reminder that Rs {amount} is due to {business}.\nThe oldest bill is {days} days old.\nBills: {bills}\nPlease pay by UPI to {upi}\nor here: {link}\nIf you have already paid, please ignore this message.\nThank you,\n{business} {phone}';
  var PLACES = ['name', 'amount', 'days', 'bills', 'oldest', 'business', 'phone', 'upi', 'link', 'upilink'];
  /* the UPI payment address with the amount (NPCI UPI linking: pa, pn, am, cu, tn). Phones open it in a UPI app; WhatsApp shows it as text. */
  function upiLink(pa, pn, am, tn) {
    pa = t(pa); if (!/^[\w.\-]{2,256}@[a-zA-Z][\w.\-]{1,64}$/.test(pa)) return '';
    var q = 'pa=' + encodeURIComponent(pa) + (pn ? '&pn=' + encodeURIComponent(t(pn).slice(0, 50)) : '') + (am > 0 ? '&am=' + r2(am).toFixed(2) : '') + '&cu=INR' + (tn ? '&tn=' + encodeURIComponent(t(tn).slice(0, 50)) : '');
    return 'upi://pay?' + q;
  }
  function billText(p, max) {
    max = max || 5;
    var b = p.bills.filter(function (x) { return x.no || x.date; }), out = b.slice(0, max).map(function (x) { return (x.no || 'bill') + (x.date ? ' of ' + ddmmyyyy(x.date) : '') + ' (Rs ' + inr(x.amt) + ')'; });
    if (b.length > max) out.push('and ' + (b.length - max) + ' more');
    return out.join(', ');
  }
  /* the text for one party; a line whose place has nothing to fill (no UPI ID, no link, no bills) is left out; an empty name, business or phone is just left blank */
  var DROP = { days: 1, bills: 1, oldest: 1, upi: 1, link: 1, upilink: 1, amount: 1 };
  function message(tpl, p, ctx) {
    ctx = ctx || {};
    var oldest = p.bills.filter(function (x) { return x.date; }).map(function (x) { return x.date; }).sort()[0] || '';
    var V = { name: p.name, amount: inr(p.due), days: p.days == null ? '' : String(p.days), bills: billText(p), oldest: ddmmyyyy(oldest), business: t(ctx.business), phone: t(ctx.phone), upi: t(ctx.upi), link: t(ctx.link),
      upilink: upiLink(ctx.upi, ctx.business, p.due, 'Dues ' + p.name) };
    return String(tpl || DEFTPL).split('\n').map(function (line) {
      var empty = false, out = line.replace(/\{(\w+)\}/g, function (all, k) { k = k.toLowerCase(); if (!(k in V)) return all; if (!V[k] && DROP[k]) empty = true; return V[k]; });
      return empty ? null : out.replace(/[ \t]+$/, '');
    }).filter(function (l) { return l != null; }).join('\n').replace(/\n{3,}/g, '\n\n').trim();
  }
  function unknownPlaces(tpl) { var bad = []; String(tpl || '').replace(/\{(\w+)\}/g, function (a, k) { if (PLACES.indexOf(k.toLowerCase()) < 0 && bad.indexOf(a) < 0) bad.push(a); return a; }); return bad; }

  /* ---------- the queue: who gets a reminder and who is left out (and why) ----------
     opts { stop: { mobile10 or 'n:'+name key: 1 }, min (smallest amount), minDays (only dues at least this old) } */
  function queue(list, opts) {
    opts = opts || {}; var go = [], left = [], seen = {};
    (list || []).forEach(function (p) {
      var why = '';
      if (opts.stop && (opts.stop['n:' + p.k] || (p.mobile && opts.stop[p.mobile]))) why = 'on your do-not-send list';
      else if (opts.min && p.due < opts.min) why = 'less than Rs ' + inr(opts.min) + ' due';
      else if (opts.minDays && (p.days == null || p.days < opts.minDays)) why = p.days == null ? 'age of the dues not known' : 'dues only ' + p.days + ' days old';
      else if (!p.mobile) why = 'no mobile number';
      else if (seen[p.mobile]) why = 'same mobile as ' + seen[p.mobile];
      if (why) left.push({ k: p.k, name: p.name, mobile: p.mobile, due: p.due, why: why }); else { seen[p.mobile] = p.name; go.push(p); }
    });
    return { go: go, left: left };
  }

  /* ---------- the sent / skipped report ---------- */
  function reportRows(list, status) {
    var out = [['Party', 'Mobile', 'Amount due (Rs)', 'Oldest bill (days)', '0-30 days', '31-60 days', '61-90 days', 'Over 90 days', 'Bills', 'What happened', 'When']];
    (list || []).forEach(function (p) {
      var s = status[p.k] || {};
      out.push([p.name, p.mobile || '', r2(p.due), p.days == null ? '' : p.days, p.ages[0], p.ages[1], p.ages[2], p.ages[3], billText(p, 50), s.what || 'not sent', s.at ? new Date(s.at).toLocaleString('en-IN') : '']);
    });
    return out;
  }

  /* ---------- the sample file (made-up parties, clearly marked) in Tally's Bills Receivable Excel layout ---------- */
  function sampleRows(asOf) {
    asOf = asOf || today();
    function back(n) { var d = new Date(Date.parse(asOf + 'T00:00:00Z') - n * 864e5); return ('0' + d.getUTCDate()).slice(-2) + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + d.getUTCFullYear(); }
    return [
      ['SAMPLE TRADERS (made-up sample data)'], ['Bills Receivable'], ['1-Apr-2026 to ' + ddmmyyyy(asOf)], [],
      ['Date', 'Ref. No.', "Party's Name", 'Mobile', 'Pending Amount', 'Due on', 'Overdue by days'],
      [back(12), 'SI/26-27/101', 'Gill Hardware (sample)', '9876500001', '12,400.00 Dr', back(0), '0'],
      [back(48), 'SI/26-27/063', 'Gill Hardware (sample)', '9876500001', '8,250.00 Dr', back(33), '33'],
      [back(75), 'SI/26-27/041', 'Mehta Paints (sample)', '98765 00002', '21,000.00 Dr', back(60), '60'],
      [back(104), 'SI/26-27/012', 'Sandhu Builders (sample)', '+91 98765 00003', '1,05,500.00 Dr', back(89), '89'],
      [back(20), 'SI/26-27/118', 'Sandhu Builders (sample)', '+91 98765 00003', '18,900.00 Dr', back(5), '5'],
      [back(9), 'SI/26-27/124', 'Arora Kirana (sample)', '', '3,450.00 Dr', back(0), '0'],
      [back(40), 'On Account', 'Kapoor Electricals (sample)', '9876500005', '2,000.00 Cr', '', ''],
      [back(150), 'SI/25-26/488', 'Verma Steels (sample)', '9876500006', '46,780.00 Dr', back(135), '135'],
      ['', '', 'Grand Total', '', '2,13,280.00 Dr', '', '']
    ];
  }

  var API = { COLS: COLS, readRows: readRows, readTallyXml: readTallyXml, parties: parties, totals: totals, bucketNames: bucketNames, message: message, DEFTPL: DEFTPL, PLACES: PLACES, unknownPlaces: unknownPlaces,
    upiLink: upiLink, billText: billText, queue: queue, reportRows: reportRows, sampleRows: sampleRows, mobile: mobile, firstMobile: firstMobile, inr: inr, key: key, today: today };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else G.NSRemind = API;
})(typeof window !== 'undefined' ? window : this);

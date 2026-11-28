/* NSConfirm: balance confirmation letters in bulk from an outstanding list (byof3-confirm, 28 Nov 2026). Pure, no screen;
   khata.'s "Balance confirmation letters for every party" section (bconf.js) and scripts/test_bconf.js use it.
   Reads the lists NSRemind (remcore.js) reads: Tally Bills Receivable / Payable, Group Outstandings, Group Summary of Sundry
   Debtors or Creditors (Excel or XML), Busy's outstanding report and any sheet with a party and a balance. Both sides are kept:
   a debit balance = the party owes us, a credit balance = we owe the party (suppliers, advances).
   The letter has the same words and layout as the party ledger's letter in engine/ledger.js (letterSheets), so a letter made
   here and one made from the NSBiz books look the same; this one can list the open bills and carries a reference number.
   Designed & Developed by: Nivesh Sahayak (https://www.niveshsahayak.in/) */
(function (G) {
  'use strict';
  var R = G.NSRemind || (typeof require === 'function' ? require('./remcore.js') : null);
  var Tpl = G.NSTpl || (typeof require === 'function' ? require('./tplcore.js') : null);
  function t(s) { return String(s == null ? '' : s).trim(); }
  function r2(n) { return Math.round((+n + Number.EPSILON) * 100) / 100; }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

  /* the outstanding columns plus the details a letter needs (address, city, PIN, GSTIN) */
  var COLS = R.COLS.concat([
    { k: 'addr', h: 'Address', a: ['address', 'address 1', 'address line 1', 'billing address', 'party address', 'addr', 'mailing address', 'full address'] },
    { k: 'city', h: 'City', a: ['town', 'place', 'city name', 'station', 'district'] },
    { k: 'pin', h: 'PIN code', a: ['pin', 'pincode', 'pin code', 'postal code', 'zip'] },
    { k: 'gstin', h: 'GSTIN', a: ['gstin', 'gst no', 'gst number', 'gstin/uin', 'gstin uin', 'gst', 'gst in'] },
    { k: 'contact', h: 'Contact person', a: ['contact person', 'attention', 'kind attention', 'person'] }
  ]);

  /* Tally Group Summary / Trial Balance exported to Excel have two heading rows: "Particulars | Closing Balance" and below it
     "Debit | Credit" (sometimes Opening, Transactions and Closing, each with Debit and Credit). The two rows are joined into one,
     taking the Debit and Credit of the Closing group (or the last pair) */
  function joinHeads(rows) {
    rows = (rows || []).slice();
    for (var r = 0; r < Math.min(rows.length - 1, 20); r++) {
      var up = rows[r] || [], dn = rows[r + 1] || [], kinds = dn.map(function (c) { var x = Tpl.hk(c); return x === 'debit' || x === 'dr' ? 'dr' : x === 'credit' || x === 'cr' ? 'cr' : ''; });
      if (kinds.indexOf('dr') < 0 || kinds.indexOf('cr') < 0) continue;
      if (up.map(Tpl.hk).some(function (x) { return x === 'debit' || x === 'credit'; })) continue;
      var grp = [], g = '';
      for (var i = 0; i < Math.max(up.length, dn.length); i++) { if (t(up[i])) g = Tpl.hk(up[i]); grp[i] = g; }
      var closing = kinds.some(function (k, i) { return k && /closing/.test(grp[i]); });
      var lastDr = kinds.lastIndexOf('dr'), lastCr = kinds.lastIndexOf('cr'), joined = [];
      for (var j = 0; j < Math.max(up.length, dn.length); j++) {
        var k = kinds[j];
        if (k) joined[j] = (closing ? /closing/.test(grp[j]) : (j === lastDr || j === lastCr)) ? (k === 'dr' ? 'Debit' : 'Credit') : t(up[j] || grp[j]) + ' ' + t(dn[j]) + ' (not used)';
        else joined[j] = t(dn[j]) || t(up[j]);
      }
      rows[r] = []; rows[r + 1] = joined;
      return { rows: rows, joined: r + 2 };
    }
    return { rows: rows, joined: 0 };
  }

  /* did the row say which side the amount is on (Debit / Credit columns, "Dr" / "Cr" beside the number or a Dr/Cr column)? */
  function sided(x) { x = x || {}; return /\d/.test(x.dr || '') || /\d/.test(x.cr || '') || /\b(dr|cr)\.?\s*$/i.test(t(x.amt)) || /^\s*-/.test(t(x.amt)) || /^(dr|cr)/i.test(t(x.side)); }

  /* rows of a sheet -> { lines, skipped, map, row, ... } ; opts as NSRemind.readRows plus plain: 'dr' (default) or 'cr' = what a
     plain amount (no Dr/Cr anywhere) means: owed to us or owed by us */
  function readRows(rows, opts) {
    opts = opts || {};
    var J = joinHeads(rows), r = R.readRows(J.rows, Object.assign({}, opts, { cols: COLS }));
    if (r.err) return r;
    if (opts.plain === 'cr') r.lines.forEach(function (l) { if (!sided(l.x) && l.amt > 0) l.amt = -l.amt; });
    r.plainRows = r.lines.filter(function (l) { return !sided(l.x); }).length;
    r.joinedHeads = J.joined;
    return r;
  }
  function readTallyXml(text, opts) { return R.readTallyXml(text, opts); }

  /* lines -> one party each, with its balance (+ = Dr, they owe us; - = Cr, we owe them), open bills and contact details */
  function parties(lines, opts) {
    opts = opts || {}; var by = {}, order = [];
    (lines || []).forEach(function (l) {
      var k = R.key(l.party); if (!k) return;
      var p = by[k]; if (!p) { p = by[k] = { k: k, name: l.party, mobile: '', email: '', addr: '', city: '', pin: '', gstin: '', contact: '', bal: 0, bills: [], rows: [] }; order.push(k); }
      var x = l.x || {};
      if (!p.mobile) p.mobile = R.firstMobile(l.mobile);
      if (!p.email) p.email = email(l.email || x.email);
      ['addr', 'city', 'pin', 'contact'].forEach(function (f) { if (!p[f] && t(x[f])) p[f] = t(x[f]); });
      if (!p.gstin && /^\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/.test(t(x.gstin).toUpperCase())) p.gstin = t(x.gstin).toUpperCase();
      p.bal = r2(p.bal + l.amt); p.rows.push(l.row);
      if (l.bill || l.date) p.bills.push({ no: l.bill, date: l.date, due: l.due, amt: l.amt });
    });
    var list = [], nil = [];
    order.forEach(function (k) {
      var p = by[k];
      p.bills.sort(function (a, b) { return (a.date || '9').localeCompare(b.date || '9'); });
      if (Math.abs(p.bal) < 0.005) { p.bal = 0; if (!opts.nil) { nil.push({ name: p.name, why: 'nil balance' }); return; } }
      list.push(p);
    });
    return { list: list, nil: nil };
  }
  function email(s) { var m = String(s || '').match(/[\w.+\-]+@[\w\-]+(\.[\w\-]+)+/); return m ? m[0] : ''; }
  function totals(list) { var o = { n: list.length, dr: 0, cr: 0, ndr: 0, ncr: 0, nil: 0 }; list.forEach(function (p) { if (p.bal > 0) { o.dr = r2(o.dr + p.bal); o.ndr++; } else if (p.bal < 0) { o.cr = r2(o.cr - p.bal); o.ncr++; } else o.nil++; }); return o; }

  /* ---------- dates, numbers, reference numbers ---------- */
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  function ldate(d) { d = t(d); if (!/^\d{4}-\d\d-\d\d$/.test(d)) return d; return d.slice(8, 10) + ' ' + MON[+d.slice(5, 7) - 1] + ' ' + d.slice(0, 4); }
  function addDays(d, n) { var x = new Date(Date.parse(d + 'T00:00:00Z') + n * 864e5); return x.toISOString().slice(0, 10); }
  /* the usual "as on" date: the last 31 March from April to September (year-end confirmations for the audit), else today */
  function defAsOn(today) { today = today || R.today(); var y = +today.slice(0, 4), m = +today.slice(5, 7); return m >= 4 && m <= 9 ? y + '-03-31' : today; }
  function fyShort(d) { var y = +d.slice(0, 4), m = +d.slice(5, 7), s = m >= 4 ? y : y - 1; return String(s).slice(2) + '-' + String(s + 1).slice(2); }
  function refNo(prefix, asOn, n) { return t(prefix || 'BC') + '/' + fyShort(asOn) + '/' + ('00' + n).slice(-3); }
  function inr(n) { return R.inr(n); }
  function amtText(v) { return '₹' + inr(Math.abs(v)); }
  function side(v) { return v > 0.004 ? 'Dr' : v < -0.004 ? 'Cr' : ''; }

  /* ---------- the letter (words and classes of ledger.js letterSheets; stsheet / lg-* styles from app.css) ---------- */
  function lines(o) { var a = [o.addr, [o.city, o.pin].filter(t).join(' - ')].map(t).filter(Boolean), out = a.map(esc).join('<br>'); if (o.gstin) out += (out ? '<br>' : '') + 'GSTIN: ' + esc(o.gstin); return out ? '<br>' + out : ''; }
  function letterHtml(p, ctx) {
    ctx = ctx || {}; var b = ctx.biz || {}, biz = esc(t(b.name) || 'Our business'), asOn = ctx.asOn, bal = p.bal, amt = amtText(bal);
    var what = bal > 0.004 ? amt + ' Debit, that is, ' + amt + ' is payable by you to us' : bal < -0.004 ? amt + ' Credit, that is, ' + amt + ' is payable by us to you' : 'Nil (no amount is payable either way)';
    var theirs = bal > 0.004 ? amt + ' payable by us to ' + biz : bal < -0.004 ? amt + ' receivable by us from ' + biz : 'Nil';
    var days = Math.max(1, +ctx.days || 15), by = ctx.date ? ldate(addDays(ctx.date, days)) : '';
    var bills = (p.bills || []).filter(function (x) { return x.no || x.date; }), maxB = ctx.maxBills == null ? 25 : ctx.maxBills;
    var billTbl = ctx.bills !== false && bills.length ? '<p>The balance is made up of these bills:</p><table class="tbl st bc-bills"><thead><tr><th>Date</th><th>Bill / reference</th><th class="n">Amount</th></tr></thead><tbody>' +
      bills.slice(0, maxB).map(function (x) { return '<tr><td>' + esc(ldate(x.date)) + '</td><td>' + esc(x.no || '') + '</td><td class="n">' + inr(Math.abs(x.amt)) + ' ' + side(x.amt) + '</td></tr>'; }).join('') +
      (bills.length > maxB ? '<tr><td colspan="3">and ' + (bills.length - maxB) + ' more bills (statement of account on request)</td></tr>' : '') + '</tbody></table>' : '';
    var contact = [t(b.phone) && 'Phone ' + esc(b.phone), t(b.email) && esc(b.email)].filter(Boolean).join(' · ');
    return '<div class="stsheet sh lg-letter bc-letter" dir="ltr" data-sheet>' +
      '<div class="st-hd"><div><b>' + biz + '</b>' + lines({ addr: b.addr, city: b.city, pin: b.pin, gstin: b.gstin }) + (contact ? '<br>' + contact : '') + '</div><div class="r">' + (ctx.ref ? 'Ref: ' + esc(ctx.ref) + '<br>' : '') + 'Date: ' + esc(ldate(ctx.date)) + '</div></div>' +
      '<p>To,<br><b>' + esc(p.name) + '</b>' + lines(p) + (p.contact ? '<br>Kind attention: ' + esc(p.contact) : '') + '</p>' +
      '<p><b>Subject: Confirmation of balance as on ' + esc(ldate(asOn)) + '</b></p>' +
      '<p>Dear Sir / Madam,</p>' +
      '<p>As per our books of account, the balance of your account with us as on ' + esc(ldate(asOn)) + ' is <b>' + what + '</b>.</p>' + billTbl +
      '<p>Please check this balance with your books and confirm it by signing the confirmation below and sending a copy back to us within ' + days + ' days' + (by ? ' (by ' + by + ')' : '') + '. ' +
      'If your books show a different balance, please send us your statement of account for the period, so that the difference can be reconciled.' +
      (ctx.noReply ? ' If we do not hear from you by then, we shall take the balance as confirmed.' : '') + '</p>' +
      '<p>Thanking you,<br>Yours faithfully,</p><p class="lg-sign">For <b>' + biz + '</b><br><br><br>' + (t(ctx.signer) ? esc(ctx.signer) + '<br>' : '') + 'Authorised signatory</p>' +
      '<div class="lg-cut">cut here and send back</div>' +
      '<p><b>CONFIRMATION</b>' + (ctx.ref ? ' <span class="soft">(your ref: ' + esc(ctx.ref) + ')</span>' : '') + '</p><p>To <b>' + biz + '</b></p>' +
      '<p>We confirm the balance of our account in your books as on ' + esc(ldate(asOn)) + ': <b>' + theirs + '</b>.</p>' +
      '<p class="lg-box">&#9744; We agree with this balance.</p>' +
      '<p class="lg-box">&#9744; We do not agree. As per our books the balance is &#8377; ______________ (payable by us / receivable by us). Our statement of account is attached.</p>' +
      '<p class="lg-sign2">For <b>' + esc(p.name) + '</b><br><br><br>Signature and stamp: ____________________ &nbsp; Name: ____________________ &nbsp; Date: ____________</p>' +
      (ctx.credit || '') + '</div>';
  }

  /* ---------- WhatsApp and email texts ---------- */
  var WATPL = 'Dear {name},\nAs per our books, the balance of your account as on {asof} is Rs {amount} {drcr}: {meaning}.\nPlease check it with your books and reply "Confirmed", or send us your statement of account if it differs, by {replyby}.\nThe confirmation letter is attached (ref {ref}).\nThank you,\n{business} {phone}';
  var SUBJ = 'Confirmation of balance as on {asof} - {business}';
  var MAILTPL = 'Dear Sir / Madam,\n\nAs per our books of account, the balance of your account ({name}) with us as on {asof} is Rs {amount} {drcr}, that is, {meaning}.\n\nPlease check this balance with your books and confirm it by replying to this email, or by signing the confirmation in the attached letter (ref {ref}), by {replyby}. If your books show a different balance, please send us your statement of account for the period, so that the difference can be reconciled.\n\nThanking you,\n{signer}\n{business}\n{phone}';
  var PLACES = ['name', 'amount', 'drcr', 'meaning', 'asof', 'replyby', 'ref', 'business', 'phone', 'signer'];
  function vals(p, ctx) {
    ctx = ctx || {}; var b = ctx.biz || {}, bal = p.bal;
    return { name: p.name, amount: inr(Math.abs(bal)), drcr: side(bal) || '(nil)', meaning: bal > 0.004 ? 'payable by you to us' : bal < -0.004 ? 'payable by us to you' : 'nothing is payable either way',
      asof: ldate(ctx.asOn), replyby: ldate(addDays(ctx.date || ctx.asOn, Math.max(1, +ctx.days || 15))), ref: ctx.ref || '', business: t(b.name), phone: t(b.phone), signer: t(ctx.signer) };
  }
  function fill(tpl, p, ctx) {
    var V = vals(p, ctx);
    return String(tpl || '').split('\n').map(function (line) {
      var empty = false, out = line.replace(/\{(\w+)\}/g, function (all, k) { k = k.toLowerCase(); if (!(k in V)) return all; if (!V[k] && (k === 'ref' || k === 'signer')) empty = true; return V[k]; });
      return empty ? null : out.replace(/[ \t]+$/, '');
    }).filter(function (l) { return l != null; }).join('\n').replace(/\n{3,}/g, '\n\n').trim();
  }
  function unknownPlaces(tpl) { var bad = []; String(tpl || '').replace(/\{(\w+)\}/g, function (a, k) { if (PLACES.indexOf(k.toLowerCase()) < 0 && bad.indexOf(a) < 0) bad.push(a); return a; }); return bad; }
  function mailto(to, subject, body) { return 'mailto:' + encodeURIComponent(t(to)).replace(/%40/g, '@') + '?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(body).slice(0, 1800); }
  function fileName(p, ref) { return ('Balance confirmation ' + p.name + (ref ? ' ' + ref.replace(/\//g, '-') : '')).replace(/[\\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 90); }

  /* ---------- the register: every letter, its texts and two empty columns for the replies ---------- */
  function registerRows(list, ctx, status) {
    status = status || {};
    var out = [['Ref no.', 'Party', 'Address', 'GSTIN', 'Mobile', 'Email', 'Balance as on ' + ldate(ctx.asOn) + ' (Rs)', 'Dr/Cr', 'Open bills', 'Sent', 'When', 'Reply received on', 'Their balance (Rs)', 'Difference / remarks', 'WhatsApp text', 'Email subject', 'Email text']];
    (list || []).forEach(function (p, i) {
      var c = Object.assign({}, ctx, { ref: refNo(ctx.prefix, ctx.asOn, (+ctx.start || 1) + i) }), s = status[p.k] || {};
      out.push([c.ref, p.name, [p.addr, p.city, p.pin].filter(t).join(', '), p.gstin, p.mobile || '', p.email || '', r2(Math.abs(p.bal)), side(p.bal), p.bills.length || '', s.what || '', s.at ? new Date(s.at).toLocaleString('en-IN') : '', '', '', '',
        fill(ctx.wa || WATPL, p, c), fill(ctx.subj || SUBJ, p, c), fill(ctx.mail || MAILTPL, p, c)]);
    });
    return out;
  }
  /* who gets a WhatsApp or an email, and who has neither */
  function queue(list, via) {
    var go = [], left = [];
    (list || []).forEach(function (p) { var ok = via === 'mail' ? !!p.email : !!R.mobile(p.mobile); if (ok) go.push(p); else left.push({ k: p.k, name: p.name, why: via === 'mail' ? 'no email' : 'no mobile number' }); });
    return { go: go, left: left };
  }

  /* ---------- the sample list (made-up parties, both sides) in the Tally Group Summary layout with two heading rows ---------- */
  function sampleRows() {
    return [
      ['SAMPLE TRADERS (made-up sample data)'], ['Sundry Debtors and Sundry Creditors'], ['1-Apr-2025 to 31-Mar-2026'], [],
      ['Particulars', 'Address', 'City', 'GSTIN', 'Mobile', 'Email', 'Closing Balance', ''],
      ['', '', '', '', '', '', 'Debit', 'Credit'],
      ['Gill Hardware (sample)', 'Shop 12, Gill Road', 'Ludhiana 141003', '03AABCG1234K1Z2', '9876500001', 'accounts@gill-sample.in', '20,650.00', ''],
      ['Mehta Paints (sample)', '45 Mall Road', 'Amritsar', '', '98765 00002', '', '21,000.00', ''],
      ['Sandhu Builders (sample)', 'Plot 7, Phase 2', 'Mohali', '03AAPFS5678L1ZG', '+91 98765 00003', 'sandhu.builders@example.com', '1,24,400.00', ''],
      ['Kapoor Electricals (sample)', '', 'Jalandhar', '', '9876500005', '', '', '2,000.00'],
      ['Bharat Cement Supply (sample)', 'G.T. Road', 'Khanna', '', '9876500007', 'billing@bharat-sample.in', '', '86,300.00'],
      ['Verma Steels (sample)', 'Industrial Area B', 'Ludhiana', '', '9876500006', '', '46,780.00', ''],
      ['Grand Total', '', '', '', '', '', '2,12,830.00', '88,300.00']
    ];
  }

  var API = { COLS: COLS, joinHeads: joinHeads, readRows: readRows, readTallyXml: readTallyXml, parties: parties, totals: totals, letterHtml: letterHtml,
    WATPL: WATPL, SUBJ: SUBJ, MAILTPL: MAILTPL, PLACES: PLACES, fill: fill, vals: vals, unknownPlaces: unknownPlaces, mailto: mailto, fileName: fileName,
    registerRows: registerRows, queue: queue, sampleRows: sampleRows, refNo: refNo, defAsOn: defAsOn, ldate: ldate, addDays: addDays, side: side, inr: inr, sided: sided };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else G.NSConfirm = API;
})(typeof window !== 'undefined' ? window : this);

/* SS builder core — parse PO follow-up workbook, index lines, build SS rows */
var SSCore = (function () {
  var norm = function (v) { return String(v == null ? '' : v).replace(/ /g, ' ').replace(/\s+/g, ' ').trim(); };
  var key = function (v) { return norm(v).toLowerCase(); };
  var ALIASES = {
    po: ['po no.', 'po number', 'po no'], poscaf: ['po scaf'], supplier: ['supplier'],
    article: ['article supplier'], rmcode: ['rm code'], rmname: ['rm name'], color: ['color'],
    scc: ['supplier color code'], sizecup: ['size cup'], tol: ['% tolerance'],
    customer: ['customer code', 'customer name'], story: ['story'], qty: ['order quantity'], unit: ['unit'],
    bu: ['bu name'], purchaser: ['purchaser'], inv: ['inv_no'], invdate: ['inv_date'], qtyinv: ['qty_inv'],
    etd: ['confirmed etd', 'confirm_etd'], issued: ['issued date']
  };
  var isNum = function (v) { return typeof v === 'number' && isFinite(v); };
  var isBlank = function (v) { var s = key(v); return s === '' || s === '(blank)' || s === '#n/a' || s === '#value!' || s === '0'; };

  function findHeader(rows) {
    for (var r = 0; r < Math.min(rows.length, 8); r++) {
      var row = rows[r] || [], map = {}, hit = false;
      for (var c = 0; c < row.length; c++) {
        var h = key(row[c]);
        if (!h) continue;
        for (var f in ALIASES) if (map[f] == null && ALIASES[f].indexOf(h) >= 0) map[f] = c;
      }
      if (map.po != null && map.supplier != null && (map.color != null || map.article != null)) { hit = true; }
      if (hit) return { row: r, map: map };
    }
    return null;
  }

  // Repair the two known column-shift patterns in hand-pasted sheets (e.g. SHIPPED)
  function repair(rec, raw, m) {
    if (isNum(rec.qty) || rec.qty === '' ) return rec;
    var sc = m.sizecup != null ? raw[m.sizecup] : null, tol = m.tol != null ? raw[m.tol] : null;
    if (isNum(sc) && typeof tol === 'string' && tol.length <= 8) {
      rec.qty = sc; rec.unit = norm(tol); rec.fixed = 'A';
    } else if (typeof sc === 'string' && /^\d+-[A-Z]/i.test(sc)) {
      rec.rmname = norm(m.rmcode != null ? raw[m.rmcode] : '');
      rec.color = norm(m.rmname != null ? raw[m.rmname] : '');
      rec.qty = isNum(rec.sccRaw) ? rec.sccRaw : '';
      rec.scc = ''; rec.unit = norm(sc).replace(/^\d+-/, ''); rec.fixed = 'B';
    } else { rec.qty = ''; rec.unit = ''; rec.fixed = '?'; }
    return rec;
  }

  function parseWorkbook(XLSX, wb) {
    var sources = [], lines = [];
    wb.SheetNames.forEach(function (name) {
      var ws = wb.Sheets[name];
      if (!ws || !ws['!ref'] || /pivot/i.test(name)) return;
      var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
      var h = findHeader(rows);
      if (!h) return;
      var m = h.map, n = 0;
      for (var r = h.row + 1; r < rows.length; r++) {
        var raw = rows[r]; if (!raw) continue;
        var g = function (f) { return m[f] != null ? raw[m[f]] : null; };
        var po = norm(g('po'));
        if (isBlank(po) || /^\d+$/.test(po) && po.length < 3) continue;
        var rec = {
          sheet: name, row: r + 1, po: po, poscaf: norm(g('poscaf')), supplier: norm(g('supplier')),
          article: norm(g('article')), rmcode: norm(g('rmcode')), rmname: norm(g('rmname')), color: norm(g('color')),
          scc: isNum(g('scc')) ? '' : norm(g('scc')), sccRaw: g('scc'), customer: norm(g('customer')), story: norm(g('story')),
          qty: g('qty') == null ? '' : g('qty'), unit: norm(g('unit')), bu: norm(g('bu')), purchaser: norm(g('purchaser')),
          inv: isBlank(g('inv')) ? '' : norm(g('inv')), invdate: g('invdate'), qtyinv: isNum(g('qtyinv')) ? g('qtyinv') : '',
          etd: g('etd'), issued: g('issued'), fixed: ''
        };
        ['bu', 'customer', 'story', 'scc', 'color', 'article', 'rmname'].forEach(function (f) { if (isBlank(rec[f])) rec[f] = ''; });
        repair(rec, raw, m);
        delete rec.sccRaw;
        lines.push(rec); n++;
      }
      if (n) sources.push({ sheet: name, rows: n, hasInv: m.inv != null, hasScc: m.scc != null });
    });
    return { sources: sources, lines: lines, index: buildIndex(lines, sources) };
  }

  // Sheet priority: SAP export (tracking) first, then manual sheets
  function rank(sheet) {
    var s = sheet.toLowerCase();
    if (s.indexOf('tracking') >= 0) return 0;
    if (s.indexOf('shipped') >= 0) return 1;
    if (s.indexOf('data') >= 0) return 2;
    return 3;
  }
  var lineKey = function (l) { return [key(l.po), key(l.article), key(l.color || l.rmname)].join('|'); };
  var FIELDS = ['poscaf', 'supplier', 'article', 'rmcode', 'rmname', 'color', 'scc', 'customer', 'story', 'qty', 'unit', 'bu', 'purchaser', 'etd', 'issued'];

  function buildIndex(lines) {
    var merged = {}, order = [];
    lines.slice().sort(function (a, b) { return rank(a.sheet) - rank(b.sheet); }).forEach(function (l) {
      var k = lineKey(l), t = merged[k];
      if (!t) { t = merged[k] = { po: l.po, sheets: [], invs: [], fixed: '' }; order.push(k); }
      if (t.sheets.indexOf(l.sheet) < 0) t.sheets.push(l.sheet);
      FIELDS.forEach(function (f) {
        var v = l[f];
        var empty = t[f] == null || t[f] === '';
        // a qty from a repaired row only fills a gap
        if (empty && v != null && v !== '') t[f] = v;
      });
      if (l.fixed && !t.fixedFrom) t.fixed = l.fixed;
      if (l.inv && t.invs.indexOf(l.inv) < 0) t.invs.push(l.inv);
    });
    var byPo = {}, byInv = {};
    order.forEach(function (k) {
      var t = merged[k];
      [t.po, t.poscaf].forEach(function (p) { p = key(p); if (p) (byPo[p] = byPo[p] || []).push(t); });
    });
    lines.forEach(function (l) {
      if (!l.inv) return;
      var m = merged[lineKey(l)];
      (byInv[key(l.inv)] = byInv[key(l.inv)] || []).push({ line: m, inv: l.inv, invdate: l.invdate, qtyinv: l.qtyinv });
    });
    return { byPo: byPo, byInv: byInv, poList: Object.keys(byPo), invList: Object.keys(byInv) };
  }

  function splitTerms(s) {
    return String(s || '').split(/[\n,;+\s]+/).map(norm).filter(Boolean);
  }

  function search(index, mode, text) {
    var out = [], missing = [];
    splitTerms(text).forEach(function (term) {
      var k = key(term);
      if (mode === 'inv') {
        var hits = index.byInv[k];
        if (!hits) { missing.push(term); return; }
        var seen = {};
        hits.forEach(function (h) {
          var id = h.inv + '|' + lineKey(h.line); if (seen[id]) return; seen[id] = 1;
          out.push(Object.assign({}, h.line, { inv: h.inv, qtyOut: h.qtyinv !== '' ? h.qtyinv : h.line.qty }));
        });
      } else {
        var lines = index.byPo[k];
        if (!lines) { missing.push(term); return; }
        lines.forEach(function (l) { out.push(Object.assign({}, l, { inv: l.invs.join(', '), qtyOut: l.qty })); });
      }
    });
    return { lines: out, missing: missing };
  }

  var SS_HEADERS = ['Supplier sending date', 'DATE OF BUYER SEND TO RM DEPT ', 'DATE QC COMMENT(J+1, except white J+2)', 'PO', 'SUPPLIER', 'Item', 'Code color', 'Color', 'Batch no#', 'Customer ', 'Story ', 'Comment', 'Final decision\n(OK/ NOT OK)', 'QC', 'supplier test report (x)', 'Buyer Name', 'BU  name', 'awb number ', 'quantity/M', 'ITS tes ', 'Report number ', 'test date'];

  function serial(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); if (!m) return null;
    return (Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000;
  }
  function toSSRow(l, opts) {
    return [serial(opts.supplierDate), serial(opts.sendDate), null, l.po, l.supplier, l.article || l.rmname, l.scc, l.color,
      l.batch || '', l.customer, l.story, '', '', '', '', opts.buyer || '', l.bu, opts.awb || '', (l.qtyOut === '' || l.qtyOut == null || isNaN(+l.qtyOut)) ? null : +l.qtyOut, '', '', ''];
  }

  function buildWorkbook(XLSX, lines, opts) {
    var aoa = [SS_HEADERS].concat(lines.map(function (l) { return toSSRow(l, opts); }));
    var ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: 'dd/mm/yyyy' });
    ws['!cols'] = [12, 14, 14, 18, 22, 42, 26, 30, 16, 26, 26, 30, 12, 12, 10, 28, 30, 16, 11, 10, 14, 11].map(function (w) { return { wch: w }; });
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: SS_HEADERS.length - 1 } }) };
    for (var r = 1; r < aoa.length; r++) [0, 1].forEach(function (c) {
      var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })]; if (cell && cell.t === 'n') cell.z = 'dd/mm/yyyy';
    });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, opts.sheetName || 'SS');
    return wb;
  }

  return { parseWorkbook: parseWorkbook, search: search, buildWorkbook: buildWorkbook, splitTerms: splitTerms, SS_HEADERS: SS_HEADERS };
})();
if (typeof module !== 'undefined') module.exports = SSCore;

/* Finance DesignTech - data pipeline
   Single source of truth: runs in Node (to build data.js) and in the browser (Excel upload).
   Reads only the BASE accounting sheets: General Ledger, Balance Sheet, Profit and Loss,
   Cash Flow Statement, Chart of Accounts, Analytic Accounts. */
(function (global) {
  'use strict';

  var EPOCH = Date.UTC(1899, 11, 30);
  var EPOCH1904 = Date.UTC(1904, 0, 1);

  function toNum(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (typeof v === 'boolean') return null;
    var s = String(v).replace(/[\s,]/g, '').replace(/[()]/g, function (m) { return m; });
    var neg = /^\(.*\)$/.test(String(v).trim());
    var n = parseFloat(String(v).replace(/[^0-9.\-eE]/g, ''));
    if (!isFinite(n)) return null;
    return neg ? -Math.abs(n) : n;
  }
  function r2(x) { return x === null || x === undefined ? null : Math.round(x * 100) / 100; }
  function norm(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, ''); }
  function clip(s, n) { s = s == null ? '' : String(s); return s.length > n ? s.slice(0, n) : s; }

  function findSheetByName(names, patterns) {
    for (var i = 0; i < patterns.length; i++) {
      var p = patterns[i];
      for (var j = 0; j < names.length; j++) {
        if (norm(names[j]) === norm(p)) return names[j];
      }
    }
    for (var k = 0; k < patterns.length; k++) {
      var q = norm(patterns[k]);
      for (var m = 0; m < names.length; m++) {
        if (norm(names[m]).indexOf(q) !== -1) return names[m];
      }
    }
    return null;
  }
  function findSheet(wb, patterns) {
    return findSheetByName(wb.SheetNames || [], patterns);
  }
  function pickSheets(SheetNames) {
    var n = SheetNames || [];
    return {
      gl: findSheetByName(n, ['General Ledger', 'Ledger', 'GL']),
      bs: findSheetByName(n, ['Balance Sheet', 'BalanceSheet', 'Balance']),
      pl: findSheetByName(n, ['Profit and Loss', 'Profit Loss', 'P&L', 'Income Statement']),
      cf: findSheetByName(n, ['Cash Flow Statement', 'Cash Flow', 'Cash Flow Statement']),
      coa: findSheetByName(n, ['Chart of Accounts', 'ChartOfAccounts', 'COA', 'Accounts']),
      an: findSheetByName(n, ['Analytic Accounts', 'Analytic', 'Analytics'])
    };
  }

  function rows(XLSX, wb, name) {
    if (!name || !wb.Sheets[name]) return [];
    return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, blankrows: false, defval: null });
  }

  function findHeaderRow(grid, needlers) {
    for (var i = 0; i < Math.min(grid.length, 30); i++) {
      var row = grid[i];
      var joined = norm((row || []).join('|'));
      var ok = true;
      for (var k = 0; k < needlers.length; k++) if (joined.indexOf(norm(needlers[k])) === -1) { ok = false; break; }
      if (ok) return i;
    }
    return -1;
  }
  function colOf(headerRow, names, fallback) {
    for (var i = 0; i < headerRow.length; i++) {
      var h = norm(headerRow[i]);
      for (var j = 0; j < names.length; j++) if (h === norm(names[j])) return i;
    }
    for (var a = 0; a < headerRow.length; a++) {
      var hh = norm(headerRow[a]);
      for (var b = 0; b < names.length; b++) if (hh.indexOf(norm(names[b])) !== -1) return a;
    }
    return fallback;
  }

  function serialToMonth(serial, date1904) {
    var days = date1904 ? serial - 1 : serial;
    var ms = (date1904 ? EPOCH1904 : EPOCH) + Math.round(days * 86400000);
    var d = new Date(ms);
    var y = d.getUTCFullYear(), m = d.getUTCMonth() + 1;
    return { y: y, m: m, key: y + '-' + (m < 10 ? '0' + m : m), iso: d.toISOString().slice(0, 10) };
  }

  // ---------- KPI framework (non-base sheets: KPI Master + KPI Calculation + Data Mapping) ----------
  function lookupLast(XLSX, wb, sheetName, label) {
    if (!sheetName) return { found: false, value: null };
    var g = rows(XLSX, wb, sheetName), target = norm(label), found = false, value = null;
    for (var i = 0; i < g.length; i++) {
      var r = g[i]; if (!r || r[0] == null) continue;
      if (norm(String(r[0])) === target) { found = true; value = toNum(r[1]); }
    }
    return { found: found, value: value };
  }

  function buildKPI(XLSX, wb, extra) {
    extra = extra || {};
    var kmName = findSheet(wb, ['KPI Master']);
    if (!kmName) return null;
    var km = rows(XLSX, wb, kmName);
    var kcName = findSheet(wb, ['KPI Calculation']);
    var dmName = findSheet(wb, ['Data Mapping']);
    var dm = dmName ? rows(XLSX, wb, dmName) : [];
    var kcs = kcName ? wb.Sheets[kcName] : null;

    function str(v) { return v == null ? '' : String(v).trim(); }
    function r6(x) { return x == null ? null : Math.round(x * 1e6) / 1e6; }
    var RE_CODE = /^F\d\d-K\d\d$/;

    // resolve each Data Mapping metric from the base sheets (find = last matching label)
    var MV = {}, METRIC = {};
    for (var i = 1; i < dm.length; i++) {
      var d = dm[i]; if (!d) continue;
      var mkey = str(d[0]); if (!mkey || MV.hasOwnProperty(mkey)) continue;
      var msrc = str(d[1]), mlabel = str(d[2]);
      var missing = true, mval = null;
      if (!/^MANUAL/i.test(msrc) && mlabel) {
        var res = lookupLast(XLSX, wb, findSheet(wb, [msrc]), mlabel);
        if (res.found) { missing = false; mval = res.value == null ? 0 : res.value; }
      }
      if (missing && extra[mkey] != null && extra[mkey] !== '') {
        var ex = extra[mkey];
        if (typeof ex === 'object' && ex.hasOwnProperty('value')) {
          missing = false; mval = ex.value;
          if (ex.source) msrc = ex.source;
          if (ex.label) mlabel = ex.label;
        } else { missing = false; mval = ex; }
      }
      MV[mkey] = missing ? null : mval;
      METRIC[mkey] = { key: mkey, source: msrc, label: mlabel, value: MV[mkey], ready: !missing };
    }
    function G(k) { return MV[k] == null ? null : MV[k]; } // null = missing (poisons the formula)
    function div(a, b) { return (a == null || b == null || b === 0) ? null : a / b; }
    function sub(a, b) { return (a == null || b == null) ? null : a - b; }
    function add(a, b) { return (a == null || b == null) ? null : a + b; }
    function mul(a, b) { return (a == null || b == null) ? null : a * b; }
    function abs(x) { return x == null ? null : Math.abs(x); }

    // one function per KPI code, mirroring the 'KPI Calculation' actual formulas
    var F = {
      'F01-K01': function () { return div(G('BS_Current_Assets'), G('BS_Current_Liabilities')); },
      'F01-K02': function () { return div(sub(sub(G('BS_Current_Assets'), G('BS_Inventory')), G('BS_Prepaid')), G('BS_Current_Liabilities')); },
      'F01-K03': function () { return div(G('CF_Operating'), G('PL_Revenue')); },
      'F01-K04': function () { return div(G('BS_Cash'), div(G('PL_Expenses'), 12)); },
      'F01-K05': function () { return div(G('BS_Total_Liabilities'), G('BS_Total_Equity')); },
      'F01-K06': function () { return div(G('PL_Operating_Income'), G('Interest_Expense')); },
      'F02-K01': function () { return div(sub(G('PL_Revenue'), G('Prior_Revenue')), G('Prior_Revenue')); },
      'F02-K02': function () { return div(G('PL_Gross_Profit'), G('PL_Revenue')); },
      'F02-K03': function () { return div(sub(G('PL_Operating_Income'), G('PL_Depreciation')), G('PL_Revenue')); },
      'F02-K04': function () { return div(G('PL_Net_Profit'), G('PL_Revenue')); },
      'F02-K05': function () { return div(G('PL_Operating_Income'), add(G('BS_Total_Liabilities'), G('BS_Total_Equity'))); },
      'F02-K06': function () { return div(G('PL_Expenses'), G('PL_Revenue')); },
      'F03-K01': function () { return G('BS_Cash'); },
      'F03-K02': function () { return G('CF_Operating'); },
      'F03-K03': function () { return sub(G('CF_Operating'), G('Capex')); },
      'F03-K04': function () { return sub(1, div(abs(sub(G('Forecast_Value'), G('Forecast_Actual'))), abs(G('Forecast_Value')))); },
      'F03-K05': function () { return G('Minimum_Cash_Breach_Rate'); },
      'F03-K06': function () { return div(G('Interest_Expense'), G('PL_Revenue')); },
      'F04-K01': function () { return mul(div(abs(G('BS_Receivables')), G('PL_Revenue')), 365); },
      'F04-K02': function () { return mul(div(G('BS_Payables'), G('PL_Cost_Revenue')), 365); },
      'F04-K03': function () { return null; },
      'F04-K04': function () { return div(G('Cash_Received_Customers'), abs(G('BS_Receivables'))); },
      'F04-K05': function () { return div(G('AR_Overdue'), abs(G('BS_Receivables'))); },
      'F04-K06': function () { return div(sub(G('BS_Current_Assets'), G('BS_Current_Liabilities')), G('PL_Revenue')); },
      'F05-K01': function () { return div(sub(G('Budget_Actual'), G('Budget_Target')), G('Budget_Target')); },
      'F05-K02': function () { return div(sub(G('CostCenter_Actual'), G('CostCenter_Budget')), G('CostCenter_Budget')); },
      'F05-K03': function () { return div(G('PL_Expenses'), G('PL_Revenue')); },
      'F05-K04': function () { return div(sub(G('Procurement_Baseline'), G('Procurement_Actual')), G('Procurement_Baseline')); },
      'F05-K05': function () { return div(G('Unallocated_Cost'), G('PL_Expenses')); },
      'F05-K06': function () { return div(G('PL_Expenses'), G('Revenue_Units')); },
      'F06-K01': function () { return div(G('Budget_Prepared_OnTime'), G('Budget_Cycles')); },
      'F06-K02': function () { return sub(1, div(abs(sub(G('Forecast_Value'), G('Forecast_Actual'))), abs(G('Forecast_Value')))); },
      'F06-K03': function () { return div(G('Budget_Covered_Items'), G('Budget_Total_Items')); },
      'F06-K04': function () { return null; }, 'F06-K05': function () { return null; }, 'F06-K06': function () { return null; },
      'F07-K01': function () { return div(sub(G('Project_Revenue'), G('Project_Cost')), G('Project_Revenue')); },
      'F07-K02': function () { return div(sub(G('Project_Cost'), G('Project_Budget')), G('Project_Budget')); },
      'F07-K03': function () { return div(sub(G('Project_EAC'), G('Project_Budget')), G('Project_Budget')); },
      'F07-K04': function () { return sub(1, div(abs(sub(G('Project_EAC'), G('Project_Cost'))), abs(G('Project_EAC')))); },
      'F07-K05': function () { return div(G('Project_Cash_Collected'), G('Project_Billing')); },
      'F07-K06': function () { return div(G('Loss_Making_Projects'), G('Total_Projects')); },
      'F08-K01': function () { return G('Close_Days'); },
      'F08-K02': function () { return div(G('Report_OnTime'), G('Reports_Total')); },
      'F08-K03': function () { return sub(1, div(G('Report_Errors'), G('Reports_Total'))); },
      'F08-K04': function () { return div(G('Reports_Covered'), G('Reports_Total')); },
      'F08-K05': function () { return G('Decision_Turnaround'); },
      'F08-K06': function () { return div(G('Lineage_Covered'), G('Lineage_Total')); },
      'F09-K01': function () { return div(G('Reconciled_Items'), G('Control_Items')); },
      'F09-K02': function () { return div(G('Policy_Compliant'), G('Policy_Total')); },
      'F09-K03': function () { return div(G('Audit_Closed'), G('Audit_Total')); },
      'F09-K04': function () { return div(G('Policy_Compliant'), G('Policy_Total')); },
      'F09-K05': function () { return div(G('Unusual_Tx'), G('Total_Tx')); },
      'F09-K06': function () { return div(G('Tax_OnTime'), G('Tax_Total')); },
      'F10-K01': function () { return G('Risk_Exposure'); },
      'F10-K02': function () { return div(G('HighRisk_AR'), abs(G('BS_Receivables'))); },
      'F10-K03': function () { return div(G('BS_Cash'), G('Stress_Cash_Need')); },
      'F10-K04': function () { return G('Customer_Concentration'); },
      'F10-K05': function () { return G('Contractual_Exposure'); },
      'F10-K06': function () { return div(G('Risk_Actions_Closed'), G('Risk_Actions_Total')); }
    };

    // targets (Target Low / High) per KPI code from 'KPI Calculation'
    var TARGET = {}, codeRefs = {};
    if (kcs) {
      var kcLen = rows(XLSX, wb, kcName).length;
      for (var c = 0; c < kcLen; c++) {
        var code = str(kcs['A' + (c + 2)] && kcs['A' + (c + 2)].v);
        if (!RE_CODE.test(code)) continue;
        TARGET[code] = { low: toNum(kcs['E' + (c + 2)] && kcs['E' + (c + 2)].v), high: toNum(kcs['F' + (c + 2)] && kcs['F' + (c + 2)].v) };
        var f = kcs['C' + (c + 2)] && kcs['C' + (c + 2)].f ? kcs['C' + (c + 2)].f : '';
        var refs = [], re = /D\$?(\d+)/g, mm;
        while ((mm = re.exec(f))) { var n = parseInt(mm[1], 10); if (refs.indexOf(n) === -1) refs.push(n); }
        codeRefs[code] = refs;
      }
    }

    function scoreOf(actual, low, high) {
      if (actual == null) return null;
      var hasL = low != null, hasH = high != null;
      if (!hasL && !hasH) return null;
      if (hasL && hasH) { if (actual >= low && actual <= high) return 100; if (actual < low) return Math.max(0, actual / low * 100); return Math.max(0, high / actual * 100); }
      if (hasL) return Math.min(100, actual / low * 100);
      // max-only bound (lower-is-better): within/at/below the ceiling scores 100
      if (actual <= high) return 100;
      if (actual <= 0) return 100;
      return Math.max(0, high / actual * 100);
    }
    function statusOf(actual, low, high, score) {
      if (actual == null) return 'Data Required';
      if (low == null && high == null) return 'Target Pending';
      if (score == null) return 'Target Pending';
      if (score >= 100) return 'On Target';
      if (score >= 80) return 'Attention';
      return 'Critical';
    }
    // derive a numeric target from the free-text target when Target Low/High are blank
    function parseTarget(text, unit, direction) {
      if (!text) return null;
      var m = /(>=|≥|<=|≤)\s*([\d.]+)\s*(%|x)?/i.exec(text);
      if (m) {
        var v = parseFloat(m[2]); if (isNaN(v)) return null;
        if (m[3] === '%') v = v / 100;
        return (m[1] === '>=' || m[1] === '≥') ? { low: v, high: null } : { low: null, high: v };
      }
      var m2 = /([\d.]+)\s*%/i.exec(text);
      if (m2) {
        var v2 = parseFloat(m2[1]); if (isNaN(v2)) return null;
        v2 = v2 / 100;
        if (direction === 'Higher') return { low: v2, high: null };
        if (direction === 'Lower') return { low: null, high: v2 };
      }
      return null;
    }

    // build objectives + KPIs from 'KPI Master'
    var objectives = [], byObj = {};
    for (var k = 1; k < km.length; k++) {
      var m = km[k]; if (!m) continue;
      var code = str(m[2]); if (!RE_CODE.test(code)) continue;
      var ocode = str(m[0]);
      if (!byObj[ocode]) { byObj[ocode] = { code: ocode, name: str(m[1]), kpis: [] }; objectives.push(byObj[ocode]); }
      var tgt = TARGET[code] || { low: null, high: null };
      if (tgt.low == null && tgt.high == null) {
        var pt = parseTarget(str(m[9]), str(m[6]), str(m[14]));
        if (pt) tgt = pt;
      }
      var fn = F[code];
      var rawActual = fn ? fn() : null;
      var actual = r6(rawActual);
      var score = rawActual == null ? null : scoreOf(rawActual, tgt.low, tgt.high);
      var inputs = (codeRefs[code] || []).map(function (n) {
        var dr = dm[n - 1]; if (!dr) return null;
        var key = str(dr[0]); var mk = METRIC[key];
        return mk ? { key: mk.key, source: mk.source, label: mk.label, value: r6(mk.value), ready: mk.ready } : null;
      }).filter(Boolean);
      byObj[ocode].kpis.push({
        code: code, name: str(m[3]), purpose: str(m[4]), formula: str(m[5]), unit: str(m[6]),
        refStandard: str(m[7]), targetText: str(m[9]), frequency: str(m[12]), tier: str(m[13]),
        direction: str(m[14]), source: str(m[15]), odooMap: str(m[16]), owner: str(m[17]),
        weight: toNum(m[18]), dashboard: str(m[19]), notes: str(m[20]),
        actual: actual, targetLow: tgt.low, targetHigh: tgt.high,
        score: score, status: actual == null ? 'Data Required' : statusOf(actual, tgt.low, tgt.high, score),
        inputs: inputs
      });
    }

    var statusCounts = {}, sum = 0, scored = 0, readyCount = 0, totalCount = 0;
    objectives.forEach(function (o) {
      var os = 0, on = 0;
      o.kpis.forEach(function (kp) {
        totalCount++;
        statusCounts[kp.status] = (statusCounts[kp.status] || 0) + 1;
        if (kp.score != null) { sum += kp.score; scored++; os += kp.score; on++; }
        if (kp.actual != null) readyCount++;
      });
      o.avgScore = on ? r6(os / on) : null;
      o.counts = o.kpis.reduce(function (a, kp) { a[kp.status] = (a[kp.status] || 0) + 1; return a; }, {});
    });
    var metrics = Object.keys(METRIC).map(function (k) { var mk = METRIC[k]; return { key: mk.key, source: mk.source, label: mk.label, value: r6(mk.value), ready: mk.ready }; });

    return {
      objectives: objectives,
      summary: { total: totalCount, computed: readyCount, scored: scored, overall: scored ? r6(sum / scored) : null, statusCounts: statusCounts },
      metrics: metrics
    };
  }

  function buildModel(XLSX, wb, opts) {
    opts = opts || {};
    var date1904 = opts.date1904 != null ? opts.date1904 : !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
    var progress = opts.progress || function () {};

    var sheetNames = pickSheets(wb.SheetNames);
    progress('sheets', sheetNames);

    // ---------- Chart of Accounts ----------
    var coa = [];
    var cgrid = rows(XLSX, wb, sheetNames.coa);
    if (cgrid.length) {
      var ch = findHeaderRow(cgrid, ['code', 'type']);
      if (ch < 0) ch = findHeaderRow(cgrid, ['code', 'name']);
      if (ch < 0) ch = 0;
      var cCode = colOf(cgrid[ch], ['Code', 'Account Code'], 0);
      var cName = colOf(cgrid[ch], ['Account Name', 'Name'], 1);
      var cType = colOf(cgrid[ch], ['Type', 'Account Type'], 2);
      for (var ci = ch + 1; ci < cgrid.length; ci++) {
        var cr = cgrid[ci]; if (!cr) continue;
        var code = cr[cCode]; if (code == null || norm(code) === 'code') continue;
        coa.push({ code: String(code).trim(), name: cr[cName] == null ? '' : String(cr[cName]).trim(), type: cr[cType] == null ? 'Unmapped' : String(cr[cType]).trim() });
      }
    }
    var typeMap = {}; var codeList = [];
    coa.forEach(function (a) { typeMap[a.code] = a.type; codeList.push(a.code); });
    codeList.sort(function (a, b) { return b.length - a.length; });
    function accType(code) {
      if (typeMap[code]) return typeMap[code];
      for (var i = 0; i < codeList.length; i++) if (code.indexOf(codeList[i]) === 0) return typeMap[codeList[i]];
      return 'Unmapped';
    }

    // ---------- General Ledger ----------
    var ggrid = rows(XLSX, wb, sheetNames.gl);
    var txRows = [], descDict = {}, descList = [], partDict = {}, partList = [];
    var accIdx = {}; // code -> index
    var nodeAgg = {}; // code -> {debit,credit,balance,name}
    var byType = {}; var monthly = {}; var monthsSet = {};
    var revStream = {}; var expGroup = {};
    var parentName = {}; // 3-digit group code -> parent account name (e.g. 501 -> ايردات المشاريع)
    var firstISO = null, lastISO = null, txCount = 0;
    var cashMove = {}; // month -> net
    var accCodesTouched = {};

    var gi0 = findHeaderRow(ggrid, ['date', 'debit']);
    if (gi0 < 0) gi0 = findHeaderRow(ggrid, ['date', 'credit']);
    if (gi0 < 0) gi0 = 0;
    var gHead = ggrid[gi0] || [];
    var colAcc = colOf(gHead, ['Account', 'Code'], 0);
    var colDate = colOf(gHead, ['Date'], 1);
    var colDesc = colOf(gHead, ['Communication', 'Label', 'Description', 'Name', 'Memo'], 2);
    var colPart = colOf(gHead, ['Partner', 'Customer', 'Contact'], 3);
    var colCur = colOf(gHead, ['Currency'], 4);
    var colDr = colOf(gHead, ['Debit'], 5);
    var colCr = colOf(gHead, ['Credit'], 6);
    var colBal = colOf(gHead, ['Balance'], 7);
    if (colAcc === 0 && colDate === 1 && colDr === 5) { /* keep positional defaults */ }

    var descCap = opts.descCap || 120;
    function pushStr(dict, list, s, cap) {
      s = clip(s, cap || descCap);
      if (s === '' || s == null) return -1;
      if (dict[s] == null) { dict[s] = list.length; list.push(s); }
      return dict[s];
    }

    var curCode = null, curName = '';
    var total = ggrid.length;
    for (var gi = gi0 + 1; gi < total; gi++) {
      var r = ggrid[gi]; if (!r) continue;
      var a = r[colAcc];
      var dv = r[colDate];
      var isNumDate = (typeof dv === 'number' && dv > 20000);

      if (a != null) {
        var s = String(a).trim();
        if (!/^total\b/i.test(s)) {
          var m = /^(\d{3,})\s*(.*)$/.exec(s);
          if (m && !isNumDate) {
            curCode = m[1]; curName = m[2] || '';
            if (curCode.length === 3) parentName[curCode] = curName;
            nodeAgg[curCode] = { debit: toNum(r[colDr]) || 0, credit: toNum(r[colCr]) || 0, balance: toNum(r[colBal]) || 0, name: curName };
            if (accIdx[curCode] == null) { accIdx[curCode] = Object.keys(accIdx).length; }
            continue;
          }
        } else {
          continue; // Total rollup line
        }
      }

      if (isNumDate && curCode) {
        var dr = toNum(r[colDr]) || 0, crv = toNum(r[colCr]) || 0;
        var mon = serialToMonth(dv, date1904);
        var t = accType(curCode);
        if (firstISO === null || mon.iso < firstISO) firstISO = mon.iso;
        if (lastISO === null || mon.iso > lastISO) lastISO = mon.iso;
        monthsSet[mon.key] = 1;
        txCount++;
        accCodesTouched[curCode] = 1;
        if (accIdx[curCode] == null) accIdx[curCode] = Object.keys(accIdx).length;
        var dIdx = pushStr(descDict, descList, r[colDesc], descCap);
        var pIdx = pushStr(partDict, partList, r[colPart], 50);
        txRows.push([accIdx[curCode], dv, r2(dr) || 0, r2(crv) || 0, dIdx, pIdx]);

        if (!byType[t]) byType[t] = { debit: 0, credit: 0 };
        byType[t].debit += dr; byType[t].credit += crv;
        if (!monthly[t]) monthly[t] = {};
        var net = ('Income' === t || 'Other Income' === t) ? (crv - dr) : (dr - crv);
        monthly[t][mon.key] = (monthly[t][mon.key] || 0) + net;
        if (t === 'Income') { var sg = curCode.slice(0, 3); revStream[sg] = (revStream[sg] || 0) + (crv - dr); revStream['#' + sg] = curName; }
        if (t === 'Cost of Revenue' || t === 'Expenses' || t === 'Depreciation') { var eg = curCode.slice(0, 3); expGroup[eg] = (expGroup[eg] || 0) + (dr - crv); expGroup['#' + eg] = expGroup['#' + eg] || curName; }
        if (t === 'Bank and Cash') cashMove[mon.key] = (cashMove[mon.key] || 0) + (dr - crv);
      }
      if (gi % 40000 === 0) progress('gl', gi, total);
    }
    progress('gl-done', txCount);

    var months = Object.keys(monthsSet).sort();

    // accounts list merged with GL aggregates
    var accountsArr = [];
    var seen = {};
    coa.forEach(function (a) {
      var g = nodeAgg[a.code];
      accountsArr.push({ code: a.code, name: a.name, type: a.type, debit: g ? r2(g.debit) : null, credit: g ? r2(g.credit) : null, balance: g ? r2(g.balance) : null });
      seen[a.code] = 1;
    });
    Object.keys(nodeAgg).forEach(function (c) {
      if (!seen[c]) accountsArr.push({ code: c, name: nodeAgg[c].name, type: accType(c), debit: r2(nodeAgg[c].debit), credit: r2(nodeAgg[c].credit), balance: r2(nodeAgg[c].balance) });
    });
    // depth by prefix nesting
    var codesOnly = accountsArr.map(function (a) { return a.code; }).sort();
    function parentOf(code) {
      var best = null;
      for (var i = 0; i < codesOnly.length; i++) {
        var c = codesOnly[i];
        if (c !== code && code.indexOf(c) === 0 && c.length < code.length) { if (best === null || c.length > best.length) best = c; }
      }
      return best;
    }
    var accountByCode = {};
    accountsArr.forEach(function (a) {
      var p = parentOf(a.code);
      a.parent = p;
      a.depth = p ? (function () { var d = 1, q = p; while (q) { d++; q = accountByCode[q] ? accountByCode[q].parent : parentOf(q); if (d > 8) break; } return d; })() : 0;
      accountByCode[a.code] = a;
    });

    function series(t) { return months.map(function (m) { return r2((monthly[t] || {})[m] || 0); }); }
    var rev = series('Income'), oi = series('Other Income'), cogs = series('Cost of Revenue'), opex = series('Expenses'), dep = series('Depreciation');
    var gp = [], np = [], ti = [];
    for (var i = 0; i < months.length; i++) {
      gp.push(r2(rev[i] - cogs[i]));
      ti.push(r2(rev[i] + oi[i]));
      np.push(r2(rev[i] + oi[i] - cogs[i] - opex[i] - dep[i]));
    }

    // ---------- statements ----------
    function statement(name) {
      var grid = rows(XLSX, wb, name);
      var out = [];
      for (var i = 0; i < grid.length; i++) {
        var r = grid[i]; if (!r) continue;
        var label = r[0] != null ? String(r[0]).trim() : (r[1] != null && r[1] !== null && typeof r[1] === 'string' && isNaN(toNum(r[1])) ? String(r[1]).trim() : null);
        var raw = r[0] != null ? r[1] : null;
        if (r[0] == null) { if (r[1] != null && isNaN(Number(String(r[1]).replace(/[^0-9.\-]/g, ''))) ) { out.push({ label: String(r[1]).trim(), value: null, head: true }); } continue; }
        var val = toNum(r[1]);
        out.push({ label: label, value: r2(val), total: /^total\b/i.test(label) });
      }
      return out;
    }
    var pl = statement(sheetNames.pl);
    var bs = statement(sheetNames.bs);
    var cf = statement(sheetNames.cf);

    // Use the LAST matching row: statement sheets repeat labels (e.g. "Total Current Assets"
    // appears as a sub-line and as the true total), and the true total is always the later row.
    function find(list, re) { for (var i = list.length - 1; i >= 0; i--) if (re.test(list[i].label)) return list[i].value; return null; }

    // ---------- Analytic Accounts ----------
    var projects = [];
    var agrid = rows(XLSX, wb, sheetNames.an);
    if (agrid.length) {
      var ah = findHeaderRow(agrid, ['analytic', 'balance']);
      if (ah < 0) ah = findHeaderRow(agrid, ['debit', 'credit']);
      if (ah < 0) ah = 0;
      var aName = colOf(agrid[ah], ['Analytic Account', 'Account', 'Name'], 0);
      var aRef = colOf(agrid[ah], ['Reference', 'Ref'], 1);
      var aCust = colOf(agrid[ah], ['Customer', 'Partner', 'Client'], 2);
      var aPlan = colOf(agrid[ah], ['Plan', 'Project'], 3);
      var aDr = colOf(agrid[ah], ['Debit'], 5);
      var aCr = colOf(agrid[ah], ['Credit'], 6);
      var aBal = colOf(agrid[ah], ['Balance'], 7);
      for (var ai = ah + 1; ai < agrid.length; ai++) {
        var ar = agrid[ai]; if (!ar) continue;
        var nm = ar[aName]; if (nm == null || String(nm).trim() === '') continue;
        projects.push({
          name: String(nm).trim(),
          reference: ar[aRef] == null ? '' : String(ar[aRef]).trim(),
          customer: ar[aCust] == null ? '' : String(ar[aCust]).trim(),
          plan: ar[aPlan] == null ? '' : String(ar[aPlan]).trim(),
          debit: r2(toNum(ar[aDr])), credit: r2(toNum(ar[aCr])), balance: r2(toNum(ar[aBal]))
        });
      }
    }

    // ---------- totals & ratios ----------
    var income = r2(byType['Income'] ? byType['Income'].credit - byType['Income'].debit : 0);
    var otherIncome = r2(byType['Other Income'] ? byType['Other Income'].credit - byType['Other Income'].debit : 0);
    var cogsT = r2(byType['Cost of Revenue'] ? byType['Cost of Revenue'].debit - byType['Cost of Revenue'].credit : 0);
    var opexT = r2(byType['Expenses'] ? byType['Expenses'].debit - byType['Expenses'].credit : 0);
    var depT = r2(byType['Depreciation'] ? byType['Depreciation'].debit - byType['Depreciation'].credit : 0);
    var grossProfit = r2(income - cogsT);
    var netProfit = r2(income + otherIncome - cogsT - opexT - depT);

    function sumType(prefixes, sign) { var s = 0; accountsArr.forEach(function (a) { if (prefixes.indexOf(a.type) !== -1 && a.balance != null) s += a.balance; }); return r2(sign * s); }
    var totalAssets = sumType(['Fixed Assets', 'Current Assets', 'Bank and Cash', 'Receivable', 'Prepayments'], 1);
    var totalLiabilities = sumType(['Current Liabilities', 'Non-current Liabilities', 'Payable'], 1);
    var totalEquity = sumType(['Equity', 'Current Year Earnings'], 1);

    var openingCash = find(cf, /beginning of period/i);
    if (openingCash != null) openingCash = Math.abs(openingCash);
    var cashBalance = [], run = openingCash != null ? openingCash : 0;
    for (var mi = 0; mi < months.length; mi++) { run += (cashMove[months[mi]] || 0); cashBalance.push(r2(run)); }
    var closingCash = find(cf, /closing balance/i);
    var netCashIncrease = find(cf, /net increase/i);

    // statement-derived headline (preferred, matches official statements)
    var bsGet = function (re) { return find(bs, re); };
    var headline = {
      revenue: income, otherIncome: otherIncome, income: r2(income + otherIncome),
      cogs: cogsT, grossProfit: grossProfit, opex: opexT, depreciation: depT, netProfit: netProfit,
      totalAssets: bsGet(/^total\s+assets/i) != null ? bsGet(/^total\s+assets/i) : totalAssets,
      currentAssets: bsGet(/total\s+current assets/i),
      fixedAssets: bsGet(/total.*fixed assets/i),
      cash: bsGet(/total\s+bank and cash/i),
      receivables: bsGet(/total\s+receivables/i),
      payables: bsGet(/total\s+payables/i),
      prepayments: bsGet(/total\s+prepayments/i),
      totalLiabilities: bsGet(/^total\s+liabilities/i),
      currentLiabilities: bsGet(/total\s+current liabilities/i),
      nonCurrentLiabilities: bsGet(/total.*non-current liabilities/i),
      totalEquity: bsGet(/^total\s+equity/i),
      openingCash: openingCash, closingCash: closingCash, netCashIncrease: netCashIncrease
    };

    var ratios = {};
    (function () {
      var ca = headline.currentAssets, cl = headline.currentLiabilities;
      var ta = headline.totalAssets, tl = headline.totalLiabilities, te = headline.totalEquity;
      var cash = headline.cash, ar = headline.receivables, ap = headline.payables;
      if (ca != null && cl) ratios.currentRatio = r2(ca / cl);
      if (tl != null && te) ratios.debtToEquity = r2(tl / te);
      if (netProfit != null && ta) ratios.roa = r2(netProfit / ta * 100);
      if (netProfit != null && te) ratios.roe = r2(netProfit / te * 100);
      if (income) { ratios.grossMargin = r2(grossProfit / income * 100); ratios.netMargin = r2(netProfit / (income + otherIncome) * 100); }
      if (cl != null) { ratios.workingCapital = r2((ca || 0) - cl); }
      if (ap) ratios.dpo = r2(ap / (cogsT || 1) * 365);
      if (ar) ratios.dso = r2(ar / (income || 1) * 365);
      if (cash != null && cl) ratios.cashRatio = r2(cash / cl);
    })();

    var revStreams = [], expGroups = [];
    Object.keys(revStream).forEach(function (k) { if (k.charAt(0) !== '#') revStreams.push({ code: k, name: parentName[k] || revStream['#' + k] || k, value: r2(revStream[k]) }); });
    Object.keys(expGroup).forEach(function (k) { if (k.charAt(0) !== '#') expGroups.push({ code: k, name: parentName[k] || expGroup['#' + k] || k, value: r2(expGroup[k]) }); });
    revStreams.sort(function (a, b) { return b.value - a.value; });
    expGroups.sort(function (a, b) { return b.value - a.value; });

    var idxToList = [];
    Object.keys(accIdx).forEach(function (code) { idxToList[accIdx[code]] = code; });

    // ---------- metrics derived from the base workbook (fed to the KPI engine) ----------
    var derived = {};
    var interest = 0, hasInterest = false;
    accountsArr.forEach(function (a) {
      if (a.balance == null) return;
      if (a.type === 'Expenses' && /فوائد تمويلية|interest expense/i.test(a.name || '')) { interest += a.balance; hasInterest = true; }
    });
    if (hasInterest) derived.Interest_Expense = r2(interest);
    var hasInventory = accountsArr.some(function (a) { return /inventor|مخزون|بضاع/i.test(a.name || ''); });
    if (!hasInventory) derived.BS_Inventory = 0;
    derived.Total_Tx = txCount;
    var extraMetrics = Object.assign({}, derived, opts.extra || {});

    var model = {
      meta: {
        company: opts.company || (coa[0] ? 'DesignTech Engineering Consultants' : 'Company'),
        sourceFile: opts.sourceFile || '', period: opts.period || 'FY 2025', currency: opts.currency || 'SAR',
        asOf: (function () { var g = rows(XLSX, wb, sheetNames.bs); for (var i = 0; i < g.length; i++) { var j = norm((g[i] || []).join(' ')); var mm = /asof(\d{4})[\/\-.]?(\d{1,2})[\/\-.]?(\d{1,2})?/.exec(j); if (mm) { return mm[1] + '-' + (mm[2].length < 2 ? '0' + mm[2] : mm[2]) + (mm[3] ? '-' + (mm[3].length < 2 ? '0' + mm[3] : mm[3]) : '-01'); } } return null; })(),
        generatedAt: opts.generatedAt || new Date().toISOString().slice(0, 19),
        txCount: txCount, dateMin: firstISO, dateMax: lastISO, date1904: date1904,
        sheets: sheetNames
      },
      headline: headline, ratios: ratios,
      monthly: { months: months, revenue: rev, otherIncome: oi, totalIncome: ti, cogs: cogs, opex: opex, depreciation: dep, grossProfit: gp, netProfit: np, cashMove: months.map(function (m) { return r2(cashMove[m] || 0); }), cashBalance: cashBalance },
      statements: { pl: pl, bs: bs, cf: cf },
      accounts: accountsArr,
      projects: projects,
      revenueStreams: revStreams, expenseGroups: expGroups,
      tx: { accCodes: idxToList, desc: descList, partner: partList, rows: txRows },
      kpi: buildKPI(XLSX, wb, extraMetrics)
    };
    progress('done', txCount);
    return model;
  }

  var api = { buildModel: buildModel, pickSheets: pickSheets, _toNum: toNum, _buildKPI: buildKPI };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof global !== 'undefined') global.FinPipeline = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);

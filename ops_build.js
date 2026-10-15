/*
 * ops_build.js — reads the six DT2025 operational workbooks and produces the
 * `ops` analytics block + an `extra` metric map for the KPI engine.
 * Node-only (used by build_data.js); the browser app only consumes the output.
 */
(function () {
  'use strict';

  function aoa(XLSX, wb, name) {
    return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false });
  }
  function norm(v) { return v == null ? '' : String(v).trim().toLowerCase().replace(/\s+/g, ' '); }
  function isNonEmpty(v) { return v != null && String(v).trim() !== ''; }
  function n(v) { return (typeof v === 'number' && isFinite(v)) ? v : 0; }
  function r2(x) { return Math.round((x + Number.EPSILON) * 100) / 100; }
  function r6(x) { return Math.round((x + Number.EPSILON) * 1e6) / 1e6; }

  function headerRow(rows, tokens) {
    for (var i = 0; i < Math.min(rows.length, 6); i++) {
      var r = rows[i] || [];
      for (var c = 0; c < r.length; c++) if (tokens.indexOf(norm(r[c])) !== -1) return i;
    }
    return rows.length > 1 ? 1 : 0;
  }
  function colOf(header, names) {
    for (var c = 0; c < header.length; c++) if (names.indexOf(norm(header[c])) !== -1) return c;
    return -1;
  }
  function ser(v) { return (typeof v === 'number' && isFinite(v)) ? Math.floor(v) : null; }
  function monthOf(v) { var s = ser(v); if (s == null) return null; var d = new Date(Date.UTC(1899, 11, 30) + s * 86400000); return d.getUTCMonth(); }
  function dict() {
    var m = {}, list = [];
    return {
      add: function (x) { x = x == null ? '' : String(x).trim(); if (!(x in m)) { m[x] = list.length; list.push(x); } return m[x]; },
      list: function () { return list; }
    };
  }
  function topN(map, size, keyField) {
    var arr = [];
    Object.keys(map).forEach(function (k) { arr.push(map[k]); });
    arr.sort(function (a, b) { return b[keyField || 'total'] - a[keyField || 'total']; });
    return arr.slice(0, size);
  }

  function buildOps(XLSX, dir, fs, path) {
    var F = {
      cust: path.join(dir, 'DT2025_Customer_Contract_Data.xlsx'),
      vat: path.join(dir, 'DT2025_VAT_Tax_Compliance.xlsx'),
      bank: path.join(dir, 'DT2025_Bank_Reconciliations.xlsx'),
      proc: path.join(dir, 'DT2025_Procurement_Data_Purchase Order1.xlsx'),
      cc: path.join(dir, 'DT2025_Cost_Center_Master_Data.xlsx'),
      loan: path.join(dir, 'DT2025_Loans_Repayment_Schedule.xlsx')
    };
    function readBook(p) { return XLSX.read(fs.readFileSync(p), { type: 'buffer', cellDates: false }); }
    var ops = {};
    var extra = {};

    /* ---------------- Customer / AR ---------------- */
    (function () {
      var wb = readBook(F.cust);
      var ig = aoa(XLSX, wb, 'Customer Invoices');
      var h = headerRow(ig, ['partner', 'payment status', 'amount due signed']);
      var hd = ig[h] || [];
      var C = {
        partner: colOf(hd, ['partner']), date: colOf(hd, ['invoice/bill date', 'invoice date', 'date']),
        due: colOf(hd, ['due date']), status: colOf(hd, ['payment status']),
        untaxed: colOf(hd, ['untaxed amount signed']), total: colOf(hd, ['total signed']),
        dueAmt: colOf(hd, ['amount due signed'])
      };
      var partners = dict(), statuses = dict();
      var invoicedMonthly = []; for (var i = 0; i < 12; i++) invoicedMonthly.push(0);
      var collectedMonthly = []; for (i = 0; i < 12; i++) collectedMonthly.push(0);
      var invoices = [], count = 0, invoiced = 0, collected = 0, outstanding = 0;
      var statusCounts = {}, asof = 0;
      for (var r = h + 1; r < ig.length; r++) {
        var row = ig[r]; if (!row || !isNonEmpty(row[C.partner])) continue;
        var tot = n(row[C.total]), due = n(row[C.dueAmt]);
        count++; invoiced += tot; outstanding += due > 0 ? due : 0;
        var st = String(row[C.status] == null ? '' : row[C.status]).trim();
        statusCounts[st] = (statusCounts[st] || 0) + 1;
        var paid = /paid/i.test(st) && !/not paid/i.test(st);
        if (paid) collected += tot;
        var ds = ser(row[C.date]); if (ds != null && ds > asof) asof = ds;
        var mo = monthOf(row[C.date]);
        if (mo != null) { invoicedMonthly[mo] += tot; if (paid) collectedMonthly[mo] += tot; }
        invoices.push([ser(row[C.date]), ser(row[C.due]), partners.add(row[C.partner]), r2(due), statuses.add(st)]);
      }
      // Customer Summary (partner totals)
      var sg = aoa(XLSX, wb, 'Customer Summary');
      var sh = headerRow(sg, ['partner', 'outstanding', 'invoices']);
      var shd = sg[sh] || [];
      var sP = colOf(shd, ['partner']), sT = colOf(shd, ['total']), sO = colOf(shd, ['outstanding']), sI = colOf(shd, ['invoices']);
      var custs = [], grand = 0;
      for (r = sh + 1; r < sg.length; r++) {
        var sr = sg[r]; if (!sr || !isNonEmpty(sr[sP])) continue;
        var nm = String(sr[sP]).trim();
        if (/^(total|المجموع|الإجمالي)/i.test(nm)) continue;
        var tt = n(sr[sT]); custs.push({ name: nm, invoices: n(sr[sI]), total: r2(tt), outstanding: r2(n(sr[sO])) }); grand += tt;
      }
      custs.sort(function (a, b) { return b.total - a.total; });
      // aging of overdue (positive due, due < asof)
      var ag = [0, 0, 0, 0], overdue = 0, high90 = 0;
      for (i = 0; i < invoices.length; i++) {
        var du = invoices[i][1], amt = invoices[i][3];
        if (du == null || amt <= 0 || du >= asof) continue;
        var days = asof - du; overdue += amt;
        if (days <= 30) ag[0] += amt; else if (days <= 60) ag[1] += amt; else if (days <= 90) ag[2] += amt; else { ag[3] += amt; high90 += amt; }
      }
      ops.ar = {
        asof: asof, count: count, invoiced: r2(invoiced), collected: r2(collected), outstanding: r2(outstanding),
        overdue: r2(overdue), high90: r2(high90), aging: ag.map(r2), statusCounts: statusCounts,
        invoicedMonthly: invoicedMonthly.map(r2), collectedMonthly: collectedMonthly.map(r2),
        monthly: invoicedMonthly.map(function (v, i) { return [r2(v), r2(collectedMonthly[i])]; }),
        customers: custs, totalCustomers: custs.length,
        invoices: invoices, partners: partners.list(), statuses: statuses.list()
      };
      extra.AR_Overdue = r2(overdue);
      extra.HighRisk_AR = r2(high90);
      extra.Cash_Received_Customers = r2(collected);
      extra.Customer_Concentration = grand ? r6(custs.length ? custs[0].total / grand : 0) : null;
      extra.AR_Total_Outstanding = r2(outstanding);
    })();

    /* ---------------- Sales orders ---------------- */
    (function () {
      var wb = readBook(F.cust);
      var g = aoa(XLSX, wb, 'Sales Orders');
      var h = headerRow(g, ['order reference', 'invoice status', 'salesperson']);
      var hd = g[h] || [];
      var C = {
        ref: colOf(hd, ['order reference']), cust: colOf(hd, ['customer']), date: colOf(hd, ['order date', 'date']),
        sales: colOf(hd, ['salesperson']), status: colOf(hd, ['status']), inv: colOf(hd, ['invoice status']),
        untaxed: colOf(hd, ['untaxed amount']), total: colOf(hd, ['total'])
      };
      var custD = dict(), salesD = dict(), statD = dict();
      var monthly = []; for (var i = 0; i < 12; i++) monthly.push(0);
      var bySales = {}, byCust = {}, statusCounts = {};
      var orders = [], count = 0, total = 0, untaxed = 0;
      for (var r = h + 1; r < g.length; r++) {
        var row = g[r]; if (!row || !isNonEmpty(row[C.ref])) continue;
        var tt = n(row[C.total]); count++; total += tt; untaxed += n(row[C.untaxed]);
        var sales = String(row[C.sales] == null ? '' : row[C.sales]).trim();
        var cu = String(row[C.cust] == null ? '' : row[C.cust]).trim();
        var stt = String(row[C.status] == null ? '' : row[C.status]).trim();
        statusCounts[stt] = (statusCounts[stt] || 0) + 1;
        if (!bySales[sales]) bySales[sales] = { name: sales || '(—)', total: 0, count: 0 };
        bySales[sales].total += tt; bySales[sales].count++;
        if (!byCust[cu]) byCust[cu] = { name: cu || '(—)', total: 0, count: 0 };
        byCust[cu].total += tt; byCust[cu].count++;
        var mo = monthOf(row[C.date]); if (mo != null) monthly[mo] += tt;
        orders.push([ser(row[C.date]), custD.add(cu), salesD.add(sales), statD.add(stt), r2(tt)]);
      }
      var salesArr = []; Object.keys(bySales).forEach(function (k) { salesArr.push({ name: bySales[k].name, total: r2(bySales[k].total), count: bySales[k].count }); });
      salesArr.sort(function (a, b) { return b.total - a.total; });
      var custArr = []; Object.keys(byCust).forEach(function (k) { custArr.push({ name: byCust[k].name, total: r2(byCust[k].total), count: byCust[k].count }); });
      custArr.sort(function (a, b) { return b.total - a.total; });
      ops.sales = {
        count: count, total: r2(total), untaxed: r2(untaxed), monthly: monthly.map(r2),
        bySales: salesArr, byCustomer: custArr, statusCounts: statusCounts,
        orders: orders, customers: custD.list(), salespeople: salesD.list(), statuses: statD.list()
      };
    })();

    /* ---------------- Procurement ---------------- */
    (function () {
      var wb = readBook(F.proc);
      var g = aoa(XLSX, wb, 'Filtered & Organized Data');
      var h = headerRow(g, ['order reference', 'vendor', 'billing status']);
      var hd = g[h] || [];
      var C = {
        ref: colOf(hd, ['order reference']), vendor: colOf(hd, ['vendor']), deadline: colOf(hd, ['order deadline']),
        confirm: colOf(hd, ['confirmation date']), buyer: colOf(hd, ['buyer']), src: colOf(hd, ['source document']),
        status: colOf(hd, ['status']), billing: colOf(hd, ['billing status']), cur: colOf(hd, ['currency']),
        untaxed: colOf(hd, ['untaxed amount']), tax: colOf(hd, ['taxes']), total: colOf(hd, ['total'])
      };
      var monthly = []; for (var i = 0; i < 12; i++) monthly.push(0);
      var byVendor = {}, statusCounts = {}, billingCounts = {}, orders = [];
      var count = 0, total = 0, untaxed = 0, tax = 0;
      for (var r = h + 1; r < g.length; r++) {
        var row = g[r]; if (!row || !isNonEmpty(row[C.ref])) continue;
        var tt = n(row[C.total]), ux = n(row[C.untaxed]), tx = n(row[C.tax]);
        count++; total += tt; untaxed += ux; tax += tx;
        var v = String(row[C.vendor] == null ? '' : row[C.vendor]).trim();
        if (!byVendor[v]) byVendor[v] = { name: v || '(—)', total: 0, count: 0 };
        byVendor[v].total += tt; byVendor[v].count++;
        var stt = String(row[C.status] == null ? '' : row[C.status]).trim(); statusCounts[stt] = (statusCounts[stt] || 0) + 1;
        var bl = String(row[C.billing] == null ? '' : row[C.billing]).trim(); billingCounts[bl] = (billingCounts[bl] || 0) + 1;
        var mo = monthOf(row[C.confirm]); if (mo != null) monthly[mo] += tt;
        orders.push([row[C.ref] == null ? '' : String(row[C.ref]),
          v, ser(row[C.deadline]), ser(row[C.confirm]),
          String(row[C.buyer] == null ? '' : row[C.buyer]), stt, bl,
          String(row[C.cur] == null ? '' : row[C.cur]), r2(ux), r2(tx), r2(tt)]);
      }
      var vend = []; Object.keys(byVendor).forEach(function (k) { vend.push({ name: byVendor[k].name, total: r2(byVendor[k].total), count: byVendor[k].count }); });
      vend.sort(function (a, b) { return b.total - a.total; });
      ops.proc = { count: count, total: r2(total), untaxed: r2(untaxed), tax: r2(tax), monthly: monthly.map(r2), byVendor: vend, statusCounts: statusCounts, billingCounts: billingCounts, orders: orders };
      extra.Procurement_Actual = r2(total);
    })();

    /* ---------------- VAT / Tax ---------------- */
    (function () {
      var wb = readBook(F.vat);
      var sg = aoa(XLSX, wb, 'Summary by Tax');
      var sh = headerRow(sg, ['originator tax', 'tax_balance', 'lines']);
      var shd = sg[sh] || [];
      var sO = colOf(shd, ['originator tax']), sL = colOf(shd, ['lines']), sB = colOf(shd, ['base']), sBal = colOf(shd, ['tax_balance']);
      var byOrigin = [], salesTax = 0, purchTax = 0, salesBase = 0, purchBase = 0;
      for (var r = sh + 1; r < sg.length; r++) {
        var row = sg[r]; if (!row || !isNonEmpty(row[sO])) continue;
        var nm = String(row[sO]).trim(), bal = n(row[sBal]);
        byOrigin.push({ name: nm, lines: n(row[sL]), base: r2(n(row[sB])), balance: r2(bal) });
        if (/مبيعات|sales/i.test(nm)) { salesTax += Math.abs(bal); salesBase += n(row[sB]); }
        else if (/مشتريات|purchase/i.test(nm)) { purchTax += bal; purchBase += n(row[sB]); }
      }
      var mg = aoa(XLSX, wb, 'Monthly by Tax');
      var mh = headerRow(mg, ['month']);
      var mhd = mg[mh] || [];
      var mMonth = colOf(mhd, ['month']);
      var mSales = -1, mPurch = -1;
      for (var c = 0; c < mhd.length; c++) { var hn = norm(mhd[c]); if (/مبيعات|sales/.test(hn)) mSales = c; else if (/مشتريات|purchase/.test(hn)) mPurch = c; }
      if (mSales < 0) mSales = 1; if (mPurch < 0) mPurch = 2;
      var monthly = [], mstat = {};
      for (var i = 0; i < 12; i++) monthly.push([0, 0]);
      var mName = dict();
      for (r = mh + 1; r < mg.length; r++) {
        var mr = mg[r]; if (!mr) continue;
        var lbl = String(mr[mMonth] == null ? '' : mr[mMonth]).trim();
        var key = lbl.slice(0, 7);
        var mm2 = /^(\d{4})-(\d{1,2})/.exec(key); var mi = mm2 ? (parseInt(mm2[2], 10) - 1) : null;
        if (mi == null || mi < 0 || mi > 11) continue;
        monthly[mi][0] += Math.abs(n(mr[mSales])); monthly[mi][1] += n(mr[mPurch]);
        mName.add(lbl);
      }
      // Tax lines detail (aggregated by month + originator for size)
      var tg = aoa(XLSX, wb, 'Tax Lines');
      var th = headerRow(tg, ['originator tax', 'base amount', 'journal entry']);
      var thd = tg[th] || [];
      var C = {
        date: colOf(thd, ['date']), je: colOf(thd, ['journal entry']), jr: colOf(thd, ['journal']),
        partner: colOf(thd, ['partner']), account: colOf(thd, ['account']), orig: colOf(thd, ['originator tax']),
        base: colOf(thd, ['base amount']), debit: colOf(thd, ['debit']), credit: colOf(thd, ['credit']), bal: colOf(thd, ['balance'])
      };
      var lineCount = 0, baseTotal = 0, debitTotal = 0, creditTotal = 0, balTotal = 0;
      var byMonthTax = []; for (i = 0; i < 12; i++) byMonthTax.push([0, 0]);
      var lines = [], partD = dict(), accD = dict(), jrD = dict(), origD = dict();
      for (r = th + 1; r < tg.length; r++) {
        var tr = tg[r]; if (!tr || !isNonEmpty(tr[C.je])) continue;
        lineCount++; baseTotal += n(tr[C.base]); debitTotal += n(tr[C.debit]); creditTotal += n(tr[C.credit]); balTotal += n(tr[C.bal]);
        var orn = String(tr[C.orig] == null ? '' : tr[C.orig]).trim();
        var mo3 = monthOf(tr[C.date]);
        if (mo3 != null) { if (/مبيعات|sales/i.test(orn)) byMonthTax[mo3][0] += Math.abs(n(tr[C.bal])); else if (/مشتريات|purchase/i.test(orn)) byMonthTax[mo3][1] += n(tr[C.bal]); }
        lines.push([ser(tr[C.date]), jrD.add(tr[C.jr]), partD.add(tr[C.partner]), accD.add(tr[C.account]), origD.add(orn), r2(n(tr[C.base])), r2(n(tr[C.debit])), r2(n(tr[C.credit]))]);
      }
      ops.vat = {
        salesTax: r2(salesTax), purchTax: r2(purchTax), salesBase: r2(salesBase), purchBase: r2(purchBase),
        netTax: r2(salesTax - purchTax), byOriginator: byOrigin, monthly: monthly.map(function (m) { return [r2(m[0]), r2(m[1])]; }),
        lines: lineCount, baseTotal: r2(baseTotal), debitTotal: r2(debitTotal), creditTotal: r2(creditTotal), balTotal: r2(balTotal),
        byMonthTax: byMonthTax.map(function (m) { return [r2(m[0]), r2(m[1])]; }),
        detail: lines, journals: jrD.list(), partners: partD.list(), accounts: accD.list(), originators: origD.list()
      };
      var monthsRecorded = 0; for (i = 0; i < 12; i++) if (monthly[i][0] || monthly[i][1]) monthsRecorded++;
      extra.Tax_Total = monthsRecorded;
      extra.Tax_OnTime = monthsRecorded;
    })();

    /* ---------------- Cost centers ---------------- */
    (function () {
      var wb = readBook(F.cc);
      var pg = aoa(XLSX, wb, 'Cost Centers 2025 P&L');
      var ph = headerRow(pg, ['analytic account', 'net', 'costs']);
      var phd = pg[ph] || [];
      var C = { name: colOf(phd, ['analytic account']), rev: colOf(phd, ['revenue']), cost: colOf(phd, ['costs']), net: colOf(phd, ['net']) };
      var mg = aoa(XLSX, wb, 'Cost Center Master');
      var mh = headerRow(mg, ['analytic account', 'reference', 'plan']);
      var mhd = mg[mh] || [];
      var M = { ref: colOf(mhd, ['reference']), name: colOf(mhd, ['analytic account']), plan: colOf(mhd, ['plan']), cust: colOf(mhd, ['customer']), active: colOf(mhd, ['active']) };
      var meta = {};
      for (var r = mh + 1; r < mg.length; r++) {
        var mr = mg[r]; if (!mr || !isNonEmpty(mr[M.name])) continue;
        meta[norm(mr[M.name])] = { plan: String(mr[M.plan] == null ? '' : mr[M.plan]), customer: String(mr[M.cust] == null ? '' : mr[M.cust]), active: String(mr[M.active] == null ? '' : mr[M.active]) };
      }
      var rows = [], count = 0, revenue = 0, cost = 0, net = 0, lossCount = 0, activeCount = 0;
      for (r = ph + 1; r < pg.length; r++) {
        var pr = pg[r]; if (!pr || !isNonEmpty(pr[C.name])) continue;
        var nm = String(pr[C.name]).trim();
        if (/^(total|المجموع|الإجمالي)/i.test(nm)) continue;
        var rv = n(pr[C.rev]), ct = n(pr[C.cost]), nt = n(pr[C.net]);
        count++; revenue += rv; cost += Math.abs(ct); net += nt; if (nt < 0) lossCount++;
        var mt = meta[norm(nm)] || {};
        var isActive = /^(true|1|yes|نشط|نعم)$/i.test(String(mt.active || 'true'));
        if (isActive) activeCount++;
        rows.push({ name: nm, ref: mt.plan || '', plan: mt.plan || '', customer: mt.customer || '', active: mt.active || '', revenue: r2(rv), cost: r2(Math.abs(ct)), net: r2(nt) });
      }
      rows.sort(function (a, b) { return b.revenue - a.revenue; });
      ops.cc = { count: count, revenue: r2(revenue), cost: r2(cost), net: r2(net), lossCount: lossCount, activeCount: activeCount, rows: rows };
      extra.Project_Revenue = r2(revenue);
      extra.Project_Cost = r2(cost);
      extra.CostCenter_Actual = r2(cost);
      extra.Loss_Making_Projects = lossCount;
      extra.Total_Projects = count;
    })();

    /* ---------------- Bank reconciliations ---------------- */
    (function () {
      var wb = readBook(F.bank);
      var sg = aoa(XLSX, wb, 'Summary by Bank');
      var sh = headerRow(sg, ['journal', 'unreconciled', 'reconciled']);
      var shd = sg[sh] || [];
      var C = { jr: colOf(shd, ['journal']), lines: colOf(shd, ['lines']), in: colOf(shd, ['inflows']), out: colOf(shd, ['outflows']), net: colOf(shd, ['net']), rec: colOf(shd, ['reconciled']), unrec: colOf(shd, ['unreconciled']) };
      var banks = [], reqTotal = 0, recTotal = 0, unrecTotal = 0;
      for (var r = sh + 1; r < sg.length; r++) {
        var row = sg[r]; if (!row || !isNonEmpty(row[C.jr])) continue;
        var ln = n(row[C.lines]), rc = n(row[C.rec]), ur = n(row[C.unrec]);
        banks.push({ journal: String(row[C.jr]).trim(), lines: ln, inflows: r2(n(row[C.in])), outflows: r2(n(row[C.out])), net: r2(n(row[C.net])), reconciled: rc, unreconciled: ur });
        reqTotal += ln; recTotal += rc; unrecTotal += ur;
      }
      var lg = aoa(XLSX, wb, 'Statement Lines');
      var lh = headerRow(lg, ['journal', 'is reconciled', 'label']);
      var lhd = lg[lh] || [];
      var L = { date: colOf(lhd, ['date']), jr: colOf(lhd, ['journal']), stmt: colOf(lhd, ['statement']), label: colOf(lhd, ['label']), partner: colOf(lhd, ['partner']), ref: colOf(lhd, ['reference']), amount: colOf(lhd, ['amount']), rec: colOf(lhd, ['is reconciled']) };
      var monthly = []; for (var i = 0; i < 12; i++) monthly.push([0, 0]);
      var statements = [], jrD = dict(), lblD = dict(), partD = dict();
      var lineCount = 0, recon = 0, unrec = 0, inflow = 0, outflow = 0;
      for (r = lh + 1; r < lg.length; r++) {
        var lr = lg[r]; if (!lr || !isNonEmpty(lr[L.amount])) continue;
        lineCount++;
        var isRec = /^(true|1|yes|نعم)$/i.test(String(lr[L.rec] == null ? '' : lr[L.rec]).trim());
        if (isRec) recon++; else unrec++;
        var amt = n(lr[L.amount]); if (amt >= 0) inflow += amt; else outflow += Math.abs(amt);
        var mo = monthOf(lr[L.date]); if (mo != null) { if (amt >= 0) monthly[mo][0] += amt; else monthly[mo][1] += Math.abs(amt); }
        statements.push([ser(lr[L.date]), jrD.add(lr[L.jr]), lblD.add(lr[L.label]), partD.add(lr[L.partner]), r2(amt), isRec]);
      }
      ops.bankrec = {
        lines: reqTotal, reconciled: recTotal, unreconciled: unrecTotal, rate: reqTotal ? r6(recTotal / reqTotal) : null,
        inflow: r2(inflow), outflow: r2(outflow), monthly: monthly.map(function (m) { return [r2(m[0]), r2(m[1])]; }),
        banks: banks, statements: statements, journals: jrD.list(), labels: lblD.list(), partners: partD.list(),
        stmtLines: lineCount, stmtReconciled: recon, stmtUnreconciled: unrec
      };
      extra.Reconciled_Items = recTotal; extra.Control_Items = reqTotal;
    })();

    /* ---------------- Loans & facilities ---------------- */
    (function () {
      var wb = readBook(F.loan);
      var bg = aoa(XLSX, wb, 'Bank Facilities');
      var bh = headerRow(bg, ['account', 'closing', 'opening']);
      var bhd = bg[bh] || [];
      var C = { acc: colOf(bhd, ['account']), op: 1, cl: 2, mv: 3 };
      for (var c = 0; c < bhd.length; c++) { var hn = norm(bhd[c]); if (/opening/.test(hn)) C.op = c; else if (/closing/.test(hn)) C.cl = c; else if (/movement/.test(hn)) C.mv = c; }
      var facilities = [], opening = 0, closing = 0, movement = 0;
      for (var r = bh + 1; r < bg.length; r++) {
        var row = bg[r]; if (!row || !isNonEmpty(row[C.acc])) continue;
        var op = n(row[C.op]), cl = n(row[C.cl]), mv = n(row[C.mv]);
        opening += op; closing += cl; movement += mv;
        facilities.push({ account: String(row[C.acc]).trim(), opening: r2(op), closing: r2(cl), movement: r2(mv) });
      }
      var fg = aoa(XLSX, wb, 'Facilities GL 2025');
      var fh = headerRow(fg, ['journal entry', 'account', 'label']);
      var fhd = fg[fh] || [];
      var G = { date: colOf(fhd, ['date']), je: colOf(fhd, ['journal entry']), jr: colOf(fhd, ['journal']), acc: colOf(fhd, ['account']), partner: colOf(fhd, ['partner']), label: colOf(fhd, ['label']), debit: colOf(fhd, ['debit']), credit: colOf(fhd, ['credit']), bal: colOf(fhd, ['balance']) };
      var facGL = [];
      for (r = fh + 1; r < fg.length; r++) {
        var fr = fg[r]; if (!fr || !isNonEmpty(fr[G.je])) continue;
        facGL.push([ser(fr[G.date]), String(fr[G.jr] == null ? '' : fr[G.jr]), String(fr[G.acc] == null ? '' : fr[G.acc]), String(fr[G.partner] == null ? '' : fr[G.partner]), String(fr[G.label] == null ? '' : fr[G.label]), r2(n(fr[G.debit])), r2(n(fr[G.credit])), r2(n(fr[G.bal]))]);
      }
      var eg = aoa(XLSX, wb, 'Employee Loans 2025');
      var eh = headerRow(eg, ['reference', 'employee', 'loan amount']);
      var ehd = eg[eh] || [];
      var E = { ref: colOf(ehd, ['reference']), emp: colOf(ehd, ['employee']), date: colOf(ehd, ['date']), amt: colOf(ehd, ['loan amount']), paid: colOf(ehd, ['total paid']), bal: colOf(ehd, ['balance']), state: colOf(ehd, ['state']) };
      var employee = [], empCount = 0, empAmount = 0, empPaid = 0, empBalance = 0;
      for (r = eh + 1; r < eg.length; r++) {
        var er = eg[r]; if (!er || !isNonEmpty(er[E.ref]) || !isNonEmpty(er[E.emp])) continue;
        empCount++; empAmount += n(er[E.amt]); empPaid += n(er[E.paid]); empBalance += n(er[E.bal]);
        employee.push({ ref: String(er[E.ref]).trim(), employee: String(er[E.emp]).trim(), date: ser(er[E.date]), amount: r2(n(er[E.amt])), paid: r2(n(er[E.paid])), balance: r2(n(er[E.bal])), state: String(er[E.state] == null ? '' : er[E.state]).trim() });
      }
      ops.loans = {
        facilities: facilities, facilOpening: r2(opening), facilClosing: r2(closing), facilMovement: r2(movement),
        facilityGL: facGL, employee: employee, empCount: empCount, empAmount: r2(empAmount), empPaid: r2(empPaid), empBalance: r2(empBalance)
      };
    })();

    return { ops: ops, extra: extra };
  }

  var api = { buildOps: buildOps };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof global !== 'undefined') global.OpsBuild = api;
})();

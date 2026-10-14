(function () {
  'use strict';
  var D = window.DASH_DATA;
  var $ = function (s, r) { return (r || document).querySelector(s); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  // ---------------- i18n ----------------
  var I18N = {
    ar: {
      'app.title': 'لوحة مؤشرات الأداء المالي', 'app.sub': 'DesignTech Engineering Consultants',
      'menu': 'القائمة', 'upload': 'رفع ملف Excel', 'up.title': 'رفع ملف Excel', 'up.desc': 'اختر ملف المحاسبة (نفس صيغة Odoo). تُقرأ الشيتات الأساسية وشيتات إطار المؤشرات.',
      'up.prompt': 'اسحب الملف هنا أو اضغط للاختيار', 'up.reading': 'جارٍ القراءة والمعالجة…', 'up.ok': 'تم التحديث', 'up.err': 'تعذّر قراءة الملف',
      'grp.lead': 'القيادة', 'grp.obj': 'الأهداف الاستراتيجية', 'grp.ref': 'الطبقات المرجعية', 'grp.exp': 'المستكشف التفصيلي',
      'nav.exec': 'اللوحة التنفيذية', 'nav.fin': 'لوحة الإدارة المالية',
      'nav.master': 'سجل المؤشرات', 'nav.mapping': 'خريطة البيانات',
      'master.title': 'سجل المؤشرات الرئيسية', 'master.sub': 'التعريفات والمعادلات والأهداف والمسؤولية والمصادر — لكل المؤشرات الستين في مكان واحد',
      'map.title': 'خريطة البيانات ومصادر Odoo', 'map.sub': 'كل مفاتيح البيانات المربوطة بمصادرها وقيمها الحالية',
      'chain.src': 'مصدر Odoo', 'chain.calc': 'حساب المؤشر', 'chain.dash': 'لوحة المؤشر',
      'f.objective': 'المحور', 'f.status': 'الحالة',
      'c.key': 'المفتاح', 'c.source': 'المصدر', 'c.label': 'البند', 'c.value': 'القيمة', 'c.avail': 'التوفر',
      'nav.tx': 'كل الحركات', 'nav.proj': 'المشاريع', 'nav.acc': 'دليل الحسابات', 'nav.cust': 'العملاء / الموردون',
      'exec.title': 'اللوحة التنفيذية المالية', 'exec.sub': 'مرجع الإدارة الأول — نظرة شاملة على محاور الأداء المالي العشرة (F01–F10)',
      'fin.title': 'لوحة الإدارة المالية', 'fin.sub': 'تجميع الإدارات عبر المحاور F01–F10 ومتابعة الاستثناءات وأولويات المعالجة',
      'score.overall': 'المؤشر المالي العام', 'crit.obj': 'محاور حرجة', 'att.obj': 'محاور تحتاج انتباهاً',
      'computed': 'مؤشرات قابلة للحساب', 'datareq': 'مؤشرات بيانات مطلوبة',
      'exec.guide': 'دليل القرار التنفيذي',
      'exec.guide.txt': 'ابدأ من هنا: حدّد المحاور الحرجة أو التي تحتاج انتباهاً ← انزل إلى لوحة المحور ← افتح المؤشر لمراجعة القيمة والهدف والنتيجة والحالة وخط الإمداد بالبيانات.',
      'exec.limit': 'حد البيانات', 'exec.limit.txt': 'مقاييس الاتجاه/المقارنة الزمنية تتطلب تصديرات Odoo لفترات متعددة. الملف الحالي يوفّر حسابات الفترة الحالية فقط.',
      'fin.priority': 'أولوية الإدارة',
      'fin.priority.txt': 'حرج ← إجراء فوري · يحتاج انتباهاً ← متابعة إدارية · بيانات مطلوبة / بانتظار الهدف ← مشكلة جاهزية · على الهدف ← مراقبة.',
      'tbl.obj': 'المحور', 'tbl.name': 'الاسم', 'tbl.scoreobj': 'النتيجة', 'tbl.exc': 'الاستثناءات', 'tbl.status': 'الحالة', 'tbl.drill': 'التنزّل', 'tbl.focus': 'التركيز التنفيذي', 'tbl.action': 'الإجراء',
      'tbl.kpicode': 'الرمز', 'tbl.kpi': 'المؤشر', 'tbl.actual': 'القيمة الفعلية', 'tbl.score': 'النتيجة', 'tbl.target': 'الهدف', 'tbl.dashboard': 'اللوحة', 'tbl.source': 'المصدر',
      'drill.obj': 'لوحة المحور', 'drill.fin': 'لوحة الإدارة المالية', 'drill.exec': 'اللوحة التنفيذية', 'drill.calc': 'المعادلة والحساب',
      'review.only': 'مراجعة الاستثناءات وجاهزية البيانات فقط', 'review.obj': 'مراجعة لوحة المحور', 'review.fin': 'فتح لوحة الإدارة المالية',
      'avg.score': 'متوسط النتيجة', 'kpi.count': 'عدد المؤشرات',
      'kpi.title': 'لوحة المؤشر', 'kpi.actual': 'القيمة الفعلية', 'kpi.target': 'الهدف / الحد', 'kpi.score100': 'النتيجة / 100', 'kpi.status': 'الحالة',
      'sec.manage': 'النظرة الإدارية', 'sec.trend': 'الاتجاه / التاريخ', 'sec.def': 'تعريف المؤشر وحاكميته', 'sec.lineage': 'خط الإمداد بالبيانات', 'sec.drill': 'التنزّل',
      'manage.decision': 'قرار الإدارة', 'perf.readiness': 'جاهزية البيانات',
      'trend.na': 'غير متاح من تصدير Odoo الحالي — يلزم توفير فترات تاريخية لتحليل الاتجاه.',
      'ready.ok': 'القيمة محسوبة من الملف الحالي.', 'ready.missing': 'المدخلات غير متوفرة في الملف الحالي.',
      'def.formula': 'طريقة الحساب', 'def.unit': 'الوحدة', 'def.ref': 'المعيار المرجعي', 'def.source': 'مصدر البيانات', 'def.freq': 'التكرار', 'def.dir': 'الاتجاه', 'def.owner': 'المسؤول', 'def.tier': 'الطبقة', 'def.odoo': 'الربط / Odoo', 'def.purpose': 'الغرض', 'def.weight': 'الوزن', 'def.freq2': 'التكرار',
      'lin.key': 'المفتاح', 'lin.src': 'المصدر', 'lin.val': 'القيمة', 'lin.st': 'الحالة', 'lin.ok': 'متاح', 'lin.no': 'مطلوب',
      'open.tx': 'فتح الحركات المرتبطة', 'lin.none': 'لا توجد مدخلات مربوطة',
      'st.on': 'على الهدف', 'st.att': 'يحتاج انتباهاً', 'st.crit': 'حرج', 'st.na': 'بيانات مطلوبة', 'st.pend': 'بانتظار الهدف',
      'days': 'يوم', 'months': 'شهر', 'hours': 'ساعة', 'times': 'مرة', 'cur': 'ر.س',
      'exp.title': 'المستكشف التفصيلي', 'exp.sub': 'كل حركة وكل مشروع وكل حساب وعميل — بفلتر بسيط (الفترة + القوائم)',
      'tab.tx': 'الحركات', 'tab.proj': 'المشاريع', 'tab.acc': 'الحسابات', 'tab.cust': 'العملاء / الموردون',
      'f.period': 'الفترة', 'f.all': 'كل الفترة', 'f.account': 'الحساب', 'f.customer': 'العميل / المورد', 'f.type': 'نوع الحساب', 'f.reset': 'تصفير',
      'c.date': 'التاريخ', 'c.acc': 'الحساب', 'c.accname': 'اسم الحساب', 'c.desc': 'البيان', 'c.partner': 'العميل / المورد', 'c.debit': 'مدين', 'c.credit': 'دائن', 'c.net': 'الصافي',
      'c.code': 'الرمز', 'c.name': 'الاسم', 'c.bal': 'الرصيد', 'c.ref': 'المرجع', 'c.plan': 'الخطة', 'c.moves': 'الحركات', 'c.unassigned': 'غير محدد',
      'c.count': 'عدد الحركات', 'c.revenue': 'إجمالي المدين', 'c.spend': 'إجمالي الدائن',
      'sum.filtered': 'الإجمالي (بعد الفلتر)', 'sum.all': 'الإجمالي الكلي',
      'rows': 'سجل', 'page': 'صفحة', 'prev': 'السابق', 'next': 'التالي', 'noRows': 'لا توجد نتائج مطابقة',
      'obj.proj': 'مشروع', 'obj.cust': 'عميل',
      'gauge.score': 'النتيجة'
    },
    en: {
      'app.title': 'Financial KPI Dashboard', 'app.sub': 'DesignTech Engineering Consultants',
      'menu': 'Menu', 'upload': 'Upload Excel', 'up.title': 'Upload Excel', 'up.desc': 'Choose the accounting workbook (Odoo format). Base sheets and KPI-framework sheets are read.',
      'up.prompt': 'Drag the file here or click to choose', 'up.reading': 'Reading and processing…', 'up.ok': 'Updated', 'up.err': 'Could not read file',
      'grp.lead': 'Leadership', 'grp.obj': 'Strategic Objectives', 'grp.ref': 'Reference Layers', 'grp.exp': 'Detailed Explorer',
      'nav.exec': 'Executive Dashboard', 'nav.fin': 'Finance Dashboard',
      'nav.master': 'KPI Master', 'nav.mapping': 'Data Mapping',
      'master.title': 'KPI Master Register', 'master.sub': 'Definitions, formulas, targets, ownership and sources — all sixty KPIs in one place',
      'map.title': 'Data Mapping & Odoo Sources', 'map.sub': 'Every data key mapped to its source and current value',
      'chain.src': 'Odoo Source', 'chain.calc': 'KPI Calculation', 'chain.dash': 'KPI Dashboard',
      'f.objective': 'Objective', 'f.status': 'Status',
      'c.key': 'Key', 'c.source': 'Source', 'c.label': 'Item', 'c.value': 'Value', 'c.avail': 'Availability',
      'nav.tx': 'All Transactions', 'nav.proj': 'Projects', 'nav.acc': 'Chart of Accounts', 'nav.cust': 'Customers / Vendors',
      'exec.title': 'Executive Finance Dashboard', 'exec.sub': 'Primary management reference — a full view across the ten financial objectives (F01–F10)',
      'fin.title': 'Finance Dashboard', 'fin.sub': 'Department aggregation across F01–F10 with exception tracking and management priorities',
      'score.overall': 'Overall Finance Score', 'crit.obj': 'Critical Objectives', 'att.obj': 'Attention Objectives',
      'computed': 'Computable Indicators', 'datareq': 'Data Required',
      'exec.guide': 'Executive Decision Guide',
      'exec.guide.txt': 'Start here: identify Critical / Attention objectives → drill to the objective → open the KPI to review Actual, Target, Score, Status and data lineage.',
      'exec.limit': 'Data limitation', 'exec.limit.txt': 'Trend / period-over-period measures require periodized Odoo exports. Current source files provide current-period calculations only.',
      'fin.priority': 'Management Priority',
      'fin.priority.txt': 'Critical → immediate action · Attention → management attention · Data Required / Target Pending → readiness issue · On Target → monitor.',
      'tbl.obj': 'Objective', 'tbl.name': 'Name', 'tbl.scoreobj': 'Score', 'tbl.exc': 'Exceptions', 'tbl.status': 'Status', 'tbl.drill': 'Drill-down', 'tbl.focus': 'Executive Focus', 'tbl.action': 'Primary Action',
      'tbl.kpicode': 'KPI Code', 'tbl.kpi': 'KPI', 'tbl.actual': 'Actual', 'tbl.score': 'Score', 'tbl.target': 'Target', 'tbl.dashboard': 'Dashboard', 'tbl.source': 'Source',
      'drill.obj': 'Objective Dashboard', 'drill.fin': 'Finance Dashboard', 'drill.exec': 'Executive Dashboard', 'drill.calc': 'Calculation',
      'review.only': 'Review only exceptions and readiness gaps', 'review.obj': 'Review objective dashboard', 'review.fin': 'Open finance dashboard',
      'avg.score': 'Average Score', 'kpi.count': 'KPIs',
      'kpi.title': 'KPI Performance Dashboard', 'kpi.actual': 'Actual', 'kpi.target': 'Target / Threshold', 'kpi.score100': 'Score / 100', 'kpi.status': 'Status',
      'sec.manage': 'Management View', 'sec.trend': 'Trend / History', 'sec.def': 'KPI Definition & Governance', 'sec.lineage': 'Data Lineage', 'sec.drill': 'Drill-down',
      'manage.decision': 'Management Decision', 'perf.readiness': 'Performance Readiness',
      'trend.na': 'Not available from current Odoo export — historical periods are required for trend analysis.',
      'ready.ok': 'Value is computed from the current file.', 'ready.missing': 'Inputs are not supplied in the current file.',
      'def.formula': 'Calculation / Formula', 'def.unit': 'Unit', 'def.ref': 'Reference Standard', 'def.source': 'Data Source', 'def.freq': 'Frequency', 'def.dir': 'Direction', 'def.owner': 'Owner', 'def.tier': 'Tier', 'def.odoo': 'Odoo / Data Mapping', 'def.purpose': 'Purpose', 'def.weight': 'Weight', 'def.freq2': 'Frequency',
      'lin.key': 'Key', 'lin.src': 'Source', 'lin.val': 'Value', 'lin.st': 'Status', 'lin.ok': 'available', 'lin.no': 'required',
      'open.tx': 'Open related transactions', 'lin.none': 'No mapped inputs',
      'st.on': 'On Target', 'st.att': 'Attention', 'st.crit': 'Critical', 'st.na': 'Data Required', 'st.pend': 'Target Pending',
      'days': 'days', 'months': 'months', 'hours': 'hours', 'times': 'times', 'cur': 'SAR',
      'exp.title': 'Detailed Explorer', 'exp.sub': 'Every transaction, project, account and customer — with a simple filter (period + dropdowns)',
      'tab.tx': 'Transactions', 'tab.proj': 'Projects', 'tab.acc': 'Accounts', 'tab.cust': 'Customers / Vendors',
      'f.period': 'Period', 'f.all': 'All period', 'f.account': 'Account', 'f.customer': 'Customer / Vendor', 'f.type': 'Account Type', 'f.reset': 'Reset',
      'c.date': 'Date', 'c.acc': 'Account', 'c.accname': 'Account Name', 'c.desc': 'Label', 'c.partner': 'Customer / Vendor', 'c.debit': 'Debit', 'c.credit': 'Credit', 'c.net': 'Net',
      'c.code': 'Code', 'c.name': 'Name', 'c.bal': 'Balance', 'c.ref': 'Reference', 'c.plan': 'Plan', 'c.moves': 'Movements', 'c.unassigned': 'Unassigned',
      'c.count': 'Transactions', 'c.revenue': 'Total Debit', 'c.spend': 'Total Credit',
      'sum.filtered': 'Total (filtered)', 'sum.all': 'Grand Total',
      'rows': 'rows', 'page': 'Page', 'prev': 'Prev', 'next': 'Next', 'noRows': 'No matching results',
      'obj.proj': 'Project', 'obj.cust': 'Customer',
      'gauge.score': 'Score'
    }
  };
  function T(k) { var L = I18N[state.lang] || I18N.ar; return L[k] != null ? L[k] : k; }

  // ---------------- state ----------------
  var state = {
    lang: 'ar', theme: 'dark', view: 'exec', obj: null, kpi: null,
    f: { period: 'all', account: null, customer: null, type: null, mobj: null, mstat: null, page: 1 }, txCache: null
  };
  (function () {
    try {
      var lg = localStorage.getItem('dtk.lang'); if (lg === 'ar' || lg === 'en') state.lang = lg;
      var th = localStorage.getItem('dtk.theme'); if (th === 'dark' || th === 'light') state.theme = th;
    } catch (e) {}
  })();

  // ---------------- formatting ----------------
  function nf(x, d) { if (x == null || isNaN(x)) return '—'; var n = Number(x), dd = d == null ? 2 : d; if (Math.abs(n) < Math.pow(10, -dd) / 2) n = 0; return n.toLocaleString('en-US', { maximumFractionDigits: dd }); }
  function money(x) { return nf(x, 2) + ' ' + T('cur'); }
  function pct(x) { return x == null ? '—' : (x * 100).toFixed(1) + '%'; }
  function kpiVal(kp) {
    if (kp.actual == null) return '—';
    var u = String(kp.unit || '');
    if (/SAR|ر\.س/i.test(u)) return money(kp.actual);
    if (u === '%') return pct(kp.actual);
    if (u === 'x') return nf(kp.actual, 2) + '×';
    if (u === 'Days') return nf(kp.actual, 1) + ' ' + T('days');
    if (u === 'Months') return nf(kp.actual, 1) + ' ' + T('months');
    if (u === 'Hours') return nf(kp.actual, 1) + ' ' + T('hours');
    return nf(kp.actual, 2);
  }
  function kpiTarget(kp) {
    var u = String(kp.unit || '');
    function f(v) { return v == null ? null : (u === '%' ? pct(v) : nf(v, 2)); }
    var a = f(kp.targetLow), b = f(kp.targetHigh);
    if (a != null && b != null) return a + ' – ' + b;
    if (a != null) return '≥ ' + a;
    if (b != null) return '≤ ' + b;
    return '—';
  }

  // ---------------- status ----------------
  var ST = {
    'On Target': { k: 'st.on', c: 'pos' }, 'Attention': { k: 'st.att', c: 'warn' },
    'Critical': { k: 'st.crit', c: 'neg' }, 'Data Required': { k: 'st.na', c: 'muted2' }, 'Target Pending': { k: 'st.pend', c: 'muted' }
  };
  function stMeta(s) { return ST[s] || { k: 'st.pend', c: 'muted' }; }
  function stLabel(s) { return T(stMeta(s).k); }
  function stColor(s) { var c = stMeta(s).c; return c === 'pos' ? 'var(--pos)' : c === 'warn' ? 'var(--warn)' : c === 'neg' ? 'var(--neg)' : c === 'muted2' ? 'var(--muted2)' : 'var(--muted)'; }
  function stPill(s) { return '<span class="kstat" style="color:' + stColor(s) + '">' + esc(stLabel(s)) + '</span>'; }
  function scoreBar(sc) { if (sc == null) return '<span class="mut">—</span>'; var col = sc >= 100 ? 'var(--pos)' : sc >= 80 ? 'var(--warn)' : 'var(--neg)'; return '<span class="sbar" style="--sbc:' + col + ';width:56px"><i style="width:' + Math.max(3, Math.min(100, sc)) + '%;background:' + col + '"></i></span><span class="sval">' + nf(sc, 1) + '</span>'; }
  function objStatus(o) { if (o.avgScore == null) return 'Data Required'; if (o.avgScore >= 100) return 'On Target'; if (o.avgScore >= 80) return 'Attention'; return 'Critical'; }

  // ---------------- indices ----------------
  var OBJ = {}, KPI = {}, KO = [];
  (D.kpi ? D.kpi.objectives : []).forEach(function (o) { OBJ[o.code] = o; KO.push(o); o.kpis.forEach(function (k) { KPI[k.code] = k; k._obj = o.code; k._objName = o.name; }); });
  var SUMMARY = D.kpi ? D.kpi.summary : { overall: null, computed: 0, total: 0, statusCounts: {} };

  var ACCB = {}; D.accounts.forEach(function (a) { ACCB[a.code] = a; });
  var TXACC = (D.tx.accCodes || []).map(function (c) { var a = ACCB[c]; return { code: c, name: a ? a.name : c, type: a ? a.type : '' }; });
  var PART = D.tx.partner || [];
  var MONTHS = (D.monthly && D.monthly.months) || [];
  function serialToKey(serial) {
    var EPOCH = Date.UTC(1899, 11, 30); var days = D.meta.date1904 ? serial - 1 : serial;
    var d = new Date(EPOCH + Math.round(days * 86400000));
    var y = d.getUTCFullYear(), m = d.getUTCMonth() + 1;
    return y + '-' + (m < 10 ? '0' + m : m);
  }
  var ROWM = null; // lazy: per-row month key
  function rowMonths() {
    if (ROWM) return ROWM;
    var rows = D.tx.rows, m = new Array(rows.length);
    for (var i = 0; i < rows.length; i++) { var s = rows[i][1]; m[i] = (typeof s === 'number' && s > 20000) ? serialToKey(s) : ''; }
    ROWM = m; return m;
  }
  function rowISO(serial) { var EPOCH = Date.UTC(1899, 11, 30); var days = D.meta.date1904 ? serial - 1 : serial; var d = new Date(EPOCH + Math.round(days * 86400000)); return d.toISOString().slice(0, 10); }

  var CUSTOMERS = null;
  function customers() {
    if (!CUSTOMERS) { var arr = []; PART.forEach(function (p, i) { if (p && String(p).trim()) arr.push({ i: i, name: p }); }); arr.sort(function (a, b) { return a.name.localeCompare(b.name); }); CUSTOMERS = arr; }
    if (hasUnassigned()) return CUSTOMERS.concat([{ i: -1, name: T('c.unassigned') }]);
    return CUSTOMERS;
  }
  var UNASSIGNED = null;
  function hasUnassigned() {
    if (UNASSIGNED != null) return UNASSIGNED;
    var u = false; for (var i = 0; i < D.tx.rows.length; i++) { if (D.tx.rows[i][5] === -1) { u = true; break; } }
    UNASSIGNED = u; return u;
  }
  function custName(i) { if (i === -1) return T('c.unassigned'); return PART[i] || String(i); }
  var TYPES = null;
  function accTypes() {
    if (TYPES) return TYPES;
    var seen = {}, arr = [];
    D.accounts.forEach(function (a) { if (a.type && a.type !== 'Unmapped' && !seen[a.type]) { seen[a.type] = 1; arr.push(a.type); } });
    arr.sort(); TYPES = arr; return arr;
  }
  var ACCLIST = null;
  function accountsList() {
    if (ACCLIST) return ACCLIST;
    var arr = [];
    D.accounts.forEach(function (a) { if (a.type !== 'Unmapped') arr.push(a); });
    arr.sort(function (x, y) { return x.code < y.code ? -1 : x.code > y.code ? 1 : 0; });
    ACCLIST = arr; return arr;
  }
  var ACCN = null; // lazy: tx-row count per account code
  function accCounts() {
    if (ACCN) return ACCN;
    var m = {}; D.tx.rows.forEach(function (r) { var c = D.tx.accCodes[r[0]]; if (c != null) m[c] = (m[c] || 0) + 1; });
    ACCN = m; return m;
  }

  // ---------------- filtering ----------------
  function txFiltered() {
    var rows = D.tx.rows, m = rowMonths(), f = state.f, out = [];
    var per = f.period, acc = f.account, cust = f.customer;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (per !== 'all' && m[i] !== per) continue;
      if (acc != null && D.tx.accCodes[r[0]] !== acc) continue;
      if (cust != null && r[5] !== cust) continue;
      out.push(i);
    }
    return out;
  }

  // ---------------- filter bar ----------------
  var FOPT = {};
  function fdropHTML(key, label, items, value, icon) {
    FOPT[key] = items;
    var cur = null; for (var i = 0; i < items.length; i++) if (items[i].v === value) { cur = items[i]; break; }
    if (!cur) cur = items[0];
    var on = value != null && value !== 'all';
    var opts = items.map(function (it, idx) {
      return '<button class="popt' + (it.v === value ? ' on' : '') + '" data-opt="' + key + '|' + idx + '">' + esc(it.label) + '</button>';
    }).join('');
    return '<div class="fdrop' + (on ? ' on' : '') + '" data-fkey="' + key + '">' +
      '<button class="fdrop-btn"><span class="fd-ic">' + (icon || '▾') + '</span><span class="fd-txt">' + esc(label) + ': ' + esc(cur ? cur.label : '') + '</span><span class="caret">▾</span></button>' +
      '<div class="fdrop-panel"><div class="pdp-cat">' + esc(label) + '</div><div class="pdp-grid">' + opts + '</div></div></div>';
  }
  function periodItems() { return [{ v: 'all', label: T('f.all') }].concat(MONTHS.map(function (m) { return { v: m, label: m }; })); }
  function accountItems() { var a = [{ v: null, label: T('f.all') }]; accountsList().forEach(function (x) { a.push({ v: x.code, label: x.code + ' · ' + x.name }); }); return a; }
  function customerItems() { var a = [{ v: null, label: T('f.all') }]; customers().forEach(function (c) { a.push({ v: c.i, label: c.name }); }); return a; }
  function typeItems() { return [{ v: null, label: T('f.all') }].concat(accTypes().map(function (t) { return { v: t, label: t }; })); }
  function objectiveFilterItems() { return [{ v: null, label: T('f.all') }].concat(KO.map(function (o) { return { v: o.code, label: o.code + ' · ' + o.name }; })); }
  function statusItems() { return [{ v: null, label: T('f.all') }].concat(['On Target', 'Attention', 'Critical', 'Target Pending', 'Data Required'].map(function (s) { return { v: s, label: stLabel(s) }; })); }

  function filterBar() {
    var f = state.f, h = '';
    var showPeriod = state.view === 'tx' || state.view === 'cust';
    if (showPeriod) h += fdropHTML('period', T('f.period'), periodItems(), f.period, '🗓');
    if (state.view === 'tx') { h += fdropHTML('account', T('f.account'), accountItems(), f.account, '▤'); h += fdropHTML('customer', T('f.customer'), customerItems(), f.customer, '☺'); }
    if (state.view === 'acc') h += fdropHTML('type', T('f.type'), typeItems(), f.type, '▤');
    if (state.view === 'proj') h += fdropHTML('customer', T('f.customer'), customerItems(), f.customer, '☺');
    if (state.view === 'master') { h += fdropHTML('mobj', T('f.objective'), objectiveFilterItems(), f.mobj, '◈'); h += fdropHTML('mstat', T('f.status'), statusItems(), f.mstat, '●'); }
    var chips = activeChips();
    return '<div class="fbar"><div class="fbar-row"><span class="fbar-title"><span class="ic">⚲</span>' + T('f.period') + '</span>' + h +
      '<div class="spacer" style="flex:1"></div><button class="btn sm" id="fReset">↺ ' + T('f.reset') + '</button></div>' + (chips ? '<div class="factive">' + chips + '</div>' : '') + '</div>';
  }
  function activeChips() {
    var f = state.f, c = [];
    function chip(key, label) { return '<span class="fchip">' + esc(label) + '<b data-chp="' + key + '">✕</b></span>'; }
    if ((state.view === 'tx' || state.view === 'cust') && f.period !== 'all') c.push(chip('period', f.period));
    if (state.view === 'tx' && f.account != null) c.push(chip('account', ACCB[f.account] ? (ACCB[f.account].code + ' · ' + ACCB[f.account].name) : f.account));
    if ((state.view === 'tx' || state.view === 'proj') && f.customer != null) c.push(chip('customer', custName(f.customer)));
    if (state.view === 'acc' && f.type != null) c.push(chip('type', f.type));
    if (state.view === 'master' && f.mobj != null) c.push(chip('mobj', OBJ[f.mobj] ? (f.mobj + ' · ' + OBJ[f.mobj].name) : f.mobj));
    if (state.view === 'master' && f.mstat != null) c.push(chip('mstat', stLabel(f.mstat)));
    return c.join('');
  }

  // ---------------- nav ----------------
  function nav() {
    var o = '';
    o += '<div class="grp">' + T('grp.lead') + '</div>';
    o += navBtn('exec', '▤', T('nav.exec'));
    o += navBtn('fin', '▦', T('nav.fin'));
    o += '<div class="grp">' + T('grp.obj') + '</div>';
    KO.forEach(function (obj) {
      var s = obj.avgScore == null ? '—' : nf(obj.avgScore, 1);
      o += '<button data-nav="obj" data-obj="' + obj.code + '" class="' + (state.view === 'obj' && state.obj === obj.code ? 'on' : '') + '"><span class="ic">◈</span><span>' + esc(obj.code + ' · ' + obj.name) + '</span><span class="nx">' + s + '</span></button>';
    });
    o += '<div class="grp">' + T('grp.ref') + '</div>';
    o += navBtn('master', '◆', T('nav.master'));
    o += navBtn('mapping', '⇄', T('nav.mapping'));
    o += '<div class="grp">' + T('grp.exp') + '</div>';
    o += navBtn('tx', '≡', T('nav.tx'));
    o += navBtn('proj', '◧', T('nav.proj'));
    o += navBtn('acc', '▤', T('nav.acc'));
    o += navBtn('cust', '☺', T('nav.cust'));
    return o;
  }
  function navBtn(v, ic, label) { return '<button data-nav="' + v + '" class="' + (state.view === v ? 'on' : '') + '"><span class="ic">' + ic + '</span><span>' + esc(label) + '</span></button>'; }

  // ---------------- views ----------------
  function viewHead(title, sub) { return '<div class="viewhead"><div><h2>' + esc(title) + '</h2><p>' + esc(sub) + '</p></div></div>'; }
  function crumb(items) {
    var h = '<div class="crumb">';
    h += '<button data-nav="exec">' + T('drill.exec') + '</button><span class="sep">›</span><button data-nav="fin">' + T('drill.fin') + '</button>';
    items.forEach(function (it) { h += '<span class="sep">›</span>' + (it.nav ? '<button data-nav="' + it.nav + '"' + (it.obj ? ' data-obj="' + it.obj + '"' : '') + '>' + esc(it.label) + '</button>' : '<span>' + esc(it.label) + '</span>'); });
    return h + '</div>';
  }

  function objTableFull() {
    var h = '<div class="tablewrap"><table><thead><tr>' +
      '<th>' + T('tbl.obj') + '</th><th>' + T('tbl.name') + '</th><th class="r">' + T('tbl.scoreobj') + '</th><th>' + T('tbl.exc') + '</th><th>' + T('tbl.status') + '</th><th>' + T('tbl.drill') + '</th><th>' + T('tbl.focus') + '</th>' +
      '</tr></thead><tbody>';
    KO.forEach(function (o) {
      var st = objStatus(o), cc = (o.counts && o.counts.Critical) || 0, ac = (o.counts && o.counts.Attention) || 0;
      h += '<tr class="clickable" data-nav="obj" data-obj="' + o.code + '">' +
        '<td class="num">' + o.code + '</td><td>' + esc(o.name) + '</td>' +
        '<td class="r">' + scoreBar(o.avgScore) + '</td>' +
        '<td>' + (cc ? '<span class="kstat" style="color:var(--neg)">' + cc + '</span> ' : '') + (ac ? '<span class="kstat" style="color:var(--warn)">' + ac + '</span>' : (cc ? '' : '<span class="mut">0</span>')) + '</td>' +
        '<td>' + stPill(st) + '</td>' +
        '<td>' + esc(o.code) + '</td><td class="mut">' + T('review.only') + '</td></tr>';
    });
    return h + '</tbody></table></div>';
  }

  function scoreCard(lbl, val, sub, color) {
    return '<div class="card kpi"><div class="lbl">' + esc(lbl) + '</div><div class="val"' + (color ? ' style="color:' + color + '"' : '') + '>' + val + '</div>' + (sub ? '<div class="sub">' + sub + '</div>' : '') + '</div>';
  }

  function renderExec() {
    var sc = SUMMARY.overall;
    var scSt = sc == null ? 'Data Required' : sc >= 80 ? 'Attention' : 'Critical';
    if (sc != null && sc >= 100) scSt = 'On Target';
    var critObjs = KO.filter(function (o) { return objStatus(o) === 'Critical'; }).length;
    var attObjs = KO.filter(function (o) { return objStatus(o) === 'Attention'; }).length;
    var h = crumb([]) + viewHead(T('exec.title'), T('exec.sub'));
    h += '<div class="grid kpis" style="margin-bottom:16px">';
    h += scoreCard(T('score.overall'), (sc == null ? '—' : nf(sc, 1) + '<small> / 100</small>'), stPill(scSt), null);
    h += scoreCard(T('crit.obj'), critObjs, '', 'var(--neg)');
    h += scoreCard(T('att.obj'), attObjs, '', 'var(--warn)');
    h += scoreCard(T('computed'), SUMMARY.computed + '<small> / ' + SUMMARY.total + '</small>', '', 'var(--pos)');
    var dr = (SUMMARY.statusCounts && SUMMARY.statusCounts['Data Required']) || 0;
    h += scoreCard(T('datareq'), dr, '', 'var(--muted)');
    h += '</div>';
    h += objTableFull();
    h += '<div class="note"><b>' + T('exec.guide') + ':</b> ' + T('exec.guide.txt') + '</div>';
    h += '<div class="note"><b>' + T('exec.limit') + ':</b> ' + T('exec.limit.txt') + '</div>';
    return h;
  }

  function renderFin() {
    var sc = SUMMARY.overall, critObjs = 0, attObjs = 0;
    KO.forEach(function (o) { var s = objStatus(o); if (s === 'Critical') critObjs++; else if (s === 'Attention') attObjs++; });
    var h = crumb([]) + viewHead(T('fin.title'), T('fin.sub'));
    h += '<div class="grid kpis" style="margin-bottom:16px">';
    h += scoreCard(T('score.overall'), (sc == null ? '—' : nf(sc, 1) + '<small> / 100</small>'), '', null);
    h += scoreCard(T('crit.obj'), critObjs, '', 'var(--neg)');
    h += scoreCard(T('att.obj'), attObjs, '', 'var(--warn)');
    h += '</div>';
    h += '<div class="tablewrap"><table><thead><tr><th>' + T('tbl.obj') + '</th><th>' + T('tbl.name') + '</th><th class="r">' + T('tbl.scoreobj') + '</th><th>' + T('tbl.exc') + '</th><th>' + T('tbl.status') + '</th><th>' + T('drill.obj') + '</th><th>' + T('tbl.action') + '</th></tr></thead><tbody>';
    KO.forEach(function (o) {
      var st = objStatus(o), cc = (o.counts && o.counts.Critical) || 0, ac = (o.counts && o.counts.Attention) || 0;
      h += '<tr class="clickable" data-nav="obj" data-obj="' + o.code + '"><td class="num">' + o.code + '</td><td>' + esc(o.name) + '</td><td class="r">' + scoreBar(o.avgScore) + '</td>' +
        '<td>' + (cc ? '<span class="kstat" style="color:var(--neg)">' + cc + '</span> ' : '') + (ac ? '<span class="kstat" style="color:var(--warn)">' + ac + '</span>' : (cc ? '' : '<span class="mut">0</span>')) + '</td>' +
        '<td>' + stPill(st) + '</td><td>' + esc(o.code) + ' Dashboard</td><td class="mut">' + T('review.obj') + '</td></tr>';
    });
    h += '</tbody></table></div>';
    h += '<div class="note"><b>' + T('fin.priority') + ':</b> ' + T('fin.priority.txt') + '</div>';
    return h;
  }

  function renderObj() {
    var o = OBJ[state.obj]; if (!o) return renderExec();
    var h = crumb([{ nav: 'obj', obj: o.code, label: o.code + ' · ' + o.name }]);
    h += viewHead(o.code + ' | ' + o.name, T('avg.score') + ': ' + (o.avgScore == null ? '—' : nf(o.avgScore, 1)) + ' · ' + T('kpi.count') + ': ' + o.kpis.length);
    var cc = (o.counts && o.counts.Critical) || 0, ac = (o.counts && o.counts.Attention) || 0, drc = (o.counts && o.counts['Data Required']) || 0, pd = (o.counts && o.counts['Target Pending']) || 0;
    h += '<div class="grid kpis" style="margin-bottom:16px">';
    h += scoreCard(T('avg.score'), (o.avgScore == null ? '—' : nf(o.avgScore, 1) + '<small> / 100</small>'), stPill(objStatus(o)), null);
    h += scoreCard(T('st.crit'), cc, '', 'var(--neg)');
    h += scoreCard(T('st.att'), ac, '', 'var(--warn)');
    h += scoreCard(T('st.na'), drc, '', 'var(--muted)');
    h += scoreCard(T('st.pend'), pd, '', 'var(--muted2)');
    h += '</div>';
    h += '<div class="tablewrap"><table><thead><tr><th>' + T('tbl.kpicode') + '</th><th>' + T('tbl.kpi') + '</th><th class="r">' + T('tbl.actual') + '</th><th>' + T('tbl.score') + '</th><th>' + T('tbl.status') + '</th><th>' + T('tbl.target') + '</th><th>' + T('tbl.source') + '</th></tr></thead><tbody>';
    o.kpis.forEach(function (kp) {
      h += '<tr class="krow" data-kpi="' + kp.code + '"><td class="num">' + kp.code + '</td><td>' + esc(kp.name) + '</td><td class="r num">' + kpiVal(kp) + '</td><td>' + scoreBar(kp.score) + '</td><td>' + stPill(kp.status) + '</td><td class="num">' + esc(kpiTarget(kp)) + '</td><td class="mut">' + esc(kp.source) + '</td></tr>';
    });
    return h + '</tbody></table></div>';
  }

  function kpiDecision(kp) {
    if (kp.status === 'Critical') return state.lang === 'ar' ? 'إجراء فوري: القيمة خارج النطاق المستهدف — مراجعة السبب وخطة تصحيحية.' : 'Immediate action: value is outside the target band — review cause and corrective plan.';
    if (kp.status === 'Attention') return state.lang === 'ar' ? 'متابعة إدارية: اقتربت من الحد — راقب الاتجاه واتخذ إجراءً وقائيًا.' : 'Management attention: approaching the threshold — monitor and act preventively.';
    if (kp.status === 'On Target') return state.lang === 'ar' ? 'على الهدف — مراقبة دورية دون إجراء.' : 'On target — routine monitoring, no action.';
    if (kp.status === 'Data Required') return state.lang === 'ar' ? 'مشكلة جاهزية بيانات: وفّر المدخلات المطلوبة لتفعيل الحساب.' : 'Data readiness issue: supply the required inputs to activate calculation.';
    return state.lang === 'ar' ? 'بانتظار اعتماد الهدف الرقمي من الإدارة.' : 'Awaiting approval of the numeric target by management.';
  }

  function defRow(k, v) { return '<div class="defrow"><div class="dk">' + esc(k) + '</div><div class="dv">' + esc(v || '—') + '</div></div>'; }
  function chain(hilite) {
    var items = [T('chain.src'), T('nav.mapping'), T('chain.calc'), T('nav.master'), T('chain.dash'), T('f.objective'), T('nav.fin'), T('nav.exec')];
    return '<div class="chain">' + items.map(function (x, i) { return '<span class="cn' + (i === hilite ? ' on' : '') + '">' + esc(x) + '</span>'; }).join('<span class="ar">→</span>') + '</div>';
  }

  function renderMaster() {
    var f = state.f, list = [];
    KO.forEach(function (o) {
      if (f.mobj != null && o.code !== f.mobj) return;
      o.kpis.forEach(function (kp) { if (f.mstat == null || kp.status === f.mstat) list.push(kp); });
    });
    var h = crumb([{ label: T('nav.master') }]) + viewHead(T('master.title'), T('master.sub')) + filterBar();
    h += '<div class="tablewrap"><table><thead><tr><th>' + T('tbl.kpicode') + '</th><th>' + T('tbl.kpi') + '</th><th>' + T('tbl.obj') + '</th><th>' + T('def.formula') + '</th><th class="r">' + T('kpi.actual') + '</th><th>' + T('kpi.target') + '</th><th>' + T('tbl.score') + '</th><th>' + T('tbl.status') + '</th></tr></thead><tbody>';
    if (!list.length) h += '<tr><td colspan="8"><div class="empty">' + T('noRows') + '</div></td></tr>';
    list.forEach(function (kp) {
      h += '<tr class="krow" data-kpi="' + kp.code + '"><td class="num">' + kp.code + '</td><td>' + esc(kp.name) + '</td><td class="num mut">' + kp._obj + '</td><td class="mut">' + esc(kp.formula) + '</td><td class="r num">' + kpiVal(kp) + '</td><td class="num">' + esc(kpiTarget(kp)) + '</td><td>' + scoreBar(kp.score) + '</td><td>' + stPill(kp.status) + '</td></tr>';
    });
    h += '</tbody><tfoot><tr><td colspan="8">' + T('sum.filtered') + ' · ' + nf(list.length, 0) + ' / ' + SUMMARY.total + '</td></tr></tfoot></table></div>';
    return h;
  }

  function renderMapping() {
    var m = (D.kpi && D.kpi.metrics) || [], ready = 0;
    m.forEach(function (x) { if (x.ready) ready++; });
    var h = crumb([{ label: T('nav.mapping') }]) + viewHead(T('map.title'), T('map.sub'));
    h += '<div class="grid kpis" style="margin-bottom:16px">';
    h += scoreCard(T('lin.ok'), ready, '', 'var(--pos)');
    h += scoreCard(T('lin.no'), m.length - ready, '', 'var(--muted2)');
    h += scoreCard(T('sum.all'), m.length, '', null);
    h += '</div>';
    h += '<div class="tablewrap"><table><thead><tr><th>' + T('c.key') + '</th><th>' + T('c.source') + '</th><th>' + T('c.label') + '</th><th class="r">' + T('c.value') + '</th><th>' + T('c.avail') + '</th></tr></thead><tbody>';
    if (!m.length) h += '<tr><td colspan="5"><div class="empty">' + T('noRows') + '</div></td></tr>';
    m.forEach(function (x) {
      h += '<tr><td class="num">' + esc(x.key) + '</td><td class="mut">' + esc(x.source) + '</td><td>' + esc(x.label) + '</td><td class="r num">' + (x.value == null ? '—' : money(x.value)) + '</td><td>' + (x.ready ? '<span class="kstat" style="color:var(--pos)">' + T('lin.ok') + '</span>' : '<span class="kstat" style="color:var(--muted2)">' + T('lin.no') + '</span>') + '</td></tr>';
    });
    h += '</tbody><tfoot><tr><td colspan="5">' + T('sum.filtered') + ' · ' + nf(m.length, 0) + '</td></tr></tfoot></table></div>';
    h += chain(1);
    return h;
  }

  function renderKpi() {
    var kp = KPI[state.kpi]; if (!kp) return renderExec();
    var o = OBJ[kp._obj];
    var h = crumb([{ nav: 'obj', obj: kp._obj, label: kp._obj + ' · ' + kp._objName }, { label: kp.code }]);
    h += viewHead(kp.code + ' · ' + kp.name, kp._obj + ' | ' + kp._objName);
    h += '<div class="statgrid" style="margin-bottom:16px">';
    h += '<div class="stat"><div class="k">' + T('kpi.actual') + '</div><div class="v">' + kpiVal(kp) + '</div></div>';
    h += '<div class="stat"><div class="k">' + T('kpi.target') + '</div><div class="v sm">' + esc(kpiTarget(kp)) + '</div><div class="sub mut" style="font-size:11px;margin-top:4px">' + esc(kp.refStandard) + '</div></div>';
    h += '<div class="stat"><div class="k">' + T('kpi.score100') + '</div><div class="v">' + (kp.score == null ? '—' : nf(kp.score, 1)) + '</div></div>';
    h += '<div class="stat"><div class="k">' + T('kpi.status') + '</div><div class="v sm">' + stPill(kp.status) + '</div></div>';
    h += '</div>';
    // management view + gauge
    h += '<div class="grid cols2" style="margin-bottom:16px">';
    h += '<div class="card"><h3>' + T('sec.manage') + '</h3>' +
      '<div class="defrow" style="border:0;padding-top:0"><div class="dk">' + T('manage.decision') + '</div><div class="dv">' + esc(kpiDecision(kp)) + '</div></div>' +
      '<div class="defrow"><div class="dk">' + T('perf.readiness') + '</div><div class="dv">' + (kp.actual != null ? '<span class="pos">' + T('ready.ok') + '</span>' : '<span class="muted2">' + T('ready.missing') + '</span>') + '</div></div>' +
      '<div class="defrow" style="border-bottom:0"><div class="dk">' + T('def.purpose') + '</div><div class="dv">' + esc(kp.purpose) + '</div></div></div>';
    h += '<div class="card"><h3>' + T('sec.trend') + '</h3>' +
      (kp.score != null ? '<div id="gauge" class="gauge"></div>' : '') +
      '<div class="note" style="margin-top:' + (kp.score != null ? '4px' : '0') + '">' + T('trend.na') + '</div></div>';
    h += '</div>';
    // definition
    h += '<div class="card" style="margin-bottom:16px"><h3>' + T('sec.def') + '</h3><div class="defgrid">';
    h += defRow(T('def.formula'), kp.formula) + defRow(T('def.unit'), kp.unit) + defRow(T('def.ref'), kp.refStandard) + defRow(T('def.source'), kp.source);
    h += defRow(T('def.freq'), kp.frequency) + defRow(T('def.dir'), kp.direction) + defRow(T('def.owner'), kp.owner) + defRow(T('def.tier'), kp.tier);
    h += defRow(T('def.odoo'), kp.odooMap) + defRow(T('def.weight'), kp.weight == null ? '' : (kp.weight * 100).toFixed(1) + '%');
    h += '</div></div>';
    // lineage
    h += '<div class="card" style="margin-bottom:16px"><h3>' + T('sec.lineage') + '</h3>';
    if (kp.inputs && kp.inputs.length) {
      h += '<div class="tablewrap" style="max-height:none"><table><thead><tr><th>' + T('lin.key') + '</th><th>' + T('lin.src') + '</th><th class="r">' + T('lin.val') + '</th><th>' + T('lin.st') + '</th></tr></thead><tbody>';
      kp.inputs.forEach(function (inp) {
        h += '<tr><td class="num">' + esc(inp.key) + '</td><td class="mut">' + esc(inp.source) + '</td><td class="r num">' + (inp.value == null ? '—' : money(inp.value)) + '</td><td>' + (inp.ready ? '<span class="kstat" style="color:var(--pos)">' + T('lin.ok') + '</span>' : '<span class="kstat" style="color:var(--muted2)">' + T('lin.no') + '</span>') + '</td></tr>';
      });
      h += '</tbody></table></div>';
    } else h += '<div class="empty">' + T('lin.none') + '</div>';
    h += chain(4);
    h += '</div>';
    // drill-down
    h += '<div class="card"><h3>' + T('sec.drill') + '</h3><div class="chain">' +
      '<button class="btn" data-nav="obj" data-obj="' + kp._obj + '">◈ ' + T('drill.obj') + '</button>' +
      '<button class="btn" data-nav="fin">▦ ' + T('drill.fin') + '</button>' +
      '<button class="btn" data-nav="exec">▤ ' + T('drill.exec') + '</button>' +
      '<button class="btn primary" data-nav="tx" data-preset="1">≡ ' + T('open.tx') + '</button>' +
      '</div></div>';
    return h;
  }

  function renderTx() {
    var idxs = txFiltered();
    var per = 150, page = state.f.page || 1, pages = Math.max(1, Math.ceil(idxs.length / per)); if (page > pages) page = pages;
    var rows = D.tx.rows, m = rowMonths();
    var start = (page - 1) * per, end = Math.min(idxs.length, start + per);
    var tD = 0, tC = 0;
    for (var z = 0; z < idxs.length; z++) { tD += rows[idxs[z]][2]; tC += rows[idxs[z]][3]; }
    var h = crumb([]) + viewHead(T('exp.title'), T('exp.sub'));
    h += tabs('tx');
    h += filterBar();
    h += '<div class="tablewrap"><table><thead><tr><th>' + T('c.date') + '</th><th>' + T('c.acc') + '</th><th>' + T('c.accname') + '</th><th>' + T('c.desc') + '</th><th>' + T('c.partner') + '</th><th class="r">' + T('c.debit') + '</th><th class="r">' + T('c.credit') + '</th></tr></thead><tbody>';
    if (!idxs.length) h += '<tr><td colspan="7"><div class="empty">' + T('noRows') + '</div></td></tr>';
    for (var i = start; i < end; i++) {
      var r = rows[idxs[i]], a = TXACC[r[0]];
      h += '<tr><td class="num">' + rowISO(r[1]) + '</td><td class="num">' + esc(a ? a.code : '') + '</td><td>' + esc(a ? a.name : '') + '</td><td class="mut">' + esc(D.tx.desc[r[4]] || '') + '</td><td>' + esc(PART[r[5]] || '') + '</td><td class="r num">' + (r[2] ? nf(r[2], 2) : '') + '</td><td class="r num">' + (r[3] ? nf(r[3], 2) : '') + '</td></tr>';
    }
    h += '</tbody><tfoot><tr><td colspan="5">' + T('sum.filtered') + ' · ' + nf(idxs.length, 0) + ' ' + T('rows') + '</td><td class="r num">' + nf(tD, 2) + '</td><td class="r num">' + nf(tC, 2) + '</td></tr></tfoot></table></div>';
    h += pager(page, pages, idxs.length);
    return h;
  }
  function pager(page, pages, total) {
    return '<div class="pager"><span class="pg">' + T('rows') + ': ' + nf(total, 0) + ' · ' + T('page') + ' ' + page + ' / ' + pages + '</span>' +
      '<button class="btn sm" data-pg="' + (page - 1) + '"' + (page <= 1 ? ' disabled' : '') + '>‹ ' + T('prev') + '</button>' +
      '<button class="btn sm" data-pg="' + (page + 1) + '"' + (page >= pages ? ' disabled' : '') + '>' + T('next') + ' ›</button></div>';
  }
  function tabs(active) {
    var t = [['tx', T('tab.tx')], ['proj', T('tab.proj')], ['acc', T('tab.acc')], ['cust', T('tab.cust')]];
    return '<div class="seg" style="margin-bottom:14px">' + t.map(function (x) { return '<button data-nav="' + x[0] + '" class="' + (active === x[0] ? 'on' : '') + '">' + esc(x[1]) + '</button>'; }).join('') + '</div>';
  }

  function renderProj() {
    var f = state.f, list = D.projects;
    if (f.customer != null && PART[f.customer]) { var cn = PART[f.customer]; list = list.filter(function (p) { return p.customer === cn; }); }
    var tD = 0, tC = 0, tB = 0;
    list.forEach(function (p) { tD += p.debit || 0; tC += p.credit || 0; tB += p.balance || 0; });
    var h = crumb([]) + viewHead(T('exp.title'), T('exp.sub')) + tabs('proj') + filterBar();
    h += '<div class="tablewrap"><table><thead><tr><th>#</th><th>' + T('c.name') + '</th><th>' + T('c.ref') + '</th><th>' + T('c.partner') + '</th><th>' + T('c.plan') + '</th><th class="r">' + T('c.debit') + '</th><th class="r">' + T('c.credit') + '</th><th class="r">' + T('c.bal') + '</th></tr></thead><tbody>';
    if (!list.length) h += '<tr><td colspan="8"><div class="empty">' + T('noRows') + '</div></td></tr>';
    list.forEach(function (p, i) {
      h += '<tr><td class="num mut">' + (i + 1) + '</td><td>' + esc(p.name) + '</td><td class="num mut">' + esc(p.reference) + '</td><td>' + esc(p.customer) + '</td><td class="mut">' + esc(p.plan) + '</td><td class="r num">' + (p.debit ? nf(p.debit, 2) : '') + '</td><td class="r num">' + (p.credit ? nf(p.credit, 2) : '') + '</td><td class="r num">' + nf(p.balance, 2) + '</td></tr>';
    });
    h += '</tbody><tfoot><tr><td colspan="5">' + T('sum.filtered') + ' · ' + nf(list.length, 0) + ' ' + T('obj.proj') + '</td><td class="r num">' + nf(tD, 2) + '</td><td class="r num">' + nf(tC, 2) + '</td><td class="r num">' + nf(tB, 2) + '</td></tr></tfoot></table></div>';
    return h;
  }

  function renderAcc() {
    var f = state.f, list = accountsList(), cnt = accCounts();
    if (f.type != null) list = list.filter(function (a) { return a.type === f.type; });
    var tD = 0, tC = 0, tB = 0, tN = 0;
    list.forEach(function (a) { tD += a.debit || 0; tC += a.credit || 0; tB += a.balance || 0; tN += cnt[a.code] || 0; });
    var h = crumb([]) + viewHead(T('exp.title'), T('exp.sub')) + tabs('acc') + filterBar();
    h += '<div class="tablewrap"><table><thead><tr><th>#</th><th>' + T('c.code') + '</th><th>' + T('c.name') + '</th><th>' + T('c.type') + '</th><th class="r">' + T('c.moves') + '</th><th class="r">' + T('c.debit') + '</th><th class="r">' + T('c.credit') + '</th><th class="r">' + T('c.bal') + '</th></tr></thead><tbody>';
    list.forEach(function (a, i) {
      var mv = cnt[a.code] || 0;
      h += '<tr' + (mv ? ' class="clickable"' : '') + ' data-accfilter="' + esc(a.code) + '" data-moves="' + mv + '"><td class="num mut">' + (i + 1) + '</td><td class="num">' + esc(a.code) + '</td><td>' + esc(a.name) + '</td><td class="mut">' + esc(a.type) + '</td><td class="r num">' + (mv ? nf(mv, 0) : '<span class="mut">—</span>') + '</td><td class="r num">' + (a.debit ? nf(a.debit, 2) : '') + '</td><td class="r num">' + (a.credit ? nf(a.credit, 2) : '') + '</td><td class="r num">' + (a.balance ? nf(a.balance, 2) : '') + '</td></tr>';
    });
    h += '</tbody><tfoot><tr><td colspan="5">' + T('sum.filtered') + ' · ' + nf(list.length, 0) + ' ' + T('tab.acc') + ' · ' + nf(tN, 0) + ' ' + T('c.moves') + '</td><td class="r num">' + nf(tD, 2) + '</td><td class="r num">' + nf(tC, 2) + '</td><td class="r num">' + nf(tB, 2) + '</td></tr></tfoot></table></div>';
    return h;
  }

  function renderCust() {
    var idxs = txFiltered(), rows = D.tx.rows;
    var map = {};
    for (var z = 0; z < idxs.length; z++) {
      var r = rows[idxs[z]], pi = r[5];
      if (!map[pi]) map[pi] = { i: pi, name: custName(pi), n: 0, dr: 0, cr: 0 };
      map[pi].n++; map[pi].dr += r[2]; map[pi].cr += r[3];
    }
    var list = Object.keys(map).map(function (k) { return map[k]; });
    list.sort(function (a, b) { return (b.dr + b.cr) - (a.dr + a.cr); });
    var tD = 0, tC = 0;
    list.forEach(function (x) { tD += x.dr; tC += x.cr; });
    var h = crumb([]) + viewHead(T('exp.title'), T('exp.sub')) + tabs('cust') + filterBar();
    h += '<div class="tablewrap"><table><thead><tr><th>#</th><th>' + T('c.name') + '</th><th class="r">' + T('c.count') + '</th><th class="r">' + T('c.revenue') + '</th><th class="r">' + T('c.spend') + '</th><th class="r">' + T('c.net') + '</th></tr></thead><tbody>';
    if (!list.length) h += '<tr><td colspan="6"><div class="empty">' + T('noRows') + '</div></td></tr>';
    list.forEach(function (x, i) {
      h += '<tr class="clickable" data-custfilter="' + x.i + '"><td class="num mut">' + (i + 1) + '</td><td>' + esc(x.name) + '</td><td class="r num">' + nf(x.n, 0) + '</td><td class="r num">' + nf(x.dr, 2) + '</td><td class="r num">' + nf(x.cr, 2) + '</td><td class="r num">' + nf(x.dr - x.cr, 2) + '</td></tr>';
    });
    h += '</tbody><tfoot><tr><td colspan="3">' + T('sum.filtered') + ' · ' + nf(list.length, 0) + '</td><td class="r num">' + nf(tD, 2) + '</td><td class="r num">' + nf(tC, 2) + '</td><td class="r num">' + nf(tD - tC, 2) + '</td></tr></tfoot></table></div>';
    return h;
  }

  // ---------------- render ----------------
  function render() {
    document.documentElement.setAttribute('dir', state.lang === 'ar' ? 'rtl' : 'ltr');
    document.documentElement.setAttribute('lang', state.lang);
    document.documentElement.setAttribute('data-theme', state.theme);
    $('#brandTitle').textContent = T('app.title');
    $('#brandSub').textContent = T('app.sub');
    $('#uploadLbl').textContent = T('upload');
    $('#sbTitle').textContent = T('menu');
    $('#upTitle').textContent = T('up.title');
    $('#upDesc').textContent = T('up.desc');
    $('#upPrompt').textContent = T('up.prompt');
    document.title = T('app.title') + ' — ' + T('app.sub');
    $('#nav').innerHTML = nav();
    var html;
    switch (state.view) {
      case 'fin': html = renderFin(); break;
      case 'obj': html = renderObj(); break;
      case 'kpi': html = renderKpi(); break;
      case 'master': html = renderMaster(); break;
      case 'mapping': html = renderMapping(); break;
      case 'tx': html = renderTx(); break;
      case 'proj': html = renderProj(); break;
      case 'acc': html = renderAcc(); break;
      case 'cust': html = renderCust(); break;
      default: html = renderExec();
    }
    $('#viewHost').innerHTML = html;
    if (state.view === 'kpi') drawGauge();
  }

  function drawGauge() {
    var node = $('#gauge'); if (!node || !window.echarts) return;
    var kp = KPI[state.kpi]; if (!kp || kp.score == null) return;
    var col = kp.score >= 100 ? '#22d3a7' : kp.score >= 80 ? '#f5b544' : '#ef4d6b';
    try {
      var ch = echarts.init(node);
      ch.setOption({
        series: [{
          type: 'gauge', min: 0, max: 100, startAngle: 210, endAngle: -30, radius: '96%', center: ['50%', '62%'],
          progress: { show: true, width: 14, itemStyle: { color: col } },
          axisLine: { lineStyle: { width: 14, color: [[1, 'rgba(128,140,170,.22)']] } },
          axisTick: { show: false }, splitLine: { show: false }, axisLabel: { show: false }, pointer: { show: false },
          anchor: { show: false },
          detail: { valueAnimation: true, offsetCenter: [0, '-4%'], fontSize: 34, fontWeight: 800, color: 'inherit', formatter: function (v) { return v.toFixed(1); } },
          title: { offsetCenter: [0, '34%'], fontSize: 12, color: '#93a1c4' },
          data: [{ value: Math.round(kp.score * 10) / 10, name: T('gauge.score') }]
        }]
      });
      window.addEventListener('resize', function () { try { ch.resize(); } catch (e) {} });
    } catch (e) {}
  }

  // ---------------- events ----------------
  function setView(v) { state.view = v; state.f.page = 1; render(); closeSidebar(); }
  function setObj(code) { state.obj = code; state.view = 'obj'; state.f.page = 1; render(); closeSidebar(); }
  function setKpi(code) { state.kpi = code; state.view = 'kpi'; render(); }
  function closeSidebar() { $('#sidebar').classList.remove('open'); $('#sbBackdrop').classList.remove('on'); document.body.classList.remove('sb-lock'); }

  document.addEventListener('click', function (e) {
    var t = e.target;
    // nav buttons
    var nav = t.closest('[data-nav]');
    if (nav) {
      var v = nav.getAttribute('data-nav');
      if (v === 'obj') setObj(nav.getAttribute('data-obj'));
      else if (v === 'tx' && nav.getAttribute('data-preset')) { state.f.period = 'all'; state.f.account = null; state.f.customer = null; setView('tx'); }
      else setView(v);
      return;
    }
    var krow = t.closest('[data-kpi]'); if (krow) { setKpi(krow.getAttribute('data-kpi')); return; }
    var accf = t.closest('[data-accfilter]'); if (accf) { state.f.account = accf.getAttribute('data-accfilter') || null; setView('tx'); return; }
    var custf = t.closest('[data-custfilter]'); if (custf) { state.f.customer = parseInt(custf.getAttribute('data-custfilter'), 10); setView('tx'); return; }
    // filter dropdown toggle
    var fbtn = t.closest('.fdrop-btn'); if (fbtn) { var d = fbtn.parentNode; var open = d.classList.contains('open'); document.querySelectorAll('.fdrop.open').forEach(function (x) { x.classList.remove('open'); }); if (!open) d.classList.add('open'); return; }
    var opt = t.closest('.popt'); if (opt) {
      var key = opt.getAttribute('data-opt').split('|')[0], idx = parseInt(opt.getAttribute('data-opt').split('|')[1], 10);
      state.f[key] = (FOPT[key] && FOPT[key][idx]) ? FOPT[key][idx].v : null;
      state.f.page = 1;
      document.querySelectorAll('.fdrop.open').forEach(function (x) { x.classList.remove('open'); });
      render(); return;
    }
    var chp = t.closest('[data-chp]'); if (chp) { var kk = chp.getAttribute('data-chp'); state.f[kk] = (kk === 'period') ? 'all' : null; state.f.page = 1; render(); return; }
    if (t.id === 'fReset') { state.f.period = 'all'; state.f.account = null; state.f.customer = null; state.f.type = null; state.f.mobj = null; state.f.mstat = null; state.f.page = 1; render(); return; }
    var pg = t.closest('[data-pg]'); if (pg && !pg.disabled) { state.f.page = parseInt(pg.getAttribute('data-pg'), 10); render(); return; }
    // close open dropdowns on outside click
    if (!t.closest('.fdrop')) document.querySelectorAll('.fdrop.open').forEach(function (x) { x.classList.remove('open'); });
    // language / theme
    var lb = t.closest('button[data-lang]'); if (lb) { state.lang = lb.getAttribute('data-lang'); try { localStorage.setItem('dtk.lang', state.lang); } catch (e2) {} document.querySelectorAll('button[data-lang]').forEach(function (b) { b.classList.toggle('on', b.getAttribute('data-lang') === state.lang); }); render(); return; }
    var tb = t.closest('button[data-theme]'); if (tb) { state.theme = tb.getAttribute('data-theme'); try { localStorage.setItem('dtk.theme', state.theme); } catch (e3) {} document.querySelectorAll('button[data-theme]').forEach(function (b) { b.classList.toggle('on', b.getAttribute('data-theme') === state.theme); }); render(); return; }
    if (t.id === 'menuBtn') { $('#sidebar').classList.add('open'); $('#sbBackdrop').classList.add('on'); document.body.classList.add('sb-lock'); return; }
    if (t.id === 'menuClose' || t.id === 'sbBackdrop') { closeSidebar(); return; }
    if (t.id === 'uploadBtn') { openUpload(); return; }
    if (t.id === 'upCloseX') { $('#uploadModal').classList.remove('on'); return; }
  });

  // ---------------- upload ----------------
  function openUpload() { $('#uploadModal').classList.add('on'); }
  function handleFile(file) {
    if (!file || !window.XLSX || !window.FinPipeline) return;
    var st = $('#upStatus'); st.innerHTML = '<span class="spin"></span> ' + T('up.reading');
    var fr = new FileReader();
    fr.onload = function () {
      try {
        var wb = XLSX.read(new Uint8Array(fr.result), { type: 'array', cellDates: false });
        var model = window.FinPipeline.buildModel(XLSX, wb, { sourceFile: file.name, descCap: 70 });
        D = model; reindex();
        st.innerHTML = '<span class="pos">✓ ' + T('up.ok') + '</span>';
        state.view = 'exec'; state.obj = null; state.kpi = null; state.f = { period: 'all', account: null, customer: null, type: null, mobj: null, mstat: null, page: 1 };
        render();
      } catch (e) { st.innerHTML = '<span class="neg">✕ ' + T('up.err') + ': ' + esc(e.message) + '</span>'; }
    };
    fr.readAsArrayBuffer(file);
  }
  function reindex() {
    OBJ = {}; KPI = {}; KO = [];
    (D.kpi ? D.kpi.objectives : []).forEach(function (o) { OBJ[o.code] = o; KO.push(o); o.kpis.forEach(function (k) { KPI[k.code] = k; k._obj = o.code; k._objName = o.name; }); });
    SUMMARY = D.kpi ? D.kpi.summary : { overall: null, computed: 0, total: 0, statusCounts: {} };
    ACCB = {}; D.accounts.forEach(function (a) { ACCB[a.code] = a; });
    TXACC = (D.tx.accCodes || []).map(function (c) { var a = ACCB[c]; return { code: c, name: a ? a.name : c, type: a ? a.type : '' }; });
    PART = D.tx.partner || []; MONTHS = (D.monthly && D.monthly.months) || [];
    ROWM = null; CUSTOMERS = null; TYPES = null; ACCLIST = null; ACCN = null; UNASSIGNED = null;
  }
  function initUpload() {
    var dz = $('#dropzone'), fi = $('#fileInput');
    dz.addEventListener('click', function () { fi.click(); });
    fi.addEventListener('change', function () { handleFile(fi.files[0]); });
    ['dragenter', 'dragover'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('drag'); }); });
    ['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('drag'); }); });
    dz.addEventListener('drop', function (e) { if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });
  }

  // ---------------- boot ----------------
  function boot() {
    document.querySelectorAll('button[data-lang]').forEach(function (b) { b.classList.toggle('on', b.getAttribute('data-lang') === state.lang); });
    document.querySelectorAll('button[data-theme]').forEach(function (b) { b.classList.toggle('on', b.getAttribute('data-theme') === state.theme); });
    initUpload();
    render();
  }
  if (!D) { $('#viewHost').innerHTML = '<div class="empty">No data. Run build_data.js.</div>'; return; }
  boot();
})();

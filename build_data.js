const fs = require('fs');
const path = require('path');
const HERE = __dirname;
const XLSX = require(HERE + '/assets/xlsx.full.min.js');
const P = require(HERE + '/pipeline.js');
const OB = require(HERE + '/ops_build.js');

const SRC = process.argv[2] || path.join(HERE, '..', 'Finance_KPI_Master_Framework_F01-F10_Odoo_Dashboard_v10_DASHBOARDS_COMPLETE.xlsx');
const OUT = process.argv[3] || path.join(HERE, 'data.js');
const OPS_DIR = process.argv[4] || path.join(HERE, '..');
const SKIP_OPS = /^(1|true|yes)$/i.test(process.argv[5] || '');

const t0 = Date.now();
const buf = fs.readFileSync(SRC);
const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
console.log('read in', Date.now() - t0, 'ms; sheets', wb.SheetNames.length);

let opsPack = { ops: null, extra: {} };
if (!SKIP_OPS) {
  try {
    opsPack = OB.buildOps(XLSX, OPS_DIR, fs, path);
    console.log('ops built:',
      Object.keys(opsPack.ops).join(', '),
      '| extra metrics:', Object.keys(opsPack.extra).length);
  } catch (e) {
    console.error('OPS BUILD FAILED:', e.message);
    throw e;
  }
}

const model = P.buildModel(XLSX, wb, {
  sourceFile: path.basename(SRC),
  descCap: 70,
  extra: opsPack.extra,
  progress: function (stage, a, b) { if (stage !== 'gl') console.log('[pipeline]', stage, a, b || ''); }
});
if (opsPack.ops) model.ops = opsPack.ops;

const json = JSON.stringify(model);
fs.writeFileSync(OUT, 'window.DASH_DATA = ' + json + ';\n', 'utf8');

console.log('WROTE', OUT);
console.log('size(MB)', (fs.statSync(OUT).size / 1048576).toFixed(2));
console.log('tx', model.tx.rows.length, 'accounts', model.accounts.length, 'projects', model.projects.length, 'kpi obj', model.kpi ? model.kpi.objectives.length : 0);
if (model.kpi) {
  const s = model.kpi.summary;
  console.log('KPI computed', s.computed, '/', s.total, '| scored', s.scored, '| overall', s.overall, '|', JSON.stringify(s.statusCounts));
}
if (opsPack.ops) {
  console.log('ops.ar', opsPack.ops.ar.count, 'inv', opsPack.ops.ar.invoiced, 'ord', opsPack.ops.ar.overdue);
  console.log('ops.sales', opsPack.ops.sales.count, opsPack.ops.sales.total);
  console.log('ops.cc', opsPack.ops.cc.count, opsPack.ops.cc.revenue, opsPack.ops.cc.cost, 'loss', opsPack.ops.cc.lossCount);
}


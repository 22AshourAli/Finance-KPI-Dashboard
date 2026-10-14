const fs = require('fs');
const path = require('path');
const HERE = __dirname;
const XLSX = require(HERE + '/assets/xlsx.full.min.js');
const P = require(HERE + '/pipeline.js');

const SRC = process.argv[2] || path.join(HERE, '..', 'Finance_KPI_Master_Framework_F01-F10_Odoo_Dashboard_v10_DASHBOARDS_COMPLETE.xlsx');
const OUT = process.argv[3] || path.join(HERE, 'data.js');

const t0 = Date.now();
const buf = fs.readFileSync(SRC);
const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
console.log('read in', Date.now() - t0, 'ms; sheets', wb.SheetNames.length);

const model = P.buildModel(XLSX, wb, {
  sourceFile: path.basename(SRC),
  descCap: 70,
  progress: function (stage, a, b) { if (stage !== 'gl') console.log('[pipeline]', stage, a, b || ''); }
});

const json = JSON.stringify(model);
fs.writeFileSync(OUT, 'window.DASH_DATA = ' + json + ';\n', 'utf8');

console.log('WROTE', OUT);
console.log('size(MB)', (fs.statSync(OUT).size / 1048576).toFixed(2));
console.log('tx', model.tx.rows.length, 'accounts', model.accounts.length, 'projects', model.projects.length, 'kpi obj', model.kpi ? model.kpi.objectives.length : 0);

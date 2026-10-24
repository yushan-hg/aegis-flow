import { createRequire } from 'node:module';
import fs from 'node:fs';
import * as core from '../core.js';
const require = createRequire(import.meta.url);
const XLSX = require('../vendor/xlsx.full.min.js');
const rd = p => XLSX.read(fs.readFileSync(p), { type: 'buffer' });
const wb = rd(process.argv[2]);
const hsd = core.parseHsd(XLSX, wb);
const vw = rd(process.argv[3]);
const vitri = core.parseVitri(core.sheetRows(XLSX, vw.Sheets[vw.SheetNames[0]]));
const list = core.buildHsd(hsd, vitri);
console.log('xe có HSD:', list.length, '| tổng dòng:', list.reduce((s, x) => s + x.count, 0));
for (const x of list.slice(0, 3)) {
  console.log(x.xe, x.count, 'dòng,', x.dons.length, 'đơn');
  for (const d of x.dons) for (const r of d.rows.slice(0, 3)) console.log('   ', d.don, r.loc, r.ma, r.ten, r.qty, r.hsd);
}

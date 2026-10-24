// Số liệu mong đợi sau khi bỏ dòng có HSD (dùng để đối chiếu với app Windows)
import { createRequire } from 'node:module';
import fs from 'node:fs';
import * as core from '../core.js';
const require = createRequire(import.meta.url);
const XLSX = require('../vendor/xlsx.full.min.js');
const wb = XLSX.read(fs.readFileSync(process.argv[2]), { type: 'buffer' });
const main = core.parseMain(core.sheetRows(XLSX, wb.Sheets['Data']));
const hsd = core.parseHsd(XLSX, wb).filter(h => h.hsd);
const left = new Map();
for (const h of hsd) { const k = [h.xe, h.don, h.ma].join('|'); left.set(k, (left.get(k) || 0) + h.qty); }
const out = [];
for (const r of main) {
  const k = [r.xe, r.don, r.ma].join('|'); const q = left.get(k) || 0;
  if (q <= 0) { out.push(r); continue; }
  const take = Math.min(q, r.qty); left.set(k, q - take);
  if (r.qty - take > 0) out.push({ ...r, qty: r.qty - take, pallet: r.pallet * (r.qty - take) / r.qty });
}
const sum = (a, f) => a.reduce((s, x) => s + f(x), 0);
console.log('trước: dòng', main.length, 'xe', new Set(main.map(r => r.xe)).size, 'qty', sum(main, r => r.qty));
console.log('sau  : dòng', out.length, 'xe', new Set(out.map(r => r.xe)).size, 'qty', sum(out, r => r.qty), 'đơn', new Set(out.map(r => r.don)).size);
const byXe = new Map();
for (const r of out) byXe.set(r.xe, (byXe.get(r.xe) || 0) + r.qty);
console.log('xe KES-5216 (toàn hàng Pchome) còn?', byXe.has('KES-5216'));

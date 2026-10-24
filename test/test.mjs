// Kiểm thử logic đọc file: so với số liệu app Windows hiển thị cho cùng file mẫu.
//   node test/test.mjs <file dữ liệu .xlsx> <vi_tri_kho_config.xlsx> <doi_ban_config.xlsx>
import { createRequire } from 'node:module';
import fs from 'node:fs';
import * as core from '../core.js';

const require = createRequire(import.meta.url);
const XLSX = require('../vendor/xlsx.full.min.js');
const [dataPath, vitriPath, doibanPath] = process.argv.slice(2);
const read = p => XLSX.read(fs.readFileSync(p), { type: 'buffer' });

const wb = read(dataPath);
console.log('sheets:', wb.SheetNames, '→ detectKind:', core.detectKind(XLSX, wb));
const dataSheet = wb.Sheets['Data'] ?? wb.Sheets[wb.SheetNames[0]];
const main = core.parseMain(core.sheetRows(XLSX, dataSheet));
const sum = (a, f) => a.reduce((s, x) => s + f(x), 0);
console.log('main rows', main.length, '| đơn', new Set(main.map(r => r.don)).size, '| xe', new Set(main.map(r => r.xe)).size,
            '| qty', sum(main, r => r.qty), '| 板位數', sum(main, r => r.pallet).toFixed(2));
console.log('  (app Windows: 1903 dòng, 100 đơn, 23 xe, 12111, 160.95)');

const hsd = core.parseHsd(XLSX, wb);
console.log('hsd rows', hsd.length);

const vwb = read(vitriPath), dwb = read(doibanPath);
console.log('detect:', core.detectKind(XLSX, vwb), core.detectKind(XLSX, dwb));
const vitri = core.parseVitri(core.sheetRows(XLSX, vwb.Sheets[vwb.SheetNames[0]]));
const doiBan = core.parseDoiBan(core.sheetRows(XLSX, dwb.Sheets[dwb.SheetNames[0]]));
console.log('vitri byMa', Object.keys(vitri.byMa).length, '| có 整板', Object.values(vitri.byMa).filter(x => x.full).length,
            '| có 底高', Object.values(vitri.byMa).filter(x => x.per && x.layers).length, '| doiBan', Object.keys(doiBan).length);

const swap = core.buildSwap(main, hsd, doiBan);
console.log('Đổi bản dòng:', swap.length, '(app Windows: 10) | có Pchome?', swap.some(r => core.isPchome(r.khach)));
for (const r of swap.slice(0, 3)) console.log('  ', r.xe, r.khach, r.ten, r.qty, r.pallet, r.swap, r.hsd);

const items = core.buildPchome(main, hsd, vitri, {});
console.log('PCHome mặt hàng:', items.length, '| dòng nguồn 23 → tổng SL', sum(items, i => i.qty));
for (const it of items) {
  const c = core.calcItem(it);
  console.log(' ', it.ten.padEnd(7), String(it.loc).padEnd(8), String(it.qty).padStart(4),
              `整板=${c.full ? '' : ''}${core.effective(it).full}`, `底×高=${c.per}×${c.layers}`,
              it.lots.map(l => l.code + (it.lots.length > 1 ? '*' + l.qty : '')).join(' / '),
              '→ bốc', c.lots.map(l => l.out).join(' | '), '| lôi', c.lots.map(l => l.pull ?? '—').join(' | '),
              it.mismatch ? '⚠' : '');
}

// Ví dụ của người dùng
const ex = core.parseLots('270416HAOO*47/270625HAOO*140/270726HAOO*13', 200);
console.log('RBR9B lots:', JSON.stringify(ex));
const rbr = core.calcItem({ ten: 'RBR9B', lots: ex.lots, spec: { full: 60, per: 10, layers: 10 }, manual: null });
console.log('RBR9B 10×10, 整板 60 →', rbr.lots.map(l => `${l.code}: bốc ${l.out}, lôi ${l.pull}`).join(' | '));
const me5 = core.calcItem({ ten: 'ME5', lots: [{ key: 'x', code: '270505HAOO', qty: 782 }], spec: { full: 104, per: 10, layers: 10 }, manual: null });
console.log('ME5 782, 104/bản, 10×10 →', me5.lots[0].out, '| lôi', me5.lots[0].pull, '(mong đợi 7/82 và 8)');

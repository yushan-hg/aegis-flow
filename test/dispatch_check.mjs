// Kiểm tra phần chia đơn + thống kê: node test/dispatch_check.mjs <file dữ liệu .xlsx> <staff_dynamic_config_v3.json>
import { createRequire } from 'node:module';
import fs from 'node:fs';
import * as core from '../core.js';
import { Dispatch } from '../dispatch.js';
const require = createRequire(import.meta.url);
const XLSX = require('../vendor/xlsx.full.min.js');
const wb = XLSX.read(fs.readFileSync(process.argv[2]), { type: 'buffer' });
const main = core.parseMain(core.sheetRows(XLSX, wb.Sheets['Data']));
const hsd = core.parseHsd(XLSX, wb);
const noHsd = core.excludeHsd(main, hsd);
const stats = core.buildStats(noHsd);
const sum = (a, f) => a.reduce((s, x) => s + f(x), 0);
console.log('main', main.length, 'qty', sum(main, r => r.qty), '| noHsd', noHsd.length, 'qty', sum(noHsd, r => r.qty), 'đơn', new Set(noHsd.map(r => r.don)).size);
console.log('xe', stats.length, '| nhặt lẻ', sum(stats, s => s.nhatLe), '| nguyên', sum(stats, s => s.nguyen), '| flip', sum(stats, s => s.flip));
console.log(stats.slice(0, 5).map(s => `${s.xe}:${s.nhatLe}`).join('  '));
const staff = JSON.parse(fs.readFileSync(process.argv[3], 'utf8').replace(/^﻿/, '')).map(s => ({ name: s.name, weight: s.weight, active: s.active }));
const d = new Dispatch(stats, staff);
console.log('autoAll →', d.autoAll(), 'xe | worst', JSON.stringify(d.worst()));
for (const s of d.workers()) console.log(' ', s.name.padEnd(7), String(d.xeOf(s.name).length).padStart(2), 'xe', String(d.load(s.name)).padStart(5), d.pct(s.name) + '%', d.status(s.name));
// kéo tay một xe nặng cho người đang rảnh nhất rồi cân bằng
const first = d.workers()[0].name;
const xe = d.dispatchable[0].xe;
d.give(xe, first); console.log('give', xe, '→', first, d.pct(first) + '%', '| rebalance moved', d.rebalance(), '| undo', d.undo());
const e = new Dispatch(stats, staff);
console.log('autoPending', e.autoPending(), 'xe | lại lần 2:', e.autoPending(), '| undo', e.undo(), e.assignedCount);

// Đọc hai file cấu hình của app Windows (vi_tri_kho_config.xlsx, doi_ban_config.xlsx) và ghi ra docs/config.json
// để app điện thoại có sẵn, không phải nạp tay. Chạy: node tools/build_config.mjs [thư mục Config]
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as core from '../core.js';

const require = createRequire(import.meta.url);
const XLSX = require('../vendor/xlsx.full.min.js');
const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const cfgDir = process.argv[2] || path.join(root, '..', 'Config');

const read = name => {
  const p = path.join(cfgDir, name);
  if (!fs.existsSync(p)) { console.warn('Không thấy', p); return null; }
  const wb = XLSX.read(fs.readFileSync(p), { type: 'buffer' });
  return core.sheetRows(XLSX, wb.Sheets[wb.SheetNames[0]]);
};

const out = { builtAt: new Date().toISOString() };
const vt = read('vi_tri_kho_config.xlsx');
if (vt) { const v = core.parseVitri(vt); out.vitri = { filename: 'vi_tri_kho_config.xlsx', ...v }; console.log('vi_tri_kho:', Object.keys(v.byMa).length, 'mã hàng'); }
const db = read('doi_ban_config.xlsx');
if (db) { const d = core.parseDoiBan(db); out.doiBan = { filename: 'doi_ban_config.xlsx', dict: d }; console.log('doi_ban:', Object.keys(d).length, 'khách'); }

const lx = read('list_xe_config.xlsx');
if (lx) { out.listXe = core.parseListXe(lx); console.log('list_xe:', out.listXe.length, 'xe'); }
const sp = path.join(cfgDir, 'staff_dynamic_config_v3.json');
if (fs.existsSync(sp)) {
  const staff = JSON.parse(fs.readFileSync(sp, 'utf8').replace(/^﻿/, ''));
  out.staff = staff.map(s => ({ name: s.name, weight: s.weight || '100%', active: s.active !== false }));
  console.log('nhân sự:', out.staff.length, 'người');
}

fs.mkdirSync(path.join(root, 'docs'), { recursive: true });
fs.writeFileSync(path.join(root, 'docs', 'config.json'), JSON.stringify(out));
console.log('Đã ghi docs/config.json');

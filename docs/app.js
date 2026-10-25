import * as core from './core.js';
import { makeT, LANGS } from './i18n.js';

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const nf = n => Math.round(n).toLocaleString('en-US');

// ── Lưu trên máy (localStorage): dữ liệu ngày, cấu hình, ngôn ngữ, giao diện ──
const store = {
  get(k, d) { try { const v = localStorage.getItem('af.' + k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem('af.' + k, JSON.stringify(v)); return true; } catch { return false; } },
  del(k) { try { localStorage.removeItem('af.' + k); } catch {} },
};

const S = {
  tab: store.get('tab', 'home'), seg: store.get('seg', 'hsd'), find: 'line',
  data: store.get('data', null),            // { filename, loadedAt, main, hsd }
  doiBan: store.get('doiban', null), vitri: store.get('vitri', null), listXe: store.get('listxe', []),
  lang: store.get('lang', 'vi'), theme: store.get('theme', ''),
  q: '', hsdXe: '', dataXe: '',
};
// Bản cũ lưu tên tab khác (disp, more, swap, …): giá trị lạ thì về tab đầu, nếu không màn hình sẽ trống
if (!['home', 'work', 'find'].includes(S.tab)) S.tab = 'home';
if (!['hsd', 'pc'].includes(S.seg)) S.seg = 'hsd';
if (!['vi', 'zh', 'en'].includes(S.lang)) S.lang = 'vi';
if (!['', 'dark', 'light'].includes(S.theme)) S.theme = '';
let T = makeT(S.lang);
let D = null;          // dữ liệu suy ra từ file ngày
let bodyFn = () => '';
let tail = '';          // phần dưới danh sách (Trang chủ: ngôn ngữ, giao diện)

// ── Dữ liệu suy ra ──────────────────────────────────────────────────────────
const GAS = ['CO2F7', 'CO2F5', 'N2F13'];   // bình khí: cùng danh sách với app Windows

function derive() {
  D = null;
  if (!S.data) return;
  const { main, hsd } = S.data;
  const noHsd = core.excludeHsd(main, hsd);
  const rowsByXe = new Map();
  for (const r of noHsd) (rowsByXe.get(r.xe) ?? rowsByXe.set(r.xe, []).get(r.xe)).push(r);
  D = { main, hsd, noHsd, xeAll: [...new Set(main.map(r => r.xe))], xeList: [...rowsByXe.keys()].sort(), items: core.buildItems(noHsd) };
}
const vitriOr = () => S.vitri || { byMa: {}, byTen: {} };
const swapRows = () => (D && S.doiBan ? core.buildSwap(D.main, D.hsd, S.doiBan.dict) : []);
// HSD và PCHome tách riêng: dòng HSD của khách Pchome十五庫 chỉ hiện ở PCHome, không hiện ở HSD
const hsdList = () => {
  if (!D) return [];
  const pc = new Set(D.main.filter(r => core.isPchome(r.khach)).map(r => [r.xe, r.don, r.ma].join('\u0001')));
  return core.buildHsd(D.hsd.filter(h => !core.isPchome(h.khach) && !pc.has([h.xe, h.don, h.ma].join('\u0001'))), vitriOr());
};
const pcItems = () => (D ? core.buildPchome(D.main, D.hsd, vitriOr(), {}) : []);
const gasQty = ma => D.main.filter(r => !core.isPchome(r.khach) && String(r.ten).toUpperCase() === ma).reduce((s, r) => s + r.qty, 0);
const missingXe = () => {
  if (!D) return [];
  const got = new Set(D.xeAll.map(x => x.toLowerCase()));
  const seen = new Set();
  return S.listXe.filter(v => { const k = v.xe.toLowerCase(); if (got.has(k) || seen.has(k)) return false; seen.add(k); return true; }).map(v => v.xe).sort();
};

// ── Thông báo ngắn ──
let toastTimer;
function toast(msg, { err = false, ms = 5000 } = {}) {
  const el = $('#toast');
  el.className = 'toast' + (err ? ' err' : '');
  el.textContent = msg;
  el.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => (el.hidden = true), ms);
}

// ── Nạp file: tự nhận loại (dữ liệu ngày / vị trí kho / danh sách đổi bản) ──
async function importFiles(files) {
  const done = [], fail = [];
  for (const f of files) {
    try { done.push(importBuffer(await f.arrayBuffer(), f.name)); }
    catch (e) { fail.push(`${f.name}: ${e.message || e}`); }
  }
  render();
  if (fail.length) toast(fail.join('\n'), { err: true, ms: 9000 });
  else if (done.length) toast(done.join('\n'));
}

export function importBuffer(buf, name) {
  const wb = XLSX.read(buf, { type: 'array' });
  const kind = core.detectKind(XLSX, wb);
  if (kind === 'data') {
    const ws = wb.Sheets['Data'] ?? wb.Sheets[wb.SheetNames[0]];
    const main = core.parseMain(core.sheetRows(XLSX, ws));
    const hsd = core.parseHsd(XLSX, wb);
    S.data = { filename: name, loadedAt: Date.now(), main, hsd };
    if (!store.set('data', S.data)) toast(T('save_fail'), { err: true });
    S.hsdXe = S.dataXe = ''; S.q = '';
    derive();
    return T('imported_data', name, main.length, hsd.length);
  }
  if (kind === 'vitri') {
    const v = core.parseVitri(core.sheetRows(XLSX, wb.Sheets[wb.SheetNames[0]]));
    S.vitri = { filename: name, ...v, source: 'import' }; store.set('vitri', S.vitri);
    return T('imported_vitri', Object.keys(v.byMa).length);
  }
  if (kind === 'doiban') {
    const dict = core.parseDoiBan(core.sheetRows(XLSX, wb.Sheets[wb.SheetNames[0]]));
    S.doiBan = { filename: name, dict, source: 'import' }; store.set('doiban', S.doiBan);
    return T('imported_doiban', Object.keys(dict).length);
  }
  throw new Error(T('bad_file'));
}

// ── Cấu hình có sẵn trong app (config.json): làm mặc định; file nạp tay thì ghi đè ──
async function loadBundledConfig() {
  try {
    const r = await fetch('config.json', { cache: 'no-cache' });
    if (!r.ok) return false;
    const b = await r.json();
    let changed = false;
    if (b.vitri && S.vitri?.source !== 'import') { S.vitri = { ...b.vitri, source: 'bundled' }; store.set('vitri', S.vitri); changed = true; }
    if (b.doiBan && S.doiBan?.source !== 'import') { S.doiBan = { ...b.doiBan, source: 'bundled' }; store.set('doiban', S.doiBan); changed = true; }
    if (b.listXe) { S.listXe = b.listXe; store.set('listxe', S.listXe); changed = true; }
    return changed;
  } catch { return false; }
}

// ── Mảnh giao diện dùng chung ──
const search = ph => `<input class="search" id="q" type="search" placeholder="${esc(ph)}" value="${esc(S.q)}" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" enterkeyhint="search">`;
const seg = (attr, list, cur) => `<div class="seg">${list.map(([k, l, n]) => `<button data-${attr}="${k}" class="${cur === k ? 'on' : ''}">${esc(l)}${n != null ? `<em>${n}</em>` : ''}</button>`).join('')}</div>`;
const chips = (attr, list, cur) => `<div class="chips">${list.map(([k, l, sm]) => `<button class="chip${cur === k ? ' on' : ''}" data-${attr}="${esc(k)}">${esc(l)}${sm ? `<small>${esc(sm)}</small>` : ''}</button>`).join('')}</div>`;
const note = s => `<div class="note">${esc(s)}</div>`;
const empty = (a, b) => `<div class="empty"><b>${esc(a)}</b>${b ? esc(b) : ''}</div>`;
const hsdText = h => String(h || '').split(/[\/+]/).map(s => esc(s.trim())).filter(Boolean).join('<br>') || '—';
const xeShort = x => x.replace(/^[A-Z]+-/, '');
const sumLine = (...parts) => `<div class="tot">${parts.map(esc).join(' · ')}</div>`;
const heading = (t, right = '') => `<h2><span>${esc(t)}</span>${right ? `<em>${esc(right)}</em>` : ''}</h2>`;
const LIMIT = 150;

// ── Bảng phẳng: mỗi dòng một mặt hàng, tiêu đề cột dính trên đầu khi cuộn ──
// cols: [{ h tiêu đề, c lớp ('m' chữ mono · 'n' số căn phải · 'b' đậm · 's' chữ nhỏ · 'a' màu nhấn), f: dòng → html }]
const col = (k, c, f) => ({ h: T(k), c, f });
const tbl = (cols, rows) => `<div class="tw"><table><thead><tr>${cols.map(c => `<th class="${c.c}">${esc(c.h)}</th>`).join('')}</tr></thead><tbody>`
  + rows.map(r => `<tr>${cols.map(c => `<td class="${c.c}">${c.f(r)}</td>`).join('')}</tr>`).join('') + '</tbody></table></div>';
const cap = total => (total > LIMIT ? note(T('more_rows', total - LIMIT)) : '');

// ── Khung chung ──
function render() {
  document.documentElement.lang = S.lang;
  document.querySelectorAll('.nav button').forEach(b => {
    b.classList.toggle('on', b.dataset.t === S.tab);
    b.querySelector('span').textContent = T('t_' + b.dataset.t);
  });
  $('#title').textContent = T('t_' + S.tab);
  const sc = $('#screen'), y = sc.scrollTop;
  bodyFn = () => ''; tail = '';
  const head = { home: vHome, work: vWork, find: vFind }[S.tab]();
  sc.innerHTML = head + '<div id="res">' + bodyFn() + '</div>' + tail;
  sc.scrollTop = y;
}
function go(tab) { S.tab = tab; S.q = ''; store.set('tab', tab); render(); $('#screen').scrollTop = 0; }

// ── Tab Trang chủ: nạp file · bình khí · xe không giao hàng · đổi bản ──────
function vHome() {
  const file = `<section class="file"><div class="fi"><b>${esc(S.data ? S.data.filename : T('noFile'))}</b>${S.data ? `<small>${esc(T('file_lines', S.data.main.length, S.data.hsd.length))}</small>` : ''}</div>
    <label class="btn pri" for="fileInput">⤒ ${esc(S.data ? T('reload') : T('load'))}</label></section>`;
  const settings = heading(T('lang')) + seg('lg', LANGS, S.lang)
    + heading(T('theme')) + seg('th', [['', T('theme_auto')], ['dark', T('theme_dark')], ['light', T('theme_light')]], S.theme)
    + `<div class="cap">Aegis Flow Mobile · v11 · ${esc(T('offline'))}</div>`;
  tail = settings;
  if (!D) return file + empty(T('need_data'), T('need_data2'));

  const gas = `<div class="kpis">${GAS.map(ma => { const q = gasQty(ma); return `<div class="${q ? '' : 'z'}"><b>${nf(q)}</b><small>${esc(ma)}</small></div>`; }).join('')}</div>`;
  const miss = missingXe();
  const missBox = miss.length ? `<div class="tags">${miss.map(x => `<span>${esc(x)}</span>`).join('')}</div>` : note(T('none'));
  const rows = swapRows();
  bodyFn = () => {
    if (!S.doiBan) return empty(T('need_swap_cfg'));
    if (!rows.length) return empty(T('no_swap'));
    return tbl([col('c_khach', 's', r => esc(r.khach)), col('c_xe', 'm', r => esc(r.xe)), col('c_ten', 'b', r => esc(r.ten)), col('c_qty', 'n b', r => nf(r.qty)),
                col('c_swap', 'b a', r => esc(r.swap)), col('c_hsd', 'm', r => hsdText(r.hsd))], rows);
  };
  return file + heading(T('gas')) + gas + heading(T('no_deliver'), miss.length ? T('xe_n', miss.length) : '') + missBox
    + heading(T('swap'), rows.length ? `${rows.length} ${T('lines')}` : '');
}

// ── Tab Việc làm: HSD · PCHome ──────────────────────────────────────────────
function vWork() {
  if (!D) return empty(T('need_data'), T('need_data2'));
  const hl = hsdList(), pcs = pcItems();
  const head = seg('seg', [['hsd', T('seg_hsd'), hl.reduce((s, x) => s + x.count, 0)], ['pc', T('pchome'), pcs.length]], S.seg);
  return head + (S.seg === 'hsd' ? wHsd(hl) : wPc(pcs));
}

// HSD: các xe hôm nay có hàng phải nhặt theo hạn sử dụng
function wHsd(list) {
  if (!list.length) return empty(T('need_hsd'));
  const chipHead = chips('hx', [['', T('all')], ...list.map(x => [x.xe, xeShort(x.xe), `${x.count}`])], S.hsdXe);
  bodyFn = () => {
    const sel = S.hsdXe ? list.filter(x => x.xe === S.hsdXe) : list;
    const rows = sel.flatMap(x => x.dons.flatMap(d => d.rows.map(r => ({ ...r, xe: x.xe }))));
    return sumLine(T('xe_lines', sel.length, rows.length))
      + tbl([col('c_xe', 'm', r => esc(r.xe)), col('c_ten', 'b', r => esc(r.ten)), col('c_ma', 'm', r => esc(r.ma)),
             col('c_hsd', 'm', r => hsdText(r.hsd)), col('c_loc', 'm', r => esc(r.loc)), col('c_qty', 'n b', r => nf(r.qty))], rows);
  };
  return chipHead;
}

// PCHome: tên hàng · số lượng · hạn sử dụng · bản chẵn/lẻ, xếp số lượng từ lớn đến nhỏ (mỗi hạn dùng một dòng)
function wPc(items) {
  if (!items.length) return empty(T('no_pc'));
  bodyFn = () => {
    const rows = items.flatMap(it => { const c = core.calcItem(it); return c.lots.map(l => ({ it, c, l })); })
      .sort((a, b) => b.l.qty - a.l.qty || a.it.ten.localeCompare(b.it.ten));
    return tbl([
      col('c_ten', 'b', r => esc(r.it.ten) + (r.it.mismatch ? ` <span class="mark" title="${esc(T('pc_mismatch'))}">⚠</span>` : '') + (r.c.cap ? '' : ' <span class="mark">?</span>')),
      col('c_qty', 'n b', r => nf(r.l.qty)), col('c_hsd', 'm', r => esc(r.l.code || '—')), col('c_out', 'n b a', r => esc(r.l.out)),
    ], rows) + note(T('pc_note'));
  };
  return '';
}

// ── Tab Tra cứu: Dòng hàng · Tìm hàng ───────────────────────────────────────
function vFind() {
  if (!D) return empty(T('need_data'), T('need_data2'));
  let head = search(S.find === 'item' ? T('ph_item') : T('ph_search')) + seg('fd', [['line', T('seg_lines')], ['item', T('seg_item')]], S.find);
  if (S.find === 'line') head += chips('dx', [['', T('all')], ...D.xeList.map(x => [x, xeShort(x)])], S.dataXe);
  bodyFn = () => {
    const q = S.q.trim().toLowerCase();
    if (S.find === 'item') {
      // Hàng ở đầu, rồi Xe · Đơn · Số lượng
      const rows = D.items.filter(i => !q || (i.ten + ' ' + i.ma.join(' ')).toLowerCase().includes(q))
        .flatMap(i => i.rows.map(r => ({ ...r, ten: i.ten })));
      if (!rows.length) return empty(T('no_match'));
      return sumLine(`${rows.length} ${T('lines')}`, nf(rows.reduce((s, r) => s + r.qty, 0)))
        + tbl([col('c_ten', 'b', r => esc(r.ten)), col('c_xe', 'm', r => esc(r.xe)), col('c_don', 'm', r => esc(r.don)), col('c_qty', 'n b', r => nf(r.qty))], rows.slice(0, LIMIT))
        + cap(rows.length);
    }
    // Khách · Xe · Đơn · Mã · Hàng · Số lượng
    const rows = D.noHsd.filter(r => (!S.dataXe || r.xe === S.dataXe) && (!q || [r.khach, r.xe, r.don, r.ma, r.ten].join(' ').toLowerCase().includes(q)));
    if (!rows.length) return empty(T('no_match'));
    return sumLine(`${rows.length} ${T('lines')}`, nf(rows.reduce((s, r) => s + r.qty, 0)))
      + tbl([col('c_khach', 's', r => esc(r.khach)), col('c_xe', 'm', r => esc(r.xe)), col('c_don', 'm', r => esc(r.don)), col('c_ma', 'm', r => esc(r.ma)),
             col('c_ten', 'b', r => esc(r.ten)), col('c_qty', 'n b', r => nf(r.qty))], rows.slice(0, LIMIT))
      + cap(rows.length);
  };
  return head;
}

// ── Sự kiện ─────────────────────────────────────────────────────────────────
function applyTheme() {
  if (S.theme) document.documentElement.dataset.theme = S.theme; else delete document.documentElement.dataset.theme;
  document.querySelector('meta[name=theme-color]')?.setAttribute('content', S.theme === 'light' ? '#eef1f7' : '#0b0f17');
}

document.addEventListener('click', e => {
  const c = e.target.closest.bind(e.target);
  let x;
  if ((x = c('.nav button'))) return go(x.dataset.t);
  if ((x = c('[data-seg]'))) { S.seg = x.dataset.seg; store.set('seg', S.seg); return render(); }
  if ((x = c('[data-fd]'))) { S.find = x.dataset.fd; S.q = ''; return render(); }
  if ((x = c('[data-hx]'))) { S.hsdXe = x.dataset.hx; return render(); }
  if ((x = c('[data-dx]'))) { S.dataXe = x.dataset.dx; return render(); }
  if ((x = c('[data-lg]'))) { S.lang = x.dataset.lg; store.set('lang', S.lang); T = makeT(S.lang); return render(); }
  if ((x = c('[data-th]'))) { S.theme = x.dataset.th; store.set('theme', S.theme); applyTheme(); return render(); }
});

document.addEventListener('input', e => {
  if (e.target.id === 'q') { S.q = e.target.value; $('#res').innerHTML = bodyFn(); }
});

$('#fileInput').addEventListener('change', e => { const f = [...e.target.files]; e.target.value = ''; if (f.length) importFiles(f); });

// ── Ngoại tuyến ──
if ('serviceWorker' in navigator && (location.protocol === 'https:' || ['localhost', '127.0.0.1'].includes(location.hostname))) {
  // Có bản mới: tải lại trang một lần để không bị lẫn trang cũ với file mới (lần cài đầu tiên thì không cần)
  const hadController = !!navigator.serviceWorker.controller;
  let reloaded = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (hadController && !reloaded) { reloaded = true; location.reload(); } });
  navigator.serviceWorker.register('sw.js').catch(() => {});
}

window.AF = { importBuffer, render, S, core, loadBundledConfig, get D() { return D; } };   // dùng cho kiểm thử
applyTheme();
derive();
render();
loadBundledConfig().then(changed => { if (changed) render(); });

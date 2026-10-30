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
  ticks: store.get('ticks', {}),            // dấu ✓ ở Tìm hàng: { [khoá]: 1 }, xoá khi nạp file mới
  q: '', hsdXe: '', dataXe: '',
};
// Bản cũ lưu tên tab khác (disp, more, swap, …): giá trị lạ thì về tab đầu, nếu không màn hình sẽ trống
if (!['home', 'work', 'find', 'more'].includes(S.tab)) S.tab = 'home';
if (!['hsd', 'pc'].includes(S.seg)) S.seg = 'hsd';
if (!['vi', 'zh', 'en'].includes(S.lang)) S.lang = 'vi';
if (!['', 'dark', 'light'].includes(S.theme)) S.theme = '';
let T = makeT(S.lang);
let D = null;          // dữ liệu suy ra từ file ngày
let bodyFn = () => '';

// ── Dữ liệu suy ra ──────────────────────────────────────────────────────────
const GAS = ['CO2F7', 'CO2F5', 'N2F13'];   // bình khí: cùng danh sách với app Windows

function derive() {
  D = null;
  if (!S.data) return;
  const { main, hsd } = S.data;
  const noHsd = core.excludeHsd(main, hsd);
  const rowsByXe = new Map();
  for (const r of noHsd) (rowsByXe.get(r.xe) ?? rowsByXe.set(r.xe, []).get(r.xe)).push(r);
  const stats = core.buildStats(noHsd);
  D = { main, hsd, noHsd, xeAll: [...new Set(main.map(r => r.xe))], xeList: [...rowsByXe.keys()].sort(), items: core.buildItems(noHsd),
        total: noHsd.reduce((s, r) => s + r.qty, 0),            // số lượng tổng (đã trừ hàng HSD)
        pick: stats.reduce((s, v) => s + v.nhatLe, 0) };        // tải lẻ = nhặt lẻ = tổng trừ phần đổi bản (板位數 ≥ 0.65)
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
    S.data = { filename: name, loadedAt: Date.now(), main, hsd, dropped: main.dropped };
    if (!store.set('data', S.data)) toast(T('save_fail'), { err: true });
    S.hsdXe = S.dataXe = ''; S.q = ''; S.ticks = {}; store.set('ticks', S.ticks);
    derive();
    const d = main.dropped;
    return T('imported_data', name, main.length, hsd.length) + (d.xe + d.khach + d.neg ? '\n' + T('dropped', d.xe, d.khach, d.neg) : '');
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
// Loại bản (換板需求) của khách theo doi_ban_config; khách không có trong danh sách thì '—'
const loaiBan = khach => { const v = S.doiBan && khach ? core.matchPallet(S.doiBan.dict, khach) : ''; return v && v !== '######' ? v : '—'; };
const hsdText = h => String(h || '').split(/[\/+]/).map(s => esc(s.trim())).filter(Boolean).join('<br>') || '—';
const xeShort = x => x.replace(/^[A-Z]+-/, '');
const sumLine = (...parts) => `<div class="tot">${parts.map(esc).join(' · ')}</div>`;
const heading = (t, right = '') => `<h2><span>${esc(t)}</span>${right ? `<em>${esc(right)}</em>` : ''}</h2>`;
const LIMIT = 150;
const itemKey = r => 'i|' + r.ten + '|' + r.xe + '|' + r.don;   // dấu ✓ ở Tìm hàng

// ── Bảng phẳng: mỗi dòng một mặt hàng, tiêu đề cột dính trên đầu khi cuộn ──
// cols: [{ h tiêu đề, c lớp ('m' chữ mono · 'n' số · 'b' đậm · 's' chữ nhỏ · 'a' màu nhấn), f: dòng → html, g: cột gộp }]
// Cột gộp (g) nằm liền nhau ở bên trái: giá trị lặp ở các dòng liên tiếp (số xe, số đơn, khách…) chỉ hiện một lần, ô được gộp
// theo chiều dọc và căn giữa; ô con chỉ gộp trong phạm vi ô cha (cùng khách mới gộp xe, cùng xe mới gộp đơn).
const col = (k, c, f, g = false, a = null) => ({ h: T(k), c, f, g, a });   // a: dòng → thuộc tính của ô (vd data-tick)
function tbl(cols, rows, rowCls = () => '') {
  const span = cols.map(() => rows.map(() => 1));
  let path = rows.map(() => '');
  cols.forEach((c, ci) => {
    if (!c.g) return;
    path = path.map((p, i) => p + '\u0001' + c.f(rows[i]));
    for (let start = 0, i = 1; i <= rows.length; i++) {
      if (i < rows.length && path[i] === path[start]) continue;
      span[ci][start] = i - start;
      for (let k = start + 1; k < i; k++) span[ci][k] = 0;
      start = i;
    }
  });
  const first = cols.findIndex(c => c.g);
  const body = rows.map((r, ri) => `<tr class="${first >= 0 && span[first][ri] > 0 ? 'gs ' : ''}${rowCls(r)}">` + cols.map((c, ci) => {
    const n = span[ci][ri];
    return n === 0 ? '' : `<td class="${c.c}${c.g ? ' g' : ''}"${c.a ? ' ' + c.a(r) : ''}${n > 1 ? ` rowspan="${n}"` : ''}>${n > 1 ? `<div class="sk">${c.f(r)}</div>` : c.f(r)}</td>`;
  }).join('') + '</tr>').join('');
  return `<div class="tw"><table${first < 0 ? ' class="z"' : ''}><thead><tr>${cols.map(c => `<th class="${c.c}">${esc(c.h)}</th>`).join('')}</tr></thead><tbody>${body}</tbody></table></div>`;
}
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
  bodyFn = () => '';
  const head = { home: vHome, work: vWork, find: vFind, more: vMore }[S.tab]();
  sc.innerHTML = head + '<div id="res">' + bodyFn() + '</div>';
  sc.scrollTop = y;
}
function go(tab) { sureClear = false; S.tab = tab; S.q = ''; store.set('tab', tab); render(); $('#screen').scrollTop = 0; }

// ── Tab Trang chủ: nạp file · bình khí · xe không giao hàng · đổi bản ──────
function vHome() {
  const file = `<section class="file"><div class="fi"><b>${esc(S.data ? S.data.filename : T('noFile'))}</b>${S.data ? `<small>${esc(T('file_lines', S.data.main.length, S.data.hsd.length))}</small>` : ''}</div>
    <label class="btn pri" for="fileInput">⤒ ${esc(S.data ? T('reload') : T('load'))}</label></section>`;
  if (!D) return file + empty(T('need_data'), T('need_data2'));

  const gas = `<div class="kpis">${GAS.map(ma => { const q = gasQty(ma); return `<div class="${q ? '' : 'z'}"><b>${nf(q)}</b><small>${esc(ma)}</small></div>`; }).join('')}</div>`;
  const miss = missingXe();
  const missBox = miss.length ? `<div class="tags">${miss.map(x => `<span>${esc(x)}</span>`).join('')}</div>` : note(T('none'));
  const rows = swapRows();
  bodyFn = () => {
    if (!S.doiBan) return empty(T('need_swap_cfg'));
    if (!rows.length) return empty(T('no_swap'));
    return tbl([col('c_khach', 's', r => esc(r.khach), true), col('c_xe', 'm', r => esc(r.xe), true), col('c_ten', 'b', r => esc(r.ten)), col('c_qty', 'n b', r => nf(r.qty)),
                col('c_swap', 'b a', r => esc(r.swap)), col('c_hsd', 'm', r => hsdText(r.hsd))], rows);
  };
  const totals = `<div class="kpis big"><div><b>${nf(D.total)}</b><small>${esc(T('k_total'))}</small></div><div class="a"><b>${nf(D.pick)}</b><small>${esc(T('k_pick'))}</small></div></div>`;
  return file + totals + heading(T('gas')) + gas + heading(T('no_deliver'), miss.length ? T('xe_n', miss.length) : '') + missBox
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
    const rows = sel.flatMap(x => x.dons.flatMap(d => d.rows.map(r => ({ ...r, xe: x.xe, khach: r.khach || d.khach }))));
    return sumLine(T('xe_lines', sel.length, rows.length))
      + tbl([col('c_xe', 'm', r => esc(r.xe), true),
             col('c_khach_lb', 's', r => `<b>${esc(r.khach || '—')}</b><div class="lb">${esc(loaiBan(r.khach))}</div>`, true),
             col('c_loc', 'm', r => esc(r.loc)),
             col('c_ten', 'b', r => esc(r.ten) + `<div class="cd">${esc(r.ma)}</div>`),   // mã hàng nằm dưới tên hàng cho đỡ rộng
             col('c_hsd', 'm', r => hsdText(r.hsd)), col('c_qty', 'n b', r => nf(r.qty))], rows);
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
        + tbl([col('c_ten', 'b', r => esc(r.ten), true), col('c_xe', 'm', r => esc(r.xe), true), col('c_don', 'm', r => esc(r.don)), col('c_qty', 'n b', r => nf(r.qty)),
               { h: '✓', c: 'k', f: () => '<i>✓</i>', a: r => `data-tick="${esc(itemKey(r))}"` }], rows.slice(0, LIMIT), r => (S.ticks[itemKey(r)] ? 'done' : ''))
        + cap(rows.length);
    }
    // Khách · Xe · Đơn · Mã · Hàng · Số lượng
    const rows = D.noHsd.filter(r => (!S.dataXe || r.xe === S.dataXe) && (!q || [r.khach, r.xe, r.don, r.ma, r.ten].join(' ').toLowerCase().includes(q)));
    if (!rows.length) return empty(T('no_match'));
    return sumLine(`${rows.length} ${T('lines')}`, nf(rows.reduce((s, r) => s + r.qty, 0)))
      + tbl([col('c_khach', 's', r => esc(r.khach), true), col('c_xe', 'm', r => esc(r.xe), true), col('c_don', 'm', r => esc(r.don), true), col('c_ma', 'm', r => esc(r.ma)),
             col('c_ten', 'b', r => esc(r.ten)), col('c_qty', 'n b', r => nf(r.qty))], rows.slice(0, LIMIT))
      + cap(rows.length);
  };
  return head;
}

// ── Tab Thêm: những việc ngoài luồng (dữ liệu & cấu hình, ngôn ngữ, giao diện, xoá dữ liệu) ──
let sureClear = false;   // xoá dữ liệu phải bấm hai lần
function vMore() {
  const src = x => (x.source === 'import' ? T('src_import') : T('src_bundled'));
  const li = (t, d, st = '', no = false) => `<div class="li"><span><div class="t">${esc(t)}</div><div class="d">${esc(d)}</div></span>${st ? `<span class="st${no ? ' no' : ''}">${esc(st)}</span>` : ''}</div>`;
  const d = S.data, v = S.vitri, o = S.doiBan;
  const imported = v?.source === 'import' || o?.source === 'import';
  return heading(T('data_cfg'))
    + '<div class="card">'
    + li(T('data_today'), d ? `${d.filename} · ${T('lines_n', d.main.length)}${d.dropped && d.dropped.xe + d.dropped.khach + d.dropped.neg ? '\n' + T('dropped', d.dropped.xe, d.dropped.khach, d.dropped.neg) : ''}` : T('not_yet'), '', !d)
    + li(T('vitri'), v ? `${v.filename} · ${Object.keys(v.byMa).length}` : T('not_yet'), v ? src(v) : '')
    + li(T('doiban'), o ? `${o.filename} · ${Object.keys(o.dict).length}` : T('not_yet'), o ? src(o) : '')
    + '</div>'
    + (imported ? `<button class="btn wide" data-a="bundled">${esc(T('use_bundled'))}</button><div class="note">${esc(T('use_bundled_sub'))}</div>` : '')
    + heading(T('lang')) + seg('lg', LANGS, S.lang)
    + heading(T('theme')) + seg('th', [['', T('theme_auto')], ['dark', T('theme_dark')], ['light', T('theme_light')]], S.theme)
    + heading(T('clear_all'))
    + `<button class="btn wide warn" data-a="clearAll">${esc(sureClear ? T('sure') : T('clear_all_btn'))}</button><div class="note">${esc(T('clear_all_sub'))}</div>`
    + `<div class="cap">Aegis Flow Mobile · v19 · ${esc(T('offline'))}</div>`;
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
  if ((x = c('[data-a]'))) {
    if (x.dataset.a === 'bundled') {
      if (S.vitri?.source === 'import') { S.vitri = null; store.del('vitri'); }
      if (S.doiBan?.source === 'import') { S.doiBan = null; store.del('doiban'); }
      loadBundledConfig().then(() => { render(); toast(T('bundled_used')); });
      return;
    }
    if (!sureClear) { sureClear = true; return render(); }
    sureClear = false;
    ['data', 'vitri', 'doiban', 'listxe', 'ticks'].forEach(store.del);
    Object.assign(S, { data: null, vitri: null, doiBan: null, listXe: [], ticks: {}, hsdXe: '', dataXe: '', q: '' });
    derive(); render(); loadBundledConfig().then(() => render());
    return toast(T('cleared'));
  }
  if ((x = c('[data-tick]'))) {
    const k = x.dataset.tick;
    if (S.ticks[k]) delete S.ticks[k]; else S.ticks[k] = 1;
    store.set('ticks', S.ticks);
    $('#res').innerHTML = bodyFn();
    return;
  }
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

import * as core from './core.js';
import { Dispatch, OVER } from './dispatch.js';
import { makeT, LANGS } from './i18n.js';

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const nf = n => Math.round(n).toLocaleString('en-US');

// ── Lưu trên máy (localStorage): dữ liệu ngày, cấu hình, dấu ✓, nhập tay, ngôn ngữ. Kết quả chia đơn KHÔNG lưu. ──
const store = {
  get(k, d) { try { const v = localStorage.getItem('af.' + k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem('af.' + k, JSON.stringify(v)); return true; } catch { return false; } },
  del(k) { try { localStorage.removeItem('af.' + k); } catch {} },
};

const S = {
  tab: store.get('tab', 'home'), seg: store.get('seg', 'hsd'), find: 'line',
  data: store.get('data', null),            // { filename, loadedAt, main, hsd }
  doiBan: store.get('doiban', null), vitri: store.get('vitri', null), listXe: store.get('listxe', []),
  manual: store.get('manual', {}),          // nhập tay PCHome: { [tên hàng]: { per, layers, full } }
  ticks: store.get('ticks', {}),            // đã xong: { [khoá]: 1 }
  lang: store.get('lang', 'vi'), theme: store.get('theme', ''),
  staff: [],                                // [{name, weight, active}] — lấy từ cấu hình có sẵn, chỉ đổi trong phiên
  q: '', swapFilter: 'all', hsdXe: '', ordXe: '', dataXe: '', pick: null,
  disp: null, sheet: null,
};
// Bản cũ lưu tên tab khác (swap, pchome, hsd, settings): giá trị lạ thì về tab đầu, nếu không màn hình sẽ trống
if (!['home', 'disp', 'work', 'find', 'more'].includes(S.tab)) S.tab = 'home';
if (!['hsd', 'swap', 'pc', 'ord'].includes(S.seg)) S.seg = 'hsd';
if (!['vi', 'zh', 'en'].includes(S.lang)) S.lang = 'vi';
let T = makeT(S.lang);
let D = null;          // dữ liệu suy ra từ file ngày
let bodyFn = () => '';

// ── Dữ liệu suy ra ──────────────────────────────────────────────────────────
function derive() {
  D = null;
  if (!S.data) { S.disp = null; return; }
  const { main, hsd } = S.data;
  const noHsd = core.excludeHsd(main, hsd);
  const stats = core.buildStats(noHsd);
  const byXe = new Map(stats.map(s => [s.xe, s]));
  const rowsByXe = new Map();
  for (const r of noHsd) (rowsByXe.get(r.xe) ?? rowsByXe.set(r.xe, []).get(r.xe)).push(r);
  D = {
    main, hsd, noHsd, stats, byXe, rowsByXe,
    xeAll: [...new Set(main.map(r => r.xe))].sort(),
    xeList: [...rowsByXe.keys()].sort(),
    items: core.buildItems(noHsd),
  };
  S.disp = new Dispatch(stats, S.staff);
}
const vitriOr = () => S.vitri || { byMa: {}, byTen: {} };
const swapRows = () => (D && S.doiBan ? core.buildSwap(D.main, D.hsd, S.doiBan.dict) : []);
const hsdList = () => (D ? core.buildHsd(D.hsd, vitriOr()) : []);
const pcItems = () => (D ? core.buildPchome(D.main, D.hsd, vitriOr(), S.manual) : []);
function pcTotals(items) {
  let full = 0, loose = 0, pull = 0, need = 0;
  for (const it of items) { const c = core.calcItem(it); full += c.nFull; loose += c.nLoose; pull += c.nPull; if (!c.complete) need++; }
  return { full, loose, pull, need };
}
const missingXe = () => {
  if (!D) return [];
  const got = new Set(D.xeAll.map(x => x.toLowerCase()));
  const seen = new Set();
  return S.listXe.filter(v => { const k = v.xe.toLowerCase(); if (got.has(k) || seen.has(k)) return false; seen.add(k); return true; }).map(v => v.xe).sort();
};

// ── Thông báo ngắn ──
let toastTimer;
function toast(msg, { err = false, undo = false, ms = 5000 } = {}) {
  const el = $('#toast');
  el.className = 'toast' + (err ? ' err' : '');
  el.innerHTML = `<span>${esc(msg)}</span>${undo ? `<button class="btn" data-a="undo">${esc(T('undo'))}</button>` : ''}`;
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
    S.ticks = {}; store.set('ticks', S.ticks);       // sang ngày mới thì xoá dấu ✓ cũ
    S.hsdXe = S.ordXe = S.dataXe = ''; S.pick = null;
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
    if (b.staff && !S.staff.length) { S.staff = b.staff.map(s => ({ ...s })); if (D) S.disp = new Dispatch(D.stats, S.staff); changed = true; }
    return changed;
  } catch { return false; }
}

// ── Mảnh giao diện dùng chung ──
const kpis = (cells, cls = '') => `<div class="kpis ${cls}">${cells.map(([v, l, c]) => `<div class="${c || ''}"><b>${v}</b><small>${esc(l)}</small></div>`).join('')}</div>`;
const pct = (d, t) => (t ? Math.round(100 * d / t) : 0);
const search = ph => `<input class="search" id="q" type="search" placeholder="${esc(ph)}" value="${esc(S.q)}" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" enterkeyhint="search">`;
const seg = (attr, list, cur) => `<div class="seg">${list.map(([k, l, n]) => `<button data-${attr}="${k}" class="${cur === k ? 'on' : ''}">${esc(l)}${n != null ? `<em>${n}</em>` : ''}</button>`).join('')}</div>`;
const chips = (attr, list, cur) => `<div class="chips">${list.map(([k, l, sm]) => `<button class="chip${cur === k ? ' on' : ''}" data-${attr}="${esc(k)}">${esc(l)}${sm ? `<small>${esc(sm)}</small>` : ''}</button>`).join('')}</div>`;
const needData = () => `<div class="empty"><b>${esc(T('need_data'))}</b>${esc(T('need_data2'))}</div>`;
const statusCls = st => ({ over: 'bad', good: 'ok', under: 'warn', idle: 'dim' }[st]);
const hsdText = h => esc(String(h || '').split(/[\/+]/).map(s => s.trim()).join('   ') || '—');

// ── Khung chung ──
function render(keepSheet = false) {
  const names = { home: 'home', disp: 'disp', work: 'work', find: 'find', more: 'more' };
  document.documentElement.lang = S.lang;
  document.querySelectorAll('.nav button').forEach(b => {
    b.classList.toggle('on', b.dataset.t === S.tab);
    b.querySelector('span').textContent = T('t_' + names[b.dataset.t]);
  });
  $('#title').textContent = T('t_' + S.tab);
  $('#fileName').textContent = S.data ? S.data.filename : T('noFile');
  const sc = $('#screen'), y = sc.scrollTop;
  bodyFn = () => '';
  const head = { home: vHome, disp: vDisp, work: vWork, find: vFind, more: vMore }[S.tab]();
  sc.innerHTML = head + '<div id="res">' + bodyFn() + '</div>';
  sc.scrollTop = y;
  const dot = document.querySelector('.nav .dot');
  if (dot) dot.hidden = !(D && pcTotals(pcItems()).need);
  if (S.sheet && !keepSheet) renderSheet();
}
function go(tab, seg) {
  S.tab = tab; if (seg) { S.seg = seg; store.set('seg', seg); }
  S.q = ''; S.pick = null; store.set('tab', tab);
  closeSheet(); render(); $('#screen').scrollTop = 0;
}

// ── Tab Hôm nay ─────────────────────────────────────────────────────────────
function vHome() {
  if (!D) return needData();
  const total = D.noHsd.reduce((s, r) => s + r.qty, 0), pick = D.stats.reduce((s, v) => s + v.nhatLe, 0);
  const dons = new Set(D.noHsd.map(r => r.don)).size;
  const pal = D.noHsd.filter(r => !core.isPchome(r.khach)).reduce((s, r) => s + r.pallet, 0);
  const hl = hsdList(), hTotal = hl.reduce((s, x) => s + x.count, 0);
  const hDone = hl.reduce((s, x) => s + x.dons.reduce((a, d) => a + d.rows.filter(r => S.ticks[hsdKey(x.xe, r)]).length, 0), 0);
  const sw = swapRows(), sDone = sw.filter(r => S.ticks[swapKey(r)]).length;
  const items = pcItems(), pt = pcTotals(items);
  const disp = S.disp, asg = disp.assignedCount, dsp = disp.dispatchable.length;
  const w = disp.worst();
  const wText = !w ? '—' : w === 'balanced' ? T('balanced') : `${w[0]} ${w[1] > 0 ? '+' : '−'}${Math.abs(w[1])}%`;
  const miss = missingXe();
  const tile = (go, title, v, sm, d, p) => `<button class="card tile" data-go="${go}"><span class="t">${esc(title)}</span><span class="v">${v}<small>${esc(sm)}</small></span><span class="d">${esc(d)}</span>${p != null ? `<div class="prog"><i style="width:${p}%"></i></div>` : ''}</button>`;
  return kpis([[nf(total), T('k_total')], [nf(pick), T('k_pick'), 'a']], 'big')
    + kpis([[D.xeAll.length, T('k_xe')], [dons, T('k_don')], [pal.toFixed(1), T('k_pal')]])
    + `<h2>${esc(T('progress'))}</h2>`
    + (hTotal ? tile('work:hsd', T('hsd_first'), `${hDone}/${hTotal}`, T('done_lines'), `${hl.length} xe · ${T('hsd_sub')}`, pct(hDone, hTotal)) : '')
    + tile('disp', T('t_disp'), `${asg}/${dsp}`, T('xe_given'), `${T('worst')} ${wText}`, pct(asg, dsp))
    + tile('work:swap', T('swap'), S.doiBan ? `${sDone}/${sw.length}` : '—', T('done_lines'), `${new Set(sw.map(r => r.xe)).size} xe`, S.doiBan ? pct(sDone, sw.length) : null)
    + (items.length ? tile('work:pc', T('pchome'), items.length, T('items'), `${T('pull_out')} ${pt.full}/${pt.loose} · ${T('pull_big')} ${pt.pull} ${T('pallets')}`, null) : '')
    + ((pt.need || miss.length) ? `<h2>${esc(T('attention'))}</h2>` : '')
    + (pt.need ? `<button class="alert" data-go="work:pc">⚠ ${esc(T('pc_miss_n', pt.need))}</button>` : '')
    + (miss.length ? `<button class="alert" data-go="more">⚠ ${esc(T('no_deliver', miss.length))}: ${esc(miss.join(', '))}</button>` : '');
}

// ── Tab Chia đơn ────────────────────────────────────────────────────────────
function vDisp() {
  if (!D) return needData();
  const d = S.disp, ws = d.workers();
  const avg = d.avgFactor();
  const sel = S.pick;
  const cards = ws.map(s => {
    const p = d.pct(s.name, 0, avg), st = d.status(s.name), c = statusCls(st), n = d.xeOf(s.name).length;
    return `<button class="sc${sel === s.name ? ' on' : ''}" data-st="${esc(s.name)}"><div class="n"><span>${esc(s.name)}</span><small>${esc(s.weight)}</small></div>
      <div class="p g-${c}">${p}%</div><div class="m"><i class="bg-${c}" style="width:${Math.min(100, p / 1.4)}%"></i></div>
      <small>${n} xe · ${nf(d.load(s.name))}</small></button>`;
  }).join('');
  const head = `<div class="staff">${cards}</div>
    <div class="act"><button class="btn pri flex" data-a="auto">${esc(T('auto'))}</button><button class="btn flex" data-a="rebal">${esc(T('rebal'))}</button><button class="btn" data-a="undo" ${d.canUndo ? '' : 'disabled'}>↶ ${esc(T('undo'))}</button></div>`;
  bodyFn = () => {
    const list = sel ? d.dispatchable.filter(v => d.owner.get(v.xe) === sel)
                     : [...d.dispatchable].sort((a, b) => (d.owner.has(a.xe) ? 1 : 0) - (d.owner.has(b.xe) ? 1 : 0) || b.nhatLe - a.nhatLe);
    return `<h2>${esc(sel ? T('xe_of', sel) : T('vehicles'))}<em>${esc(sel ? T('tap_change') : T('tap_give'))}</em></h2><div class="card">`
      + (list.length ? list.map(v => {
          const who = d.owner.get(v.xe);
          return `<button class="veh" data-v="${esc(v.xe)}"><b>${esc(v.xe)}</b><span class="tg ${who ? 'who' : 'free'}">${esc(who || T('unassigned'))}</span><span class="l">${nf(v.nhatLe)}</span>
            <span class="s">${T('col_nguyen')} ${nf(v.nguyen)}${v.flip ? ` · ${T('col_flip')} ${nf(v.flip)}` : ''}</span></button>`;
        }).join('') : `<div class="note">${esc(T('no_xe'))}</div>`) + '</div>';
  };
  return head;
}

function vehSheet() {
  const sh = S.sheet, d = S.disp, v = D.byXe.get(sh.xe);
  if (!v) return '';
  const who = d.owner.get(v.xe);
  let h = `<div class="grab"></div><h4><span class="mono">${esc(v.xe)}</span><span class="num">${nf(v.nhatLe)}</span></h4>
    <div class="sub">${T('col_nguyen')} ${nf(v.nguyen)} · ${T('col_flip')} ${nf(v.flip)} · ${T('col_pick')} ${nf(v.nhatLe)}${who ? ' · ' + esc(who) : ''}</div>
    <div style="height:8px"></div>${seg('sh', [['give', T('give')], ['detail', T('detail')]], sh.tab)}`;
  if (sh.tab === 'give') {
    const avg = d.avgFactor();
    h += `<div class="sub" style="margin-bottom:4px">${esc(T('give_to'))}</div>`
      + d.workers().filter(s => s.active || who === s.name).map(s => {
        const cur = who === s.name, p = d.pct(s.name, cur ? 0 : v.nhatLe, avg), c = p > OVER ? 'bad' : p < 85 ? 'warn' : 'ok';
        return `<button class="opt${cur ? ' cur' : ''}" data-give="${esc(s.name)}"><b>${esc(s.name)}${cur ? ' · ' + esc(T('current')) : ''}</b><span class="sub">${d.xeOf(s.name).length} xe · ${T('load')} ${nf(d.load(s.name))}</span><span class="pv g-${c}">${p}%</span></button>`;
      }).join('')
      + (who ? `<button class="opt" data-give=""><b class="g-warn">${esc(T('unassign'))}</b><span class="sub">${esc(T('unassign_sub'))}</span></button>` : '');
  } else {
    const rows = [...(D.rowsByXe.get(v.xe) || [])].sort((a, b) => a.don.localeCompare(b.don));
    const dons = new Set(rows.map(r => r.don)).size;
    const line = r => `<div class="row plain" style="padding-inline:0"><div class="mm" style="min-width:0"><div class="nm">${esc(r.ten)}${r.pallet >= core.PALLET_THRESHOLD ? `<span class="tag mark">板 ${r.pallet.toFixed(2)}</span>` : ''}</div>
      <div class="s">${sh.flat ? esc(r.don) + ' · ' + esc(r.khach) + ' · ' : ''}${esc(r.ma)}</div></div><div class="q">${nf(r.qty)}</div></div>`;
    h += `<div class="sub" style="margin-bottom:6px">${esc(T('orders_n', dons, rows.length))}</div>${seg('shf', [['0', T('by_order')], ['1', T('flat')]], sh.flat ? '1' : '0')}`;
    if (sh.flat) h += rows.map(line).join('');
    else {
      const by = new Map();
      for (const r of rows) (by.get(r.don) ?? by.set(r.don, []).get(r.don)).push(r);
      for (const [don, list] of by) h += `<div class="sum" style="margin:6px -14px 0;padding-inline:14px"><span>${esc(don)} · ${esc(list[0].khach)}</span><span class="num">${nf(list.reduce((s, r) => s + r.qty, 0))}</span></div>` + list.map(line).join('');
    }
  }
  return h;
}

// ── Tab Việc làm ────────────────────────────────────────────────────────────
const swapKey = r => 's|' + r.xe + '|' + r.don + '|' + r.ma;
const hsdKey = (xe, r) => 'h|' + xe + '|' + r.don + '|' + r.ma + '|' + r.hsd;

function vWork() {
  if (!D) return needData();
  const hl = hsdList(), sw = swapRows(), pcs = pcItems();
  const hCount = hl.reduce((s, x) => s + x.count, 0);
  const head = seg('seg', [['hsd', T('seg_hsd'), hCount], ['swap', T('swap'), sw.length], ['pc', T('pchome'), pcs.length], ['ord', T('orders')]], S.seg);
  if (S.seg === 'hsd') return head + wHsd(hl);
  if (S.seg === 'swap') return head + wSwap(sw);
  if (S.seg === 'pc') return head + wPc(pcs);
  return head + wOrders();
}

function wHsd(list) {
  if (!list.length) return `<div class="empty">${esc(T('need_hsd'))}</div>`;
  const chipHead = chips('hx', [['', T('all')], ...list.map(x => [x.xe, x.xe.replace(/^[A-Z]+-/, ''), `${x.count}`])], S.hsdXe);
  bodyFn = () => (S.hsdXe ? list.filter(x => x.xe === S.hsdXe) : list).map(x => {
    const rows = x.dons.flatMap(d => d.rows.map(r => ({ r, d })));
    const dn = rows.filter(({ r }) => S.ticks[hsdKey(x.xe, r)]).length;
    return `<section class="card grp"><h3><b>${esc(x.xe)}</b><span>${dn}/${rows.length}</span></h3>` + rows.map(({ r, d }) => {
      const k = hsdKey(x.xe, r);
      return `<button class="row${S.ticks[k] ? ' done' : ''}" data-tick="${esc(k)}"><span class="ck">✓</span><span class="mm"><div class="nm">${esc(r.ten)}</div><div class="hs">${hsdText(r.hsd)}</div>
        <div class="s">${esc(r.loc)} · ${esc(r.ma)} · ${esc(d.don)}${d.khach ? ' · ' + esc(d.khach) : ''}</div></span><span class="q">${nf(r.qty)}</span></button>`;
    }).join('') + '</section>';
  }).join('');
  return chipHead;
}

function wSwap(all) {
  if (!S.doiBan) return `<div class="empty">${esc(T('need_swap_cfg'))}</div>`;
  const doneN = all.filter(r => S.ticks[swapKey(r)]).length;
  const head = kpis([[all.length, T('lines')], [new Set(all.map(r => r.xe)).size, 'xe'], [`${doneN}/${all.length}`, '✓', doneN === all.length && all.length ? 'g' : '']])
    + search(T('ph_search')) + seg('sf', [['all', T('all')], ['todo', T('todo')]], S.swapFilter);
  bodyFn = () => {
    if (!all.length) return `<div class="empty">${esc(T('no_swap'))}</div>`;
    const q = S.q.trim().toLowerCase();
    const rows = all.filter(r => (!q || [r.xe, r.khach, r.ten, r.ma, r.don].join(' ').toLowerCase().includes(q)) && (S.swapFilter !== 'todo' || !S.ticks[swapKey(r)]));
    if (!rows.length) return `<div class="empty">${esc(T('no_match'))}</div>`;
    const by = new Map();
    for (const r of rows) (by.get(r.xe) ?? by.set(r.xe, []).get(r.xe)).push(r);
    return [...by].map(([xe, list]) => {
      const dn = list.filter(r => S.ticks[swapKey(r)]).length;
      return `<section class="card grp"><h3><b>${esc(xe)}</b><span>${dn}/${list.length}</span></h3>` + list.map(r =>
        `<button class="row${S.ticks[swapKey(r)] ? ' done' : ''}" data-tick="${esc(swapKey(r))}"><span class="ck">✓</span><span class="mm"><div class="nm">${esc(r.ten)}<span class="tag">${esc(r.swap)}</span></div>
          ${r.hsd ? `<div class="hs">${hsdText(r.hsd)}</div>` : ''}<div class="s">${esc(r.khach)} · ${esc(r.don)} · ${r.pallet}板</div></span><span class="q">${nf(r.qty)}</span></button>`).join('') + '</section>';
    }).join('');
  };
  return head;
}

function wPc(items) {
  if (!items.length) return `<div class="empty">${esc(T('no_pc'))}</div>`;
  const t = pcTotals(items);
  const head = kpis([[`${t.full}/${t.loose}`, T('pc_out')], [t.pull, T('pc_pull'), 'a'], [t.need, T('pc_miss_k'), t.need ? 'w' : 'g']])
    + search(T('ph_search')) + seg('pf', [['all', T('all')], ['need', T('need')]], S.pcFilter || 'all');
  bodyFn = () => {
    const q = S.q.trim().toLowerCase();
    const shown = items.filter(it => (!q || (it.ten + ' ' + it.codes.join(' ')).toLowerCase().includes(q)) && ((S.pcFilter || 'all') !== 'need' || !core.calcItem(it).complete));
    if (!shown.length) return `<div class="empty">${esc(T('no_match'))}</div>`;
    return '<div class="card">' + shown.map(pcCard).join('') + '</div><div class="note">' + esc(T('pc_note')) + '</div>';
  };
  return head;
}

function pcCard(it) {
  const c = core.calcItem(it);
  const lots = c.lots.map(l => {
    const tk = 'p|' + it.ten + '|' + l.code, on = S.ticks[tk] ? ' done' : '';
    return `<span class="${on}" data-tick="${esc(tk)}">${esc(l.code || '—')}${c.lots.length > 1 ? '*' + l.qty : ''}</span><span class="r" data-tick="${esc(tk)}">${esc(l.out)}</span><span class="p" data-tick="${esc(tk)}">${l.pull ?? '—'}</span>`;
  }).join('');
  const allDone = c.lots.length && c.lots.every(l => S.ticks['p|' + it.ten + '|' + l.code]);
  return `<div class="pc${allDone ? ' done' : ''}" data-pc="${esc(it.ten)}"><div class="h"><b>${esc(it.ten)}${it.mismatch ? ` <span class="mark" title="${esc(T('pc_mismatch'))}">⚠</span>` : ''}</b><span class="num" style="font-weight:800;font-size:15px">${nf(it.qty)}</span></div>
    <div class="sub num">${esc(it.codes.join(','))}${c.complete ? ` · ${c.full} · ${c.per}×${c.layers}` : ''}</div><div class="lots">${lots}</div>${c.complete ? '' : `<div class="miss">${esc(T('pc_miss'))}</div>`}</div>`;
}

function pcSheet() {
  const it = pcItems().find(i => i.ten === S.sheet.ten);
  if (!it) return '';
  const c = core.calcItem(it);
  const box = (f, v, lab) => `<label>${esc(lab)}<input data-fld="${f}" class="${v > 0 ? '' : 'miss'}" type="text" inputmode="numeric" pattern="[0-9]*" maxlength="4" value="${v || ''}" placeholder="?"></label>`;
  return `<div class="grab"></div><h4><span>${esc(it.ten)}</span><span class="num">${nf(it.qty)}</span></h4><div class="sub num">${esc(it.codes.join(','))}</div>
    <div class="fields">${box('full', c.full, T('pc_full'))}${box('per', c.per, T('pc_base'))}${box('layers', c.layers, T('pc_high'))}</div>
    <div class="sub" id="pcCap">${pcCapText(c)}</div><div class="note">${esc(T('pc_hint'))}</div><div id="pcRes">${pcResHtml(c)}</div>
    <div class="act" style="margin-top:10px"><button class="btn pri flex" data-a="close">${esc(T('pc_close'))}</button></div>`;
}
const pcCapText = c => c.cap ? `= ${c.cap} ${T('pc_cap')}` : T('per_pallet');
const pcResHtml = c => c.lots.map(l => `<div class="row plain" style="padding-inline:0"><div class="mm"><div class="hs">${esc(l.code || '—')}${c.lots.length > 1 ? '*' + l.qty : ''}</div></div>
  <div class="q" style="font-size:15px">${esc(l.out)} <span class="g-dim">|</span> <span style="color:var(--acc)">${l.pull ?? '—'}</span></div></div>`).join('');

function wOrders() {
  if (!D.xeList.length) return '';
  if (!S.ordXe || !D.xeList.includes(S.ordXe)) S.ordXe = D.xeList[0];
  const head = chips('ox', D.xeList.map(x => [x, x.replace(/^[A-Z]+-/, '')]), S.ordXe);
  bodyFn = () => {
    const rows = D.rowsByXe.get(S.ordXe) || [];
    const by = new Map();
    for (const r of [...rows].sort((a, b) => a.don.localeCompare(b.don))) (by.get(r.don) ?? by.set(r.don, []).get(r.don)).push(r);
    return `<div class="sub" style="margin-bottom:6px"><b class="mono">${esc(S.ordXe)}</b> · ${esc(T('orders_n', by.size, rows.length))}</div>`
      + [...by].map(([don, list]) => `<section class="card grp"><h3><b>${esc(don)}</b><span>${esc(list[0].khach)} · ${nf(list.reduce((s, r) => s + r.qty, 0))}</span></h3>`
        + list.map(r => `<div class="row plain"><div class="mm" style="min-width:0"><div class="nm">${esc(r.ten)}${r.pallet >= core.PALLET_THRESHOLD ? `<span class="tag mark">板 ${r.pallet.toFixed(2)}</span>` : ''}</div><div class="s">${esc(r.ma)}</div></div><div class="q">${nf(r.qty)}</div></div>`).join('') + '</section>').join('');
  };
  return head;
}

// ── Tab Tra cứu ─────────────────────────────────────────────────────────────
function vFind() {
  if (!D) return needData();
  const ph = S.find === 'item' ? T('ph_item') : S.find === 'veh' ? T('ph_xe') : T('ph_search');
  let head = search(ph) + seg('fd', [['line', T('seg_lines')], ['item', T('seg_item')], ['veh', T('seg_veh')]], S.find);
  if (S.find === 'line') head += chips('dx', [['', T('all')], ...D.xeList.map(x => [x, x.replace(/^[A-Z]+-/, '')])], S.dataXe);
  bodyFn = () => {
    const q = S.q.trim().toLowerCase();
    if (S.find === 'item') {
      const it = D.items.filter(i => !q || (i.ten + ' ' + i.ma.join(' ')).toLowerCase().includes(q));
      if (!it.length) return `<div class="empty">${esc(T('no_match'))}</div>`;
      return it.slice(0, 60).map(i => `<section class="card grp"><div class="sum"><span>${esc(i.ten)} <span class="sub">${esc(i.ma.join(','))}</span></span><span class="num">${nf(i.qty)}</span></div>`
        + i.rows.map(r => `<button class="res" data-v="${esc(r.xe)}"><span class="a num">${esc(r.xe)}</span><span class="b">${esc(r.don)} · ${esc(r.khach)}</span><span class="c">${nf(r.qty)}</span></button>`).join('') + '</section>').join('')
        + (it.length > 60 ? `<div class="note">${esc(T('more_rows', it.length - 60))}</div>` : '') + `<div class="note">${esc(T('item_note'))}</div>`;
    }
    if (S.find === 'veh') {
      const list = D.stats.filter(v => !q || v.xe.toLowerCase().includes(q)).sort((a, b) => a.xe.localeCompare(b.xe));
      if (!list.length) return `<div class="empty">${esc(T('hsd_none_xe'))}</div>`;
      return '<div class="card">' + list.map(v => {
        const who = S.disp.owner.get(v.xe);
        return `<button class="res" data-v="${esc(v.xe)}" data-detail="1"><span class="a num">${esc(v.xe)}</span><span class="b">${who ? esc(who) : esc(T('unassigned'))} · ${T('col_nguyen')} ${nf(v.nguyen)} · ${T('col_flip')} ${nf(v.flip)}</span><span class="c">${nf(v.nhatLe)}</span></button>`;
      }).join('') + '</div>';
    }
    const rows = D.noHsd.filter(r => (!S.dataXe || r.xe === S.dataXe) && (!q || [r.xe, r.don, r.khach, r.ma, r.ten].join(' ').toLowerCase().includes(q)));
    if (!rows.length) return `<div class="empty">${esc(T('no_match'))}</div>`;
    return '<div class="card">' + rows.slice(0, 150).map(r => `<div class="res"><span class="a">${esc(r.ten)}</span><span class="b">${esc(r.xe)} · ${esc(r.don)} · ${esc(r.khach)} · ${esc(r.ma)}</span><span class="c">${nf(r.qty)}</span></div>`).join('') + '</div>'
      + (rows.length > 150 ? `<div class="note">${esc(T('more_rows', rows.length - 150))}</div>` : '');
  };
  return head;
}

// ── Tab Thêm ────────────────────────────────────────────────────────────────
function vMore() {
  const d = S.data, v = S.vitri, o = S.doiBan;
  const src = x => (x.source === 'import' ? T('src_import') : T('src_bundled'));
  const li = (i, t, ds, attrs = '', right = '<span class="chev">›</span>') => `<button class="li" ${attrs}><span class="i">${i}</span><span><div class="t">${esc(t)}</div><div class="d">${ds}</div></span>${right}</button>`;
  const miss = missingXe();
  const themeName = S.theme === 'dark' ? T('theme_dark') : S.theme === 'light' ? T('theme_light') : T('theme_auto');
  return `<h2>${esc(T('staff'))}</h2><div class="card list">${S.staff.map((s, i) =>
      `<div class="stf${s.active ? '' : ' off'}"><span class="i" style="width:30px;height:30px;border-radius:9px;background:var(--s2);display:grid;place-items:center">👤</span><span class="nm">${esc(s.name)}</span>
       <button class="wt" data-wt="${i}">${esc(s.weight)}</button><button class="sw${s.active ? ' on' : ''}" data-act="${i}" aria-label="${esc(s.name)}"></button></div>`).join('') || `<div class="note">—</div>`}</div>
    <div class="note">${esc(T('staff_note'))}</div>
    ${miss.length ? `<h2>${esc(T('missing_xe'))}</h2><div class="card" style="padding:10px 12px"><span class="mono">${esc(miss.join(' · '))}</span></div>` : ''}
    <h2>${esc(T('data_cfg'))}</h2><div class="card list">
      ${li('📄', T('data_today'), d ? `${esc(d.filename)} · ${esc(T('lines_n', d.main.length))} · ${esc(new Date(d.loadedAt).toLocaleString())}` : esc(T('not_yet')), 'data-a="pick"')}
      ${li('🗂', T('vitri'), v ? `${esc(v.filename)} · ${Object.keys(v.byMa).length} · ${esc(src(v))}` : esc(T('not_yet')), 'disabled', '')}
      ${li('🔁', T('doiban'), o ? `${esc(o.filename)} · ${Object.keys(o.dict).length} · ${esc(src(o))}` : esc(T('not_yet')), 'disabled', '')}
      ${(v?.source === 'import' || o?.source === 'import') ? li('↺', T('use_bundled'), esc(T('use_bundled_sub')), 'data-a="bundled"') : ''}
      ${li('✓', T('ticks'), String(Object.keys(S.ticks).length), 'data-a="clearTicks"', `<span class="btn">${esc(T('clear'))}</span>`)}
      ${li('▦', T('manual'), String(Object.keys(S.manual).length), 'data-a="clearManual"', `<span class="btn">${esc(T('clear'))}</span>`)}
      ${li('🗑', T('clear_all'), esc(T('clear_all_sub')), 'data-a="clearAll"', `<span class="btn warn" id="clearAllLbl">${esc(T('clear_all_btn'))}</span>`)}
    </div>
    <h2>${esc(T('lang'))}</h2><div class="card" style="padding:8px">${seg('lg', LANGS.map(([k, l]) => [k, l]), S.lang)}<div style="height:2px"></div></div>
    <h2>${esc(T('theme'))}</h2><div class="card list">${li('◐', T('theme'), esc(themeName), 'data-a="theme"')}</div>
    <div class="cap">Aegis Flow Mobile · v7 · ${esc(T('offline'))}</div>`;
}

// ── Bảng trượt từ dưới lên ──────────────────────────────────────────────────
function openSheet(s) { S.sheet = s; $('#toast').hidden = true; renderSheet(); $('#sheet').classList.add('on'); $('#scrim').classList.add('on'); }
function closeSheet() { S.sheet = null; $('#sheet').classList.remove('on'); $('#scrim').classList.remove('on'); }
function renderSheet() {
  if (!S.sheet) return;
  const el = $('#sheet'), y = el.scrollTop;
  el.innerHTML = S.sheet.type === 'veh' ? vehSheet() : pcSheet();
  el.scrollTop = y;
  if (!el.innerHTML) closeSheet();
}

// ── Sự kiện ─────────────────────────────────────────────────────────────────
function applyTheme() {
  if (S.theme) document.documentElement.dataset.theme = S.theme; else delete document.documentElement.dataset.theme;
  document.querySelector('meta[name=theme-color]')?.setAttribute('content', S.theme === 'light' ? '#eef1f7' : '#0b0f17');
}
const worstText = () => { const w = S.disp.worst(); return !w ? '—' : w === 'balanced' ? T('balanced') : `${w[0]} ${w[1] > 0 ? '+' : '−'}${Math.abs(w[1])}%`; };

document.addEventListener('click', e => {
  const c = e.target.closest.bind(e.target);
  let x;
  if ((x = c('.nav button'))) return go(x.dataset.t);
  if ((x = c('[data-go]'))) { const [a, b] = x.dataset.go.split(':'); return go(a, b); }
  if (e.target.id === 'scrim') return closeSheet();
  if ((x = c('[data-seg]'))) { S.seg = x.dataset.seg; S.q = ''; store.set('seg', S.seg); return render(); }
  if ((x = c('[data-fd]'))) { S.find = x.dataset.fd; S.q = ''; return render(); }
  if ((x = c('[data-sf]'))) { S.swapFilter = x.dataset.sf; return render(); }
  if ((x = c('[data-pf]')) && x.tagName === 'BUTTON') { S.pcFilter = x.dataset.pf; return render(); }
  if ((x = c('[data-hx]'))) { S.hsdXe = x.dataset.hx; return render(); }
  if ((x = c('[data-ox]'))) { S.ordXe = x.dataset.ox; return render(); }
  if ((x = c('[data-dx]'))) { S.dataXe = x.dataset.dx; return render(); }
  if ((x = c('[data-lg]'))) { S.lang = x.dataset.lg; store.set('lang', S.lang); T = makeT(S.lang); return render(); }
  if ((x = c('[data-sh]'))) { S.sheet.tab = x.dataset.sh; return renderSheet(); }
  if ((x = c('[data-shf]'))) { S.sheet.flat = x.dataset.shf === '1'; return renderSheet(); }
  if ((x = c('[data-st]'))) { S.pick = S.pick === x.dataset.st ? null : x.dataset.st; return render(); }
  if ((x = c('[data-v]'))) { return openSheet({ type: 'veh', xe: x.dataset.v, tab: x.dataset.detail ? 'detail' : 'give', flat: false }); }
  if ((x = c('[data-give]'))) {
    const xe = S.sheet.xe, name = x.dataset.give;
    S.disp.give(xe, name || null);
    toast(name ? T('given', xe, name) : T('unassigned_t', xe), { undo: true });
    closeSheet(); return render();
  }
  if ((x = c('[data-tick]'))) {
    const k = x.dataset.tick;
    if (S.ticks[k]) delete S.ticks[k]; else S.ticks[k] = 1;
    store.set('ticks', S.ticks);
    return render();
  }
  if ((x = c('[data-pc]'))) return openSheet({ type: 'pc', ten: x.dataset.pc });
  if ((x = c('[data-wt]'))) { const s = S.staff[+x.dataset.wt]; s.weight = core.WEIGHTS[(core.WEIGHTS.indexOf(s.weight) + 1) % core.WEIGHTS.length]; return render(); }
  if ((x = c('[data-act]'))) { const s = S.staff[+x.dataset.act]; s.active = !s.active; return render(); }
  if (!(x = c('[data-a]'))) return;
  switch (x.dataset.a) {
    case 'close': closeSheet(); return;
    case 'pick': $('#fileInput').click(); return;
    case 'undo': if (S.disp?.undo()) { toast(T('undone')); render(); } return;
    case 'auto': {
      const n = S.disp.autoPending();
      if (n === -1) return toast(T('no_staff'), { err: true });
      if (n === 0) return toast(T('auto_none'));
      render(); return toast(T('auto_done', n, worstText()), { undo: true });
    }
    case 'rebal': {
      const n = S.disp.rebalance();
      if (n === -1) return toast(T('rebal_none'));
      if (n === 0) return toast(T('rebal_stuck'), { err: true });
      render(); return toast(T('rebal_done', n, worstText()), { undo: true });
    }
    case 'theme': S.theme = S.theme === '' ? 'dark' : S.theme === 'dark' ? 'light' : ''; store.set('theme', S.theme); applyTheme(); return render();
    case 'bundled':
      if (S.vitri?.source === 'import') { S.vitri = null; store.del('vitri'); }
      if (S.doiBan?.source === 'import') { S.doiBan = null; store.del('doiban'); }
      loadBundledConfig().then(() => { render(); toast(T('bundled_used')); }); return;
    case 'clearTicks': S.ticks = {}; store.set('ticks', {}); return render();
    case 'clearManual': S.manual = {}; store.set('manual', {}); return render();
    case 'clearAll':
      if (!x.dataset.sure) { x.dataset.sure = 1; $('#clearAllLbl').textContent = T('sure'); return; }
      ['data', 'vitri', 'doiban', 'listxe', 'manual', 'ticks'].forEach(store.del);
      Object.assign(S, { data: null, vitri: null, doiBan: null, listXe: [], manual: {}, ticks: {}, staff: [] });
      derive(); loadBundledConfig().then(() => render()); render(); return;
  }
});

document.addEventListener('input', e => {
  const el = e.target;
  if (el.id === 'q') { S.q = el.value; $('#res').innerHTML = bodyFn(); return; }
  if (!el.dataset.fld || el.tagName !== 'INPUT' || S.sheet?.type !== 'pc') return;
  // Nhập thông số PCHome: ghi nhập tay (chỉ giá trị khác file cấu hình), cập nhật bảng kết quả — không dựng lại bảng để không mất con trỏ
  const key = S.sheet.ten, item = pcItems().find(i => i.ten === key); if (!item) return;
  el.value = el.value.replace(/\D/g, '');
  const get = f => parseInt($(`#sheet [data-fld="${f}"]`).value, 10) || 0;
  const per = get('per'), layers = get('layers'), full = get('full'), s = item.spec;
  const baseManual = per > 0 && layers > 0 && !(s.per === per && s.layers === layers);
  const fullManual = full > 0 && full !== s.full;
  if (baseManual || fullManual) S.manual[key] = { per: baseManual ? per : 0, layers: baseManual ? layers : 0, full: fullManual ? full : 0 };
  else delete S.manual[key];
  store.set('manual', S.manual);
  const c = core.calcItem({ ...item, manual: S.manual[key] || null });
  for (const [f, v] of [['full', full], ['per', per], ['layers', layers]]) $(`#sheet [data-fld="${f}"]`).classList.toggle('miss', !(v > 0));
  $('#pcCap').textContent = pcCapText(c);
  $('#pcRes').innerHTML = pcResHtml(c);
  render(true);
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

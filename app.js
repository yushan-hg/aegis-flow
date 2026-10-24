import * as core from './core.js';

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

// ── Lưu trên máy (localStorage) — mọi thứ nằm trong điện thoại, mở lại không cần mạng ──
const store = {
  get(k, d) { try { const v = localStorage.getItem('af.' + k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem('af.' + k, JSON.stringify(v)); return true; } catch { return false; } },
  del(k) { try { localStorage.removeItem('af.' + k); } catch {} },
};

const state = {
  tab: store.get('tab', 'swap'),
  data: store.get('data', null),            // { filename, loadedAt, main, hsd }
  doiBan: store.get('doiban', null),        // { filename, dict, source }
  vitri: store.get('vitri', null),          // { filename, byMa, byTen, source }
  manual: store.get('manual', {}),          // nhập tay PCHome: { [tên hàng]: { per, layers, full } }
  ticks: store.get('ticks', {}),            // đã xong: { [khoá]: 1 }
  q: '', filter: 'all',
  hsdXe: '',                                // xe đang xem ở tab In HSD ('' = chưa chọn xe)
  open: new Set(),                          // các mặt hàng PCHome đang mở hàng ô nhập (trong phiên)
};

// ── Thông báo ngắn ──
let toastTimer;
function toast(msg, err = false, ms = 5000) {
  const el = $('#toast');
  el.textContent = msg; el.className = 'toast' + (err ? ' err' : ''); el.hidden = false;
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
  if (fail.length) toast(fail.join('\n'), true, 9000);
  else if (done.length) toast(done.join('\n'));
}

export function importBuffer(buf, name) {
  const wb = XLSX.read(buf, { type: 'array' });
  const kind = core.detectKind(XLSX, wb);
  if (kind === 'data') {
    const ws = wb.Sheets['Data'] ?? wb.Sheets[wb.SheetNames[0]];
    const main = core.parseMain(core.sheetRows(XLSX, ws));
    const hsd = core.parseHsd(XLSX, wb);
    state.data = { filename: name, loadedAt: Date.now(), main, hsd };
    if (!store.set('data', state.data)) toast('Máy không lưu nổi dữ liệu — vẫn dùng được nhưng mở lại phải nạp lại.', true);
    state.ticks = {}; store.set('ticks', state.ticks);       // sang ngày mới thì xoá dấu ✓ cũ
    state.hsdXe = '';
    return `Dữ liệu: ${name} (${main.length} dòng, ${hsd.length} dòng HSD)`;
  }
  if (kind === 'vitri') {
    const v = core.parseVitri(core.sheetRows(XLSX, wb.Sheets[wb.SheetNames[0]]));
    state.vitri = { filename: name, ...v, source: 'import' }; store.set('vitri', state.vitri);
    return `Vị trí kho: ${Object.keys(v.byMa).length} mã hàng`;
  }
  if (kind === 'doiban') {
    const dict = core.parseDoiBan(core.sheetRows(XLSX, wb.Sheets[wb.SheetNames[0]]));
    state.doiBan = { filename: name, dict, source: 'import' }; store.set('doiban', state.doiBan);
    return `Danh sách đổi bản: ${Object.keys(dict).length} khách`;
  }
  throw new Error('không nhận ra loại file (cần file dữ liệu ngày, vi_tri_kho_config hoặc doi_ban_config)');
}

// ── Cấu hình có sẵn trong app (config.json): dùng làm mặc định; file nạp tay thì ghi đè ──
async function loadBundledConfig() {
  try {
    const r = await fetch('config.json', { cache: 'no-cache' });
    if (!r.ok) return false;
    const b = await r.json();
    let changed = false;
    if (b.vitri && state.vitri?.source !== 'import') { state.vitri = { ...b.vitri, source: 'bundled' }; store.set('vitri', state.vitri); changed = true; }
    if (b.doiBan && state.doiBan?.source !== 'import') { state.doiBan = { ...b.doiBan, source: 'bundled' }; store.set('doiban', state.doiBan); changed = true; }
    return changed;
  } catch { return false; }
}

// ── Mảnh giao diện dùng chung ──
const kpi = cells => `<div class="kpi">${cells.map(([v, l, c]) => `<div class="${c || ''}"><b>${v}</b><small>${l}</small></div>`).join('')}</div>`;
const pct = (d, t) => (t ? Math.round(100 * d / t) : 0);
const bar = (d, t) => `<div class="bar"><i style="width:${pct(d, t)}%"></i></div>`;
const tools = (placeholder, segs) => `<div class="tools"><input class="search" id="q" type="search" placeholder="${placeholder}" value="${esc(state.q)}" autocomplete="off" autocapitalize="off">
  ${segs ? `<div class="seg">${segs.map(([k, l]) => `<button data-f="${k}" class="${state.filter === k ? 'on' : ''}">${l}</button>`).join('')}</div>` : ''}</div>`;
const needData = what => `<div class="empty">Chưa có ${what}.<br>Bấm <b>Nạp file</b> ở góc trên rồi chọn file Excel.</div>`;

// ── Màn hình ──
function render() {
  document.querySelectorAll('.tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === state.tab));
  $('#fileName').textContent = state.data ? state.data.filename : 'Chưa nạp file';
  $('#view').innerHTML = state.tab === 'swap' ? renderSwap() : state.tab === 'pchome' ? renderPchome()
                       : state.tab === 'hsd' ? renderHsd() : renderSettings();
}

// ── Tab Đổi bản ──
const swapKey = r => 's|' + r.xe + '|' + r.don + '|' + r.ma;

function renderSwap() {
  if (!state.data) return needData('dữ liệu hôm nay');
  if (!state.doiBan) return needData('danh sách đổi bản (<b>doi_ban_config.xlsx</b>)');
  const all = core.buildSwap(state.data.main, state.data.hsd, state.doiBan.dict);
  const doneN = all.filter(r => state.ticks[swapKey(r)]).length;
  const q = state.q.trim().toLowerCase();
  const rows = all.filter(r => (!q || [r.xe, r.khach, r.ten, r.ma, r.don].join(' ').toLowerCase().includes(q))
                            && (state.filter !== 'todo' || !state.ticks[swapKey(r)]));
  let h = kpi([[all.length, 'dòng'], [new Set(all.map(r => r.xe)).size, 'xe'], [`${doneN}/${all.length}`, 'xong', doneN === all.length && all.length ? 'g' : '']])
        + tools('Tìm xe, khách, tên hàng…', [['all', 'Tất cả'], ['todo', 'Chưa xong']]);
  if (!rows.length) return h + `<div class="empty">${all.length ? 'Không có dòng nào khớp.' : 'Hôm nay không có hàng cần đổi bản.'}</div>`;
  const byXe = new Map();
  for (const r of rows) (byXe.get(r.xe) ?? byXe.set(r.xe, []).get(r.xe)).push(r);
  for (const [xe, list] of byXe) {
    const d = list.filter(r => state.ticks[swapKey(r)]).length;
    h += `<section class="grp"><h3><span class="x">${esc(xe)}</span><span class="g">${d}/${list.length}</span></h3>${bar(d, list.length)}`;
    for (const r of list) {
      const hsd = r.hsd ? `<div class="hs">${esc(r.hsd.split(/[\/+]/).join('  '))}</div>` : '';
      h += `<div class="r${state.ticks[swapKey(r)] ? ' done' : ''}" data-tick="${esc(swapKey(r))}"><span class="ck">✓</span>
        <div class="m"><div class="t"><b>${esc(r.ten)}</b><span class="tag">${esc(r.swap)}</span></div>
          <div class="s">${esc(r.khach)} · ${esc(r.don)} · ${esc(r.ma)} · ${r.pallet}板</div>${hsd}</div>
        <div class="q">${r.qty}</div></div>`;
    }
    h += '</section>';
  }
  return h;
}

// ── Tab In HSD: hàng có hạn sử dụng làm đầu ca, theo xe rồi theo đơn ──
const hsdKey = (xe, r) => 'h|' + xe + '|' + r.don + '|' + r.ma + '|' + r.hsd;

function renderHsd() {
  if (!state.data) return needData('dữ liệu hôm nay');
  const list = core.buildHsd(state.data.hsd, state.vitri || { byMa: {}, byTen: {} });
  if (!list.length) return `<div class="empty">File này không có hàng có hạn sử dụng (sheet HSD).</div>`;
  const left = x => x.dons.reduce((s, d) => s + d.rows.filter(r => !state.ticks[hsdKey(x.xe, r)]).length, 0);
  const total = list.reduce((s, x) => s + x.count, 0), totalLeft = list.reduce((s, x) => s + left(x), 0);
  const q = state.q.trim().toLowerCase();

  if (!state.hsdXe || !list.some(x => x.xe === state.hsdXe)) {
    const shown = list.filter(x => !q || x.xe.toLowerCase().includes(q));
    let h = kpi([[list.length, 'xe'], [total, 'dòng'], [`${total - totalLeft}/${total}`, 'xong', totalLeft === 0 ? 'g' : '']]) + tools('Gõ số xe, vd 9911…');
    if (!shown.length) return h + '<div class="empty">Không có xe nào khớp.</div>';
    return h + '<div class="chips">' + shown.map(x => {
      const l = left(x);
      return `<button class="chip" data-xe="${esc(x.xe)}"><b>${esc(x.xe)}</b><small>${x.count} dòng${l < x.count ? ` · còn ${l}` : ''}</small></button>`;
    }).join('') + '</div>';
  }

  const x = list.find(v => v.xe === state.hsdXe), l = left(x);
  let h = `<div class="tools"><button class="btn" id="hsdBack">← Xe khác</button><span class="ctx" style="margin:0;flex:1;text-align:right"><b class="mono">${esc(x.xe)}</b></span></div>`
        + kpi([[x.dons.length, 'đơn'], [x.count, 'dòng'], [`${x.count - l}/${x.count}`, 'xong', l === 0 ? 'g' : '']]);
  for (const d of x.dons) {
    const dn = d.rows.filter(r => state.ticks[hsdKey(x.xe, r)]).length;
    h += `<section class="grp"><h3><span class="x">${esc(d.don)}</span><span class="g">${esc(d.khach)} · ${dn}/${d.rows.length}</span></h3>${bar(dn, d.rows.length)}`;
    for (const r of d.rows) {
      const k = hsdKey(x.xe, r);
      h += `<div class="r${state.ticks[k] ? ' done' : ''}" data-tick="${esc(k)}"><span class="ck">✓</span>
        <div class="m"><div class="t"><b>${esc(r.ten)}</b><span class="s mono" style="flex:none">${esc(r.loc)}</span></div>
          <div class="hs">${esc(r.hsd.split(/[\/+]/).join('  ') || '—')}</div>
          <div class="s">${esc(r.ma)}${r.desc ? ' · ' + esc(r.desc) : ''}</div></div>
        <div class="q">${r.qty}</div></div>`;
    }
    h += '</section>';
  }
  return h;
}

// ── Tab PCHome ──
function pchomeItems() {
  return core.buildPchome(state.data.main, state.data.hsd, state.vitri || { byMa: {}, byTen: {} }, state.manual);
}

function pchomeTotals(items) {
  let full = 0, loose = 0, pull = 0, need = 0;
  for (const it of items) { const c = core.calcItem(it); full += c.nFull; loose += c.nLoose; pull += c.nPull; if (!c.complete) need++; }
  return { full, loose, pull, need };
}

function renderPchome() {
  if (!state.data) return needData('dữ liệu hôm nay');
  const items = pchomeItems();
  if (!items.length) return `<div class="empty">File này không có hàng của ${core.PCHOME}.</div>`;
  const rows = state.data.main.filter(r => core.isPchome(r.khach));
  const t = pchomeTotals(items);
  const q = state.q.trim().toLowerCase();
  const shown = items.filter(it => {
    if (q && !(it.ten + ' ' + it.codes.join(' ')).toLowerCase().includes(q)) return false;
    return state.filter !== 'need' || !core.calcItem(it).complete;
  });
  let h = `<div id="pk">${pcKpi(t)}</div>
    <div class="ctx">${core.PCHOME} · ${new Set(rows.map(r => r.don)).size} đơn · ${esc([...new Set(rows.map(r => r.xe))].join(', '))}${state.vitri ? '' : ' · <span class="no">chưa có file vị trí kho</span>'}</div>`
    + tools('Tìm mã / tên hàng…', [['all', 'Tất cả'], ['need', 'Cần nhập']]);
  if (!shown.length) return h + '<div class="empty">Không có mặt hàng nào khớp.</div>';
  h += `<div class="ph"><span>Hàng</span><span>Cần</span><span>HSD</span><span>Bốc</span><span>Lôi</span></div>`;
  return h + shown.map(pcBlock).join('');
}

const pcKpi = t => kpi([[`${t.full}/${t.loose}`, 'bốc ra bản/lẻ'], [t.pull, 'lôi ra (bản)'], [t.need, 'thiếu thông số', t.need ? 'w' : 'g']]);

function pcBlock(it) {
  const c = core.calcItem(it), L = c.lots.length, key = it.ten;
  const box = (f, v) => `<input data-f="${f}" class="${v > 0 ? '' : 'miss'}" type="text" inputmode="numeric" pattern="[0-9]*" maxlength="4" value="${v || ''}" placeholder="?">`;
  const cells = c.lots.map((l, i) => {
    const tk = 'p|' + key + '|' + l.code, on = state.ticks[tk] ? ' done' : '';
    return `<span class="hs${on}" data-tick="${esc(tk)}" style="grid-column:3;grid-row:${i + 1}">${esc(l.code || '—')}${L > 1 ? `*${l.qty}` : ''}</span>`
         + `<span class="rs${on}" data-tick="${esc(tk)}" data-i="${i}" style="grid-column:4;grid-row:${i + 1}">${esc(l.out)}</span>`
         + `<span class="pl${on}" data-tick="${esc(tk)}" data-i="${i}" style="grid-column:5;grid-row:${i + 1}">${l.pull ?? '—'}</span>`;
  }).join('');
  // Đủ thông số thì chỉ hiện gọn "整板 · 底×高" dưới tên; thiếu thông số (hoặc bấm vào tên) mới mở hàng ô nhập
  const open = !c.complete || state.open.has(key);
  const spec = c.complete ? ` · ${c.full} · ${c.per}×${c.layers}` : '';
  return `<article class="p" data-key="${esc(key)}">
    <div class="n" data-open="${esc(key)}" style="grid-row:1 / span ${L}"><b>${esc(it.ten)}</b>${it.mismatch ? ' <span class="mm" title="Tổng các hạn dùng không bằng số lượng">⚠</span>' : ''}<small>${esc(it.codes.join(','))}${spec}</small></div>
    <div class="need" style="grid-row:1 / span ${L}">${it.qty}</div>${cells}
    ${open ? `<div class="sp" style="grid-row:${L + 1}"><label>整板${box('full', c.full)}</label><label>底${box('per', c.per)}</label>×<label>高${box('layers', c.layers)}</label><span class="cap">${c.cap ? '=' + c.cap + '/bản' : 'chưa cài'}</span></div>` : ''}
  </article>`;
}

/** Sửa 1 ô: ghi nhập tay (chỉ giá trị khác file cấu hình), cập nhật đúng khối đó — không dựng lại danh sách để không mất con trỏ. */
function onInput(e) {
  const inp = e.target;
  if (!inp.dataset.f) return;
  const card = inp.closest('.p'); const key = card.dataset.key;
  const item = pchomeItems().find(i => i.ten === key); if (!item) return;
  inp.value = inp.value.replace(/\D/g, '');
  const get = f => parseInt(card.querySelector(`[data-f="${f}"]`).value, 10) || 0;
  const per = get('per'), layers = get('layers'), full = get('full');
  const s = item.spec;
  const baseManual = per > 0 && layers > 0 && !(s.per === per && s.layers === layers);
  const fullManual = full > 0 && full !== s.full;
  if (baseManual || fullManual) state.manual[key] = { per: baseManual ? per : 0, layers: baseManual ? layers : 0, full: fullManual ? full : 0 };
  else delete state.manual[key];
  store.set('manual', state.manual);

  const c = core.calcItem({ ...item, manual: state.manual[key] || null });
  card.querySelector('[data-f="full"]').classList.toggle('miss', !(full > 0));
  card.querySelector('[data-f="per"]').classList.toggle('miss', !(per > 0));
  card.querySelector('[data-f="layers"]').classList.toggle('miss', !(layers > 0));
  card.querySelector('.cap').textContent = c.cap ? '=' + c.cap + '/bản' : 'chưa cài';
  card.querySelector('.n small').textContent = item.codes.join(',') + (c.complete ? ` · ${c.full} · ${c.per}×${c.layers}` : '');
  state.open.add(key);   // vừa sửa xong thì giữ ô nhập mở, tránh bị đóng giữa chừng khi đủ thông số
  c.lots.forEach((l, i) => {
    card.querySelector(`.rs[data-i="${i}"]`).textContent = l.out;
    card.querySelector(`.pl[data-i="${i}"]`).textContent = l.pull ?? '—';
  });
  $('#pk').innerHTML = pcKpi(pchomeTotals(pchomeItems()));
}

// ── Cài đặt ──
function renderSettings() {
  const line = (title, info, ok) => `<div class="li"><div><b>${title}</b><span class="${ok ? 'ok' : 'no'}">${info}</span></div></div>`;
  const d = state.data, v = state.vitri, o = state.doiBan;
  const src = x => (x.source === 'import' ? 'nạp tay' : 'có sẵn trong app');
  return `<div class="panel">
    ${line('Dữ liệu hôm nay', d ? `${esc(d.filename)} · ${d.main.length} dòng · nạp lúc ${new Date(d.loadedAt).toLocaleString('vi-VN')}` : 'chưa nạp', !!d)}
    ${line('Vị trí kho', v ? `${esc(v.filename)} · ${Object.keys(v.byMa).length} mã hàng · ${src(v)}` : 'chưa có', !!v)}
    ${line('Danh sách đổi bản', o ? `${esc(o.filename)} · ${Object.keys(o.dict).length} khách · ${src(o)}` : 'chưa có', !!o)}
    ${(v?.source === 'import' || o?.source === 'import') ? `<div class="li"><div><b>Dùng lại cấu hình có sẵn</b><span>bỏ file đã nạp tay</span></div><button class="btn" id="useBundled">Dùng</button></div>` : ''}
  </div>
  <div class="panel">
    <div class="li"><div><b>Dấu ✓ đã xong</b><span>${Object.keys(state.ticks).length} dòng</span></div><button class="btn" id="clearTicks">Xoá</button></div>
    <div class="li"><div><b>Cách xếp nhập tay (PCHome)</b><span>${Object.keys(state.manual).length} mặt hàng</span></div><button class="btn warn" id="clearManual">Xoá</button></div>
    <div class="li"><div><b>Xoá toàn bộ dữ liệu trên máy</b><span>dữ liệu, cấu hình, dấu ✓, nhập tay</span></div><button class="btn warn" id="clearAll">Xoá hết</button></div>
  </div>
  <div class="note"><b>Cách dùng:</b> bấm <b>Nạp file</b>, chọn file dữ liệu ngày (全虹…xlsx). Cấu hình vị trí kho và đổi bản đã có sẵn trong app.
    Nạp một lần là xong, mở lại không cần mạng.<br><b>Cài lên màn hình chính:</b> Safari → Chia sẻ → Thêm vào Màn hình chính.</div>`;
}

// ── Sự kiện ──
document.addEventListener('click', e => {
  const tab = e.target.closest('.tabs button');
  if (tab) { state.tab = tab.dataset.tab; state.q = ''; state.filter = 'all'; state.hsdXe = ''; store.set('tab', state.tab); render(); window.scrollTo(0, 0); return; }
  const chip = e.target.closest('.chip[data-xe]');
  if (chip) { state.hsdXe = chip.dataset.xe; render(); window.scrollTo(0, 0); return; }
  if (e.target.id === 'hsdBack') { state.hsdXe = ''; state.q = ''; render(); window.scrollTo(0, 0); return; }
  const opener = e.target.closest('[data-open]');
  if (opener) {
    const k = opener.dataset.open;
    if (state.open.has(k)) state.open.delete(k); else state.open.add(k);
    const y = window.scrollY; render(); window.scrollTo(0, y); return;
  }
  const seg = e.target.closest('.seg button[data-f]');
  if (seg) { state.filter = seg.dataset.f; render(); return; }
  const tick = e.target.closest('[data-tick]');
  if (tick && !e.target.closest('input')) {
    const k = tick.dataset.tick;
    if (state.ticks[k]) delete state.ticks[k]; else state.ticks[k] = 1;
    store.set('ticks', state.ticks);
    if (state.tab === 'pchome') document.querySelectorAll('[data-tick]').forEach(el => { if (el.dataset.tick === k) el.classList.toggle('done', !!state.ticks[k]); });
    else { const y = window.scrollY; render(); window.scrollTo(0, y); }
    return;
  }
  if (e.target.id === 'useBundled') {
    if (state.vitri?.source === 'import') { state.vitri = null; store.del('vitri'); }
    if (state.doiBan?.source === 'import') { state.doiBan = null; store.del('doiban'); }
    loadBundledConfig().then(() => { render(); toast('Đã dùng lại cấu hình có sẵn trong app.'); });
    return;
  }
  if (e.target.id === 'clearTicks') { state.ticks = {}; store.set('ticks', {}); render(); }
  if (e.target.id === 'clearManual') { state.manual = {}; store.set('manual', {}); render(); }
  if (e.target.id === 'clearAll') {
    if (!e.target.dataset.sure) { e.target.dataset.sure = 1; e.target.textContent = 'Bấm lần nữa'; return; }
    ['data', 'vitri', 'doiban', 'manual', 'ticks'].forEach(store.del);
    Object.assign(state, { data: null, vitri: null, doiBan: null, manual: {}, ticks: {} });
    loadBundledConfig().then(() => render()); render();
  }
});
document.addEventListener('input', e => {
  if (e.target.id === 'q') {
    state.q = e.target.value; const pos = e.target.selectionStart;
    render(); const q = $('#q'); if (q) { q.focus(); q.setSelectionRange(pos, pos); }
  } else onInput(e);
});
$('#fileInput').addEventListener('change', e => { const f = [...e.target.files]; e.target.value = ''; if (f.length) importFiles(f); });

// ── Ngoại tuyến ──
if ('serviceWorker' in navigator && (location.protocol === 'https:' || ['localhost', '127.0.0.1'].includes(location.hostname))) {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}

window.AF = { importBuffer, render, state, core, loadBundledConfig };   // dùng cho kiểm thử
render();
loadBundledConfig().then(changed => { if (changed) render(); });

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
  doiBan: store.get('doiban', null),        // { filename, dict }
  vitri: store.get('vitri', null),          // { filename, byMa, byTen }
  manual: store.get('manual', {}),          // nhập tay PCHome: { [tên hàng]: { per, layers, full } }
  ticks: store.get('ticks', {}),            // đã xong: { [khoá]: 1 }
  q: '', filter: 'all',
  hsdXe: '',                                // xe đang xem ở tab In HSD ('' = chưa chọn xe)
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
    try {
      const buf = await f.arrayBuffer();
      const msg = importBuffer(buf, f.name);
      done.push(msg);
    } catch (e) { fail.push(`${f.name}: ${e.message || e}`); }
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

// ── Cấu hình có sẵn trong app (config.json): dùng làm mặc định; file nạp tay thì ghi đè, đến khi bấm "Dùng bản có sẵn" ──
async function loadBundledConfig() {
  try {
    const r = await fetch('config.json', { cache: 'no-cache' });
    if (!r.ok) return false;
    const b = await r.json();
    let changed = false;
    if (b.vitri && state.vitri?.source !== 'import') {
      state.vitri = { ...b.vitri, source: 'bundled' }; store.set('vitri', state.vitri); changed = true;
    }
    if (b.doiBan && state.doiBan?.source !== 'import') {
      state.doiBan = { ...b.doiBan, source: 'bundled' }; store.set('doiban', state.doiBan); changed = true;
    }
    return changed;
  } catch { return false; }
}

// ── Màn hình ──
function render() {
  document.querySelectorAll('.tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === state.tab));
  $('#fileName').textContent = state.data ? state.data.filename : 'Chưa nạp file';
  const v = $('#view');
  v.innerHTML = state.tab === 'swap' ? renderSwap() : state.tab === 'pchome' ? renderPchome()
              : state.tab === 'hsd' ? renderHsd() : renderSettings();
  if (state.tab === 'pchome') updateTotals();
}

function needData(what) {
  return `<div class="empty">Chưa có ${what}.<br>Bấm <b>Nạp file</b> ở góc trên rồi chọn file Excel.<br>
    <span class="muted">Chọn được nhiều file một lần.</span></div>`;
}

// ── Tab Đổi bản ──
function swapRows() {
  const rows = core.buildSwap(state.data.main, state.data.hsd, state.doiBan.dict);
  const q = state.q.trim().toLowerCase();
  return rows.filter(r => (!q || [r.xe, r.khach, r.ten, r.ma, r.don].join(' ').toLowerCase().includes(q))
                       && (state.filter !== 'todo' || !state.ticks['s|' + r.xe + '|' + r.don + '|' + r.ma]));
}

function renderSwap() {
  if (!state.data) return needData('dữ liệu hôm nay');
  if (!state.doiBan) return needData('danh sách đổi bản (file <b>doi_ban_config.xlsx</b>)');
  const rows = swapRows();
  const all = core.buildSwap(state.data.main, state.data.hsd, state.doiBan.dict);
  const left = all.filter(r => !state.ticks['s|' + r.xe + '|' + r.don + '|' + r.ma]).length;
  let h = `<div class="tools"><input class="search" id="q" type="search" placeholder="Tìm xe, khách, tên hàng…" value="${esc(state.q)}" autocomplete="off">
    <div class="seg"><button data-f="all" class="${state.filter === 'all' ? 'on' : ''}">Tất cả</button>
    <button data-f="todo" class="${state.filter === 'todo' ? 'on' : ''}">Chưa xong (${left})</button></div></div>
    <div class="summary">${all.length} dòng đổi bản · ${new Set(all.map(r => r.xe)).size} xe · còn ${left} chưa xong</div>`;
  if (!rows.length) return h + `<div class="empty">${all.length ? 'Không có dòng nào khớp.' : 'Hôm nay không có hàng cần đổi bản.'}</div>`;
  const byXe = new Map();
  for (const r of rows) (byXe.get(r.xe) ?? byXe.set(r.xe, []).get(r.xe)).push(r);
  for (const [xe, list] of byXe) {
    h += `<section class="group"><h3>${esc(xe)} <small>${list.length} dòng</small></h3>`;
    for (const r of list) {
      const k = 's|' + r.xe + '|' + r.don + '|' + r.ma;
      const hsd = r.hsd ? `<div class="hsd">${esc(r.hsd.split(/[\/+]/).join('\n'))}</div>` : '';
      h += `<div class="row${state.ticks[k] ? ' done' : ''}" data-tick="${esc(k)}">
        <span class="ck">✓</span>
        <div class="main"><div class="name">${esc(r.ten)}</div>
          <div class="sub">${esc(r.khach)}<br><span class="mono">${esc(r.don)} · ${esc(r.ma)}</span> · ${r.pallet} 板</div>
          <div class="sub"><span class="tag">${esc(r.swap)}</span></div>${hsd}</div>
        <div class="qty">${r.qty}</div></div>`;
    }
    h += '</section>';
  }
  return h;
}

// ── Tab In HSD: hàng có hạn sử dụng làm đầu ca, xem theo xe rồi theo đơn ──
function renderHsd() {
  if (!state.data) return needData('dữ liệu hôm nay');
  const list = core.buildHsd(state.data.hsd, state.vitri || { byMa: {}, byTen: {} });
  if (!list.length) return `<div class="empty">File này không có hàng có hạn sử dụng (sheet HSD).</div>`;
  const keyOf = (xe, r) => 'h|' + xe + '|' + r.don + '|' + r.ma + '|' + r.hsd;
  const q = state.q.trim().toLowerCase();

  if (!state.hsdXe || !list.some(x => x.xe === state.hsdXe)) {
    // Chưa chọn xe: gõ vài số của biển số để lọc, bấm vào xe để mở
    const shown = list.filter(x => !q || x.xe.toLowerCase().includes(q));
    let h = `<div class="tools"><input class="search" id="q" type="search" inputmode="search" placeholder="Gõ số xe, vd 9911…" value="${esc(state.q)}" autocomplete="off"></div>
      <div class="summary">${list.length} xe có hàng HSD · ${list.reduce((s, x) => s + x.count, 0)} dòng</div>`;
    if (!shown.length) return h + '<div class="empty">Không có xe nào khớp.</div>';
    h += '<div class="chips">' + shown.map(x => {
      const left = x.dons.reduce((s, d) => s + d.rows.filter(r => !state.ticks[keyOf(x.xe, r)]).length, 0);
      return `<button class="chip" data-xe="${esc(x.xe)}"><b>${esc(x.xe)}</b><small>${x.count} dòng${left < x.count ? ` · còn ${left}` : ''}</small></button>`;
    }).join('') + '</div>';
    return h;
  }

  const x = list.find(v => v.xe === state.hsdXe);
  const left = x.dons.reduce((s, d) => s + d.rows.filter(r => !state.ticks[keyOf(x.xe, r)]).length, 0);
  let h = `<div class="tools"><button class="btn small" id="hsdBack">← Chọn xe khác</button></div>
    <div class="summary">${esc(x.xe)} · ${x.dons.length} đơn · ${x.count} dòng · còn ${left} chưa xong</div>`;
  for (const d of x.dons) {
    h += `<section class="group"><h3>${esc(d.don)} <small>${esc(d.khach)}</small></h3>`;
    for (const r of d.rows) {
      const k = keyOf(x.xe, r);
      const hsd = r.hsd ? `<div class="hsd">${esc(r.hsd.split(/[\/+]/).join('\n'))}</div>` : '';
      h += `<div class="row${state.ticks[k] ? ' done' : ''}" data-tick="${esc(k)}">
        <span class="ck">✓</span>
        <div class="main"><div class="name">${esc(r.ten)}</div>
          <div class="sub"><span class="mono">${esc(r.loc)} · ${esc(r.ma)}</span></div>${hsd}
          ${r.desc ? `<div class="sub">${esc(r.desc)}</div>` : ''}</div>
        <div class="qty">${r.qty}</div></div>`;
    }
    h += '</section>';
  }
  return h;
}

// ── Tab PCHome ──
function pchomeItems() {
  return core.buildPchome(state.data.main, state.data.hsd, state.vitri || { byMa: {}, byTen: {} }, state.manual);
}

function renderPchome() {
  if (!state.data) return needData('dữ liệu hôm nay');
  const items = pchomeItems();
  if (!items.length) return `<div class="empty">File này không có hàng của ${core.PCHOME}.</div>`;
  const q = state.q.trim().toLowerCase();
  const shown = items.filter(it => {
    const c = core.calcItem(it);
    if (q && !(it.ten + ' ' + it.codes.join(' ')).toLowerCase().includes(q)) return false;
    return state.filter === 'all' || (state.filter === 'need' ? !c.complete : c.complete);
  });
  const orders = new Set(state.data.main.filter(r => core.isPchome(r.khach)).map(r => r.don)).size;
  const xes = [...new Set(state.data.main.filter(r => core.isPchome(r.khach)).map(r => r.xe))];
  const noCfg = !state.vitri ? `<div class="warnline" style="margin:0 2px 10px">Chưa nạp <b>vi_tri_kho_config.xlsx</b> — 整板書量 và 底高 phải nhập tay.</div>` : '';
  let h = `<div class="tools"><input class="search" id="q" type="search" placeholder="Tìm mã hoặc tên hàng…" value="${esc(state.q)}" autocomplete="off">
    <div class="seg"><button data-f="all" class="${state.filter === 'all' ? 'on' : ''}">Tất cả</button>
    <button data-f="ok" class="${state.filter === 'ok' ? 'on' : ''}">Đủ thông số</button>
    <button data-f="need" class="${state.filter === 'need' ? 'on' : ''}">Cần nhập</button></div></div>
    <div class="summary" id="sum"></div>${noCfg}`;
  if (!shown.length) return h + '<div class="empty">Không có mặt hàng nào khớp.</div>';
  h += shown.map(it => cardHtml(it, orders, xes)).join('');
  return h;
}

function cardHtml(it) {
  const c = core.calcItem(it);
  const key = it.ten;
  const miss = v => (v > 0 ? '' : ' miss');
  const multi = it.lots.length > 1;
  const lots = c.lots.map(l => {
    const tk = 'p|' + key + '|' + l.code;
    return `<div class="l${state.ticks[tk] ? ' done' : ''}" data-tick="${esc(tk)}">
      <span class="ck">✓</span>
      <span class="code">${esc(l.code || '—')}${multi ? ` <small>*${l.qty}</small>` : ''}</span>
      <span class="res">${esc(l.out)}</span><span class="pull">${l.pull ?? '—'}</span></div>`;
  }).join('');
  return `<article class="card" data-key="${esc(key)}">
    <div class="head"><span class="nm">${esc(it.ten)}</span><span class="need">${it.qty}</span></div>
    <div class="meta">${esc(it.codes.join(', '))}</div>
    ${it.mismatch ? `<div class="warnline">Tổng các hạn dùng không bằng số lượng — kiểm tra lại cột HSD.</div>` : ''}
    <div class="inputs">
      <label>整板書量<input data-f="full" class="${miss(c.full)}" type="text" inputmode="numeric" pattern="[0-9]*" maxlength="4" value="${c.full || ''}" placeholder="?"></label>
      <label>底<input data-f="per" class="${miss(c.per)}" type="text" inputmode="numeric" pattern="[0-9]*" maxlength="4" value="${c.per || ''}" placeholder="?"></label>
      <span class="x">×</span>
      <label>高<input data-f="layers" class="${miss(c.layers)}" type="text" inputmode="numeric" pattern="[0-9]*" maxlength="4" value="${c.layers || ''}" placeholder="?"></label>
    </div>
    <div class="lots"><div class="h"><span></span><span>HSD</span><span>Bốc ra</span><span>Lôi ra</span></div>${lots}</div>
  </article>`;
}

/** Sửa 1 ô: ghi nhập tay (chỉ giá trị khác file cấu hình), tính lại đúng thẻ đó — không dựng lại cả danh sách để không mất con trỏ. */
function onInput(e) {
  const inp = e.target;
  if (!inp.dataset.f) return;
  const card = inp.closest('.card'); const key = card.dataset.key;
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
  const fresh = { ...item, manual: state.manual[key] || null };
  const c = core.calcItem(fresh);
  // Ô nào chưa đủ thì viền vàng; ô đang gõ giữ nguyên chữ người dùng nhập
  card.querySelector('[data-f="full"]').classList.toggle('miss', !(full > 0));
  card.querySelector('[data-f="per"]').classList.toggle('miss', !(per > 0));
  card.querySelector('[data-f="layers"]').classList.toggle('miss', !(layers > 0));
  card.querySelector('.lots').outerHTML = `<div class="lots"><div class="h"><span></span><span>HSD</span><span>Bốc ra</span><span>Lôi ra</span></div>${
    c.lots.map(l => { const tk = 'p|' + key + '|' + l.code;
      return `<div class="l${state.ticks[tk] ? ' done' : ''}" data-tick="${esc(tk)}"><span class="ck">✓</span>
        <span class="code">${esc(l.code || '—')}${item.lots.length > 1 ? ` <small>*${l.qty}</small>` : ''}</span>
        <span class="res">${esc(l.out)}</span><span class="pull">${l.pull ?? '—'}</span></div>`; }).join('')}</div>`;
  updateTotals();
}

function updateTotals() {
  const el = $('#sum'); if (!el) return;
  const items = pchomeItems(); let full = 0, loose = 0, pull = 0, need = 0;
  for (const it of items) { const c = core.calcItem(it); full += c.nFull; loose += c.nLoose; pull += c.nPull; if (!c.complete) need++; }
  const rows = state.data.main.filter(r => core.isPchome(r.khach));
  el.textContent = `${core.PCHOME} · ${new Set(rows.map(r => r.don)).size} đơn · ${[...new Set(rows.map(r => r.xe))].join(', ')}`
    + ` — bốc ra ${full} bản + ${loose} lẻ · lôi ra ${pull} bản nguyên` + (need ? ` · còn ${need} mặt hàng chưa đủ thông số` : '');
}

// ── Cài đặt ──
function renderSettings() {
  const item = (title, info, ok) => `<div class="li"><div><b>${title}</b><span class="${ok ? 'ok' : 'no'}">${info}</span></div></div>`;
  const d = state.data, v = state.vitri, o = state.doiBan;
  const when = d ? new Date(d.loadedAt).toLocaleString('vi-VN') : '';
  return `<div class="panel">
    ${item('Dữ liệu hôm nay', d ? `${esc(d.filename)} · ${d.main.length} dòng · nạp lúc ${when}` : 'chưa nạp', !!d)}
    ${item('Vị trí kho (vi_tri_kho_config)', v ? `${esc(v.filename)} · ${Object.keys(v.byMa).length} mã hàng · ${v.source === 'import' ? 'nạp tay' : 'có sẵn trong app'}` : 'chưa có', !!v)}
    ${item('Danh sách đổi bản (doi_ban_config)', o ? `${esc(o.filename)} · ${Object.keys(o.dict).length} khách · ${o.source === 'import' ? 'nạp tay' : 'có sẵn trong app'}` : 'chưa có', !!o)}
    ${(v?.source === 'import' || o?.source === 'import') ? `<div class="li"><div><b>Dùng lại cấu hình có sẵn</b><span>bỏ file đã nạp tay</span></div><button class="btn small" id="useBundled">Dùng</button></div>` : ''}
  </div>
  <div class="panel">
    <div class="li"><div><b>Dấu ✓ đã xong</b><span>${Object.keys(state.ticks).length} dòng</span></div><button class="btn small" id="clearTicks">Xoá</button></div>
    <div class="li"><div><b>Cách xếp nhập tay (PCHome)</b><span>${Object.keys(state.manual).length} mặt hàng</span></div><button class="btn small danger" id="clearManual">Xoá</button></div>
    <div class="li"><div><b>Xoá toàn bộ dữ liệu trên máy</b><span>dữ liệu, cấu hình, dấu ✓, nhập tay</span></div><button class="btn small danger" id="clearAll">Xoá hết</button></div>
  </div>
  <div class="note"><b>Cách dùng:</b> bấm <b>Nạp file</b>, chọn file dữ liệu ngày (全虹…xlsx). Hai file cấu hình chỉ cần nạp một lần,
    khi nào sửa thì nạp lại. Nạp một lần là xong, mở lại app không cần mạng.<br>
    <b>Cài lên màn hình chính:</b> Safari → nút Chia sẻ → Thêm vào Màn hình chính.</div>`;
}

// ── Sự kiện ──
document.addEventListener('click', e => {
  const tab = e.target.closest('.tabs button');
  if (tab) { state.tab = tab.dataset.tab; state.q = ''; state.filter = 'all'; state.hsdXe = ''; store.set('tab', state.tab); render(); window.scrollTo(0, 0); return; }
  const chip = e.target.closest('.chip[data-xe]');
  if (chip) { state.hsdXe = chip.dataset.xe; render(); window.scrollTo(0, 0); return; }
  if (e.target.id === 'hsdBack') { state.hsdXe = ''; state.q = ''; render(); window.scrollTo(0, 0); return; }
  const seg = e.target.closest('.seg button[data-f]');
  if (seg) { state.filter = seg.dataset.f; render(); return; }
  const tick = e.target.closest('[data-tick]');
  if (tick && !e.target.closest('input')) {
    const k = tick.dataset.tick;
    if (state.ticks[k]) delete state.ticks[k]; else state.ticks[k] = 1;
    store.set('ticks', state.ticks);
    if (state.tab === 'swap' || state.tab === 'hsd') { const y = window.scrollY; render(); window.scrollTo(0, y); }
    else { tick.classList.toggle('done', !!state.ticks[k]); }
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
    if (!e.target.dataset.sure) { e.target.dataset.sure = 1; e.target.textContent = 'Bấm lần nữa để xoá'; return; }
    ['data', 'vitri', 'doiban', 'manual', 'ticks'].forEach(store.del);
    Object.assign(state, { data: null, vitri: null, doiBan: null, manual: {}, ticks: {} }); render();
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

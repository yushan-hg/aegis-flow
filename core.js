// Logic đọc file + tính đổi bản, chuyển từ app Windows (DataManager, HsdProcessor, HsdLotParser, PchomeViewModel).
// Không đụng tới giao diện: chạy được cả trong trình duyệt lẫn Node (để kiểm thử).

export const PALLET_THRESHOLD = 0.65;      // 板位數 từ mức này trở lên là hàng đổi bản
export const PCHOME = 'Pchome十五庫';
const EXCLUDE_XE = 'KEV-5601';

const COL = { KHACH: '送達方小名', XE: '車牌號', DON: '裝運單號', MA: '物料號', TEN: '物料小名',
              QTY: '數量', QTY_SRC: '數量(按基礎單位)', PALLET: '板位數' };
const REQUIRED = [COL.KHACH, COL.XE, COL.DON, COL.MA, COL.TEN, COL.QTY, COL.PALLET];
const HSD_HEADERS = ['出貨效期', '效期', '新效', 'HạnSD', 'Hạn SD'];

const t = v => String(v ?? '').trim();
const num = s => { const n = parseFloat(String(s ?? '').replace(/,/g, '')); return Number.isFinite(n) ? n : 0; };
const ci = (a, b) => a.toLowerCase() === b.toLowerCase();

/** Mảng các dòng (mỗi dòng là mảng ô dạng chữ) của một sheet — giống ws.Cells[..].Text bên app Windows. */
export function sheetRows(XLSX, ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: true });
}

export function isPchome(khach) { return ci(t(khach), PCHOME); }

// ── Sheet Data → các dòng hàng ──────────────────────────────────────────────

function findHeaderRow(rows, a, b) {
  for (let r = 0; r < rows.length; r++) {
    const cells = (rows[r] || []).map(t);
    if (cells.includes(a) && cells.includes(b)) return r;
  }
  return -1;
}

export function parseMain(rows) {
  let h = findHeaderRow(rows, COL.XE, COL.MA);
  if (h < 0) h = 0;
  const map = {};
  (rows[h] || []).forEach((v, i) => { const k = t(v); if (k && !(k in map)) map[k] = i; });
  if (!(COL.QTY in map) && COL.QTY_SRC in map) map[COL.QTY] = map[COL.QTY_SRC];
  const missing = REQUIRED.filter(c => !(c in map));
  if (missing.length) throw new Error('Thiếu cột: ' + missing.join(', '));

  const out = [];
  for (let r = h + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const xe = t(row[map[COL.XE]]);
    if (!xe || xe === '0' || ci(xe, 'nan') || xe === COL.XE) continue;
    const khach = t(row[map[COL.KHACH]]);
    if (!khach || ci(khach, 'nan')) continue;
    if (xe.toLowerCase().includes(EXCLUDE_XE.toLowerCase())) continue;
    const qty = Math.round(num(row[map[COL.QTY]]));
    if (qty < 0) continue;
    out.push({ khach, xe, don: t(row[map[COL.DON]]), ma: t(row[map[COL.MA]]), ten: t(row[map[COL.TEN]]),
               qty, pallet: num(row[map[COL.PALLET]]) });
  }
  // Gộp theo (xe, đơn, mã, tên)
  const agg = new Map();
  for (const r of out) {
    const k = [r.xe, r.don, r.ma, r.ten].join('\u0001');
    const g = agg.get(k);
    if (g) { g.qty += r.qty; g.pallet += r.pallet; } else agg.set(k, { ...r });
  }
  return [...agg.values()].sort((a, b) => a.xe.localeCompare(b.xe) || a.don.localeCompare(b.don));
}

// ── Sheet HSD ───────────────────────────────────────────────────────────────

const HSD_START = /^\d{6}[A-Za-z]/;
const HSD_CODE  = /^\d{6}[A-Za-z][A-Za-z0-9]*$/;

function hsdHeaderRows(rows) {
  const res = [];
  for (let r = 0; r < rows.length; r++) {
    const cells = (rows[r] || []).map(t);
    if (cells.includes(COL.XE) && cells.includes(COL.MA)) res.push(r);
  }
  return res;
}

function isHsdSheet(rows) {
  for (const row of rows) for (const v of row || []) {
    if (HSD_HEADERS.some(h => ci(h, t(v)))) return true;
  }
  for (const row of rows) for (const v of row || []) if (HSD_CODE.test(t(v))) return true;
  return false;
}

export function parseHsd(XLSX, wb) {
  const fixed = wb.SheetNames.find(n => ci(n, 'HSD'));
  const names = fixed ? [fixed] : wb.SheetNames.filter(n => !ci(n, 'Data'));
  let block = 0;
  const all = [];
  for (const name of names) {
    const rows = sheetRows(XLSX, wb.Sheets[name]);
    if (!rows.length) continue;
    if (!fixed && !isHsdSheet(rows)) continue;
    let headers = hsdHeaderRows(rows);
    if (!headers.length) headers = [0];
    const lastByDon = new Map();
    headers.forEach((hr, bi) => {
      const end = bi + 1 < headers.length ? headers[bi + 1] - 1 : rows.length - 1;
      const id = ++block;
      const cm = {}; let hsdCol = -1;
      (rows[hr] || []).forEach((v, i) => {
        const k = t(v); if (!k) return;
        if (!(k in cm)) cm[k] = i;
        if (HSD_HEADERS.some(x => ci(x, k))) hsdCol = i;
      });
      const col = { xe: cm[COL.XE] ?? -1, don: cm[COL.DON] ?? -1, ma: cm[COL.MA] ?? -1, ten: cm[COL.TEN] ?? -1,
                    qty: cm[COL.QTY] ?? cm[COL.QTY_SRC] ?? -1, khach: cm[COL.KHACH] ?? -1 };
      if (hsdCol < 0) {
        outer: for (let c = 0; c < (rows[hr] || []).length + 30; c++)
          for (let r = hr + 1; r <= end; r++) if (HSD_CODE.test(t((rows[r] || [])[c]))) { hsdCol = c; break outer; }
      }
      if (col.xe < 0 || col.don < 0 || col.ma < 0) return;
      const headerTexts = new Set((rows[hr] || []).map(t).filter(Boolean));
      for (let r = hr + 1; r <= end; r++) {
        const row = rows[r] || [];
        const xe = t(row[col.xe]);
        if (!xe || xe === '0' || ci(xe, 'nan') || headerTexts.has(xe)) continue;
        const don = t(row[col.don]);
        let raw = hsdCol >= 0 ? t(row[hsdCol]) : '';
        if (!raw) {
          for (let c = 0; c < row.length; c++) {
            if ([col.xe, col.don, col.ma, col.ten, col.qty, col.khach].includes(c)) continue;
            const v = t(row[c]);
            if (v && HSD_START.test(v)) { raw = v; break; }
          }
        }
        if (!raw) raw = lastByDon.get(don) ?? ''; else lastByDon.set(don, raw);
        all.push({ blockId: id, item: { khach: col.khach >= 0 ? t(row[col.khach]) : '', xe, don, ma: t(row[col.ma]),
          ten: col.ten >= 0 ? t(row[col.ten]) : '', qty: Math.round(num(row[col.qty])), hsd: raw } });
      }
    });
  }
  // Khử trùng lặp 2 lớp như HsdProcessor (A: trùng nguyên vẹn cùng khối; B: chỉ giữ khối sau cùng của mỗi (xe, đơn, mã))
  const seen = new Set(); const a = [];
  for (const x of all) {
    const i = x.item;
    const k = [i.xe, i.don, i.ma, i.ten, i.qty, i.hsd, x.blockId].join('\u0001');
    if (!seen.has(k)) { seen.add(k); a.push(x); }
  }
  const last = new Map();
  for (const x of a) { const k = [x.item.xe, x.item.don, x.item.ma].join('\u0001'); last.set(k, Math.max(last.get(k) ?? 0, x.blockId)); }
  return a.filter(x => x.blockId === last.get([x.item.xe, x.item.don, x.item.ma].join('\u0001'))).map(x => x.item);
}

// ── Chuỗi HSD → các lô (đúng như HsdLotParser bên app Windows) ───────────────

const LOT = /^(\d{8}|\d{6})\s*([A-Za-z][A-Za-z0-9]*)?\s*(?:\*\s*(\d+))?$/;

/** @returns {{lots: {key:string, code:string, qty:number}[], mismatch:boolean}} code = mã HSD nguyên văn, vd "270416HAOO" */
export function parseLots(raw, totalQty) {
  const text = t(raw);
  if (!text) return { lots: [{ key: '', code: '', qty: totalQty }], mismatch: false };
  let mismatch = false;
  const parts = text.split('/').map(t).filter(Boolean);
  const lots = parts.map(p => {
    const m = LOT.exec(p);
    if (!m) { mismatch = true; return { d: '', q: null, code: p }; }
    return { d: m[1], q: m[3] != null ? parseInt(m[3], 10) : null, code: p.split('*')[0].trim() };
  });
  const known = lots.reduce((s, l) => s + (l.q ?? 0), 0);
  const unknown = lots.filter(l => l.q == null).length;
  const res = lots.map(l => {
    const q = l.q ?? (unknown === 1 ? Math.max(0, totalQty - known) : 0);
    return { key: l.d ? (l.d.length === 8 ? l.d : '20' + l.d) : '~' + l.code, code: l.code, qty: q };
  });
  if (unknown > 1 || res.reduce((s, l) => s + l.qty, 0) !== totalQty) mismatch = true;
  // Gộp lô trùng, sắp theo ngày
  const g = new Map();
  for (const l of res) { const e = g.get(l.key); if (e) e.qty += l.qty; else g.set(l.key, { ...l }); }
  return { lots: [...g.values()].sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0)), mismatch };
}

// ── File cấu hình ───────────────────────────────────────────────────────────

/** doi_ban_config.xlsx → { 'tên khách': 'loại bản' } */
export function parseDoiBan(rows) {
  const dict = {};
  let cK = -1, cD = -1, hr = -1;
  for (let r = 0; r < Math.min(5, rows.length) && hr < 0; r++) {
    (rows[r] || []).forEach((v, i) => {
      const k = t(v);
      if (k === COL.KHACH) cK = i;
      if (['換板需求', '棧板類型', '板類型'].includes(k)) cD = i;
    });
    if (cK >= 0 && cD >= 0) hr = r;
  }
  if (hr < 0) return dict;
  for (let r = hr + 1; r < rows.length; r++) {
    const k = t((rows[r] || [])[cK]), d = t((rows[r] || [])[cD]);
    if (k && d) dict[k] = d;
  }
  return dict;
}

export function matchPallet(dict, khach) {
  if (!t(khach)) return '######';
  for (const [k, v] of Object.entries(dict)) if (khach.toLowerCase().includes(k.toLowerCase())) return v;
  return '######';
}

/** vi_tri_kho_config.xlsx → { byMa: {...}, byTen: {...} } với {loc, desc, full, per, layers} */
export function parseVitri(rows) {
  const out = { byMa: {}, byTen: {} };
  const hdr = (rows[0] || []).map(t);
  let cMa = -1, cTen = -1, cLoc = -1, cDesc = -1, cFull = -1, cBase = -1;
  hdr.forEach((h, i) => {
    if (h === COL.MA) cMa = i;
    if (h === COL.TEN) cTen = i;
    if (h === '儲位') cLoc = i;
    if (h === '物料描述') cDesc = i;
    if (h.includes('整板')) cFull = i;
    if (h.includes('底高') && (cBase < 0 || ci(h.slice(0, 6), 'PCHome') || h.toLowerCase().includes('pchome'))) cBase = i;
  });
  // Tiêu đề "PCHome 底高" gộp ô 2 cột, hoặc tách "PCHome" | "底高"
  if (cBase > 0 && !hdr[cBase].toLowerCase().includes('pchome') && hdr[cBase - 1].toLowerCase().includes('pchome')) cBase--;
  if (cMa < 0) return out;
  const firstInt = s => { const m = /\d+/.exec(t(s)); return m ? parseInt(m[0], 10) : 0; };
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const ma = t(row[cMa]); if (!ma) continue;
    const ten = cTen >= 0 ? t(row[cTen]) : '';
    let per = 0, layers = 0;
    if (cBase >= 0) {
      const nums = t(row[cBase]).match(/\d+/g) || [];
      if (nums.length === 1) { const nx = /\d+/.exec(t(row[cBase + 1])); if (nx) { per = +nums[0]; layers = +nx[0]; } }
      else if (nums.length >= 2) { per = +nums[0]; layers = +nums[1]; }
    }
    const rec = { loc: cLoc >= 0 ? t(row[cLoc]) : '', desc: cDesc >= 0 ? t(row[cDesc]) : '',
                  full: cFull >= 0 ? firstInt(row[cFull]) : 0, per, layers };
    out.byMa[ma] = rec;
    if (ten && !(ten in out.byTen)) out.byTen[ten] = rec;
  }
  return out;
}

// ── Nhận diện loại file khi nạp ─────────────────────────────────────────────

/** @returns {'data'|'vitri'|'doiban'|null} */
export function detectKind(XLSX, wb) {
  for (const name of wb.SheetNames) {
    const rows = sheetRows(XLSX, wb.Sheets[name]).slice(0, 12);
    for (const row of rows) {
      const c = (row || []).map(t);
      if (c.includes(COL.XE) && c.includes(COL.MA) && c.includes(COL.PALLET)) return 'data';
    }
    const h = (rows[0] || []).map(t);
    if (h.includes(COL.MA) && h.includes('儲位')) return 'vitri';
    for (const row of rows.slice(0, 5)) {
      const c = (row || []).map(t);
      if (c.includes(COL.KHACH) && c.some(x => ['換板需求', '棧板類型', '板類型'].includes(x))) return 'doiban';
    }
  }
  return null;
}

// ── Tab Đổi bản ─────────────────────────────────────────────────────────────

/** Các dòng cần đổi bản (板位數 ≥ 0.65, khách có trong danh sách đổi bản), kèm HSD đối chiếu theo (xe, đơn, mã). */
export function buildSwap(main, hsd, doiBan) {
  const hmap = new Map();
  for (const h of hsd) {
    const k = [h.xe, h.don, h.ma].join('\u0001');
    (hmap.get(k) ?? hmap.set(k, []).get(k)).push(h.hsd);
  }
  const rows = [];
  for (const r of main) {
    if (r.pallet < PALLET_THRESHOLD) continue;
    if (isPchome(r.khach)) continue;   // Pchome十五庫 chỉ xử lý ở tab PCHome
    const swap = matchPallet(doiBan, r.khach);
    if (swap === '######') continue;
    const hs = [...new Set((hmap.get([r.xe, r.don, r.ma].join('\u0001')) || []).filter(Boolean))];
    rows.push({ ...r, pallet: Math.round(r.pallet * 100) / 100, swap, hsd: hs.join('/') });
  }
  return rows.sort((a, b) => a.xe.localeCompare(b.xe));
}

// ── Tab In HSD ──────────────────────────────────────────────────────────────

const locKey = loc => (/\d/.test(loc || '') ? 0 : 1);   // vị trí có số đi trước, "Chưa cài" xuống cuối (giống phiếu in HSD)

/** Hàng có hạn sử dụng, gom theo xe rồi theo đơn; trong đơn xếp theo 儲位 để đi kho một vòng. */
export function buildHsd(hsd, vitri) {
  const noise = new Set(['物料小名', '0', '', 'nan']);
  const byXe = new Map();
  for (const h of hsd) {
    if (noise.has(h.ten)) continue;
    const v = vitri.byMa[h.ma] || {};
    const row = { ...h, loc: v.loc || 'Chưa cài', desc: v.desc || '' };
    const dons = byXe.get(h.xe) ?? byXe.set(h.xe, new Map()).get(h.xe);
    (dons.get(h.don) ?? dons.set(h.don, []).get(h.don)).push(row);
  }
  const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
  return [...byXe].sort((a, b) => cmp(a[0], b[0])).map(([xe, dons]) => {
    const list = [...dons].sort((a, b) => cmp(a[0], b[0])).map(([don, rows]) => ({
      don,
      khach: rows.find(r => r.khach)?.khach || '',
      rows: rows.sort((a, b) => locKey(a.loc) - locKey(b.loc) || cmp(a.loc, b.loc) || cmp(a.ma, b.ma)),
    }));
    return { xe, dons: list, count: list.reduce((s, d) => s + d.rows.length, 0) };
  });
}

// ── Tab PCHome ──────────────────────────────────────────────────────────────

/** Gộp hàng Pchome theo tên hàng; mỗi hạn dùng một lô. manual = ghi đè nhập tay { [ten]: {per, layers, full} } */
export function buildPchome(main, hsd, vitri, manual) {
  const rows = main.filter(r => isPchome(r.khach));
  const hs = new Map();
  for (const h of hsd) { const k = [h.xe, h.don, h.ma].join('\u0001'); (hs.get(k) ?? hs.set(k, []).get(k)).push(h); }
  const groups = new Map();
  for (const r of rows) (groups.get(r.ten) ?? groups.set(r.ten, []).get(r.ten)).push(r);

  const items = [];
  for (const [ten, rs] of [...groups].sort((a, b) => a[0].localeCompare(b[0]))) {
    let mismatch = false;
    const lots = new Map();
    for (const r of rs) {
      const list = hs.get([r.xe, r.don, r.ma].join('\u0001')) || [];
      let parsed = [];
      if (!list.length) parsed = [{ key: '', code: '', qty: r.qty }];
      else {
        for (const h of list) { const p = parseLots(h.hsd, h.qty); parsed.push(...p.lots); mismatch ||= p.mismatch; }
        if (parsed.reduce((s, l) => s + l.qty, 0) !== r.qty) mismatch = true;
      }
      for (const l of parsed) {
        const k = l.key, e = lots.get(k);
        if (e) e.qty += l.qty; else lots.set(k, { ...l });
      }
    }
    const codes = [...new Set(rs.map(r => r.ma).filter(Boolean))];
    const spec = codes.map(c => vitri.byMa[c]).find(Boolean) || vitri.byTen[ten] || { loc: '', desc: '', full: 0, per: 0, layers: 0 };
    items.push({ ten, codes, loc: spec.loc || '', qty: rs.reduce((s, r) => s + r.qty, 0), mismatch,
                 lots: [...lots.values()].sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0)),
                 spec: { full: spec.full, per: spec.per, layers: spec.layers }, manual: manual?.[ten] || null });
  }
  return items;
}

/** Giá trị đang dùng của một mặt hàng: nhập tay ghi đè file cấu hình. */
export function effective(item) {
  const m = item.manual || {};
  const hasBase = m.per > 0 && m.layers > 0;
  return {
    per:    hasBase ? m.per    : item.spec.per,
    layers: hasBase ? m.layers : item.spec.layers,
    full:   m.full > 0 ? m.full : item.spec.full,
  };
}

/** Kết quả từng hạn: bốc ra theo cách xếp PCHome (bản đầy/lẻ) và số bản nguyên phải lôi ra. */
export function calcItem(item) {
  const e = effective(item);
  const cap = e.per > 0 && e.layers > 0 ? e.per * e.layers : 0;
  let full = 0, loose = 0, pull = 0;
  const lots = item.lots.map(l => {
    const lf = cap ? Math.floor(l.qty / cap) : null;
    const ll = cap ? l.qty % cap : null;
    const p = e.full > 0 ? Math.ceil(l.qty / e.full) : null;
    if (cap) { full += lf; loose += ll; }
    if (p != null) pull += p;
    // "2/20" = 2 bản đầy + lẻ 20; "13" = chỉ lẻ 13; "2/0" = 2 bản không lẻ
    const out = !cap ? '—' : lf > 0 ? `${lf}/${ll}` : ll > 0 ? String(ll) : '—';
    return { ...l, full: lf, loose: ll, out, pull: p };
  });
  // full = 整板書量 (nguyên bản); nFull/nLoose/nPull = tổng bản đầy / lẻ / bản nguyên phải lôi
  return { ...e, cap, nFull: full, nLoose: loose, nPull: pull, lots, complete: cap > 0 && e.full > 0 };
}

// ── Bỏ dòng có HSD, thống kê xe, tìm hàng (đúng như DataManager bên app Windows) ─────────────

/** Trừ số lượng đã có trong sheet HSD (cùng xe + đơn + mã, có ghi hạn) khỏi từng dòng; 板位數 phần còn lại tính theo tỉ lệ. */
export function excludeHsd(main, hsd) {
  const left = new Map();
  for (const h of hsd) {
    if (!t(h.hsd)) continue;
    const k = [h.xe, h.don, h.ma].join('\u0001');
    left.set(k, (left.get(k) || 0) + h.qty);
  }
  const out = [];
  for (const r of main) {
    const k = [r.xe, r.don, r.ma].join('\u0001');
    const q = left.get(k) || 0;
    if (q <= 0) { out.push(r); continue; }
    const take = Math.min(q, r.qty);
    left.set(k, q - take);
    const rest = r.qty - take;
    if (rest <= 0) continue;
    out.push({ ...r, qty: rest, pallet: r.qty > 0 ? r.pallet * rest / r.qty : 0 });
  }
  return out;
}

/** Thống kê từng xe: nguyên bản, bản đổi (flip), nhặt lẻ = nguyên bản − flip. Sắp theo nhặt lẻ giảm dần. */
export function buildStats(rows) {
  const m = new Map();
  for (const r of rows) {
    const s = m.get(r.xe) ?? m.set(r.xe, { xe: r.xe, nguyen: 0, flip: 0, nhatLe: 0, flipPallets: 0 }).get(r.xe);
    s.nguyen += r.qty;
    if (r.pallet >= PALLET_THRESHOLD) { s.flip += r.qty; s.flipPallets += r.pallet; }
  }
  const list = [...m.values()];
  for (const s of list) s.nhatLe = s.nguyen - s.flip;
  return list.sort((a, b) => b.nhatLe - a.nhatLe);
}

/** Tìm hàng: mỗi mặt hàng → tổng và từng (xe, đơn). */
export function buildItems(rows) {
  const m = new Map();
  for (const r of rows) {
    if (!r.ten) continue;
    const it = m.get(r.ten) ?? m.set(r.ten, { ten: r.ten, ma: new Set(), qty: 0, per: new Map() }).get(r.ten);
    it.ma.add(r.ma); it.qty += r.qty;
    const k = r.xe + '\u0001' + r.don;
    const e = it.per.get(k);
    if (e) e.qty += r.qty; else it.per.set(k, { xe: r.xe, don: r.don, khach: r.khach, qty: r.qty });
  }
  return [...m.values()].sort((a, b) => a.ten.localeCompare(b.ten)).map(it => ({
    ten: it.ten, ma: [...it.ma].filter(Boolean), qty: it.qty,
    rows: [...it.per.values()].sort((a, b) => a.xe.localeCompare(b.xe) || a.don.localeCompare(b.don)),
  }));
}

/** list_xe_config.xlsx → [{taiXe, xe}] */
export function parseListXe(rows) {
  const hdr = (rows[0] || []).map(t);
  const cX = hdr.indexOf(COL.XE), cT = hdr.indexOf('司機姓名');
  if (cX < 0) return [];
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const xe = t((rows[r] || [])[cX]); if (!xe) continue;
    out.push({ taiXe: cT >= 0 ? t(rows[r][cT]) : '', xe });
  }
  return out;
}

/** Danh sách nhân sự (staff_dynamic_config_v3.json) → [{name, weight:'100%', active}] */
export const WEIGHTS = ['100%', '75%', '65%', '50%'];
export const weightRatio = w => ({ '100%': 1, '75%': 0.75, '65%': 0.65, '50%': 0.5 }[w] ?? 1);
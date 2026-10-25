// Chữ trên giao diện: Tiếng Việt · 中文 · English. Thiếu bản dịch thì dùng tiếng Việt.
export const LANGS = [['vi', 'Tiếng Việt'], ['zh', '中文'], ['en', 'English']];

const D = {
  vi: {
    noFile: 'Chưa nạp file', load: 'Nạp file', reload: 'Nạp file khác',
    t_home: 'Trang chủ', t_work: 'Việc làm', t_find: 'Tra cứu',
    file_lines: (m, h) => `${m} dòng · ${h} dòng HSD`,
    need_data: 'Chưa có dữ liệu hôm nay', need_data2: 'Bấm "Nạp file" rồi chọn file Excel dữ liệu ngày (全虹…xlsx).',
    gas: 'Bình khí', no_deliver: 'Xe không giao hàng', none: 'Không có', swap: 'Đổi bản', pchome: 'PCHome',
    need_swap_cfg: 'Chưa có danh sách đổi bản (doi_ban_config).', no_swap: 'Hôm nay không có hàng cần đổi bản.',
    need_hsd: 'File này không có hàng có hạn sử dụng.', no_pc: 'File này không có hàng của Pchome十五庫.', no_match: 'Không có dòng nào khớp.',
    lines: 'dòng', xe_n: n => `${n} xe`, xe_lines: (x, l) => `${x} xe · ${l} dòng`, all: 'Tất cả',
    seg_hsd: 'HSD', seg_lines: 'Dòng hàng', seg_item: 'Tìm hàng',
    c_xe: 'Xe', c_don: 'Đơn', c_khach: 'Khách', c_ma: 'Mã', c_ten: 'Hàng', c_qty: 'SL', c_hsd: 'HSD', c_loc: 'Vị trí', c_swap: 'Đổi sang', c_out: 'Bản/Lẻ',
    ph_search: 'Tìm khách, xe, đơn, mã, tên hàng…', ph_item: 'Tên hàng, vd CK6N', more_rows: n => `Còn ${n} dòng nữa — gõ thêm để lọc.`,
    pc_note: 'Bản/Lẻ: 2/20 = 2 bản chẵn + 20 lẻ · 13 = chỉ lẻ · — hoặc ? = chưa có thông số xếp bản.', pc_mismatch: 'Tổng các hạn dùng không bằng số lượng',
    lang: 'Ngôn ngữ', theme: 'Giao diện', theme_auto: 'Theo máy', theme_dark: 'Tối', theme_light: 'Sáng', offline: 'chạy ngoại tuyến',
    imported_data: (f, m, h) => `Đã nạp ${f} (${m} dòng, ${h} dòng HSD)`, imported_vitri: n => `Vị trí kho: ${n} mã hàng`, imported_doiban: n => `Danh sách đổi bản: ${n} khách`,
    bad_file: 'không nhận ra loại file (cần file dữ liệu ngày, vi_tri_kho_config hoặc doi_ban_config)', save_fail: 'Máy không lưu nổi dữ liệu — vẫn dùng được nhưng mở lại phải nạp lại.',
  },
  zh: {
    noFile: '尚未載入檔案', load: '載入檔案', reload: '載入其他檔案',
    t_home: '首頁', t_work: '作業', t_find: '查詢',
    file_lines: (m, h) => `${m} 行 · 效期 ${h} 行`,
    need_data: '尚無今日資料', need_data2: '點「載入檔案」，選擇出貨資料 Excel（全虹…xlsx）。',
    gas: '鋼瓶', no_deliver: '無出貨車輛', none: '無', swap: '換板', pchome: 'PCHome',
    need_swap_cfg: '尚無換板名單（doi_ban_config）。', no_swap: '今日沒有需要換板的貨。',
    need_hsd: '此檔案沒有效期品。', no_pc: '此檔案沒有 Pchome十五庫 的貨。', no_match: '沒有符合的資料。',
    lines: '行', xe_n: n => `${n} 輛`, xe_lines: (x, l) => `${x} 輛 · ${l} 行`, all: '全部',
    seg_hsd: '效期', seg_lines: '明細', seg_item: '找貨',
    c_xe: '車', c_don: '單號', c_khach: '客戶', c_ma: '料號', c_ten: '品名', c_qty: '數量', c_hsd: '效期', c_loc: '儲位', c_swap: '換成', c_out: '板/零',
    ph_search: '搜尋客戶、車、單號、料號、品名…', ph_item: '品名，例如 CK6N', more_rows: n => `還有 ${n} 行 — 輸入關鍵字篩選。`,
    pc_note: '板/零：2/20 = 2 整板 + 20 零 · 13 = 只有零 · — 或 ? = 尚未設定疊法。', pc_mismatch: '各效期合計不等於數量',
    lang: '語言', theme: '外觀', theme_auto: '跟隨系統', theme_dark: '深色', theme_light: '淺色', offline: '可離線使用',
    imported_data: (f, m, h) => `已載入 ${f}（${m} 行，效期 ${h} 行）`, imported_vitri: n => `儲位：${n} 個料號`, imported_doiban: n => `換板名單：${n} 家客戶`,
    bad_file: '無法辨識檔案類型（需要出貨資料、vi_tri_kho_config 或 doi_ban_config）', save_fail: '裝置無法儲存資料 — 仍可使用，但重開後需重新載入。',
  },
  en: {
    noFile: 'No file loaded', load: 'Load file', reload: 'Load another file',
    t_home: 'Home', t_work: 'Tasks', t_find: 'Lookup',
    file_lines: (m, h) => `${m} lines · ${h} expiry lines`,
    need_data: 'No data for today yet', need_data2: 'Tap "Load file" and choose the daily Excel file (全虹…xlsx).',
    gas: 'Gas cylinders', no_deliver: 'Vehicles not delivering', none: 'None', swap: 'Pallet swap', pchome: 'PCHome',
    need_swap_cfg: 'Pallet swap list (doi_ban_config) missing.', no_swap: 'No pallet swaps today.',
    need_hsd: 'This file has no expiry lines.', no_pc: 'This file has no Pchome十五庫 goods.', no_match: 'Nothing matches.',
    lines: 'lines', xe_n: n => `${n} vehicles`, xe_lines: (x, l) => `${x} vehicles · ${l} lines`, all: 'All',
    seg_hsd: 'Expiry', seg_lines: 'Lines', seg_item: 'Find item',
    c_xe: 'Veh.', c_don: 'Order', c_khach: 'Customer', c_ma: 'Code', c_ten: 'Item', c_qty: 'Qty', c_hsd: 'Expiry', c_loc: 'Location', c_swap: 'Swap to', c_out: 'Plt/Loose',
    ph_search: 'Search customer, vehicle, order, code, item…', ph_item: 'Item name, e.g. CK6N', more_rows: n => `${n} more lines — type to filter.`,
    pc_note: 'Plt/Loose: 2/20 = 2 full pallets + 20 loose · 13 = loose only · — or ? = stacking not set.', pc_mismatch: 'Expiry quantities do not add up to the line quantity',
    lang: 'Language', theme: 'Appearance', theme_auto: 'Device', theme_dark: 'Dark', theme_light: 'Light', offline: 'works offline',
    imported_data: (f, m, h) => `Loaded ${f} (${m} lines, ${h} expiry lines)`, imported_vitri: n => `Locations: ${n} items`, imported_doiban: n => `Pallet-swap list: ${n} customers`,
    bad_file: 'unrecognised file (need the daily data, vi_tri_kho_config or doi_ban_config)', save_fail: 'Could not save data on this phone — it still works but must be reloaded next time.',
  },
};

export function makeT(lang) {
  const d = D[lang] || D.vi;
  return (k, ...a) => { const v = d[k] ?? D.vi[k] ?? k; return typeof v === 'function' ? v(...a) : v; };
}

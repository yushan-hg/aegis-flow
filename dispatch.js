// Chia đơn cho nhân viên — chuyển từ DispatchViewModel / DataManager.SmartDistribute bên app Windows.
// Tải của một xe = số lượng nhặt lẻ (nhatLe). % = (tải / năng lực) so với phần công bằng của cả nhóm.
import { weightRatio } from './core.js';

export const OVER = 115, UNDER = 85, UNDO_LIMIT = 50;

export class Dispatch {
  /** @param stats [{xe, nguyen, flip, nhatLe, flipPallets}]  @param staff [{name, weight, active}] */
  constructor(stats, staff) {
    this.stats = stats;
    this.byXe = new Map(stats.map(s => [s.xe, s]));
    this.staff = staff;            // dùng chung với giao diện: bật/tắt người đi làm ở đây
    this.owner = new Map();        // xe → tên nhân viên
    this.undoStack = [];
  }

  // ── Người làm ──
  w(name) { return weightRatio(this.staff.find(s => s.name === name)?.weight); }
  /** Người hiện thẻ: đang đi làm, hoặc đã có xe (giống SyncWorkerCards: không bỏ thẻ đã gán dở) */
  workers() {
    return this.staff.filter(s => s.active || this.xeOf(s.name).length > 0);
  }
  xeOf(name) { return [...this.owner].filter(([, n]) => n === name).map(([x]) => x); }
  load(name, owner = this.owner) {
    let sum = 0;
    for (const [x, n] of owner) if (n === name) sum += this.byXe.get(x)?.nhatLe ?? 0;
    return sum;
  }

  // ── Chỉ số ──
  /** Mức tải/năng lực mà mỗi người "được chia công bằng": tổng tải mọi xe cần chia / tổng năng lực của nhóm */
  avgFactor(workers = this.workers()) {
    const weighted = workers.filter(s => weightRatio(s.weight) > 0);
    const totalW = weighted.reduce((a, s) => a + weightRatio(s.weight), 0);
    if (totalW <= 0) return 0;
    let total = this.stats.filter(v => v.nhatLe > 0 || v.flipPallets > 0).reduce((a, v) => a + v.nhatLe, 0);
    if (total <= 0) total = weighted.reduce((a, s) => a + this.load(s.name), 0);
    return total / totalW;
  }
  /** % tải; extra = cộng thêm số lượng (xem trước khi nhận một xe) */
  pct(name, extra = 0, avg = this.avgFactor()) {
    const w = this.w(name);
    return avg > 0 && w > 0 ? Math.round(((this.load(name) + extra) / w / avg) * 100) : 0;
  }
  status(name) {
    const n = this.xeOf(name).length, p = this.pct(name);
    if (n && p > OVER) return 'over';
    return p < UNDER ? (n ? 'under' : 'idle') : 'good';
  }
  /** Người lệch nhất so với 100% (chỉ tính người đã có xe) — hiện trên thông báo */
  worst() {
    const a = this.workers().filter(s => this.xeOf(s.name).length);
    if (!a.length) return null;
    const avg = this.avgFactor();
    const best = a.map(s => [s.name, this.pct(s.name, 0, avg) - 100]).sort((x, y) => Math.abs(y[1]) - Math.abs(x[1]))[0];
    return Math.abs(best[1]) < 1 ? 'balanced' : best;
  }
  get assignedCount() { return this.owner.size; }
  get dispatchable() { return this.stats.filter(v => v.nhatLe > 0 || v.flipPallets > 0); }

  // ── Hoàn tác ──
  snapshot() { return new Map(this.owner); }
  push(snap = this.snapshot()) { this.undoStack.push(snap); if (this.undoStack.length > UNDO_LIMIT) this.undoStack.shift(); }
  undo() { const s = this.undoStack.pop(); if (!s) return false; this.owner = s; return true; }
  get canUndo() { return this.undoStack.length > 0; }

  // ── Thao tác ──
  give(xe, name) {
    if (!this.byXe.has(xe)) return;
    this.push();
    if (name) this.owner.set(xe, name); else this.owner.delete(xe);
  }

  /** Chia tự động chỉ các xe chưa chia: xe nặng trước, giao người có tải/năng lực thấp nhất. Trả về số xe đã chia. */
  autoPending() {
    const pending = this.dispatchable.filter(v => !this.owner.has(v.xe)).sort((a, b) => b.nhatLe - a.nhatLe);
    if (!pending.length) return 0;
    const ws = this.workers();
    if (!ws.length) return -1;
    this.push();
    this.#greedy(pending, ws);
    return pending.length;
  }
  /** Chia lại TẤT CẢ từ đầu (SmartDistribute): chỉ người đang đi làm */
  autoAll() {
    const ws = this.staff.filter(s => s.active);
    if (!ws.length) return -1;
    this.push();
    this.owner = new Map();
    const list = [...this.dispatchable].sort((a, b) => b.nhatLe - a.nhatLe);
    this.#greedy(list, ws);
    return list.length;
  }
  #greedy(vehicles, ws) {
    const targets = ws.filter(s => weightRatio(s.weight) > 0);
    const pool = targets.length ? targets : ws;
    const loads = new Map(pool.map(s => [s.name, this.load(s.name)]));
    for (const v of vehicles) {
      let best = null, bestF = Infinity;
      for (const s of pool) {
        const w = weightRatio(s.weight), f = w > 0 ? loads.get(s.name) / w : Infinity;
        if (best === null || f < bestF) { best = s; bestF = f; }
      }
      this.owner.set(v.xe, best.name);
      loads.set(best.name, loads.get(best.name) + v.nhatLe);
    }
  }

  /** Cân bằng lại mọi người đang quá tải: dời xe nhỏ nhất sang người rảnh nhất, không đẩy người nhận quá 115%. Trả về số xe đã dời, -1 nếu không ai quá tải. */
  rebalance() {
    const snap = this.snapshot();
    const over = this.workers().filter(s => this.status(s.name) === 'over').sort((a, b) => this.pct(b.name) - this.pct(a.name));
    if (!over.length) return -1;
    let moved = 0;
    for (const from of over) moved += this.#rebalanceOne(from);
    if (moved) this.push(snap);
    return moved;
  }
  #rebalanceOne(from) {
    let moved = 0;
    for (let i = 0; i < 50; i++) {
      if (this.status(from.name) !== 'over') break;
      const mine = this.xeOf(from.name).map(x => [x, this.byXe.get(x)?.nhatLe ?? 0]).sort((a, b) => a[1] - b[1]);
      const targets = this.workers().filter(s => s.name !== from.name && weightRatio(s.weight) > 0)
        .sort((a, b) => this.load(a.name) / weightRatio(a.weight) - this.load(b.name) / weightRatio(b.weight));
      const avg = this.avgFactor();
      let mv = null;
      for (const t of targets) {
        const fit = mine.find(([, e]) => avg > 0 && ((this.load(t.name) + e) / weightRatio(t.weight) / avg) * 100 <= OVER);
        if (fit) { mv = [fit[0], t.name]; break; }
      }
      if (!mv) break;
      this.owner.set(mv[0], mv[1]);
      moved++;
    }
    return moved;
  }
}

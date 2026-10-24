/* 统一月份筛选组件（所有含时间字段的页面共用）
 *
 * 规则（与后端 app/time_range.py 对齐）：
 * - 默认当前月份；支持年/月选择、上一月/下一月；
 * - 快捷筛选：本月 / 上月 / 近3个月 / 近6个月 / 今年 / 全部；
 * - 自定义开始/结束日期（from/to，左闭右开由后端转换）；
 * - 显示当前筛选条件；支持清除筛选；
 * - 筛选状态写入 URL（?month= / ?from=&to= / ?t=all），刷新后按统一规则恢复；
 * - 本地日期计算，不经 UTC 转换，避免时区偏移。
 *
 * 用法：
 *   const MF = new MonthFilter({
 *     container: "month-filter",            // 容器元素 id
 *     onChange: () => { STATE.page = 1; load(); },  // 筛选变化回调
 *     timeLabel: "下单时间",                 // 时间口径显示（可选）
 *   });
 *   const params = MF.getParams();          // {month} | {from,to} | {}
 */
class MonthFilter {
  constructor(opts) {
    this.el = typeof opts.container === "string"
      ? document.getElementById(opts.container) : opts.container;
    if (!this.el) throw new Error("MonthFilter: 容器不存在");
    this.onChange = opts.onChange || (() => {});
    this.timeLabel = opts.timeLabel || "";
    this.state = { mode: "month", month: null, from: "", to: "" };
    this._initFromUrl();
    this._render();
  }

  /* ---------- URL 恢复规则（全站统一） ---------- */
  _initFromUrl() {
    const q = new URLSearchParams(location.search);
    const now = this._now();
    if (q.get("t") === "all") {
      this.state = { mode: "all", month: null, from: "", to: "" };
    } else if (q.get("month")) {
      this.state = { mode: "month", month: q.get("month"), from: "", to: "" };
    } else if (q.get("from") || q.get("to")) {
      this.state = { mode: "range", month: null,
                     from: q.get("from") || "", to: q.get("to") || "" };
    } else {
      this.state = { mode: "month", month: `${now.y}-${String(now.m).padStart(2, "0")}`,
                     from: "", to: "" };
    }
  }

  _syncUrl() {
    const q = new URLSearchParams(location.search);
    ["month", "from", "to", "t"].forEach(k => q.delete(k));
    if (this.state.mode === "month") q.set("month", this.state.month);
    else if (this.state.mode === "range") {
      if (this.state.from) q.set("from", this.state.from);
      if (this.state.to) q.set("to", this.state.to);
      if (!this.state.from && !this.state.to) q.set("t", "all");
    } else if (this.state.mode === "all") q.set("t", "all");
    const s = q.toString();
    history.replaceState(null, "", location.pathname + (s ? `?${s}` : "") + location.hash);
  }

  /* ---------- 日期工具（本地时区，禁止 toISOString） ---------- */
  _now() {
    const d = new Date();
    return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() };
  }
  _fmt(y, m, d) {
    return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  _monthStr(y, m) { return `${y}-${String(m).padStart(2, "0")}`; }
  _shiftMonth(base, delta) {
    const [y, m] = base.split("-").map(Number);
    const total = y * 12 + (m - 1) + delta;
    return { y: Math.floor(total / 12), m: (total % 12 + 12) % 12 + 1 };
  }
  _monthEnd(y, m) {
    const last = new Date(y, m, 0).getDate();
    return this._fmt(y, m, last);
  }

  /* ---------- 渲染 ---------- */
  _render() {
    const now = this._now();
    const years = [];
    for (let y = now.y - 3; y <= now.y + 1; y++) years.push(y);
    const cur = this.state.mode === "month" && this.state.month
      ? this.state.month.split("-").map(Number) : [now.y, now.m];
    this.el.className = "month-filter";
    this.el.innerHTML = `
      <div class="mf-main">
        <button class="mf-nav" data-act="prev" title="上一月">‹</button>
        <select class="mf-year" title="年份">${years.map(y =>
          `<option value="${y}" ${y === cur[0] ? "selected" : ""}>${y}年</option>`).join("")}</select>
        <select class="mf-month" title="月份">${Array.from({ length: 12 }, (_, i) =>
          `<option value="${i + 1}" ${i + 1 === cur[1] ? "selected" : ""}>${i + 1}月</option>`).join("")}</select>
        <button class="mf-nav" data-act="next" title="下一月">›</button>
        <span class="mf-quick">
          <button data-q="this_month">本月</button>
          <button data-q="last_month">上月</button>
          <button data-q="last_3m">近3个月</button>
          <button data-q="last_6m">近6个月</button>
          <button data-q="this_year">今年</button>
          <button data-q="all">全部</button>
        </span>
        <span class="mf-range">
          <input type="date" class="mf-from" value="${this.state.from}" title="自定义开始日期">
          <span class="mf-tilde">~</span>
          <input type="date" class="mf-to" value="${this.state.to}" title="自定义结束日期">
          <button class="btn secondary mf-apply" data-act="apply">应用</button>
        </span>
        <button class="btn secondary mf-clear" data-act="clear" title="清除时间筛选，显示全部">清除筛选</button>
      </div>
      <div class="mf-label-row">
        <span class="mf-label"></span>
        ${this.timeLabel ? `<span class="mf-caliber">时间口径：${this.timeLabel}</span>` : ""}
      </div>`;
    this._syncLabel();
    this._bind();
  }

  _bind() {
    this.el.querySelector('[data-act="prev"]').onclick = () => {
      const base = this.state.mode === "month" && this.state.month
        ? this.state.month : this._monthStr(this._now().y, this._now().m);
      const { y, m } = this._shiftMonth(base, -1);
      this._set({ mode: "month", month: this._monthStr(y, m), from: "", to: "" });
    };
    this.el.querySelector('[data-act="next"]').onclick = () => {
      const base = this.state.mode === "month" && this.state.month
        ? this.state.month : this._monthStr(this._now().y, this._now().m);
      const { y, m } = this._shiftMonth(base, 1);
      this._set({ mode: "month", month: this._monthStr(y, m), from: "", to: "" });
    };
    this.el.querySelector(".mf-year").onchange = e => {
      const m = Number(this.el.querySelector(".mf-month").value);
      this._set({ mode: "month", month: this._monthStr(Number(e.target.value), m), from: "", to: "" });
    };
    this.el.querySelector(".mf-month").onchange = e => {
      const y = Number(this.el.querySelector(".mf-year").value);
      this._set({ mode: "month", month: this._monthStr(y, Number(e.target.value)), from: "", to: "" });
    };
    this.el.querySelectorAll(".mf-quick button").forEach(b => {
      b.onclick = () => this._applyQuick(b.dataset.q);
    });
    this.el.querySelector('[data-act="apply"]').onclick = () => {
      const from = this.el.querySelector(".mf-from").value;
      const to = this.el.querySelector(".mf-to").value;
      if (from && to && from > to) {
        toast("开始日期不得晚于结束日期", "error");
        return;
      }
      this._set({ mode: "range", month: null, from, to });
    };
    this.el.querySelector('[data-act="clear"]').onclick = () => {
      this._set({ mode: "all", month: null, from: "", to: "" });
    };
  }

  _applyQuick(kind) {
    const now = this._now();
    const curMonth = this._monthStr(now.y, now.m);
    if (kind === "all") return this._set({ mode: "all", month: null, from: "", to: "" });
    if (kind === "this_month")
      return this._set({ mode: "month", month: curMonth, from: "", to: "" });
    if (kind === "last_month") {
      const { y, m } = this._shiftMonth(curMonth, -1);
      return this._set({ mode: "month", month: this._monthStr(y, m), from: "", to: "" });
    }
    if (kind === "last_3m" || kind === "last_6m") {
      const delta = kind === "last_3m" ? -2 : -5;
      const { y, m } = this._shiftMonth(curMonth, delta);
      return this._set({ mode: "range", month: null,
                         from: this._fmt(y, m, 1), to: this._monthEnd(now.y, now.m) });
    }
    if (kind === "this_year") {
      return this._set({ mode: "range", month: null,
                         from: this._fmt(now.y, 1, 1), to: this._fmt(now.y, 12, 31) });
    }
  }

  _set(next) {
    this.state = Object.assign({}, this.state, next);
    this._syncUrl();
    this._render();
    this.onChange(this._summary());
  }

  _summary() {
    if (this.state.mode === "month") return { kind: "month", month: this.state.month };
    if (this.state.mode === "range") return { kind: "range", from: this.state.from, to: this.state.to };
    return { kind: "all" };
  }

  /* ---------- 对外：查询参数与标签 ---------- */
  reset() {
    const now = this._now();
    this._set({ mode: "month", month: this._monthStr(now.y, now.m), from: "", to: "" });
  }

  getParams() {
    if (this.state.mode === "month") return { month: this.state.month };
    if (this.state.mode === "range") {
      const p = {};
      if (this.state.from) p.from = this.state.from;
      if (this.state.to) p.to = this.state.to;
      return p;
    }
    return {};
  }

  getLabel() {
    if (this.state.mode === "month") {
      const [y, m] = this.state.month.split("-").map(Number);
      return `${y}年${m}月`;
    }
    if (this.state.mode === "range") {
      if (this.state.from && this.state.to) return `${this.state.from} ~ ${this.state.to}`;
      if (this.state.from) return `${this.state.from} 起`;
      if (this.state.to) return `至 ${this.state.to}`;
      return "全部时间";
    }
    return "全部时间";
  }

  _syncLabel() {
    this.el.querySelector(".mf-label").textContent = `当前筛选：${this.getLabel()}`;
    const quick = this._activeQuick();
    this.el.querySelectorAll(".mf-quick button").forEach(b =>
      b.classList.toggle("cur", b.dataset.q === quick));
  }

  _activeQuick() {
    const s = this.state;
    if (s.mode === "all") return "all";
    if (s.mode !== "range" || !s.from || !s.to) return null;
    const now = this._now();
    const curMonth = this._monthStr(now.y, now.m);
    if (s.from === this._fmt(now.y, 1, 1) && s.to === this._fmt(now.y, 12, 31))
      return "this_year";
    for (const [q, delta] of [["last_3m", -2], ["last_6m", -5]]) {
      const { y, m } = this._shiftMonth(curMonth, delta);
      if (s.from === this._fmt(y, m, 1) && s.to === this._monthEnd(now.y, now.m))
        return q;
    }
    return null;
  }
}

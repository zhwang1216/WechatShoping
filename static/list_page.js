/* ==================== 通用列表页 ListPage ====================
   把订单/退款/供应商成本/异常/日志等页面的重复逻辑收敛到一处：
   - 月份筛选（MonthFilter）+ 成员/店铺筛选（StoreScope）
   - 顶部筛选条件 → 查询参数（声明式 filters，不再逐页手写 queryStr）
   - 请求、加载中/空/错误态、表体渲染、分页（复用 CommonTable）
   - 统计卡片、导出、重置、回车/下拉自动查询

   用法（页面保留静态表头与筛选 HTML，仅声明数据与列）：
     const PAGE = new ListPage({
       tableId: "orders-table", url: "/api/orders", itemsKey: "orders",
       rowKey: "order_id",
       columns: [{ title: "订单号", render: r => esc(r.order_id) }, ...],
       filters: [{ id: "f-order-id", param: "order_id" }, ...],
       stats: s => renderStatCards(document.getElementById("stat-grid"), [...]),
       exportUrl: "/api/orders/export",
       emptyHtml: mf => `...`,
     });
*/
function ListPage(options) {
  this.opts = options || {};
  this.stateEl = document.getElementById(this.opts.stateEl || "table-state");
  this.pagerEl = document.getElementById(this.opts.pagerEl || "pager");
  this.statEl = this.opts.statEl
    ? document.getElementById(this.opts.statEl)
    : document.getElementById("stat-grid");
  this.rows = [];
  this.data = null;

  var self = this;
  this.mf = new MonthFilter({
    container: this.opts.monthContainer || "month-filter",
    timeLabel: this.opts.timeLabel || "",
    onChange: function () { self.reload(); },
  });
  this.ss = new StoreScope({
    container: this.opts.storeContainer || "store-scope",
    onChange: function () { self.reload(); },
  });

  this.table = new CommonTable({
    tableId: this.opts.tableId,
    pagerEl: this.pagerEl,
    stateEl: this.stateEl,
    columns: this.opts.columns,
    rowKey: this.opts.rowKey,
    staticHead: true,             // 表头保留在 HTML 中，组件只渲染表体
    emptyText: "暂无数据",
    pagination: { pageSize: this.opts.pageSize || 50 },
    dataSource: {
      url: this.opts.url,
      itemsKey: this.opts.itemsKey,
      buildParams: function () { return self.collectParams(); },
    },
    events: {
      onRefresh: function (r) { self._onRefresh(r); },
      onError: function (r) { self._onError(r.error); },
    },
  });

  this.bind();
}

/* 顶部筛选 → 请求参数。filters 声明：
   {id: 元素 id, param: 后端参数名, always: 是否总是携带, reset: 重置值} */
ListPage.prototype.collectParams = function () {
  var p = {};
  var mf = this.mf.getParams();
  Object.keys(mf).forEach(function (k) { p[k] = mf[k]; });
  this.ss.appendTo(p);
  var filters = this.opts.filters || [];
  for (var i = 0; i < filters.length; i++) {
    var f = filters[i];
    var el = document.getElementById(f.id);
    if (!el) continue;
    var v = (el.value || "").trim();
    if (v !== "" || f.always) p[f.param || f.id] = v;
  }
  if (typeof this.opts.extraParams === "function") {
    var extra = this.opts.extraParams() || {};
    Object.keys(extra).forEach(function (k) {
      if (extra[k] !== null && extra[k] !== undefined && extra[k] !== "") {
        p[k] = extra[k];
      }
    });
  }
  return p;
};

ListPage.prototype.reload = function () {
  this.table.emptyText = typeof this.opts.emptyText === "function"
    ? this.opts.emptyText(this.mf) : (this.opts.emptyText || "暂无数据");
  this.table.refresh({ resetPage: true, filters: this.collectParams() });
};

ListPage.prototype.refresh = function () {
  this.table.refresh({ filters: this.collectParams() });
};

ListPage.prototype.reset = function () {
  var filters = this.opts.filters || [];
  for (var i = 0; i < filters.length; i++) {
    var el = document.getElementById(filters[i].id);
    if (!el) continue;
    el.value = filters[i].reset != null ? filters[i].reset : "";
  }
  if (typeof this.opts.onReset === "function") this.opts.onReset(this);
  this.mf.reset();   // MonthFilter.reset() 触发 onChange → reload()
};

ListPage.prototype.exportFile = function () {
  if (!this.opts.exportUrl) return;
  var qs = new URLSearchParams(this.collectParams()).toString();
  var url = this.opts.exportUrl + (this.opts.exportUrl.indexOf("?") >= 0 ? "&" : "?") + qs;
  window.open(url);
  if (typeof toast === "function") toast("导出文件生成中…", "info");
};

ListPage.prototype._onRefresh = function (r) {
  this.data = r;
  this.rows = r.data || [];
  this.renderStats(r.statistics);
  if (typeof this.opts.onData === "function") this.opts.onData(r, this);
  var hint = document.getElementById("table-hint");
  if (hint) hint.textContent = r.total ? ("第 " + r.page + "/" + r.pages + " 页") : "";
  if (!this.rows.length && typeof this.opts.emptyHtml === "function") {
    this.stateEl.innerHTML = this.opts.emptyHtml(this.mf);
  }
};

ListPage.prototype._onError = function (err) {
  this.rows = [];
  this.renderStats(null);
  var hint = document.getElementById("table-hint");
  if (hint) hint.textContent = "";
  if (!this.stateEl) return;
  var self = this;
  this.stateEl.innerHTML =
    '<div class="error-state"><div class="msg">' +
    esc((err && err.message) || "请求失败，请稍后重试") +
    '</div><button class="btn secondary" type="button">重试</button></div>';
  var retry = this.stateEl.querySelector("button");
  if (retry) retry.onclick = function () { self.reload(); };
};

ListPage.prototype.renderStats = function (statistics) {
  if (!this.statEl) return;
  if (typeof this.opts.stats === "function") {
    var html = this.opts.stats(statistics, this);
    if (html !== false) this.statEl.innerHTML = html || "";
  }
};

ListPage.prototype.bind = function () {
  var self = this;
  var queryBtn = document.getElementById("btn-query");
  if (queryBtn) queryBtn.onclick = function () { self.reload(); };
  var resetBtn = document.getElementById("btn-reset");
  if (resetBtn) resetBtn.onclick = function () { self.reset(); };
  var exportBtn = document.getElementById("btn-export");
  if (exportBtn && this.opts.exportUrl) {
    exportBtn.onclick = function () { self.exportFile(); };
  }
  var grid = document.querySelector(".filter-grid");
  if (grid) {
    grid.addEventListener("change", function (e) {
      if (e.target.matches("select")) self.reload();
    });
    grid.addEventListener("keydown", function (e) {
      if (e.key === "Enter" && e.target.matches("input,select")) {
        e.preventDefault();
        self.reload();
      }
    });
  }
};

/* 统计卡片统一渲染：cards 为 [label, value, className, sub] 或对象数组；
   note 为卡片下方的口径说明（可含 HTML，仅接受页面内静态文案）。
   传入 el 时直接写入并返回 HTML；不传 el 时仅返回 HTML，供 ListPage.stats 使用。 */
function renderStatCards(el, cards, note) {
  if (!cards) { if (el) el.innerHTML = ""; return ""; }
  var html = cards.map(function (c) {
    var label, value, cls = "", sub = "", title = "";
    if (Array.isArray(c)) {
      label = c[0]; value = c[1]; cls = c[2] || ""; sub = c[3] || "";
      title = c[4] || "";
    } else {
      label = c.label; value = c.value; cls = c.cls || "";
      sub = c.sub || ""; title = c.title || "";
    }
    return '<div class="stat-item"' + (title ? ' title="' + esc(title) + '"' : "") + ">" +
      '<div class="label">' + label + "</div>" +
      '<div class="value ' + cls + '">' + value + "</div>" +
      (sub ? '<div class="sub">' + sub + "</div>" : "") + "</div>";
  }).join("");
  if (note) html += '<div class="stat-note">' + note + "</div>";
  if (el) el.innerHTML = html;
  return html;
}

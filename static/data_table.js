/* ==================== DataTable 统一表格组件 ====================
 * 基于 CommonTable 的数据层，统一 DOM 结构、滚动、列管理与格式化。
 *
 * API:
 *   new DataTable({
 *     tableId: "orders-table",       // 必填（或 mount + 空 table）
 *     columns: [...],                // 同 CommonTable，扩展 type / sortable / fixed: "right"
 *     rowKey, dataSource, pagination, stateEl, pagerEl, events,
 *     storageKey: "orders-table",    // 列宽/显隐 localStorage 键
 *     listPage: ListPage 实例,       // 可选：用于表头排序联动筛选
 *     sortFilterIds: { field: "f-order-by", dir: "f-order-dir" },
 *   })
 *
 * 列 type: orderId | amount | datetime | status | text | custom（默认 custom，用 render）
 * 列 fixed: "right" → 操作列固定
 */
(function (global) {
  "use strict";

  var COL_STORE_PREFIX = "wxr_dt_cols_";
  var WIDTH_STORE_PREFIX = "wxr_dt_width_";

  function DataTable(options) {
    if (!options || !options.columns || !options.columns.length) {
      console.warn("[DataTable] columns required");
      return null;
    }
    if (!options.tableId && !options.mount) {
      console.warn("[DataTable] tableId or mount required");
      return null;
    }
    this.options = options;
    this.storageKey = options.storageKey || options.tableId || "data-table";
    this.columns = DataTable.normalizeColumns(options.columns);
    this._applyVisibilityFromStore();
    this._initDom();
    var self = this;
    this._ct = new CommonTable(Object.assign({}, options, {
      columns: this.columns,
      el: this.tableEl,
      staticHead: false,
      stateEl: this.stateEl,
      pagerEl: this.pagerEl,
      stickyHeader: false,
      scrollX: false,
      scrollY: false,
    }));
    this._bindChrome();
    this.refresh = function (p) { return self._ct.refresh(p); };
    this.setData = function (d) { return self._ct.setData(d); };
    this.destroy = function () {
      self._teardownChrome();
      if (self._ct) self._ct.destroy();
    };
    return this;
  }

  DataTable.normalizeColumns = function (columns) {
    return columns.map(function (c) {
      var col = Object.assign({}, c);
      if (col.fixed === true) col.fixed = "right";
      if (col.type && !col.render) {
        col.render = DataTable.cellByType(col.type, col);
      }
      if (col.type === "orderId") {
        col.className = ((col.className || "") + " cell-order-id").trim();
        col.ellipsis = true;
        col.minWidth = col.minWidth || 140;
      }
      if (col.type === "amount") col.align = col.align || "right";
      if (col.fixed === "right") {
        col.className = ((col.className || "") + " col-fixed-right").trim();
      }
      return col;
    });
  };

  DataTable.cellByType = function (type, col) {
    var key = col.key;
    return function (row) {
      var val = key ? row[key] : undefined;
      switch (type) {
        case "orderId":
          return DataTable.format.orderId(val != null ? val : row.order_id);
        case "amount":
          return DataTable.format.amount(val);
        case "datetime":
          return DataTable.format.datetime(val);
        case "status":
          return typeof col.statusRender === "function"
            ? col.statusRender(val, row)
            : esc(val != null ? val : "—");
        case "text":
          if (col.ellipsis) {
            return '<span class="cell-ellipsis" title="' + esc(val != null ? val : "") + '">' +
              esc(val != null ? val : "") + "</span>";
          }
          return esc(val != null ? val : "");
        default:
          return esc(val != null ? val : "");
      }
    };
  };

  DataTable.format = {
    orderId: function (id) {
      if (id == null || id === "") return "—";
      var s = String(id);
      return '<span class="copyable cell-order-id" title="' + esc(s) + '" data-copy="' +
        esc(s) + '">' + esc(s) + "</span>";
    },
    amount: function (v, tone) {
      var cls = "amount amount--default";
      if (tone === "success") cls = "amount amount--success";
      else if (tone === "warning") cls = "amount amount--warning";
      else if (tone === "danger") cls = "amount amount--danger";
      else if (tone === "muted") cls = "amount amount--muted";
      var inner = typeof fmtMoney === "function" ? fmtMoney(v) : esc(v);
      return '<span class="' + cls + '">' + inner + "</span>";
    },
    datetime: function (v) {
      return typeof fmtTime === "function" ? fmtTime(v) : esc(v || "—");
    },
  };

  DataTable.prototype._initDom = function () {
    var table = this.options.tableId
      ? document.getElementById(this.options.tableId)
      : null;
    if (!table && this.options.mount) {
      var mount = typeof this.options.mount === "string"
        ? document.querySelector(this.options.mount) : this.options.mount;
      mount.innerHTML =
        '<div class="data-table__viewport"><table id="' + esc(this.options.tableId || "") +
        '"><thead></thead><tbody></tbody></table></div>';
      table = mount.querySelector("table");
    }
    if (!table) {
      console.warn("[DataTable] table element not found");
      return;
    }
    this.tableEl = table;
    var root = table.closest(".data-table");
    if (!root) {
      var wrap = document.createElement("div");
      wrap.className = "data-table";
      wrap.dataset.tableId = this.options.tableId || "";
      var toolbar = document.createElement("div");
      toolbar.className = "data-table__toolbar";
      var viewport = document.createElement("div");
      viewport.className = "data-table__viewport";
      var parent = table.parentElement;
      if (parent) {
        parent.insertBefore(wrap, table);
        parent.removeChild(table);
      }
      wrap.appendChild(toolbar);
      viewport.appendChild(table);
      wrap.appendChild(viewport);
      root = wrap;
    }
    this.rootEl = root;
    var host = table.closest(".table-wrap");
    if (host) {
      host.classList.add("data-table-host");
    }
    this.viewportEl = table.closest(".data-table__viewport") || root.querySelector(".data-table__viewport");
    if (!this.viewportEl) {
      this.viewportEl = document.createElement("div");
      this.viewportEl.className = "data-table__viewport";
      table.parentNode.insertBefore(this.viewportEl, table);
      this.viewportEl.appendChild(table);
    }
    this.toolbarEl = root.querySelector(".data-table__toolbar");
    if (!this.toolbarEl) {
      this.toolbarEl = document.createElement("div");
      this.toolbarEl.className = "data-table__toolbar";
      root.insertBefore(this.toolbarEl, this.viewportEl);
    }
    var card = root.closest(".table-card");
    this.stateEl = this.options.stateEl
      || (card && card.querySelector("#table-state"))
      || document.getElementById("table-state");
    this.pagerEl = this.options.pagerEl
      || (card && card.querySelector("#pager"))
      || document.getElementById("pager");
    if (this.stateEl && !this.stateEl.classList.contains("data-table__state")) {
      this.stateEl.classList.add("data-table__state");
    }
    if (this.pagerEl) this.pagerEl.classList.add("data-table__pager");
    this._widths = this._loadWidths();
  };

  DataTable.prototype._applyVisibilityFromStore = function () {
    try {
      var raw = localStorage.getItem(COL_STORE_PREFIX + this.storageKey);
      if (!raw) return;
      var hidden = JSON.parse(raw);
      if (!Array.isArray(hidden)) return;
      this.columns.forEach(function (c) {
        if (c.key && hidden.indexOf(c.key) >= 0) c.hidden = true;
      });
    } catch (e) { /* 忽略 */ }
  };

  DataTable.prototype._saveVisibility = function () {
    try {
      var hidden = this.columns.filter(function (c) { return c.hidden && c.key; }).map(function (c) { return c.key; });
      localStorage.setItem(COL_STORE_PREFIX + this.storageKey, JSON.stringify(hidden));
    } catch (e) { /* 忽略 */ }
  };

  DataTable.prototype._loadWidths = function () {
    try {
      var raw = localStorage.getItem(WIDTH_STORE_PREFIX + this.storageKey);
      return raw ? JSON.parse(raw) : {};
    } catch (e) { return {}; }
  };

  DataTable.prototype._saveWidths = function () {
    try {
      localStorage.setItem(WIDTH_STORE_PREFIX + this.storageKey, JSON.stringify(this._widths));
    } catch (e) { /* 忽略 */ }
  };

  DataTable.prototype._bindChrome = function () {
    var self = this;
    this._renderToolbar();
    this._ct._eventBus.on("onRefresh", function () {
      self._syncHeadSortState();
      self._applyColumnWidths();
    });
    this.viewportEl.addEventListener("click", function (e) {
      var th = e.target.closest("th.sortable");
      if (!th || !self._ct) return;
      if (e.target.closest(".dt-col-resizer")) return;
      var field = th.dataset.sort;
      if (!field) return;
      var next = "desc";
      if (th.classList.contains("dt-sorted-desc")) next = "asc";
      else if (th.classList.contains("dt-sorted-asc")) next = "desc";
      self._applySort(field, next);
    });
  };

  DataTable.prototype._applySort = function (field, dir) {
    var sf = this.options.sortFilterIds;
    if (sf && sf.field) {
      var el = document.getElementById(sf.field);
      var elDir = sf.dir && document.getElementById(sf.dir);
      if (el) el.value = field;
      if (elDir) elDir.value = dir;
      if (this.options.listPage) {
        this.options.listPage.reload();
        return;
      }
    }
    this._ct.refresh({ sortField: field, sortOrder: dir, resetPage: true });
  };

  DataTable.prototype._syncHeadSortState = function () {
    var sf = this.options.sortFilterIds;
    var field = "";
    var dir = "desc";
    if (sf && sf.field) {
      var el = document.getElementById(sf.field);
      var elDir = sf.dir && document.getElementById(sf.dir);
      if (el) field = el.value || "";
      if (elDir) dir = elDir.value || "desc";
    } else {
      field = this._ct.state.sortField || "";
      dir = this._ct.state.sortOrder || "desc";
    }
    var ths = this.tableEl.querySelectorAll("thead th[data-sort]");
    ths.forEach(function (th) {
      th.classList.remove("dt-sorted-asc", "dt-sorted-desc");
      if (th.dataset.sort === field) {
        th.classList.add(dir === "asc" ? "dt-sorted-asc" : "dt-sorted-desc");
      }
    });
  };

  DataTable.prototype._renderToolbar = function () {
    if (!this.toolbarEl) return;
    var self = this;
    this.toolbarEl.innerHTML =
      '<div class="more-wrap" style="position:relative">' +
      '<button type="button" class="btn btn-sm secondary" id="dt-cols-' + esc(this.storageKey) + '">列显示</button>' +
      '<div class="dt-col-panel" id="dt-col-panel-' + esc(this.storageKey) + '"></div></div>';
    var btn = this.toolbarEl.querySelector("button");
    var panel = this.toolbarEl.querySelector(".dt-col-panel");
  panel.innerHTML = this.columns.map(function (c, i) {
      if (!c.key || c.fixed === "right") return "";
      var id = "dt-col-" + self.storageKey + "-" + i;
      return '<label><input type="checkbox" data-col-idx="' + i + '" id="' + id + '"' +
        (c.hidden ? "" : " checked") + "> " + esc(c.title) + "</label>";
    }).join("");
    btn.onclick = function (e) {
      e.stopPropagation();
      panel.classList.toggle("open");
    };
    document.addEventListener("click", this._docClick = function () { panel.classList.remove("open"); });
    panel.addEventListener("click", function (e) { e.stopPropagation(); });
    panel.querySelectorAll("input[type=checkbox]").forEach(function (cb) {
      cb.onchange = function () {
        var idx = parseInt(cb.dataset.colIdx, 10);
        if (isNaN(idx) || !self.columns[idx]) return;
        self.columns[idx].hidden = !cb.checked;
        self._saveVisibility();
        self._ct.columns = self.columns;
        self._ct.adapter.renderHead(self.columns);
        self._ct.setData(self._ct.state.data);
        self._bindResizeHandles();
      };
    });
    this._ct.adapter.renderHead(this.columns);
    this._bindResizeHandles();
    this._applyColumnWidths();
  };

  DataTable.prototype._bindResizeHandles = function () {
    var self = this;
    var ths = this.tableEl.querySelectorAll("thead th");
    ths.forEach(function (th) {
      if (th.querySelector(".dt-col-resizer")) return;
      th.style.position = "sticky";
      var colKey = th.dataset.colKey;
      var col = colKey ? self.columns.find(function (c) { return c.key === colKey; }) : null;
      var h = document.createElement("span");
      h.className = "dt-col-resizer";
      h.addEventListener("mousedown", function (e) {
        e.preventDefault();
        e.stopPropagation();
        var startX = e.clientX;
        var startW = th.offsetWidth;
        th.classList.add("dt-resizing");
        function move(ev) {
          var w = Math.max(60, startW + (ev.clientX - startX));
          th.style.width = w + "px";
          if (col && col.key) self._widths[col.key] = w;
        }
        function up() {
          th.classList.remove("dt-resizing");
          document.removeEventListener("mousemove", move);
          document.removeEventListener("mouseup", up);
          self._saveWidths();
          self._syncBodyColWidths();
        }
        document.addEventListener("mousemove", move);
        document.addEventListener("mouseup", up);
      });
      th.appendChild(h);
      if (!th.querySelector(".dt-sort-ind") && th.classList.contains("sortable")) {
        var ind = document.createElement("span");
        ind.className = "dt-sort-ind";
        th.appendChild(ind);
      }
    });
  };

  DataTable.prototype._applyColumnWidths = function () {
    var self = this;
    var ths = this.tableEl.querySelectorAll("thead th");
    ths.forEach(function (th) {
      var key = th.dataset.colKey;
      if (key && self._widths[key]) th.style.width = self._widths[key] + "px";
    });
    this._syncBodyColWidths();
  };

  DataTable.prototype._syncBodyColWidths = function () {
    var ths = this.tableEl.querySelectorAll("thead th");
    var rows = this.tableEl.querySelectorAll("tbody tr");
    if (!rows.length) return;
    var widths = [];
    ths.forEach(function (th) { widths.push(th.style.width || th.offsetWidth + "px"); });
    rows.forEach(function (tr) {
      var tds = tr.querySelectorAll("td");
      tds.forEach(function (td, i) {
        if (widths[i]) td.style.width = widths[i];
      });
    });
  };

  DataTable.prototype._teardownChrome = function () {
    if (this._docClick) document.removeEventListener("click", this._docClick);
  };

  global.DataTable = DataTable;
})(typeof window !== "undefined" ? window : this);

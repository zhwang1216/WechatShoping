/* ==================== 公共表格组件 CommonTable ====================
   统一全站表格行为：表头固定(R1-R7)、双向滚动、列配置驱动渲染、
   统一事件/分页/数据适配，后端零改动。
   复用既有：esc/renderPager(API.js)、R1-R7 公共表格滚动规范(style.css)。
   不内置业务字段/格式/接口——由业务页面在列配置 formatter/render 中提供。 */

/* ---- EventBus 统一事件转发 ---- */
function _EventBus() {
  this._handlers = {};
}
_EventBus.prototype.on = function (name, handler) {
  (this._handlers[name] = this._handlers[name] || []).push(handler);
};
_EventBus.prototype.off = function (name) {
  delete this._handlers[name];
};
_EventBus.prototype.emit = function (name, payload) {
  var list = this._handlers[name];
  if (!list) return;
  for (var i = 0; i < list.length; i++) {
    try { list[i](payload); }
    catch (e) {
      console.error("[CommonTable] event " + name + " handler error:", e);
      if (typeof toast === "function") toast("表格事件处理异常", "error");
    }
  }
};

/* ---- ColumnRenderer 列配置渲染器 ---- */
var ColumnRenderer = {
  renderHead: function (columns) {
    var ths = [];
    for (var i = 0; i < columns.length; i++) {
      var c = columns[i];
      if (c.hidden) continue;
      var cls = [];
      if (c.align === "right" || c.align === "center") cls.push(c.align === "right" ? "num" : "center");
      if (c.className) cls.push(c.className);
      if (c.fixed === "right") cls.push("col-fixed-right");
      else if (c.fixed === "left" || c.fixed === true) cls.push("col-op");
      var style = "";
      if (c.width != null) style += "width:" + (typeof c.width === "number" ? c.width + "px" : c.width) + ";";
      if (c.minWidth != null) style += "min-width:" + (typeof c.minWidth === "number" ? c.minWidth + "px" : c.minWidth) + ";";
      var sortKey = null;
      if (c.sortable) {
        if (typeof c.sortable === "object") sortKey = c.sortable.field || c.key;
        else if (typeof c.sortable === "string") sortKey = c.sortable;
        else sortKey = c.key;
      }
      var sortAttr = sortKey ? ' data-sort="' + esc(sortKey) + '"' : "";
      var sortClass = c.sortable ? " sortable" : "";
      var keyAttr = c.key ? ' data-col-key="' + esc(String(c.key)) + '"' : "";
      ths.push("<th" + (cls.length ? ' class="' + cls.join(" ") + sortClass + '"' : ' class="' + sortClass.trim() + '"') +
        (style ? ' style="' + style + '"' : "") + sortAttr + keyAttr + ">" + esc(c.title || "") + "</th>");
    }
    return "<thead><tr>" + ths.join("") + "</tr></thead>";
  },
  renderBody: function (data, columns, rowKey) {
    if (!data || !data.length) return "<tbody></tbody>";
    var rows = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rk = typeof rowKey === "function" ? rowKey(row) : (row[rowKey] != null ? row[rowKey] : i);
      var tds = [];
      for (var j = 0; j < columns.length; j++) {
        var c = columns[j];
        if (c.hidden) continue;
        var cls = [];
        if (c.align === "right" || c.align === "center") cls.push(c.align === "right" ? "num" : "center");
        if (c.className) cls.push(c.className);
        if (c.fixed === "right") cls.push("col-fixed-right");
      else if (c.fixed === "left" || c.fixed === true) cls.push("col-op");
        var content = this.renderCell(row, c, i);
        var style = "";
        if (c.ellipsis) style += "overflow:hidden;text-overflow:ellipsis;";
        tds.push("<td" + (cls.length ? ' class="' + cls.join(" ") + '"' : "") +
          (style ? ' style="' + style + '"' : "") + ">" + content + "</td>");
      }
      rows.push('<tr data-row-key="' + esc(String(rk)) + '">' + tds.join("") + "</tr>");
    }
    return "<tbody>" + rows.join("") + "</tbody>";
  },
  renderCell: function (row, column, index) {
    try {
      if (typeof column.render === "function") return column.render(row, index);
      var val = row[column.key];
      if (typeof column.formatter === "function") return column.formatter(val, row, index);
      return esc(val != null ? val : "");
    } catch (e) {
      console.warn("[CommonTable] renderCell error for key=" + column.key + ":", e);
      return "";
    }
  }
};

/* ---- PagerBinder 分页绑定（复用既有 renderPager） ---- */
function _PagerBinder() {}
_PagerBinder.prototype.render = function (el, pagination, eventBus) {
  if (pagination === false || !el) return;
  var page = pagination.current || 1;
  var size = pagination.pageSize || 50;
  var total = pagination.total || 0;
  var pages = pagination.pages || (size > 0 ? Math.ceil(total / size) : 1);
  if (typeof renderPager === "function") {
    renderPager(el, { page: page, pages: pages, total: total, size: size }, {
      onPage: function (p) { eventBus.emit("onPageChange", { page: p, pageSize: size }); },
      onSize: function (s) { eventBus.emit("onPageChange", { page: 1, pageSize: s }); }
    });
  }
};

/* ---- TableDataAdapter 数据请求/响应适配层 ---- */
function _TableDataAdapter() {}
_TableDataAdapter.prototype.buildRequest = function (params, dataSource) {
  var req = {};
  req.page = params.page || 1;
  req.page_size = params.pageSize || 50;
  if (params.sortField) {
    var sortKey = dataSource && dataSource.sortFieldMap ? (dataSource.sortFieldMap[params.sortField] || params.sortField) : params.sortField;
    req.order_by = sortKey;
    req.order_dir = params.sortOrder || "desc";
  }
  if (dataSource && typeof dataSource.buildParams === "function") {
    var extra = dataSource.buildParams(params.filters || {});
    for (var k in extra) { if (extra[k] != null && extra[k] !== "") req[k] = extra[k]; }
  }
  return req;
};
_TableDataAdapter.prototype.parseResponse = function (rawData, itemsKey) {
  if (!rawData || typeof rawData !== "object") return { items: [], total: 0, page: 1, pageSize: 50, pages: 1, success: false };
  var items = rawData[itemsKey] || [];
  if (!Array.isArray(items)) { console.warn("[CommonTable] itemsKey '" + itemsKey + "' not array"); items = []; }
  var pageSize = rawData.size || rawData.page_size || 50;
  var total = rawData.total || 0;
  var page = rawData.page || 1;
  var pages = rawData.pages || (pageSize > 0 ? Math.ceil(total / pageSize) : 1);
  return { items: items, total: total, page: page, pageSize: pageSize, pages: pages, success: true, statistics: rawData.statistics };
};

/* ---- NativeTableAdapter 主表格适配器 ---- */
function _NativeTableAdapter() {}
_NativeTableAdapter.prototype.bindTable = function (el) {
  this.tableEl = el;
  this.tbodyEl = el.querySelector("tbody");
  this.theadEl = el.querySelector("thead");
  this.wrapEl = el.closest(".table-wrap");
  return this;
};
_NativeTableAdapter.prototype.renderHead = function (columns) {
  if (!this.theadEl) return;
  var full = ColumnRenderer.renderHead(columns);
  var m = full.match(/<tr>([\s\S]*)<\/tr>/);
  this.theadEl.innerHTML = m ? "<tr>" + m[1] + "</tr>" : full;
};
_NativeTableAdapter.prototype.renderBody = function (data, columns, rowKey) {
  if (this.tbodyEl) this.tbodyEl.innerHTML = ColumnRenderer.renderBody(data, columns, rowKey).replace(/<\/?tbody>/g, "");
};
_NativeTableAdapter.prototype.showLoading = function (stateEl) {
  if (stateEl) stateEl.innerHTML = '<div class="loading-state"><div class="spinner"></div>加载中…</div>';
  if (this.tbodyEl) this.tbodyEl.innerHTML = "";
};
_NativeTableAdapter.prototype.showEmpty = function (stateEl, text) {
  if (stateEl) stateEl.innerHTML = '<div class="empty-state"><div class="icon">📋</div>' + esc(text || "暂无数据") + "</div>";
};
_NativeTableAdapter.prototype.showError = function (stateEl, msg) {
  if (stateEl) stateEl.innerHTML = '<div class="error-state">' + esc(msg || "加载失败") + "</div>";
};
_NativeTableAdapter.prototype.bindRowClick = function (eventBus) {
  var self = this;
  if (this.tbodyEl) {
    this.tbodyEl.addEventListener("click", function (e) {
      var tr = e.target.closest("tr");
      if (tr && tr.dataset.rowKey != null) eventBus.emit("onRowClick", { rowKey: tr.dataset.rowKey, row: tr });
    });
  }
};
_NativeTableAdapter.prototype.destroy = function () {
  this.tableEl = this.tbodyEl = this.theadEl = this.wrapEl = null;
};

/* ---- ModalTableAdapter 弹窗嵌套适配器 ---- */
function _ModalTableAdapter() {}
_ModalTableAdapter.prototype.bindContainer = function (el) {
  this.containerEl = el;
  return this;
};
_ModalTableAdapter.prototype.renderTable = function (data, columns, rowKey) {
  if (!this.containerEl) return;
  this.containerEl.innerHTML = '<div class="table-wrap"><table>' +
    ColumnRenderer.renderHead(columns) + ColumnRenderer.renderBody(data, columns, rowKey) + "</table></div>";
};
_ModalTableAdapter.prototype.showLoading = function () {
  if (this.containerEl) this.containerEl.innerHTML = '<div class="loading-state"><div class="spinner"></div>加载中…</div>';
};
_ModalTableAdapter.prototype.showEmpty = function (text) {
  if (this.containerEl) this.containerEl.innerHTML = '<div class="empty-state">' + esc(text || "暂无数据") + "</div>";
};
_ModalTableAdapter.prototype.destroy = function () { this.containerEl = null; };

/* ---- CommonTable 公共组件 ---- */
function CommonTable(options) {
  if (!options) { console.warn("[CommonTable] options required"); return null; }
  if (!options.columns || !options.columns.length) { console.warn("[CommonTable] columns required"); return null; }
  if (!options.rowKey) { console.warn("[CommonTable] rowKey required"); return null; }
  if (!options.el && !options.tableId) { console.warn("[CommonTable] el or tableId required"); return null; }

  this.options = options;
  this.tableId = options.tableId || "";
  this.columns = options.columns;
  this.rowKey = options.rowKey;
  this.stickyHeader = options.stickyHeader !== false;
  this.scrollX = options.scrollX !== false;
  this.scrollY = options.scrollY !== false;
  this.emptyText = options.emptyText || "暂无数据";
  this.dataSource = options.dataSource || {};
  this.events = options.events || {};
  this.rowSelection = options.rowSelection || false;
  this.summary = options.summary || null;
  this.rowClassName = options.rowClassName || null;
  this.toolbar = options.toolbar || null;

  this.state = {
    page: 1, pageSize: 50, total: 0, pages: 1,
    sortField: null, sortOrder: null, filters: {},
    loading: false, data: [], selectedKeys: [], selectedRows: []
  };
  this._reqId = 0;

  this._eventBus = new _EventBus();
  for (var name in this.events) { this._eventBus.on(name, this.events[name]); }
  /* 分页控件事件 → 组件内部刷新（调用方无需重复绑定翻页逻辑） */
  if (!this.events.onPageChange) {
    var self = this;
    this._eventBus.on("onPageChange", function (p) {
      self.refresh({ page: p.page, pageSize: p.pageSize });
    });
  }

  this._pager = new _PagerBinder();
  this._dataAdapter = new _TableDataAdapter();

  var AdapterCtor = options.adapter || _NativeTableAdapter;
  this.adapter = new AdapterCtor();

  this.stateEl = options.stateEl || (options.tableId ? document.getElementById(options.tableId + "-state") : null);
  this.pagerEl = options.pagerEl || (options.tableId ? document.getElementById(options.tableId + "-pager") : null);

  if (options.pagination === false) { this.pagination = false; }
  else {
    this.pagination = Object.assign({ current: 1, pageSize: 50, pageSizeOptions: [20, 50, 100], showTotal: true, showSizeChanger: true }, options.pagination || {});
    this.state.pageSize = this.pagination.pageSize;
  }

  if (options.adapter === _ModalTableAdapter || options.adapter === CommonTable.ModalAdapter) {
    if (options.el) this.adapter.bindContainer(typeof options.el === "string" ? document.querySelector(options.el) : options.el);
  } else {
    var el = options.el ? (typeof options.el === "string" ? document.querySelector(options.el) : options.el) : document.getElementById(options.tableId);
    if (el) {
      this.adapter.bindTable(el);
      /* staticHead=true 时保留页面静态表头（渐进增强/首屏可读），仅由组件渲染表体 */
      if (options.staticHead !== true) this.adapter.renderHead(this.columns);
      if (this.adapter.bindRowClick) this.adapter.bindRowClick(this._eventBus);
    }
  }

  return this;
}

CommonTable.prototype.refresh = function (params) {
  var self = this;
  if (params) {
    if (params.page != null) self.state.page = params.page;
    if (params.pageSize != null) self.state.pageSize = params.pageSize;
    if (params.sortField != null) self.state.sortField = params.sortField;
    if (params.sortOrder != null) self.state.sortOrder = params.sortOrder;
    if (params.filters != null) self.state.filters = params.filters;
    if (params.resetPage) self.state.page = 1;
  }
  var reqId = ++self._reqId;
  self.state.loading = true;
  if (self.adapter.showLoading) {
    if (self.adapter instanceof _NativeTableAdapter) self.adapter.showLoading(self.stateEl);
    else self.adapter.showLoading();
  }
  var reqParams = self._dataAdapter.buildRequest({
    page: self.state.page, pageSize: self.state.pageSize,
    sortField: self.state.sortField, sortOrder: self.state.sortOrder,
    filters: self.state.filters
  }, self.dataSource);
  var qs = new URLSearchParams();
  for (var k in reqParams) { if (reqParams[k] != null && reqParams[k] !== "") qs.set(k, reqParams[k]); }
  var url = self.dataSource.url + (self.dataSource.url.indexOf("?") >= 0 ? "&" : "?") + qs.toString();
  return API.get(url).then(function (rawData) {
    if (self._reqId !== reqId) return null;
    var resp = self._dataAdapter.parseResponse(rawData, self.dataSource.itemsKey || "items");
    self.state.data = resp.items;
    self.state.total = resp.total;
    self.state.page = resp.page;
    self.state.pageSize = resp.pageSize;
    self.state.pages = resp.pages;
    self.state.loading = false;
    if (self.adapter instanceof _NativeTableAdapter) {
      if (!resp.items.length) { self.adapter.showEmpty(self.stateEl, self.emptyText); }
      else { if (self.stateEl) self.stateEl.innerHTML = ""; self.adapter.renderBody(resp.items, self.columns, self.rowKey); }
      if (self.pagination !== false) {
        self._pager.render(self.pagerEl, { current: resp.page, pageSize: resp.pageSize, total: resp.total, pages: resp.pages }, self._eventBus);
      }
    } else {
      self.adapter.renderTable(resp.items, self.columns, self.rowKey);
    }
    self._eventBus.emit("onRefresh", { data: resp.items, total: resp.total, statistics: resp.statistics });
    return resp;
  }).catch(function (err) {
    if (self._reqId !== reqId) return null;
    self.state.loading = false;
    if (self.adapter instanceof _NativeTableAdapter) self.adapter.showError(self.stateEl, (err && err.message) || "加载失败");
    else if (self.adapter.showEmpty) self.adapter.showEmpty("加载失败");
    self._eventBus.emit("onError", { error: err });
    if (typeof toast === "function") toast("数据加载失败", "error");
    console.error("[CommonTable] refresh error:", err);
    return null;
  });
};

CommonTable.prototype.setData = function (data) {
  if (!Array.isArray(data)) { console.warn("[CommonTable] setData: data is not array"); data = []; }
  this.state.data = data;
  if (this.adapter instanceof _NativeTableAdapter) {
    if (!data.length) this.adapter.showEmpty(this.stateEl, this.emptyText);
    else { if (this.stateEl) this.stateEl.innerHTML = ""; this.adapter.renderBody(data, this.columns, this.rowKey); }
  } else {
    this.adapter.renderTable(data, this.columns, this.rowKey);
  }
};

CommonTable.prototype.getSelectedRows = function () { return this.state.selectedRows; };

CommonTable.prototype.destroy = function () {
  if (this.adapter && this.adapter.destroy) this.adapter.destroy();
  if (this._eventBus) this._eventBus.off();
  this.adapter = this._eventBus = this._pager = this._dataAdapter = null;
};

CommonTable.NativeAdapter = _NativeTableAdapter;
CommonTable.ModalAdapter = _ModalTableAdapter;
CommonTable.DataAdapter = _TableDataAdapter;
CommonTable.ColumnRenderer = ColumnRenderer;

# DataTable 统一表格组件

现代企业 ERP 列表表格：**单视口滚动**、表头 sticky、列宽/显隐/排序、后端分页，基于 `CommonTable` 数据层。

## 文件

| 文件 | 职责 |
|------|------|
| `static/data_table.js` | DataTable 组件（DOM 壳层 + 列管理） |
| `static/data_table.css` | 视口滚动、固定操作列、列工具条 |
| `static/common_table.js` | 请求、分页、列渲染、状态态 |
| `static/list_page.js` | 月份/店铺/筛选 + 导出 + `useDataTable` 开关 |

`style.css` 已 `@import data_table.css`。

## 快速接入（ListPage）

```javascript
new ListPage({
  tableId: "orders-table",
  useDataTable: true,
  sortFilterIds: { field: "f-order-by", dir: "f-order-dir" }, // 可选：与筛选区排序下拉联动
  url: "/api/orders",
  itemsKey: "orders",
  rowKey: "order_id",
  exportUrl: "/api/orders/export",
  filters: [ /* 不变 */ ],
  columns: [ /* 见下 */ ],
});
```

HTML 表格改为：

```html
<div class="table-wrap">
  <table id="orders-table"><thead></thead><tbody></tbody></table>
</div>
<div id="table-state"></div>
<div class="pager" id="pager"></div>
```

脚本顺序：`api.js` → `common_table.js` → **`data_table.js`** → `list_page.js`。

## DataTable / 列 API

### 列字段

| 字段 | 说明 |
|------|------|
| `key` | 列 ID（列宽/显隐 localStorage、表头 `data-col-key`） |
| `title` | 表头文案 |
| `type` | `orderId` / `amount` / `datetime` / `status` / `text` / 默认 `custom` |
| `render(row, index?)` | 自定义单元格（优先于 type 默认） |
| `sortable` | `true` 或后端 `order_by` 字段名；表头点击排序 |
| `fixed` | `"right"` 固定操作列 |
| `hidden` | 初始隐藏 |
| `width` / `minWidth` | 列宽 |
| `align` | `right` / `center` |
| `ellipsis` | 文本省略 + title |
| `amountTone` | 配合 `type: "amount"` 或 `DataTable.format.amount(v, tone)` |

### 统一格式化

```javascript
DataTable.format.orderId(id);           // 等宽 + 省略 + 可复制 data-copy
DataTable.format.amount(v, "success");  // amount--success|warning|danger|muted
DataTable.format.datetime(v);         // fmtTime
```

### 功能对照（需求 1–12）

| # | 能力 | 实现 |
|---|------|------|
| 1–2 | 固定表头 + 横纵同步 | 仅 `.data-table__viewport` 滚动；单 `<table>` + `thead sticky` |
| 3 | 列宽调整 | 表头 `.dt-col-resizer`，持久化 `wxr_dt_width_{storageKey}` |
| 4 | 列排序 | `sortable` + 表头点击；可联动 `sortFilterIds` |
| 5 | 列显隐 | 工具条「列显示」，`wxr_dt_cols_{storageKey}` |
| 6 | 右侧操作列 | `fixed: "right"` → `col-fixed-right` |
| 7 | 金额/日期/状态/订单号 | `type` + `DataTable.format` |
| 8 | 分页 | `CommonTable` + `renderPager` |
| 9 | 筛选/搜索 | 仍由页面 `filter-grid` + `ListPage.filters` |
| 10 | 加载/空/错 | `table-state` + `loading-state` / `empty-state` / `error-state` |
| 11 | 自适应宽度 | `min-width:0`、视口内横向滚动 |
| 12 | 后端分页 | `API.get` + `page` / `page_size` 不变 |

### 约束

- **无双纵向滚动条**：`.table-wrap.data-table-host { overflow: visible }`，只滚 viewport
- **表头不错位**：单表结构；列宽同步 `_syncBodyColWidths`
- **长订单号**：`.cell-order-id` 省略
- **窄屏不撑破页面**：viewport 内横滚
- **筛选回第一页**：`ListPage.reload()` → `resetPage: true`
- **导出一致**：`exportFile()` → `collectParams()`（不含页码）

## 迁移进度

| 页面 | 状态 |
|------|------|
| 普通订单 `orders.html` | ✅ DataTable |
| 退款订单 `refund_orders.html` | ✅ DataTable |
| 供应商成本 `supplier_costs.html` | ✅ DataTable |
| 异常处理 `anomalies.html` | ✅ DataTable |
| 操作日志 `logs.html` | ✅ DataTable |
| 对账成本 `reconciliation.html` | 待迁移（含弹窗内嵌表） |
| 利润核算 `profit.html` | 待迁移 |
| 每月盈利 `monthly_profit.html` | 待迁移 |
| 数据导入 `index.html` | 待迁移（批次/预览表） |
| 用户/店铺 `users.html` / `stores.html` | 待迁移（自定义加载逻辑） |

## 低级 API（非 ListPage）

```javascript
new DataTable({
  tableId: "my-table",
  columns, rowKey, dataSource: { url, itemsKey, buildParams },
  pagination: { pageSize: 50 },
  storageKey: "my-table",
});
```

底层仍为 `new CommonTable({ ... })`（弹窗内表格可用 `CommonTable.ModalAdapter`）。

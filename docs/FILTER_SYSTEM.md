# 全站统一筛选系统

## 组件

| 文件 | 职责 |
|------|------|
| `static/month_filter.js` | 月份/日期区间 UI、URL 同步、默认当月 |
| `static/store_scope.js` | 店铺/成员筛选、写操作 `writeTarget()` |
| `static/filter_system.js` | **合并查询参数** `FilterSystem.buildQueryParams` / `toSearchParams` |
| `static/list_page.js` | 列表页：`collectParams()` → 表格/统计/导出 |
| `app/time_range.py` | 后端 **左闭右开** `[start, end)` 解析 |

## 前端交互（统一）

1. **默认当前月**（无 URL 参数时）
2. 年/月选择、上一月/下一月
3. 快捷：本月、上月、近 3/6 月、今年、全部
4. 自定义 `from` / `to`（`to` 为含当日，后端转次日 00:00 排他）
5. **店铺 + 成员**（管理员）/ 本店只读（成员）
6. 业务页 **关键词、状态、排序** 等与时间参数一并提交
7. 切换筛选 → **重新请求后端**；`ListPage.reload()` 带 `resetPage: true`
8. 统计卡片、表格、图表、导出共用 `MF + SS + filters`
9. URL 参数：`month` / `from` & `to` / `t=all`，与 `member`、`store_id` 可组合

## 查询参数优先级（后端）

`from` + `to` > `month` > 旧 `start`/`end`  
区间：**左闭右开** — `month=2026-05` → `[2026-05-01 00:00:00, 2026-06-01 00:00:00)`

## 各页面时间字段映射

| 页面路由 | 主要 API | UI 时间口径文案 | 后端过滤字段 | 可切换口径 |
|----------|----------|-----------------|--------------|------------|
| `/dashboard` 经营总览 | `GET /api/summary` | 订单下单时间 | 订单 `order_time` 汇总区间 | 否 |
| `/data-import` 数据导入 | `GET /api/import/batches` | 导入时间 | `imported_at` | 否 |
| `/monthly-profit` 每月盈利 | `GET /api/monthly-profit` | 订单下单时间 | 月维度与 `profits.monthly_operating(start,end)` 对齐 | 否 |
| `/reconciliation` 对账成本 | `GET /api/reconciliation` | 订单下单时间 | `orders.order_time` | 否 |
| `/profit` 利润核算 | `GET /api/profit` | 订单下单时间 | `orders.order_time` | 否 |
| `/orders` 普通订单 | `GET /api/orders` | 订单下单时间 | `order_time` | `order_by` 排序字段 |
| `/refund-orders` 退款订单 | `GET /api/refund-orders` | 退款申请时间（缺失回退下单） | 默认 `apply_time` | `time_field`: `apply_time` / `success_time` / `order_time` |
| 退款明细（详情） | `GET /api/orders/{id}` 内明细 | 同退款订单页说明 | 逐笔 `refund_details` | 在订单详情抽屉中展示 |
| `/supplier-costs` 供应商成本 | `GET /api/supplier-costs` | 供应商订单创建时间 | 默认 `created_at` | `time_field`: `created_at` / `ship_time` |
| `/anomalies` 异常处理 | `GET /api/anomalies` | 异常发现时间 | `created_at` | 否 |
| `/logs` 操作日志 | `GET /api/logs` | 操作发生时间 | `time` | 否 |

### 退款跨月（需求 9）

- **列表归属**：默认按 **退款申请时间**；缺失时回退 **下单时间**（`order_time_fallback`）
- **统计独立口径**：`success_refund_amount_by_success_time` 按 **退款成功时间** 统计，与申请时间月份筛选并列展示
- 切换 `time_field=success_time` 或 `order_time` 时，列表与统计整体按所选字段过滤（后端 `order_queries`）

## 代码用法

```javascript
// 列表页（ListPage 已内置）
FilterSystem.buildQueryParams(mf, ss, { keyword: "..." });

// 自定义页
const qs = FilterSystem.toSearchParams(MF, SS, { page: 1, page_size: 50 });
await API.get(`/api/reconciliation?${qs}`);

// 导出（与筛选一致）
const exportQs = new URLSearchParams(ListPage.collectParams()).toString();
```

## 页面布局（统一）

```html
<div class="card">
  <div class="list-filters">
    <div id="month-filter"></div>
    <div id="store-scope"></div>
  </div>
</div>
```

下方为统计卡片 + 业务 `filter-grid` + 表格。

## 禁止事项

- **禁止**仅在前端对已加载页数据做日期过滤
- **禁止**各页自创月份 UI；必须使用 `MonthFilter`

## 测试

```bash
python -m pytest tests/test_month_filter.py -q
```

覆盖：月末边界、跨月退款、空月、统计与表格一致、翻页保留筛选、导出、参数优先级、各主列表 API。

/* 全站统一筛选：MonthFilter + StoreScope + 业务筛选参数合并
 * 与 app/time_range.py 左闭右开规则对齐；所有列表/统计/导出应使用 buildQueryParams。
 */
(function (global) {
  "use strict";

  /** 各页面时间口径（展示文案 + 后端主时间字段） */
  var PAGE_TIME_FIELDS = {
    "/dashboard": {
      label: "订单下单时间",
      api: "/api/summary",
      backend: "order_time（利润汇总按订单下单时间归属月份）",
    },
    "/data-import": {
      label: "导入批次时间",
      api: "/api/import/batches",
      backend: "imported_at（批次入库时间）",
    },
    "/monthly-profit": {
      label: "自然月归属",
      api: "/api/monthly-profit",
      backend: "profit_statements.month + 费用月份对齐（后端按月聚合）",
    },
    "/reconciliation": {
      label: "订单下单时间",
      api: "/api/reconciliation",
      backend: "orders.order_time",
    },
    "/profit": {
      label: "订单下单时间",
      api: "/api/profit",
      backend: "orders.order_time",
    },
    "/orders": {
      label: "订单下单时间",
      api: "/api/orders",
      backend: "orders.order_time",
    },
    "/refund-orders": {
      label: "退款申请时间（缺失回退下单时间）",
      api: "/api/refund-orders",
      backend: "默认 apply_time；time_field=success_time|order_time 可切换",
      extra: "success_refund_amount_by_success_time 按退款成功时间独立统计",
    },
    "refund-details": {
      label: "同退款订单",
      api: "（订单详情内 /api/orders/{id} 退款明细）",
      backend: "refund_details.refund_time / 订单上下文",
    },
    "/supplier-costs": {
      label: "供应商订单创建时间",
      api: "/api/supplier-costs",
      backend: "默认 created_at；time_field=ship_time 可切换",
    },
    "/anomalies": {
      label: "异常发现时间（记录生成时间）",
      api: "/api/anomalies",
      backend: "anomalies.created_at",
    },
    "/logs": {
      label: "操作发生时间",
      api: "/api/logs",
      backend: "operation_log.time",
    },
  };

  function buildQueryParams(mf, ss, extra) {
    var p = {};
    if (mf && typeof mf.getParams === "function") {
      Object.assign(p, mf.getParams());
    }
    if (ss && typeof ss.getParams === "function") {
      Object.assign(p, ss.getParams());
    } else if (ss && typeof ss.appendTo === "function") {
      ss.appendTo(p);
    }
    if (extra && typeof extra === "object") {
      Object.keys(extra).forEach(function (k) {
        var v = extra[k];
        if (v !== null && v !== undefined && v !== "") p[k] = v;
      });
    }
    return p;
  }

  function toSearchParams(mf, ss, extra) {
    var p = new URLSearchParams();
    var obj = buildQueryParams(mf, ss, extra);
    Object.keys(obj).forEach(function (k) {
      if (obj[k] != null && obj[k] !== "") p.set(k, obj[k]);
    });
    return p;
  }

  function metaForPath(pathname) {
    pathname = pathname || (typeof location !== "undefined" ? location.pathname : "");
    return PAGE_TIME_FIELDS[pathname] || null;
  }

  function suggestTimeLabel(pathname) {
    var m = metaForPath(pathname);
    return m ? m.label : "";
  }

  global.FilterSystem = {
    PAGE_TIME_FIELDS: PAGE_TIME_FIELDS,
    buildQueryParams: buildQueryParams,
    toSearchParams: toSearchParams,
    metaForPath: metaForPath,
    suggestTimeLabel: suggestTimeLabel,
  };
})(typeof window !== "undefined" ? window : this);

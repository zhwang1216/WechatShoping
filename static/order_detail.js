/* 订单详情共享视图（普通订单 / 退款订单共用）
   依赖：api.js（格式化/徽章）、modal.js（openDrawer 右侧抽屉）
   两页只保留列表列定义，详情结构、数据请求、错误/加载态集中在此处。 */

let _orderDetailSeq = 0;

function _odKv(label, valueHtml) {
  return '<div class="kv"><div class="k">' + esc(label) + '</div><div class="v">' +
    valueHtml + "</div></div>";
}

function _odGrid(pairs) {
  return '<div class="kv-grid">' + pairs.map(function (p) {
    return _odKv(p[0], p[1]);
  }).join("") + "</div>";
}

function _odSection(title, inner) {
  return '<div class="detail-section"><h4>' + title + "</h4>" + inner + "</div>";
}

function _odRefundSummary(v) {
  if (!v) return '<div style="color:var(--text-light)">无退款汇总</div>';
  var rf = v.refund || {}, bs = v.refund_by_status || {};
  var applySource = rf.refund_apply_time_source === "order_time_fallback"
    ? "回退自下单时间（申请时间缺失）" : "平台退款记录";
  return _odGrid([
    ["退款状态", refundStatusBadge(rf.refund_status)],
    ["退款总额", fmtMoneyYuan(rf.refund_amount)],
    ["退款申请时间", rf.refund_apply_time ? fmtTime(rf.refund_apply_time) : "—"],
    ["退款成功时间", rf.refund_success_time ? fmtTime(rf.refund_success_time) : "—"],
    ["申请时间来源", esc(applySource)],
    ["退款成功金额", fmtMoneyYuan(bs.success)],
    ["退款处理中金额", fmtMoneyYuan(bs.pending)],
    ["退款关闭金额", fmtMoneyYuan(bs.closed)],
    ["退款笔数", rf.refund_count],
    ["退款粒度", rf.refund_amount_granularity === "unknown"
      ? "待核对（订单级/商品行级无法判断，不参与自动利润终算）" : "商品行级"],
    ["状态来源", rf.status_source === "derived" ? "推导" : "平台原始"],
  ]);
}

function _odOrderSummary(d) {
  var a = d.amounts || {};
  return _odGrid([
    ["订单号", esc(d.order_id)],
    ["下单时间", fmtTime(d.order_time)],
    ["订单状态", orderStatusBadge(d.order_status)],
    ["原始状态", esc(d.source_status || "—")],
    ["实际支付金额", fmtMoneyYuan(a.actual_paid_amount)],
    ["实际收款金额", fmtMoneyYuan(a.actual_received_amount)],
    ["技术服务费", fmtMoneyYuan(a.technical_service_fee)],
    ["商品总价", fmtMoneyYuan(a.product_total)],
    ["商品行数", d.product_line_count],
    ["是否有退款", d.has_refund ? "是（见退款明细）" : "否"],
    ["订单级金额冲突", d.amount_conflict
      ? "是（多行不同值，已取唯一值并生成异常）" : "否"],
  ]);
}

function _odRefundRows(refunds) {
  if (!refunds || !refunds.length) {
    return '<div style="color:var(--text-light)">该订单无退款记录</div>';
  }
  return '<div class="table-wrap"><table><thead><tr>' +
    '<th>#</th><th class="num">退款金额</th><th>退款状态</th><th>状态来源</th>' +
    '<th>退款时间</th><th>原售后文本</th><th>源工作表/行号</th><th>导入批次</th>' +
    '</tr></thead><tbody>' +
    refunds.map(function (r, i) {
      return "<tr><td>" + (i + 1) + '</td><td class="num">' +
        fmtMoneyYuan(r.refund_amount) + "</td><td>" +
        refundStatusBadge(r.refund_status) + "</td><td>" +
        statusSourceBadge(r.status_source) + "</td><td>" +
        fmtTime(r.refund_time) + "</td><td>" +
        esc(r.raw_after_sale_text || "—") + "</td><td>" +
        esc(r.sheet_name || "—") + " 第 " + r.row_number + " 行</td>" +
        '<td style="font-size:12px">' + esc(r.batch_id || "—") + "</td></tr>";
    }).join("") + "</tbody></table></div>";
}

function _odLineRows(d) {
  var lines = d.lines || [];
  if (!lines.length) {
    return '<div style="color:var(--text-light)">无商品行明细</div>';
  }
  return '<div class="table-wrap"><table><thead><tr>' +
    '<th>商品标题</th><th>规格</th><th>SKU</th><th class="num">数量</th>' +
    '<th class="num">实际总价</th><th class="num">退款金额</th>' +
    '<th>退款状态</th><th>源表行号</th></tr></thead><tbody>' +
    lines.map(function (l) {
      var p = l.product || {}, rf = l.refund || {}, am = l.amounts || {};
      return "<tr><td>" + esc(p.title || "—") + "</td><td>" +
        esc(p.sku_attribute || "—") + "</td><td>" +
        esc(p.sku_code || "—") + '</td><td class="num">' + p.quantity +
        '</td><td class="num">' + fmtMoney(p.actual_total_price) +
        '</td><td class="num">' + fmtMoney(am.refund_amount) + "</td><td>" +
        refundStatusBadge(rf.refund_status) + " " +
        statusSourceBadge(rf.status_source) + "</td><td>" +
        (l.source_record_id ? "可追溯" : "—") + "</td></tr>";
    }).join("") + "</tbody></table></div>";
}

function _odTrace(d) {
  var batches = d.batches || [];
  if (!batches.length) {
    return '<div style="color:var(--text-light)">无批次信息</div>';
  }
  return batches.map(function (b) {
    return '<div class="kv-grid" style="margin-bottom:8px">' +
      _odKv("导入批次", esc(b.batch_id)) +
      _odKv("源文件", esc(b.file_name)) +
      _odKv("导入时间", esc(b.imported_at)) + "</div>";
  }).join("");
}

function renderOrderDetailBody(d, refunds, mode) {
  var sections = [];
  if (mode === "refund") {
    sections.push(_odSection("退款汇总（按订单去重）", _odRefundSummary(d.refund_view)));
    sections.push(_odSection("订单信息", _odOrderSummary(d)));
    sections.push(_odSection("退款记录（逐笔，" + refunds.length + " 笔）",
      _odRefundRows(refunds)));
    sections.push(_odSection("商品行明细（" + (d.lines || []).length + " 行）",
      _odLineRows(d)));
  } else {
    sections.push(_odSection("订单汇总", _odOrderSummary(d)));
    sections.push(_odSection("商品行明细（" + (d.lines || []).length + " 行）",
      _odLineRows(d)));
    sections.push(_odSection("退款记录（逐笔，" + refunds.length + " 笔）",
      _odRefundRows(refunds) +
      (d.refund_view
        ? '<div style="margin-top:8px;color:var(--text-light);font-size:13px">退款汇总（按订单去重）：' +
          "总额 " + fmtMoneyYuan(d.refund_view.refund.refund_amount) +
          "，成功 " + fmtMoneyYuan(d.refund_view.refund_by_status.success) +
          "，处理中 " + fmtMoneyYuan(d.refund_view.refund_by_status.pending) +
          "，关闭 " + fmtMoneyYuan(d.refund_view.refund_by_status.closed) +
          "，粒度 " + (d.refund_view.refund.refund_amount_granularity === "unknown"
            ? "待核对" : "商品行级") + "</div>"
        : "")));
  }
  sections.push(_odSection("数据溯源", _odTrace(d)));
  return sections.join("");
}

/* 打开订单详情抽屉：mode="order"（默认）|"refund" */
function openOrderDetail(orderId, options) {
  var mode = (options && options.mode) || "order";
  var seq = ++_orderDetailSeq;
  var title = (mode === "refund" ? "退款订单详情：" : "订单详情：") + orderId;
  openDrawer({
    title: title,
    width: mode === "refund" ? 860 : 800,
    bodyHtml: '<div class="loading-state"><div class="spinner"></div>加载中…</div>',
  });
  var overlay = document.getElementById("um-modal");
  var body = overlay ? overlay.querySelector(".um-body") : null;
  if (!body) return;
  var tasks = [API.get("/api/orders/" + encodeURIComponent(orderId))];
  if (mode === "refund") {
    tasks.push(API.get("/api/orders/" + encodeURIComponent(orderId) + "/refunds"));
  }
  Promise.all(tasks).then(function (res) {
    if (seq !== _orderDetailSeq || !document.body.contains(overlay)) return;
    var d = res[0];
    var refunds = mode === "refund"
      ? ((res[1] && res[1].refunds) || [])
      : (d.refunds || []);
    body.innerHTML = renderOrderDetailBody(d, refunds, mode);
  }).catch(function (err) {
    if (seq !== _orderDetailSeq || !document.body.contains(overlay)) return;
    body.innerHTML = '<div class="error-state"><div class="msg">' +
      esc((err && err.message) || "加载失败") + "</div></div>";
  });
}

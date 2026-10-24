"""退款订单视图构建（导入落盘与查询实时构建共用的单一逻辑，spec 6.6）

退款判定不依赖订单状态，仅依据订单行 refund_status != none：
部分退款、全额退款、多次退款、处理中、退款关闭但历史存在退款记录均纳入。

月份筛选时间字段（spec 月份筛选统一规则）：
- refund_apply_time  退款申请时间 = 该订单逐笔退款明细中最早非空 refund_time；
  全部缺失时回退订单下单时间（source=order_time_fallback），保证跨月退款
  不会因退款时间缺失而被所有月份排除；
- refund_success_time 退款成功时间 = 成功退款（partial/full）明细中最晚非空
  refund_time，缺失即 None（不参与时间筛选）。
"""
from decimal import Decimal

from .normalizer import money_str

STATUS_PRIORITY = ("full", "unknown", "pending", "partial", "closed")
BUCKET_KEYS = {"partial": "success", "full": "success", "pending": "pending",
               "closed": "closed"}
SUCCESS_STATUSES = ("partial", "full")


def _dec(s):
    if s in (None, ""):
        return Decimal("0")
    return Decimal(s)


def order_level_amount(lines, getter):
    """订单级金额取值：多行取唯一值；存在冲突取首个非空并标记（RC-004）"""
    values = []
    for l in lines:
        v = getter(l)
        if v is not None and v not in values:
            values.append(v)
    if not values:
        return None, False
    return values[0], len(values) > 1


def order_refund_times(order_id, lines, refund_details):
    """计算订单级退款申请/成功时间（供月份筛选与统计口径展示）"""
    details = [d for d in (refund_details or [])
               if d.get("order_id") == order_id]
    apply_times = [d["refund_time"] for d in details if d.get("refund_time")]
    success_times = [d["refund_time"] for d in details
                     if d.get("refund_time")
                     and d.get("refund_status") in SUCCESS_STATUSES]
    order_time = next((l.get("order_time") for l in lines
                       if l.get("order_time")), None)
    if apply_times:
        apply_time, source = min(apply_times), "refund_time"
    elif order_time:
        apply_time, source = order_time, "order_time_fallback"
    else:
        apply_time, source = None, None
    success_time = max(success_times) if success_times else None
    return apply_time, source, success_time


def build_refund_views(order_records, refund_details=None):
    """按订单去重构建退款订单汇总视图列表（同一原订单仅一条）"""
    by_order = {}
    for o in order_records:
        if o.get("order_id"):
            by_order.setdefault(o["order_id"], []).append(o)

    views = []
    for order_id, lines in by_order.items():
        refund_lines = [l for l in lines
                        if (l.get("refund") or {}).get("refund_status") not in (None, "none")]
        if not refund_lines:
            continue

        granularities = {(l["refund"].get("granularity") or "line") for l in refund_lines}
        granularity = "unknown" if "unknown" in granularities else "line"
        buckets = {"success": Decimal("0"), "pending": Decimal("0"), "closed": Decimal("0")}

        if granularity == "unknown":
            seen, uniq = set(), []
            for l in refund_lines:
                amt = l["refund"].get("refund_amount")
                if amt not in seen:
                    seen.add(amt)
                    uniq.append(l["refund"])
            refund_amount = sum((_dec(r.get("refund_amount")) for r in uniq), Decimal("0"))
            for r in uniq:
                bk = BUCKET_KEYS.get(r.get("refund_status"))
                if bk:
                    buckets[bk] += _dec(r.get("refund_amount"))
        else:
            refund_amount = sum((_dec(l["refund"].get("refund_amount"))
                                 for l in refund_lines), Decimal("0"))
            for l in refund_lines:
                bk = BUCKET_KEYS.get(l["refund"].get("refund_status"))
                if bk:
                    buckets[bk] += _dec(l["refund"].get("refund_amount"))
        refund_count = len(refund_lines)

        statuses = {l["refund"]["refund_status"] for l in refund_lines}
        status = next((s for s in STATUS_PRIORITY if s in statuses), "unknown")
        status_source = ("derived" if any(l["refund"].get("status_source") == "derived"
                                          for l in refund_lines) else "platform")

        received, recv_conflict = order_level_amount(
            lines, lambda l: (l.get("amounts") or {}).get("actual_received_amount"))
        tech, _ = order_level_amount(
            lines, lambda l: (l.get("amounts") or {}).get("technical_service_fee"))
        paid, _ = order_level_amount(
            lines, lambda l: (l.get("amounts") or {}).get("actual_paid_amount"))
        product_total = sum((_dec((l.get("product") or {}).get("actual_total_price"))
                             for l in lines), Decimal("0"))
        apply_time, apply_source, success_time = order_refund_times(
            order_id, lines, refund_details)

        merged = {}
        qty_total = 0
        for l in lines:
            p = l.get("product") or {}
            name = p.get("title") or p.get("sku_code") or "未命名商品"
            attr = p.get("sku_attribute") or ""
            merged[(name, attr)] = merged.get((name, attr), 0) + (p.get("quantity") or 0)
            qty_total += p.get("quantity") or 0
        product_summary = "; ".join(
            f"{name}{'(' + attr + ')' if attr else ''}×{qty}"
            for (name, attr), qty in merged.items())

        views.append({
            "record_id": "refund_view_" + order_id,
            "record_type": "refund_order_view",
            "order_id": order_id,
            "order_time": next((l.get("order_time") for l in lines if l.get("order_time")), None),
            "order_status": next((l.get("order_status") for l in lines
                                  if l.get("order_status") not in (None, "unknown")), "unknown"),
            "source_status": next((l.get("source_status") for l in lines
                                   if l.get("source_status")), None),
            "product_summary": product_summary,
            "product_line_count": len(lines),
            "product_quantity_total": qty_total,
            "has_refund": True,
            "refund": {
                "refund_status": status,
                "refund_amount": money_str(refund_amount),
                "refund_amount_granularity": granularity,
                "status_source": status_source,
                "refund_deduction_status": "unknown",
                "refund_count": refund_count,
                "last_refund_time": None,
                "refund_apply_time": apply_time,
                "refund_apply_time_source": apply_source,
                "refund_success_time": success_time,
                "refund_reason": None,
            },
            "refund_by_status": {k: money_str(v) for k, v in buckets.items()},
            "amounts": {
                "actual_paid_amount": paid,
                "actual_received_amount": received,
                "technical_service_fee": tech,
                "product_total": money_str(product_total),
                "matched_supplier_total_cost": None,
            },
            "amount_conflict": recv_conflict,
            "pending_review": granularity == "unknown" or status == "unknown",
            "reconciliation": {
                "match_status": "unmatched",
                "profit_status": "pending_confirmation",
                "exception_codes": [],
            },
            "source": {
                "batch_ids": sorted({l.get("batch_id") for l in lines if l.get("batch_id")}),
                "raw_record_ids": [l.get("source_record_id") for l in lines
                                   if l.get("source_record_id")],
            },
            "updated_at": max((l.get("updated_at") or "" for l in lines), default=None),
        })
    return views

"""供应商成本业务有效性：按对应微信订单发货状态判定是否参与统计与核算。

validity_status:
  - valid: 可参与有效业务统计、对账匹配与利润核算
  - invalid: 对应订单未发货等，保留原始记录但不参与统计
  - pending: 无法匹配订单或发货状态待核对，暂不参与统计
"""
from .data_schemas import now_iso
from .normalizer import INVISIBLE_CHARS, ORDER_STATUS_MAP, normalize_order_status
from .profit_service import (
    PENDING_REVIEW_STATUSES,
    is_calculable_status,
    is_unshipped_order,
)

INVALID_REASON_UNSHIPPED = "订单发货状态为未发货"
PENDING_NO_ORDER = "未匹配到微信小店订单，待核对"
PENDING_SHIP_UNKNOWN = "订单发货状态无法确认，待核对"

VALIDITY_LABELS = {
    "valid": "有效",
    "invalid": "无效",
    "pending": "待核对",
}


def _norm_key(s):
    if not s:
        return None
    return str(s).translate(INVISIBLE_CHARS).strip()


def build_orders_lookup(orders_records):
    """order_id -> [wechat order lines]"""
    by_id = {}
    for o in orders_records or []:
        oid = o.get("order_id")
        if oid:
            by_id.setdefault(oid, []).append(o)
    return by_id


def cost_linked_order_keys(cost):
    keys = []
    for k in ("main_order_id", "supplier_order_id"):
        v = cost.get(k)
        if v:
            keys.append(v)
    return keys


def resolve_order_lines(cost, orders_by_id):
    """匹配微信订单行：(canonical_order_id, lines) 或 (None, None)"""
    for raw in cost_linked_order_keys(cost):
        if raw in orders_by_id:
            return raw, orders_by_id[raw]
        nk = _norm_key(raw)
        if not nk:
            continue
        for oid, lines in orders_by_id.items():
            if _norm_key(oid) == nk:
                return oid, lines
    return None, None


_CANONICAL_ORDER = set(ORDER_STATUS_MAP.values())


def _canonical_order_status(raw):
    if not raw:
        return None
    if raw in _CANONICAL_ORDER:
        return raw
    std, known = normalize_order_status(raw)
    return std if known else raw


def order_ship_snapshot(lines):
    raw_status = next(
        (l.get("order_status") for l in lines if l.get("order_status")), None)
    first_status = _canonical_order_status(raw_status)
    first_ship = next(
        (l.get("ship_status") for l in lines if l.get("ship_status") is not None),
        None,
    )
    return first_status, first_ship


def evaluate_supplier_cost_validity(
        cost, orders_by_id, batch_id=None, evaluated_at=None,
        matched_order_id=None):
    """根据当前订单库状态评估一条供应商成本记录的业务有效性。

    matched_order_id: 对账/人工匹配已确认的微信订单号（供货主单号与微信单号不一致时使用）。
    """
    evaluated_at = evaluated_at or now_iso()
    order_id, lines = resolve_order_lines(cost, orders_by_id)
    if not lines and matched_order_id and matched_order_id in orders_by_id:
        order_id, lines = matched_order_id, orders_by_id[matched_order_id]
    bid = batch_id or cost.get("batch_id")

    if not lines:
        return {
            "validity_status": "pending",
            "invalid_reason": PENDING_NO_ORDER,
            "shipment_status": None,
            "order_id": order_id,
            "evaluated_at": evaluated_at,
            "import_batch_id": bid,
        }

    order_status, ship_status = order_ship_snapshot(lines)
    shipment_status = ship_status if ship_status is not None else order_status

    if is_unshipped_order(order_status, ship_status):
        return {
            "validity_status": "invalid",
            "invalid_reason": INVALID_REASON_UNSHIPPED,
            "shipment_status": shipment_status,
            "order_id": order_id,
            "evaluated_at": evaluated_at,
            "import_batch_id": bid,
        }

    if not is_calculable_status(order_status, ship_status):
        reason = PENDING_SHIP_UNKNOWN
        if order_status is None:
            reason = PENDING_SHIP_UNKNOWN
        elif order_status in PENDING_REVIEW_STATUSES:
            reason = PENDING_SHIP_UNKNOWN
        return {
            "validity_status": "pending",
            "invalid_reason": reason,
            "shipment_status": shipment_status,
            "order_id": order_id,
            "evaluated_at": evaluated_at,
            "import_batch_id": bid,
        }

    return {
        "validity_status": "valid",
        "invalid_reason": None,
        "shipment_status": shipment_status,
        "order_id": order_id,
        "evaluated_at": evaluated_at,
        "import_batch_id": bid,
    }


def is_effective_supplier_cost(cost, orders_by_id):
    """是否参与有效业务统计（统一过滤入口）。"""
    bv = cost.get("business_validity")
    if not bv or not bv.get("validity_status"):
        bv = evaluate_supplier_cost_validity(cost, orders_by_id)
    return bv.get("validity_status") == "valid"


def apply_validity_to_cost(
        cost, orders_by_id, batch_id=None, evaluated_at=None,
        matched_order_id=None):
    """写入 business_validity 并返回 (new_validity, changed, prev_validity)。"""
    prev = cost.get("business_validity") or {}
    new = evaluate_supplier_cost_validity(
        cost, orders_by_id, batch_id=batch_id, evaluated_at=evaluated_at,
        matched_order_id=matched_order_id)
    cost["business_validity"] = new
    changed = (prev.get("validity_status") != new.get("validity_status")
               or prev.get("invalid_reason") != new.get("invalid_reason"))
    return new, changed, prev


def revalidate_supplier_cost_records(
        repo, record_ids, op_log=None, operator="system", batch_id=None,
        matched_order_hints=None):
    """按供应商成本 record_id 重判有效性（人工匹配等场景）。"""
    if not record_ids:
        return {"updated": 0}
    record_ids = set(record_ids)
    matched_order_hints = matched_order_hints or {}
    orders_by_id = build_orders_lookup(repo.read("orders")["records"])
    costs = repo.read("supplier_costs")["records"]
    now = now_iso()
    updated = 0
    for c in costs:
        if c.get("record_id") not in record_ids:
            continue
        hint = matched_order_hints.get(c.get("record_id"))
        new, changed, prev = apply_validity_to_cost(
            c, orders_by_id, batch_id=batch_id, evaluated_at=now,
            matched_order_id=hint)
        if changed and op_log:
            op_log.append("supplier_cost.validity_revised", detail={
                "supplier_record_id": c.get("record_id"),
                "order_id": new.get("order_id"),
                "before": prev,
                "after": new,
            }, operator=operator)
        if changed:
            updated += 1
    if updated:
        def mutator(d):
            d["records"] = costs
            d["updated_at"] = now
        repo.update("supplier_costs", mutator)
    return {"updated": updated}


def revalidate_supplier_costs_for_order_ids(
        repo, order_ids, op_log=None, operator="system", batch_id=None):
    """订单状态变更后，重判关联供应商成本有效性并落盘。"""
    if not order_ids:
        return {"updated": 0}
    order_ids = set(order_ids)
    orders_by_id = build_orders_lookup(repo.read("orders")["records"])
    costs = repo.read("supplier_costs")["records"]
    now = now_iso()
    updated = 0
    for c in costs:
        keys = cost_linked_order_keys(c)
        if not any(k in order_ids or _norm_key(k) in {_norm_key(x) for x in order_ids}
                   for k in keys):
            continue
        new, changed, prev = apply_validity_to_cost(
            c, orders_by_id, batch_id=batch_id, evaluated_at=now)
        if changed and op_log:
            detail = {
                "supplier_record_id": c.get("record_id"),
                "order_id": new.get("order_id"),
                "before": prev,
                "after": new,
                "batch_id": batch_id,
            }
            if new.get("validity_status") == "valid" and prev.get("validity_status") == "invalid":
                op_log.append("supplier_cost.validity_restored", detail=detail,
                              operator=operator)
            elif new.get("validity_status") == "invalid":
                op_log.append("supplier_cost.validity_invalid", detail=detail,
                              operator=operator)
            else:
                op_log.append("supplier_cost.validity_revised", detail=detail,
                              operator=operator)
        if changed:
            updated += 1

    if updated:
        def mutator(d):
            d["records"] = costs
            d["updated_at"] = now
        repo.update("supplier_costs", mutator)
    return {"updated": updated}

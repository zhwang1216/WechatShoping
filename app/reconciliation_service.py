"""订单与供应商成本对账服务（阶段4，spec 5.7 / 6.5 / 6.8）

核心规则：
1. 匹配优先级：① 微信订单号 = 供货表主订单编号 → ② 微信订单号 = 供货表订单编号 →
   ③ 标准化（去空格/不可见字符）后相等 → ④ 辅助候选（日期 + 商品编码/SKUID + 数量）。
   辅助匹配只能进入"疑似匹配/待人工确认"，禁止静默自动认定。
2. 一对多：一个订单可关联多条供应商明细（supplier_record_ids 数组），属正常匹配；
   一条成本记录被多个订单匹配时生成多对一异常。
3. 成本粒度区分：order_level=订单级（同额多行去重只扣一次）、
   package_level=物流单级（快递单号互不相同的同额行分别计入）、
   supplier_line=商品行级（逐行累加）。
4. 重复风险：总金额在多行重复出现时禁止直接重复累加，先识别再按粒度处理；
   疑似订单级重复在粒度经人工确认前不得用于最终利润计算（usable_for_profit=False）。
5. 人工确认：记录操作人、时间、匹配规则与修改前后值；重跑对账与重新导入均不覆盖。
   未匹配订单可另录货品+快递作为总成本；一旦后续供应商成本入库并匹配成功，
   以入库数据为准，丢弃这笔手工金额。
6. 对账范围：仅商品发货状态为「已发货」的订单进入对账统计（无商品发货列时按订单状态
   排除未发货/待核对，与历史数据兼容）；未发货/部分发货等不生成对账记录。
7. 异常：未匹配订单(RC-001)、未匹配成本(RC-002)、总金额重复(RC-007)、
   数量差异(RC-005)、成本缺失(RC-006)、金额差异(RC-012)、多对一、疑似匹配待确认、
   人工关系失效；异常禁止静默丢弃，自动关闭必须留痕。
"""
import io
import uuid
from datetime import datetime
from decimal import Decimal

from openpyxl import Workbook

from .data_schemas import now_iso
from .errors import (FieldNotModifiableError, InvalidFieldValueError,
                     ManualMatchConflictError, ModifyRejectedError,
                     OrderNotFoundError, ReconcileServiceError,
                     SupplierRecordNotFoundError)
from .excel_export import append_row
from .normalizer import INVISIBLE_CHARS, money_str
from .cost_validity import (
    build_orders_lookup,
    is_effective_supplier_cost,
    revalidate_supplier_cost_records,
    revalidate_supplier_costs_for_order_ids,
)
from .profit_service import is_reconciliation_scope
from .refund_view import order_level_amount
from .time_range import in_range

AMOUNT_TOLERANCE = Decimal("0.01")
PAGE_SIZES = (20, 50, 100)

MATCH_STATUS_LABELS = {
    "matched": "已匹配",
    "manual_confirmed": "人工确认",
    "suspected": "疑似匹配（待人工确认）",
    "unmatched_order": "未匹配订单",
    "unmatched_supplier": "未匹配成本",
}
MATCH_METHOD_LABELS = {
    "main_order_id_exact": "① 主订单编号精确匹配",
    "supplier_order_id_exact": "② 订单编号精确匹配",
    "normalized_equal": "③ 标准化后相等",
    "assisted_candidate": "④ 辅助候选匹配",
    "manual": "人工确认匹配",
}
GRANULARITY_LABELS = {
    "order_level": "订单级（同额去重只扣一次）",
    "package_level": "物流单级（按快递单号）",
    "supplier_line": "商品行级",
    "mixed": "混合粒度",
    "unknown": "未知",
}
CONFIRMED_GRANULARITY_LABELS = {
    "order_level": "订单级：同额多行只扣一次",
    "line_level": "行级：各明细行分别累加",
}
ANOMALY_TYPE_LABELS = {
    "UNMATCHED_ORDER": "未匹配订单（RC-001）",
    "UNMATCHED_SUPPLIER_COST": "未匹配供应商成本（RC-002）",
    "DUPLICATE_TOTAL_AMOUNT_RISK": "总金额重复风险（RC-007）",
    "QUANTITY_DIFF": "商品数量差异（RC-005）",
    "COST_MISSING": "成本缺失（RC-006）",
    "AMOUNT_DIFF": "金额差异（RC-012）",
    "MULTI_ORDER_ONE_COST": "多对一匹配异常",
    "SUSPECTED_MATCH_NEED_REVIEW": "疑似匹配待人工确认",
    "MANUAL_MATCH_BROKEN": "人工匹配关系失效",
}
RECON_ANOMALY_TYPES = set(ANOMALY_TYPE_LABELS)

MODIFIABLE_ORDER_FIELDS = {
    "order_time": {"label": "下单时间", "type": "iso_time"},
    "order_status": {"label": "订单状态", "type": "status"},
    "product.quantity": {"label": "商品数量", "type": "positive_int"},
    "product.actual_total_price": {"label": "实际总价", "type": "nonneg_amount"},
    "amounts.actual_received_amount": {"label": "实收金额", "type": "nonneg_amount"},
    "amounts.product_total": {"label": "商品总金额", "type": "nonneg_amount"},
    "amounts.shipping_fee": {"label": "运费", "type": "amount_or_null"},
}
MODIFIABLE_COST_FIELDS = {
    "cost_breakdown.total_amount": {"label": "成本总金额", "type": "nonneg_amount"},
    "cost_breakdown.product_supply_cost": {"label": "供货成本", "type": "nonneg_amount"},
    "cost_breakdown.courier_fee": {"label": "快递费", "type": "amount_or_null"},
    "supplier_actual_received": {"label": "供应商实收", "type": "nonneg_amount"},
    "product.quantity": {"label": "供货数量", "type": "positive_int"},
}
SYSTEM_FIELD_BLACKLIST = {
    "order_id", "record_id", "batch_id", "business_key", "store_id",
    "main_order_id", "supplier_order_id", "source_record_id",
    "reconciliation_id", "match_status", "match_method",
    "exception_codes", "preserved_manual",
}
_VALID_STATUSES = {
    "已发货", "已完成", "已关闭", "待发货", "待备货", "备货中",
    "已支付", "已付款", "待付款", "待支付", "退款中", "已退款",
    "部分退款", "退款失败", "",
}


def _validate_field(field_path, value):
    """校验字段是否可修改且值合法，返回 (ok, error_message)"""
    spec = MODIFIABLE_ORDER_FIELDS.get(field_path) \
        or MODIFIABLE_COST_FIELDS.get(field_path)
    if field_path in SYSTEM_FIELD_BLACKLIST:
        return False, f"字段 {field_path} 为系统关键字段，不可修改"
    if spec is None:
        return False, f"字段 {field_path} 不在可修改白名单内"
    vtype = spec["type"]
    if vtype == "positive_int":
        if not isinstance(value, int) or isinstance(value, bool) or value <= 0:
            return False, f"字段 {field_path} 要求正整数，收到 {value!r}"
    elif vtype == "nonneg_amount":
        if value is None:
            return False, f"字段 {field_path} 不可为空"
        try:
            d = Decimal(str(value))
        except Exception:
            return False, f"字段 {field_path} 要求非负金额，收到 {value!r}"
        if d < 0:
            return False, f"字段 {field_path} 要求非负金额，收到 {value!r}"
    elif vtype == "amount_or_null":
        if value is not None:
            try:
                d = Decimal(str(value))
            except Exception:
                return False, f"字段 {field_path} 要求金额或空，收到 {value!r}"
            if d < 0:
                return False, f"字段 {field_path} 要求非负金额或空，收到 {value!r}"
    elif vtype == "iso_time":
        if not isinstance(value, str) or not value.strip():
            return False, f"字段 {field_path} 要求 ISO 时间字符串"
        try:
            datetime.fromisoformat(value.replace("Z", "+00:00"))
        except Exception:
            return False, f"字段 {field_path} 不是合法 ISO 时间: {value!r}"
    elif vtype == "status":
        if not isinstance(value, str) or value not in _VALID_STATUSES:
            return False, f"字段 {field_path} 不是合法订单状态: {value!r}"
    return True, None


def _dec(s):
    return Decimal(s) if s not in (None, "") else Decimal("0")


def _norm_key(s):
    """订单键标准化：去空格与不可见字符"""
    if not s:
        return None
    return str(s).translate(INVISIBLE_CHARS).strip()


class ReconciliationService:
    def __init__(self, repo, op_log):
        self.repo = repo
        self.op_log = op_log

    # ==================== 执行对账 ====================
    def run_reconciliation(self, start=None, end=None, operator="admin"):
        orders = self.repo.read("orders")["records"]
        costs = self.repo.read("supplier_costs")["records"]
        existing = self.repo.read("reconciliation_results")["records"]

        recon_id = ("recon_" + datetime.now().strftime("%Y%m%d%H%M%S")
                    + "_" + uuid.uuid4().hex[:8])
        now = now_iso()

        manual_by_order = {}
        old_by_order = {}
        for r in existing:
            if r.get("record_type") != "order_match" or not r.get("order_id"):
                continue
            old_by_order[r["order_id"]] = r
            if r.get("match_method") == "manual":
                manual_by_order[r["order_id"]] = r

        orders_by_id = {}
        for o in orders:
            if o.get("order_id"):
                orders_by_id.setdefault(o["order_id"], []).append(o)

        orders_lookup = orders_by_id
        main_index, sub_index, normalized_index = {}, {}, {}
        seen_norm = set()
        for c in costs:
            if not is_effective_supplier_cost(c, orders_lookup):
                continue
            rid = c.get("record_id")
            if c.get("main_order_id"):
                main_index.setdefault(c["main_order_id"], []).append(c)
            if c.get("supplier_order_id"):
                sub_index.setdefault(c["supplier_order_id"], []).append(c)
            for k in (c.get("main_order_id"), c.get("supplier_order_id")):
                nk = _norm_key(k)
                if nk and (rid, nk) not in seen_norm:
                    seen_norm.add((rid, nk))
                    normalized_index.setdefault(nk, []).append(c)

        time_scope_order_ids = set()
        reconcile_scope_order_ids = set()
        for oid, wlines in orders_by_id.items():
            ot = next((l.get("order_time") for l in wlines if l.get("order_time")), None)
            if start and (ot is None or ot < start):
                continue
            if end and (ot is None or ot > end):
                continue
            time_scope_order_ids.add(oid)
            first_status = next((l.get("order_status") for l in wlines
                                 if l.get("order_status")), None)
            first_ship = next((l.get("ship_status") for l in wlines
                               if l.get("ship_status")), None)
            if is_reconciliation_scope(first_status, first_ship):
                reconcile_scope_order_ids.add(oid)

        results = []
        matched_cost_ids = set()
        fresh = []

        def add_anom(order_id, etype, severity, message, related_ids):
            fresh.append({
                "order_id": order_id,
                "exception_type": etype,
                "severity": severity,
                "message": message,
                "related_record_ids": related_ids or [],
            })

        def inherit_legacy(rec, oid):
            """继承旧记录的人工操作历史与创建时间（历史禁止丢失）"""
            old = old_by_order.get(oid)
            if not old:
                return
            rec["created_at"] = old.get("created_at") or rec["created_at"]
            old_history = (old.get("manual_review") or {}).get("history") or []
            if old_history:
                rec["manual_review"]["history"] = list(old_history)

        cost_by_id = {c.get("record_id"): c for c in costs}

        for oid in sorted(reconcile_scope_order_ids):
            wlines = orders_by_id[oid]
            wx_ids = [l.get("record_id") for l in wlines if l.get("record_id")]

            manual = manual_by_order.get(oid)
            if manual is not None:
                # 人工确认的匹配关系：重跑对账时保留，禁止被自动匹配覆盖
                sids = [s for s in (manual.get("supplier_record_ids") or [])]
                sup_records = [cost_by_id[s] for s in sids if s in cost_by_id]
                broken = [s for s in sids if s not in cost_by_id]
                rec, anoms = self._build_match_record(
                    oid, wlines, sup_records, method="manual",
                    status="manual_confirmed", confidence="high", recon_id=recon_id,
                    rule_hit="人工确认匹配（重跑对账保留原关系）",
                    confirmed_granularity=(manual.get("cost") or {}).get(
                        "confirmed_granularity"))
                rec["preserved_manual"] = True
                rec["match_detail"]["operator"] = (
                    (manual.get("manual_review") or {}).get("reviewer"))
                if broken:
                    anoms.append({
                        "order_id": oid,
                        "exception_type": "MANUAL_MATCH_BROKEN",
                        "severity": "high",
                        "message": (f"订单 {oid} 人工匹配引用的成本记录已不存在: "
                                    f"{broken}，关系已保留待人工处理"),
                        "related_record_ids": wx_ids,
                    })
                inherit_legacy(rec, oid)
                results.append(rec)
                matched_cost_ids.update(sids)
                for a in anoms:
                    add_anom(a["order_id"], a["exception_type"], a["severity"],
                             a["message"], a.get("related_record_ids"))
                rec["exception_codes"] = sorted({a["exception_type"] for a in anoms})
                continue

            sup_records, method, rule_hit = self._auto_match(
                oid, main_index, sub_index, normalized_index)
            if method:
                rec, anoms = self._build_match_record(
                    oid, wlines, sup_records, method=method, status="matched",
                    confidence="high", recon_id=recon_id, rule_hit=rule_hit)
                inherit_legacy(rec, oid)
                results.append(rec)
                matched_cost_ids.update(r.get("record_id") for r in sup_records)
                for a in anoms:
                    add_anom(a["order_id"], a["exception_type"], a["severity"],
                             a["message"], a.get("related_record_ids"))
                continue

            available = [c for c in costs if c.get("record_id") not in matched_cost_ids]
            candidates = self._assisted_candidates(wlines, available)
            if candidates:
                cand_ids = [c.get("record_id") for c in candidates]
                rec, anoms = self._build_match_record(
                    oid, wlines, candidates, method="assisted_candidate",
                    status="suspected", confidence="low", recon_id=recon_id,
                    rule_hit="辅助候选匹配（日期/商品编码/数量线索），待人工确认，禁止自动认定")
                anoms.append({
                    "order_id": oid,
                    "exception_type": "SUSPECTED_MATCH_NEED_REVIEW",
                    "severity": "medium",
                    "message": (f"订单 {oid} 仅能通过辅助线索匹配到 {len(candidates)} 条供货记录"
                                f"（{cand_ids}），已进入待人工确认，不参与自动利润计算"),
                    "related_record_ids": wx_ids + cand_ids,
                })
                inherit_legacy(rec, oid)
                matched_cost_ids.update(cand_ids)
                for a in anoms:
                    add_anom(a["order_id"], a["exception_type"], a["severity"],
                             a["message"], a.get("related_record_ids"))
                rec["exception_codes"] = sorted({a["exception_type"] for a in anoms})
                results.append(rec)
            else:
                rec, anoms = self._build_match_record(
                    oid, wlines, [], method=None, status="unmatched_order",
                    confidence=None, recon_id=recon_id,
                    rule_hit="未找到对应供货成本记录")
                anoms.append({
                    "order_id": oid,
                    "exception_type": "UNMATCHED_ORDER",
                    "severity": "high",
                    "message": f"订单 {oid} 未找到对应供货成本记录（RC-001），支持人工绑定",
                    "related_record_ids": wx_ids,
                })
                inherit_legacy(rec, oid)
                self._inherit_unmatched_manual_cost(rec, oid, old_by_order)
                for a in anoms:
                    add_anom(a["order_id"], a["exception_type"], a["severity"],
                             a["message"], a.get("related_record_ids"))
                rec["exception_codes"] = sorted({a["exception_type"] for a in anoms})
                results.append(rec)

        # 时间范围外订单：保留已有记录。若该单只有手工成本、且现在能匹配到入库的供应商成本，
        # 以入库数据重建匹配并丢掉手工金额。（未发货等不对账订单不保留，重跑时清除旧记录。）
        for r in existing:
            if not (r.get("record_type") == "order_match" and r.get("order_id")
                    and r["order_id"] not in time_scope_order_ids):
                continue
            oid = r["order_id"]
            replaced = False
            if (r.get("manual_cost") and r.get("match_status") == "unmatched_order"
                    and not r.get("supplier_record_ids")):
                wlines = orders_by_id.get(oid) or []
                sup_records, method, rule_hit = self._auto_match(
                    oid, main_index, sub_index, normalized_index)
                if method and wlines:
                    rec, anoms = self._build_match_record(
                        oid, wlines, sup_records, method=method,
                        status="matched", confidence="high", recon_id=recon_id,
                        rule_hit=rule_hit)
                    inherit_legacy(rec, oid)
                    results.append(rec)
                    matched_cost_ids.update(
                        x.get("record_id") for x in sup_records)
                    for a in anoms:
                        add_anom(a["order_id"], a["exception_type"], a["severity"],
                                 a["message"], a.get("related_record_ids"))
                    replaced = True
            if not replaced:
                results.append(r)
                matched_cost_ids.update(r.get("supplier_record_ids") or [])

        # 孤儿供应商记录（未匹配成本，RC-002）；无效/待核对成本不参与对账统计
        for c in costs:
            rid = c.get("record_id")
            if rid in matched_cost_ids:
                continue
            if not is_effective_supplier_cost(c, orders_lookup):
                continue
            key = c.get("main_order_id") or c.get("supplier_order_id") or "未知订单键"
            results.append({
                "record_id": "orphan_" + rid,
                "record_type": "order_match",
                "reconciliation_id": recon_id,
                "order_id": None,
                "supplier_order_key": key,
                "order_time": c.get("created_at"),
                "wechat_record_ids": [],
                "supplier_record_ids": [rid],
                "match_method": None,
                "match_status": "unmatched_supplier",
                "confidence": None,
                "match_detail": {"rule_hit": "供应商成本记录未关联到任何微信订单"},
                "cost": self._aggregate_costs([c]),
                "checks": {},
                "exception_codes": ["UNMATCHED_SUPPLIER_COST"],
                "manual_review": {"required": True, "reviewer": None,
                                  "note": None, "history": []},
                "preserved_manual": False,
                "created_at": now,
                "updated_at": now,
            })
            add_anom(key, "UNMATCHED_SUPPLIER_COST", "medium",
                     f"供货记录 {key}（{rid}）未关联到微信订单，可能未同步、错单或时间范围不一致（RC-002）",
                     [rid])

        # 多对一检查：一条成本记录被多个订单匹配
        cost_to_orders = {}
        for r in results:
            if r.get("record_type") != "order_match" or not r.get("order_id"):
                continue
            for sid in r.get("supplier_record_ids") or []:
                cost_to_orders.setdefault(sid, set()).add(r["order_id"])
        for sid, oids in sorted(cost_to_orders.items()):
            if len(oids) > 1:
                add_anom(None, "MULTI_ORDER_ONE_COST", "high",
                         f"成本记录 {sid} 同时被多个订单匹配: {sorted(oids)}，需人工核实（多对一）",
                         [sid])

        anomalies, added, auto_resolved = self._merge_anomalies(recon_id, fresh)

        def repl_results(d):
            d["records"] = results
            d["updated_at"] = now
        self.repo.update("reconciliation_results", repl_results)

        def repl_anomalies(d):
            d["records"] = anomalies
            d["updated_at"] = now
        self.repo.update("anomalies", repl_anomalies)

        match_hints = {}
        for r in results:
            oid = r.get("order_id")
            if not oid:
                continue
            for sid in r.get("supplier_record_ids") or []:
                match_hints[sid] = oid
        if match_hints:
            revalidate_supplier_cost_records(
                self.repo, match_hints.keys(),
                op_log=self.op_log, operator=operator,
                matched_order_hints=match_hints)

        summary = self._stats_of(results)
        summary["wechat_orders"] = len(reconcile_scope_order_ids)
        status = "completed_with_exceptions" if fresh else "completed"
        self.op_log.append("reconciliation.run", detail={
            "reconciliation_id": recon_id, "start": start, "end": end,
            "status": status, "summary": summary,
            "anomalies_created": added, "anomalies_auto_resolved": auto_resolved,
        }, operator=operator)
        return {
            "reconciliation_id": recon_id,
            "status": status,
            "summary": summary,
            "anomalies_created": added,
            "anomalies_auto_resolved": auto_resolved,
        }

    # ==================== 人工确认匹配 ====================
    def manual_match(self, order_id, supplier_record_ids=None, action="confirm",
                     cost_granularity=None, note=None, operator="admin"):
        if action not in ("confirm", "unbind"):
            raise ReconcileServiceError("action 仅支持 confirm / unbind")
        if cost_granularity not in (None, "order_level", "line_level"):
            raise ReconcileServiceError(
                "cost_granularity 仅支持 order_level（同额多行只扣一次）/ "
                "line_level（各明细行分别累加）")

        orders = self.repo.read("orders")["records"]
        wlines = [o for o in orders if o.get("order_id") == order_id]
        if not wlines:
            raise OrderNotFoundError(f"微信订单不存在: {order_id}")

        costs = self.repo.read("supplier_costs")["records"]
        by_id = {c.get("record_id"): c for c in costs}
        results = self.repo.read("reconciliation_results")["records"]
        now = now_iso()

        if action == "unbind":
            sids = []
        else:
            sids = list(dict.fromkeys(str(s) for s in (supplier_record_ids or [])))
            if not sids:
                raise ReconcileServiceError("confirm 需要至少一条 supplier_record_ids")
            for s in sids:
                if s not in by_id:
                    raise SupplierRecordNotFoundError(f"供应商成本记录不存在: {s}")
            # 冲突保护：已被其他订单人工匹配占用的成本记录禁止重复绑定
            for r in results:
                if (r.get("record_type") == "order_match"
                        and r.get("match_method") == "manual"
                        and r.get("order_id") != order_id):
                    occupied = sorted(set(r.get("supplier_record_ids") or []) & set(sids))
                    if occupied:
                        raise ManualMatchConflictError(
                            f"成本记录 {occupied} 已被订单 {r['order_id']} 的人工匹配占用，请先解除该绑定",
                            data={"conflict_order_id": r["order_id"],
                                  "conflict_record_ids": occupied})

        old = next((r for r in results
                    if r.get("record_type") == "order_match"
                    and r.get("record_id") == "match_" + order_id), None)
        before = self._snapshot(old)
        wx_ids = [l.get("record_id") for l in wlines if l.get("record_id")]

        fresh = []
        if action == "confirm":
            sup_records = [by_id[s] for s in sids]
            rec, anoms = self._build_match_record(
                order_id, wlines, sup_records, method="manual",
                status="manual_confirmed", confidence="high",
                recon_id=(old or {}).get("reconciliation_id") or "manual",
                rule_hit="人工确认匹配", confirmed_granularity=cost_granularity)
            rec["match_detail"]["operator"] = operator
        else:
            rec, anoms = self._build_match_record(
                order_id, wlines, [], method=None, status="unmatched_order",
                confidence=None,
                recon_id=(old or {}).get("reconciliation_id") or "manual",
                rule_hit="人工解除匹配")
            anoms.append({
                "order_id": order_id,
                "exception_type": "UNMATCHED_ORDER",
                "severity": "high",
                "message": f"订单 {order_id} 人工解除匹配后处于未匹配状态（RC-001）",
                "related_record_ids": wx_ids,
            })
        after = self._snapshot(rec)

        history = list((old or {}).get("manual_review", {}).get("history") or [])
        history.append({
            "operator": operator,
            "operated_at": now,
            "action": action,
            "match_rule": "manual",
            "cost_granularity": cost_granularity,
            "note": note,
            "before": before,
            "after": after,
        })
        rec["manual_review"] = {
            "required": rec["manual_review"]["required"],
            "reviewer": operator if action == "confirm" else None,
            "note": note,
            "history": history,
        }
        rec["created_at"] = (old or {}).get("created_at") or rec["created_at"]

        for a in anoms:
            fresh.append(a)

        # 写入匹配记录：替换本订单记录，移除被绑定成本记录的孤儿记录
        new_results = []
        for r in results:
            if (r.get("record_type") == "order_match"
                    and r.get("record_id") == rec["record_id"]):
                continue
            if (r.get("record_type") == "order_match"
                    and r.get("match_status") == "unmatched_supplier"
                    and set(r.get("supplier_record_ids") or []) & set(sids)):
                continue
            new_results.append(r)
        new_results.append(rec)

        def repl_results(d):
            d["records"] = new_results
            d["updated_at"] = now
        self.repo.update("reconciliation_results", repl_results)

        if action == "confirm" and sids:
            revalidate_supplier_cost_records(
                self.repo, sids, op_log=self.op_log, operator=operator,
                matched_order_hints={s: order_id for s in sids})

        # 异常处理：人工确认后关闭相关待审异常（留痕），并合入本次新异常
        anomalies = self.repo.read("anomalies")["records"]
        if action == "confirm":
            self._resolve_pending(
                anomalies, order_id=order_id,
                types=("UNMATCHED_ORDER", "SUSPECTED_MATCH_NEED_REVIEW",
                       "MANUAL_MATCH_BROKEN"),
                text=f"订单 {order_id} 已人工确认匹配关系（操作人: {operator}）",
                operator=operator)
            if sids:
                self._resolve_pending(
                    anomalies, types=("UNMATCHED_SUPPLIER_COST",), record_ids=sids,
                    text=f"成本记录已人工绑定至订单 {order_id}（操作人: {operator}）",
                    operator=operator)
        merged, _, _ = self._merge_anomalies("manual", fresh, auto_close=False,
                                             anomalies=anomalies)

        def repl_anomalies(d):
            d["records"] = merged
            d["updated_at"] = now
        self.repo.update("anomalies", repl_anomalies)

        self.op_log.append("reconciliation.manual_match", detail={
            "order_id": order_id, "action": action,
            "supplier_record_ids": sids, "cost_granularity": cost_granularity,
            "note": note, "before": before, "after": after,
        }, operator=operator)
        return rec

    # ==================== 人工修改业务字段 ====================
    def _assert_modifiable(self, order_id):
        """锁定校验：已匹配订单拒绝人工修改"""
        results = self.repo.read("reconciliation_results")["records"]
        rec = next((r for r in results
                    if r.get("record_type") == "order_match"
                    and r.get("record_id") == "match_" + order_id), None)
        if rec is None:
            raise OrderNotFoundError(f"对账记录不存在: {order_id}")
        status = rec.get("match_status")
        if status in ("matched", "manual_confirmed"):
            raise ModifyRejectedError(
                "该订单已匹配导入数据，不能手动修改。",
                data={"order_id": order_id, "match_status": status})

    def _apply_field_modifications(self, order_id, fields, cost_updates, operator):
        """逐字段校验并写入业务字段，返回 (before_snapshot, after_snapshot, changed_fields)

        订单字段写入 orders 表，成本字段按 supplier_record_id 写入 supplier_costs 表。
        """
        orders = self.repo.read("orders")["records"]
        wlines = [o for o in orders if o.get("order_id") == order_id]
        costs = self.repo.read("supplier_costs")["records"]
        results = self.repo.read("reconciliation_results")["records"]
        rec = next((r for r in results
                    if r.get("record_type") == "order_match"
                    and r.get("record_id") == "match_" + order_id), None)
        if rec is None:
            raise OrderNotFoundError(f"对账记录不存在: {order_id}")

        before = self._snapshot(rec)
        changed = []
        now = now_iso()
        order_changed = False
        cost_changed = False

        for field_path, new_value in fields.items():
            ok, err = _validate_field(field_path, new_value)
            if not ok:
                if field_path in SYSTEM_FIELD_BLACKLIST:
                    raise FieldNotModifiableError(err)
                raise InvalidFieldValueError(err)

            top, _, sub = field_path.partition(".")
            if sub:
                container = (wlines[0] if wlines else {})
                old_val = (container.get(top) or {}).get(sub)
                for wl in wlines:
                    wl.setdefault(top, {})[sub] = new_value
            else:
                container = wlines[0] if wlines else {}
                old_val = container.get(top)
                for wl in wlines:
                    wl[top] = new_value

            changed.append({
                "field": field_path,
                "before": old_val,
                "after": new_value,
            })
            order_changed = True

        valid_sids = set(rec.get("supplier_record_ids") or [])
        cost_by_id = {c.get("record_id"): c for c in costs}
        for cu in (cost_updates or []):
            field_path = cu.get("field")
            sid = cu.get("supplier_record_id")
            new_value = cu.get("value")
            if sid not in valid_sids:
                raise SupplierRecordNotFoundError(
                    f"供应商成本记录 {sid} 不属于订单 {order_id}")
            ok, err = _validate_field(field_path, new_value)
            if not ok:
                if field_path in SYSTEM_FIELD_BLACKLIST:
                    raise FieldNotModifiableError(err)
                raise InvalidFieldValueError(err)

            sc = cost_by_id.get(sid)
            if sc is None:
                raise SupplierRecordNotFoundError(
                    f"供应商成本记录不存在: {sid}")

            top, _, sub = field_path.partition(".")
            if sub:
                old_val = (sc.get(top) or {}).get(sub)
                sc.setdefault(top, {})[sub] = new_value
            else:
                old_val = sc.get(top)
                sc[top] = new_value

            changed.append({
                "field": field_path,
                "supplier_record_id": sid,
                "before": old_val,
                "after": new_value,
            })
            cost_changed = True

        if order_changed and wlines:
            def repl_orders(d):
                d["records"] = orders
                d["updated_at"] = now
            self.repo.update("orders", repl_orders)

        if cost_changed:
            def repl_costs(d):
                d["records"] = costs
                d["updated_at"] = now
            self.repo.update("supplier_costs", repl_costs)

        self._mark_manual(rec, [c["field"] for c in changed], operator)
        rec["updated_at"] = now

        after = self._snapshot(rec)
        return before, after, changed

    def manual_modify(self, order_id, fields, reason, operator="admin", cost_updates=None):
        """人工修改异常订单业务字段并重新校验

        fields 修改订单字段（写入 orders 表）；
        cost_updates 修改成本字段（按 supplier_record_id 写入 supplier_costs 表）。
        两者至少一个非空，向后兼容 cost_updates=None。
        """
        if fields is None:
            fields = {}
        if not isinstance(fields, dict):
            raise InvalidFieldValueError("fields 必须是字典")
        if cost_updates is None:
            cost_updates = []
        if not isinstance(cost_updates, list):
            raise InvalidFieldValueError("cost_updates 必须是列表")
        if not fields and not cost_updates:
            raise InvalidFieldValueError("fields 与 cost_updates 至少一个非空")
        if not reason or not isinstance(reason, str):
            raise InvalidFieldValueError("请提供修改原因")

        orders = self.repo.read("orders")["records"]
        if not any(o.get("order_id") == order_id for o in orders):
            raise OrderNotFoundError(f"微信订单不存在: {order_id}")

        try:
            self._assert_modifiable(order_id)
        except ModifyRejectedError:
            self.op_log.append("modify_rejected", detail={
                "order_id": order_id,
                "reason": "已匹配订单拒绝人工修改",
            }, operator=operator)
            raise

        before, after, changed = self._apply_field_modifications(
            order_id, fields, cost_updates, operator)

        if fields and "order_status" in fields:
            revalidate_supplier_costs_for_order_ids(
                self.repo, {order_id}, op_log=self.op_log, operator=operator)

        recon = self.run_reconciliation(operator=operator)

        results = self.repo.read("reconciliation_results")["records"]
        new_rec = next((r for r in results
                        if r.get("record_type") == "order_match"
                        and r.get("record_id") == "match_" + order_id), None)
        if new_rec:
            self._mark_manual(new_rec, [c["field"] for c in changed], operator)
            new_rec["updated_at"] = now_iso()
            new_results = [r for r in results if not (
                r.get("record_type") == "order_match"
                and r.get("record_id") == "match_" + order_id)]
            new_results.append(new_rec)

            def repl_results(d):
                d["records"] = new_results
                d["updated_at"] = now_iso()
            self.repo.update("reconciliation_results", repl_results)

            if new_rec.get("match_status") in ("matched", "manual_confirmed") \
                    and new_rec.get("manual_modified_fields"):
                self.op_log.append("match_locked", detail={
                    "order_id": order_id,
                    "match_status": new_rec.get("match_status"),
                    "locked_fields": new_rec.get("manual_modified_fields"),
                }, operator=operator)

        self.op_log.append("manual_modify", detail={
            "order_id": order_id,
            "reason": reason,
            "before": before,
            "after": after,
            "changed_fields": changed,
        }, operator=operator)

        return new_rec or rec

    def manual_cost_entry(self, supplier_record_id, total_amount,
                          reason=None, operator="admin"):
        """未匹配供应商成本记录手动录入综合成本

        校验链: 记录不存在 404 → 已匹配拒绝 403 → 金额非法 400;
        保留原始导入值 original_total_amount（仅首次设置）;
        更新 cost_breakdown.total_amount 为人工录入值;
        添加 manual_cost_modified_at/by/source 标记;
        不改变 match_status，不自动认定匹配成功。
        """
        costs = self.repo.read("supplier_costs")["records"]
        sc = next((c for c in costs if c.get("record_id") == supplier_record_id),
                   None)
        if sc is None:
            raise SupplierRecordNotFoundError(
                f"供应商成本记录不存在: {supplier_record_id}")

        results = self.repo.read("reconciliation_results")["records"]
        match_status = "unmatched_supplier"
        for r in results:
            if (r.get("record_type") == "order_match"
                    and supplier_record_id in (r.get("supplier_record_ids") or [])
                    and r.get("match_status") in ("matched", "manual_confirmed")):
                match_status = r["match_status"]
                break
        if match_status in ("matched", "manual_confirmed"):
            raise ModifyRejectedError(
                "该记录已匹配订单，不能手动修改综合成本。",
                data={"supplier_record_id": supplier_record_id,
                      "match_status": match_status})

        if not isinstance(total_amount, str):
            raise InvalidFieldValueError("综合成本必须为字符串")
        total_amount = total_amount.strip()
        if not total_amount:
            raise InvalidFieldValueError("综合成本不能为空")
        try:
            amount = Decimal(total_amount)
        except (ArithmeticError, ValueError):
            raise InvalidFieldValueError("综合成本不是合法金额")
        if amount < 0:
            raise InvalidFieldValueError("综合成本不能为负数")
        quantized = str(amount.quantize(Decimal("0.01")))
        if Decimal(quantized) != amount:
            raise InvalidFieldValueError("综合成本最多保留两位小数")

        breakdown = sc.get("cost_breakdown") or {}
        old_value = breakdown.get("total_amount")
        if "original_total_amount" not in breakdown:
            breakdown["original_total_amount"] = old_value
        breakdown["total_amount"] = quantized
        sc["cost_breakdown"] = breakdown
        now = now_iso()
        sc["manual_cost_modified_at"] = now
        sc["manual_cost_modified_by"] = (
            operator.get("username") if isinstance(operator, dict)
            else operator)
        sc["manual_cost_source"] = "manual"
        sc["updated_at"] = now

        def repl_costs(d):
            d["records"] = costs
            d["updated_at"] = now
        self.repo.update("supplier_costs", repl_costs)

        self.op_log.append("supplier.manual_cost", detail={
            "supplier_record_id": supplier_record_id,
            "reason": reason,
            "before": old_value,
            "after": quantized,
            "match_status": match_status,
        }, operator=operator)

        return {
            "supplier_record_id": supplier_record_id,
            "old_total_amount": old_value,
            "new_total_amount": quantized,
            "original_total_amount": breakdown.get("original_total_amount"),
            "manual_cost_modified_at": now,
            "manual_cost_modified_by": sc["manual_cost_modified_by"],
            "match_status": match_status,
        }

    @staticmethod
    def _parse_nonneg_money(label, value):
        """非负金额字符串，最多两位小数。"""
        if not isinstance(value, str):
            raise InvalidFieldValueError(f"{label}必须为字符串")
        text = value.strip()
        if not text:
            raise InvalidFieldValueError(f"{label}不能为空")
        try:
            amount = Decimal(text)
        except (ArithmeticError, ValueError):
            raise InvalidFieldValueError(f"{label}不是合法金额")
        if amount < 0:
            raise InvalidFieldValueError(f"{label}不能为负数")
        quantized = str(amount.quantize(Decimal("0.01")))
        if Decimal(quantized) != amount:
            raise InvalidFieldValueError(f"{label}最多保留两位小数")
        return quantized

    @staticmethod
    def _apply_manual_order_cost(rec, manual):
        """把手工货品+快递写进对账成本汇总，供列表和利润读取。"""
        cost = dict(rec.get("cost") or {})
        total = manual["total_amount"]
        cost["total_amount"] = total
        cost["raw_line_sum"] = total
        cost["dedup_amount"] = "0.00"
        cost["usable_for_profit"] = True
        cost["missing_amount_lines"] = 0
        cost["duplicate_total_amount_risk"] = False
        cost["order_level_dedup_pending"] = False
        cost["granularity"] = "manual"
        cost["line_count"] = 0
        cost["lines"] = []
        rec["cost"] = cost
        rec["manual_cost"] = dict(manual)

    @staticmethod
    def _inherit_unmatched_manual_cost(rec, order_id, old_by_order):
        """仍未匹配时保留手工成本。已匹配到入库成本的记录不继承。"""
        if rec.get("match_status") != "unmatched_order":
            return
        if rec.get("supplier_record_ids"):
            return
        old = old_by_order.get(order_id) or {}
        manual = old.get("manual_cost") or {}
        if manual.get("total_amount") is None:
            return
        ReconciliationService._apply_manual_order_cost(rec, manual)

    def set_unmatched_order_cost(self, order_id, product_amount, courier_amount,
                                 operator="admin"):
        """未匹配订单录入总成本 = 货品 + 快递。"""
        saved = self.set_unmatched_order_costs([{
            "order_id": order_id,
            "product_amount": product_amount,
            "courier_amount": courier_amount,
        }], operator)
        return saved[0]

    def validate_unmatched_order_costs(self, items):
        """只校验，不写入。批量保存前用来确认全部订单都能改。"""
        self._plan_unmatched_order_costs(items, operator="admin")

    def set_unmatched_order_costs(self, items, operator="admin"):
        """一次写入多笔未匹配订单的货品+快递。任一笔不合法则整批不写。

        不创建供应商记录，匹配状态保持未匹配。
        之后若供应商成本入库并匹配到其中某单，重跑对账时该单以入库金额为准。
        """
        planned, results = self._plan_unmatched_order_costs(items, operator)
        if not planned:
            return []
        now = planned[0]["manual"]["modified_at"]

        def repl(d):
            d["records"] = results
            d["updated_at"] = now
        self.repo.update("reconciliation_results", repl)
        self.op_log.append("reconciliation.manual_order_cost_batch", detail={
            "count": len(planned),
            "orders": [{"order_id": p["order_id"], "before": p["before"],
                        "after": p["manual"]["total_amount"]} for p in planned],
        }, operator=operator)
        return [p["summary"] for p in planned]

    def _plan_unmatched_order_costs(self, items, operator):
        if not isinstance(items, list) or not items:
            raise InvalidFieldValueError("请至少提供一条订单成本")
        if len(items) > 200:
            raise InvalidFieldValueError("一次最多修改 200 条")
        orders = self.repo.read("orders")["records"]
        order_ids = {o.get("order_id") for o in orders if o.get("order_id")}
        results = self.repo.read("reconciliation_results")["records"]
        by_order = {}
        for rec in results:
            if rec.get("record_type") == "order_match" and rec.get("order_id"):
                by_order[rec["order_id"]] = rec
        who = (operator.get("username") if isinstance(operator, dict)
               else operator)
        now = now_iso()
        seen = set()
        planned = []
        for item in items:
            if not isinstance(item, dict):
                raise InvalidFieldValueError("订单成本条目格式不正确")
            order_id = item.get("order_id")
            if not isinstance(order_id, str) or not order_id:
                raise InvalidFieldValueError("订单号不能为空")
            if order_id in seen:
                raise InvalidFieldValueError(f"订单 {order_id} 在本次提交中重复")
            seen.add(order_id)
            if order_id not in order_ids:
                raise OrderNotFoundError(f"微信订单不存在: {order_id}")
            rec = by_order.get(order_id)
            if rec is None:
                raise ModifyRejectedError(
                    f"订单 {order_id} 还没有对账记录，请先执行对账",
                    data={"order_id": order_id})
            if rec.get("match_status") != "unmatched_order" \
                    or rec.get("supplier_record_ids"):
                raise ModifyRejectedError(
                    f"订单 {order_id} 已有入库成本，不能再手工修改总成本",
                    data={"order_id": order_id,
                          "match_status": rec.get("match_status")})
            try:
                product = self._parse_nonneg_money(
                    "货品成本", item.get("product_amount"))
                courier = self._parse_nonneg_money(
                    "快递费", item.get("courier_amount"))
            except InvalidFieldValueError as e:
                raise InvalidFieldValueError(f"订单 {order_id}：{e}")
            total = money_str(Decimal(product) + Decimal(courier))
            before = (rec.get("manual_cost") or {}).get("total_amount")
            manual = {
                "product_amount": product,
                "courier_amount": courier,
                "total_amount": total,
                "source": "manual",
                "modified_at": now,
                "modified_by": who,
            }
            planned.append({
                "order_id": order_id,
                "rec": rec,
                "before": before,
                "manual": manual,
                "summary": {
                    "order_id": order_id,
                    "match_status": rec.get("match_status"),
                    "product_amount": product,
                    "courier_amount": courier,
                    "total_amount": total,
                    "modified_at": now,
                    "modified_by": who,
                },
            })
        for item in planned:
            rec = item["rec"]
            self._apply_manual_order_cost(rec, item["manual"])
            self._mark_manual(
                rec, ["manual_cost.product_amount", "manual_cost.courier_amount",
                      "manual_cost.total_amount"], operator)
            rec["updated_at"] = now
        return planned, results

    def get_modifiable_fields(self, order_id):
        """返回订单匹配状态与可修改字段定义"""
        orders = self.repo.read("orders")["records"]
        wlines = [o for o in orders if o.get("order_id") == order_id]
        if not wlines:
            raise OrderNotFoundError(f"微信订单不存在: {order_id}")

        results = self.repo.read("reconciliation_results")["records"]
        rec = next((r for r in results
                    if r.get("record_type") == "order_match"
                    and r.get("record_id") == "match_" + order_id), None)
        match_status = rec.get("match_status") if rec else None
        modifiable = match_status in (None, "unmatched_order", "suspected",
                                      "unmatched_supplier")

        def field_info(field_path, spec):
            top, _, sub = field_path.partition(".")
            wl = wlines[0] if wlines else {}
            current = (wl.get(top) or {}).get(sub) if sub else wl.get(top)
            return {
                "field": field_path,
                "label": spec["label"],
                "type": spec["type"],
                "current_value": current,
            }

        order_fields = [field_info(f, s) for f, s in MODIFIABLE_ORDER_FIELDS.items()]
        cost_fields = []
        if rec and rec.get("supplier_record_ids"):
            costs = self.repo.read("supplier_costs")["records"]
            by_id = {c.get("record_id"): c for c in costs}
            for sid in rec["supplier_record_ids"]:
                sc = by_id.get(sid)
                if sc:
                    for f, s in MODIFIABLE_COST_FIELDS.items():
                        top, _, sub = f.partition(".")
                        current = (sc.get(top) or {}).get(sub) if sub else sc.get(top)
                        cost_fields.append({
                            "field": f,
                            "label": s["label"],
                            "type": s["type"],
                            "current_value": current,
                            "supplier_record_id": sid,
                        })

        return {
            "order_id": order_id,
            "match_status": match_status,
            "modifiable": modifiable,
            "order_fields": order_fields,
            "cost_fields": cost_fields,
            "system_fields": sorted(SYSTEM_FIELD_BLACKLIST),
        }

    # ==================== 查询与导出 ====================
    def list_results(self, start=None, end=None, order_id=None, match_status=None,
                     anomaly_type=None, page=1, size=50):
        records = [self._init_data_source(r) for r in self.repo.read("reconciliation_results")["records"]
                   if r.get("record_type") == "order_match"]
        rows = self._filter_rows(records, start=start, end=end, order_id=order_id,
                                  match_status=match_status, anomaly_type=anomaly_type)
        total = len(rows)
        start_idx = (page - 1) * size
        return {
            "total": total, "page": page, "size": size,
            "pages": (total + size - 1) // size,
            "results": rows[start_idx:start_idx + size],
            "statistics": self._stats_of(rows),
        }

    def export_excel(self, operator="admin", **filters):
        records = [self._init_data_source(r) for r in self.repo.read("reconciliation_results")["records"]
                   if r.get("record_type") == "order_match"]
        rows = self._filter_rows(records, **filters)
        if not rows:
            raise ReconcileServiceError("当前筛选无数据可导出")

        headers = ["订单号", "日期时间", "匹配状态", "匹配方式", "供货明细数",
                   "微信实收金额", "供货实收金额", "金额差异", "数量差异",
                   "综合成本(去重后)", "直接累加值", "去重扣除", "成本粒度",
                   "重复风险", "粒度待确认", "可用于利润计算", "异常标记",
                   "最后操作人", "更新时间"]

        def row_of(r):
            c = r.get("cost") or {}
            k = r.get("checks") or {}
            oid = r.get("order_id") or f"(供货){r.get('supplier_order_key')}"
            return [
                oid, r.get("order_time"),
                MATCH_STATUS_LABELS.get(r.get("match_status"), r.get("match_status") or "—"),
                MATCH_METHOD_LABELS.get(r.get("match_method"), r.get("match_method") or "—"),
                c.get("line_count") or 0,
                _dec(k["wechat_received_amount"])
                if k.get("wechat_received_amount") is not None else None,
                _dec(k["supplier_received_amount"])
                if k.get("supplier_received_amount") is not None else None,
                _dec(k["amount_diff"]) if k.get("amount_diff") is not None else None,
                k.get("quantity_diff") if k.get("quantity_diff") is not None else None,
                _dec(c["total_amount"]) if c.get("total_amount") is not None else None,
                _dec(c["raw_line_sum"]) if c.get("raw_line_sum") is not None else None,
                _dec(c["dedup_amount"]) if c.get("dedup_amount") is not None else None,
                GRANULARITY_LABELS.get(c.get("granularity"), c.get("granularity") or "—"),
                "是" if c.get("duplicate_total_amount_risk") else "否",
                "是" if c.get("order_level_dedup_pending") else "否",
                "是" if c.get("usable_for_profit") else "否",
                "；".join(ANOMALY_TYPE_LABELS.get(t, t)
                          for t in r.get("exception_codes") or []),
                (r.get("manual_review") or {}).get("reviewer") or "",
                r.get("updated_at"),
            ]

        wb = Workbook()
        ws = wb.active
        ws.title = "对账结果"
        append_row(ws, headers)
        for r in rows:
            append_row(ws, row_of(r))
        for col in (6, 7, 8, 10, 11, 12):
            for cell in ws.iter_cols(min_col=col, max_col=col, min_row=2):
                for c in cell:
                    c.number_format = "0.00"
        buf = io.BytesIO()
        wb.save(buf)

        filename = f"对账结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        record = {
            "record_id": ("exp_" + datetime.now().strftime("%Y%m%d%H%M%S") + "_"
                          + str(len(self.repo.read("export_records")["records"]) + 1)),
            "kind": "reconciliation", "file_name": filename, "row_count": len(rows),
            "filters": {k: v for k, v in filters.items() if v is not None},
            "exported_at": now_iso(), "exported_by": operator,
        }

        def mutator(d):
            d["records"].append(record)
            d["updated_at"] = now_iso()
        self.repo.update("export_records", mutator)
        self.op_log.append("data.export", detail={
            "kind": "reconciliation", "row_count": len(rows), "file_name": filename,
        }, operator=operator)
        return filename, buf.getvalue()

    def get_options(self):
        return {
            "match_status": MATCH_STATUS_LABELS,
            "match_method": MATCH_METHOD_LABELS,
            "granularity": GRANULARITY_LABELS,
            "confirmed_granularity": CONFIRMED_GRANULARITY_LABELS,
            "anomaly_types": ANOMALY_TYPE_LABELS,
            "amount_tolerance": str(AMOUNT_TOLERANCE),
            "page_sizes": list(PAGE_SIZES),
        }

    # ==================== 内部：匹配 ====================
    @staticmethod
    def _auto_match(order_key, main_index, sub_index, normalized_index):
        """按优先级①②③自动匹配，返回 (记录列表, 匹配方式, 规则说明)"""
        recs = main_index.get(order_key)
        if recs:
            return recs, "main_order_id_exact", "微信订单号 = 供货表主订单编号"
        recs = sub_index.get(order_key)
        if recs:
            return recs, "supplier_order_id_exact", "微信订单号 = 供货表订单编号"
        nk = _norm_key(order_key)
        if nk:
            recs = normalized_index.get(nk)
            if recs:
                return recs, "normalized_equal", "订单号标准化（去空格/不可见字符）后相等"
        return [], None, None

    @staticmethod
    def _assisted_candidates(wlines, available_costs, limit=5):
        """辅助候选匹配：SKU 线索命中且（日期或数量）命中；仅作疑似候选"""
        ot = next((l.get("order_time") for l in wlines if l.get("order_time")), None)
        ot_day = ot[:10] if ot else None
        wx_skus = {(l.get("product") or {}).get("sku_code") for l in wlines} - {None}
        wx_qtys = {(l.get("product") or {}).get("quantity") for l in wlines} - {None}
        out = []
        for c in available_costs:
            p = c.get("product") or {}
            sku_hit = bool(p.get("sku_code") and p.get("sku_code") in wx_skus)
            if not sku_hit:
                continue
            cat_day = (c.get("created_at") or "")[:10] or None
            date_hit = bool(ot_day and cat_day and ot_day == cat_day)
            qty_hit = bool(p.get("quantity") and p.get("quantity") in wx_qtys)
            if date_hit or qty_hit:
                out.append(c)
                if len(out) >= limit:
                    break
        return out

    # ==================== 内部：成本汇总与差异检查 ====================
    @staticmethod
    def _has_repeated_amount(lines):
        seen = {}
        for ln in lines:
            amt = ln["total_amount"]
            if amt is None:
                continue
            seen[amt] = seen.get(amt, 0) + 1
        return any(n > 1 for n in seen.values())

    def _aggregate_costs(self, sup_records, confirmed_granularity=None):
        """成本汇总：区分订单级/物流单级/商品行级，同额重复禁止直接累加"""
        lines = []
        for r in sup_records:
            cb = r.get("cost_breakdown") or {}
            lines.append({
                "record_id": r.get("record_id"),
                "total_amount": cb.get("total_amount"),
                "tracking_number": (r.get("tracking") or {}).get("tracking_number"),
            })
        if not lines:
            return {
                "granularity": "unknown", "confirmed_granularity": confirmed_granularity,
                "line_count": 0, "raw_line_sum": None, "total_amount": None,
                "dedup_amount": None, "duplicate_total_amount_risk": False,
                "order_level_dedup_pending": False, "missing_amount_lines": 0,
                "usable_for_profit": False, "lines": [],
            }

        missing_count = sum(1 for ln in lines if ln["total_amount"] is None)
        entries = []
        duplicate_risk = False
        dedup_pending = False

        if confirmed_granularity == "line_level":
            for ln in lines:
                entries.append({
                    "granularity": ("package_level" if ln["tracking_number"]
                                    else "supplier_line"),
                    "total_amount": ln["total_amount"],
                    "record_ids": [ln["record_id"]],
                    "duplicated_record_ids": [],
                })
            if self._has_repeated_amount(lines):
                duplicate_risk = True
        elif confirmed_granularity == "order_level":
            by_amount = {}
            for ln in lines:
                by_amount.setdefault(ln["total_amount"], []).append(ln)
            for amt, lns in by_amount.items():
                entries.append({
                    "granularity": "order_level",
                    "total_amount": amt,
                    "record_ids": [lns[0]["record_id"]],
                    "duplicated_record_ids": [ln["record_id"] for ln in lns[1:]],
                })
            if self._has_repeated_amount(lines):
                duplicate_risk = True
        else:
            by_amount = {}
            for ln in lines:
                by_amount.setdefault(ln["total_amount"], []).append(ln)
            for amt, lns in by_amount.items():
                if amt is None or len(lns) == 1:
                    for ln in lns:
                        entries.append({
                            "granularity": ("package_level" if ln["tracking_number"]
                                            else "supplier_line"),
                            "total_amount": ln["total_amount"],
                            "record_ids": [ln["record_id"]],
                            "duplicated_record_ids": [],
                        })
                    continue
                trk_numbers = [ln["tracking_number"] for ln in lns]
                if all(trk_numbers) and len(set(trk_numbers)) == len(trk_numbers):
                    # 物流单级：同额但快递单号互不相同 → 每个包裹分别计入
                    for ln in lns:
                        entries.append({
                            "granularity": "package_level",
                            "total_amount": ln["total_amount"],
                            "record_ids": [ln["record_id"]],
                            "duplicated_record_ids": [],
                        })
                    duplicate_risk = True
                else:
                    # 疑似订单级综合成本：去重只计一次，
                    # 粒度未经人工确认前不得用于最终利润计算
                    entries.append({
                        "granularity": "order_level",
                        "total_amount": amt,
                        "record_ids": [lns[0]["record_id"]],
                        "duplicated_record_ids": [ln["record_id"] for ln in lns[1:]],
                    })
                    duplicate_risk = True
                    dedup_pending = True

        raw_sum = sum((_dec(ln["total_amount"]) for ln in lines), Decimal("0"))
        total = sum((_dec(e["total_amount"]) for e in entries), Decimal("0"))
        granularities = {e["granularity"] for e in entries
                         if e["total_amount"] is not None}
        if len(granularities) > 1:
            granularity = "mixed"
        elif granularities:
            granularity = next(iter(granularities))
        else:
            granularity = "unknown"
        return {
            "granularity": granularity,
            "confirmed_granularity": confirmed_granularity,
            "line_count": len(lines),
            "raw_line_sum": money_str(raw_sum),
            "total_amount": money_str(total),
            "dedup_amount": money_str(raw_sum - total),
            "duplicate_total_amount_risk": duplicate_risk,
            "order_level_dedup_pending": dedup_pending,
            "missing_amount_lines": missing_count,
            "usable_for_profit": (not dedup_pending) and (not missing_count),
            "lines": entries,
        }

    @staticmethod
    def _check_differences(order_id, wlines, sup_records):
        """差异检查：金额（微信实收 vs 供货实收，容差 0.01）与商品数量"""
        anomalies = []
        wx_received, _ = order_level_amount(
            wlines, lambda l: (l.get("amounts") or {}).get("actual_received_amount"))
        wx_qty = sum((l.get("product") or {}).get("quantity") or 0 for l in wlines)
        sup_vals = [r.get("supplier_actual_received") for r in sup_records
                    if r.get("supplier_actual_received") not in (None, "")]
        sup_received, sup_conflict = None, False
        if sup_vals:
            uniq = []
            for v in sup_vals:
                if v not in uniq:
                    uniq.append(v)
            sup_received = uniq[0]
            sup_conflict = len(uniq) > 1
        sup_qty = sum((r.get("product") or {}).get("quantity") or 0 for r in sup_records)

        amount_diff = None
        amount_match = None
        if wx_received is not None and sup_received is not None:
            amount_diff = _dec(wx_received) - _dec(sup_received)
            amount_match = abs(amount_diff) <= AMOUNT_TOLERANCE
            if not amount_match:
                anomalies.append({
                    "order_id": order_id,
                    "exception_type": "AMOUNT_DIFF",
                    "severity": "medium",
                    "message": (f"订单 {order_id} 微信实收 {wx_received} 与供货实收 "
                                f"{sup_received} 差异 {money_str(amount_diff)}，"
                                f"超出容差 {AMOUNT_TOLERANCE}（RC-012）"),
                })
        if sup_conflict:
            anomalies.append({
                "order_id": order_id,
                "exception_type": "AMOUNT_DIFF",
                "severity": "medium",
                "message": (f"订单 {order_id} 供货侧订单实收款多行存在不同值 "
                            f"{sorted(set(sup_vals))}，已取首值核对，请人工核实"),
            })

        quantity_diff = (wx_qty - sup_qty) if sup_records else None
        quantity_match = (wx_qty == sup_qty) if sup_records else None
        if quantity_match is False:
            anomalies.append({
                "order_id": order_id,
                "exception_type": "QUANTITY_DIFF",
                "severity": "medium",
                "message": (f"订单 {order_id} 微信商品数量 {wx_qty} 与供货数量 {sup_qty} "
                            f"相差 {quantity_diff}，请核实拆单/漏发/退款换货（RC-005）"),
            })

        checks = {
            "wechat_received_amount": wx_received,
            "supplier_received_amount": sup_received,
            "amount_diff": money_str(amount_diff) if amount_diff is not None else None,
            "amount_match": amount_match,
            "wechat_quantity": wx_qty,
            "supplier_quantity": sup_qty if sup_records else None,
            "quantity_diff": quantity_diff,
            "quantity_match": quantity_match,
        }
        return checks, anomalies

    def _init_data_source(self, rec):
        """历史记录兼容兜底：为缺失数据来源字段填充默认值"""
        if "data_source" not in rec:
            rec["data_source"] = {}
        if "manual_modified_fields" not in rec:
            rec["manual_modified_fields"] = []
        if "last_import_batch_id" not in rec:
            rec["last_import_batch_id"] = None
        if "last_modified_at" not in rec:
            rec["last_modified_at"] = None
        if "last_modified_by" not in rec:
            rec["last_modified_by"] = None
        return rec

    def _mark_manual(self, rec, fields, operator):
        """标记字段为人工修改"""
        now = now_iso()
        ds = rec.setdefault("data_source", {})
        mf = rec.setdefault("manual_modified_fields", [])
        for f in fields:
            ds[f] = "manual"
            if f not in mf:
                mf.append(f)
        rec["last_modified_at"] = now
        rec["last_modified_by"] = operator

    def _mark_import(self, rec, fields, batch_id):
        """标记字段为导入数据覆盖，清除人工修改标记"""
        ds = rec.setdefault("data_source", {})
        mf = rec.setdefault("manual_modified_fields", [])
        for f in fields:
            ds[f] = "import"
            if f in mf:
                mf.remove(f)
        if batch_id:
            rec["last_import_batch_id"] = batch_id

    def _build_match_record(self, order_id, wlines, sup_records, method, status,
                            confidence, recon_id, rule_hit=None,
                            confirmed_granularity=None):
        cost = self._aggregate_costs(sup_records, confirmed_granularity)
        checks, anomalies = self._check_differences(order_id, wlines, sup_records)
        wx_ids = [l.get("record_id") for l in wlines if l.get("record_id")]
        sup_ids = [r.get("record_id") for r in sup_records]
        related = wx_ids + sup_ids

        if cost["duplicate_total_amount_risk"]:
            if confirmed_granularity == "order_level":
                anomalies.append({
                    "order_id": order_id,
                    "exception_type": "DUPLICATE_TOTAL_AMOUNT_RISK",
                    "severity": "medium",
                    "message": (f"订单 {order_id} 总金额在多行重复出现，"
                                f"已人工确认为订单级综合成本，去重只计一次"
                                f"（{cost['total_amount']} 元，RC-007）"),
                })
            elif confirmed_granularity == "line_level":
                anomalies.append({
                    "order_id": order_id,
                    "exception_type": "DUPLICATE_TOTAL_AMOUNT_RISK",
                    "severity": "medium",
                    "message": (f"订单 {order_id} 总金额在多行重复出现，"
                                f"已人工确认为行级明细，各明细行分别累加"
                                f"（{cost['total_amount']} 元，RC-007）"),
                })
            elif cost["order_level_dedup_pending"]:
                anomalies.append({
                    "order_id": order_id,
                    "exception_type": "DUPLICATE_TOTAL_AMOUNT_RISK",
                    "severity": "high",
                    "message": (f"订单 {order_id} 总金额在多行重复出现，疑似订单级综合成本，"
                                f"已去重只计一次（{cost['total_amount']} 元，直接累加为 "
                                f"{cost['raw_line_sum']} 元）；粒度未经人工确认前"
                                f"不得用于最终利润计算（RC-007）"),
                })
            else:
                anomalies.append({
                    "order_id": order_id,
                    "exception_type": "DUPLICATE_TOTAL_AMOUNT_RISK",
                    "severity": "medium",
                    "message": (f"订单 {order_id} 总金额在多行重复出现，快递单号互不相同，"
                                f"已按物流单级分别计入（RC-007）"),
                })
        if cost["missing_amount_lines"]:
            anomalies.append({
                "order_id": order_id,
                "exception_type": "COST_MISSING",
                "severity": "high",
                "message": (f"订单 {order_id} 有 {cost['missing_amount_lines']} 行供应商明细"
                            f"总金额为空，利润不可计算，不会默认按 0 处理（RC-006）"),
            })
        for a in anomalies:
            a.setdefault("related_record_ids", related)

        rec = {
            "record_id": "match_" + order_id,
            "record_type": "order_match",
            "reconciliation_id": recon_id,
            "order_id": order_id,
            "order_time": next((l.get("order_time") for l in wlines
                                if l.get("order_time")), None),
            "wechat_record_ids": wx_ids,
            "supplier_record_ids": sup_ids,
            "match_method": method,
            "match_status": status,
            "confidence": confidence,
            "match_detail": {"rule_hit": rule_hit},
            "cost": cost,
            "checks": checks,
            "exception_codes": sorted({a["exception_type"] for a in anomalies}),
            "manual_review": {
                "required": status in ("suspected", "unmatched_order"),
                "reviewer": None, "note": None, "history": [],
            },
            "preserved_manual": False,
            "data_source": {},
            "manual_modified_fields": [],
            "last_import_batch_id": None,
            "last_modified_at": None,
            "last_modified_by": None,
            "created_at": now_iso(),
            "updated_at": now_iso(),
        }
        return rec, anomalies

    # ==================== 内部：异常与工具 ====================
    def _merge_anomalies(self, recon_id, fresh, auto_close=True, anomalies=None):
        """幂等合并异常：同键 pending 更新不重复新增；
        重跑对账后未再出现的旧对账异常自动关闭（必须留痕）。
        anomalies 可由调用方传入已修改的列表（如人工处理后预置 resolved 状态）"""
        if anomalies is None:
            anomalies = self.repo.read("anomalies")["records"]
        now = now_iso()
        index = {}
        for i, a in enumerate(anomalies):
            index.setdefault((a.get("order_id"), a.get("exception_type")), []).append(i)

        added = 0
        for na in fresh:
            key = (na["order_id"], na["exception_type"])
            updated = False
            for i in index.get(key, []):
                if anomalies[i].get("status") == "pending_review":
                    anomalies[i]["message"] = na["message"]
                    anomalies[i]["severity"] = na["severity"]
                    anomalies[i]["reconciliation_id"] = recon_id
                    anomalies[i]["related_record_ids"] = na.get(
                        "related_record_ids", [])
                    anomalies[i]["updated_at"] = now
                    updated = True
                    break
            if not updated:
                anomalies.append({
                    "record_id": "exc_" + uuid.uuid4().hex,
                    "batch_id": None,
                    "reconciliation_id": recon_id,
                    "order_id": na["order_id"],
                    "exception_type": na["exception_type"],
                    "severity": na["severity"],
                    "message": na["message"],
                    "related_record_ids": na.get("related_record_ids", []),
                    "status": "pending_review",
                    "resolution": None,
                    "created_at": now,
                })
                index.setdefault(key, []).append(len(anomalies) - 1)
                added += 1

        resolved = 0
        if auto_close:
            fresh_keys = {(na["order_id"], na["exception_type"]) for na in fresh}
            for a in anomalies:
                if (a.get("status") == "pending_review"
                        and a.get("reconciliation_id")
                        and a.get("batch_id") is None
                        and a.get("exception_type") in RECON_ANOMALY_TYPES
                        and (a.get("order_id"), a.get("exception_type"))
                        not in fresh_keys):
                    a["status"] = "resolved"
                    a["resolution"] = {
                        "text": "重新对账后该异常未再出现，系统自动关闭",
                        "resolved_by": "system",
                        "resolved_at": now,
                    }
                    resolved += 1
        return anomalies, added, resolved

    @staticmethod
    def _resolve_pending(anomalies, order_id=None, types=(), record_ids=None,
                         text="", operator="admin"):
        now = now_iso()
        for a in anomalies:
            if a.get("status") != "pending_review":
                continue
            if order_id is not None and a.get("order_id") != order_id:
                continue
            if types and a.get("exception_type") not in types:
                continue
            if record_ids is not None:
                rel = set(a.get("related_record_ids") or [])
                if not (rel & set(record_ids)):
                    continue
            a["status"] = "resolved"
            a["resolution"] = {"text": text, "resolved_by": operator,
                               "resolved_at": now}

    @staticmethod
    def _snapshot(r):
        if not r:
            return {
                "match_status": "unmatched_order", "match_method": None,
                "supplier_record_ids": [], "cost_total_amount": None,
                "cost_granularity": None, "usable_for_profit": False,
            }
        cost = r.get("cost") or {}
        return {
            "match_status": r.get("match_status"),
            "match_method": r.get("match_method"),
            "supplier_record_ids": list(r.get("supplier_record_ids") or []),
            "cost_total_amount": cost.get("total_amount"),
            "cost_granularity": cost.get("confirmed_granularity"),
            "usable_for_profit": cost.get("usable_for_profit"),
        }

    @staticmethod
    def _filter_rows(records, start=None, end=None, order_id=None,
                     match_status=None, anomaly_type=None):
        rows = []
        for r in records:
            if order_id and (r.get("order_id") != order_id
                             and r.get("supplier_order_key") != order_id):
                continue
            if match_status and r.get("match_status") != match_status:
                continue
            if anomaly_type and anomaly_type not in (r.get("exception_codes") or []):
                continue
            if not in_range(r.get("order_time"), start, end):
                continue
            rows.append(r)
        rows.sort(key=lambda r: (r.get("order_time") or "", r.get("record_id") or ""))
        rows.reverse()
        return rows

    @staticmethod
    def _stats_of(rows):
        stats = {
            "total": len(rows), "matched": 0, "manual_confirmed": 0, "suspected": 0,
            "unmatched_order": 0, "unmatched_supplier": 0, "duplicate_risk": 0,
            "profit_blocked": 0,
        }
        for r in rows:
            ms = r.get("match_status")
            if ms in stats:
                stats[ms] += 1
            cost = r.get("cost") or {}
            if cost.get("duplicate_total_amount_risk"):
                stats["duplicate_risk"] += 1
            if r.get("supplier_record_ids") and not cost.get("usable_for_profit", True):
                stats["profit_blocked"] += 1
        return stats

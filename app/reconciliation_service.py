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
6. 异常：未匹配订单(RC-001)、未匹配成本(RC-002)、总金额重复(RC-007)、
   数量差异(RC-005)、成本缺失(RC-006)、金额差异(RC-012)、多对一、疑似匹配待确认、
   人工关系失效；异常禁止静默丢弃，自动关闭必须留痕。
"""
import io
import uuid
from datetime import datetime
from decimal import Decimal

from openpyxl import Workbook

from .data_schemas import now_iso
from .errors import (ManualMatchConflictError, OrderNotFoundError,
                     ReconcileServiceError, SupplierRecordNotFoundError)
from .normalizer import INVISIBLE_CHARS, money_str
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

        main_index, sub_index, normalized_index = {}, {}, {}
        seen_norm = set()
        for c in costs:
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

        scope_order_ids = set()
        for oid, wlines in orders_by_id.items():
            ot = next((l.get("order_time") for l in wlines if l.get("order_time")), None)
            if start and (ot is None or ot < start):
                continue
            if end and (ot is None or ot > end):
                continue
            scope_order_ids.add(oid)

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

        for oid in sorted(scope_order_ids):
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
                for a in anoms:
                    add_anom(a["order_id"], a["exception_type"], a["severity"],
                             a["message"], a.get("related_record_ids"))
                rec["exception_codes"] = sorted({a["exception_type"] for a in anoms})
                results.append(rec)

        # 范围外订单：保留已有记录（人工关系与自动结果均保留），其成本引用继续占用
        for r in existing:
            if (r.get("record_type") == "order_match" and r.get("order_id")
                    and r["order_id"] not in scope_order_ids):
                results.append(r)
                matched_cost_ids.update(r.get("supplier_record_ids") or [])

        # 孤儿供应商记录（未匹配成本，RC-002）
        for c in costs:
            rid = c.get("record_id")
            if rid in matched_cost_ids:
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

        summary = self._stats_of(results)
        summary["wechat_orders"] = len(scope_order_ids)
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

    # ==================== 查询与导出 ====================
    def list_results(self, start=None, end=None, order_id=None, match_status=None,
                     anomaly_type=None, page=1, size=50):
        records = [r for r in self.repo.read("reconciliation_results")["records"]
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
        records = [r for r in self.repo.read("reconciliation_results")["records"]
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
        ws.append(headers)
        for r in rows:
            ws.append(row_of(r))
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

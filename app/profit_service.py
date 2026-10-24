"""利润核算与经营总览服务（阶段5，spec 5.1 / 5.8）

核心规则（全部使用 Decimal 计算，禁止 float）：
1. 净收入 = 实际收款金额 − 已确认成功退款金额；仅当退款口径=not_deducted（实收未扣退款）
   时才扣减；口径=deducted 时禁止重复扣减；退款处理中(pending)金额禁止计入成功退款。
2. 退款口径=unknown 时，存在退款的订单利润状态必须为"待确认"（pending_confirmation）。
3. 成本只扣供应商"总金额"（综合成本）一次；成本/快递费/耗材/加收/燃油费等分项
   仅用于展示与核对，禁止重复扣减。
4. 平台补贴仅在口径确认=not_included（未含于实收）时计入利润；unknown 时单列展示不计入。
5. 经营贡献利润 = 净收入 + 计入补贴 − 供应商综合总金额 − 技术服务费 − 其他经营费用。
6. 利润率 = 经营贡献利润 / 净收入，仅净收入 > 0 时计算；净收入 ≤ 0 时不计算并说明原因。
7. 口径未确认、成本重复风险未解决、必要数据缺失（缺实收/缺成本/未匹配）时，
   利润状态为"待确认"或"不可计算"，禁止输出误导性最终利润。
8. 每次核算记录规则版本、数据批次、核算时间与结果组成项；重算追加新版本，禁止覆盖历史。

P1.1 新增（利润页四核心指标，2026-10 真实数据口径确认）：
9. 商品销售收入 = Σ 微信订单行 product.actual_total_price（源字段"商品实际价格(总共)"）。
10. 产品成本 = Σ 供应商 cost_breakdown.product_supply_cost（源字段"成本"，整行产品成本，
    已含数量；同一订单按对账去重后的供应商行汇总，避免同一产品成本重复计入）；
    任一匹配行分项缺失时为 null，显示"待核对"，禁止默认按 0 当作已确认。
11. 快递费 = Σ cost_breakdown.courier_fee_total（源字段"快递费合计"= 快递费 + 加收 +
    燃油费/打包费等，已包含全部物流相关费用，禁止再逐项累加导致重复计算；
    统计粒度为供应商成本行（物流单级去重后），同一物流单不重复计算）；
    分项缺失时为 null 显示"待核对"。
12. 综合成本 = 供应商"总金额"（产品成本 + 快递费合计等组成），利润核算中只扣一次；
    禁止将综合成本与产品成本、快递费再次相加作为总成本（重复计算）。
13. 产品毛利 = 商品销售收入 − 产品成本（独立经营指标，不参与经营贡献利润计算，
    不得与经营贡献利润相加；分项缺失时显示"待核对"）。
14. 扣除广告费后经营利润 = 经营贡献利润合计 − 对应统计月份广告花费总金额；
    广告费独立核算（月度登记，唯一可人工编辑的财务金额），不得混入供应商综合成本，
    亦不做订单/商品级分摊（订单明细中广告费用分摊列显示"未分摊"）。
"""
import io
import uuid
from datetime import datetime
from decimal import Decimal

from openpyxl import Workbook

from .data_schemas import now_iso
from .errors import ProfitServiceError
from .normalizer import money_str
from .refund_view import build_refund_views, order_level_amount
from .time_range import in_range

RULES_VERSION = "P1.1"
PAGE_SIZES = (20, 50, 100)

PROFIT_STATUS_LABELS = {
    "calculated": "已计算",
    "pending_confirmation": "待确认",
    "not_calculable": "不可计算",
    "needs_review": "需复核",
}
COST_MATCH_LABELS = {
    "matched": "已匹配（成本分项齐全）",
    "parts_missing": "待核对（供应商分项缺失）",
    "suspected": "疑似匹配待确认",
    "unmatched": "未匹配",
    "cost_missing": "成本总金额缺失",
}
FIELD_SOURCES = {
    "product_sales_amount": "微信订单·商品实际价格(总共)",
    "product_cost": "供应商·成本（整行产品成本，含数量）",
    "courier_fee": "供应商·快递费合计（=快递费+加收+燃油费等，不再逐项累加）",
    "composite_cost": "供应商·总金额（综合成本，利润中唯一扣减项）",
    "product_gross_profit": "= 商品销售收入 − 产品成本",
    "technical_service_fee": "微信订单·技术服务费",
    "ad_expense": "广告费用管理页·月度广告花费总金额（人工登记）",
}
SORT_FIELDS = {
    "order_time": "下单时间",
    "order_id": "订单号",
    "product_sales": "商品销售收入",
    "product_cost": "产品成本",
    "courier_fee": "快递费",
    "composite_cost": "综合成本",
    "product_gross_profit": "产品毛利",
    "net_revenue": "净收入",
    "operating_profit": "经营贡献利润",
}
REFUND_DEDUCTION_LABELS = {
    "unknown": "未确认（存在退款时利润显示待确认）",
    "not_deducted": "未扣除（净收入 = 实收 − 成功退款）",
    "deducted": "已扣除（净收入 = 实收，禁止重复扣减）",
}
SUBSIDY_POLICY_LABELS = {
    "unknown": "未确认（补贴单列展示，不计入利润）",
    "not_included": "未含于实收（已确认，计入利润）",
    "included": "已含于实收（不重复计入）",
}
STATUS_REASON_LABELS = {
    "RECEIVED_AMOUNT_MISSING": "实际收款金额缺失",
    "UNMATCHED_NO_COST": "订单未匹配到供应商成本，成本未知",
    "SUPPLIER_COST_MISSING": "供应商成本总金额缺失",
    "SUSPECTED_MATCH_NOT_CONFIRMED": "匹配关系为疑似匹配，待人工确认",
    "REFUND_DEDUCTION_UNCONFIRMED": "退款口径未确认（实收是否已扣退款未知）",
    "DUPLICATE_RISK_UNRESOLVED": "供应商总金额重复风险未解决（粒度未确认）",
    "RECEIVED_AMOUNT_CONFLICT": "订单多行实收金额不一致，已取首值需复核",
}
MARGIN_BLOCKED_LABELS = {
    "NET_REVENUE_NOT_POSITIVE": "净收入不大于零，不计算常规利润率",
    "NOT_CALCULABLE": "利润不可计算",
    "PENDING_CONFIRMATION": "利润待确认",
}
ANOMALY_TYPE_LABELS = {
    "REFUND_DEDUCTION_UNCONFIRMED": "退款口径未确认",
    "NEGATIVE_PROFIT": "负利润订单（RC-009）",
}
PROFIT_ANOMALY_TYPES = set(ANOMALY_TYPE_LABELS)
CALCULATED_LIKE = ("calculated", "needs_review")

CALIBER_DECLARATION = (
    "口径声明：本利润为经营贡献利润；未导入广告推广、达人佣金、运费险、赔付、"
    "人工、仓储、税费等数据时，不代表完整会计净利润。"
)
COST_PART_LABELS = {
    "product_supply_cost": "成本（产品供货成本）",
    "courier_fee": "快递费",
    "foam_material_fee": "泡沫耗材费",
    "surcharge": "加收",
    "fuel_or_tianjin_packing_fee": "燃油费/打包费",
    "courier_fee_total": "快递费合计",
}


def _dec(s):
    return Decimal(s) if s not in (None, "") else Decimal("0")


class ProfitService:
    def __init__(self, repo, op_log):
        self.repo = repo
        self.op_log = op_log

    # ==================== 口径配置 ====================
    def _policies(self):
        settings = self.repo.read("settings").get("settings") or {}
        return {
            "refund_deduction_status": settings.get(
                "refund_deduction_status", "unknown"),
            "platform_subsidy_included": settings.get(
                "platform_subsidy_included", "unknown"),
        }

    def update_config(self, refund_deduction_status=None,
                      platform_subsidy_included=None, operator="admin"):
        updates = {}
        if refund_deduction_status is not None:
            if refund_deduction_status not in REFUND_DEDUCTION_LABELS:
                raise ProfitServiceError(
                    "refund_deduction_status 仅支持 unknown / not_deducted / deducted")
            updates["refund_deduction_status"] = refund_deduction_status
        if platform_subsidy_included is not None:
            if platform_subsidy_included not in SUBSIDY_POLICY_LABELS:
                raise ProfitServiceError(
                    "platform_subsidy_included 仅支持 unknown / not_included / included")
            updates["platform_subsidy_included"] = platform_subsidy_included
        if not updates:
            raise ProfitServiceError("请至少提供一项待更新口径")

        def mutator(d):
            d["settings"].update(updates)
            d["updated_at"] = now_iso()
        self.repo.update("settings", mutator)
        self.op_log.append("profit.config", detail=updates, operator=operator)
        return self.get_config()

    def get_config(self):
        policies = self._policies()
        return {
            "policies": policies,
            "labels": {
                "refund_deduction_status": REFUND_DEDUCTION_LABELS[
                    policies["refund_deduction_status"]],
                "platform_subsidy_included": SUBSIDY_POLICY_LABELS[
                    policies["platform_subsidy_included"]],
            },
            "rules_version": RULES_VERSION,
            "declaration": CALIBER_DECLARATION,
        }

    def get_options(self):
        return {
            "profit_status": PROFIT_STATUS_LABELS,
            "cost_match_status": COST_MATCH_LABELS,
            "field_sources": FIELD_SOURCES,
            "sort_fields": SORT_FIELDS,
            "refund_deduction_status": REFUND_DEDUCTION_LABELS,
            "platform_subsidy_included": SUBSIDY_POLICY_LABELS,
            "status_reasons": STATUS_REASON_LABELS,
            "margin_blocked": MARGIN_BLOCKED_LABELS,
            "anomaly_types": ANOMALY_TYPE_LABELS,
            "rules_version": RULES_VERSION,
            "page_sizes": list(PAGE_SIZES),
            "declaration": CALIBER_DECLARATION,
            "config": self.get_config(),
        }

    # ==================== 执行核算 ====================
    def run_calculation(self, start=None, end=None, operator="admin"):
        policies = self._policies()
        orders = self.repo.read("orders")["records"]
        costs = self.repo.read("supplier_costs")["records"]
        recons = self.repo.read("reconciliation_results")["records"]

        orders_by_id = {}
        for o in orders:
            if o.get("order_id"):
                orders_by_id.setdefault(o["order_id"], []).append(o)

        scope = []
        for oid, wlines in orders_by_id.items():
            ot = next((l.get("order_time") for l in wlines
                       if l.get("order_time")), None)
            if not in_range(ot, start, end):
                continue
            scope.append(oid)

        recon_by_order = {}
        for r in recons:
            if r.get("record_type") == "order_match" and r.get("order_id"):
                recon_by_order[r["order_id"]] = r
        cost_by_id = {c.get("record_id"): c for c in costs}
        refund_by_order = {v["order_id"]: v for v in build_refund_views(orders)}

        calc_id = ("calc_" + datetime.now().strftime("%Y%m%d%H%M%S%f")
                   + "_" + uuid.uuid4().hex[:8])
        now = now_iso()

        records = []
        fresh = []
        wx_batches_all, sup_batches_all = set(), set()
        for oid in sorted(scope):
            rec, anoms = self._calc_order(
                oid, orders_by_id[oid], recon_by_order.get(oid),
                refund_by_order.get(oid), cost_by_id, policies, calc_id, now)
            records.append(rec)
            fresh.extend(anoms)
            wx_batches_all.update(rec["trace"]["batch_ids"]["wechat_orders"])
            sup_batches_all.update(rec["trace"]["batch_ids"]["supplier_costs"])

        anomalies, added, resolved = self._merge_anomalies(calc_id, fresh)
        run_record = {
            "record_id": calc_id,
            "record_type": "calculation_run",
            "rules_version": RULES_VERSION,
            "scope": {"start": start, "end": end},
            "policies": dict(policies),
            "statistics": self._stats_of(records),
            "source_batch_ids": {
                "wechat_orders": sorted(wx_batches_all),
                "supplier_costs": sorted(sup_batches_all),
            },
            "order_count": len(records),
            "created_at": now,
            "created_by": operator,
        }

        def mut_statements(d):
            d["records"].extend(records)
            d["records"].append(run_record)
            d["updated_at"] = now
        self.repo.update("profit_statements", mut_statements)

        def mut_anomalies(d):
            d["records"] = anomalies
            d["updated_at"] = now
        self.repo.update("anomalies", mut_anomalies)

        summary = dict(run_record["statistics"])
        summary.update({
            "calculation_id": calc_id,
            "rules_version": RULES_VERSION,
            "scope": run_record["scope"],
            "policies": dict(policies),
            "source_batch_ids": run_record["source_batch_ids"],
            "calculated_at": now,
            "anomalies_created": added,
            "anomalies_auto_resolved": resolved,
        })
        self.op_log.append("profit.calculate", detail={
            "calculation_id": calc_id, "start": start, "end": end,
            "rules_version": RULES_VERSION,
            "statistics": run_record["statistics"],
        }, operator=operator)
        return summary

    def _calc_order(self, oid, wlines, recon, refund_view, cost_by_id,
                    policies, calc_id, now):
        received, recv_conflict = order_level_amount(
            wlines, lambda l: (l.get("amounts") or {}).get(
                "actual_received_amount"))
        tech, _ = order_level_amount(
            wlines, lambda l: (l.get("amounts") or {}).get(
                "technical_service_fee"))
        wx_ids = [l.get("record_id") for l in wlines if l.get("record_id")]
        wx_batches = sorted({l.get("batch_id") for l in wlines
                             if l.get("batch_id")})

        rb = (refund_view or {}).get("refund_by_status") or {}
        refund_success = _dec(rb.get("success"))
        refund_pending = _dec(rb.get("pending"))
        has_refund = refund_success > 0 or refund_pending > 0

        sup_records = []
        if recon:
            sup_records = [cost_by_id[s]
                           for s in (recon.get("supplier_record_ids") or [])
                           if s in cost_by_id]
        cost = (recon or {}).get("cost") or {}
        supplier_total = None
        cost_state = "missing"
        if sup_records and cost.get("total_amount") is not None \
                and not cost.get("missing_amount_lines"):
            supplier_total = _dec(cost["total_amount"])
            cost_state = ("dedup_pending" if cost.get("order_level_dedup_pending")
                          else "ok")

        parts = {}
        for key in COST_PART_LABELS:
            parts[key] = money_str(sum(
                (_dec((r.get("cost_breakdown") or {}).get(key))
                 for r in sup_records), Decimal("0")))

        # ---- P1.1 四核心指标（防重复计算） ----
        # 商品销售收入：微信行级"商品实际价格(总共)"求和
        product_sales = sum((_dec((l.get("product") or {}).get(
            "actual_total_price")) for l in wlines), Decimal("0"))
        # 产品成本 / 快递费：供应商"成本" / "快递费合计"求和；
        # 任一匹配行分项缺失 → None（待核对，禁止默认 0 当作已确认）
        cost_parts_all_present = all(
            (r.get("cost_breakdown") or {}).get("product_supply_cost")
            is not None for r in sup_records) if sup_records else False
        courier_parts_all_present = all(
            (r.get("cost_breakdown") or {}).get("courier_fee_total")
            is not None for r in sup_records) if sup_records else False
        product_cost = (sum((_dec((r.get("cost_breakdown") or {}).get(
            "product_supply_cost")) for r in sup_records), Decimal("0"))
            if cost_parts_all_present else None)
        courier_fee = (sum((_dec((r.get("cost_breakdown") or {}).get(
            "courier_fee_total")) for r in sup_records), Decimal("0"))
            if courier_parts_all_present else None)
        # 产品毛利 = 商品销售收入 − 产品成本（独立指标，不参与贡献利润）
        product_gross = (None if product_cost is None
                         else product_sales - product_cost)
        # 成本匹配状态（缺分项 → 待核对）
        if not sup_records:
            cost_match_status = "unmatched"
        elif supplier_total is None:
            cost_match_status = "cost_missing"
        elif (recon or {}).get("match_status") == "suspected":
            cost_match_status = "suspected"
        elif not (cost_parts_all_present and courier_parts_all_present):
            cost_match_status = "parts_missing"
        else:
            cost_match_status = "matched"
        # 商品信息（明细表展示；多行拼接）
        titles = [t for t in dict.fromkeys(
            (l.get("product") or {}).get("title") for l in wlines
            if (l.get("product") or {}).get("title"))]
        quantity_total = sum((int((l.get("product") or {}).get("quantity")
                                  or 0) for l in wlines), 0)

        subsidy_total = sum((_dec(r.get("platform_subsidy"))
                             for r in sup_records), Decimal("0"))
        subsidy_credited = (subsidy_total
                            if policies["platform_subsidy_included"] == "not_included"
                            else Decimal("0"))

        if received is None:
            net = None
        elif policies["refund_deduction_status"] == "deducted":
            net = _dec(received)
        else:
            net = _dec(received) - refund_success

        reasons = []
        if received is None:
            status = "not_calculable"
            reasons.append("RECEIVED_AMOUNT_MISSING")
        elif cost_state == "missing":
            status = "not_calculable"
            reasons.append("UNMATCHED_NO_COST" if not sup_records
                           else "SUPPLIER_COST_MISSING")
        elif recon and recon.get("match_status") == "suspected":
            status = "pending_confirmation"
            reasons.append("SUSPECTED_MATCH_NOT_CONFIRMED")
        elif (policies["refund_deduction_status"] == "unknown"
              and has_refund):
            status = "pending_confirmation"
            reasons.append("REFUND_DEDUCTION_UNCONFIRMED")
        elif cost_state == "dedup_pending":
            status = "pending_confirmation"
            reasons.append("DUPLICATE_RISK_UNRESOLVED")
        elif recv_conflict:
            status = "needs_review"
            reasons.append("RECEIVED_AMOUNT_CONFLICT")
        else:
            status = "calculated"

        profit = None
        margin_pct = None
        margin_blocked = None
        if status == "not_calculable" or net is None or supplier_total is None:
            margin_blocked = ("NOT_CALCULABLE" if status == "not_calculable"
                              else "PENDING_CONFIRMATION")
        else:
            tech_d = _dec(tech)
            profit = net + subsidy_credited - supplier_total - tech_d
            if net > 0:
                margin_pct = (profit / net * Decimal("100")).quantize(
                    Decimal("0.01"))
            else:
                margin_blocked = "NET_REVENUE_NOT_POSITIVE"

        anoms = []
        if "REFUND_DEDUCTION_UNCONFIRMED" in reasons:
            anoms.append({
                "order_id": oid,
                "exception_type": "REFUND_DEDUCTION_UNCONFIRMED",
                "severity": "medium",
                "message": (f"订单 {oid} 存在退款（成功 {money_str(refund_success)} 元）"
                            f"但实收是否已扣退款口径未确认，利润显示待确认，"
                            f"请先在利润页确认退款口径"),
                "related_record_ids": wx_ids,
            })
        if profit is not None and profit < 0:
            anoms.append({
                "order_id": oid,
                "exception_type": "NEGATIVE_PROFIT",
                "severity": "low",
                "message": (f"订单 {oid} 经营贡献利润为负（{money_str(profit)} 元），"
                            f"请查看成本结构核实（RC-009）"),
                "related_record_ids": wx_ids,
            })

        m = money_str
        formula = {
            "net_revenue": self._fmt_net(received, refund_success,
                                          policies["refund_deduction_status"], net),
            "profit": (None if profit is None else
                       f"净收入 {m(net)} + 计入补贴 {m(subsidy_credited)}"
                       f" − 供应商综合总金额 {m(supplier_total)}"
                       f" − 技术服务费 {m(_dec(tech))}"
                       f" − 其他经营费用 {m(Decimal('0'))} = {m(profit)}"),
            "product_gross_profit": (
                "待核对：供应商成本分项缺失，无法确认产品成本，禁止按 0 计算"
                if product_gross is None else
                f"产品毛利 = 商品销售收入 {m(product_sales)}"
                f" − 产品成本 {m(product_cost)} = {m(product_gross)}"
                f"（独立经营指标，不参与经营贡献利润）"),
            "margin": (f"{m(profit)} ÷ 净收入 {m(net)} = {margin_pct}%"
                       if margin_pct is not None else None),
            "cost_deduction_basis": (
                f"仅扣除供应商「总金额」（综合成本）一次：{m(supplier_total)}"
                if supplier_total is not None else "供应商综合总金额缺失"),
            "parts_note": "综合总金额中已包含的成本/物流分项仅用于展示核对，不重复扣减",
        }
        margin_blocked_label = (MARGIN_BLOCKED_LABELS.get(margin_blocked)
                                if margin_blocked else None)

        rec = {
            "record_id": f"profit_{calc_id}_{oid}",
            "record_type": "order_profit",
            "calculation_id": calc_id,
            "rules_version": RULES_VERSION,
            "order_id": oid,
            "order_time": next((l.get("order_time") for l in wlines
                                if l.get("order_time")), None),
            "profit_status": status,
            "status_reasons": reasons,
            "amounts": {
                "actual_received_amount": received,
                "refund_success_amount": m(refund_success),
                "refund_pending_amount": m(refund_pending),
                "net_revenue": m(net) if net is not None else None,
                "product_sales_amount": m(product_sales),
                "product_cost": (m(product_cost)
                                 if product_cost is not None else None),
                "courier_fee": (m(courier_fee)
                                if courier_fee is not None else None),
                "composite_cost": (m(supplier_total)
                                   if supplier_total is not None else None),
                "product_gross_profit": (m(product_gross)
                                         if product_gross is not None else None),
                "platform_subsidy_total": m(subsidy_total),
                "platform_subsidy_credited": m(subsidy_credited),
                "platform_subsidy_excluded": m(subsidy_total - subsidy_credited),
                "supplier_total_amount": (m(supplier_total)
                                          if supplier_total is not None else None),
                "technical_service_fee": tech,
                "other_expenses": m(Decimal("0")),
                "operating_contribution_profit": m(profit) if profit is not None else None,
            },
            "product_info": {
                "titles": titles,
                "title": "；".join(titles) if titles else None,
                "quantity_total": quantity_total,
                "line_count": len(wlines),
            },
            "cost_match_status": cost_match_status,
            "margin": {
                "percent": f"{margin_pct}%" if margin_pct is not None else None,
                "numeric": str(margin_pct) if margin_pct is not None else None,
                "blocked_reason": margin_blocked,
                "blocked_reason_label": margin_blocked_label,
            },
            "formula": formula,
            "cost_parts_display_only": parts,
            "cost_granularity": cost.get("granularity"),
            "confirmed_granularity": cost.get("confirmed_granularity"),
            "policies": dict(policies),
            "trace": {
                "order_record_ids": wx_ids,
                "supplier_record_ids": [r.get("record_id")
                                        for r in sup_records],
                "refund_view_id": (refund_view or {}).get("record_id"),
                "reconciliation_id": (recon or {}).get("reconciliation_id"),
                "batch_ids": {
                    "wechat_orders": wx_batches,
                    "supplier_costs": sorted({r.get("batch_id")
                                              for r in sup_records
                                              if r.get("batch_id")}),
                },
            },
            "created_at": now,
        }
        return rec, anoms

    @staticmethod
    def _fmt_net(received, refund_success, policy, net):
        m = money_str
        if received is None:
            return "实际收款金额缺失，净收入不可计算"
        recv_d = _dec(received)
        if policy == "deducted":
            return (f"口径=已扣除：净收入 = 实收 {m(recv_d)}"
                    f"（不重复扣减退款）= {m(net)}")
        tag = "口径=未扣除" if policy == "not_deducted" else "口径=未确认（参考值）"
        return (f"{tag}：净收入 = 实收 {m(recv_d)}"
                f" − 成功退款 {m(refund_success)}（处理中退款不计入）= {m(net)}")

    # ==================== 经营总览 ====================
    def get_summary(self, start=None, end=None):
        orders = self.repo.read("orders")["records"]
        orders_by_id = {}
        for o in orders:
            if o.get("order_id"):
                orders_by_id.setdefault(o["order_id"], []).append(o)

        scope = []
        for oid, wlines in orders_by_id.items():
            ot = next((l.get("order_time") for l in wlines
                       if l.get("order_time")), None)
            if not in_range(ot, start, end):
                continue
            scope.append(oid)

        empty = {
            "has_orders": False, "declaration": CALIBER_DECLARATION,
            "policies": self._policies(),
        }
        if not scope:
            return empty

        received_total = Decimal("0")
        paid_total = Decimal("0")
        tech_total = Decimal("0")
        product_sales = Decimal("0")
        for oid in scope:
            wlines = orders_by_id[oid]
            received, _ = order_level_amount(
                wlines, lambda l: (l.get("amounts") or {}).get(
                    "actual_received_amount"))
            if received is not None:
                received_total += _dec(received)
            paid, _ = order_level_amount(
                wlines, lambda l: (l.get("amounts") or {}).get(
                    "actual_paid_amount"))
            if paid is not None:
                paid_total += _dec(paid)
            tech, _ = order_level_amount(
                wlines, lambda l: (l.get("amounts") or {}).get(
                    "technical_service_fee"))
            if tech is not None:
                tech_total += _dec(tech)
            product_sales += sum((_dec((l.get("product") or {}).get(
                "actual_total_price")) for l in wlines), Decimal("0"))

        refund_by_order = {v["order_id"]: v
                           for v in build_refund_views(orders)}
        refund_order_count = sum(1 for oid in scope if oid in refund_by_order)
        refund_success_total = sum((_dec(
            (refund_by_order[oid].get("refund_by_status") or {}).get("success"))
            for oid in scope if oid in refund_by_order), Decimal("0"))
        refund_pending_total = sum((_dec(
            (refund_by_order[oid].get("refund_by_status") or {}).get("pending"))
            for oid in scope if oid in refund_by_order), Decimal("0"))
        refund_rate = ((refund_success_total / received_total * Decimal("100"))
                       .quantize(Decimal("0.01"))
                       if received_total > 0 else None)

        recons = self.repo.read("reconciliation_results")["records"]
        recon_by_order = {}
        for r in recons:
            if r.get("record_type") == "order_match" and r.get("order_id"):
                recon_by_order[r["order_id"]] = r
        cost_total = Decimal("0")
        unmatched_order_count = 0
        for oid in scope:
            recon = recon_by_order.get(oid)
            if not recon:
                unmatched_order_count += 1
                continue
            if recon.get("match_status") == "unmatched_order":
                unmatched_order_count += 1
            cost = (recon.get("cost") or {}).get("total_amount")
            if cost is not None:
                cost_total += _dec(cost)

        open_anomaly_count = sum(
            1 for a in self.repo.read("anomalies")["records"]
            if a.get("status") == "pending_review")

        latest = self._latest_run()
        status_counts = {"calculated": 0, "pending_confirmation": 0,
                         "not_calculable": 0, "needs_review": 0}
        profit_total = Decimal("0")
        net_total = Decimal("0")
        margin = None
        margin_blocked = None
        calculated_count = 0
        latest_calc_id = latest.get("record_id") if latest else None
        if latest_calc_id:
            profit_records = self.repo.read("profit_statements")["records"]
            by_order = {r["order_id"]: r for r in profit_records
                        if r.get("record_type") == "order_profit"
                        and r.get("calculation_id") == latest_calc_id}
            for oid in scope:
                r = by_order.get(oid)
                if not r:
                    continue
                st = r.get("profit_status")
                if st in status_counts:
                    status_counts[st] += 1
                if st in CALCULATED_LIKE:
                    calculated_count += 1
                    profit_total += _dec(
                        (r.get("amounts") or {}).get(
                            "operating_contribution_profit"))
                    net_total += _dec((r.get("amounts") or {}).get("net_revenue"))
            if calculated_count > 0:
                if net_total > 0:
                    margin = (profit_total / net_total * Decimal("100")) \
                        .quantize(Decimal("0.01"))
                else:
                    margin_blocked = "NET_REVENUE_NOT_POSITIVE"

        m = money_str
        return {
            "has_orders": True,
            "order_count": len(scope),
            "refund_order_count": refund_order_count,
            "product_sales_amount": m(product_sales),
            "actual_paid_amount": m(paid_total),
            "actual_received_amount": m(received_total),
            "refund_success_amount": m(refund_success_total),
            "refund_pending_amount": m(refund_pending_total),
            "refund_rate": f"{refund_rate}%" if refund_rate is not None else None,
            "supplier_total_cost": m(cost_total),
            "technical_service_fee": m(tech_total),
            "operating_contribution_profit": m(profit_total),
            "net_revenue_total": m(net_total),
            "profit_margin": f"{margin}%" if margin is not None else None,
            "profit_margin_blocked_reason": margin_blocked,
            "profit_margin_blocked_label": MARGIN_BLOCKED_LABELS.get(
                margin_blocked) if margin_blocked else None,
            "profit_status_counts": status_counts,
            "calculable_order_count": calculated_count,
            "profit_coverage": (money_str(Decimal(calculated_count)
                                          / Decimal(len(scope)) * Decimal("100"))
                                + "%"),
            "unmatched_order_count": unmatched_order_count,
            "open_anomaly_count": open_anomaly_count,
            "has_calculation": bool(latest_calc_id),
            "latest_calculation": {
                "calculation_id": latest_calc_id,
                "rules_version": (latest or {}).get("rules_version"),
                "calculated_at": (latest or {}).get("created_at"),
                "policies": (latest or {}).get("policies"),
            } if latest else None,
            "declaration": CALIBER_DECLARATION,
            "policies": self._policies(),
        }

    # ==================== 查询与导出 ====================
    def _latest_run(self):
        runs = [r for r in self.repo.read("profit_statements")["records"]
                if r.get("record_type") == "calculation_run"]
        if not runs:
            return None
        return max(runs, key=lambda r: (r.get("created_at") or "",
                                        r.get("record_id") or ""))

    def list_results(self, start=None, end=None, order_id=None,
                     profit_status=None, calc_id=None, page=1, size=50):
        target = calc_id or (self._latest_run() or {}).get("record_id")
        records = [r for r in self.repo.read("profit_statements")["records"]
                   if r.get("record_type") == "order_profit"]
        if target:
            records = [r for r in records
                       if r.get("calculation_id") == target]
        rows = []
        for r in records:
            if order_id and r.get("order_id") != order_id:
                continue
            if profit_status and r.get("profit_status") != profit_status:
                continue
            if not in_range(r.get("order_time"), start, end):
                continue
            rows.append(r)
        rows.sort(key=lambda r: (r.get("order_time") or "",
                                 r.get("order_id") or ""))
        rows.reverse()
        total = len(rows)
        start_idx = (page - 1) * size
        return {
            "total": total, "page": page, "size": size,
            "pages": (total + size - 1) // size,
            "results": rows[start_idx:start_idx + size],
            "statistics": self._stats_of(rows),
            "calculation_id": target,
            "has_calculation": bool(target),
            "latest_run": self._latest_run(),
        }

    def export_excel(self, operator="admin", include_summary_row=True,
                     stats_sink=None, **filters):
        data = self.list_results(page=1, size=10 ** 9, **filters)
        rows = data["results"]
        statistics = dict(data["statistics"])
        if stats_sink is not None:
            stats_sink.append(dict(statistics))
        if not rows:
            raise ProfitServiceError("当前筛选无数据可导出")

        headers = ["订单号", "下单时间", "利润状态", "状态原因", "实际收款金额",
                   "成功退款", "处理中退款", "净收入", "补贴合计", "计入补贴",
                   "未计入补贴", "综合总金额(唯一扣减)", "技术服务费", "其他费用",
                   "经营贡献利润", "利润率", "退款口径", "补贴口径", "规则版本",
                   "核算时间", "微信批次", "供应商批次",
                   "产品成本", "快递费合计", "综合成本"]

        def row_of(r):
            a = r.get("amounts") or {}
            t = (r.get("trace") or {}).get("batch_ids") or {}
            return [
                r.get("order_id"), r.get("order_time"),
                PROFIT_STATUS_LABELS.get(r.get("profit_status"),
                                         r.get("profit_status")),
                "；".join(STATUS_REASON_LABELS.get(x, x)
                          for x in r.get("status_reasons") or []),
                _dec(a["actual_received_amount"]) if a.get(
                    "actual_received_amount") is not None else None,
                _dec(a["refund_success_amount"]) if a.get(
                    "refund_success_amount") is not None else None,
                _dec(a["refund_pending_amount"]) if a.get(
                    "refund_pending_amount") is not None else None,
                _dec(a["net_revenue"]) if a.get("net_revenue") is not None else None,
                _dec(a["platform_subsidy_total"]) if a.get(
                    "platform_subsidy_total") is not None else None,
                _dec(a["platform_subsidy_credited"]) if a.get(
                    "platform_subsidy_credited") is not None else None,
                _dec(a["platform_subsidy_excluded"]) if a.get(
                    "platform_subsidy_excluded") is not None else None,
                _dec(a["supplier_total_amount"]) if a.get(
                    "supplier_total_amount") is not None else None,
                _dec(a["technical_service_fee"]) if a.get(
                    "technical_service_fee") is not None else None,
                _dec(a["other_expenses"]) if a.get(
                    "other_expenses") is not None else None,
                _dec(a["operating_contribution_profit"]) if a.get(
                    "operating_contribution_profit") is not None else None,
                (r.get("margin") or {}).get("percent") or "—",
                REFUND_DEDUCTION_LABELS.get(
                    (r.get("policies") or {}).get("refund_deduction_status")),
                SUBSIDY_POLICY_LABELS.get(
                    (r.get("policies") or {}).get("platform_subsidy_included")),
                r.get("rules_version"), r.get("created_at"),
                "、".join(t.get("wechat_orders") or []),
                "、".join(t.get("supplier_costs") or []),
                _dec(a["product_cost"]) if a.get(
                    "product_cost") is not None else None,
                _dec(a["courier_fee"]) if a.get(
                    "courier_fee") is not None else None,
                _dec(a["composite_cost"]) if a.get(
                    "composite_cost") is not None else None,
            ]

        wb = Workbook()
        ws = wb.active
        ws.title = "订单利润明细"
        ws.append(headers)
        for r in rows:
            ws.append(row_of(r))
        if include_summary_row:
            notes = []
            pc_missing = statistics.get("product_cost_missing_count")
            cf_missing = statistics.get("courier_fee_missing_count")
            if isinstance(pc_missing, int) and pc_missing > 0:
                notes.append(f"产品成本待核对 {pc_missing} 单")
            if isinstance(cf_missing, int) and cf_missing > 0:
                notes.append(f"快递费合计待核对 {cf_missing} 单")
            summary_row = [None] * len(headers)
            summary_row[0] = "合计"
            summary_row[1] = f"共 {len(rows)} 单"
            summary_row[3] = "；".join(notes) if notes else None
            summary_row[22] = _dec(statistics.get("product_cost_total"))
            summary_row[23] = _dec(statistics.get("courier_fee_total"))
            summary_row[24] = _dec(statistics.get("composite_cost_total"))
            ws.append(summary_row)
        for col in (5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 23, 24, 25):
            for cell in ws.iter_cols(min_col=col, max_col=col, min_row=2):
                for c in cell:
                    c.number_format = "0.00"
        buf = io.BytesIO()
        wb.save(buf)

        filename = f"订单利润明细_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        record = {
            "record_id": ("exp_" + datetime.now().strftime("%Y%m%d%H%M%S") + "_"
                          + str(len(self.repo.read("export_records")["records"]) + 1)),
            "kind": "profit", "file_name": filename, "row_count": len(rows),
            "filters": {k: v for k, v in filters.items() if v is not None},
            "exported_at": now_iso(), "exported_by": operator,
        }

        def mutator(d):
            d["records"].append(record)
            d["updated_at"] = now_iso()
        self.repo.update("export_records", mutator)
        self.op_log.append("data.export", detail={
            "kind": "profit", "row_count": len(rows), "file_name": filename,
        }, operator=operator)
        return filename, buf.getvalue()

    # ==================== 内部 ====================
    @staticmethod
    def _stats_of(rows):
        stats = {"total": len(rows), "calculated": 0,
                 "pending_confirmation": 0, "not_calculable": 0,
                 "needs_review": 0, "profit_total": Decimal("0"),
                 "product_cost_total": Decimal("0"),
                 "courier_fee_total": Decimal("0"),
                 "composite_cost_total": Decimal("0"),
                 "product_cost_missing_count": 0,
                 "courier_fee_missing_count": 0,
                 "net_revenue_total": Decimal("0"),
                 "technical_service_fee_total": Decimal("0"),
                 "subsidy_credited_total": Decimal("0")}
        for r in rows:
            st = r.get("profit_status")
            if st in stats:
                stats[st] += 1
            a = r.get("amounts") or {}
            if st in CALCULATED_LIKE:
                stats["profit_total"] += _dec(
                    (r.get("amounts") or {}).get(
                        "operating_contribution_profit"))
                stats["net_revenue_total"] += _dec(a.get("net_revenue"))
                stats["technical_service_fee_total"] += _dec(
                    a.get("technical_service_fee"))
                stats["subsidy_credited_total"] += _dec(
                    a.get("platform_subsidy_credited"))
            unmatched = r.get("cost_match_status") == "unmatched"
            pc = a.get("product_cost")
            if pc is not None:
                stats["product_cost_total"] += _dec(pc)
            elif not unmatched:
                stats["product_cost_missing_count"] += 1
            cf = a.get("courier_fee")
            if cf is not None:
                stats["courier_fee_total"] += _dec(cf)
            elif not unmatched:
                stats["courier_fee_missing_count"] += 1
            cc = a.get("composite_cost")
            if cc is not None:
                stats["composite_cost_total"] += _dec(cc)
        # 成本核对：综合成本 − 产品成本 − 快递费 = 其他分项（泡沫耗材费等）。
        # 综合成本仍以供应商"总金额"为唯一扣减口径，此处仅用于页面核对展示。
        stats["composite_other_parts_total"] = money_str(
            stats["composite_cost_total"] - stats["product_cost_total"]
            - stats["courier_fee_total"])
        # 利润率 = 可计算订单贡献利润 ÷ 同口径净收入合计（净收入 > 0 才计算）
        margin = None
        margin_blocked = None
        calculated_like_count = sum(
            1 for r in rows if r.get("profit_status") in CALCULATED_LIKE)
        if calculated_like_count > 0:
            if stats["net_revenue_total"] > 0:
                margin = (stats["profit_total"]
                          / stats["net_revenue_total"] * Decimal("100")) \
                    .quantize(Decimal("0.01"))
            else:
                margin_blocked = "NET_REVENUE_NOT_POSITIVE"
        stats["profit_margin"] = f"{margin}%" if margin is not None else None
        stats["profit_margin_blocked_reason"] = margin_blocked
        stats["profit_margin_blocked_label"] = (
            MARGIN_BLOCKED_LABELS.get(margin_blocked)
            if margin_blocked else None)
        stats["profit_total"] = money_str(stats["profit_total"])
        stats["product_cost_total"] = money_str(
            stats["product_cost_total"])
        stats["courier_fee_total"] = money_str(
            stats["courier_fee_total"])
        stats["composite_cost_total"] = money_str(
            stats["composite_cost_total"])
        stats["net_revenue_total"] = money_str(stats["net_revenue_total"])
        stats["technical_service_fee_total"] = money_str(
            stats["technical_service_fee_total"])
        stats["subsidy_credited_total"] = money_str(
            stats["subsidy_credited_total"])
        return stats

    def _merge_anomalies(self, calc_id, fresh):
        """幂等合并利润异常：同键 pending 更新不重复新增；
        重算后未再出现的利润类异常自动关闭（留痕）。
        仅处理利润类异常，不影响对账类异常。"""
        anomalies = self.repo.read("anomalies")["records"]
        now = now_iso()
        index = {}
        for i, a in enumerate(anomalies):
            index.setdefault((a.get("order_id"), a.get("exception_type")),
                             []).append(i)

        added = 0
        for na in fresh:
            key = (na["order_id"], na["exception_type"])
            updated = False
            for i in index.get(key, []):
                if anomalies[i].get("status") == "pending_review":
                    anomalies[i]["message"] = na["message"]
                    anomalies[i]["severity"] = na["severity"]
                    anomalies[i]["reconciliation_id"] = calc_id
                    anomalies[i]["related_record_ids"] = na.get(
                        "related_record_ids", [])
                    anomalies[i]["updated_at"] = now
                    updated = True
                    break
            if not updated:
                anomalies.append({
                    "record_id": "exc_" + uuid.uuid4().hex,
                    "batch_id": None,
                    "reconciliation_id": calc_id,
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
        fresh_keys = {(na["order_id"], na["exception_type"]) for na in fresh}
        for a in anomalies:
            if (a.get("status") == "pending_review"
                    and a.get("exception_type") in PROFIT_ANOMALY_TYPES
                    and (a.get("order_id"), a.get("exception_type"))
                    not in fresh_keys):
                a["status"] = "resolved"
                a["resolution"] = {
                    "text": "重新核算后该异常未再出现，系统自动关闭",
                    "resolved_by": "system",
                    "resolved_at": now,
                }
                resolved += 1
        return anomalies, added, resolved

"""通用列表查询服务：供应商成本明细、异常记录（月份筛选统一规则接入）

- 供应商成本：默认按供应商订单创建时间（created_at）筛选，可切换发货时间
  （tracking.ship_time）；统计口径与表格一致；
- 异常记录：按异常发现时间（created_at，即记录生成时间）筛选，页面明示所用字段。
所有时间筛选为左闭右开区间 [start, end)，end 为排他上界。
"""
import io
from datetime import datetime
from decimal import Decimal

from openpyxl import Workbook

from .cost_validity import build_orders_lookup, is_effective_supplier_cost
from .data_schemas import now_iso
from .excel_export import append_row
from .normalizer import money_str
from .time_range import in_range

PAGE_SIZES = (20, 50, 100)
SUPPLIER_TIME_FIELDS = ("created_at", "ship_time")


def _dec(v):
    if v in (None, ""):
        return Decimal("0")
    return Decimal(v)


def _paginate(rows, page, size, stats, items_key):
    total = len(rows)
    page = max(1, int(page))
    size = max(1, int(size))
    start_idx = (page - 1) * size
    return {
        "total": total,
        "page": page,
        "size": size,
        "pages": (total + size - 1) // size,
        items_key: rows[start_idx:start_idx + size],
        "statistics": stats,
    }


class SupplierCostQueryService:
    """供应商成本明细列表（成本行级记录，即成本明细）"""

    def __init__(self, repo, op_log):
        self.repo = repo
        self.op_log = op_log

    def list_costs(self, time_field="created_at", start=None, end=None,
                   keyword=None, match_status=None, page=1, size=50):
        if time_field not in SUPPLIER_TIME_FIELDS:
            time_field = "created_at"
        records = self.repo.read("supplier_costs")["records"]
        orders_by_id = build_orders_lookup(self.repo.read("orders")["records"])
        rows = []
        for r in records:
            if not is_effective_supplier_cost(r, orders_by_id):
                continue
            if keyword:
                kw = keyword.lower()
                tracking = (r.get("tracking") or {}).get("tracking_number") or ""
                title = (r.get("product") or {}).get("title") or ""
                if not (kw in (r.get("main_order_id") or "").lower()
                        or kw in (r.get("supplier_order_id") or "").lower()
                        or kw in tracking.lower() or kw in title.lower()):
                    continue
            if match_status and r.get("match_status") != match_status:
                continue
            tv = (r.get("created_at") if time_field == "created_at"
                  else (r.get("tracking") or {}).get("ship_time"))
            if not in_range(tv, start, end):
                continue
            rows.append(r)
        rows.sort(key=lambda r: ((r.get("created_at") or ""), r.get("record_id") or ""))
        rows.reverse()

        total_amt = supply_t = courier_t = subsidy_t = Decimal("0")
        for r in rows:
            cb = r.get("cost_breakdown") or {}
            total_amt += _dec(cb.get("total_amount"))
            supply_t += _dec(cb.get("product_supply_cost"))
            courier_t += _dec(cb.get("courier_fee_total"))
            subsidy_t += _dec(r.get("platform_subsidy"))
        stats = {
            "cost_line_count": len(rows),
            "total_amount_sum": money_str(total_amt),
            "product_supply_cost_sum": money_str(supply_t),
            "courier_fee_sum": money_str(courier_t),
            "platform_subsidy_sum": money_str(subsidy_t),
            "time_field": time_field,
        }
        return _paginate(rows, page, size, stats, "costs")

    def export_excel(self, operator="admin", **filters):
        data = self.list_costs(
            time_field=filters.get("time_field") or "created_at",
            start=filters.get("start"), end=filters.get("end"),
            keyword=filters.get("keyword"),
            match_status=filters.get("match_status"),
            page=1, size=10 ** 9)
        rows = data["costs"]
        if not rows:
            raise ValueError("当前筛选无数据可导出")

        headers = ["主订单编号", "订单编号", "供应商订单创建时间", "发货时间",
                   "商品标题", "数量", "产品成本", "快递费", "泡沫耗材费", "加收",
                   "燃油/打包费", "总金额(综合成本)", "补贴", "快递单号", "成本粒度",
                   "导入批次"]
        wb = Workbook()
        ws = wb.active
        ws.title = "供应商成本明细"
        append_row(ws, headers)
        for r in rows:
            cb = r.get("cost_breakdown") or {}
            p = r.get("product") or {}
            t = r.get("tracking") or {}
            append_row(ws, [
                r.get("main_order_id"), r.get("supplier_order_id"),
                r.get("created_at"), t.get("ship_time"),
                p.get("title"), p.get("quantity"),
                _dec(cb.get("product_supply_cost")),
                _dec(cb.get("courier_fee")),
                _dec(cb.get("foam_material_fee")),
                _dec(cb.get("surcharge")),
                _dec(cb.get("fuel_or_tianjin_packing_fee")),
                _dec(cb.get("total_amount")),
                _dec(r.get("platform_subsidy")),
                t.get("tracking_number"), r.get("cost_granularity"),
                r.get("batch_id"),
            ])
        for col in (7, 8, 9, 10, 11, 12, 13):
            for cell in ws.iter_cols(min_col=col, max_col=col, min_row=2):
                for c in cell:
                    c.number_format = "0.00"
        buf = io.BytesIO()
        wb.save(buf)
        return self._log_export("供应商成本明细", rows, filters, buf, operator)

    def _log_export(self, name, rows, filters, buf, operator):
        filename = f"{name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        record = {
            "record_id": "exp_" + datetime.now().strftime("%Y%m%d%H%M%S") + "_"
                         + str(len(self.repo.read("export_records")["records"]) + 1),
            "kind": "supplier_costs", "file_name": filename,
            "row_count": len(rows),
            "filters": {k: v for k, v in filters.items() if v is not None},
            "exported_at": now_iso(), "exported_by": operator,
        }

        def mutator(d):
            d["records"].append(record)
            d["updated_at"] = now_iso()
            return d
        self.repo.update("export_records", mutator)
        self.op_log.append("data.export", detail={
            "kind": "supplier_costs", "row_count": len(rows),
            "file_name": filename,
        }, operator=operator)
        return filename, buf.getvalue()


ANOMALY_STATUS_LABELS = {
    "pending_review": "待处理",
    "resolved": "已解决",
}
ANOMALY_SEVERITY_LABELS = {"high": "高", "medium": "中", "low": "低"}


def _anomaly_handler(r):
    res = r.get("resolution")
    if isinstance(res, dict):
        return res.get("resolved_by") or res.get("operator") or None
    return None


def _enrich_anomaly(r):
    """列表/导出统一结构化字段（不修改磁盘记录）。"""
    handler = _anomaly_handler(r)
    et = r.get("exception_type")
    st = r.get("status")
    out = dict(r)
    out["exception_type_label"] = et or "—"
    out["status_label"] = ANOMALY_STATUS_LABELS.get(st, st or "—")
    out["severity_label"] = ANOMALY_SEVERITY_LABELS.get(
        r.get("severity"), r.get("severity") or "—")
    out["handler"] = handler
    out["handler_label"] = handler or "—"
    return out


class AnomalyQueryService:
    """异常记录列表：时间字段为异常发现时间（created_at）"""

    def __init__(self, repo, op_log):
        self.repo = repo
        self.op_log = op_log

    def list_anomalies(self, start=None, end=None, status=None,
                       exception_type=None, severity=None, keyword=None,
                       order_id=None, page=1, size=50):
        records = self.repo.read("anomalies")["records"]
        rows = []
        for r in records:
            if status and r.get("status") != status:
                continue
            if exception_type and r.get("exception_type") != exception_type:
                continue
            if severity and r.get("severity") != severity:
                continue
            if order_id and r.get("order_id") != order_id:
                continue
            if keyword:
                kw = keyword.lower()
                if not (kw in (r.get("message") or "").lower()
                        or kw in (r.get("order_id") or "").lower()
                        or kw in (r.get("exception_type") or "").lower()):
                    continue
            if not in_range(r.get("created_at"), start, end):
                continue
            rows.append(_enrich_anomaly(r))
        rows.sort(key=lambda r: ((r.get("created_at") or ""), r.get("record_id") or ""))
        rows.reverse()

        pending = high = 0
        for r in rows:
            if r.get("status") == "pending_review":
                pending += 1
            if r.get("severity") == "high":
                high += 1
        stats = {
            "anomaly_count": len(rows),
            "pending_count": pending,
            "resolved_count": len(rows) - pending,
            "high_severity_count": high,
            "time_field": "created_at",
        }
        return _paginate(rows, page, size, stats, "anomalies")

    def get_options(self):
        types = {}
        for r in self.repo.read("anomalies")["records"]:
            t = r.get("exception_type")
            if t:
                types[t] = t
        return {
            "status": ANOMALY_STATUS_LABELS,
            "severity": ANOMALY_SEVERITY_LABELS,
            "exception_types": types,
            "time_field_label": "异常发现时间（记录生成时间）",
        }

    def export_excel(self, operator="admin", **filters):
        data = self.list_anomalies(page=1, size=10 ** 9, **filters)
        rows = data["anomalies"]
        if not rows:
            raise ValueError("当前筛选无数据可导出")

        headers = ["异常发现时间", "异常类型", "级别", "状态", "关联订单号",
                   "处理人", "异常信息", "处理结果", "导入批次", "对账批次"]
        wb = Workbook()
        ws = wb.active
        ws.title = "异常记录"
        append_row(ws, headers)
        for r in rows:
            append_row(ws, [
                r.get("created_at"), r.get("exception_type"),
                ANOMALY_SEVERITY_LABELS.get(r.get("severity"), r.get("severity")),
                ANOMALY_STATUS_LABELS.get(r.get("status"), r.get("status")),
                r.get("order_id") or "", r.get("handler_label") or "",
                r.get("message") or "",
                (r.get("resolution") or {}).get("note") or ""
                if isinstance(r.get("resolution"), dict) else (r.get("resolution") or ""),
                r.get("batch_id") or "", r.get("reconciliation_id") or "",
            ])
        buf = io.BytesIO()
        wb.save(buf)

        filename = f"异常记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        record = {
            "record_id": "exp_" + datetime.now().strftime("%Y%m%d%H%M%S") + "_"
                         + str(len(self.repo.read("export_records")["records"]) + 1),
            "kind": "anomalies", "file_name": filename, "row_count": len(rows),
            "filters": {k: v for k, v in filters.items() if v is not None},
            "exported_at": now_iso(), "exported_by": operator,
        }

        def mutator(d):
            d["records"].append(record)
            d["updated_at"] = now_iso()
            return d
        self.repo.update("export_records", mutator)
        self.op_log.append("data.export", detail={
            "kind": "anomalies", "row_count": len(rows), "file_name": filename,
        }, operator=operator)
        return filename, buf.getvalue()

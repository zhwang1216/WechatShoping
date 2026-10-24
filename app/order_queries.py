"""订单查询服务：普通订单列表/详情、退款订单列表/统计、逐笔退款明细、Excel 导出

普通订单页（spec 5.2）：默认只展示退款金额为 0 且无退款相关状态的订单；
退款订单页（spec 5.3）：按订单去重一行，分状态统计金额，支持筛选/分页/排序/导出。

月份筛选（统一规则）：
- 所有时间筛选为左闭右开区间 [start, end)，end 为排他上界；
- 普通订单按订单下单时间归属月份；
- 退款订单默认按退款申请时间（refund_apply_time）筛选，可切换退款成功时间
  或下单时间；统计卡片主口径与表格一致，成功退款金额另附"按退款成功时间"
  统计口径，跨月退款不得因原订单下单月份不同而漏记。
"""
import io
from datetime import datetime
from decimal import Decimal

from openpyxl import Workbook

from .data_schemas import now_iso
from .normalizer import money_str
from .refund_view import (SUCCESS_STATUSES, build_refund_views,
                          order_level_amount)
from .time_range import in_range

PAGE_SIZES = (20, 50, 100)
NORMAL_SORT_FIELDS = ("order_time", "order_id", "actual_received_amount")
REFUND_SORT_FIELDS = ("order_time", "order_id", "refund_amount", "actual_received_amount")
REFUND_TIME_FIELDS = ("apply_time", "success_time", "order_time")
REFUND_DETAIL_TIME_FIELDS = ("apply_time", "success_time")

ORDER_STATUS_LABELS = {
    "completed": "已完成", "shipped": "已发货", "pending_shipment": "待发货",
    "paid": "已支付", "pending_payment": "待付款", "closed": "已关闭",
    "refund_processing": "退款处理中", "partial_refunded": "部分退款",
    "full_refunded": "全额退款", "unknown": "未知状态",
}
REFUND_STATUS_LABELS = {
    "none": "无退款", "pending": "退款处理中", "partial": "部分退款",
    "full": "全额退款", "closed": "退款关闭", "unknown": "未知/待核对",
}


def _dec(s):
    if s in (None, ""):
        return Decimal("0")
    return Decimal(s)


def _has_refund(line):
    return (line.get("refund") or {}).get("refund_status") not in (None, "none")


class OrderQueryService:
    def __init__(self, repo, op_log):
        self.repo = repo
        self.op_log = op_log

    # ==================== 普通订单 ====================
    def list_normal_orders(self, order_id=None, status=None, start=None, end=None,
                           keyword=None,
                           order_by="order_time", order_dir="desc", page=1, size=50):
        rows, stats = self._load_normal_orders(order_id, status, start, end,
                                               keyword=keyword)
        rows = self._sort(rows, order_by, order_dir, NORMAL_SORT_FIELDS)
        return self._paginate(rows, page, size, stats, "orders")

    def _load_normal_orders(self, order_id=None, status=None, start=None, end=None,
                            keyword=None):
        orders = self.repo.read("orders")["records"]
        by_order = {}
        for o in orders:
            if o.get("order_id"):
                by_order.setdefault(o["order_id"], []).append(o)

        rows = []
        paid_t = recv_t = tech_t = prod_t = Decimal("0")
        qty_t = 0
        for oid, lines in by_order.items():
            if any(_has_refund(l) for l in lines):
                continue
            first = lines[0]
            if order_id and oid != order_id:
                continue
            if keyword and keyword.lower() not in oid.lower():
                continue
            if status and first.get("order_status") != status:
                continue
            ot = first.get("order_time")
            if not in_range(ot, start, end):
                continue

            received, conflict = order_level_amount(
                lines, lambda l: (l.get("amounts") or {}).get("actual_received_amount"))
            tech, _ = order_level_amount(
                lines, lambda l: (l.get("amounts") or {}).get("technical_service_fee"))
            paid, _ = order_level_amount(
                lines, lambda l: (l.get("amounts") or {}).get("actual_paid_amount"))
            product_total = sum((_dec((l.get("product") or {}).get("actual_total_price"))
                                 for l in lines), Decimal("0"))
            qty = sum((l.get("product") or {}).get("quantity") or 0 for l in lines)
            merged = {}
            for l in lines:
                p = l.get("product") or {}
                name = p.get("title") or p.get("sku_code") or "未命名商品"
                attr = p.get("sku_attribute") or ""
                merged[(name, attr)] = merged.get((name, attr), 0) + (p.get("quantity") or 0)
            summary = "; ".join(f"{n}{'(' + a + ')' if a else ''}×{q}"
                                for (n, a), q in merged.items())

            rows.append({
                "order_id": oid,
                "order_time": ot,
                "order_status": first.get("order_status"),
                "source_status": first.get("source_status"),
                "amounts": {
                    "actual_paid_amount": paid,
                    "actual_received_amount": received,
                    "technical_service_fee": tech,
                    "product_total": money_str(product_total),
                },
                "amount_conflict": conflict,
                "has_refund": False,
                "product_summary": summary,
                "product_line_count": len(lines),
                "product_quantity_total": qty,
                "batch_ids": sorted({l.get("batch_id") for l in lines if l.get("batch_id")}),
                "updated_at": max((l.get("updated_at") or "" for l in lines), default=None),
            })
            paid_t += _dec(paid)
            recv_t += _dec(received)
            tech_t += _dec(tech)
            prod_t += product_total
            qty_t += qty

        stats = {
            "order_count": len(rows),
            "actual_paid_total": money_str(paid_t),
            "actual_received_total": money_str(recv_t),
            "technical_service_fee_total": money_str(tech_t),
            "product_total_sum": money_str(prod_t),
            "quantity_total": qty_t,
        }
        return rows, stats

    # ==================== 退款订单 ====================
    def list_refund_orders(self, order_id=None, status=None, start=None, end=None,
                           min_amount=None, max_amount=None, keyword=None,
                           time_field="apply_time",
                           order_by="order_time", order_dir="desc", page=1, size=50):
        rows, stats = self._load_refund_orders(order_id, status, start, end,
                                                min_amount, max_amount,
                                                keyword=keyword,
                                                time_field=time_field)
        rows = self._sort(rows, order_by, order_dir, REFUND_SORT_FIELDS)
        return self._paginate(rows, page, size, stats, "orders")

    def _load_refund_orders(self, order_id=None, status=None, start=None, end=None,
                            min_amount=None, max_amount=None, keyword=None,
                            time_field="apply_time"):
        if time_field not in REFUND_TIME_FIELDS:
            time_field = "apply_time"
        refund_details = self.repo.read("refund_details")["records"]
        views = build_refund_views(self.repo.read("orders")["records"],
                                   refund_details=refund_details)
        rows = []
        for v in views:
            if order_id and v["order_id"] != order_id:
                continue
            if keyword and keyword.lower() not in v["order_id"].lower():
                continue
            if status and v["refund"]["refund_status"] != status:
                continue
            if time_field == "apply_time":
                tv = v["refund"].get("refund_apply_time")
            elif time_field == "success_time":
                tv = v["refund"].get("refund_success_time")
            else:
                tv = v.get("order_time")
            if not in_range(tv, start, end):
                continue
            amt = _dec(v["refund"]["refund_amount"])
            if min_amount is not None and amt < min_amount:
                continue
            if max_amount is not None and amt > max_amount:
                continue
            rows.append(v)

        success_t = pending_t = closed_t = total_t = Decimal("0")
        partial_n = full_n = review_n = fallback_n = 0
        for v in rows:
            bs = v["refund_by_status"]
            success_t += _dec(bs.get("success"))
            pending_t += _dec(bs.get("pending"))
            closed_t += _dec(bs.get("closed"))
            total_t += _dec(v["refund"]["refund_amount"])
            st = v["refund"]["refund_status"]
            if st == "partial":
                partial_n += 1
            elif st == "full":
                full_n += 1
            if v.get("pending_review"):
                review_n += 1
            if v["refund"].get("refund_apply_time_source") == "order_time_fallback":
                fallback_n += 1

        stats = {
            "refund_order_count": len(rows),
            "total_refund_amount": money_str(total_t),
            "success_refund_amount": money_str(success_t),
            "pending_refund_amount": money_str(pending_t),
            "closed_refund_amount": money_str(closed_t),
            "partial_refund_order_count": partial_n,
            "full_refund_order_count": full_n,
            "pending_review_count": review_n,
            "apply_time_fallback_count": fallback_n,
            "time_field": time_field,
            # 成功退款金额的独立统计口径：按退款成功时间（refund_time 且状态为成功）
            "success_refund_amount_by_success_time": money_str(
                self._success_amount_by_success_time(refund_details, start, end)),
        }
        return rows, stats

    @staticmethod
    def _success_amount_by_success_time(refund_details, start, end):
        """按退款成功时间统计成功退款金额（跨月退款按成功发生月份归属）"""
        total = Decimal("0")
        for d in refund_details:
            if d.get("refund_status") not in SUCCESS_STATUSES:
                continue
            if not in_range(d.get("refund_time"), start, end):
                continue
            total += _dec(d.get("refund_amount"))
        return total

    # ==================== 退款明细（逐笔） ====================
    def list_refund_details(self, time_field="apply_time", start=None, end=None,
                            order_id=None, refund_status=None, keyword=None,
                            page=1, size=50):
        if time_field not in REFUND_DETAIL_TIME_FIELDS:
            time_field = "apply_time"
        records = self.repo.read("refund_details")["records"]
        rows = []
        for d in records:
            if order_id and d.get("order_id") != order_id:
                continue
            if keyword and keyword.lower() not in (d.get("order_id") or "").lower():
                continue
            if refund_status and d.get("refund_status") != refund_status:
                continue
            rt = d.get("refund_time")
            if time_field == "apply_time":
                if not in_range(rt, start, end):
                    continue
            else:
                if d.get("refund_status") not in SUCCESS_STATUSES:
                    continue
                if not in_range(rt, start, end):
                    continue
            rows.append(d)
        rows.sort(key=lambda d: ((d.get("refund_time") or d.get("created_at") or ""),
                                 d.get("record_id") or ""))
        rows.reverse()

        total_amt = success_t = pending_t = closed_t = Decimal("0")
        for d in rows:
            amt = _dec(d.get("refund_amount"))
            total_amt += amt
            st = d.get("refund_status")
            if st in SUCCESS_STATUSES:
                success_t += amt
            elif st == "pending":
                pending_t += amt
            elif st == "closed":
                closed_t += amt

        stats = {
            "refund_detail_count": len(rows),
            "total_refund_amount": money_str(total_amt),
            "success_refund_amount": money_str(success_t),
            "pending_refund_amount": money_str(pending_t),
            "closed_refund_amount": money_str(closed_t),
            "time_field": time_field,
        }
        return self._paginate(rows, page, size, stats, "details")

    # ==================== 订单详情 ====================
    def get_order(self, order_id):
        orders = self.repo.read("orders")["records"]
        lines = [o for o in orders if o.get("order_id") == order_id]
        if not lines:
            raise KeyError(order_id)

        refunds = self._refunds_of(order_id)
        batches = self.repo.read("import_records")["records"]
        batch_ids = sorted({l.get("batch_id") for l in lines if l.get("batch_id")})
        batch_info = [
            {"batch_id": b.get("batch_id"), "file_name": b.get("file_name"),
             "imported_at": b.get("imported_at"), "status": b.get("status"),
             "statistics": b.get("statistics")}
            for b in batches if b.get("batch_id") in batch_ids
        ]

        received, conflict = order_level_amount(
            lines, lambda l: (l.get("amounts") or {}).get("actual_received_amount"))
        tech, _ = order_level_amount(
            lines, lambda l: (l.get("amounts") or {}).get("technical_service_fee"))
        paid, _ = order_level_amount(
            lines, lambda l: (l.get("amounts") or {}).get("actual_paid_amount"))
        product_total = sum((_dec((l.get("product") or {}).get("actual_total_price"))
                             for l in lines), Decimal("0"))

        has_refund = any(_has_refund(l) for l in lines)
        refund_view = next((v for v in build_refund_views(lines)), None) if has_refund else None

        first = lines[0]
        return {
            "order_id": order_id,
            "order_time": first.get("order_time"),
            "order_status": first.get("order_status"),
            "source_status": first.get("source_status"),
            "has_refund": has_refund,
            "amounts": {
                "actual_paid_amount": paid,
                "actual_received_amount": received,
                "technical_service_fee": tech,
                "product_total": money_str(product_total),
            },
            "amount_conflict": conflict,
            "product_line_count": len(lines),
            "lines": lines,
            "refunds": refunds,
            "refund_view": refund_view,
            "batches": batch_info,
            "batch_ids": batch_ids,
        }

    def get_order_refunds(self, order_id):
        """逐笔退款明细：订单不存在抛 KeyError；无退款返回空数组"""
        orders = self.repo.read("orders")["records"]
        if not any(o.get("order_id") == order_id for o in orders):
            raise KeyError(order_id)
        return self._refunds_of(order_id)

    def _refunds_of(self, order_id):
        refunds = [r for r in self.repo.read("refund_details")["records"]
                   if r.get("order_id") == order_id]
        refunds.sort(key=lambda r: ((r.get("created_at") or ""),
                                    (r.get("row_number") or 0)))
        return refunds

    # ==================== 导出 ====================
    def export_excel(self, kind, operator="admin", **filters):
        """导出与筛选完全一致的 Excel（含月份/时间区间/关键词）；无数据抛 ValueError"""
        if kind == "normal":
            rows, _ = self._load_normal_orders(
                filters.get("order_id"), filters.get("status"),
                filters.get("start"), filters.get("end"),
                keyword=filters.get("keyword"))
            filename = f"普通订单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            headers = ["订单号", "下单时间", "订单状态", "订单实际支付金额",
                       "订单实际收款金额", "技术服务费", "商品总价", "商品数量",
                       "商品明细", "金额冲突", "导入批次", "更新时间"]
            money_cols = (4, 5, 6, 7)
            def row_of(r):
                return [r["order_id"], r["order_time"],
                        ORDER_STATUS_LABELS.get(r["order_status"], r["order_status"]),
                        _dec(r["amounts"]["actual_paid_amount"]),
                        _dec(r["amounts"]["actual_received_amount"]),
                        _dec(r["amounts"]["technical_service_fee"]),
                        _dec(r["amounts"]["product_total"]),
                        r["product_quantity_total"], r["product_summary"],
                        "是" if r["amount_conflict"] else "否",
                        ",".join(r["batch_ids"]), r["updated_at"]]
            sheet_title = "普通订单"
        else:
            rows, _ = self._load_refund_orders(
                filters.get("order_id"), filters.get("status"),
                filters.get("start"), filters.get("end"),
                filters.get("min_amount"), filters.get("max_amount"),
                keyword=filters.get("keyword"),
                time_field=filters.get("time_field") or "apply_time")
            filename = f"退款订单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            headers = ["订单号", "下单时间", "退款申请时间", "退款成功时间", "订单状态",
                       "退款状态", "退款总额",
                       "退款成功金额", "退款处理中金额", "退款关闭金额", "退款笔数",
                       "退款粒度", "状态来源", "待核对", "订单实际收款金额",
                       "商品明细", "导入批次", "更新时间"]
            money_cols = (7, 8, 9, 10, 11, 15)
            def row_of(r):
                rf = r["refund"]
                bs = r["refund_by_status"]
                return [r["order_id"], r["order_time"],
                        rf.get("refund_apply_time"), rf.get("refund_success_time"),
                        ORDER_STATUS_LABELS.get(r["order_status"], r["order_status"]),
                        REFUND_STATUS_LABELS.get(rf["refund_status"], rf["refund_status"]),
                        _dec(rf["refund_amount"]), _dec(bs.get("success")),
                        _dec(bs.get("pending")), _dec(bs.get("closed")),
                        rf["refund_count"],
                        "订单级待核对" if rf["refund_amount_granularity"] == "unknown" else "商品行级",
                        "推导" if rf["status_source"] == "derived" else "平台原始",
                        "是" if r.get("pending_review") else "否",
                        _dec(r["amounts"]["actual_received_amount"]),
                        r["product_summary"], ",".join(r["source"]["batch_ids"]),
                        r["updated_at"]]
            sheet_title = "退款订单"

        if not rows:
            raise ValueError("当前筛选无数据可导出")

        wb = Workbook()
        ws = wb.active
        ws.title = sheet_title
        ws.append(headers)
        for r in rows:
            ws.append(row_of(r))
        for col in money_cols:
            for cell in ws.iter_cols(min_col=col, max_col=col, min_row=2):
                for c in cell:
                    c.number_format = "0.00"
        buf = io.BytesIO()
        wb.save(buf)

        record = {
            "record_id": "exp_" + datetime.now().strftime("%Y%m%d%H%M%S") + "_"
                         + str(len(self.repo.read("export_records")["records"]) + 1),
            "kind": kind, "file_name": filename, "row_count": len(rows),
            "filters": {k: v for k, v in filters.items() if v is not None},
            "exported_at": now_iso(), "exported_by": operator,
        }

        def mutator(d):
            d["records"].append(record)
            d["updated_at"] = now_iso()
            return d
        self.repo.update("export_records", mutator)
        self.op_log.append("data.export", detail={
            "kind": kind, "row_count": len(rows), "file_name": filename,
        }, operator=operator)
        return filename, buf.getvalue()

    def export_refund_details_excel(self, operator="admin", **filters):
        """导出逐笔退款明细（与列表筛选完全一致）；无数据抛 ValueError"""
        data = self.list_refund_details(
            time_field=filters.get("time_field") or "apply_time",
            start=filters.get("start"), end=filters.get("end"),
            order_id=filters.get("order_id"),
            refund_status=filters.get("refund_status"),
            keyword=filters.get("keyword"), page=1, size=10 ** 9)
        rows = data["details"]
        if not rows:
            raise ValueError("当前筛选无数据可导出")

        headers = ["订单号", "退款金额", "退款状态", "状态来源", "退款时间",
                   "退款粒度", "原售后文本", "源工作表/行号", "导入批次", "记录时间"]
        wb = Workbook()
        ws = wb.active
        ws.title = "退款明细"
        ws.append(headers)
        for r in rows:
            ws.append([
                r.get("order_id"), _dec(r.get("refund_amount")),
                REFUND_STATUS_LABELS.get(r.get("refund_status"), r.get("refund_status")),
                "推导" if r.get("status_source") == "derived" else "平台原始",
                r.get("refund_time"),
                "订单级待核对" if r.get("refund_amount_granularity") == "unknown" else "商品行级",
                r.get("raw_after_sale_text") or "",
                f"{r.get('sheet_name') or ''} 第{r.get('row_number')}行",
                r.get("batch_id") or "", r.get("created_at") or "",
            ])
        for cell in ws.iter_cols(min_col=2, max_col=2, min_row=2):
            for c in cell:
                c.number_format = "0.00"
        buf = io.BytesIO()
        wb.save(buf)

        filename = f"退款明细_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        record = {
            "record_id": "exp_" + datetime.now().strftime("%Y%m%d%H%M%S") + "_"
                         + str(len(self.repo.read("export_records")["records"]) + 1),
            "kind": "refund_details", "file_name": filename, "row_count": len(rows),
            "filters": {k: v for k, v in filters.items() if v is not None},
            "exported_at": now_iso(), "exported_by": operator,
        }

        def mutator(d):
            d["records"].append(record)
            d["updated_at"] = now_iso()
            return d
        self.repo.update("export_records", mutator)
        self.op_log.append("data.export", detail={
            "kind": "refund_details", "row_count": len(rows),
            "file_name": filename,
        }, operator=operator)
        return filename, buf.getvalue()

    # ==================== 内部工具 ====================
    @staticmethod
    def _sort(rows, order_by, order_dir, allowed_fields):
        if order_by not in allowed_fields:
            order_by = allowed_fields[0]
        key_map = {
            "order_time": lambda r: r.get("order_time") or "",
            "order_id": lambda r: r["order_id"],
            "actual_received_amount": lambda r: _dec(
                r.get("amounts", {}).get("actual_received_amount")),
            "refund_amount": lambda r: _dec(r.get("refund", {}).get("refund_amount")),
        }
        key = key_map[order_by]
        rows.sort(key=lambda r: r["order_id"])
        rows.sort(key=key, reverse=(order_dir == "desc"))
        return rows

    @staticmethod
    def _paginate(rows, page, size, stats, items_key):
        total = len(rows)
        start_idx = (page - 1) * size
        return {
            "total": total,
            "page": page,
            "size": size,
            "pages": (total + size - 1) // size,
            items_key: rows[start_idx:start_idx + size],
            "statistics": stats,
        }

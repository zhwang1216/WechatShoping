"""月份筛选统一规则测试（spec：全站月份筛选）

覆盖：
1. 月末订单归属（左闭右开，23:59:59.999 毫秒精度不漏）
2. 跨月退款不漏记（申请时间缺失回退下单时间；成功金额按成功时间独立口径）
3. 空月份返回空结果不报错
4. 统计卡片与表格数据一致（数量/金额）
5. 翻页保留筛选（total 恒定、页间不重叠）
6. 导出遵守当前筛选
7. 全部 / 自定义范围 / 参数优先级（from/to > month）
8. 时区与格式一致性（ISO T 分隔与空格分隔同口径；本地零点边界）
9. 各页面接口（订单/退款订单/退款明细/供应商成本/对账/利润/异常/导入历史/日志）
"""
import io
import json
import unittest
from datetime import datetime
from decimal import Decimal

from openpyxl import load_workbook

from tests.http_client import TestEnv


def _order(order_id, order_time, received, refund_status="none",
           refund_amount="0.00", status="completed"):
    return {
        "record_id": f"ord_{order_id}",
        "order_id": order_id,
        "order_time": order_time,
        "order_status": status,
        "source_status": status,
        "amounts": {
            "actual_paid_amount": received,
            "actual_received_amount": received,
            "technical_service_fee": "1.00",
        },
        "product": {
            "title": f"商品-{order_id}",
            "sku_attribute": "",
            "quantity": 1,
            "actual_total_price": received,
        },
        "refund": {
            "refund_status": refund_status,
            "refund_amount": refund_amount,
            "granularity": "line",
            "status_source": "platform",
        },
        "batch_id": "bat_test_1",
        "updated_at": order_time,
    }


def _refund_detail(record_id, order_id, amount, status, refund_time):
    return {
        "record_id": record_id,
        "order_id": order_id,
        "refund_amount": amount,
        "refund_status": status,
        "refund_time": refund_time,
        "refund_amount_granularity": "line",
        "status_source": "platform",
        "sheet_name": "退款明细",
        "row_number": 1,
        "batch_id": "bat_test_1",
        "created_at": "2026-06-01T00:00:00",
    }


class TestMonthFilter(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member_store = cls.env.member_store

        normal_orders = [
            _order("ORD-0501", "2026-05-01 10:00:00", "100.00"),
            _order("ORD-0531A", "2026-05-31T23:59:59", "50.50"),
            _order("ORD-0531B", "2026-05-31 23:59:59.999", "30.00"),
            _order("ORD-0601", "2026-06-01 00:00:00", "200.00"),
            _order("ORD-0615", "2026-06-15 12:30:00", "80.00"),
        ]
        # 7 月批量订单（分页测试：25 条，page_size=20 时分 2 页）
        for i in range(1, 26):
            normal_orders.append(_order(
                f"ORD-07{i:02d}", f"2026-07-01 08:00:{i % 60:02d}", "10.00"))
        refund_orders = [
            # 9 月下单、10 月退款（跨月退款，申请时间完整）
            _order("REF-SEP", "2026-09-30 20:00:00", "66.60",
                   refund_status="partial", refund_amount="66.60"),
            # 9 月下单、退款时间缺失（应回退下单时间，跨月退款保护）
            _order("REF-NULL", "2026-09-20 10:00:00", "20.00",
                   refund_status="partial", refund_amount="20.00"),
            # 10 月下单、10 月退款成功
            _order("REF-OCT", "2026-10-02 08:00:00", "10.00",
                   refund_status="full", refund_amount="10.00"),
        ]
        refund_details = [
            _refund_detail("rd_1", "REF-SEP", "66.60", "full",
                           "2026-10-05 09:00:00"),
            _refund_detail("rd_2", "REF-NULL", "20.00", "pending", None),
            _refund_detail("rd_3", "REF-OCT", "10.00", "full",
                           "2026-10-03 08:00:00"),
        ]
        supplier_costs = [
            {
                "record_id": "sc_1", "main_order_id": "ORD-0501",
                "supplier_order_id": "SUP-1",
                "created_at": "2026-05-10 09:00:00",
                "tracking": {"ship_time": None, "tracking_number": "SF001"},
                "product": {"title": "商品-A", "quantity": 1},
                "cost_breakdown": {
                    "product_supply_cost": "30.00", "courier_fee": "10.00",
                    "total_amount": "40.00",
                },
                "platform_subsidy": "2.00", "cost_granularity": "supplier_line",
                "match_status": "unmatched", "batch_id": "bat_sc_1",
            },
            {
                "record_id": "sc_2", "main_order_id": "ORD-0615",
                "supplier_order_id": "SUP-2",
                "created_at": "2026-06-11 09:00:00",
                "tracking": {"ship_time": "2026-06-12 09:00:00",
                             "tracking_number": "SF002"},
                "product": {"title": "商品-B", "quantity": 2},
                "cost_breakdown": {
                    "product_supply_cost": "45.00", "courier_fee": "15.00",
                    "total_amount": "60.00",
                },
                "platform_subsidy": "0.00", "cost_granularity": "supplier_line",
                "match_status": "unmatched", "batch_id": "bat_sc_2",
            },
        ]
        anomalies = [
            {"record_id": "an_1", "created_at": "2026-05-11 09:00:00",
             "exception_type": "UNMATCHED_ORDER", "severity": "high",
             "status": "pending_review", "order_id": "ORD-0501",
             "message": "订单未匹配成本", "resolution": None},
            {"record_id": "an_2", "created_at": "2026-06-12 09:00:00",
             "exception_type": "QUANTITY_DIFF", "severity": "medium",
             "status": "resolved", "order_id": "ORD-0615",
             "message": "数量差异 1", "resolution": {"note": "已核对"}},
        ]
        import_records = [
            {"batch_id": "BAT-20260501-1", "source_type": "wechat_orders",
             "file_name": "orders_may.xlsx", "imported_at": "2026-05-01T08:00:00",
             "status": "success",
             "statistics": {"total_rows": 10, "success_rows": 10,
                            "updated_rows": 0, "skipped_rows": 0,
                            "duplicate_rows": 0, "error_rows": 0}},
            {"batch_id": "BAT-20260601-2", "source_type": "supplier_cost",
             "file_name": "cost_june.xlsx", "imported_at": "2026-06-01T08:00:00",
             "status": "success",
             "statistics": {"total_rows": 5, "success_rows": 5,
                            "updated_rows": 0, "skipped_rows": 0,
                            "duplicate_rows": 0, "error_rows": 0}},
            {"batch_id": "BAT-20260602-3", "source_type": "wechat_orders",
             "file_name": "orders_june2.xlsx", "imported_at": "2026-06-02T08:00:00",
             "status": "partial_success",
             "statistics": {"total_rows": 8, "success_rows": 7,
                            "updated_rows": 0, "skipped_rows": 0,
                            "duplicate_rows": 0, "error_rows": 1}},
        ]

        cls._write("orders", normal_orders + refund_orders)
        cls._write("refund_details", refund_details)
        cls._write("supplier_costs", supplier_costs)
        cls._write("anomalies", anomalies)
        cls._write("import_records", import_records)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 请求与写入工具 ----------
    @classmethod
    def _write(cls, name, records):
        status, body = cls.admin.post(f"/api/data/{name}",
                                      {"records": records,
                                       "target_store_id": cls.member_store})
        assert status == 200 and body["code"] == 0, f"写入 {name} 失败: {body}"

    def _get(self, path):
        return self.admin.get(path)

    def _data(self, path):
        status, resp = self._get(path)
        self.assertEqual(status, 200, f"{path} HTTP {status}: {resp}")
        self.assertEqual(resp["code"], 0, f"{path} 业务错误: {resp}")
        return resp["data"]

    # ---------- 1. 月末订单归属（左闭右开 + 毫秒精度） ----------
    def test_month_end_and_next_month_first_instant(self):
        data = self._data("/api/orders?month=2026-05")
        ids = {r["order_id"] for r in data["orders"]}
        self.assertEqual(ids, {"ORD-0501", "ORD-0531A", "ORD-0531B"})
        self.assertEqual(data["total"], 3)

        data = self._data("/api/orders?month=2026-06")
        ids = {r["order_id"] for r in data["orders"]}
        self.assertEqual(ids, {"ORD-0601", "ORD-0615"})

    def test_half_open_interval_by_from_to(self):
        # to 为含当日语义（次日零点排他）：05-31 单日 = [05-31 00:00, 06-01 00:00)
        data = self._data("/api/orders?from=2026-05-31&to=2026-05-31")
        ids = {r["order_id"] for r in data["orders"]}
        self.assertEqual(ids, {"ORD-0531A", "ORD-0531B"})

    def test_time_range_meta_half_open(self):
        data = self._data("/api/orders?month=2026-05")
        tr = data["time_range"]
        self.assertEqual(tr["start"], "2026-05-01 00:00:00")
        self.assertEqual(tr["end"], "2026-06-01 00:00:00")
        self.assertEqual(tr["month"], "2026-05")
        self.assertIn("2026-05-01", tr["label"])

    # ---------- 2. 跨月退款不漏记 ----------
    def test_cross_month_refund_in_refund_orders(self):
        # 9 月下单、10 月退款 → 归属 10 月（按申请时间）
        data = self._data("/api/refund-orders?month=2026-10")
        ids = {r["order_id"] for r in data["orders"]}
        self.assertEqual(ids, {"REF-SEP", "REF-OCT"})
        self.assertEqual(data["statistics"]["apply_time_fallback_count"], 0)

    def test_refund_apply_time_fallback_to_order_time(self):
        # 退款时间缺失 → 回退下单时间归属 9 月，不因时间缺失被所有月份排除
        data = self._data("/api/refund-orders?month=2026-09")
        ids = {r["order_id"] for r in data["orders"]}
        self.assertEqual(ids, {"REF-NULL"})
        stats = data["statistics"]
        self.assertEqual(stats["apply_time_fallback_count"], 1)
        row = data["orders"][0]["refund"]
        self.assertEqual(row["refund_apply_time"], "2026-09-20 10:00:00")
        self.assertEqual(row["refund_apply_time_source"], "order_time_fallback")

    def test_refund_orders_time_field_switch(self):
        # 切换为下单时间口径：9 月应同时看到 REF-SEP（9 月下单）与 REF-NULL
        data = self._data("/api/refund-orders?month=2026-09&time_field=order_time")
        ids = {r["order_id"] for r in data["orders"]}
        self.assertEqual(ids, {"REF-SEP", "REF-NULL"})
        self.assertEqual(data["statistics"]["time_field"], "order_time")

    def test_success_amount_by_success_time_separate_caliber(self):
        # 成功退款金额独立口径：按退款成功时间统计
        data = self._data("/api/refund-orders?month=2026-10")
        stats = data["statistics"]
        self.assertEqual(Decimal(stats["success_refund_amount_by_success_time"]),
                         Decimal("76.60"))
        # 9 月无成功退款（REF-NULL 处理中且无退款时间）
        data = self._data("/api/refund-orders?month=2026-09")
        self.assertEqual(
            Decimal(data["statistics"]["success_refund_amount_by_success_time"]),
            Decimal("0"))


    # ---------- 3. 空月份 ----------
    def test_empty_month_returns_empty_everywhere(self):
        for path in (
            "/api/orders?month=2025-01",
            "/api/refund-orders?month=2025-01",
            "/api/supplier-costs?month=2025-01",
            "/api/anomalies?month=2025-01",
            "/api/import/batches?month=2025-01",
            "/api/reconciliation?month=2025-01",
            "/api/profit?month=2025-01",
        ):
            data = self._data(path)
            self.assertEqual(data["total"], 0, f"{path} 应为空")
        # 空月份接口仍应返回统计结构
        data = self._data("/api/orders?month=2025-01")
        self.assertEqual(data["statistics"]["order_count"], 0)
        self.assertEqual(data["orders"], [])

    # ---------- 4. 统计与表格一致 ----------
    def test_statistics_consistent_with_table_orders(self):
        data = self._data("/api/orders?month=2026-05&page_size=20")
        stats = data["statistics"]
        self.assertEqual(stats["order_count"], data["total"])
        self.assertEqual(stats["order_count"], len(data["orders"]))
        total_recv = sum(Decimal(r["amounts"]["actual_received_amount"])
                         for r in data["orders"])
        self.assertEqual(Decimal(stats["actual_received_total"]), total_recv)
        self.assertEqual(Decimal(stats["actual_received_total"]),
                         Decimal("180.50"))

    def test_statistics_consistent_with_table_refund_orders(self):
        data = self._data("/api/refund-orders?month=2026-10")
        stats = data["statistics"]
        self.assertEqual(stats["refund_order_count"], data["total"])
        self.assertEqual(stats["refund_order_count"], len(data["orders"]))
        total = sum(Decimal(r["refund"]["refund_amount"]) for r in data["orders"])
        self.assertEqual(Decimal(stats["total_refund_amount"]), total)
        self.assertEqual(Decimal(stats["total_refund_amount"]), Decimal("76.60"))

    def test_supplier_costs_statistics_consistent(self):
        data = self._data("/api/supplier-costs?month=2026-05")
        self.assertEqual(data["total"], 1)
        self.assertEqual(Decimal(data["statistics"]["total_amount_sum"]),
                         Decimal("40.00"))
        self.assertEqual(data["costs"][0]["record_id"], "sc_1")

    # ---------- 5. 翻页保留筛选 ----------
    def test_pagination_keeps_filter(self):
        p1 = self._data("/api/orders?month=2026-07&page=1&page_size=20")
        p2 = self._data("/api/orders?month=2026-07&page=2&page_size=20")
        self.assertEqual(p1["total"], 25)
        self.assertEqual(p2["total"], 25)
        self.assertEqual(p1["pages"], 2)
        self.assertEqual(len(p1["orders"]), 20)
        self.assertEqual(len(p2["orders"]), 5)
        ids1 = {r["order_id"] for r in p1["orders"]}
        ids2 = {r["order_id"] for r in p2["orders"]}
        self.assertFalse(ids1 & ids2, "页间数据不应重叠")
        self.assertEqual(len(ids1 | ids2), 25)
        # 两页统计口径一致
        self.assertEqual(p1["statistics"], p2["statistics"])

    # ---------- 6. 导出遵守筛选 ----------
    def test_export_respects_month_filter(self):
        status, content, headers = self.admin.get_binary(
            "/api/orders/export?month=2026-05")
        self.assertEqual(status, 200)
        self.assertIn("spreadsheetml", headers.get("Content-Type", ""))
        ws = load_workbook(io.BytesIO(content)).active
        rows = list(ws.iter_rows(values_only=True))
        header, data_rows = rows[0], rows[1:]
        self.assertEqual(len(data_rows), 3, "导出行数应与筛选结果一致")
        exported_ids = {r[header.index("订单号")] for r in data_rows}
        self.assertEqual(exported_ids, {"ORD-0501", "ORD-0531A", "ORD-0531B"})

    # ---------- 7. 全部 / 自定义范围 / 参数优先级 ----------
    def test_no_params_returns_all(self):
        data = self._data("/api/orders")
        self.assertEqual(data["total"], 30)
        tr = data["time_range"]
        self.assertIsNone(tr["start"])
        self.assertIsNone(tr["end"])
        self.assertEqual(tr["label"], "全部时间")

    def test_custom_range(self):
        data = self._data("/api/orders?from=2026-06-01&to=2026-06-16")
        ids = {r["order_id"] for r in data["orders"]}
        self.assertEqual(ids, {"ORD-0601", "ORD-0615"})

    def test_from_to_takes_priority_over_month(self):
        data = self._data("/api/orders?month=2026-05&from=2026-06-01&to=2026-06-30")
        ids = {r["order_id"] for r in data["orders"]}
        self.assertEqual(ids, {"ORD-0601", "ORD-0615"})

    def test_legacy_start_end_still_works(self):
        data = self._data("/api/orders?start=2026-05-01&end=2026-05-31")
        ids = {r["order_id"] for r in data["orders"]}
        self.assertEqual(ids, {"ORD-0501", "ORD-0531A", "ORD-0531B"})

    # ---------- 8. 时区/格式一致性 ----------
    def test_iso_t_and_space_and_millisecond_same_month(self):
        # ISO T 分隔（normalizer 标准格式）与空格分隔 + 毫秒精度均归 5 月
        data = self._data("/api/orders?month=2026-05")
        ids = {r["order_id"] for r in data["orders"]}
        self.assertIn("ORD-0531A", ids)
        self.assertIn("ORD-0531B", ids)
        # 6 月首瞬 00:00:00 归 6 月（不被 5 月的 23:59:59 截断逻辑误吞）
        data = self._data("/api/orders?month=2026-06")
        self.assertIn("ORD-0601", {r["order_id"] for r in data["orders"]})

    # ---------- 9. 其余页面接口 ----------
    def test_supplier_costs_time_field_switch(self):
        data = self._data("/api/supplier-costs?month=2026-06&time_field=ship_time")
        self.assertEqual(data["total"], 1)
        self.assertEqual(data["costs"][0]["record_id"], "sc_2")
        # 按发货时间口径：ship_time 缺失的 sc_1 不出现
        data = self._data("/api/supplier-costs?month=2026-05&time_field=ship_time")
        self.assertEqual(data["total"], 0)

    def test_anomalies_month_filter(self):
        data = self._data("/api/anomalies?month=2026-06")
        self.assertEqual(data["total"], 1)
        self.assertEqual(data["anomalies"][0]["record_id"], "an_2")
        stats = data["statistics"]
        self.assertEqual(stats["anomaly_count"], 1)
        self.assertEqual(stats["resolved_count"], 1)
        self.assertEqual(stats["pending_count"], 0)
        data = self._data("/api/anomalies?month=2026-05")
        self.assertEqual(data["total"], 1)
        self.assertEqual(data["statistics"]["pending_count"], 1)

    def test_import_batches_month_filter(self):
        data = self._data("/api/import/batches?month=2026-06")
        self.assertEqual(data["total"], 2)
        stats = data["statistics"]
        self.assertEqual(stats["batch_count"], 2)
        self.assertEqual(stats["partial_success"], 1)
        self.assertEqual(stats["total_rows"], 13)
        data = self._data("/api/import/batches?month=2026-05")
        self.assertEqual(data["total"], 1)

    def test_summary_respects_month_filter(self):
        data = self._data("/api/summary?month=2026-05")
        self.assertGreaterEqual(data.get("order_count", 0), 1)
        tr = data.get("time_range") or {}
        self.assertEqual(tr.get("month"), "2026-05")

    def test_monthly_profit_respects_month_filter(self):
        data = self._data("/api/monthly-profit?month=2026-05")
        self.assertIn("months", data)
        months = {row["month"] for row in data["months"]}
        self.assertTrue(all(m == "2026-05" for m in months) or len(months) == 0)

    def test_logs_month_filter_and_pagination(self):
        now = datetime.now()
        month = f"{now.year}-{now.month:02d}"
        data = self._data(f"/api/logs?page=1&page_size=20&month={month}")
        self.assertGreater(data["total"], 0)
        self.assertLessEqual(len(data["entries"]), 20)
        self.assertEqual(data["statistics"]["entry_count"], data["total"])
        # 非当月日志被排除
        data = self._data("/api/logs?page=1&page_size=20&month=2025-01")
        self.assertEqual(data["total"], 0)

    # ---------- 参数校验 ----------
    def test_invalid_month_rejected(self):
        for bad in ("2026-13", "2026-1", "abc", "202605"):
            status, resp = self._get(f"/api/orders?month={bad}")
            self.assertEqual(status, 400, f"month={bad} 应返回 400")
            self.assertEqual(resp["code"], 400)
            self.assertIn("时间筛选参数错误", resp["message"])

    def test_invalid_from_to_rejected(self):
        status, _ = self._get("/api/orders?from=2026-06-10&to=2026-06-01")
        self.assertEqual(status, 400)
        status, _ = self._get("/api/orders?from=not-a-date")
        self.assertEqual(status, 400)
        status, _ = self._get("/api/orders?to=2026/06/01")
        self.assertEqual(status, 400)

    def test_invalid_time_field_rejected(self):
        status, _ = self._get("/api/refund-orders?time_field=nope")
        self.assertEqual(status, 400)
        status, _ = self._get("/api/supplier-costs?time_field=nope")
        self.assertEqual(status, 400)


if __name__ == "__main__":
    unittest.main()

"""阶段6交付验收：端到端全流程集成测试

覆盖：原始数据保存、标准化、订单/退款展示与隔离、退款分类、成本匹配（一对一/一对多/
重复/未匹配/疑似/多对一）、人工匹配、利润核算（只扣一次/口径未确认/净收入为零）、
异常处理留痕、四种数据导出、经营总览同口径一致性。
"""
import io
import unittest
from decimal import Decimal

from openpyxl import Workbook, load_workbook

from tests.http_client import TestEnv

WX_HEADERS = ["订单号", "订单下单时间", "订单状态", "订单实际支付金额",
              "订单实际收款金额", "商品数量", "商品标题", "sku编码",
              "技术服务费", "商品已退款金额", "商品售后"]

# (订单号, 时间, 状态, 支付, 实收, 数量, 标题, SKU, 技术费, 退款额, 售后)
WX_ROWS = [
    ("E2E0001", "2026-03-01 10:00:00", "已完成", 100, 100, 1, "商品A", "SKU-A", 3, 0, None),
    ("E2E0002", "2026-03-01 11:00:00", "已完成", 90, 90, 1, "商品A2", "SKU-A2", 2, 0, None),
    ("E2E0010", "2026-03-01 12:00:00", "已完成", 120, 120, 1, "商品B1", "SKU-B1", 4, 0, None),
    ("E2E0010", "2026-03-01 12:00:00", "已完成", 120, 120, 1, "商品B2", "SKU-B2", None, 0, None),
    ("E2E0020", "2026-03-02 10:00:00", "已完成", 200, 200, 1, "商品C1", "SKU-C1", 5, 0, None),
    ("E2E0030", "2026-03-02 11:00:00", "已完成", 60, 60, 1, "商品D", "SKU-D", 1, 0, None),
    ("E2E0040", "2026-03-03 10:00:00", "已完成", 80, 80, 1, "商品E", "SKU-E", 2, 10, "部分退款"),
    ("E2E0050", "2026-03-03 11:00:00", "已完成", 70, 70, 1, "商品F", "SKU-F", 1, 70, "已全额退款"),
    ("E2E0060", "2026-03-04 10:00:00", "已完成", 100, 100, 1, "商品G", "SKU-G", 3, 0, None),
    ("E2E0070", "2026-03-04 11:00:00", "已完成", 100, 100, 1, "商品H", "SKU-H", 3, 0, None),
    ("E2E0080", "2026-03-05 10:00:00", "已完成", 85, 85, 2, "商品X", "SKU-X", 2, 0, None),
    ("E2E0091", "2026-03-05 11:00:00", "已完成", 50, 50, 1, "商品M", "SKU-M", 1, 0, None),
    ("E2E0092", "2026-03-05 12:00:00", "已完成", 50, 50, 1, "商品N", "SKU-N", 1, 0, None),
]

SUP_HEADERS = ["主订单编号", "订单编号", "订单创建时间", "商品数量", "成本", "总金额",
               "快递费合计", "快递费", "泡沫耗材费", "加收", "燃油费", "订单运费",
               "平台补贴", "订单实收款", "sku编码", "快递单号", "商品标题"]

# (主订单, 子订单, 时间, 数量, 成本, 总金额, 快递合计, 快递, 耗材, 加收, 燃油, 运费, 补贴, 实收, SKU, 单号, 标题)
SUP_ROWS = [
    ("E2E0001", None, "2026-03-01 10:05:00", 1, 30, 45, None, None, None, None, None, None, None, 100, "SKU-A", None, "商品A"),
    ("E2E0002", None, "2026-03-01 11:05:00", 1, 20, 35, None, None, None, None, None, None, None, 90, "SKU-A2", None, "商品A2"),
    ("E2E0010", None, "2026-03-01 12:05:00", 1, 10, 25, None, None, None, None, None, None, None, 120, "SKU-B1", "SF001", "商品B1"),
    ("E2E0010", None, "2026-03-01 12:05:00", 1, 10, 25, None, None, None, None, None, None, None, 120, "SKU-B2", "SF002", "商品B2"),
    ("E2E0020", None, "2026-03-02 10:05:00", 1, 30, 48, None, None, None, None, None, None, None, 200, "SKU-C1", None, "商品C-1"),
    ("E2E0020", None, "2026-03-02 10:05:00", 1, 30, 48, None, None, None, None, None, None, 200, "SKU-C2", None, "商品C-2"),
    ("E2E0020", None, "2026-03-02 10:05:00", 1, 30, 48, None, None, None, None, None, None, 200, "SKU-C3", None, "商品C-3"),
    # E2E0030 故意无成本 → 未匹配
    ("E2E0040", None, "2026-03-03 10:05:00", 1, 10, 25, None, None, None, None, None, None, None, 80, "SKU-E", None, "商品E"),
    ("E2E0050", None, "2026-03-03 11:05:00", 1, 20, 30, None, None, None, None, None, None, None, 70, "SKU-F", None, "商品F"),
    # E2E0060：分项总和 56（成本40+快递5+耗材1+加收1+燃油1+快递合计8）但总金额 48 → 只扣 48
    ("E2E0060", None, "2026-03-04 10:05:00", 1, 40, 48, 8, 5, 1, 1, 1, None, None, 100, "SKU-G", None, "商品G"),
    # E2E0070：供应商实收 98 ≠ 微信实收 100 → 金额差异
    ("E2E0070", None, "2026-03-04 11:05:00", 1, 20, 40, None, None, None, None, None, None, None, 98, "SKU-H", None, "商品H"),
    # E2E0080：无主订单编号，仅辅助线索（同日 + SKU-X + 数量 2）→ 疑似匹配
    (None, "SUP-X1", "2026-03-05 09:30:00", 2, 25, 40, None, None, None, None, None, None, None, 85, "SKU-X", None, "商品X"),
    # 多对一：同一成本记录主订单=0091、订单编号=0092
    ("E2E0091", "E2E0092", "2026-03-05 11:05:00", 1, 15, 30, None, None, None, None, None, None, None, 50, "SKU-M", None, "商品M"),
    # 孤儿成本：不属于任何订单
    ("SUP-ORPHAN", None, "2026-03-06 09:00:00", 1, 10, 22, None, None, None, None, None, None, None, None, "SKU-Z", None, "孤儿商品"),
]


def make_wechat_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.append(WX_HEADERS)
    for r in WX_ROWS:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_supplier_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = "供应商"
    ws.append(SUP_HEADERS)
    for r in SUP_ROWS:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestDeliveryE2E(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # admin_user="boss"：操作人取会话身份，保持既有断言 operator == "boss" 口径
        cls.env = TestEnv(admin_user="boss")
        cls.admin = cls.env.admin
        cls.member_store = cls.env.member_store
        cls.wx_content = make_wechat_xlsx()
        cls.sup_content = make_supplier_xlsx()
        cls.admin.upload_confirm(
            f"/api/import/wechat-orders?store_id={cls.member_store}",
            "交付订单.xlsx", cls.wx_content,
            target_store_id=cls.member_store)
        cls.admin.upload_confirm(
            f"/api/import/supplier-costs?store_id={cls.member_store}",
            "交付成本.xlsx", cls.sup_content,
            target_store_id=cls.member_store)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 请求辅助 ----------
    def _get(self, path):
        return self.admin.get(path)

    def _post(self, path, body):
        body = dict(body or {})
        body.setdefault("target_store_id", self.member_store)
        return self.admin.post(path, body)

    def _upload(self, path, filename, content):
        sep = "&" if "?" in path else "?"
        st, resp = self.admin.upload(
            f"{path}{sep}store_id={self.member_store}", filename, content)
        return resp["data"]

    def _upload_confirm(self, path, filename, content, force=False):
        sep = "&" if "?" in path else "?"
        return self.admin.upload_confirm(
            f"{path}{sep}store_id={self.member_store}",
            filename, content, force,
            target_store_id=self.member_store)

    def _get_binary(self, path):
        status, content, _ = self.admin.get_binary(path)
        return status, content

    def _orders_all(self, kind):
        ids = set()
        page = 1
        while True:
            st, r = self._get(f"/api/{kind}?page={page}&size=100")
            self.assertEqual(st, 200)
            for o in r["data"]["orders"]:
                ids.add(o["order_id"])
            if page >= r["data"]["pages"]:
                break
            page += 1
        return ids

    def _profit_of(self, order_id):
        st, r = self._get(f"/api/profit?order_id={order_id}")
        self.assertEqual(st, 200)
        rows = r["data"]["results"]
        self.assertEqual(len(rows), 1)
        return rows[0]

    # ---------- 1. 原始数据保存与标准化 ----------
    def test_01_raw_records_preserved(self):
        st, r = self._get("/api/data/raw_records")
        self.assertEqual(st, 200)
        records = r["data"]["records"]
        self.assertEqual(len(records), len(WX_ROWS) + len(SUP_ROWS))
        wx_raw = [x for x in records
                  if (x.get("raw_fields") or {}).get("订单号") == "E2E0001"]
        self.assertEqual(len(wx_raw), 1)
        raw = wx_raw[0]
        # 原始字段原样留存 + 行号/工作表/批次可追溯
        self.assertEqual(raw["row_number"], 2)
        self.assertEqual(raw["sheet_name"], "订单")
        self.assertTrue(raw["batch_id"])
        self.assertEqual(raw["raw_fields"]["订单实际收款金额"], "100")

        st, r = self._get("/api/data/orders")
        orders = r["data"]["records"]
        self.assertEqual(len(orders), len(WX_ROWS))
        o1 = next(o for o in orders if o["order_id"] == "E2E0001")
        self.assertEqual(o1["amounts"]["actual_received_amount"], "100.00")
        self.assertEqual(o1["amounts"]["technical_service_fee"], "3.00")
        self.assertEqual(o1["source_record_id"], raw["record_id"])

        st, r = self._get("/api/data/supplier_costs")
        costs = r["data"]["records"]
        self.assertEqual(len(costs), len(SUP_ROWS))

    # ---------- 2/4. 订单展示与退款隔离 ----------
    def test_02_order_pages_and_refund_isolation(self):
        normal_ids = self._orders_all("orders")
        refund_ids = self._orders_all("refund-orders")
        refund_like = {"E2E0040", "E2E0050"}
        # 退款订单禁止混入普通订单页
        self.assertEqual(normal_ids & refund_like, set())
        self.assertEqual(refund_ids, refund_like)
        # 退款分类：部分退款 / 全额退款 + 处理中金额不混入成功金额
        st, r = self._get("/api/refund-orders?order_id=E2E0040")
        v = r["data"]["orders"][0]
        self.assertEqual(v["refund"]["refund_status"], "partial")
        self.assertEqual(v["refund_by_status"]["success"], "10.00")
        st, r = self._get("/api/refund-orders?order_id=E2E0050")
        v = r["data"]["orders"][0]
        self.assertEqual(v["refund"]["refund_status"], "full")
        self.assertEqual(v["refund_by_status"]["success"], "70.00")

    # ---------- 3. 成本匹配全场景 ----------
    def test_03_reconciliation_scenarios(self):
        st, r = self._post("/api/reconciliation/run", {})
        self.assertEqual(st, 200)
        st, r = self._get("/api/reconciliation?order_id=E2E0001&size=100")
        rec = r["data"]["results"][0]
        self.assertEqual(rec["match_status"], "matched")
        self.assertEqual(rec["match_method"], "main_order_id_exact")
        # 一对多：两个包裹分别计入
        st, r = self._get("/api/reconciliation?order_id=E2E0010&size=100")
        rec = r["data"]["results"][0]
        self.assertEqual(rec["cost"]["granularity"], "package_level")
        self.assertEqual(rec["cost"]["total_amount"], "50.00")
        # 重复成本：同额 3 行去重只扣一次，粒度待确认
        st, r = self._get("/api/reconciliation?order_id=E2E0020&size=100")
        rec = r["data"]["results"][0]
        self.assertEqual(rec["cost"]["total_amount"], "48.00")
        self.assertTrue(rec["cost"]["order_level_dedup_pending"])
        # 未匹配订单
        st, r = self._get("/api/reconciliation?order_id=E2E0030&size=100")
        rec = r["data"]["results"][0]
        self.assertEqual(rec["match_status"], "unmatched_order")
        # 疑似匹配（辅助线索，禁止自动认定）
        st, r = self._get("/api/reconciliation?order_id=E2E0080&size=100")
        rec = r["data"]["results"][0]
        self.assertEqual(rec["match_status"], "suspected")
        # 金额差异异常
        st, r = self._get("/api/reconciliation?order_id=E2E0070&size=100")
        rec = r["data"]["results"][0]
        self.assertIn("AMOUNT_DIFF", rec["exception_codes"])
        # 多对一异常
        st, r = self._get("/api/data/anomalies")
        multi = [a for a in r["data"]["records"]
                 if a["exception_type"] == "MULTI_ORDER_ONE_COST"
                 and a["status"] == "pending_review"]
        self.assertTrue(any("E2E0091" in a["message"] for a in multi))
        # 孤儿成本
        st, r = self._get("/api/reconciliation?order_id=SUP-ORPHAN")
        rec = r["data"]["results"][0]
        self.assertEqual(rec["match_status"], "unmatched_supplier")

    # ---------- 4. 人工匹配与异常留痕 ----------
    def test_04_manual_match_and_anomaly_resolution(self):
        st, r = self._get("/api/data/supplier_costs")
        sid = next(c["record_id"] for c in r["data"]["records"]
                   if c["main_order_id"] == "SUP-ORPHAN")
        st, r = self._post("/api/reconciliation/manual-match", {
            "order_id": "E2E0030", "supplier_record_ids": [sid],
            "note": "交付测试人工绑定"})
        self.assertEqual(st, 200)
        rec = r["data"]
        self.assertEqual(rec["match_status"], "manual_confirmed")
        h = rec["manual_review"]["history"][0]
        self.assertEqual(h["operator"], "boss")
        self.assertEqual(h["before"]["match_status"], "unmatched_order")
        # 异常关闭且留痕
        st, r = self._get("/api/data/anomalies")
        a = next(x for x in r["data"]["records"]
                 if x["order_id"] == "E2E0030"
                 and x["exception_type"] == "UNMATCHED_ORDER")
        self.assertEqual(a["status"], "resolved")
        self.assertIn("boss", a["resolution"]["text"])

    # ---------- 5/6. 利润核算：只扣一次 + 口径未确认提示 ----------
    def test_05_profit_pending_confirmation_hint(self):
        self._post("/api/profit/config",
                   {"refund_deduction_status": "unknown"})
        st, r = self._post("/api/profit/recalculate", {})
        self.assertEqual(st, 200)
        self.assertGreater(r["data"]["pending_confirmation"], 0)
        rec = self._profit_of("E2E0040")
        self.assertEqual(rec["profit_status"], "pending_confirmation")
        self.assertIn("REFUND_DEDUCTION_UNCONFIRMED", rec["status_reasons"])
        # 接口提示文案 + 页面口径声明
        st, r = self._get("/api/profit/options")
        self.assertIn("待确认",
                      r["data"]["config"]["labels"]["refund_deduction_status"])

    def test_06_profit_single_cost_deduction(self):
        self._post("/api/profit/config", {
            "refund_deduction_status": "deducted",
            "platform_subsidy_included": "unknown"})
        st, r = self._post("/api/profit/recalculate", {})
        self.assertEqual(st, 200)
        # E2E0060：分项和 56（40+8+5+1+1+1），总金额 48 → 只扣 48
        rec = self._profit_of("E2E0060")
        self.assertEqual(rec["profit_status"], "calculated")
        self.assertEqual(rec["amounts"]["supplier_total_amount"], "48.00")
        # 100 − 48 − 3 = 49（若误扣分项则不会是 49）
        self.assertEqual(rec["amounts"]["operating_contribution_profit"], "49.00")
        self.assertEqual(rec["margin"]["percent"], "49.00%")
        # 分项仅展示
        self.assertEqual(rec["cost_parts_display_only"]["product_supply_cost"], "40.00")
        self.assertEqual(rec["cost_parts_display_only"]["courier_fee_total"], "8.00")
        # 净收入为零：全额退款单不计算利润率
        rec = self._profit_of("E2E0050")
        self.assertEqual(rec["amounts"]["net_revenue"], "70.00")
        st, r = self._post("/api/profit/config",
                           {"refund_deduction_status": "not_deducted"})
        self._post("/api/profit/recalculate", {})
        rec = self._profit_of("E2E0050")
        self.assertEqual(rec["amounts"]["net_revenue"], "0.00")
        self.assertIsNone(rec["margin"]["percent"])
        self.assertEqual(rec["margin"]["blocked_reason"], "NET_REVENUE_NOT_POSITIVE")
        # 处理中退款不计入成功退款（not_deducted 下 E2E0040 只扣 10）
        rec = self._profit_of("E2E0040")
        self.assertEqual(rec["amounts"]["net_revenue"], "70.00")

    # ---------- 7. 四种导出与筛选一致性 ----------
    def test_07_exports_consistency(self):
        cases = [
            ("/api/orders/export?status=completed", "订单号"),
            ("/api/refund-orders/export", "订单号"),
            ("/api/reconciliation/export?anomaly_type=DUPLICATE_TOTAL_AMOUNT_RISK", "订单号"),
            ("/api/profit/export?profit_status=calculated", "订单号"),
        ]
        for path, first_header in cases:
            st, content = self._get_binary(path)
            self.assertEqual(st, 200, path)
            wb = load_workbook(io.BytesIO(content))
            rows = list(wb.active.iter_rows(values_only=True))
            self.assertGreaterEqual(len(rows), 2, path)
            self.assertEqual(rows[0][0], first_header, path)
        # 导出记录留痕
        st, r = self._get("/api/data/export_records")
        kinds = {e["kind"] for e in r["data"]["records"]}
        self.assertTrue({"normal", "refund", "reconciliation", "profit"} <= kinds)
        # 利润导出行数与筛选一致
        st, r = self._get("/api/profit?profit_status=calculated&size=100")
        expected = len(r["data"]["results"])
        st, content = self._get_binary(
            "/api/profit/export?profit_status=calculated")
        rows = list(load_workbook(io.BytesIO(content)).active.iter_rows(
            values_only=True))
        self.assertEqual(len(rows) - 1, expected)

    # ---------- 8. 经营总览同口径一致性 ----------
    def test_08_summary_consistency(self):
        self._post("/api/profit/recalculate", {})
        st, r = self._get("/api/summary")
        self.assertEqual(st, 200)
        s = r["data"]
        self.assertTrue(s["has_orders"])
        self.assertIn("口径声明", s["declaration"])
        # 明细页逐单求和（同口径）：技术费覆盖全部订单，
        # 利润/净收入仅统计可计算订单（CALCULATED_LIKE）
        received = net = tech = profit = Decimal("0")
        count = 0
        page = 1
        while True:
            st, r = self._get(f"/api/profit?page={page}&size=100")
            for rec in r["data"]["results"]:
                a = rec["amounts"]
                if a.get("technical_service_fee") is not None:
                    tech += Decimal(a["technical_service_fee"])
                if rec["profit_status"] in ("calculated", "needs_review"):
                    received += Decimal(a["actual_received_amount"])
                    net += Decimal(a["net_revenue"])
                    profit += Decimal(a["operating_contribution_profit"])
                    count += 1
            if page >= r["data"]["pages"]:
                break
            page += 1
        self.assertEqual(s["order_count"], 12)
        self.assertEqual(s["calculable_order_count"], count)
        self.assertAlmostEqual(Decimal(s["operating_contribution_profit"]),
                               profit, delta=Decimal("0.01"))
        self.assertAlmostEqual(Decimal(s["technical_service_fee"]), tech,
                               delta=Decimal("0.01"))
        # 未匹配订单数与异常数量
        self.assertGreaterEqual(s["unmatched_order_count"], 0)
        self.assertGreaterEqual(s["open_anomaly_count"], 1)


if __name__ == "__main__":
    unittest.main()

"""每月盈利：销售笔数、销售额、货品成本、快递成本，广告费和其他费用手填后重算。"""
import io
import unittest

from openpyxl import Workbook

from tests.http_client import TestEnv

WX_HEADERS = ["订单号", "订单下单时间", "订单状态", "订单实际支付金额",
              "订单实际收款金额", "商品数量", "商品标题", "sku编码",
              "商品实际价格(总共)", "技术服务费", "商品已退款金额", "商品售后"]
SUP_HEADERS = ["主订单编号", "订单编号", "订单创建时间", "商品数量", "总金额",
               "成本", "快递费", "泡沫耗材费", "加收", "燃油费", "快递费合计",
               "平台补贴", "sku编码", "快递单号", "商品标题"]


def _xlsx(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestMonthlyProfit(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.member = cls.env.member
        cls.store = cls.env.member_store
        wx = _xlsx(WX_HEADERS, [
            ["M4001", "2026-04-02 10:00:00", "已完成", 100, 100, 1,
             "甲", "SKU-A", 100, 3, 10, "部分退款"],
            ["M4002", "2026-05-03 10:00:00", "已完成", 60, 60, 1,
             "乙", "SKU-B", 60, 2, None, None],
        ])
        sup = _xlsx(SUP_HEADERS, [
            ["M4001", None, "2026-04-02 11:00:00", 1, 48, 40, 6.5,
             None, None, None, 8, None, "SKU-A", None, "甲"],
            ["M4002", None, "2026-05-03 11:00:00", 1, 30, None, None,
             None, None, None, None, None, "SKU-B", None, "乙"],
        ])
        cls._import("wechat-orders", "月盈利订单.xlsx", wx)
        cls._import("supplier-costs", "月盈利成本.xlsx", sup)
        status, resp = cls.member.post("/api/reconciliation/run", {
            "target_store_id": cls.store})
        assert status == 200 and resp["code"] == 0, resp
        status, resp = cls.member.post("/api/profit/recalculate", {
            "target_store_id": cls.store})
        assert status == 200 and resp["code"] == 0, resp

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    @classmethod
    def _import(cls, kind, filename, content):
        status, resp = cls.env.admin.upload(
            f"/api/import/{kind}?store_id={cls.store}", filename, content)
        assert status == 200 and resp["code"] == 0, resp
        status, resp = cls.env.admin.post("/api/import/confirm", {
            "upload_id": resp["data"]["upload_id"],
            "target_store_id": cls.store})
        assert status == 200 and resp["code"] == 0, resp

    def _month(self, month):
        status, resp = self.member.get(
            f"/api/monthly-profit?month={month}&store_id={self.store}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        rows = resp["data"]["months"]
        self.assertEqual(len(rows), 1)
        return rows[0], resp["data"]["totals"]

    def test_01_month_sales_and_costs(self):
        row, totals = self._month("2026-04")
        self.assertEqual(row["order_count"], 1)
        self.assertEqual(row["sales_amount"], "100.00")
        self.assertEqual(row["product_cost_total"], "40.00")
        self.assertEqual(row["courier_fee_total"], "8.00")
        self.assertEqual(row["technical_service_fee"], "3.00")
        self.assertEqual(row["refund_amount"], "10.00")
        self.assertEqual(row["ad_expense"], "0.00")
        self.assertFalse(row["ad_recorded"])
        self.assertEqual(row["profit"], "39.00")
        self.assertEqual(totals["profit"], "39.00")
        self.assertEqual(totals["technical_service_fee"], "3.00")
        self.assertEqual(totals["refund_amount"], "10.00")
        self.assertEqual(row["product_cost_missing_count"], 0)

    def test_02_manual_ad_recalculates_profit(self):
        status, resp = self.member.post("/api/ad-expenses", {
            "month": "2026-04", "amount": "10.00",
            "target_store_id": self.store})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        row, _ = self._month("2026-04")
        self.assertTrue(row["ad_recorded"])
        self.assertEqual(row["ad_expense"], "10.00")
        self.assertEqual(row["refund_amount"], "10.00")
        self.assertEqual(row["profit"], "29.00")

    def test_03_may_ad_and_costs(self):
        status, resp = self.member.post("/api/ad-expenses", {
            "month": "2026-05", "amount": "5.00",
            "target_store_id": self.store})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        april, _ = self._month("2026-04")
        self.assertEqual(april["ad_expense"], "10.00")
        self.assertEqual(april["profit"], "29.00")
        may, _ = self._month("2026-05")
        self.assertEqual(may["sales_amount"], "60.00")
        self.assertEqual(may["product_cost_total"], "0.00")
        self.assertEqual(may["product_cost_missing_count"], 1)
        self.assertEqual(may["courier_fee_missing_count"], 1)
        self.assertEqual(may["technical_service_fee"], "2.00")
        self.assertEqual(may["refund_amount"], "0.00")
        self.assertEqual(may["ad_expense"], "5.00")
        self.assertEqual(may["profit"], "53.00")

    def test_05_page_is_in_nav(self):
        status, body, _ = self.member.get_page("/monthly-profit")
        self.assertEqual(status, 200)
        self.assertIn("每月盈利", body)
        self.assertIn('href="/monthly-profit" class="active"', body)
        self.assertIn("销售笔数", body)
        self.assertIn("货品成本合计", body)
        self.assertIn("快递成本合计", body)
        self.assertIn("技术服务费合计", body)
        self.assertIn("退款金额", body)
        self.assertIn("其他费用", body)
        self.assertNotIn("导入并重置", body)
        self.assertNotIn("/api/ad-expenses/import", body)
        self.assertIn(">合计<", body)

    def test_06_empty_month_points_at_latest_fee_month(self):
        status, resp = self.member.get(
            f"/api/monthly-profit?month=2026-10&store_id={self.store}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        data = resp["data"]
        self.assertEqual(data["months"], [])
        self.assertEqual(data["totals"]["technical_service_fee"], "0.00")
        self.assertEqual(data["latest_data_month"], "2026-05")
        self.assertEqual(data["latest_technical_service_fee"], "2.00")
        status, resp = self.member.get(
            f"/api/profit?month=2026-10&store_id={self.store}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["total"], 0)
        self.assertEqual(resp["data"]["statistics"]["technical_service_fee"], "0.00")
        self.assertEqual(resp["data"]["latest_order_month"], "2026-05")
        status, resp = self.member.get(
            f"/api/monthly-profit?month=2026-05&store_id={self.store}")
        self.assertEqual(resp["data"]["totals"]["technical_service_fee"], "2.00")
        status, resp = self.member.get(
            f"/api/profit?month=2026-05&store_id={self.store}&size=100")
        self.assertEqual(resp["data"]["statistics"]["technical_service_fee"], "2.00")

    def test_07_other_expense_is_editable_and_deducted(self):
        status, resp = self.member.post("/api/other-expenses", {
            "month": "2026-04", "amount": "4.00", "note": "包装",
            "target_store_id": self.store})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        row, totals = self._month("2026-04")
        self.assertEqual(row["other_expense"], "4.00")
        self.assertEqual(row["other_note"], "包装")
        self.assertTrue(row["other_recorded"])
        self.assertEqual(row["profit"], "25.00")
        self.assertEqual(totals["other_expense"], "4.00")
        self.assertEqual(totals["profit"], "25.00")

        status, resp = self.member.post("/api/other-expenses", {
            "month": "2026-04", "amount": "7.50", "note": "包材",
            "target_store_id": self.store})
        self.assertEqual(status, 200)
        row, _ = self._month("2026-04")
        self.assertEqual(row["other_expense"], "7.50")
        self.assertEqual(row["other_note"], "包材")
        self.assertEqual(row["profit"], "21.50")

        status, resp = self.member.post("/api/other-expenses", {
            "month": "2026-04", "amount": "-1",
            "target_store_id": self.store})
        self.assertEqual(status, 400)
        row, _ = self._month("2026-04")
        self.assertEqual(row["other_expense"], "7.50")

        status, resp = self.member.post("/api/other-expenses", {
            "month": "2026-05", "amount": "1.00",
            "target_store_id": self.store})
        self.assertEqual(status, 200)
        may, _ = self._month("2026-05")
        self.assertEqual(may["other_expense"], "1.00")
        self.assertEqual(may["profit"], "52.00")

        status, resp = self.member.post("/api/other-expenses", {
            "month": "2026-06", "amount": "6.00", "note": "杂费",
            "target_store_id": self.store})
        self.assertEqual(status, 200)
        june, _ = self._month("2026-06")
        self.assertEqual(june["order_count"], 0)
        self.assertEqual(june["other_expense"], "6.00")
        self.assertEqual(june["profit"], "-6.00")

    def test_08_partial_recalc_still_shows_may(self):
        """导入/重算若只覆盖部分月份，每月盈利仍应保留其他月份汇总。"""
        may, _ = self._month("2026-05")
        self.assertEqual(may["sales_amount"], "60.00")
        status, resp = self.member.post("/api/profit/recalculate", {
            "month": "2026-04", "target_store_id": self.store})
        self.assertEqual(status, 200, resp)
        may2, _ = self._month("2026-05")
        self.assertEqual(may2["sales_amount"], "60.00",
                         "局部重算后 5 月汇总不应丢失")

"""未匹配订单手工总成本 = 货品 + 快递。

不要求先选择店铺。后续供应商成本入库并匹配后，以入库金额为准。
"""
import io
import unittest

from openpyxl import Workbook

from tests.http_client import TestEnv

WX_HEADERS = ["订单号", "订单下单时间", "订单状态", "订单实际支付金额",
              "订单实际收款金额", "商品数量", "商品标题", "sku编码"]
SUP_HEADERS = ["主订单编号", "订单编号", "订单创建时间", "商品数量", "总金额",
               "订单实收款", "sku编码", "快递单号", "商品标题"]


def _xlsx(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestUnmatchedOrderCost(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv(admin_user="cost_admin")
        cls.admin = cls.env.admin
        cls.store = cls.env.member_store
        wx = _xlsx(WX_HEADERS, [
            ["O9001", "2026-04-02 10:00:00", "已完成", 100, 100, 1, "商品A", "SKU-A"],
            ["O9002", "2026-04-03 10:00:00", "已完成", 60, 60, 1, "商品B", "SKU-B"],
        ])
        status, resp = cls.admin.upload(
            f"/api/import/wechat-orders?store_id={cls.store}", "订单.xlsx", wx)
        assert status == 200 and resp["code"] == 0, resp
        status, resp = cls.admin.post("/api/import/confirm", {
            "upload_id": resp["data"]["upload_id"],
            "target_store_id": cls.store,
        })
        assert status == 200 and resp["code"] == 0, resp

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _recon(self):
        status, resp = self.admin.get("/api/reconciliation?order_id=O9001")
        self.assertEqual(status, 200)
        rows = resp["data"]["results"]
        self.assertEqual(len(rows), 1)
        return rows[0]

    def _profit(self):
        status, resp = self.admin.get("/api/profit?order_id=O9001")
        self.assertEqual(status, 200)
        rows = resp["data"]["results"]
        self.assertEqual(len(rows), 1)
        return rows[0]

    def test_01_manual_total_without_selecting_store(self):
        row = self._recon()
        self.assertEqual(row["match_status"], "unmatched_order")
        self.assertFalse(row.get("supplier_record_ids"))

        status, resp = self.admin.post("/api/reconciliation/manual-order-cost", {
            "order_id": "O9001",
            "product_amount": "10.00",
            "courier_amount": "5.50",
        })
        self.assertEqual(status, 200, resp)
        self.assertEqual(resp["code"], 0)
        self.assertEqual(resp["data"]["total_amount"], "15.50")
        self.assertEqual(resp["data"]["match_status"], "unmatched_order")

        row = self._recon()
        self.assertEqual(row["cost"]["total_amount"], "15.50")
        self.assertEqual(row["manual_cost"]["product_amount"], "10.00")
        self.assertEqual(row["manual_cost"]["courier_amount"], "5.50")
        self.assertTrue(row["cost"]["usable_for_profit"])

        profit = self._profit()
        self.assertEqual(profit["profit_status"], "calculated")
        self.assertEqual(profit["cost_match_status"], "manual")
        self.assertEqual(profit["amounts"]["product_cost"], "10.00")
        self.assertEqual(profit["amounts"]["courier_fee"], "5.50")
        self.assertEqual(profit["amounts"]["supplier_total_amount"], "15.50")

    def test_01b_batch_two_orders(self):
        status, resp = self.admin.post("/api/reconciliation/manual-order-cost", {
            "items": [
                {"order_id": "O9001", "product_amount": "10.00",
                 "courier_amount": "5.50"},
                {"order_id": "O9001", "product_amount": "1.00",
                 "courier_amount": "1.00"},
            ],
        })
        self.assertEqual(status, 400, resp)
        self.assertEqual(resp["code"], 3006)

        status, resp = self.admin.post("/api/reconciliation/manual-order-cost", {
            "items": [
                {"order_id": "O9001", "product_amount": "10.00",
                 "courier_amount": "5.50"},
                {"order_id": "O9002", "product_amount": "3.00",
                 "courier_amount": "4.00"},
            ],
        })
        self.assertEqual(status, 200, resp)
        self.assertEqual(resp["data"]["count"], 2)
        by_id = {item["order_id"]: item for item in resp["data"]["items"]}
        self.assertEqual(by_id["O9001"]["total_amount"], "15.50")
        self.assertEqual(by_id["O9002"]["total_amount"], "7.00")

        status, resp = self.admin.get("/api/reconciliation?order_id=O9002")
        self.assertEqual(status, 200)
        row = resp["data"]["results"][0]
        self.assertEqual(row["match_status"], "unmatched_order")
        self.assertEqual(row["manual_cost"]["product_amount"], "3.00")
        self.assertEqual(row["manual_cost"]["courier_amount"], "4.00")

    def test_02_rerun_keeps_manual_until_import(self):
        status, resp = self.admin.post("/api/reconciliation/run", {
            "target_store_id": self.store,
        })
        self.assertEqual(status, 200)
        row = self._recon()
        self.assertEqual(row["match_status"], "unmatched_order")
        self.assertEqual(row["cost"]["total_amount"], "15.50")

    def test_03_invalid_amount_rejected(self):
        status, resp = self.admin.post("/api/reconciliation/manual-order-cost", {
            "order_id": "O9001",
            "product_amount": "-1",
            "courier_amount": "1.00",
        })
        self.assertEqual(status, 400, resp)
        self.assertEqual(resp["code"], 3006)
        row = self._recon()
        self.assertEqual(row["cost"]["total_amount"], "15.50")

    def test_04_import_overwrites_manual_cost(self):
        sup = _xlsx(SUP_HEADERS, [[
            "O9001", None, "2026-04-02 11:00:00", 1, 80, 100, "SKU-A", None, "商品A",
        ]])
        status, resp = self.admin.upload(
            f"/api/import/supplier-costs?store_id={self.store}", "成本.xlsx", sup)
        self.assertEqual(status, 200, resp)
        status, resp = self.admin.post("/api/import/confirm", {
            "upload_id": resp["data"]["upload_id"],
            "target_store_id": self.store,
        })
        self.assertEqual(status, 200, resp)

        row = self._recon()
        self.assertEqual(row["match_status"], "matched")
        self.assertEqual(row["cost"]["total_amount"], "80.00")
        self.assertNotIn("manual_cost", row)

        status, resp = self.admin.post("/api/profit/recalculate", {
            "target_store_id": self.store,
        })
        self.assertEqual(status, 200, resp)
        profit = self._profit()
        self.assertEqual(profit["amounts"]["supplier_total_amount"], "80.00")
        self.assertNotEqual(profit["cost_match_status"], "manual")

        status, resp = self.admin.post("/api/reconciliation/manual-order-cost", {
            "order_id": "O9001",
            "product_amount": "1.00",
            "courier_amount": "1.00",
        })
        self.assertEqual(status, 403, resp)
        self.assertEqual(resp["code"], 3004)

        status, resp = self.admin.get("/api/reconciliation?order_id=O9002")
        self.assertEqual(status, 200)
        other = resp["data"]["results"][0]
        self.assertEqual(other["match_status"], "unmatched_order")
        self.assertEqual(other["cost"]["total_amount"], "7.00")

"""导入时自动剔除未发货订单验收"""
import io
import unittest

from openpyxl import Workbook

from app.profit_service import is_unshipped_order
from tests.http_client import TestEnv

WX_HEADERS = [
    "订单号", "订单下单时间", "订单状态", "订单实际支付金额",
    "订单实际收款金额", "商品数量", "商品标题", "sku编码",
    "技术服务费", "商品已退款金额", "商品售后",
]

WX_HEADERS_SHIP = WX_HEADERS + ["商品发货"]


def _xlsx(rows, headers=WX_HEADERS):
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _row(oid, status, recv=100):
    return (oid, "2026-03-01 10:00:00", status, recv, recv,
            1, "商品", "SKU", 3, 0, None)


class TestImportUnshippedSkip(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.member = cls.env.create_user_client(
            "imp_unship", "ImpUnship123456", store_name="导入剔除测试店")
        cls.store_id = cls.member.get("/api/auth/me")[1]["data"]["store_id"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _precheck(self, xlsx):
        st, r = self.member.upload(
            "/api/import/wechat-orders?store_id=" + self.store_id,
            "t.xlsx", xlsx)
        self.assertEqual(st, 200, r)
        return r.get("data", r)

    def _confirm(self, upload_id):
        st, r = self.member.post("/api/import/confirm", {
            "upload_id": upload_id,
            "target_store_id": self.store_id,
        })
        self.assertEqual(st, 200, r)
        return r.get("data", r)

    def test_pending_shipment_not_in_orders(self):
        xlsx = _xlsx([_row("IU001", "待发货"), _row("IU002", "已发货")])
        prev = self._precheck(xlsx)
        pc = prev.get("precheck", {})
        self.assertEqual(pc.get("excluded_unshipped_rows"), 1)
        self.assertEqual(pc.get("success_rows"), 1)
        self._confirm(prev["upload_id"])
        st, r = self.member.get("/api/orders?page=1&size=50")
        ids = {o.get("order_id") for o in r.get("data", {}).get("orders", [])}
        self.assertNotIn("IU001", ids)
        self.assertIn("IU002", ids)

    def test_ship_column_unshipped_skipped(self):
        row_ship = ("IU010", "2026-03-01 10:00:00", "已完成", 100, 100,
                    1, "商品", "SKU", 3, 0, None, "未发货")
        xlsx = _xlsx([row_ship], headers=WX_HEADERS_SHIP)
        prev = self._precheck(xlsx)
        self.assertEqual(prev["precheck"]["excluded_unshipped_rows"], 1)
        self._confirm(prev["upload_id"])
        st, r = self.member.get("/api/orders")
        ids = {o.get("order_id") for o in r.get("data", {}).get("orders", [])}
        self.assertNotIn("IU010", ids)

    def test_is_unshipped_order_unit(self):
        self.assertTrue(is_unshipped_order("pending_shipment", None))
        self.assertTrue(is_unshipped_order("completed", "unshipped"))
        self.assertFalse(is_unshipped_order("completed", "shipped"))
        self.assertFalse(is_unshipped_order("completed", "partial_shipped"))


if __name__ == "__main__":
    unittest.main()

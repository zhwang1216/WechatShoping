"""对账统计商品发货状态过滤验收

仅有商品发货=已发货的订单进入对账；未发货/部分发货排除。
无商品发货列时按订单状态向后兼容（与 is_reconciliation_scope 一致）。
"""
import io
import unittest

from openpyxl import Workbook

from app.profit_service import is_reconciliation_scope
from tests.http_client import TestEnv

WX_HEADERS_WITH_SHIP = [
    "订单号", "订单下单时间", "订单状态", "订单实际支付金额",
    "订单实际收款金额", "商品数量", "商品标题", "sku编码",
    "技术服务费", "商品已退款金额", "商品售后", "商品发货",
]

SUP_HEADERS = [
    "主订单编号", "订单编号", "订单创建时间", "商品数量", "总金额",
    "订单实收款", "sku编码", "快递单号", "商品标题",
]


def _make_order_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.append(WX_HEADERS_WITH_SHIP)
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _make_sup_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "供应商成本"
    ws.append(SUP_HEADERS)
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _order(order_id, ship_status, received=100):
    return (order_id, "2026-03-01 10:00:00", "已完成", received, received,
            1, "商品", "SKU-X", 3, 0, None, ship_status)


def _sup(main_oid, total=50):
    return (main_oid, None, "2026-03-01 10:05:00", 1, total, total,
            "SKU-X", None, "商品")


class TestReconciliationShipStatus(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.member = cls.env.create_user_client(
            "recon_ship_member", "ReconShip123456", store_name="对账发货测试店")
        cls.store_id = cls.member.get("/api/auth/me")[1]["data"]["store_id"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _import_orders(self, rows):
        st, r = self.member.upload_confirm(
            "/api/import/wechat-orders?store_id=" + self.store_id,
            "orders.xlsx", _make_order_xlsx(rows), target_store_id=self.store_id)
        self.assertEqual(st, 200, r)

    def _import_costs(self, rows):
        st, r = self.member.upload_confirm(
            "/api/import/supplier-costs?store_id=" + self.store_id,
            "costs.xlsx", _make_sup_xlsx(rows), target_store_id=self.store_id)
        self.assertEqual(st, 200, r)

    def _run_reconcile(self):
        st, r = self.member.post("/api/reconciliation/run", {})
        self.assertEqual(st, 200, r)
        return r.get("data", r)

    def _recon_order_ids(self):
        st, r = self.member.get("/api/reconciliation?page=1&size=100")
        self.assertEqual(st, 200, r)
        data = r.get("data", {})
        items = data.get("results", data.get("items", data.get("records", [])))
        return {x.get("order_id") for x in items if x.get("order_id")}

    def test_unshipped_excluded_from_reconciliation(self):
        self._import_orders([
            _order("RS001", "未发货"),
            _order("RS002", "已发货"),
        ])
        self._import_costs([_sup("RS001"), _sup("RS002")])
        self._run_reconcile()
        ids = self._recon_order_ids()
        self.assertNotIn("RS001", ids, "商品发货=未发货 不应进入对账")
        self.assertIn("RS002", ids, "商品发货=已发货 应进入对账")

    def test_partial_ship_excluded(self):
        self._import_orders([_order("RS010", "部分发货")])
        self._import_costs([_sup("RS010")])
        self._run_reconcile()
        ids = self._recon_order_ids()
        self.assertNotIn("RS010", ids, "部分发货不应进入对账统计")

    def test_is_reconciliation_scope_unit(self):
        self.assertFalse(is_reconciliation_scope("completed", "unshipped"))
        self.assertFalse(is_reconciliation_scope("completed", "partial_shipped"))
        self.assertTrue(is_reconciliation_scope("completed", "shipped"))
        self.assertTrue(is_reconciliation_scope("completed", None),
                       "无商品发货列时已完成订单应对账（兼容）")
        self.assertFalse(is_reconciliation_scope("pending_shipment", None))


if __name__ == "__main__":
    unittest.main()

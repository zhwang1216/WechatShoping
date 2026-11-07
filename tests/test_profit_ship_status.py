"""商品发货状态过滤 验收测试

导入表格中"商品发货"列值为"未发货"的订单不进入利润统计范畴。
覆盖场景:
  test_01: 商品发货=未发货 → 排除
  test_02: 商品发货=已发货 → 包含
  test_03: 商品发货=部分发货 → 包含
  test_04: 商品发货列缺失 → 按 order_status 判断（向后兼容）
  test_05: 订单状态=已完成 但 商品发货=未发货 → 排除（ship_status 优先）
  test_06: is_calculable_status 单元测试
"""
import io
import unittest

from openpyxl import Workbook

from app.profit_service import is_calculable_status
from tests.http_client import TestEnv


WX_HEADERS_WITH_SHIP = [
    "订单号", "订单下单时间", "订单状态", "订单实际支付金额",
    "订单实际收款金额", "商品数量", "商品标题", "sku编码",
    "技术服务费", "商品已退款金额", "商品售后", "商品发货",
]

WX_HEADERS_NO_SHIP = [
    "订单号", "订单下单时间", "订单状态", "订单实际支付金额",
    "订单实际收款金额", "商品数量", "商品标题", "sku编码",
    "技术服务费", "商品已退款金额", "商品售后",
]


def _make_order_xlsx(rows, headers=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.append(headers or WX_HEADERS_WITH_SHIP)
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _order_with_ship(order_id, order_status, received, ship_status, qty=1):
    return (order_id, "2026-03-01 10:00:00", order_status, received, received,
            qty, "商品", "SKU-X", 3, 0, None, ship_status)


def _order_no_ship(order_id, order_status, received, qty=1):
    return (order_id, "2026-03-01 10:00:00", order_status, received, received,
            qty, "商品", "SKU-X", 3, 0, None)


class TestProfitShipStatus(unittest.TestCase):
    """商品发货状态过滤 验收测试"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            "ps_member", "PsMember123456", store_name="PS测试店")
        cls.store_id = cls.member.get("/api/auth/me")[1]["data"]["store_id"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _import_orders(self, rows, headers=None):
        xlsx = _make_order_xlsx(rows, headers)
        st, r = self.member.upload_confirm(
            "/api/import/wechat-orders?store_id=" + self.store_id,
            "test.xlsx", xlsx, target_store_id=self.store_id)
        self.assertEqual(st, 200, f"导入失败: {r}")
        return r

    def _recalculate(self):
        st, r = self.admin.post(
            "/api/profit/recalculate", {"store_id": self.store_id})
        self.assertEqual(st, 200, f"重算失败: {r}")

    def _profit_order_ids(self):
        ids = set()
        page = 1
        while True:
            st, r = self.member.get(f"/api/profit?page={page}&size=100")
            self.assertEqual(st, 200, r)
            data = r.get("data", {})
            items = data.get("results", data.get("items", data.get("records", [])))
            for item in items:
                ids.add(item.get("order_id"))
            if len(items) < 100:
                break
            page += 1
        return ids

    def test_01_ship_unshipped_excluded(self):
        """商品发货=未发货 的订单不进入利润核算"""
        self._import_orders([
            _order_with_ship("PS001", "已完成", 100, "未发货"),
            _order_with_ship("PS002", "已完成", 200, "已发货"),
        ])
        self._recalculate()
        ids = self._profit_order_ids()
        self.assertNotIn("PS001", ids, "商品发货=未发货 的订单不应进入利润核算")
        self.assertIn("PS002", ids, "商品发货=已发货 的订单应进入利润核算")

    def test_02_ship_shipped_included(self):
        """商品发货=已发货 的订单正常进入利润核算"""
        self._import_orders([
            _order_with_ship("PS010", "已完成", 300, "已发货"),
        ])
        self._recalculate()
        ids = self._profit_order_ids()
        self.assertIn("PS010", ids, "商品发货=已发货 的订单应进入利润核算")

    def test_03_ship_partial_included(self):
        """商品发货=部分发货 的订单进入利润核算"""
        self._import_orders([
            _order_with_ship("PS020", "已完成", 150, "部分发货"),
        ])
        self._recalculate()
        ids = self._profit_order_ids()
        self.assertIn("PS020", ids, "商品发货=部分发货 的订单应进入利润核算")

    def test_04_ship_missing_fallback(self):
        """商品发货列缺失时按 order_status 判断（向后兼容）"""
        self._import_orders(
            [_order_no_ship("PS030", "待发货", 100),
             _order_no_ship("PS031", "已发货", 200)],
            headers=WX_HEADERS_NO_SHIP)
        self._recalculate()
        ids = self._profit_order_ids()
        self.assertNotIn("PS030", ids, "无商品发货列时 待发货订单应被排除")
        self.assertIn("PS031", ids, "无商品发货列时 已发货订单应进入核算")

    def test_05_ship_unshipped_overrides_completed(self):
        """订单状态=已完成 但 商品发货=未发货 → 排除（ship_status 优先）"""
        self._import_orders([
            _order_with_ship("PS040", "已完成", 500, "未发货"),
            _order_with_ship("PS041", "已完成", 600, "已发货"),
        ])
        self._recalculate()
        ids = self._profit_order_ids()
        self.assertNotIn("PS040", ids,
                         "订单状态=已完成但商品发货=未发货 应被排除")
        self.assertIn("PS041", ids, "订单状态=已完成且商品发货=已发货 应进入核算")

    def test_06_is_calculable_status_unit(self):
        """is_calculable_status 单元测试：ship_status 维度"""
        self.assertFalse(is_calculable_status("completed", "unshipped"),
                         "ship_status=unshipped 应排除，即使 order_status=completed")
        self.assertTrue(is_calculable_status("completed", "shipped"),
                        "ship_status=shipped 且 order_status=completed 应可核算")
        self.assertTrue(is_calculable_status("completed", "partial_shipped"),
                        "ship_status=partial_shipped 应可核算")
        self.assertTrue(is_calculable_status("completed", None),
                        "ship_status 缺失时按 order_status 判断")
        self.assertFalse(is_calculable_status("pending_shipment", None),
                         "ship_status 缺失且 order_status=未发货 应排除")
        self.assertFalse(is_calculable_status(None, None),
                         "两者均缺失应排除")


if __name__ == "__main__":
    unittest.main()

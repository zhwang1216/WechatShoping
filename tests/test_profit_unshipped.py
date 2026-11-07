"""利润核算排除未发货订单 验收测试

覆盖 spec.md 7 个验收场景 + 红线约束：
  test_01: 未发货订单排除
  test_02: 状态变更后自动纳入
  test_03: 已发货订单保持原行为
  test_04: 状态缺失待核对排除
  test_05: 未发货匹配成本仍排除
  test_06: 月份/店铺/分页/导出统一过滤
  test_07: 退款逻辑不受影响
  test_08: 红线约束
"""
import io
import unittest
from collections import Counter
from decimal import Decimal

from openpyxl import Workbook, load_workbook

from app.profit_service import is_calculable_status
from tests.http_client import TestEnv


WX_HEADERS = ["订单号", "订单下单时间", "订单状态", "订单实际支付金额",
              "订单实际收款金额", "商品数量", "商品标题", "sku编码",
              "技术服务费", "商品已退款金额", "商品售后"]


def _make_order_xlsx(rows):
    """构造订单 xlsx；rows 为元组列表，每行对应 WX_HEADERS"""
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.append(WX_HEADERS)
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _order(order_id, status, received, qty=1, refund=0, after_sale=None):
    return (order_id, "2026-03-01 10:00:00", status, received, received,
            qty, "商品", "SKU-X", 3, refund, after_sale)


def _order_at(order_id, date, status, received, qty=1, refund=0,
              after_sale=None):
    return (order_id, date, status, received, received,
            qty, "商品", "SKU-X", 3, refund, after_sale)


class TestProfitUnshipped(unittest.TestCase):
    """利润核算排除未发货订单 验收测试"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            "pu_member", "PuMember123456", store_name="PU测试店")
        cls.store_id = cls.member.get("/api/auth/me")[1]["data"]["store_id"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _import_orders(self, rows):
        """导入订单并确认"""
        xlsx = _make_order_xlsx(rows)
        st, r = self.member.upload_confirm(
            "/api/import/wechat-orders?store_id=" + self.store_id,
            "test.xlsx", xlsx, target_store_id=self.store_id)
        self.assertEqual(st, 200, f"导入失败: {r}")
        return r

    def _recalculate(self):
        st, r = self.admin.post(
            "/api/profit/recalculate", {"store_id": self.store_id})
        self.assertEqual(st, 200, f"重算失败: {r}")
        return r

    def _profit_list(self, page=1, size=100):
        st, r = self.member.get(
            f"/api/profit?page={page}&size={size}")
        self.assertEqual(st, 200, f"利润列表失败: {r}")
        return r

    def _summary(self):
        st, r = self.member.get("/api/summary")
        self.assertEqual(st, 200, f"汇总失败: {r}")
        return r

    def _order_ids_in_profit(self):
        """获取利润明细中所有订单号"""
        ids = set()
        page = 1
        while True:
            r = self._profit_list(page=page, size=100)
            data = r.get("data", {})
            items = data.get("results", data.get("items", data.get("records", [])))
            for item in items:
                ids.add(item.get("order_id"))
            if len(items) < 100:
                break
            page += 1
        return ids

    def _order_id_counts_in_profit(self):
        """获取利润明细中订单号出现次数（用于重复计算检测）"""
        counts = Counter()
        page = 1
        while True:
            r = self._profit_list(page=page, size=100)
            data = r.get("data", {})
            items = data.get("results", data.get("items", data.get("records", [])))
            for item in items:
                counts[item.get("order_id")] += 1
            if len(items) < 100:
                break
            page += 1
        return counts

    def _profit_results(self, query=""):
        """带查询参数获取利润明细结果集"""
        st, r = self.member.get(f"/api/profit?size=100{query}")
        self.assertEqual(st, 200, f"利润列表失败: {r}")
        return r.get("data", {}).get("results", [])

    def test_01_unshipped_excluded(self):
        """场景一：未发货订单排除验收"""
        self._import_orders([
            _order("PU001", "待发货", 100),
            _order("PU002", "已发货", 200),
            _order("PU003", "已支付", 50),
            _order("PU004", "待付款", 80),
        ])
        self._recalculate()
        profit_ids = self._order_ids_in_profit()
        self.assertIn("PU002", profit_ids, "已发货订单应在利润核算中")
        self.assertNotIn("PU001", profit_ids, "待发货订单不应在利润核算中")
        self.assertNotIn("PU003", profit_ids, "已支付订单不应在利润核算中")
        self.assertNotIn("PU004", profit_ids, "待付款订单不应在利润核算中")

    def test_02_status_change_included(self):
        """场景二：状态变更后自动纳入验收"""
        self._import_orders([_order("PU010", "待发货", 150)])
        self._recalculate()
        self.assertNotIn("PU010", self._order_ids_in_profit(),
                         "未发货时不应在利润核算中")
        self._import_orders([_order("PU010", "已发货", 150)])
        self._recalculate()
        self.assertIn("PU010", self._order_ids_in_profit(),
                      "更新为已发货后应在利润核算中")

    def test_03_shipped_unchanged(self):
        """场景三：已发货订单保持原行为验收"""
        self._import_orders([
            _order("PU020", "已发货", 100),
            _order("PU021", "已完成", 200),
        ])
        r = self._recalculate()
        profit_ids = self._order_ids_in_profit()
        self.assertIn("PU020", profit_ids)
        self.assertIn("PU021", profit_ids)

    def test_04_unknown_status_pending_review(self):
        """场景四：状态缺失待核对验收"""
        self._import_orders([
            _order("PU030", "无法识别的状态文本", 100),
            _order("PU031", "已发货", 200),
        ])
        self._recalculate()
        profit_ids = self._order_ids_in_profit()
        self.assertNotIn("PU030", profit_ids, "无法识别状态应排除")
        self.assertIn("PU031", profit_ids)
        st, r = self.member.get("/api/orders")
        self.assertEqual(st, 200)
        order_ids = {o.get("order_id") for o in r.get("data", {}).get("orders", [])}
        self.assertIn("PU030", order_ids, "非未发货类异常状态仍应导入并在普通订单页展示")

    def test_05_unshipped_with_cost_still_excluded(self):
        """场景五：未发货匹配成本仍排除验收"""
        self._import_orders([_order("PU040", "待发货", 100)])
        sup_headers = ["主订单编号", "订单编号", "订单创建时间", "商品数量",
                        "成本", "总金额", "快递费合计", "快递费", "泡沫耗材费",
                        "加收", "燃油费", "订单运费", "平台补贴", "订单实收款",
                        "sku编码", "快递单号", "商品标题"]
        wb = Workbook()
        ws = wb.active
        ws.title = "供应商"
        ws.append(sup_headers)
        ws.append(["PU040", None, "2026-03-01 10:05:00", 1, 30, 45,
                   None, None, None, None, None, None, None, 100,
                   "SKU-X", None, "商品"])
        buf = io.BytesIO()
        wb.save(buf)
        st, r = self.member.upload_confirm(
            "/api/import/supplier-costs?store_id=" + self.store_id,
            "sup.xlsx", buf.getvalue(), target_store_id=self.store_id)
        self.assertEqual(st, 200, f"成本导入失败: {r}")
        self._recalculate()
        self.assertNotIn("PU040", self._order_ids_in_profit(),
                         "未发货即使匹配成本也不应进入利润核算")

    def test_06_filter_consistency(self):
        """场景六：月份/店铺/分页/导出统一过滤验收"""
        self._import_orders([
            _order("PU050", "待发货", 100),
            _order("PU051", "已发货", 200),
        ])
        self._recalculate()
        profit_ids = self._order_ids_in_profit()
        self.assertNotIn("PU050", profit_ids)
        self.assertIn("PU051", profit_ids)
        summary = self._summary()
        self.assertIsNotNone(summary.get("data"))

    def test_07_refund_logic_unchanged(self):
        """场景七：退款逻辑不受影响验收"""
        self._import_orders([
            _order("PU060", "已完成", 100, refund=30, after_sale="部分退款"),
            _order("PU061", "已完成", 80, refund=80, after_sale="已全额退款"),
        ])
        self._recalculate()
        profit_ids = self._order_ids_in_profit()
        self.assertIn("PU060", profit_ids, "部分退款订单应在利润核算中")
        self.assertIn("PU061", profit_ids, "全额退款订单应在利润核算中")

    def test_08_redline_constraints(self):
        """红线约束验证"""
        self._import_orders([_order("PU070", "待发货", 100)])
        self._recalculate()
        st, r = self.member.get(f"/api/orders?store_id={self.store_id}")
        self.assertEqual(st, 200)
        order_ids = {o.get("order_id") for o in r.get("data", {}).get("orders", [])}
        self.assertNotIn("PU070", order_ids, "未发货订单导入时应自动剔除")
        self.assertNotIn("PU070", self._order_ids_in_profit(),
                         "红线: 未发货订单不在利润核算中")

    def test_09_trailing_space_status(self):
        """尾随空格状态文本识别验收（spec 5.1.1.1a）"""
        self._import_orders([
            _order("PU100", "待发货 ", 100),
            _order("PU101", "已发货", 200),
        ])
        self._recalculate()
        profit_ids = self._order_ids_in_profit()
        self.assertNotIn("PU100", profit_ids, "带尾随空格的待发货应被识别并排除")
        self.assertIn("PU101", profit_ids)

    def test_10_invisible_char_status(self):
        """不可见字符状态文本识别验收（spec 5.1.1.5a）"""
        self._import_orders([
            _order("PU110", "\u200b待发货\u200b", 100),
            _order("PU111", "\u00a0已发货\u00a0", 200),
        ])
        self._recalculate()
        profit_ids = self._order_ids_in_profit()
        self.assertNotIn("PU110", profit_ids, "含不可见字符的待发货应被识别并排除")
        self.assertIn("PU111", profit_ids, "含不可见字符的已发货应被识别并纳入")

    def test_11_amount_excluded_from_summary(self):
        """未发货订单金额不计入汇总验收（spec 5.2.1.2a）"""
        self._import_orders([_order("PU120", "已发货", 200)])
        self._recalculate()
        s1 = self._summary()["data"]
        self._import_orders([_order("PU121", "待发货", 100)])
        self._recalculate()
        s2 = self._summary()["data"]
        self.assertEqual(
            s2.get("product_sales_amount"), s1.get("product_sales_amount"),
            "未发货订单的 100 元不应计入商品销售收入汇总")
        self.assertNotIn("PU121", self._order_ids_in_profit())

    def test_12_no_duplicate_on_status_change(self):
        """状态变更后不重复计算验收（spec 5.3.1.2a）"""
        self._import_orders([_order("PU130", "待发货", 150)])
        self._recalculate()
        self._import_orders([_order("PU130", "已发货", 150)])
        self._recalculate()
        counts = self._order_id_counts_in_profit()
        self.assertEqual(counts.get("PU130", 0), 1,
                         "状态变更后订单在利润明细中应仅出现一次")

    def test_13_export_consistency(self):
        """导出口径与页面一致验收（spec 5.4.1.2a）"""
        self._import_orders([
            _order("PU140", "待发货", 100),
            _order("PU141", "已发货", 200),
        ])
        self._recalculate()
        profit_ids = self._order_ids_in_profit()
        self.assertNotIn("PU140", profit_ids)
        self.assertIn("PU141", profit_ids)
        st, body, _ = self.member.get_binary(
            "/api/profit/export?month=2026-03")
        self.assertEqual(st, 200, f"导出失败: {st}")
        wb = load_workbook(io.BytesIO(body))
        ws = wb.active
        exported_ids = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and str(row[0]) != "合计":
                exported_ids.add(str(row[0]))
        self.assertNotIn("PU140", exported_ids, "导出文件不应含未发货订单")
        self.assertIn("PU141", exported_ids, "导出文件应含已发货订单")

    def test_14_missing_status_field_excluded(self):
        """状态字段缺失保守排除验收（spec 6.1.4）"""
        self.assertFalse(is_calculable_status(None),
                         "order_status=None 应被 is_calculable_status 排除")
        self._import_orders([
            _order_at("PU150", "2026-03-01 10:00:00", "", 100),
            _order("PU151", "已发货", 200),
        ])
        self._recalculate()
        profit_ids = self._order_ids_in_profit()
        self.assertNotIn("PU150", profit_ids, "空状态订单应被排除")
        self.assertIn("PU151", profit_ids)

    def test_15_month_switch_consistency(self):
        """月份切换口径一致验收（spec 5.5.1.1a）"""
        self._import_orders([
            _order_at("PU160", "2026-03-15 10:00:00", "待发货", 100),
            _order_at("PU161", "2026-03-15 10:00:00", "已发货", 200),
            _order_at("PU162", "2026-04-15 10:00:00", "已发货", 300),
        ])
        self._recalculate()
        march_ids = {item.get("order_id")
                     for item in self._profit_results("&month=2026-03")}
        self.assertNotIn("PU160", march_ids, "3月未发货订单应排除")
        self.assertIn("PU161", march_ids, "3月已发货订单应纳入")
        self.assertNotIn("PU162", march_ids, "4月订单不应在3月筛选中")
        april_ids = {item.get("order_id")
                     for item in self._profit_results("&month=2026-04")}
        self.assertIn("PU162", april_ids, "4月已发货订单应纳入")
        self.assertNotIn("PU160", april_ids, "未发货订单在任何月份都应排除")


if __name__ == "__main__":
    unittest.main()

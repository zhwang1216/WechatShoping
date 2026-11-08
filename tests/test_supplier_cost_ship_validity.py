"""订单成本导入：未发货有效性校验验收"""
import io
import json
import os
import unittest
from decimal import Decimal

from openpyxl import Workbook

from app.cost_validity import (
    INVALID_REASON_UNSHIPPED,
    build_orders_lookup,
    evaluate_supplier_cost_validity,
    is_effective_supplier_cost,
)
from app.profit_service import is_unshipped_order
from tests.http_client import TestEnv

WX_HEADERS = [
    "订单号", "订单下单时间", "订单状态", "订单实际支付金额",
    "订单实际收款金额", "商品数量", "商品标题", "sku编码",
    "技术服务费", "商品已退款金额", "商品售后",
]

SUP_HEADERS = [
    "主订单编号", "订单编号", "订单创建时间", "商品数量",
    "成本", "总金额", "快递费合计", "快递费", "泡沫耗材费",
    "加收", "燃油费", "订单运费", "平台补贴", "订单实收款",
    "sku编码", "快递单号", "商品标题",
]


def _wx_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.append(WX_HEADERS)
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _sup_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.append(SUP_HEADERS)
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _order_row(oid, status, recv=100):
    return (oid, "2026-03-01 10:00:00", status, recv, recv,
            1, "商品", "SKU", 3, 0, None)


def _sup_row(main, total=45, supply=30, courier=15):
    return (main, None, "2026-03-01 10:05:00", 1, supply, total,
            courier, courier, None, None, None, None, None, 100,
            "SKU", "SF001", "商品")


class TestCostValidityUnit(unittest.TestCase):
  def test_unshipped_invalid(self):
    orders = build_orders_lookup([{
      "order_id": "U1",
      "order_status": "pending_shipment",
      "ship_status": None,
    }])
    cost = {"main_order_id": "U1", "batch_id": "b1"}
    bv = evaluate_supplier_cost_validity(cost, orders)
    self.assertEqual(bv["validity_status"], "invalid")
    self.assertEqual(bv["invalid_reason"], INVALID_REASON_UNSHIPPED)

  def test_ship_column_unshipped(self):
    orders = build_orders_lookup([{
      "order_id": "U2",
      "order_status": "completed",
      "ship_status": "unshipped",
    }])
    bv = evaluate_supplier_cost_validity(
        {"main_order_id": "U2"}, orders)
    self.assertEqual(bv["validity_status"], "invalid")
    self.assertTrue(is_unshipped_order("completed", "unshipped"))

  def test_no_order_pending(self):
    bv = evaluate_supplier_cost_validity(
        {"main_order_id": "NOPE"}, {})
    self.assertEqual(bv["validity_status"], "pending")

  def test_shipped_valid(self):
    orders = build_orders_lookup([{
      "order_id": "U3",
      "order_status": "shipped",
      "ship_status": "shipped",
    }])
    bv = evaluate_supplier_cost_validity({"main_order_id": "U3"}, orders)
    self.assertEqual(bv["validity_status"], "valid")
    self.assertTrue(is_effective_supplier_cost(
        {"main_order_id": "U3", "business_validity": bv}, orders))


class TestSupplierCostShipValidityHttp(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.member = cls.env.create_user_client(
            "scv_member", "ScvMember123456", store_name="成本有效性店")
        cls.store_id = cls.member.get("/api/auth/me")[1]["data"]["store_id"]
        cls.admin = cls.env.admin

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _import_orders(self, rows):
        st, r = self.member.upload_confirm(
            f"/api/import/wechat-orders?store_id={self.store_id}",
            "o.xlsx", _wx_xlsx(rows), target_store_id=self.store_id)
        self.assertEqual(st, 200, r)

    def _import_costs(self, rows, force=False, filename="c.xlsx"):
        st, r = self.member.upload(
            f"/api/import/supplier-costs?store_id={self.store_id}",
            filename, _sup_xlsx(rows))
        self.assertEqual(st, 200, r)
        body = {
            "upload_id": r["data"]["upload_id"],
            "target_store_id": self.store_id,
            "force": force,
        }
        st, r = self.member.post("/api/import/confirm", body)
        self.assertEqual(st, 200, r)
        return r

    def _supplier_stats(self, keyword=None):
        path = "/api/supplier-costs?page=1&size=50"
        if keyword:
            path += f"&keyword={keyword}"
        st, r = self.member.get(path)
        self.assertEqual(st, 200, r)
        return r.get("data", {}).get("statistics", {})

    def _cost_records(self):
        path = os.path.join(
            self.env.data_dir, "stores", self.store_id, "data",
            "supplier_costs.json")
        with open(path, encoding="utf-8") as f:
            return json.load(f)["records"]

    def test_pending_no_order_not_in_stats(self):
        self._import_costs([_sup_row("SCV_NONE", total=99)])
        stats = self._supplier_stats("SCV_NONE")
        self.assertEqual(stats.get("cost_line_count"), 0)
        recs = [r for r in self._cost_records() if r.get("main_order_id") == "SCV_NONE"]
        self.assertEqual(len(recs), 1)
        self.assertEqual(recs[0]["business_validity"]["validity_status"], "pending")

    def test_valid_shipped_in_stats(self):
        self._import_orders([_order_row("SCV_OK", "已发货", 200)])
        self._import_costs([_sup_row("SCV_OK", total=45, supply=30, courier=15)])
        stats = self._supplier_stats("SCV_OK")
        self.assertEqual(stats.get("cost_line_count"), 1)
        self.assertEqual(Decimal(stats.get("total_amount_sum")), Decimal("45.00"))
        self.assertEqual(Decimal(stats.get("product_supply_cost_sum")), Decimal("30.00"))
        self.assertEqual(Decimal(stats.get("courier_fee_sum")), Decimal("15.00"))

    def test_unshipped_invalid_not_in_stats(self):
        self._import_orders([_order_row("SCV_BAD", "已发货", 100)])
        self._import_costs([_sup_row("SCV_BAD")])
        path = os.path.join(
            self.env.data_dir, "stores", self.store_id, "data", "orders.json")
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        for line in data["records"]:
            if line.get("order_id") == "SCV_BAD":
                line["ship_status"] = "unshipped"
                line["source_ship_status"] = "未发货"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        from app.json_repository import JsonRepository
        from app.cost_validity import revalidate_supplier_costs_for_order_ids
        from app.operation_log import OperationLog
        store_root = os.path.join(
            self.env.data_dir, "stores", self.store_id)
        repo = JsonRepository(
            data_dir=os.path.join(store_root, "data"),
            backup_dir=os.path.join(store_root, "backups"))
        op = OperationLog(log_dir=os.path.join(store_root, "logs"))
        revalidate_supplier_costs_for_order_ids(
            repo, {"SCV_BAD"}, op_log=op, operator="test")
        stats = self._supplier_stats("SCV_BAD")
        self.assertEqual(stats.get("cost_line_count"), 0)
        rec = [r for r in self._cost_records() if r.get("main_order_id") == "SCV_BAD"][0]
        self.assertEqual(rec["business_validity"]["validity_status"], "invalid")

    def test_status_restore_to_shipped(self):
        self._import_orders([_order_row("SCV_REV", "已发货")])
        self._import_costs([_sup_row("SCV_REV")])
        path = os.path.join(
            self.env.data_dir, "stores", self.store_id, "data", "orders.json")
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        for line in data["records"]:
            if line.get("order_id") == "SCV_REV":
                line["ship_status"] = "unshipped"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        from app.json_repository import JsonRepository
        from app.cost_validity import revalidate_supplier_costs_for_order_ids
        from app.operation_log import OperationLog
        store_root = os.path.join(
            self.env.data_dir, "stores", self.store_id)
        repo = JsonRepository(
            data_dir=os.path.join(store_root, "data"),
            backup_dir=os.path.join(store_root, "backups"))
        op = OperationLog(log_dir=os.path.join(store_root, "logs"))
        revalidate_supplier_costs_for_order_ids(
            repo, {"SCV_REV"}, op_log=op, operator="test")
        self.assertEqual(self._supplier_stats("SCV_REV").get("cost_line_count"), 0)
        for line in data["records"]:
            if line.get("order_id") == "SCV_REV":
                line["ship_status"] = "shipped"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        revalidate_supplier_costs_for_order_ids(
            repo, {"SCV_REV"}, op_log=op, operator="test")
        self.assertEqual(self._supplier_stats("SCV_REV").get("cost_line_count"), 1)

    def test_duplicate_import_not_double_stats(self):
        self._import_orders([_order_row("SCV_DD", "已发货")])
        self._import_costs([_sup_row("SCV_DD")])
        self._import_costs([_sup_row("SCV_DD")], force=True, filename="c2.xlsx")
        stats = self._supplier_stats("SCV_DD")
        self.assertEqual(stats.get("cost_line_count"), 1)
        self.assertEqual(Decimal(stats.get("total_amount_sum")), Decimal("45.00"))

    def test_export_excludes_invalid(self):
        self._import_orders([_order_row("SCV_EX", "已发货")])
        self._import_costs([_sup_row("SCV_EX"), _sup_row("SCV_NOP")])
        st, body, _ = self.member.get_binary("/api/supplier-costs/export")
        self.assertEqual(st, 200)
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(body))
        mains = {str(r[0]) for r in wb.active.iter_rows(min_row=2, values_only=True)}
        self.assertIn("SCV_EX", mains)
        self.assertNotIn("SCV_NOP", mains)


if __name__ == "__main__":
    unittest.main()

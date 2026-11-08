"""阶段6测试：未匹配供应商成本记录手动录入综合成本

验收场景：
1. 未匹配记录成功录入综合成本；
2. 空值/负数/非法金额被拒绝；
3. 刷新/重启后金额仍存在；
4. 已匹配记录不能通过 API 修改；
5. 手动录入不改变匹配状态；
6. 后续导入覆盖冲突金额；
7. 无效导入不清空人工录入值；
8. 匹配成功后综合成本参与核算；
9. 未发货订单不进入利润核算；
10. 其他店铺数据隔离正常。
"""
import io
import unittest

from openpyxl import Workbook

from tests.http_client import TestEnv

WX_HEADERS = ["订单号", "订单下单时间", "订单状态", "订单实际支付金额",
              "订单实际收款金额", "商品数量", "商品标题", "sku编码"]

# (订单号, 时间, 状态, 支付, 实收, 数量, SKU)
WX_ROWS = [
    ("O5001", "2026-03-05 10:00:00", "已完成", 100, 100, 1, "SKU-A"),
    ("O5002", "2026-03-05 11:00:00", "已完成", 200, 200, 1, "SKU-B"),
    ("O5003", "2026-03-05 12:00:00", "待发货", 80, 80, 1, "SKU-C"),
]

SUP_HEADERS = ["主订单编号", "订单编号", "订单创建时间", "商品数量", "总金额",
               "订单实收款", "sku编码", "快递单号", "商品标题"]

# (主订单编号, 订单编号, 创建时间, 数量, 总金额, 实收款, SKU, 快递单号)
SUP_ROWS = [
    ("S8888", None, "2026-03-05 09:00:00", 1, 50, None, "SKU-X", None),
    ("O5002", None, "2026-03-05 11:05:00", 1, 180, 200, "SKU-B", None),
]

def make_wechat_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.append(WX_HEADERS)
    for oid, t, st, paid, recv, qty, sku in WX_ROWS:
        ws.append([oid, t, st, paid, recv, qty, f"商品{oid}", sku])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_supplier_xlsx(rows=None):
    rows = rows or SUP_ROWS
    wb = Workbook()
    ws = wb.active
    ws.title = "供应商成本"
    ws.append(SUP_HEADERS)
    for main, sub, t, qty, total, recv, sku, trk in rows:
        ws.append([main, sub, t, qty, total, recv, sku, trk, f"商品{sku}"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestManualCostEntry(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv(admin_user="tester")
        cls.admin = cls.env.admin
        cls.member_store = cls.env.member_store

        status, resp = cls.admin.upload(
            f"/api/import/wechat-orders?store_id={cls.member_store}",
            "订单.xlsx", make_wechat_xlsx())
        assert status == 200 and resp["code"] == 0, resp
        status, resp = cls.admin.post(
            "/api/import/confirm",
            {"upload_id": resp["data"]["upload_id"],
             "target_store_id": cls.member_store})
        assert status == 200 and resp["code"] == 0, resp

        status, resp = cls.admin.upload(
            f"/api/import/supplier-costs?store_id={cls.member_store}",
            "成本.xlsx", make_supplier_xlsx())
        assert status == 200 and resp["code"] == 0, resp
        status, resp = cls.admin.post(
            "/api/import/confirm",
            {"upload_id": resp["data"]["upload_id"],
             "target_store_id": cls.member_store})
        assert status == 200 and resp["code"] == 0, resp

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

    def _upload_static(self, path, filename, content):
        sep = "&" if "?" in path else "?"
        return self.admin.upload(f"{path}{sep}store_id={self.member_store}",
                                 filename, content)

    # ---------- 查询辅助 ----------
    def _run_recon(self):
        status, resp = self._post("/api/reconciliation/run", {})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        return resp["data"]

    def _orphan_supplier_cost(self, key="S8888"):
        status, resp = self._get("/api/data/supplier_costs")
        self.assertEqual(status, 200)
        rec = next(
            (r for r in resp["data"]["records"]
             if r.get("main_order_id") == key), None)
        self.assertIsNotNone(rec, f"应保留 {key} 供应商成本原始记录")
        self.assertEqual(
            (rec.get("business_validity") or {}).get("validity_status"),
            "pending",
            "无微信订单匹配的成本应为待核对")
        return rec

    def _supplier_record_id(self, key="S8888"):
        return self._orphan_supplier_cost(key)["record_id"]

    def _supplier_cost_record(self, record_id):
        status, resp = self._get("/api/data/supplier_costs")
        self.assertEqual(status, 200)
        for r in resp["data"]["records"]:
            if r.get("record_id") == record_id:
                return r
        raise AssertionError(f"未找到供应商记录: {record_id}")

    def _matched_supplier_id(self, main="O5002"):
        status, resp = self._get("/api/data/supplier_costs")
        self.assertEqual(status, 200)
        for r in resp["data"]["records"]:
            if r.get("main_order_id") == main:
                return r["record_id"]
        raise AssertionError(f"未找到供应商记录: {main}")

    # ---------- 1. 未匹配记录成功录入综合成本 ----------
    def test_01_unmatched_manual_cost_entry(self):
        self._run_recon()
        rec = self._orphan_supplier_cost("S8888")
        self.assertEqual(rec["cost_breakdown"]["total_amount"], "50.00")

        sid = rec["record_id"]
        status, resp = self._post("/api/reconciliation/manual-cost", {
            "supplier_record_id": sid,
            "total_amount": "45.00",
            "reason": "核实后录入实际综合成本",
        })
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        data = resp["data"]
        self.assertEqual(data["supplier_record_id"], sid)
        self.assertEqual(data["old_total_amount"], "50.00")
        self.assertEqual(data["new_total_amount"], "45.00")
        self.assertEqual(data["original_total_amount"], "50.00")
        self.assertEqual(data["manual_cost_modified_by"], "tester")
        self.assertTrue(data["manual_cost_modified_at"])

    # ---------- 2. 空值/负数/非法金额被拒绝 ----------
    def test_02_invalid_amounts_rejected(self):
        sid = self._supplier_record_id("S8888")
        for bad in ["", "-1.00", "abc", "1.001"]:
            status, resp = self._post("/api/reconciliation/manual-cost", {
                "supplier_record_id": sid,
                "total_amount": bad,
            })
            self.assertEqual(status, 400, f"非法金额 {bad!r} 应被拒绝")
            self.assertEqual(resp["code"], 3006)
        status, resp = self._post("/api/reconciliation/manual-cost", {
            "supplier_record_id": sid,
        })
        self.assertEqual(status, 400)

    # ---------- 3. 已匹配记录不能通过 API 修改 ----------
    def test_03_matched_record_rejected(self):
        sid = self._matched_supplier_id("O5002")
        status, resp = self._post("/api/reconciliation/manual-cost", {
            "supplier_record_id": sid,
            "total_amount": "99.00",
        })
        self.assertEqual(status, 403)
        self.assertEqual(resp["code"], 3004)

    # ---------- 4. 手动录入不改变匹配状态 ----------
    def test_04_match_status_unchanged(self):
        self._run_recon()
        rec = self._orphan_supplier_cost("S8888")
        self.assertEqual(rec["cost_breakdown"]["total_amount"], "45.00")

    # ---------- 5. 刷新/重启后金额仍存在 ----------
    def test_05_persistence_after_restart(self):
        self.env.restart()
        self._run_recon()
        orphan = self._orphan_supplier_cost("S8888")
        self.assertEqual(orphan["cost_breakdown"]["total_amount"], "45.00")
        rec = self._supplier_cost_record(orphan["record_id"])
        self.assertEqual(rec["manual_cost_source"], "manual")
        self.assertEqual(rec["manual_cost_modified_by"], "tester")

    # ---------- 6. 后续导入覆盖冲突金额 ----------
    def test_06_import_override(self):
        override_rows = [
            ("S8888", None, "2026-03-05 09:00:00", 1, 55, None, "SKU-X", None),
            ("O5002", None, "2026-03-05 11:05:00", 1, 180, 200, "SKU-B", None),
        ]
        status, resp = self._upload_static(
            "/api/import/supplier-costs", "成本_覆盖.xlsx",
            make_supplier_xlsx(override_rows))
        self.assertEqual(status, 200)
        status, resp = self._post("/api/import/confirm",
                                  {"upload_id": resp["data"]["upload_id"],
                                   "force": True})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)

        self._run_recon()
        orphan = self._orphan_supplier_cost("S8888")
        self.assertEqual(orphan["cost_breakdown"]["total_amount"], "55.00")
        rec = self._supplier_cost_record(orphan["record_id"])
        self.assertNotIn("manual_cost_source", rec)
        self.assertNotIn("manual_cost_modified_at", rec)
        self.assertNotIn("manual_cost_modified_by", rec)

    # ---------- 7. 无效导入不清空人工录入值 ----------
    def test_07_no_override_without_amount_column(self):
        sid = self._supplier_record_id("S8888")
        status, resp = self._post("/api/reconciliation/manual-cost", {
            "supplier_record_id": sid,
            "total_amount": "45.00",
            "reason": "重新录入",
        })
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)

        partial_rows = [
            ("O5002", None, "2026-03-05 11:05:00", 1, 180, 200, "SKU-B", None),
        ]
        status, resp = self._upload_static(
            "/api/import/supplier-costs", "成本_仅O5002.xlsx",
            make_supplier_xlsx(partial_rows))
        self.assertEqual(status, 200)
        status, resp = self._post("/api/import/confirm",
                                  {"upload_id": resp["data"]["upload_id"],
                                   "force": True})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)

        self._run_recon()
        orphan = self._orphan_supplier_cost("S8888")
        self.assertEqual(orphan["cost_breakdown"]["total_amount"], "45.00")
        rec = self._supplier_cost_record(orphan["record_id"])
        self.assertEqual(rec["manual_cost_source"], "manual")
        self.assertEqual(rec["manual_cost_modified_by"], "tester")

    # ---------- 8. 匹配成功后综合成本参与核算 ----------
    def test_08_matched_cost_in_profit(self):
        sid = self._supplier_record_id("S8888")
        status, resp = self._post("/api/reconciliation/manual-match", {
            "order_id": "O5001",
            "supplier_record_ids": [sid],
            "note": "人工绑定后核算利润",
        })
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        rec = resp["data"]
        self.assertEqual(rec["match_status"], "manual_confirmed")
        self.assertEqual(rec["cost"]["total_amount"], "45.00")

        status, resp = self._post("/api/profit/recalculate", {})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)

        status, resp = self._get("/api/profit?order_id=O5001")
        self.assertEqual(status, 200)
        rows = resp["data"]["results"]
        self.assertEqual(len(rows), 1)
        p = rows[0]
        self.assertEqual(p["profit_status"], "calculated")
        self.assertEqual(p["amounts"]["supplier_total_amount"], "45.00")
        self.assertEqual(p["amounts"]["composite_cost"], "45.00")
        self.assertEqual(p["amounts"]["operating_contribution_profit"], "55.00")

    # ---------- 9. 未发货订单不进入利润核算 ----------
    def test_09_unshipped_not_in_profit(self):
        status, resp = self._get("/api/profit?order_id=O5003")
        self.assertEqual(status, 200)
        rows = resp["data"]["results"]
        self.assertEqual(len(rows), 0)

    # ---------- 10. 其他店铺数据隔离正常 ----------
    def test_10_store_isolation(self):
        other = self.env.create_user_client(
            "iso_member", "Iso123456", role="member", store_name="隔离店铺")
        other_store = other.store_id

        sup_rows = [
            ("S7777", None, "2026-03-05 09:00:00", 1, 30, None, "SKU-Y", None),
        ]
        wb = Workbook()
        ws = wb.active
        ws.title = "供应商成本"
        ws.append(SUP_HEADERS)
        for main, sub, t, qty, total, recv, sku, trk in sup_rows:
            ws.append([main, sub, t, qty, total, recv, sku, trk, f"商品{sku}"])
        buf = io.BytesIO()
        wb.save(buf)
        content = buf.getvalue()

        status, resp = other.upload(
            f"/api/import/supplier-costs?store_id={other_store}",
            "隔离成本.xlsx", content)
        self.assertEqual(status, 200)
        status, resp = other.post("/api/import/confirm", {
            "upload_id": resp["data"]["upload_id"],
            "target_store_id": other_store})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)

        status, resp = other.post("/api/reconciliation/run", {
            "target_store_id": other_store})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)

        status, resp = other.get(
            f"/api/data/supplier_costs?target_store_id={other_store}")
        self.assertEqual(status, 200)
        orphan = next(
            (r for r in resp["data"]["records"]
             if r.get("main_order_id") == "S7777"), None)
        self.assertIsNotNone(orphan)
        other_sid = orphan["record_id"]

        status, resp = other.post("/api/reconciliation/manual-cost", {
            "supplier_record_id": other_sid,
            "total_amount": "25.00",
            "target_store_id": other_store,
        })
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        self.assertEqual(resp["data"]["new_total_amount"], "25.00")

        status, resp = self._post("/api/reconciliation/manual-cost", {
            "supplier_record_id": other_sid,
            "total_amount": "99.00",
        })
        self.assertEqual(status, 404)


if __name__ == "__main__":
    unittest.main()

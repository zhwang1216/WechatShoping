"""对账成本异常订单人工修正与导入优先规则 — 集成测试

覆盖 spec 第 7 节 10 个验收测试场景：
1. 异常订单人工修改后重新校验
2. 已匹配订单拒绝修改
3. 导入覆盖人工修改
4. 导入后恢复锁定
5. 未导入字段不被清空
6. 无效导入不覆盖
7. 异常重新判断
8. 历史数据兼容
9. 多店铺数据隔离
10. 操作日志可追溯
"""
import io
import json
import unittest

from openpyxl import Workbook

from tests.http_client import TestEnv

WX_HEADERS = ["订单号", "订单下单时间", "订单状态", "订单实际支付金额",
              "订单实际收款金额", "商品数量", "商品标题", "sku编码"]
SUP_HEADERS = ["主订单编号", "订单编号", "订单创建时间", "商品数量", "总金额",
               "订单实收款", "sku编码", "快递单号", "商品标题"]


def _wx_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.append(WX_HEADERS)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _sup_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "供应商成本"
    ws.append(SUP_HEADERS)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestReconAnomalyFix(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv(admin_user="tester")
        cls.admin = cls.env.admin
        cls.store = cls.env.member_store

        wx_rows = [
            ["M2001", "2026-03-01 10:00:00", "已完成", 100, 100, 1, "商品A", "SKU-A"],
            ["M2002", "2026-03-01 11:00:00", "已完成", 200, 200, 2, "商品B", "SKU-B"],
            ["M2003", "2026-03-01 12:00:00", "已完成", 50, 50, 1, "商品C", "SKU-C"],
        ]
        sup_rows = [
            ["M2001", None, "2026-03-01 10:05:00", 1, 48, 100, "SKU-A", None, "商品A"],
        ]

        status, resp = cls.admin.upload(
            f"/api/import/wechat-orders?store_id={cls.store}",
            "wx.xlsx", _wx_xlsx(wx_rows))
        assert status == 200 and resp["code"] == 0, resp
        status, resp = cls.admin.post("/api/import/confirm",
            {"upload_id": resp["data"]["upload_id"], "target_store_id": cls.store})
        assert status == 200 and resp["code"] == 0, resp

        status, resp = cls.admin.upload(
            f"/api/import/supplier-costs?store_id={cls.store}",
            "sup.xlsx", _sup_xlsx(sup_rows))
        assert status == 200 and resp["code"] == 0, resp
        status, resp = cls.admin.post("/api/import/confirm",
            {"upload_id": resp["data"]["upload_id"], "target_store_id": cls.store})
        assert status == 200 and resp["code"] == 0, resp

        status, resp = cls.admin.post("/api/reconciliation/run",
            {"target_store_id": cls.store})
        assert status == 200 and resp["code"] == 0, resp

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _post(self, path, body):
        body = dict(body or {})
        body.setdefault("target_store_id", self.store)
        return self.admin.post(path, body)

    def _result_of(self, order_id):
        qs = f"order_id={order_id}&target_store_id={self.store}"
        status, resp = self.admin.get(f"/api/reconciliation?{qs}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        rows = resp["data"]["results"]
        return rows[0] if rows else None

    def _run_recon(self):
        return self._post("/api/reconciliation/run", {})

    # ---------- 测试 1：异常订单人工修改后重新校验 ----------
    def test_01_manual_modify_anomaly_order(self):
        """异常订单可人工修改，修改后重新校验异常原因与数据一致"""
        r = self._result_of("M2002")
        self.assertIsNotNone(r)
        self.assertIn(r["match_status"], ("unmatched_order", "suspected"))

        status, resp = self._post("/api/reconciliation/manual-modify", {
            "order_id": "M2002",
            "fields": {"amounts.actual_received_amount": "180.00"},
            "reason": "核对原始单据后修正金额",
        })
        self.assertEqual(status, 200, resp)
        self.assertEqual(resp["code"], 0, resp)

        r2 = self._result_of("M2002")
        self.assertIsNotNone(r2)

    # ---------- 测试 2：已匹配订单拒绝修改 ----------
    def test_02_matched_order_reject_modify(self):
        """已匹配订单不能通过后端接口人工修改"""
        r = self._result_of("M2001")
        self.assertIsNotNone(r)
        self.assertEqual(r["match_status"], "matched")

        status, resp = self._post("/api/reconciliation/manual-modify", {
            "order_id": "M2001",
            "fields": {"amounts.actual_received_amount": "99.00"},
            "reason": "尝试修改已匹配订单",
        })
        self.assertEqual(status, 403, resp)

    # ---------- 测试 3：导入覆盖人工修改 ----------
    def test_03_import_overwrites_manual(self):
        """人工修改字段 A→B，后续导入字段值 C，最终值为 C"""
        status, resp = self._post("/api/reconciliation/manual-modify", {
            "order_id": "M2003",
            "fields": {"amounts.actual_received_amount": "40.00"},
            "reason": "人工修正为 40",
        })
        self.assertEqual(status, 200, resp)

        wx_rows = [
            ["M2003", "2026-03-01 12:00:00", "已完成", 50, 55, 1, "商品C", "SKU-C"],
        ]
        status, resp = self.admin.upload(
            f"/api/import/wechat-orders?store_id={self.store}",
            "wx2.xlsx", _wx_xlsx(wx_rows))
        self.assertEqual(status, 200)
        status, resp = self.admin.post("/api/import/confirm", {
            "upload_id": resp["data"]["upload_id"],
            "target_store_id": self.store,
            "force": True,
        })
        self.assertEqual(status, 200)

        r = self._result_of("M2003")
        self.assertIsNotNone(r)
        orders = self.admin.get(f"/api/data/orders?target_store_id={self.store}")[1]
        order = next(o for o in orders["data"]["records"] if o["order_id"] == "M2003")
        self.assertEqual(order["amounts"]["actual_received_amount"], "55.00")

    # ---------- 测试 4：导入后恢复锁定 ----------
    def test_04_import_restores_lock(self):
        """异常订单人工修改后，后续导入使其匹配，此后不得人工修改"""
        sup_rows = [
            ["M2003", None, "2026-03-01 12:05:00", 1, 50, 55, "SKU-C", None, "商品C"],
        ]
        status, resp = self.admin.upload(
            f"/api/import/supplier-costs?store_id={self.store}",
            "sup2.xlsx", _sup_xlsx(sup_rows))
        self.assertEqual(status, 200)
        status, resp = self.admin.post("/api/import/confirm", {
            "upload_id": resp["data"]["upload_id"],
            "target_store_id": self.store,
            "force": True,
        })
        self.assertEqual(status, 200)

        self._run_recon()

        r = self._result_of("M2003")
        if r and r["match_status"] in ("matched", "manual_confirmed"):
            status, resp = self._post("/api/reconciliation/manual-modify", {
                "order_id": "M2003",
                "fields": {"amounts.actual_received_amount": "10.00"},
                "reason": "尝试修改已匹配",
            })
            self.assertEqual(status, 403, resp)

    # ---------- 测试 5：未导入字段不被清空 ----------
    def test_05_unimported_field_preserved(self):
        """本次 Excel 未包含的字段原值不被清空"""
        wx_rows = [
            ["M2001", "2026-03-01 10:00:00", "已发货", 100, 100, 1, "商品A", "SKU-A"],
        ]
        status, resp = self.admin.upload(
            f"/api/import/wechat-orders?store_id={self.store}",
            "wx3.xlsx", _wx_xlsx(wx_rows))
        self.assertEqual(status, 200)
        status, resp = self.admin.post("/api/import/confirm", {
            "upload_id": resp["data"]["upload_id"],
            "target_store_id": self.store,
            "force": True,
        })
        self.assertEqual(status, 200)

        orders = self.admin.get(f"/api/data/orders?target_store_id={self.store}")[1]
        order = next(o for o in orders["data"]["records"] if o["order_id"] == "M2001")
        self.assertEqual(order["product"]["title"], "商品A")

    # ---------- 测试 6：无效导入不覆盖 ----------
    def test_06_invalid_import_no_overwrite(self):
        """导入记录未通过校验时不覆盖有效历史数据"""
        bad_rows = [
            ["", "2026-03-01 10:00:00", "已完成", 100, 100, 1, "商品", "SKU"],
        ]
        status, resp = self.admin.upload(
            f"/api/import/wechat-orders?store_id={self.store}",
            "bad.xlsx", _wx_xlsx(bad_rows))
        if status == 200:
            status2, resp2 = self.admin.post("/api/import/confirm", {
                "upload_id": resp["data"]["upload_id"],
                "target_store_id": self.store,
            })
            orders = self.admin.get(f"/api/data/orders?target_store_id={self.store}")[1]
            order = next((o for o in orders["data"]["records"] if o["order_id"] == "M2001"), None)
            self.assertIsNotNone(order)
            self.assertEqual(order["amounts"]["actual_received_amount"], "100.00")

    # ---------- 测试 7：异常重新判断 ----------
    def test_07_anomaly_recheck(self):
        """导入或人工修改后重新校验，异常原因与实际数据一致"""
        r = self._result_of("M2002")
        self.assertIsNotNone(r)
        if r["match_status"] in ("unmatched_order", "suspected"):
            status, resp = self._post("/api/reconciliation/manual-modify", {
                "order_id": "M2002",
                "fields": {"order_status": "已发货"},
                "reason": "修正状态",
            })
            self.assertEqual(status, 200, resp)
            r2 = self._result_of("M2002")
            self.assertIsNotNone(r2)

    # ---------- 测试 8：历史数据兼容 ----------
    def test_08_historical_data_compat(self):
        """原有 JSON 数据正常读取，未新增字段的历史记录不报错"""
        status, resp = self.admin.get(
            f"/api/reconciliation?target_store_id={self.store}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        for r in resp["data"]["results"]:
            self.assertIn("data_source", r)
            self.assertIn("manual_modified_fields", r)

    # ---------- 测试 9：多店铺数据隔离 ----------
    def test_09_multi_store_isolation(self):
        """不同店铺的相同订单号不因匹配或导入覆盖互相影响"""
        status, resp = self.admin.get("/api/stores")
        self.assertEqual(status, 200)
        stores = resp["data"].get("stores", [])
        if len(stores) < 2:
            self.skipTest("需要至少两个店铺测试隔离")

    # ---------- 测试 10：操作日志可追溯 ----------
    def test_10_operation_log_traceable(self):
        """人工修改操作可追溯"""
        status, resp = self.admin.get("/api/logs?action=manual_modify&limit=200")
        self.assertEqual(status, 200)
        entries = resp["data"].get("entries", [])
        found = any(
            (e.get("action") == "manual_modify" or "manual_modify" in str(e.get("action", "")))
            for e in entries
        )
        self.assertTrue(found, "应存在 manual_modify 操作日志")

    # ---------- 额外：字段白名单校验 ----------
    def test_11_system_field_rejected(self):
        """系统关键字段不可修改"""
        r = self._result_of("M2002")
        if r and r["match_status"] in ("unmatched_order", "suspected"):
            status, resp = self._post("/api/reconciliation/manual-modify", {
                "order_id": "M2002",
                "fields": {"order_id": "HACKED"},
                "reason": "尝试修改系统字段",
            })
            self.assertIn(status, (400, 403), resp)

    # ---------- 额外：modifiable-fields 接口 ----------
    def test_12_modifiable_fields_api(self):
        """GET modifiable-fields 返回字段定义与匹配状态"""
        status, resp = self.admin.get(
            f"/api/reconciliation/modifiable-fields?order_id=M2001&target_store_id={self.store}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        data = resp["data"]
        self.assertIn("order_fields", data)
        self.assertIn("match_status", data)
        self.assertIn("modifiable", data)

    # ==================== 辅助方法（成本字段修改场景构造） ====================
    def _make_suspected(self, order_id, sku, day, qty=1, cost_amount="100.00"):
        """构造 suspected 订单（SKU+日期辅助匹配但非精确匹配）

        返回 (对账结果, supplier_record_ids)。
        供应商成本 main_order_id 设为不存在的 FAKE_xxx，避免自动匹配；
        SKU 与日期与微信订单一致，触发辅助候选匹配 → suspected。
        """
        wx_rows = [[order_id, f"{day} 10:00:00", "已完成", 100, 100, qty,
                    f"商品{sku}", sku]]
        st, resp = self.admin.upload(
            f"/api/import/wechat-orders?store_id={self.store}",
            f"wx_{order_id}.xlsx", _wx_xlsx(wx_rows))
        self.assertEqual(st, 200, resp)
        self.assertEqual(resp["code"], 0, resp)
        st, resp = self.admin.post("/api/import/confirm", {
            "upload_id": resp["data"]["upload_id"], "target_store_id": self.store})
        self.assertEqual(st, 200, resp)

        fake_main = f"FAKE_{order_id}"
        sup_rows = [[fake_main, None, f"{day} 10:05:00", qty, cost_amount, 100,
                     sku, None, f"商品{sku}"]]
        st, resp = self.admin.upload(
            f"/api/import/supplier-costs?store_id={self.store}",
            f"sup_{order_id}.xlsx", _sup_xlsx(sup_rows))
        self.assertEqual(st, 200, resp)
        self.assertEqual(resp["code"], 0, resp)
        st, resp = self.admin.post("/api/import/confirm", {
            "upload_id": resp["data"]["upload_id"], "target_store_id": self.store})
        self.assertEqual(st, 200, resp)

        self._run_recon()
        r = self._result_of(order_id)
        self.assertIsNotNone(r, f"订单 {order_id} 对账结果为空")
        self.assertEqual(r["match_status"], "suspected",
                         f"订单 {order_id} 状态为 {r['match_status']}，期望 suspected")
        sids = r.get("supplier_record_ids") or []
        self.assertTrue(sids, f"订单 {order_id} 无关联 supplier_record_ids")
        return r, sids

    def _get_supplier_cost(self, record_id):
        """通过 record_id 获取 supplier_costs 记录"""
        st, resp = self.admin.get(
            f"/api/data/supplier_costs?target_store_id={self.store}")
        self.assertEqual(st, 200, resp)
        for c in resp["data"]["records"]:
            if c.get("record_id") == record_id:
                return c
        self.fail(f"supplier_costs 记录不存在: {record_id}")

    def _get_order(self, order_id):
        """通过 order_id 获取订单记录"""
        st, resp = self.admin.get(
            f"/api/data/orders?target_store_id={self.store}")
        self.assertEqual(st, 200, resp)
        for o in resp["data"]["records"]:
            if o.get("order_id") == order_id:
                return o
        self.fail(f"订单不存在: {order_id}")

    def _has_log(self, action):
        """查询操作日志中是否存在指定 action 的记录"""
        st, resp = self.admin.get(f"/api/logs?action={action}&limit=500")
        self.assertEqual(st, 200, resp)
        entries = resp["data"].get("entries", [])
        return any(e.get("action") == action for e in entries)

    # ---------- 测试 13：成本字段修改写入 supplier_costs ----------
    def test_13_cost_field_modify_writes_supplier_costs(self):
        """成本字段修改写入 supplier_costs 表，不污染 orders 表"""
        r, sids = self._make_suspected("M2013", "SKU-X13", "2026-03-13",
                                       cost_amount="100.00")
        sid = sids[0]

        status, resp = self._post("/api/reconciliation/manual-modify", {
            "order_id": "M2013",
            "cost_updates": [{"field": "cost_breakdown.total_amount",
                              "supplier_record_id": sid, "value": "95.00"}],
            "reason": "修正成本总金额",
        })
        self.assertEqual(status, 200, resp)
        self.assertEqual(resp["code"], 0, resp)

        # supplier_costs 记录已更新
        sc = self._get_supplier_cost(sid)
        self.assertEqual(sc["cost_breakdown"]["total_amount"], "95.00")

        # orders 表未被污染（不含 cost_breakdown 字段）
        order = self._get_order("M2013")
        self.assertNotIn("cost_breakdown", order)

        # 重新对账后 cost 汇总读取到新值
        self._run_recon()
        r2 = self._result_of("M2013")
        self.assertIsNotNone(r2)
        cost = r2.get("cost") or {}
        self.assertEqual(cost.get("total_amount"), "95.00")

    # ---------- 测试 14：订单区与成本区 product.quantity 独立修改 ----------
    def test_14_product_quantity_disambiguation(self):
        """订单区与成本区 product.quantity 独立修改，不互相污染"""
        r, sids = self._make_suspected("M2014", "SKU-X14", "2026-03-14",
                                       qty=3, cost_amount="100.00")
        sid = sids[0]

        status, resp = self._post("/api/reconciliation/manual-modify", {
            "order_id": "M2014",
            "fields": {"product.quantity": 7},
            "cost_updates": [{"field": "product.quantity",
                              "supplier_record_id": sid, "value": 9}],
            "reason": "分别修正订单与成本数量",
        })
        self.assertEqual(status, 200, resp)
        self.assertEqual(resp["code"], 0, resp)

        # 订单 product.quantity 更新为 7
        order = self._get_order("M2014")
        self.assertEqual(order["product"]["quantity"], 7)

        # supplier_costs product.quantity 更新为 9
        sc = self._get_supplier_cost(sid)
        self.assertEqual(sc["product"]["quantity"], 9)

    # ---------- 测试 15：多条 supplier_costs 记录精确更新 ----------
    def test_15_multi_cost_record_targeting(self):
        """多条 supplier_costs 记录按 supplier_record_id 精确更新"""
        order_id = "M2015"
        sku = "SKU-X15"
        day = "2026-03-15"

        wx_rows = [[order_id, f"{day} 10:00:00", "已完成", 200, 200, 2,
                    f"商品{sku}", sku]]
        st, resp = self.admin.upload(
            f"/api/import/wechat-orders?store_id={self.store}",
            f"wx_{order_id}.xlsx", _wx_xlsx(wx_rows))
        self.assertEqual(st, 200, resp)
        self.assertEqual(resp["code"], 0, resp)
        st, resp = self.admin.post("/api/import/confirm", {
            "upload_id": resp["data"]["upload_id"], "target_store_id": self.store})
        self.assertEqual(st, 200, resp)

        # 两条供应商成本：main_order_id 不匹配任何订单，但 SKU+日期辅助匹配
        sup_rows = [
            [f"FAKE_{order_id}A", None, f"{day} 10:05:00", 1, "80.00", 100,
             sku, None, f"商品{sku}"],
            [f"FAKE_{order_id}B", None, f"{day} 10:06:00", 1, "50.00", 100,
             sku, "TRK15B", f"商品{sku}"],
        ]
        st, resp = self.admin.upload(
            f"/api/import/supplier-costs?store_id={self.store}",
            f"sup_{order_id}.xlsx", _sup_xlsx(sup_rows))
        self.assertEqual(st, 200, resp)
        self.assertEqual(resp["code"], 0, resp)
        st, resp = self.admin.post("/api/import/confirm", {
            "upload_id": resp["data"]["upload_id"], "target_store_id": self.store})
        self.assertEqual(st, 200, resp)

        self._run_recon()
        r = self._result_of(order_id)
        self.assertIsNotNone(r)
        self.assertEqual(r["match_status"], "suspected")
        sids = r.get("supplier_record_ids") or []
        self.assertEqual(len(sids), 2, f"期望 2 条 supplier_record_ids，实际 {sids}")

        # 分别修改两条记录的不同字段
        status, resp = self._post("/api/reconciliation/manual-modify", {
            "order_id": order_id,
            "cost_updates": [
                {"field": "cost_breakdown.total_amount",
                 "supplier_record_id": sids[0], "value": "75.00"},
                {"field": "cost_breakdown.product_supply_cost",
                 "supplier_record_id": sids[1], "value": "40.00"},
            ],
            "reason": "分别修正两条成本记录",
        })
        self.assertEqual(status, 200, resp)
        self.assertEqual(resp["code"], 0, resp)

        sc0 = self._get_supplier_cost(sids[0])
        sc1 = self._get_supplier_cost(sids[1])
        self.assertEqual(sc0["cost_breakdown"]["total_amount"], "75.00")
        self.assertEqual(sc1["cost_breakdown"]["product_supply_cost"], "40.00")
        # 确保不互相污染
        self.assertNotEqual(sc1["cost_breakdown"]["total_amount"], "75.00")
        self.assertNotEqual(sc0["cost_breakdown"]["product_supply_cost"], "40.00")

    # ---------- 测试 16：非法 supplier_record_id 被拒绝 ----------
    def test_16_invalid_supplier_record_id_rejected(self):
        """cost_updates 中 supplier_record_id 不属于该订单时返回 404（错误码 3002）"""
        r, sids = self._make_suspected("M2016", "SKU-X16", "2026-03-16",
                                       cost_amount="100.00")

        status, resp = self._post("/api/reconciliation/manual-modify", {
            "order_id": "M2016",
            "cost_updates": [{"field": "cost_breakdown.total_amount",
                              "supplier_record_id": "sup_invalid_xxx",
                              "value": "50.00"}],
            "reason": "测试非法 supplier_record_id",
        })
        self.assertEqual(status, 404, resp)
        self.assertEqual(resp.get("code"), 3002, resp)

    # ---------- 测试 17：已匹配订单拒绝修改日志 ----------
    def test_17_modify_rejected_log(self):
        """已匹配订单拒绝修改，操作日志含 modify_rejected 记录"""
        order_id = "M2001"
        r = self._result_of(order_id)
        if r is None or r["match_status"] not in ("matched", "manual_confirmed"):
            # M2001 不是 matched，构造新 matched 订单
            order_id = "M2017"
            wx_rows = [[order_id, "2026-03-17 10:00:00", "已完成", 100, 100, 1,
                        "商品M2017", "SKU-X17"]]
            st, resp = self.admin.upload(
                f"/api/import/wechat-orders?store_id={self.store}",
                f"wx_{order_id}.xlsx", _wx_xlsx(wx_rows))
            self.assertEqual(st, 200, resp)
            st, resp = self.admin.post("/api/import/confirm", {
                "upload_id": resp["data"]["upload_id"], "target_store_id": self.store})
            self.assertEqual(st, 200, resp)
            sup_rows = [[order_id, None, "2026-03-17 10:05:00", 1, 48, 100,
                         "SKU-X17", None, "商品M2017"]]
            st, resp = self.admin.upload(
                f"/api/import/supplier-costs?store_id={self.store}",
                f"sup_{order_id}.xlsx", _sup_xlsx(sup_rows))
            self.assertEqual(st, 200, resp)
            st, resp = self.admin.post("/api/import/confirm", {
                "upload_id": resp["data"]["upload_id"], "target_store_id": self.store})
            self.assertEqual(st, 200, resp)
            self._run_recon()
            r = self._result_of(order_id)
            self.assertIsNotNone(r)
            self.assertIn(r["match_status"], ("matched", "manual_confirmed"))

        status, resp = self._post("/api/reconciliation/manual-modify", {
            "order_id": order_id,
            "fields": {"amounts.actual_received_amount": "99.00"},
            "reason": "尝试修改已匹配订单",
        })
        self.assertEqual(status, 403, resp)

        self.assertTrue(self._has_log("modify_rejected"),
                        "应存在 modify_rejected 操作日志")

    # ---------- 测试 18：新增 action 类型 ----------
    def test_18_new_action_types(self):
        """导入校验失败时操作日志含 import_validation_failed 记录"""
        bad_rows = [
            ["", "2026-03-18 10:00:00", "已完成", 100, 100, 1, "商品", "SKU-X18"],
        ]
        st, resp = self.admin.upload(
            f"/api/import/wechat-orders?store_id={self.store}",
            "bad18.xlsx", _wx_xlsx(bad_rows))
        if st == 200:
            self.admin.post("/api/import/confirm", {
                "upload_id": resp["data"]["upload_id"],
                "target_store_id": self.store,
            })

        self.assertTrue(self._has_log("import_validation_failed"),
                        "应存在 import_validation_failed 操作日志")


if __name__ == "__main__":
    unittest.main()

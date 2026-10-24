"""多店铺管理与隔离强化验收测试（16 项）

场景：管理员（无业务店铺）+ 承载成员 A（TestEnv 自动创建）+ member B（店铺 B）
+ member C（店铺 C）。
验证：
- 管理员可查看/聚合全部成员店铺数据，可按店铺筛选；
- 管理员可代任意店铺导入 Excel（上传目标店铺 + confirm 路由一致）；
- 管理员可重命名店铺（store_id 与数据归属不变）；
- 管理员可删除导入历史（仅历史记录）与撤销导入批次（含按月重算）；
- 普通成员严格锁定本店：查询/上传/写操作携带他人 store_id 一律 403。
"""
import io
import unittest
from decimal import Decimal

from openpyxl import load_workbook

from tests.http_client import TestEnv
from tests.test_auth_isolation import make_sup_xlsx, make_wx_xlsx

MONTH = "2026-03"
B_PREFIX = "B2E"
C_PREFIX = "C2E"
M_PREFIX = "M2E"

PASSWORD = "Store123456"


class TestMultiStoreAdmin(unittest.TestCase):
    """管理员全店铺数据管理 + 多店铺隔离强化（16 项验收）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.b = cls.env.create_user_client("mstore_b", PASSWORD,
                                           store_name="B 分店")
        cls.c = cls.env.create_user_client("mstore_c", PASSWORD,
                                           store_name="C 分店")
        for client, prefix, name in ((cls.b, B_PREFIX, "B"),
                                     (cls.c, C_PREFIX, "C")):
            st, _ = client.upload_confirm("/api/import/wechat-orders",
                                          f"{name}订单.xlsx",
                                          make_wx_xlsx(prefix))
            assert st == 200, f"{name} 订单导入失败"
            st, _ = client.upload_confirm("/api/import/supplier-costs",
                                          f"{name}成本.xlsx",
                                          make_sup_xlsx(prefix))
            assert st == 200, f"{name} 成本导入失败"
            client.post("/api/reconciliation/run", {"month": MONTH})
            client.post("/api/profit/recalculate", {"month": MONTH})
        cls.a_store = cls.env.member_store  # 管理员无店，A 为业务承载成员店铺
        cls.b_store = cls.b.get("/api/auth/me")[1]["data"]["store_id"]
        cls.c_store = cls.c.get("/api/auth/me")[1]["data"]["store_id"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 工具 ----------
    @classmethod
    def _store_name(cls, client, store_id):
        st, r = client.get("/api/stores")
        for s in r["data"]["stores"]:
            if s["store_id"] == store_id:
                return s["name"]
        return None

    def _c_latest_wechat_batch(self):
        st, r = self.admin.get(
            f"/api/import/batches?store_id={self.c_store}"
            f"&source_type=wechat_orders&page=1&page_size=20")
        self.assertEqual(st, 200)
        return r["data"]["batches"][0]["batch_id"]

    # ---------- 验收 1：店铺列表按身份裁剪 ----------
    def test_01_stores_listing_scoped(self):
        st, r = self.admin.get("/api/stores")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 3)
        by_id = {s["store_id"]: s for s in r["data"]["stores"]}
        self.assertEqual(set(by_id),
                         {self.a_store, self.b_store, self.c_store})
        self.assertIn("orders", by_id[self.b_store]["record_counts"])
        self.assertGreater(by_id[self.c_store]["record_counts"]["orders"], 0)
        # 成员仅返回本店
        st, r = self.b.get("/api/stores")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 1)
        self.assertEqual(r["data"]["stores"][0]["store_id"], self.b_store)
        # 业务页面已接入店铺筛选组件
        st, html, _ = self.admin.get_page("/orders")
        self.assertEqual(st, 200)
        self.assertIn('id="store-scope"', html)
        self.assertIn("store_scope.js", html)

    # ---------- 验收 2：管理员代任意店铺导入 ----------
    def test_02_admin_import_to_target_store(self):
        st, r = self.admin.upload(
            f"/api/import/wechat-orders?store_id={self.c_store}",
            "代C导入.xlsx", make_wx_xlsx(M_PREFIX))
        self.assertEqual(st, 200)
        upload_id = r["data"]["upload_id"]
        st, r = self.admin.post("/api/import/confirm", {
            "upload_id": upload_id, "target_store_id": self.c_store})
        self.assertEqual(st, 200)
        # 数据真实落 C 店铺
        st, r = self.c.get("/api/orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], 20)
        # 管理员无自有店铺：数据只落目标店铺 C，承载店 A 与 B 均不受影响
        st, r = self.admin.get(
            f"/api/orders?store_id={self.a_store}&page=1&size=100")
        self.assertEqual(r["data"]["total"], 0)
        # B 店铺不受影响
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], 10)

    # ---------- 验收 3：成员上传指定他店被拒 ----------
    def test_03_member_upload_target_forbidden(self):
        st, r = self.b.upload(
            f"/api/import/wechat-orders?store_id={self.c_store}",
            "越权上传.xlsx", make_wx_xlsx("H2E"))
        self.assertEqual(st, 403)
        st, r = self.b.post("/api/import/confirm", {
            "upload_id": "any", "target_store_id": self.c_store})
        self.assertEqual(st, 403)

    # ---------- 验收 4：管理员全部店铺聚合查询 ----------
    def test_04_admin_all_stores_orders_merged(self):
        st, r = self.admin.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        d = r["data"]
        self.assertEqual(d["store_scope"], "all")
        self.assertEqual(d["store_count"], 3)
        st, rb = self.b.get("/api/orders?page=1&size=100")
        st, rc = self.c.get("/api/orders?page=1&size=100")
        self.assertEqual(d["total"],
                         rb["data"]["total"] + rc["data"]["total"])
        for o in d["orders"]:
            self.assertIn(o.get("store_id"),
                          (self.a_store, self.b_store, self.c_store))
            self.assertTrue(o.get("store_name"))

    # ---------- 验收 5：管理员按单店筛选 ----------
    def test_05_admin_single_store_filter(self):
        st, r = self.admin.get(
            f"/api/orders?store_id={self.b_store}&page=1&size=100")
        self.assertEqual(st, 200)
        self.assertNotEqual(r["data"]["store_scope"], "all")
        self.assertGreater(r["data"]["total"], 0)
        for o in r["data"]["orders"]:
            self.assertEqual(o["store_id"], self.b_store)
        # 不存在的店铺 404
        st, r = self.admin.get("/api/orders?store_id=s_nope999")
        self.assertEqual(st, 404)

    # ---------- 验收 6：成员跨店查询被拒 ----------
    def test_06_member_cross_store_denied(self):
        st, r = self.b.get(f"/api/orders?store_id={self.c_store}")
        self.assertEqual(st, 403)
        st, r = self.b.get(f"/api/orders?store_id={self.b_store}")
        self.assertEqual(st, 200)

    # ---------- 验收 7：全部店铺汇总不重复计算 ----------
    def test_07_summary_aggregation_no_double_count(self):
        st, rb = self.b.get(f"/api/summary?month={MONTH}")
        st, rc = self.c.get(f"/api/summary?month={MONTH}")
        st, ra = self.admin.get(f"/api/summary?month={MONTH}")
        self.assertEqual(ra["data"]["store_scope"], "all")
        self.assertEqual(ra["data"]["store_count"], 3)
        self.assertEqual(ra["data"]["order_count"],
                         rb["data"]["order_count"] + rc["data"]["order_count"])
        merged_paid = (Decimal(rb["data"]["actual_paid_amount"])
                       + Decimal(rc["data"]["actual_paid_amount"]))
        self.assertEqual(Decimal(ra["data"]["actual_paid_amount"]), merged_paid)
        # 单店视图汇总 = 聚合视图（管理员按店筛选）
        st, rb2 = self.admin.get(
            f"/api/summary?month={MONTH}&store_id={self.b_store}")
        self.assertEqual(Decimal(rb2["data"]["actual_paid_amount"]),
                         Decimal(rb["data"]["actual_paid_amount"]))

    # ---------- 验收 8：全部店铺导出合并 ----------
    def test_08_admin_export_merged_excel(self):
        st, content, _ = self.admin.get_binary("/api/orders/export")
        self.assertEqual(st, 200)
        wb = load_workbook(io.BytesIO(content))
        all_rows = list(wb.active.iter_rows(values_only=True))
        st, content_b, _ = self.admin.get_binary(
            f"/api/orders/export?store_id={self.b_store}")
        self.assertEqual(st, 200)
        rows_b = list(load_workbook(io.BytesIO(content_b))
                      .active.iter_rows(values_only=True))
        st, content_c, _ = self.admin.get_binary(
            f"/api/orders/export?store_id={self.c_store}")
        self.assertEqual(st, 200)
        rows_c = list(load_workbook(io.BytesIO(content_c))
                      .active.iter_rows(values_only=True))
        self.assertEqual(len(all_rows), len(rows_b) + len(rows_c) - 1)
        self.assertEqual(all_rows[0], rows_b[0])  # 表头一致

    # ---------- 验收 9：店铺重命名 ----------
    def test_09_store_rename(self):
        st, r = self.c.get("/api/orders?page=1&size=100")
        orders_before = r["data"]["total"]
        # 成员无权重命名
        st, r = self.b.post(f"/api/stores/{self.b_store}/rename",
                            {"name": "B 自己改名"})
        self.assertEqual(st, 403)
        # 空名 400
        st, r = self.admin.post(f"/api/stores/{self.c_store}/rename",
                                {"name": "   "})
        self.assertEqual(st, 400)
        # 管理员重命名 C 店
        st, r = self.admin.post(f"/api/stores/{self.c_store}/rename",
                                {"name": "C 分店（新名称）"})
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["name"], "C 分店（新名称）")
        self.assertEqual(r["data"]["store_id"], self.c_store)
        # store_id 不变、数据不丢
        st, r = self.admin.get("/api/stores")
        c = [s for s in r["data"]["stores"]
             if s["store_id"] == self.c_store][0]
        self.assertEqual(c["name"], "C 分店（新名称）")
        self.assertGreater(c["record_counts"]["orders"], 0)
        st, r = self.c.get("/api/orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], orders_before)

    # ---------- 验收 10：跨店导入批次列表 ----------
    def test_10_batch_list_cross_store(self):
        st, r = self.admin.get("/api/import/batches?page=1&page_size=100")
        self.assertEqual(st, 200)
        sids = {b.get("store_id") for b in r["data"]["batches"]}
        self.assertIn(self.b_store, sids)
        self.assertIn(self.c_store, sids)
        # 成员批次列表 = 管理员视角本店批次集合
        st, r = self.admin.get(
            f"/api/import/batches?store_id={self.b_store}&page=1&page_size=100")
        admin_b_ids = {b["batch_id"] for b in r["data"]["batches"]}
        st, r = self.b.get("/api/import/batches?page=1&page_size=100")
        self.assertEqual({b["batch_id"] for b in r["data"]["batches"]},
                         admin_b_ids)

    # ---------- 验收 11：撤销影响预览 ----------
    def test_11_revoke_preview(self):
        batch_id = self._c_latest_wechat_batch()
        st, r = self.admin.get(
            f"/api/import/batches/{batch_id}/revoke-preview")
        self.assertEqual(st, 200)
        d = r["data"]
        self.assertEqual(d["store_id"], self.c_store)
        self.assertGreaterEqual(d["counts"].get("orders", 0), 1)
        self.assertGreaterEqual(d["counts"].get("raw_records", 0), 1)
        self.assertGreaterEqual(d["affected_order_count"], 1)
        self.assertIn(MONTH, d["affected_months"])
        # 成员无权预览他店批次
        st, r = self.b.get(
            f"/api/import/batches/{batch_id}/revoke-preview")
        self.assertEqual(st, 403)

    # ---------- 验收 12：仅删除导入历史 ----------
    def test_12_delete_batch_history_only(self):
        st, r = self.admin.get(
            f"/api/import/batches?store_id={self.c_store}"
            f"&source_type=supplier_cost&page=1&page_size=20")
        cost_batch = r["data"]["batches"][0]["batch_id"]
        # 成员无权
        st, r = self.b.post(
            f"/api/import/batches/{cost_batch}/delete-history",
            {"confirm": True})
        self.assertEqual(st, 403)
        # 缺二次确认 400
        st, r = self.admin.post(
            f"/api/import/batches/{cost_batch}/delete-history", {})
        self.assertEqual(st, 400)
        # 管理员确认删除：仅历史记录消失，业务数据保留
        st, r = self.admin.get(
            f"/api/supplier-costs?store_id={self.c_store}&page=1&page_size=100")
        costs_before = r["data"]["total"]
        st, r = self.admin.post(
            f"/api/import/batches/{cost_batch}/delete-history",
            {"confirm": True})
        self.assertEqual(st, 200)
        st, r = self.admin.get(
            f"/api/import/batches?store_id={self.c_store}"
            f"&source_type=supplier_cost&page=1&page_size=100")
        self.assertNotIn(cost_batch,
                         {b["batch_id"] for b in r["data"]["batches"]})
        st, r = self.admin.get(
            f"/api/supplier-costs?store_id={self.c_store}&page=1&page_size=100")
        self.assertEqual(r["data"]["total"], costs_before)

    # ---------- 验收 13：撤销导入批次（含按月重算） ----------
    def test_13_revoke_batch_full_flow(self):
        batch_id = self._c_latest_wechat_batch()  # M_PREFIX 批次
        st, r = self.admin.get(
            f"/api/orders?store_id={self.c_store}&page=1&size=100")
        c_before = r["data"]["total"]
        st, r = self.b.get("/api/orders?page=1&size=100")
        b_before = r["data"]["total"]
        # 成员无权撤销
        st, r = self.b.post(f"/api/import/batches/{batch_id}/revoke",
                            {"confirm": True})
        self.assertEqual(st, 403)
        # 缺二次确认 400
        st, r = self.admin.post(f"/api/import/batches/{batch_id}/revoke", {})
        self.assertEqual(st, 400)
        # 管理员撤销
        st, r = self.admin.post(f"/api/import/batches/{batch_id}/revoke",
                                {"confirm": True})
        self.assertEqual(st, 200)
        d = r["data"]
        self.assertEqual(d["store_id"], self.c_store)
        self.assertGreaterEqual(d["removed"].get("orders", 0), 1)
        self.assertGreaterEqual(len(d.get("recalculated", [])), 1)
        self.assertTrue(d.get("backup"))
        # C 店数据回到撤销前（仅剩 C_PREFIX 批次）
        st, r = self.admin.get(
            f"/api/orders?store_id={self.c_store}&page=1&size=100")
        self.assertEqual(r["data"]["total"], c_before - 10)
        # B 店数据完全不受影响
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], b_before)
        # 批次记录消失，C_PREFIX 批次仍在
        st, r = self.admin.get(
            f"/api/import/batches?store_id={self.c_store}"
            f"&source_type=wechat_orders&page=1&page_size=100")
        ids = {b["batch_id"] for b in r["data"]["batches"]}
        self.assertNotIn(batch_id, ids)
        self.assertTrue(ids, "C 店应仍保留 C_PREFIX 订单批次")

    # ---------- 验收 14：成员管理类 API 全部 403 ----------
    def test_14_member_admin_apis_denied(self):
        batch_id = self._c_latest_wechat_batch()
        checks = [
            ("delete-history",
             self.b.post(f"/api/import/batches/{batch_id}/delete-history",
                         {"confirm": True})),
            ("revoke",
             self.b.post(f"/api/import/batches/{batch_id}/revoke",
                         {"confirm": True})),
            ("revoke-preview",
             self.b.get(f"/api/import/batches/{batch_id}/revoke-preview")),
            ("rename",
             self.b.post(f"/api/stores/{self.c_store}/rename",
                         {"name": "越权重命名"})),
        ]
        for name, (st, _r) in checks:
            self.assertEqual(st, 403, f"{name} 应拒绝成员访问")

    # ---------- 验收 15：跨店操作日志合并 ----------
    def test_15_logs_cross_store(self):
        b_name = self._store_name(self.admin, self.b_store)
        c_name = self._store_name(self.admin, self.c_store)
        st, r = self.admin.get("/api/logs?page=1&page_size=100")
        self.assertEqual(st, 200)
        names = {e.get("store_name") for e in r["data"]["entries"]}
        self.assertIn(b_name, names)
        self.assertIn(c_name, names)
        # 按店过滤仅该店日志
        st, r = self.admin.get(
            f"/api/logs?store_id={self.b_store}&page=1&page_size=100")
        self.assertEqual(st, 200)
        self.assertTrue(r["data"]["entries"])
        for e in r["data"]["entries"]:
            self.assertEqual(e.get("store_id"), self.b_store)
        # 成员日志仅本店
        st, r = self.b.get("/api/logs?page=1&page_size=100")
        self.assertTrue(all(e.get("store_id") in (None, self.b_store)
                            for e in r["data"]["entries"]))

    # ---------- 验收 16：全部店铺批量对账/重算 ----------
    def test_16_run_all_stores(self):
        st, r = self.admin.post("/api/reconciliation/run",
                                {"target_store_id": "all", "month": MONTH})
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["store_count"], 3)
        self.assertEqual(len(r["data"]["stores"]), 3)
        st, r = self.admin.post("/api/profit/recalculate",
                                {"target_store_id": "all", "month": MONTH})
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["store_count"], 3)
        # 成员传 all 等价本店（单店结果直接返回，不返回 stores 数组）
        st, r = self.b.post("/api/reconciliation/run",
                            {"target_store_id": "all", "month": MONTH})
        self.assertEqual(st, 200)
        self.assertNotIn("stores", r["data"])
        # 成员指定他店 403
        st, r = self.b.post("/api/reconciliation/run",
                            {"target_store_id": self.c_store,
                             "month": MONTH})
        self.assertEqual(st, 403)


if __name__ == "__main__":
    unittest.main()

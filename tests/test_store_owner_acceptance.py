"""店铺归属者修改功能 — 增量验收测试（对照需求文档 8 项验收 + 操作日志）

覆盖 test_store_owner_fix.py 未覆盖项：
  test_09_audit_log_recorded: 操作日志记录归属变更（操作人/店铺/原归属/新归属/结果）
  test_10_member_a_api_access_revoked: A 重新登录后直接调 API 访问已转移店 → 403
  test_11_member_b_api_access_granted: B 重新登录后访问转移店 → 200
  test_12_history_refund_cost_intact: 退款、供应商成本前后一致
  test_13_nonexistent_store_transfer_fails: 不存在店铺转移 → 404，不虚报成功
  test_14_other_functions_regression: 创建/编辑/删除店铺、用户列表正常
  test_15_admin_views_all_stores: admin 仍可查看全部店铺
"""
import unittest

from tests.http_client import TestEnv


class TestStoreOwnerAcceptance(unittest.TestCase):
    """店铺归属者修改功能增量验收测试"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member_pwd = "Carrier123456"
        for uname, sname in [("owner_acc_a", "A-验收店"),
                             ("owner_acc_b", "B-验收店")]:
            if cls.env.get_user_id(uname) is None:
                cls.admin.post("/api/users", {
                    "username": uname, "password": cls.member_pwd,
                    "confirm_password": cls.member_pwd, "role": "member",
                    "create_store": True, "store_name": sname})

    def _get_members(self):
        resp = self.admin.get("/api/users")[1]
        users = resp.get("users", resp.get("data", {}).get("users", []))
        return [u for u in users if u.get("role") == "member"
                and u.get("status") == "active"]

    def _get_member(self, username):
        for m in self._get_members():
            if m["username"] == username:
                return m
        return None

    def test_09_audit_log_recorded(self):
        """操作日志记录归属变更（操作人/店铺/原归属/新归属/结果）"""
        admin = self.admin
        ma = self._get_member("owner_acc_a")
        mb = self._get_member("owner_acc_b")
        self.assertIsNotNone(ma)
        self.assertIsNotNone(mb)

        r = admin.post("/api/stores", {"name": "T9-日志测试店",
                                       "owner_user_id": ma["user_id"]})
        self.assertEqual(r[0], 200)
        store_id = r[1]["data"]["store_id"]

        try:
            admin.post(f"/api/stores/{store_id}/transfer-ownership",
                       {"target_user_id": mb["user_id"]})

            lr = admin.get("/api/logs?action=store.ownership_change&page=1&page_size=50")
            self.assertEqual(lr[0], 200, f"查询日志失败: {lr[1]}")
            entries = lr[1].get("data", {}).get("entries", [])
            matched = [e for e in entries
                       if e.get("store_id") == store_id
                       and e.get("result") == "success"]
            self.assertTrue(matched, "应存在归属变更成功日志")
            e = matched[-1]
            self.assertEqual(e.get("action"), "store.ownership_change")
            self.assertEqual(e.get("operator"), admin.username)
            detail = e.get("detail", {})
            self.assertEqual(detail.get("store_id"), store_id)
            prev = detail.get("previous_owner", {})
            new = detail.get("new_owner", {})
            self.assertEqual(prev.get("user_id"), ma["user_id"],
                             "日志应记录原归属者 A")
            self.assertEqual(new.get("user_id"), mb["user_id"],
                             "日志应记录新归属者 B")
        finally:
            admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})

    def test_10_member_a_api_access_revoked(self):
        """A 重新登录后直接调 API 访问已转移店 → 403（数据隔离）"""
        admin = self.admin
        ma = self._get_member("owner_acc_a")
        mb = self._get_member("owner_acc_b")
        self.assertIsNotNone(ma)
        self.assertIsNotNone(mb)

        r = admin.post("/api/stores", {"name": "T10-隔离测试店",
                                       "owner_user_id": ma["user_id"]})
        self.assertEqual(r[0], 200)
        store_id = r[1]["data"]["store_id"]

        try:
            admin.post(f"/api/stores/{store_id}/transfer-ownership",
                       {"target_user_id": mb["user_id"]})

            client_a = self.env.login_client("owner_acc_a", self.member_pwd)
            st, _ = client_a.get(f"/api/orders?store_id={store_id}")
            self.assertEqual(st, 403, "A 重新登录后访问已转移店应 403")

            st2, _ = client_a.get(f"/api/refund-orders?store_id={store_id}")
            self.assertEqual(st2, 403, "A 访问已转移店退款数据应 403")
        finally:
            admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})

    def test_11_member_b_api_access_granted(self):
        """B 重新登录后访问转移店 → 200（获得访问权限）"""
        admin = self.admin
        ma = self._get_member("owner_acc_a")
        mb = self._get_member("owner_acc_b")
        self.assertIsNotNone(ma)
        self.assertIsNotNone(mb)

        r = admin.post("/api/stores", {"name": "T11-授权测试店",
                                       "owner_user_id": ma["user_id"]})
        self.assertEqual(r[0], 200)
        store_id = r[1]["data"]["store_id"]

        try:
            admin.post(f"/api/stores/{store_id}/transfer-ownership",
                       {"target_user_id": mb["user_id"]})

            client_b = self.env.login_client("owner_acc_b", self.member_pwd)
            st, _ = client_b.get(f"/api/orders?store_id={store_id}")
            self.assertEqual(st, 200, "B 重新登录后访问转移店应 200")
        finally:
            admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})

    def test_12_history_refund_cost_intact(self):
        """历史数据完整 — 退款、供应商成本前后一致"""
        admin = self.admin
        ma = self._get_member("owner_acc_a")
        mb = self._get_member("owner_acc_b")
        self.assertIsNotNone(ma)
        self.assertIsNotNone(mb)

        r = admin.post("/api/stores", {"name": "T12-历史数据测试店",
                                       "owner_user_id": ma["user_id"]})
        self.assertEqual(r[0], 200)
        store_id = r[1]["data"]["store_id"]

        try:
            refund_before = admin.get(
                f"/api/refund-orders?store_id={store_id}")[1]
            cost_before = admin.get(
                f"/api/supplier-costs?store_id={store_id}")[1]

            admin.post(f"/api/stores/{store_id}/transfer-ownership",
                       {"target_user_id": mb["user_id"]})

            refund_after = admin.get(
                f"/api/refund-orders?store_id={store_id}")[1]
            cost_after = admin.get(
                f"/api/supplier-costs?store_id={store_id}")[1]

            self.assertEqual(refund_before.get("total", 0),
                             refund_after.get("total", 0),
                             "修改归属者前后退款总数应一致")
            self.assertEqual(cost_before.get("total", 0),
                             cost_after.get("total", 0),
                             "修改归属者前后供应商成本总数应一致")
        finally:
            admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})

    def test_13_nonexistent_store_transfer_fails(self):
        """不存在店铺转移 → 404，不虚报成功"""
        admin = self.admin
        mb = self._get_member("owner_acc_b")
        self.assertIsNotNone(mb)

        st, resp = admin.post("/api/stores/nonexistent_store_xyz/transfer-ownership",
                              {"target_user_id": mb["user_id"]})
        self.assertEqual(st, 404, "不存在的店铺应返回 404")
        self.assertNotIn("成功", str(resp.get("message", "")),
                         "失败时不得显示成功")

    def test_14_other_functions_regression(self):
        """其他功能回归 — 创建/编辑/删除店铺、用户列表正常"""
        admin = self.admin
        ma = self._get_member("owner_acc_a")
        self.assertIsNotNone(ma)

        st, resp = admin.post("/api/stores", {"name": "T14-回归测试店",
                                              "owner_user_id": ma["user_id"]})
        self.assertEqual(st, 200, "创建店铺应成功")
        store_id = resp["data"]["store_id"]

        try:
            st2, resp2 = admin.post(f"/api/stores/{store_id}/rename",
                                    {"name": "T14-已改名"})
            self.assertEqual(st2, 200, "编辑店铺名称应成功")

            sr = admin.get(f"/api/stores/{store_id}")[1]
            store_data = sr.get("data", sr)
            self.assertEqual(store_data.get("name"), "T14-已改名",
                             "店铺名称应已更新")

            st4, _ = admin.get("/api/users")
            self.assertEqual(st4, 200, "用户列表应可正常查询")
        finally:
            admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})

    def test_15_admin_views_all_stores(self):
        """admin 仍可查看全部店铺"""
        admin = self.admin
        ma = self._get_member("owner_acc_a")
        mb = self._get_member("owner_acc_b")
        self.assertIsNotNone(ma)
        self.assertIsNotNone(mb)

        r = admin.post("/api/stores", {"name": "T15-admin查看测试店",
                                       "owner_user_id": ma["user_id"]})
        self.assertEqual(r[0], 200)
        store_id = r[1]["data"]["store_id"]

        try:
            st, resp = admin.get("/api/stores")
            self.assertEqual(st, 200, "admin 查询店铺列表应成功")
            stores = resp.get("data", {}).get("stores", [])
            ids = [s.get("store_id") for s in stores]
            self.assertIn(store_id, ids, "admin 应能看到全部店铺")

            st2, resp2 = admin.get(f"/api/orders?store_id={store_id}")
            self.assertEqual(st2, 200, "admin 应可访问任意店铺数据")
        finally:
            admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})


if __name__ == "__main__":
    unittest.main()

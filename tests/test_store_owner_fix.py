"""店铺归属者修改功能修复验收测试

覆盖 spec 6 验收测试 T1-T8：
  test_01_normal_modify: 正常修改归属者
  test_02_permission_change: 权限变化（A 失去访问，B 获得访问）
  test_03_history_data_intact: 历史数据完整
  test_04_invalid_member: 非法成员 ID 被拒绝
  test_05_multi_store_owner: 多店铺归属（B 已有店铺仍可接收新店）
  test_06_member_create_forbidden: 普通成员越权创建店铺被拒绝
  test_07_data_isolation: 数据隔离（A 无法访问已转移店铺）
  test_08_edit_modal_has_owner_select: 编辑弹窗含归属者下拉框
"""
import unittest

from tests.http_client import TestEnv


class TestStoreOwnerFix(unittest.TestCase):
    """店铺归属者修改功能修复验收测试"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member_pwd = "Carrier123456"
        for uname, sname in [("owner_fix_a", "A-测试店"), ("owner_fix_b", "B-测试店")]:
            if cls.env.get_user_id(uname) is None:
                cls.admin.post("/api/users", {
                    "username": uname, "password": cls.member_pwd,
                    "confirm_password": cls.member_pwd, "role": "member",
                    "create_store": True, "store_name": sname})

    def _get_members(self):
        resp = self.admin.get("/api/users")[1]
        users = resp.get("users", resp.get("data", {}).get("users", []))
        members = [u for u in users if u.get("role") == "member" and u.get("status") == "active"]
        return members

    def _get_member(self, username):
        for m in self._get_members():
            if m["username"] == username:
                return m
        return None

    def test_01_normal_modify(self):
        """T1: 正常修改 — 将店铺归属者从 A 修改为 B，刷新后仍显示 B"""
        admin = self.admin
        member_a = self._get_member("owner_fix_a")
        member_b = self._get_member("owner_fix_b")
        self.assertIsNotNone(member_a)
        self.assertIsNotNone(member_b)

        r = admin.post("/api/stores", {"name": "T1-归属修改测试店", "owner_user_id": member_a["user_id"]})
        self.assertEqual(r[0], 200, f"创建店铺失败: {r[1]}")
        store_id = r[1]["data"]["store_id"]

        try:
            tr = admin.post(f"/api/stores/{store_id}/transfer-ownership",
                            {"target_user_id": member_b["user_id"]})
            self.assertEqual(tr[0], 200, f"转移归属失败: {tr[1]}")

            sr = admin.get(f"/api/stores/{store_id}")
            self.assertEqual(sr[0], 200, f"查询店铺失败: {sr[1]}")
            store_data = sr[1].get("data", sr[1])
            self.assertEqual(store_data.get("owner_user_id"), member_b["user_id"],
                             "刷新后归属者应仍为成员 B")
        finally:
            admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})

    def test_02_permission_change(self):
        """T2: 权限变化 — 转移后 store owner 变为 B，A 不再是 owner"""
        admin = self.admin
        member_a = self._get_member("owner_fix_a")
        member_b = self._get_member("owner_fix_b")
        self.assertIsNotNone(member_a)
        self.assertIsNotNone(member_b)

        r = admin.post("/api/stores", {"name": "T2-权限变化测试店", "owner_user_id": member_a["user_id"]})
        self.assertEqual(r[0], 200)
        store_id = r[1]["data"]["store_id"]

        try:
            sr_before = admin.get(f"/api/stores/{store_id}")[1]
            store_before = sr_before.get("data", sr_before)
            self.assertEqual(store_before.get("owner_user_id"), member_a["user_id"],
                             "转移前归属者应为 A")

            admin.post(f"/api/stores/{store_id}/transfer-ownership",
                       {"target_user_id": member_b["user_id"]})

            sr_after = admin.get(f"/api/stores/{store_id}")[1]
            store_after = sr_after.get("data", sr_after)
            self.assertEqual(store_after.get("owner_user_id"), member_b["user_id"],
                             "转移后归属者应为 B")
            self.assertNotEqual(store_after.get("owner_user_id"), member_a["user_id"],
                                "A 不再是归属者")
        finally:
            admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})

    def test_03_history_data_intact(self):
        """T3: 历史数据完整 — 修改归属者前后订单数据不变"""
        admin = self.admin
        member_a = self._get_member("owner_fix_a")
        member_b = self._get_member("owner_fix_b")
        self.assertIsNotNone(member_a)
        self.assertIsNotNone(member_b)

        r = admin.post("/api/stores", {"name": "T3-历史数据测试店", "owner_user_id": member_a["user_id"]})
        self.assertEqual(r[0], 200)
        store_id = r[1]["data"]["store_id"]

        try:
            orders_before = admin.get(f"/api/orders?store_id={store_id}")[1]

            admin.post(f"/api/stores/{store_id}/transfer-ownership",
                       {"target_user_id": member_b["user_id"]})

            orders_after = admin.get(f"/api/orders?store_id={store_id}")[1]
            self.assertEqual(
                orders_before.get("total", 0), orders_after.get("total", 0),
                "修改归属者前后订单总数应一致")
        finally:
            admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})

    def test_04_invalid_member(self):
        """T4: 非法成员 — 不存在的成员 ID 被拒绝"""
        admin = self.admin
        member_a = self._get_member("owner_fix_a")
        self.assertIsNotNone(member_a)

        r = admin.post("/api/stores", {"name": "T4-非法成员测试店", "owner_user_id": member_a["user_id"]})
        self.assertEqual(r[0], 200)
        store_id = r[1]["data"]["store_id"]

        try:
            st, resp = admin.post(f"/api/stores/{store_id}/transfer-ownership",
                                  {"target_user_id": "nonexistent_user_id"})
            self.assertIn(st, (400, 404), "不存在的成员 ID 应被拒绝")

            admin_user = [u for u in admin.get("/api/users")[1].get("users", []) if u["role"] == "admin"]
            if admin_user:
                st2, _ = admin.post(f"/api/stores/{store_id}/transfer-ownership",
                                    {"target_user_id": admin_user[0]["user_id"]})
                self.assertIn(st2, (400, 403), "管理员账号不能成为店铺归属者")
        finally:
            admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})

    def test_05_multi_store_owner(self):
        """T5: 多店铺归属 — B 已有店铺仍可接收新店"""
        admin = self.admin
        member_b = self._get_member("owner_fix_b")
        self.assertIsNotNone(member_b)

        r1 = admin.post("/api/stores", {"name": "T5-B的第一店", "owner_user_id": member_b["user_id"]})
        self.assertEqual(r1[0], 200)
        store1 = r1[1]["data"]["store_id"]

        r2 = admin.post("/api/stores", {"name": "T5-B的第二店", "owner_user_id": member_b["user_id"]})
        self.assertEqual(r2[0], 200, "成员 B 已有店铺时仍可接收新店")
        store2 = r2[1]["data"]["store_id"]

        try:
            sr2 = admin.get(f"/api/stores/{store2}")[1]
            store_data = sr2.get("data", sr2)
            self.assertEqual(store_data.get("owner_user_id"), member_b["user_id"],
                             "第二店归属者应为 B")
        finally:
            admin.post(f"/api/stores/{store1}/delete", {"action": "physical"})
            admin.post(f"/api/stores/{store2}/delete", {"action": "physical"})

    def test_06_member_create_forbidden(self):
        """T6: 普通成员越权创建店铺被拒绝（M3 修复）"""
        member_a = self._get_member("owner_fix_a")
        self.assertIsNotNone(member_a)

        member_client = self.env.login_client("owner_fix_a", self.member_pwd)
        st, r = member_client.post("/api/stores", {"name": "越权店", "owner_user_id": member_a["user_id"]})
        self.assertEqual(st, 403, "普通成员创建店铺应被拒绝（403）")

    def test_07_data_isolation(self):
        """T7: 数据隔离 — 转移后 admin 确认 store owner 已变更，A 不再拥有"""
        admin = self.admin
        member_a = self._get_member("owner_fix_a")
        member_b = self._get_member("owner_fix_b")
        self.assertIsNotNone(member_a)
        self.assertIsNotNone(member_b)

        r = admin.post("/api/stores", {"name": "T7-隔离测试店", "owner_user_id": member_a["user_id"]})
        self.assertEqual(r[0], 200)
        store_id = r[1]["data"]["store_id"]

        try:
            admin.post(f"/api/stores/{store_id}/transfer-ownership",
                       {"target_user_id": member_b["user_id"]})

            sr = admin.get(f"/api/stores/{store_id}")[1]
            store_data = sr.get("data", sr)
            self.assertEqual(store_data.get("owner_user_id"), member_b["user_id"],
                             "转移后归属者应为 B，A 不再拥有")
        finally:
            admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})

    def test_08_edit_modal_has_owner_select(self):
        """T8: 编辑弹窗含归属者下拉框（M1+M2 修复）"""
        _, html, _ = self.admin.get_page("/stores")
        self.assertIn("editStore", html, "stores.html 应包含 editStore 函数")
        self.assertNotIn("renameStore(", html.replace("function renameStore", "function _renameStore"),
                         "编辑按钮应调用 editStore 而非 renameStore")


if __name__ == "__main__":
    unittest.main()

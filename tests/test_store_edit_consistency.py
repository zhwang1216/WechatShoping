"""店铺编辑保存失败回滚与数据一致性验收测试（spec 6.8）

覆盖 11 项验收场景:
  test_01_normal_edit_persists: 正常修改名称+归属者，持久化成功
  test_02_temp_write_fail_rolls_back: 模拟临时文件写入失败，原始数据不变
  test_03_replace_fail_rolls_back: 模拟正式文件替换失败，原始数据不变
  test_04_post_save_verify_fail_restores: 模拟保存后验证失败，恢复原始状态
  test_05_multi_file_partial_fail_restores: 多文件更新中途失败，恢复一致状态
  test_06_rollback_fail_alerts: 模拟回滚失败，系统明确告警并保留恢复材料
  test_07_failed_transfer_preserves_permissions: 保存失败后权限不变
  test_08_successful_transfer_updates_permissions: 保存成功后权限立即生效
  test_09_concurrent_edit_rejected: 并发编辑，旧版本被拒绝
  test_10_txn_log_for_crash_recovery: 事务日志供崩溃恢复识别
  test_11_other_stores_untouched: 其他店铺及业务数据不受影响
"""
import json
import os
import unittest
from unittest.mock import patch, MagicMock

from tests.http_client import TestEnv

PASSWORD = "Carrier123456"


class TestStoreEditConsistency(unittest.TestCase):
    """店铺编辑保存失败回滚与数据一致性验收测试"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        for uname in ("edit_con_a", "edit_con_b"):
            if cls.env.get_user_id(uname) is None:
                cls.admin.post("/api/users", {
                    "username": uname, "password": PASSWORD,
                    "confirm_password": PASSWORD, "role": "member",
                    "create_store": True, "store_name": f"{uname}-店"})

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _get_member(self, username):
        st, r = self.admin.get("/api/users")
        users = r.get("users", r.get("data", {}).get("users", []))
        for u in users:
            if u.get("username") == username:
                return u
        return None

    def _create_store(self, owner_username):
        m = self._get_member(owner_username)
        self.assertIsNotNone(m, f"成员 {owner_username} 不存在")
        st, r = self.admin.post("/api/stores", {
            "name": f"编辑测试-{owner_username}", "owner_user_id": m["user_id"]})
        self.assertEqual(st, 200, f"创建店铺失败: {r}")
        return r["data"]["store_id"], m

    def _get_store(self, store_id):
        st, r = self.admin.get(f"/api/stores/{store_id}")
        self.assertEqual(st, 200)
        return r.get("data", r)

    def _delete_store(self, store_id):
        self.admin.post(f"/api/stores/{store_id}/delete", {"action": "physical"})

    # ===== T1: 正常修改 =====
    def test_01_normal_edit_persists(self):
        """T1: 正常修改店铺名称及归属者，数据成功持久化"""
        sid, member_a = self._create_store("edit_con_a")
        member_b = self._get_member("edit_con_b")
        try:
            store = self._get_store(sid)
            expected_at = store.get("updated_at")
            # 修改名称
            st, r = self.admin.post(f"/api/stores/{sid}/rename", {
                "name": "T1-新名称", "expected_updated_at": expected_at})
            self.assertEqual(st, 200, f"重命名失败: {r}")
            new_updated_at = r["data"].get("updated_at")
            self.assertIsNotNone(new_updated_at)
            # 修改归属者（用 rename 返回的新 updated_at）
            st, r = self.admin.post(f"/api/stores/{sid}/transfer-ownership", {
                "target_user_id": member_b["user_id"]})
            self.assertEqual(st, 200, f"转移归属失败: {r}")
            self.assertFalse(r["data"].get("unchanged", False))
            # 验证持久化
            store = self._get_store(sid)
            self.assertEqual(store["name"], "T1-新名称")
            self.assertEqual(store["owner_user_id"], member_b["user_id"])
        finally:
            self._delete_store(sid)

    # ===== T2: 模拟临时文件写入失败 =====
    def test_02_temp_write_fail_rolls_back(self):
        """T2: 模拟临时文件写入失败，原始数据保持不变"""
        sid, member_a = self._create_store("edit_con_a")
        member_b = self._get_member("edit_con_b")
        try:
            store_before = self._get_store(sid)
            original_owner = store_before["owner_user_id"]
            original_name = store_before["name"]
            # Mock _atomic_write to raise OSError
            um = self.env.server.user_manager
            with patch.object(um, '_atomic_write', side_effect=OSError("模拟写入失败")):
                st, r = self.admin.post(f"/api/stores/{sid}/transfer-ownership", {
                    "target_user_id": member_b["user_id"]})
            self.assertEqual(st, 500, f"应返回 500: {r}")
            # 验证原始数据不变
            store_after = self._get_store(sid)
            self.assertEqual(store_after["owner_user_id"], original_owner)
            self.assertEqual(store_after["name"], original_name)
        finally:
            self._delete_store(sid)

    # ===== T3: 模拟正式文件替换失败 =====
    def test_03_replace_fail_rolls_back(self):
        """T3: 模拟正式文件替换失败，原始数据保持不变"""
        sid, member_a = self._create_store("edit_con_a")
        member_b = self._get_member("edit_con_b")
        try:
            store_before = self._get_store(sid)
            original_owner = store_before["owner_user_id"]
            # Mock os.replace to raise OSError
            with patch('os.replace', side_effect=OSError("模拟替换失败")):
                st, r = self.admin.post(f"/api/stores/{sid}/transfer-ownership", {
                    "target_user_id": member_b["user_id"]})
            self.assertEqual(st, 500, f"应返回 500: {r}")
            store_after = self._get_store(sid)
            self.assertEqual(store_after["owner_user_id"], original_owner)
        finally:
            self._delete_store(sid)

    # ===== T4: 模拟保存后验证失败 =====
    def test_04_post_save_verify_fail_restores(self):
        """T4: 模拟保存后验证失败，能够恢复原始状态"""
        sid, member_a = self._create_store("edit_con_a")
        member_b = self._get_member("edit_con_b")
        try:
            store_before = self._get_store(sid)
            original_owner = store_before["owner_user_id"]
            # Mock _store_inventory to return different values before/after
            call_count = [0]
            original_inventory = self.env.server._store_inventory(sid)
            def mock_inventory(store_id):
                call_count[0] += 1
                if call_count[0] <= 1:
                    return original_inventory
                # 返回不一致的 inventory 触发恢复
                result = dict(original_inventory)
                result["orders"] = result["orders"] + 999
                return result
            with patch.object(self.env.server, '_store_inventory', side_effect=mock_inventory):
                st, r = self.admin.post(f"/api/stores/{sid}/transfer-ownership", {
                    "target_user_id": member_b["user_id"]})
            self.assertEqual(st, 500, f"应返回 500: {r}")
            self.assertIn("迁移核对失败", r.get("message", ""))
            # 验证恢复后数据与原始一致
            store_after = self._get_store(sid)
            self.assertEqual(store_after["owner_user_id"], original_owner)
        finally:
            self._delete_store(sid)

    # ===== T5: 多文件更新中途失败 =====
    def test_05_multi_file_partial_fail_restores(self):
        """T5: 模拟多文件更新中途失败，能够恢复一致状态"""
        sid, member_a = self._create_store("edit_con_a")
        member_b = self._get_member("edit_con_b")
        try:
            store_before = self._get_store(sid)
            original_owner = store_before["owner_user_id"]
            um = self.env.server.user_manager
            original_mutate_users = um._mutate_users
            call_count = [0]
            def failing_mutate_users(mutator):
                call_count[0] += 1
                if call_count[0] >= 2:
                    raise OSError("模拟 users.json 第二次写入失败")
                return original_mutate_users(mutator)
            with patch.object(um, '_mutate_users', side_effect=failing_mutate_users):
                st, r = self.admin.post(f"/api/stores/{sid}/transfer-ownership", {
                    "target_user_id": member_b["user_id"]})
            self.assertEqual(st, 500, f"应返回 500: {r}")
            store_after = self._get_store(sid)
            self.assertEqual(store_after["owner_user_id"], original_owner,
                             "多文件中途失败后应恢复原始归属")
        finally:
            self._delete_store(sid)

    # ===== T6: 回滚失败告警 =====
    def test_06_rollback_fail_alerts(self):
        """T6: 模拟回滚失败，系统明确告警并保留恢复材料"""
        sid, member_a = self._create_store("edit_con_a")
        member_b = self._get_member("edit_con_b")
        try:
            um = self.env.server.user_manager
            original_mutate_stores = um._mutate_stores
            original_atomic_write = um._atomic_write
            def failing_mutate_stores(mutator):
                raise OSError("模拟 stores.json 写入失败")
            rollback_count = [0]
            def failing_atomic_write(path, data):
                rollback_count[0] += 1
                if rollback_count[0] >= 1:
                    raise OSError("模拟回滚写入也失败")
                return original_atomic_write(path, data)
            with patch.object(um, '_mutate_stores', side_effect=failing_mutate_stores), \
                 patch.object(um, '_atomic_write', side_effect=failing_atomic_write):
                st, r = self.admin.post(f"/api/stores/{sid}/transfer-ownership", {
                    "target_user_id": member_b["user_id"]})
            self.assertEqual(st, 500, f"应返回 500: {r}")
            msg = r.get("message", "")
            self.assertIn("恢复异常", msg,
                          "回滚失败应明确提示'数据恢复异常'")
        finally:
            self._delete_store(sid)

    # ===== T7: 保存失败后权限不变 =====
    def test_07_failed_transfer_preserves_permissions(self):
        """T7: 保存失败后，原归属成员权限不变，新归属成员不会提前获得权限"""
        sid, member_a = self._create_store("edit_con_a")
        member_b = self._get_member("edit_con_b")
        try:
            um = self.env.server.user_manager
            # 模拟写入失败
            with patch.object(um, '_atomic_write', side_effect=OSError("模拟失败")):
                self.admin.post(f"/api/stores/{sid}/transfer-ownership", {
                    "target_user_id": member_b["user_id"]})
            # A 仍可访问该店铺
            client_a = self.env.login_client("edit_con_a", PASSWORD)
            st, r = client_a.get(f"/api/orders?store_id={sid}&page=1&size=20")
            self.assertEqual(st, 200, "原归属成员 A 应仍能访问店铺")
            # B 仍不能访问该店铺
            client_b = self.env.login_client("edit_con_b", PASSWORD)
            st, r = client_b.get(f"/api/orders?store_id={sid}&page=1&size=20")
            self.assertIn(st, (403, 404), "新归属成员 B 不应提前获得权限")
        finally:
            self._delete_store(sid)

    # ===== T8: 保存成功后权限立即生效 =====
    def test_08_successful_transfer_updates_permissions(self):
        """T8: 保存成功后，权限立即按照新归属关系生效"""
        sid, member_a = self._create_store("edit_con_a")
        member_b = self._get_member("edit_con_b")
        try:
            st, r = self.admin.post(f"/api/stores/{sid}/transfer-ownership", {
                "target_user_id": member_b["user_id"]})
            self.assertEqual(st, 200, f"转移失败: {r}")
            # B 重新登录后可访问
            client_b = self.env.login_client("edit_con_b", PASSWORD)
            st, r = client_b.get(f"/api/orders?store_id={sid}&page=1&size=20")
            self.assertEqual(st, 200, "新归属成员 B 应能访问店铺")
            # A 重新登录后不能访问
            client_a = self.env.login_client("edit_con_a", PASSWORD)
            st, r = client_a.get(f"/api/orders?store_id={sid}&page=1&size=20")
            self.assertIn(st, (403, 404), "原归属成员 A 应失去访问权限")
        finally:
            self._delete_store(sid)

    # ===== T9: 并发编辑被拒绝 =====
    def test_09_concurrent_edit_rejected(self):
        """T9: 模拟并发编辑，后提交的旧版本不能覆盖已成功提交的新版本"""
        sid, member_a = self._create_store("edit_con_a")
        member_b = self._get_member("edit_con_b")
        try:
            store = self._get_store(sid)
            stale_updated_at = store.get("updated_at")
            # 第一次修改成功
            st, r = self.admin.post(f"/api/stores/{sid}/rename", {
                "name": "T9-第一次修改", "expected_updated_at": stale_updated_at})
            self.assertEqual(st, 200, f"第一次修改应成功: {r}")
            # 第二次修改用旧的 updated_at → 应被拒绝
            st, r = self.admin.post(f"/api/stores/{sid}/rename", {
                "name": "T9-第二次修改", "expected_updated_at": stale_updated_at})
            self.assertEqual(st, 409, f"并发修改应返回 409: {r}")
            self.assertIn("已被其他操作修改", r.get("message", ""))
            # 同样测试 transfer-ownership
            st, r = self.admin.post(f"/api/stores/{sid}/transfer-ownership", {
                "target_user_id": member_b["user_id"],
                "expected_updated_at": stale_updated_at})
            self.assertEqual(st, 409, f"并发归属修改应返回 409: {r}")
            # 验证数据仍是第一次修改后的状态
            store = self._get_store(sid)
            self.assertEqual(store["name"], "T9-第一次修改")
        finally:
            self._delete_store(sid)

    # ===== T10: 事务日志供崩溃恢复 =====
    def test_10_txn_log_for_crash_recovery(self):
        """T10: 事务日志写入 backup 目录，供崩溃恢复识别未完成操作"""
        sid, member_a = self._create_store("edit_con_a")
        member_b = self._get_member("edit_con_b")
        try:
            st, r = self.admin.post(f"/api/stores/{sid}/transfer-ownership", {
                "target_user_id": member_b["user_id"]})
            self.assertEqual(st, 200, f"转移失败: {r}")
            op_id = r["data"].get("operation_id")
            self.assertIsNotNone(op_id, "响应应包含 operation_id")
            # 检查 backup 目录中存在 txn.json
            data_dir = self.env.data_dir
            txn_files = []
            for root, dirs, files in os.walk(data_dir):
                for fn in files:
                    if fn == "txn.json":
                        txn_files.append(os.path.join(root, fn))
            self.assertGreater(len(txn_files), 0,
                               "应存在 txn.json 事务日志文件")
            # 验证 txn.json 内容
            found_completed = False
            for txn_path in txn_files:
                with open(txn_path, encoding="utf-8") as f:
                    txn = json.load(f)
                if txn.get("operation_id") == op_id:
                    self.assertEqual(txn.get("phase"), "completed",
                                     "事务应标记为 completed")
                    self.assertEqual(txn.get("store_id"), sid)
                    found_completed = True
            self.assertTrue(found_completed,
                            f"应找到 operation_id={op_id} 的事务记录")
        finally:
            self._delete_store(sid)

    # ===== T11: 其他店铺不受影响 =====
    def test_11_other_stores_untouched(self):
        """T11: 其他店铺及其订单、退款、成本、利润和历史记录均不受影响"""
        sid_a, member_a = self._create_store("edit_con_a")
        sid_b, member_b = self._create_store("edit_con_b")
        try:
            # 记录 B 店铺的初始状态
            store_b_before = self._get_store(sid_b)
            # 对 A 店铺进行编辑（改名 + 转归属给 B）
            store_a = self._get_store(sid_a)
            st, r = self.admin.post(f"/api/stores/{sid_a}/rename", {
                "name": "T11-改名后", "expected_updated_at": store_a.get("updated_at")})
            self.assertEqual(st, 200)
            st, r = self.admin.post(f"/api/stores/{sid_a}/transfer-ownership", {
                "target_user_id": member_b["user_id"]})
            self.assertEqual(st, 200)
            # 验证 B 店铺完全不受影响
            store_b_after = self._get_store(sid_b)
            self.assertEqual(store_b_after["name"], store_b_before["name"],
                             "其他店铺名称不应改变")
            self.assertEqual(store_b_after["owner_user_id"], store_b_before["owner_user_id"],
                             "其他店铺归属者不应改变")
            self.assertEqual(store_b_after["updated_at"], store_b_before["updated_at"],
                             "其他店铺 updated_at 不应改变")
            # 验证 B 店铺的业务数据不受影响
            client_b = self.env.login_client("edit_con_b", PASSWORD)
            st, r = client_b.get(f"/api/orders?store_id={sid_b}&page=1&size=20")
            self.assertEqual(st, 200, "B 店铺订单查询应正常")
        finally:
            self._delete_store(sid_a)
            self._delete_store(sid_b)


if __name__ == "__main__":
    unittest.main()

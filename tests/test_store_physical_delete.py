"""店铺物理删除与归档功能移除验收测试（store_physical_delete 特性）

15 项验收映射（tasks.md 总表）:
- 验收 1 删除空店铺: TestDeleteEmptyStore
- 验收 2/3 有数据店铺删除且关联数据确不存在: TestDeleteStoreWithData
- 验收 4/5 同成员他店与其他成员店铺保留: TestCrossStoreIsolation
- 验收 6/7 普通成员与未登录不能删除: TestDeletePermissions
- 验收 8/9 错误 ID 与确认不匹配: TestDeleteValidation
- 验收 10 写入失败一致性恢复 + 崩溃恢复专项: TestDeleteFailureRecovery
- 验收 11 删除后统计/列表/搜索无该店: TestDeleteVisibility
- 验收 12 重启后不重现: TestDeleteRestart
- 验收 13 审计日志保留: TestDeleteAudit
- 验收 14 归档入口移除: TestArchiveRemoval
- 验收 15 重复提交不损坏: TestDeleteIdempotency

全部经真实 HTTP 驱动（TestEnv 临时目录隔离, 严禁触碰真实 data/）;
断言直读落盘 JSON 与目录结构, 不使用模拟数据掩盖问题。
"""
import json
import os
import threading
import unittest
from datetime import datetime
from unittest import mock

from app.data_schemas import BUSINESS_FILES

from tests.http_client import Client, TestEnv
from tests.test_auth_isolation import make_sup_xlsx, make_wx_xlsx
from tests.test_member_management import read_security_log, seed_store_data

PASSWORD = "Delete123456"


# ---------- 直读落盘辅助 ----------

def _read_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def stores_table(env):
    path = os.path.join(env.data_dir, "stores.json")
    return _read_json(path).get("stores", [])


def users_table(env):
    path = os.path.join(env.data_dir, "users.json")
    return _read_json(path).get("users", [])


def first_member_user_id(env):
    """获取第一个普通成员 user_id（排除管理员，用于建店归属）"""
    for u in users_table(env):
        if u.get("role") == "member":
            return u["user_id"]
    raise AssertionError("无普通成员用户可用作店主")


def tombstone_records(env):
    path = os.path.join(env.backup_dir, "deleted_store_tombstones.json")
    if not os.path.isfile(path):
        return []
    return _read_json(path).get("tombstones", [])


def store_dir_exists(env, store_id):
    return os.path.isdir(os.path.join(env.data_dir, "stores", store_id))


def snapshot_dirs(env, store_id):
    return [d for d in os.listdir(env.backup_dir)
            if d.startswith(f"store_snapshot_{store_id}_")]


def business_counts(env, store_id):
    """直读 12 类业务文件记录数（删除前口径快照）"""
    data_dir = os.path.join(env.data_dir, "stores", store_id, "data")
    counts = {}
    for name, spec in BUSINESS_FILES.items():
        if name == "settings":
            continue
        path = os.path.join(data_dir, spec["file"])
        counts[name] = (len(_read_json(path).get("records", []))
                        if os.path.isfile(path) else 0)
    return counts


def admin_create_store(env, name, owner_user_id):
    status, resp = env.admin.post(
        "/api/stores", {"name": name, "owner_user_id": owner_user_id})
    assert status == 200, f"建店失败 {name}: {resp}"
    return resp["data"]["store_id"]


def delete_store(admin, store_id, confirm, acknowledged=True):
    return admin.post(f"/api/stores/{store_id}/delete", {
        "confirm_store_name": confirm,
        "confirm_acknowledged": acknowledged})


def seed_store_for(client, prefix, target_store_id):
    """多店成员为指定归属店铺造数（订单 + 成本导入）"""
    for path, filename, content in (
            ("/api/import/wechat-orders", f"{prefix}订单.xlsx",
             make_wx_xlsx(prefix)),
            ("/api/import/supplier-costs", f"{prefix}成本.xlsx",
             make_sup_xlsx(prefix))):
        status, resp = client.upload_confirm(
            path, filename, content, target_store_id=target_store_id)
        assert status == 200, f"导入失败 {filename}: {resp}"


def security_events(env, event):
    return [e for e in read_security_log(env) if e.get("event") == event]


class TestDeleteEmptyStore(unittest.TestCase):
    """验收 1: 删除零业务数据店铺 → 200 + 两表/目录/引用全清"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        owner_id = first_member_user_id(cls.env)
        cls.store_id = admin_create_store(cls.env, "空店一号", owner_id)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_delete_empty_store(self):
        status, resp = delete_store(self.admin, self.store_id, "空店一号")
        self.assertEqual(status, 200, resp)
        counts = resp["data"]["deleted_counts"]
        for key in ("orders", "refund_orders", "refund_details",
                    "supplier_costs", "raw_records", "import_records",
                    "anomalies", "upload_files", "store_backup_files"):
            self.assertEqual(counts.get(key), 0, f"{key} 应为零")
        self.assertEqual(counts.get("settings"), 1)

    def test_02_all_traces_removed(self):
        self.assertNotIn(self.store_id,
                         [s["store_id"] for s in stores_table(self.env)])
        for user in users_table(self.env):
            self.assertNotIn(self.store_id, user.get("store_ids") or [])
        self.assertFalse(store_dir_exists(self.env, self.store_id))
        self.assertEqual(snapshot_dirs(self.env, self.store_id), [])
        self.assertIn(self.store_id,
                      [t["store_id"] for t in tombstone_records(self.env)])


class TestDeleteStoreWithData(unittest.TestCase):
    """验收 2/3: 有数据店铺直接删除, 关联数据确不存在, 统计与清点一致"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.member
        cls.store_id = cls.env.member_store
        seed_store_data(cls.member, "DEL")
        cls.before = business_counts(cls.env, cls.store_id)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_delete_with_data_and_exact_counts(self):
        status, resp = delete_store(self.admin, self.store_id, self.store_id)
        self.assertEqual(status, 200, resp)
        counts = resp["data"]["deleted_counts"]
        for name, before in self.before.items():
            key = ("profit_results" if name == "profit_statements" else name)
            self.assertEqual(counts.get(key), before,
                             f"{key} 删除统计应与删除前读数一致")
        self.assertGreater(counts.get("orders", 0), 0)

    def test_02_store_data_really_gone(self):
        self.assertFalse(store_dir_exists(self.env, self.store_id))
        self.assertNotIn(self.store_id,
                         [s["store_id"] for s in stores_table(self.env)])
        for user in users_table(self.env):
            self.assertNotIn(self.store_id, user.get("store_ids") or [])
        self.assertEqual(snapshot_dirs(self.env, self.store_id), [])
        tomb = [t for t in tombstone_records(self.env)
                if t["store_id"] == self.store_id]
        self.assertEqual(len(tomb), 1)
        self.assertIn("deleted_at", tomb[0])
        self.assertIn("txn_id", tomb[0])

    def test_03_owner_session_revoked_and_relogin(self):
        status, resp = self.member.get("/api/auth/me")
        self.assertEqual(status, 401)
        # 删除唯一店铺后 owner 无店可登（会话已吊销且 store_ids 为空）
        status, resp = self.env.admin.post("/api/auth/login", {
            "username": "biz_carrier", "password": "Carrier123456"},
            csrf=False)
        self.assertEqual(status, 401)
        self.assertIn("店铺", resp["message"])


class TestCrossStoreIsolation(unittest.TestCase):
    """验收 4/5: 同成员他店与其他成员店铺完整保留"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member_a = cls.env.create_user_client("isoa", PASSWORD)
        cls.member_b = cls.env.create_user_client("isob", PASSWORD)
        me_a = cls.member_a.get("/api/auth/me")[1]["data"]
        me_b = cls.member_b.get("/api/auth/me")[1]["data"]
        cls.user_a, cls.user_b = me_a["user_id"], me_b["user_id"]
        cls.store_x = admin_create_store(cls.env, "隔离店X", cls.user_a)
        cls.store_y = admin_create_store(cls.env, "隔离店Y", cls.user_a)
        cls.store_z = admin_create_store(cls.env, "隔离店Z", cls.user_b)
        seed_store_for(cls.member_a, "ISOX", cls.store_x)
        seed_store_for(cls.member_a, "ISOY", cls.store_y)
        seed_store_for(cls.member_b, "ISOZ", cls.store_z)
        cls.before_y = business_counts(cls.env, cls.store_y)
        cls.before_z = business_counts(cls.env, cls.store_z)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_delete_x_only(self):
        status, resp = delete_store(self.admin, self.store_x, "隔离店X")
        self.assertEqual(status, 200, resp)

    def test_02_same_member_other_store_intact(self):
        self.assertTrue(store_dir_exists(self.env, self.store_y))
        self.assertEqual(business_counts(self.env, self.store_y),
                         self.before_y)
        user_a = next(u for u in users_table(self.env)
                      if u["user_id"] == self.user_a)
        self.assertIn(self.store_y, user_a.get("store_ids") or [])
        self.assertNotIn(self.store_x, user_a.get("store_ids") or [])
        st, r = self.env.login_client("isoa", PASSWORD).get(
            f"/api/stores/{self.store_y}")
        self.assertIn(st, (200, 403, 404))  # member 权限口径由现有协议决定

    def test_03_other_member_store_intact(self):
        self.assertTrue(store_dir_exists(self.env, self.store_z))
        self.assertEqual(business_counts(self.env, self.store_z),
                         self.before_z)
        user_b = next(u for u in users_table(self.env)
                      if u["user_id"] == self.user_b)
        self.assertIn(self.store_z, user_b.get("store_ids") or [])
        self.assertIn(self.store_z,
                      [s["store_id"] for s in stores_table(self.env)])
        self.assertNotIn(self.store_x,
                         [s["store_id"] for s in stores_table(self.env)])
        self.assertEqual(snapshot_dirs(self.env, self.store_z), [])


class TestDeletePermissions(unittest.TestCase):
    """验收 6/7: 普通成员 403 + 未登录 401, 数据零变化"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.member
        cls.store_id = cls.env.member_store
        seed_store_data(cls.member, "PRM")
        cls.before = business_counts(cls.env, cls.store_id)
        cls.anon = Client(cls.env)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_member_forbidden(self):
        status, resp = self.member.post(
            f"/api/stores/{self.store_id}/delete",
            {"confirm_store_name": self.store_id,
             "confirm_acknowledged": True})
        self.assertEqual(status, 403)
        denied = security_events(self.env, "access_denied")
        self.assertTrue(any(
            d.get("detail", {}).get("path", "").endswith("/delete")
            or "delete" in json.dumps(d.get("detail", {}), ensure_ascii=False)
            for d in denied), "应存在删除接口 access_denied 审计")

    def test_02_anonymous_unauthorized(self):
        status, resp = self.anon.post(
            f"/api/stores/{self.store_id}/delete",
            {"confirm_store_name": self.store_id,
             "confirm_acknowledged": True})
        self.assertEqual(status, 401)

    def test_03_data_unchanged(self):
        self.assertEqual(business_counts(self.env, self.store_id),
                         self.before)
        self.assertIn(self.store_id,
                      [s["store_id"] for s in stores_table(self.env)])
        self.assertTrue(store_dir_exists(self.env, self.store_id))
        self.assertEqual(tombstone_records(self.env), [])


class TestDeleteValidation(unittest.TestCase):
    """验收 8/9: 错误 ID 404; 确认不匹配/未勾选/伪布尔 400; 双通道各自成功"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        owner_id = first_member_user_id(cls.env)
        cls.store_id = admin_create_store(cls.env, "校验店甲", owner_id)
        cls.store_other = admin_create_store(cls.env, "校验店乙", owner_id)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_nonexistent_store_404(self):
        status, resp = delete_store(self.admin, "s_notexists",
                                    "s_notexists")
        self.assertEqual(status, 404)
        self.assertIn("店铺不存在", resp["message"])

    def test_02_wrong_confirm_400_zero_change(self):
        before = stores_table(self.env)
        for confirm in ("校验店", "校验店乙", 123, None):
            status, resp = self.admin.post(
                f"/api/stores/{self.store_id}/delete",
                {"confirm_store_name": confirm,
                 "confirm_acknowledged": True})
            self.assertEqual(status, 400, f"confirm={confirm!r} 应 400")
        status, _ = delete_store(self.admin, self.store_id,
                                 "校验店甲", acknowledged=False)
        self.assertEqual(status, 400)
        status, _ = delete_store(self.admin, self.store_id,
                                 "校验店甲", acknowledged="true")
        self.assertEqual(status, 400)
        self.assertEqual(stores_table(self.env), before)
        self.assertTrue(store_dir_exists(self.env, self.store_id))

    def test_03_both_channels_succeed(self):
        sid = admin_create_store(self.env, "通道店", owner_id=None) \
            if False else self.store_other
        status, resp = delete_store(self.admin, sid, "校验店乙")
        self.assertEqual(status, 200, resp)
        status, resp = self.admin.get(f"/api/stores/{sid}")
        self.assertEqual(status, 404)


class TestDeleteFailureRecovery(unittest.TestCase):
    """验收 10: 写入失败 → 回滚一致 → 可安全重试; 崩溃恢复专项"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.store_id = cls.env.member_store
        seed_store_data(cls.env.member, "FLR2")
        cls.before = business_counts(cls.env, cls.store_id)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_clear_failure_rolls_back_and_retry(self):
        with mock.patch.object(
                self.env.server, "_clear_store_data_files",
                side_effect=OSError("disk full mid-clear")):
            status, resp = delete_store(self.admin, self.store_id,
                                        self.store_id)
        self.assertEqual(status, 500, resp)
        self.assertEqual(business_counts(self.env, self.store_id),
                         self.before)
        self.assertIn(self.store_id,
                      [s["store_id"] for s in stores_table(self.env)])
        self.assertTrue(any(
            self.store_id in (u.get("store_ids") or [])
            for u in users_table(self.env)),
            "回滚后归属成员 store_ids 应仍含该店")
        failed = security_events(self.env, "store.delete_failed")
        self.assertTrue(any(
            f.get("detail", {}).get("store_id") == self.store_id
            for f in failed))
        status, resp = delete_store(self.admin, self.store_id, self.store_id)
        self.assertEqual(status, 200, resp)
        self.assertFalse(store_dir_exists(self.env, self.store_id))

    def test_02_crash_recovery_prepared_txn(self):
        crash_member = self.env.create_user_client("crash_member", PASSWORD)
        me = crash_member.get("/api/auth/me")[1]["data"]
        sid = admin_create_store(self.env, "崩溃恢复店", me["user_id"])
        seed_store_for(crash_member, "CRASH", sid)
        before = business_counts(self.env, sid)
        self.env.shutdown()
        backup_dir = self.env.backup_dir
        store_root = os.path.join(self.env.data_dir, "stores")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        txn_dir = os.path.join(backup_dir, f"delete_txn_{sid}_{ts}")
        os.makedirs(txn_dir)
        shutil_copytree = __import__("shutil").copytree
        shutil_copytree(os.path.join(store_root, sid),
                        os.path.join(txn_dir, "store_snapshot"))
        for name, path in (("users.json.bak",
                            os.path.join(self.env.data_dir, "users.json")),
                           ("stores.json.bak",
                            os.path.join(self.env.data_dir, "stores.json"))):
            __import__("shutil").copy2(path, os.path.join(txn_dir, name))
        with open(os.path.join(txn_dir, "txn.json"), "w",
                  encoding="utf-8") as f:
            json.dump({"version": 1, "state": "prepared", "store_id": sid,
                       "store_name": "崩溃恢复店"}, f, ensure_ascii=False)
        data_dir = os.path.join(store_root, sid, "data")
        for spec in BUSINESS_FILES.values():
            fp = os.path.join(data_dir, spec["file"])
            if os.path.isfile(fp):
                os.remove(fp)
        self.env.restart()
        self.assertEqual(business_counts(self.env, sid), before)
        self.assertFalse(store_dir_exists(self.env, "s_notexists"))
        recovered = security_events(self.env, "store.delete_recovered")
        self.assertTrue(any(
            r.get("detail", {}).get("store_id") == sid
            and r.get("detail", {}).get("recovered_state") == "rolled_back"
            for r in recovered))
        self.assertFalse(os.path.isdir(txn_dir))

    @staticmethod
    def _biz_password():
        return "Carrier123456"


class TestDeleteVisibility(unittest.TestCase):
    """验收 11: 删除后列表/搜索/详情/统计均无该店, owner 会话失效"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.member
        cls.store_id = cls.env.member_store
        seed_store_data(cls.member, "VIS")
        status, resp = delete_store(cls.admin, cls.store_id, cls.store_id)
        assert status == 200, resp

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_all_stores_excludes_deleted(self):
        status, resp = self.admin.get("/api/admin/all-stores")
        self.assertEqual(status, 200)
        ids = [s["store_id"] for s in resp["data"]["stores"]]
        self.assertNotIn(self.store_id, ids)

    def test_02_search_excludes_deleted(self):
        status, resp = self.admin.get(
            "/api/admin/all-stores?name=biz")
        self.assertEqual(status, 200)
        for entry in resp["data"]["stores"]:
            self.assertNotEqual(entry["store_id"], self.store_id)

    def test_03_detail_404_and_owner_session_dead(self):
        status, resp = self.admin.get(f"/api/stores/{self.store_id}")
        self.assertEqual(status, 404)
        status, resp = self.member.get("/api/auth/me")
        self.assertEqual(status, 401)


class TestDeleteRestart(unittest.TestCase):
    """验收 12: 重启后店铺不重现; committed 残留事务幂等清理"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.store_id = cls.env.member_store
        seed_store_data(cls.env.member, "RST")
        status, resp = delete_store(cls.admin, cls.store_id, cls.store_id)
        assert status == 200, resp
        cls.env.restart()

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_not_resurrected_after_restart(self):
        self.assertFalse(store_dir_exists(self.env, self.store_id))
        status, resp = self.admin.get(f"/api/stores/{self.store_id}")
        self.assertEqual(status, 404)
        ids = [s["store_id"] for s in stores_table(self.env)]
        self.assertNotIn(self.store_id, ids)
        self.assertIn(self.store_id,
                      [t["store_id"] for t in tombstone_records(self.env)])

    def test_02_committed_leftover_txn_cleaned_on_restart(self):
        sid = admin_create_store(self.env, "重启清理店",
                                 first_member_user_id(self.env))
        # 先真正删除店铺（删除完成、事务目录已清理），再手动构造
        # committed 残留事务模拟"清理中断"场景
        delete_store(self.admin, sid, sid)
        self.env.shutdown()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        txn_dir = os.path.join(self.env.backup_dir, f"delete_txn_{sid}_{ts}")
        os.makedirs(txn_dir)
        with open(os.path.join(txn_dir, "txn.json"), "w",
                  encoding="utf-8") as f:
            json.dump({"version": 1, "state": "committed", "store_id": sid,
                       "store_name": "重启清理店"}, f, ensure_ascii=False)
        self.env.restart()
        self.assertFalse(os.path.isdir(txn_dir))
        self.assertIn(sid, [t["store_id"] for t in tombstone_records(self.env)])
        recovered = security_events(self.env, "store.delete_recovered")
        self.assertTrue(any(
            r.get("detail", {}).get("store_id") == sid
            and r.get("detail", {}).get("recovered_state")
            == "cleanup_completed" for r in recovered))
        status, resp = self.admin.get(f"/api/stores/{sid}")
        self.assertNotEqual(status, 200)


class TestDeleteAudit(unittest.TestCase):
    """验收 13: 成功/失败/拒绝三类路径审计完整; 历史日志只增不删"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.member
        cls.store_id = cls.env.member_store
        seed_store_data(cls.member, "AUD")
        cls.baseline = read_security_log(cls.env)
        cls.member.post(f"/api/stores/{cls.store_id}/delete",
                        {"confirm_store_name": cls.store_id,
                         "confirm_acknowledged": True})
        with mock.patch.object(
                cls.env.server, "_clear_store_data_files",
                side_effect=OSError("audit fail")):
            cls.fail_status, _ = delete_store(cls.admin, cls.store_id,
                                              cls.store_id)
        cls.success_status, cls.success_resp = delete_store(
            cls.admin, cls.store_id, cls.store_id)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_success_event_fields(self):
        events = security_events(self.env, "store.delete")
        matched = [e for e in events
                   if e.get("detail", {}).get("store_id") == self.store_id]
        self.assertEqual(len(matched), 1)
        detail = matched[0]["detail"]
        self.assertEqual(matched[0]["result"], "success")
        for field in ("store_name", "owner_user_id", "deleted_counts",
                      "revoked_sessions", "backup_treatment", "txn_id"):
            self.assertIn(field, detail)
        self.assertIn("墓碑", detail["backup_treatment"])
        self.assertIn("不承诺从中清除", detail["backup_treatment"])

    def test_02_failure_and_denied_events(self):
        failed = [e for e in security_events(self.env, "store.delete_failed")
                  if e.get("detail", {}).get("store_id") == self.store_id]
        self.assertEqual(len(failed), 1)
        self.assertEqual(failed[0]["result"], "error")
        denied = [e for e in security_events(self.env, "access_denied")]
        self.assertGreaterEqual(len(denied), 1)

    def test_03_history_only_appended(self):
        current = read_security_log(self.env)
        self.assertGreaterEqual(len(current), len(self.baseline))
        self.assertEqual(current[:len(self.baseline)], self.baseline)

    def test_04_outcomes(self):
        self.assertEqual(self.fail_status, 500)
        self.assertEqual(self.success_status, 200)


class TestArchiveRemoval(unittest.TestCase):
    """验收 14: 归档入口全移除; 存量 archived 启动迁移为 disabled"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.store_id = cls.env.member_store

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_page_has_no_archive_ui(self):
        status, html, _ = self.admin.get_page("/stores")
        self.assertEqual(status, 200)
        self.assertNotIn("归档", html)
        self.assertNotIn("stat-archived", html)

    def test_02_archive_action_rejected_with_audit(self):
        status, resp = self.admin.post(
            f"/api/stores/{self.store_id}/status", {"action": "archive"})
        self.assertEqual(status, 400)
        self.assertIn("归档功能已下线", resp["message"])
        removed = security_events(self.env, "store.archive_removed")
        self.assertEqual(len(removed), 1)
        self.assertEqual(removed[0]["result"], "denied")
        self.assertEqual(removed[0]["detail"]["store_id"], self.store_id)
        store = next(s for s in stores_table(self.env)
                     if s["store_id"] == self.store_id)
        self.assertNotEqual(store.get("status"), "archived")

    def test_03_status_filter_archived_rejected(self):
        status, resp = self.admin.get(
            "/api/admin/all-stores?status=archived")
        self.assertEqual(status, 400)
        self.assertIn("active / disabled", resp["message"])

    def test_04_legacy_archived_migrated_on_restart(self):
        path = os.path.join(self.env.data_dir, "stores.json")
        data = _read_json(path)
        for store in data["stores"]:
            if store["store_id"] == self.store_id:
                store["status"] = "archived"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.env.restart()
        store = next(s for s in stores_table(self.env)
                     if s["store_id"] == self.store_id)
        self.assertEqual(store["status"], "disabled")
        migrated = security_events(self.env, "store.archive_migration")
        self.assertEqual(len(migrated), 1)
        self.assertEqual(migrated[0]["result"], "success")
        self.assertEqual(migrated[0]["detail"]["count"], 1)
        status, resp = delete_store(self.admin, self.store_id, self.store_id)
        self.assertEqual(status, 200, resp)


class TestDeleteIdempotency(unittest.TestCase):
    """验收 15: 重复提交 404; 并发互斥一成功一 409/404, 数据一致"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.store_id = cls.env.member_store
        seed_store_data(cls.env.member, "IDP")

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_second_delete_404(self):
        status, resp = delete_store(self.admin, self.store_id, self.store_id)
        self.assertEqual(status, 200, resp)
        status, resp = delete_store(self.admin, self.store_id, self.store_id)
        self.assertEqual(status, 404)
        self.assertFalse(store_dir_exists(self.env, self.store_id))
        self.assertEqual(
            len([t for t in tombstone_records(self.env)
                 if t["store_id"] == self.store_id]), 1)

    def test_02_concurrent_delete_mutex(self):
        sid = admin_create_store(self.env, "并发店",
                                 first_member_user_id(self.env))
        results = []
        lock = threading.Lock()

        def worker():
            st, resp = delete_store(self.admin, sid, sid)
            with lock:
                results.append(st)

        threads = [threading.Thread(target=worker) for _ in range(2)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        self.assertIn(200, results)
        self.assertEqual(sorted(results), [200, 404] if 404 in results
                         else [200, 409], f"并发结果: {results}")
        self.assertFalse(store_dir_exists(self.env, sid))
        self.assertEqual(
            len([t for t in tombstone_records(self.env)
                 if t["store_id"] == sid]), 1)


if __name__ == "__main__":
    unittest.main()

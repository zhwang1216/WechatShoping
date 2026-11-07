"""成员管理验收测试：账号删除 / 店铺数据清理 / 自我保护 / 异常恢复 / 跨店铺隔离

对应故障报告"删除用户提示接口不存在"与七类验收（测试 A~G），全部经真实 HTTP 驱动
（隔离临时数据环境, 严禁触碰真实 data/）：
- 测试 A 接口存在性 + 测试 G 页面交互: TestMemberManagementApiAndPage
- 测试 B 仅删账号（场景 A, 保留数据可追溯）: TestDeleteKeepData
- 测试 C 删账号并清店铺数据（场景 B, 快照可读）: TestDeleteClearData
- 测试 D 跨店铺隔离 + 业务回归: TestCrossStoreIsolation
- 测试 E 权限与自我保护: TestPermissionProtection
- 测试 F 异常恢复（快照失败/清理中断/删账号失败）: TestFailureRecovery
"""
import json
import os
import unittest
from unittest import mock

from app.data_schemas import BUSINESS_FILES
from app.server import ApiError
from tests.http_client import Client, TestEnv
from tests.test_auth_isolation import make_sup_xlsx, make_wx_xlsx

PASSWORD = "Member123456"
ADMIN_STORE = "admin的店铺"


def read_security_log(env):
    """读取全局安全日志（jsonl）→ 事件字典列表"""
    path = os.path.join(env.log_dir, "security_log.jsonl")
    entries = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                entries.append(json.loads(line))
    return entries


def store_records(env, store_id, name):
    """直读店铺业务 JSON（数据落盘断言用）"""
    path = os.path.join(env.data_dir, "stores", store_id, "data",
                        BUSINESS_FILES[name]["file"])
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def read_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def me_of(client):
    status, resp = client.get("/api/auth/me")
    assert status == 200, f"/api/auth/me 失败: {resp}"
    return resp["data"]


def seed_store_data(client, prefix):
    """为成员店铺生成完整业务数据: 订单+成本导入 → 对账 → 利润 → 广告费"""
    for path, filename, content in (
            ("/api/import/wechat-orders", f"{prefix}订单.xlsx",
             make_wx_xlsx(prefix)),
            ("/api/import/supplier-costs", f"{prefix}成本.xlsx",
             make_sup_xlsx(prefix))):
        status, resp = client.upload_confirm(path, filename, content)
        assert status == 200, f"导入失败 {filename}: {resp}"
    client.post("/api/reconciliation/run", {"month": "2026-03"})
    client.post("/api/profit/recalculate", {"month": "2026-03"})
    status, resp = client.post("/api/ad-expenses",
                               {"month": "2026-03", "amount": 88.5})
    assert status == 200, f"广告费保存失败: {resp}"


def order_count(client):
    status, resp = client.get("/api/orders?page=1&size=100")
    assert status == 200
    return resp["data"]["total"]


class TestMemberManagementApiAndPage(unittest.TestCase):
    """测试 A: 三个新接口真实注册（不再返回"接口不存在"）; 测试 G: 页面交互入口"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client("mcheck", PASSWORD)
        cls.member_id = me_of(cls.member)["user_id"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _assert_user_not_found(self, status, resp):
        """路由已匹配但用户不存在: 404 且消息是"用户不存在"而非"接口不存在\""""
        self.assertEqual(status, 404)
        self.assertNotIn("接口不存在", resp["message"])
        self.assertIn("用户不存在", resp["message"])

    def test_a01_delete_route_registered(self):
        status, resp = self.admin.post(
            "/api/users/u_notexists/delete",
            {"confirm_username": "nobody", "clear_store_data": False})
        self._assert_user_not_found(status, resp)

    def test_a02_clear_route_registered(self):
        status, resp = self.admin.post(
            "/api/users/u_notexists/store-data-clear",
            {"confirm_store_name": "某店"})
        self._assert_user_not_found(status, resp)

    def test_a03_summary_route_registered(self):
        status, resp = self.admin.get(
            "/api/users/u_notexists/store-data-summary")
        self._assert_user_not_found(status, resp)

    def test_a04_delete_requires_login(self):
        anon = Client(self.env)
        status, resp = anon.post(
            f"/api/users/{self.member_id}/delete",
            {"confirm_username": "mcheck", "clear_store_data": False})
        self.assertEqual(status, 401)
        self.assertEqual(resp["code"], 1001)

    def test_g01_admin_page_has_delete_controls(self):
        status, html, _ = self.admin.get_page("/users")
        self.assertEqual(status, 200)
        for needle in ('data-act="delete"', 'data-act="clear"',
                       "仅删除账号，保留业务数据", "删除账号并清除店铺数据",
                       "输入店铺名称以确认"):
            self.assertIn(needle, html)

    def test_g02_member_page_forbidden(self):
        """member/未登录/静态路径直访均无法打开用户管理页面"""
        status, _html, headers = self.member.get_page("/users")
        self.assertIn(status, (301, 302))
        self.assertNotEqual((headers.get("Location") or ""), "/users")
        anon = Client(self.env)
        status, _html, headers = anon.get_page("/users")
        self.assertEqual(status, 302)
        self.assertIn("/login", headers.get("Location") or "")
        status, _html, _h = self.member.get_page("/static/users.html")
        self.assertIn(status, (301, 302, 404))


class TestDeleteKeepData(unittest.TestCase):
    """测试 B（场景 A）: 仅删除账号, 业务数据完整保留可追溯"""

    USER = "keepd"
    STORE = "保留数据测试店"

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            cls.USER, PASSWORD, store_name=cls.STORE)
        me = me_of(cls.member)
        cls.user_id = me["user_id"]
        cls.store_id = me["store_id"]
        seed_store_data(cls.member, "KPA")
        cls.orders_before = len(
            store_records(cls.env, cls.store_id, "orders")["records"])

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_b01_delete_keeps_data(self):
        status, resp = self.admin.post(
            f"/api/users/{self.user_id}/delete",
            {"confirm_username": self.USER, "clear_store_data": False})
        self.assertEqual(status, 200)
        data = resp["data"]
        self.assertEqual(data["username"], self.USER)
        self.assertFalse(data["clear_store_data"])
        self.assertEqual(data["store_ids"], [self.store_id])
        self.assertIn("保留", resp["message"])

    def test_b02_cannot_login_after_delete(self):
        client = Client(self.env)
        status, resp = client.post("/api/auth/login",
                                   {"username": self.USER,
                                    "password": PASSWORD}, csrf=False)
        self.assertNotEqual(status, 200)
        self.assertNotEqual(resp.get("code"), 0)

    def test_b03_old_session_revoked(self):
        status, resp = self.member.get("/api/auth/me")
        self.assertEqual(status, 401)
        status, resp = self.member.get("/api/orders?page=1&size=20")
        self.assertEqual(status, 401)

    def test_b04_store_and_data_preserved(self):
        stores = read_json(os.path.join(cls_dir := self.env.data_dir,
                                        "stores.json"))["stores"]
        self.assertIn(self.store_id, [s["store_id"] for s in stores])
        orders = store_records(self.env, self.store_id, "orders")
        self.assertEqual(len(orders["records"]), self.orders_before)
        for name in BUSINESS_FILES:
            path = os.path.join(self.env.data_dir, "stores", self.store_id,
                                "data", BUSINESS_FILES[name]["file"])
            self.assertTrue(os.path.exists(path), f"{name} 文件缺失")

    def test_b05_audit_log_written(self):
        entries = read_security_log(self.env)
        hits = [e for e in entries if e["event"] == "user.delete"
                and e["detail"]["target_username"] == self.USER]
        self.assertEqual(len(hits), 1)
        detail = hits[0]["detail"]
        self.assertFalse(detail["clear_store_data"])
        self.assertEqual(detail["store_ids"], [self.store_id])
        self.assertIn("revoked_sessions", detail)
        line = json.dumps(hits[0], ensure_ascii=False).lower()
        self.assertNotIn("password", line)
        self.assertNotIn(PASSWORD.lower(), line)

    def test_b06_removed_from_user_list(self):
        status, resp = self.admin.get("/api/users")
        self.assertEqual(status, 200)
        names = [u["username"] for u in resp["data"]["users"]]
        self.assertNotIn(self.USER, names)


class TestDeleteClearData(unittest.TestCase):
    """测试 C（场景 B）: 删除账号并清除店铺数据, 快照可读、记录彻底移除"""

    USER = "cleare"
    STORE = "清理数据测试店"

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            cls.USER, PASSWORD, store_name=cls.STORE)
        me = me_of(cls.member)
        cls.user_id = me["user_id"]
        cls.store_id = me["store_id"]
        seed_store_data(cls.member, "CLC")

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_c01_summary_preview(self):
        status, resp = self.admin.get(
            f"/api/users/{self.user_id}/store-data-summary")
        self.assertEqual(status, 200)
        data = resp["data"]
        self.assertEqual(data["username"], self.USER)
        self.assertEqual(len(data["stores"]), 1)
        store = data["stores"][0]
        self.assertEqual(store["store_name"], self.STORE)
        for key in ("orders", "supplier_costs", "ad_expenses",
                    "reconciliation_results", "profit_statements",
                    "import_records", "raw_records"):
            self.assertIsInstance(store["categories"][key], int)
            self.assertGreater(store["categories"][key], 0, f"{key} 应有记录")
        # 可清理业务数据不含 settings；新增其他费用后为 13 类
        self.assertEqual(len(store["categories"]), 13)
        self.assertNotIn("settings", store["categories"])

    def test_c02_delete_with_clear(self):
        status, resp = self.admin.post(
            f"/api/users/{self.user_id}/delete",
            {"confirm_username": self.USER, "clear_store_data": True})
        self.assertEqual(status, 200)
        data = resp["data"]
        self.assertTrue(data["clear_store_data"])
        self.assertEqual(data["store_ids"], [self.store_id])
        self.assertEqual(len(data["cleared"]), 1)
        cleared = data["cleared"][0]
        self.assertGreater(cleared["records_before"]["orders"], 0)
        self.assertIn("uploads_removed", cleared)
        self.assertEqual(len(data["backup"]), 1)
        self.assertTrue(data["backup"][0]["snapshot"]
                        .startswith(f"store_snapshot_{self.store_id}"))
        self.assertGreater(data["backup"][0]["files"], 0)

    def test_c03_snapshot_readable(self):
        snaps = [d for d in os.listdir(self.env.backup_dir)
                 if d.startswith(f"store_snapshot_{self.store_id}_")]
        self.assertEqual(len(snaps), 1)
        snap_data = os.path.join(self.env.backup_dir, snaps[0], "data")
        for name in BUSINESS_FILES:
            self.assertTrue(os.path.exists(
                os.path.join(snap_data, BUSINESS_FILES[name]["file"])),
                f"快照缺少 {name}")
        orders = read_json(os.path.join(snap_data, "orders.json"))
        self.assertGreater(len(orders["records"]), 0)
        settings = read_json(os.path.join(snap_data, "settings.json"))
        self.assertIn("settings", settings)

    def test_c04_store_records_removed(self):
        store_dir = os.path.join(self.env.data_dir, "stores", self.store_id)
        self.assertFalse(os.path.exists(store_dir))
        stores = read_json(os.path.join(self.env.data_dir, "stores.json"))["stores"]
        self.assertNotIn(self.store_id, [s["store_id"] for s in stores])
        users = read_json(os.path.join(self.env.data_dir, "users.json"))["users"]
        self.assertNotIn(self.USER, [u["username"] for u in users])

    def test_c05_cannot_login_after_delete(self):
        client = Client(self.env)
        status, resp = client.post("/api/auth/login",
                                   {"username": self.USER,
                                    "password": PASSWORD}, csrf=False)
        self.assertNotEqual(status, 200)
        self.assertNotEqual(resp.get("code"), 0)

    def test_c06_audit_log_written(self):
        entries = read_security_log(self.env)
        hits = [e for e in entries if e["event"] == "user.delete"
                and e["detail"]["target_username"] == self.USER]
        self.assertEqual(len(hits), 1)
        detail = hits[0]["detail"]
        self.assertTrue(detail["clear_store_data"])
        self.assertTrue(detail["backup"])
        self.assertIn("records_before", json.dumps(detail))
        line = json.dumps(hits[0], ensure_ascii=False).lower()
        self.assertNotIn("password", line)
        self.assertNotIn(PASSWORD.lower(), line)


class TestCrossStoreIsolation(unittest.TestCase):
    """测试 D + 业务回归: 清理 B 店铺后, C 店铺/admin/系统功能全部完好"""

    USER_B = "isob"
    USER_C = "isoc"
    STORE_B = "隔离被清店"
    STORE_C = "隔离旁观店"

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.b = cls.env.create_user_client(cls.USER_B, PASSWORD,
                                           store_name=cls.STORE_B)
        cls.c = cls.env.create_user_client(cls.USER_C, PASSWORD,
                                           store_name=cls.STORE_C)
        me_b, me_c = me_of(cls.b), me_of(cls.c)
        cls.b_user_id, cls.b_store_id = me_b["user_id"], me_b["store_id"]
        cls.c_user_id, cls.c_store_id = me_c["user_id"], me_c["store_id"]
        seed_store_data(cls.b, "ISB")
        seed_store_data(cls.c, "ISC")
        cls.c_orders_before = order_count(cls.c)
        cls.admin_users_before = 4  # admin + 承载成员 biz_carrier + B + C

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_d01_clear_b_data_only(self):
        """清数据是独立操作: 不删账号, 只清目标店铺"""
        status, resp = self.admin.post(
            f"/api/users/{self.b_user_id}/store-data-clear",
            {"confirm_store_name": self.STORE_B})
        self.assertEqual(status, 200)
        data = resp["data"]
        self.assertEqual(len(data["cleared"]), 1)
        self.assertEqual(data["cleared"][0]["store_id"], self.b_store_id)
        self.assertGreater(data["cleared"][0]["records_before"]["orders"], 0)
        self.assertTrue(data["backup"])

    def test_d02_b_account_survives_and_data_empty(self):
        """场景 C: 账号保留, 可重新登录, 数据为空, 可重新导入"""
        self.b = self.env.login_client(self.USER_B, PASSWORD)
        self.assertEqual(order_count(self.b), 0)
        status, resp = self.b.get("/api/summary?range=2026-03")
        self.assertEqual(status, 200)
        status, resp = self.b.upload_confirm(
            "/api/import/wechat-orders", "B重导.xlsx", make_wx_xlsx("ISB2"))
        self.assertEqual(status, 200)
        self.assertGreater(order_count(self.b), 0)

    def test_d03_c_store_intact(self):
        """C 店铺数据与会话完全不受影响"""
        self.assertEqual(order_count(self.c), self.c_orders_before)
        for name in ("orders", "supplier_costs", "ad_expenses",
                     "reconciliation_results", "profit_statements"):
            records = store_records(self.env, self.c_store_id, name)["records"]
            self.assertGreater(len(records), 0, f"C 店铺 {name} 数据丢失")
        status, resp = self.c.get("/api/auth/me")
        self.assertEqual(status, 200)
        status, resp = self.c.get("/api/profit?month=2026-03")
        self.assertEqual(status, 200)

    def test_d04_admin_and_system_intact(self):
        status, resp = self.admin.get("/api/users")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["total"], self.admin_users_before)
        status, resp = self.admin.get("/api/orders?page=1&size=20")
        self.assertEqual(status, 200)
        new = self.env.create_user_client("isod", PASSWORD)
        self.assertEqual(order_count(new), 0)

    def test_d05_clear_b_audit_log(self):
        entries = read_security_log(self.env)
        hits = [e for e in entries if e["event"] == "store.clear_data"
                and e["detail"]["target_username"] == self.USER_B]
        self.assertEqual(len(hits), 1)
        detail = hits[0]["detail"]
        self.assertEqual(detail["stores"][0]["store_id"], self.b_store_id)
        self.assertTrue(detail["backup"])


class TestPermissionProtection(unittest.TestCase):
    """测试 E: member 全接口 403; 管理员自我保护; 确认口令不匹配拒绝"""

    USER_B = "permb"
    USER_C = "permc"
    STORE_B = "权限B店"
    STORE_C = "权限C店"

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.b = cls.env.create_user_client(cls.USER_B, PASSWORD,
                                           store_name=cls.STORE_B)
        cls.c = cls.env.create_user_client(cls.USER_C, PASSWORD,
                                           store_name=cls.STORE_C)
        me_admin, me_b, me_c = (me_of(cls.admin), me_of(cls.b), me_of(cls.c))
        cls.admin_id = me_admin["user_id"]
        cls.b_user_id, cls.b_store_id = me_b["user_id"], me_b["store_id"]
        cls.c_user_id, cls.c_store_id = me_c["user_id"], me_c["store_id"]
        seed_store_data(cls.b, "PEB")
        seed_store_data(cls.c, "PEC")
        cls.b_orders_before = len(
            store_records(cls.env, cls.b_store_id, "orders")["records"])
        cls.c_orders_before = len(
            store_records(cls.env, cls.c_store_id, "orders")["records"])

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_e01_member_delete_forbidden(self):
        status, resp = self.b.post(
            f"/api/users/{self.c_user_id}/delete",
            {"confirm_username": self.USER_C, "clear_store_data": False})
        self.assertEqual(status, 403)

    def test_e02_member_clear_forbidden(self):
        status, resp = self.b.post(
            f"/api/users/{self.c_user_id}/store-data-clear",
            {"confirm_store_name": self.STORE_C})
        self.assertEqual(status, 403)

    def test_e03_member_summary_forbidden(self):
        status, resp = self.b.get(
            f"/api/users/{self.c_user_id}/store-data-summary")
        self.assertEqual(status, 403)

    def test_e04_member_user_list_forbidden(self):
        status, resp = self.b.get("/api/users")
        self.assertEqual(status, 403)

    def test_e05_admin_cannot_delete_self(self):
        status, resp = self.admin.post(
            f"/api/users/{self.admin_id}/delete",
            {"confirm_username": "admin", "clear_store_data": False})
        self.assertEqual(status, 400)
        self.assertIn("自己", resp["message"])
        status, resp = self.admin.get("/api/auth/me")
        self.assertEqual(status, 200)  # admin 会话与账号完好

    def test_e06_admin_cannot_clear_self(self):
        status, resp = self.admin.post(
            f"/api/users/{self.admin_id}/store-data-clear",
            {"confirm_store_name": ADMIN_STORE})
        self.assertEqual(status, 400)
        self.assertIn("自己", resp["message"])

    def test_e07_admin_cannot_reset_own_password(self):
        status, resp = self.admin.post(
            f"/api/users/{self.admin_id}/reset-password",
            {"new_password": "Whatever123"})
        self.assertEqual(status, 400)
        self.assertIn("修改密码", resp["message"])
        # 原会话仍有效（未被重置吊销）
        status, _ = self.admin.get("/api/auth/me")
        self.assertEqual(status, 200)

    def test_e08_admin_cannot_toggle_self(self):
        for action in ("disable", "enable"):
            status, resp = self.admin.post(
                f"/api/users/{self.admin_id}/status", {"action": action})
            self.assertEqual(status, 400, f"{action} 自己应被拒绝")
            self.assertIn("自己", resp["message"])

    def test_e09_confirm_username_mismatch_rejected(self):
        status, resp = self.admin.post(
            f"/api/users/{self.b_user_id}/delete",
            {"confirm_username": "错误的名字", "clear_store_data": False})
        self.assertEqual(status, 400)
        self.assertIn("用户名确认不一致", resp["message"])
        status, resp = self.admin.get("/api/users")
        names = [u["username"] for u in resp["data"]["users"]]
        self.assertIn(self.USER_B, names)  # B 账号未被动过

    def test_e10_confirm_store_name_mismatch_rejected(self):
        status, resp = self.admin.post(
            f"/api/users/{self.b_user_id}/store-data-clear",
            {"confirm_store_name": "错误的店铺"})
        self.assertEqual(status, 400)
        self.assertIn("店铺名称确认不一致", resp["message"])
        orders = store_records(self.env, self.b_store_id, "orders")
        self.assertEqual(len(orders["records"]), self.b_orders_before)

    def test_e11_nonexistent_user_404(self):
        status, resp = self.admin.post(
            "/api/users/u_ghost123/delete",
            {"confirm_username": "ghost", "clear_store_data": False})
        self.assertEqual(status, 404)
        self.assertIn("用户不存在", resp["message"])

    def test_e12_forged_ids_cannot_touch_other_store(self):
        """请求体伪造 user_id/store_id 字段: 后端按路径 user_id 解析, C 店铺不受影响"""
        status, resp = self.admin.post(
            f"/api/users/{self.b_user_id}/store-data-clear",
            {"confirm_store_name": self.STORE_B,
             "user_id": self.c_user_id, "store_id": self.c_store_id})
        self.assertEqual(status, 200)
        cleared_ids = [c["store_id"] for c in resp["data"]["cleared"]]
        self.assertEqual(cleared_ids, [self.b_store_id])
        self.assertEqual(
            len(store_records(self.env, self.c_store_id, "orders")["records"]),
            self.c_orders_before)


class TestFailureRecovery(unittest.TestCase):
    """测试 F: 快照失败/清理中断/删账号失败 → 明确失败 + 数据回滚, 绝不虚假成功"""

    USER = "failf"
    STORE = "故障恢复测试店"

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            cls.USER, PASSWORD, store_name=cls.STORE)
        me = me_of(cls.member)
        cls.user_id = me["user_id"]
        cls.store_id = me["store_id"]
        seed_store_data(cls.member, "FLR")
        cls.orders_before = len(
            store_records(cls.env, cls.store_id, "orders")["records"])

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _assert_untouched(self):
        """断言目标账号与数据回到操作前状态"""
        orders = store_records(self.env, self.store_id, "orders")
        self.assertEqual(len(orders["records"]), self.orders_before)
        status, resp = self.admin.get("/api/users")
        names = [u["username"] for u in resp["data"]["users"]]
        self.assertIn(self.USER, names)

    def test_f01_snapshot_failure_no_change(self):
        with mock.patch.object(
                self.env.server, "_snapshot_store",
                side_effect=OSError("disk full")):
            status, resp = self.admin.post(
                f"/api/users/{self.user_id}/delete",
                {"confirm_username": self.USER, "clear_store_data": True})
        self.assertEqual(status, 500)
        self.assertNotEqual(resp["code"], 0)
        self._assert_untouched()
        snaps = [d for d in os.listdir(self.env.backup_dir)
                 if d.startswith("store_snapshot_")]
        self.assertEqual(snaps, [])  # 无半成品快照残留

    def test_f02_clear_interrupted_rolls_back(self):
        def partial_fail(store_id, snapshot_dir):
            # 先真实清掉订单文件, 再抛错模拟清理中断
            bundle = self.env.server.registry.get(store_id)
            bundle.repo.write("orders", BUSINESS_FILES["orders"]["initial"](),
                              backup=False)
            raise ApiError("模拟清理中断", code=500, http_status=500)

        with mock.patch.object(self.env.server, "_clear_store_data_files",
                               side_effect=partial_fail):
            status, resp = self.admin.post(
                f"/api/users/{self.user_id}/delete",
                {"confirm_username": self.USER, "clear_store_data": True})
        self.assertEqual(status, 500)
        self.assertEqual(resp["code"], 500)
        self.assertIn("模拟清理中断", resp["message"])
        self._assert_untouched()

    def test_f03_delete_user_failure_restores_data(self):
        with mock.patch.object(
                self.env.server.user_manager, "delete_user",
                side_effect=OSError("disk full")):
            status, resp = self.admin.post(
                f"/api/users/{self.user_id}/delete",
                {"confirm_username": self.USER, "clear_store_data": True})
        self.assertEqual(status, 500)
        self.assertIn("恢复", resp["message"])
        self._assert_untouched()
        # 账号还在: 应能正常登录
        client = Client(self.env)
        status, _ = client.post("/api/auth/login",
                                {"username": self.USER,
                                 "password": PASSWORD}, csrf=False)
        self.assertEqual(status, 200)

    def test_f04_failures_never_report_success(self):
        """失败场景响应 code 非 0, 前端不会展示成功提示（汇总断言）"""
        for status_code in (500,):
            with mock.patch.object(
                    self.env.server, "_snapshot_store",
                    side_effect=OSError("disk full")):
                status, resp = self.admin.post(
                    f"/api/users/{self.user_id}/store-data-clear",
                    {"confirm_store_name": self.STORE})
                self.assertEqual(status, status_code)
                self.assertNotEqual(resp["code"], 0)


if __name__ == "__main__":
    unittest.main()

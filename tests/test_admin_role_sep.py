"""管理员角色与业务数据分离验收测试（spec 5.7 验收基准）

- TestAdminNoStore：验收 1/2 —— 管理员双路径创建均无业务店铺、
  无店管理员逐页全功能可用、member 对照与无店 member 登录拒绝
- TestTransferOwnership：验收 3/4/5/10/12 —— 店铺归属变更、待分配、
  孤儿店铺转交、备份凭证与原子性回滚（组 5）
- TestMemberFilter：验收 8/11 —— 成员+店铺双维筛选范围正确、
  列表/利润/导出三口径一致、member 参数越权 403（组 6）
- TestMigration：spec 5.6 历史数据迁移全流程 —— admin 遗留/孤儿样本、
  七步流程（发现→清点→备份→人工确认→执行→核对→审计）、失败回滚、
  禁删规则（组 8）
- TestPermissionMatrix：spec 5.4 权限矩阵 —— member 越权全扫描
  （他店 store_id/记录 ID/伪造身份）、admin 正向操作矩阵、管理接口
  越权与未登录拦截（组 8）
- TestFrontendAssets：spec 5.5 前端资产回归 —— 10 页面双维筛选资产、
  users.html 归属管理区块、index.html 目标级联、nav 10 项完整性、
  spec 5.7.2 验收 1-12 ↔ 用例映射表（组 8）
"""
import hashlib
import io
import json
import os
import re
import time
import unittest
from decimal import Decimal
from unittest import mock
from urllib.parse import quote

from openpyxl import load_workbook

from app.auth import UserManager
from app.data_schemas import BUSINESS_FILES
from tests.http_client import Client, TestEnv
from tests.test_auth_isolation import make_sup_xlsx, make_wx_xlsx
from tests.test_multi_store_admin import TestMultiStoreAdmin
from tests.test_page_routes import MGMT_ORDER, NAV_ORDER, PAGE_FEATURES

CARRIER_PASSWORD = "Carrier123456"
TR_PASSWORD = "Tr012345678"


def _read_users(env):
    with open(os.path.join(env.data_dir, "users.json"),
              encoding="utf-8") as f:
        return json.load(f).get("users", [])


def _read_stores(env):
    path = os.path.join(env.data_dir, "stores.json")
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return json.load(f).get("stores", [])


def _find_user(env, username):
    return next((u for u in _read_users(env)
                 if u.get("username") == username), None)


def _find_store(env, store_id):
    return next((s for s in _read_stores(env)
                 if s.get("store_id") == store_id), None)


class TestAdminNoStore(unittest.TestCase):
    """验收 1/2：管理员不拥有业务店铺，无店管理员全功能可用"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 验收 1a：setup/init 创建的管理员无店 ----------
    def test_01_setup_init_admin_has_no_store(self):
        env = TestEnv(initialize=False)
        try:
            # store_name 仅为兼容保留：服务端不得据此为管理员建店
            st, r = env.setup_admin(store_name="管理员店铺不应被创建")
            self.assertEqual(st, 200, r)
            self.assertIsNone(r["data"]["store_id"])
            self.assertIsNone(r["data"]["store_name"])
            user = _find_user(env, env.admin_user)
            self.assertEqual(user["store_ids"], [])
            self.assertEqual(user["role"], "admin")
            self.assertEqual(_read_stores(env), [])
            # 登录响应与 me 均为 null 店铺
            st, r = Client(env).post("/api/auth/login", {
                "username": env.admin_user,
                "password": env.admin_password}, csrf=False)
            self.assertEqual(st, 200)
            self.assertIsNone(r["data"]["store_id"])
            self.assertIsNone(r["data"]["store_name"])
            st, r = env.admin.get("/api/auth/me")
            self.assertEqual(st, 200)
            self.assertIsNone(r["data"]["store_id"])
        finally:
            env.cleanup()

    # ---------- 验收 1b：/api/users 创建 admin 角色同样无店 ----------
    def test_02_create_admin_role_no_store(self):
        stores_before = _read_stores(self.env)
        st, r = self.admin.post("/api/users", {
            "username": "admin2", "password": "Admin2123456",
            "role": "admin"})
        self.assertEqual(st, 200, r)
        self.assertIsNone(r["data"]["store_id"])
        self.assertIsNone(r["data"]["store_name"])
        self.assertEqual(_find_user(self.env, "admin2")["store_ids"], [])
        self.assertEqual(_read_stores(self.env), stores_before)
        # 新管理员无店登录合法
        st, r = Client(self.env).post("/api/auth/login", {
            "username": "admin2", "password": "Admin2123456"}, csrf=False)
        self.assertEqual(st, 200)
        self.assertIsNone(r["data"]["store_id"])

    # ---------- 验收 2：无店管理员逐页全功能可用 ----------
    def test_03_admin_no_store_all_pages_accessible(self):
        for page, feature in PAGE_FEATURES.items():
            st, body, _ = self.admin.get_page(page)
            self.assertEqual(st, 200, page)
            self.assertIn(feature, body, page)
        # 用户管理页（admin 专属）与代导入入口同样可访问
        for page in ("/users",):
            st, body, _ = self.admin.get_page(page)
            self.assertEqual(st, 200, page)
        # 核心管理接口无 500
        for api in ("/api/auth/me", "/api/stores", "/api/users",
                    "/api/health", "/api/orders?page=1&size=20",
                    "/api/summary?month=2026-03"):
            st, r = self.admin.get(api)
            self.assertEqual(st, 200, api)

    # ---------- 验收 1 对照：member 创建必有店铺 ----------
    def test_04_member_contrast_with_store(self):
        stores_before = _read_stores(self.env)
        data = self.env.create_user("member_contrast", "Memc123456",
                                    store_name="对照成员店铺")
        self.assertIsNotNone(data["store_id"])
        self.assertEqual(data["store_name"], "对照成员店铺")
        user = _find_user(self.env, "member_contrast")
        self.assertEqual(user["store_ids"], [data["store_id"]])
        store = _find_store(self.env, data["store_id"])
        self.assertEqual(store["owner_user_id"], user["user_id"])
        self.assertEqual(len(_read_stores(self.env)), len(stores_before) + 1)
        # member 登录后 me 返回本店
        c = self.env.login_client("member_contrast", "Memc123456")
        st, r = c.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["store_id"], data["store_id"])
        self.assertEqual(r["data"]["store_name"], "对照成员店铺")

    # ---------- 无店 member 登录被拒（独立环境，手工构造） ----------
    def test_05_member_without_store_login_denied(self):
        env = TestEnv()
        try:
            users_path = os.path.join(env.data_dir, "users.json")
            with open(users_path, encoding="utf-8") as f:
                data = json.load(f)
            for u in data["users"]:
                if u["username"] == "biz_carrier":
                    u["store_ids"] = []
            with open(users_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
            st, r = Client(env).post("/api/auth/login", {
                "username": "biz_carrier",
                "password": CARRIER_PASSWORD}, csrf=False)
            self.assertEqual(st, 401)
            self.assertEqual(r["message"], "账号未绑定店铺，请联系管理员")
        finally:
            env.cleanup()

    # ---------- 服务重启后无店管理员会话仍有效 ----------
    def test_06_restart_keeps_admin_session(self):
        st, _ = self.admin.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.env.restart()
        st, r = self.admin.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.assertIsNone(r["data"]["store_id"])
        st, r = self.admin.get("/api/orders?page=1&size=20")
        self.assertEqual(st, 200)


def _read_security_log(env):
    """读取全局安全日志（jsonl）→ 事件字典列表"""
    path = os.path.join(env.log_dir, "security_log.jsonl")
    entries = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                entries.append(json.loads(line))
    return entries


class TestTransferOwnership(unittest.TestCase):
    """验收 3/4/5/10/12：店铺归属变更、待分配、孤儿转交、备份与原子性"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member_a = cls.env.create_user_client(
            "tr_a", TR_PASSWORD, store_name="转出成员店")
        cls.member_b = cls.env.create_user_client(
            "tr_b", TR_PASSWORD, store_name="转入方占位店")
        cls.a_user_id = cls.env.get_user_id("tr_a")
        cls.b_user_id = cls.env.get_user_id("tr_b")
        cls.a_store = cls.member_a.get(
            "/api/auth/me")[1]["data"]["store_id"]
        cls.b_placeholder = cls.member_b.get(
            "/api/auth/me")[1]["data"]["store_id"]
        # A 店铺灌入完整业务数据
        st, _ = cls.member_a.upload_confirm(
            "/api/import/wechat-orders", "转出订单.xlsx", make_wx_xlsx("TR"))
        assert st == 200
        st, _ = cls.member_a.upload_confirm(
            "/api/import/supplier-costs", "转出成本.xlsx",
            make_sup_xlsx("TR"))
        assert st == 200
        st, _ = cls.member_a.post("/api/reconciliation/run",
                                  {"month": "2026-03"})
        assert st == 200
        st, _ = cls.member_a.post("/api/profit/recalculate",
                                  {"month": "2026-03"})
        assert st == 200
        # B 腾出店铺位: 占位店标记待分配, B 成为无店成员（会话随之吊销）
        st, r = cls.admin.post(
            f"/api/stores/{cls.b_placeholder}/mark-unassigned", {})
        assert st == 200, r
        cls.orphan_id = None

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _all_stores(self):
        st, r = self.admin.get("/api/admin/all-stores")
        self.assertEqual(st, 200)
        return {s["store_id"]: s for s in r["data"]["stores"]}

    # ---------- A3：店铺管理全量视图 ----------
    def test_01_all_stores_view(self):
        by_id = self._all_stores()
        self.assertEqual(len(by_id), 3)
        a = by_id[self.a_store]
        self.assertEqual(a["assignment_status"], "assigned")
        self.assertEqual(a["owner_username"], "tr_a")
        self.assertEqual(a["owner_user_id"], self.a_user_id)
        self.assertGreater(a["record_counts"]["orders"], 0)
        p = by_id[self.b_placeholder]
        self.assertEqual(p["assignment_status"], "unassigned")
        self.assertIsNone(p["owner_user_id"])
        self.assertIsNone(p["owner_username"])
        carrier = by_id[self.env.member_store]
        self.assertEqual(carrier["assignment_status"], "assigned")

    # ---------- A4：迁移清点十项指标 ----------
    def test_02_migration_overview(self):
        st, r = self.admin.get("/api/admin/migration/overview")
        self.assertEqual(st, 200)
        d = r["data"]
        self.assertEqual(d["admin_stores"], [])
        self.assertEqual(d["member_stores_count"], 2)
        # setUpClass 标记待分配时已产生备份目录，overview 应报告且真实存在
        self.assertIsNotNone(d["backup_dir"])
        self.assertTrue(os.path.isdir(d["backup_dir"]))
        un = {s["store_id"]: s for s in d["unassigned_stores"]}
        self.assertIn(self.b_placeholder, un)
        inv = un[self.b_placeholder]["inventory"]
        for key in ("orders", "refund_orders", "refund_details",
                    "refund_amount", "supplier_costs", "cost_amount",
                    "ad_expenses", "profit_results", "import_records",
                    "anomalies"):
            self.assertIn(key, inv)
        self.assertEqual(inv["orders"], 0)  # 占位店为空店
        self.assertEqual(inv["refund_amount"], "0.00")

    # ---------- 权限：member 调用全部 403 ----------
    def test_03_member_forbidden(self):
        checks = [
            self.member_a.get("/api/admin/all-stores"),
            self.member_a.get("/api/admin/migration/overview"),
            self.member_a.post(
                f"/api/stores/{self.a_store}/transfer-ownership",
                {"target_user_id": self.b_user_id}),
            self.member_a.post(
                f"/api/stores/{self.a_store}/mark-unassigned", {}),
        ]
        for st, r in checks:
            self.assertEqual(st, 403, r)
        entries = _read_security_log(self.env)
        denied = [e for e in entries if e["event"] == "access_denied"
                  and e.get("username") == "tr_a"]
        self.assertGreaterEqual(len(denied), 4)
        for e in denied:
            self.assertEqual(e["detail"].get("reason"), "admin_only")

    # ---------- 校验链：六类异常均拒绝且归属不变 ----------
    def test_04_validation_chain(self):
        base = f"/api/stores/{self.a_store}/transfer-ownership"
        # 店铺不存在 → 404
        st, r = self.admin.post("/api/stores/s_nope999/transfer-ownership",
                                {"target_user_id": self.b_user_id})
        self.assertEqual(st, 404)
        self.assertIn("店铺不存在", r["message"])
        # 目标用户不存在 → 404
        st, r = self.admin.post(base, {"target_user_id": "u_nope999"})
        self.assertEqual(st, 404)
        self.assertIn("目标用户不存在", r["message"])
        # 目标为 admin 角色 → 400（验收 10）
        st, r = self.admin.post(base, {
            "target_user_id": self.env.get_user_id(self.env.admin_user)})
        self.assertEqual(st, 400)
        self.assertIn("不能将店铺分配给管理员账号", r["message"])
        # 目标已有店铺 → 追加绑定成功（一对多归属模型, spec store_ownership）
        c2 = self.env.create_user("tr_c2", TR_PASSWORD,
                                  store_name="追加绑定C2店")
        c3 = self.env.create_user("tr_c3", TR_PASSWORD,
                                  store_name="追加绑定C3店")
        st, r = self.admin.post(
            f"/api/stores/{c3['store_id']}/transfer-ownership",
            {"target_user_id": c2["user_id"]})
        self.assertEqual(st, 200, r)
        u2 = _find_user(self.env, "tr_c2")
        self.assertEqual(len(u2["store_ids"]), 2)
        self.assertIn(c2["store_id"], u2["store_ids"])
        self.assertIn(c3["store_id"], u2["store_ids"])
        self.assertEqual(_find_store(self.env, c3["store_id"])["owner_user_id"],
                         c2["user_id"])
        # 幂等: 重复转给当前归属成员 → 200 且归属不变
        st, r = self.admin.post(
            f"/api/stores/{c3['store_id']}/transfer-ownership",
            {"target_user_id": c2["user_id"]})
        self.assertEqual(st, 200, r)
        self.assertEqual(len(_find_user(self.env, "tr_c2")["store_ids"]), 2)
        # 目标已停用 → 400
        c = self.env.create_user("tr_c", TR_PASSWORD, store_name="停用成员店")
        st, r = self.admin.post(f"/api/users/{c['user_id']}/status",
                                {"action": "disable"})
        self.assertEqual(st, 200)
        st, r = self.admin.post(base, {"target_user_id": c["user_id"]})
        self.assertEqual(st, 400)
        self.assertIn("目标成员已停用，请先启用", r["message"])
        # 缺 target_user_id → 400
        st, r = self.admin.post(base, {})
        self.assertEqual(st, 400)
        self.assertIn("目标成员", r["message"])
        # 归属不变：A 仍正常访问本店
        st, r = self.member_a.get("/api/orders?page=1&size=20")
        self.assertEqual(st, 200)

    # ---------- 正向流：转店后数据不变、A 失权 B 得权 ----------
    def test_05_transfer_positive_flow(self):
        st, r = self.member_a.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        orders_before = r["data"]["total"]
        st, r = self.member_a.get("/api/summary?month=2026-03")
        self.assertEqual(st, 200)
        summary_before = r["data"]
        store_data = os.path.join(self.env.data_dir, "stores",
                                  self.a_store, "data")
        files_before = {}
        for spec in BUSINESS_FILES.values():
            fp = os.path.join(store_data, spec["file"])
            if os.path.exists(fp):
                with open(fp, "rb") as f:
                    files_before[spec["file"]] = f.read()
        # 执行转店 A → B
        st, r = self.admin.post(
            f"/api/stores/{self.a_store}/transfer-ownership",
            {"target_user_id": self.b_user_id})
        self.assertEqual(st, 200, r)
        d = r["data"]
        self.assertEqual(d["store_id"], self.a_store)
        self.assertEqual(d["previous_owner"]["username"], "tr_a")
        self.assertEqual(d["new_owner"]["username"], "tr_b")
        self.assertGreater(d["record_counts"]["orders"], 0)
        self.assertTrue(d["data_preserved"])
        # store_id 与目录不变、业务文件字节级一致（验收 4）
        for fn, content in files_before.items():
            with open(os.path.join(store_data, fn), "rb") as f:
                self.assertEqual(f.read(), content, fn)
        # B 得权：重新登录并完整访问
        b = self.env.login_client("tr_b", TR_PASSWORD)
        st, r = b.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["store_id"], self.a_store)
        st, r = b.get("/api/orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], orders_before)
        st, r = b.get("/api/summary?month=2026-03")
        self.assertEqual(r["data"]["order_count"], summary_before["order_count"])
        self.assertEqual(r["data"]["actual_paid_amount"],
                         summary_before["actual_paid_amount"])
        # A 失权：旧会话立即失效、重登被无店拦截（验收 5）
        st, r = self.member_a.get("/api/orders?page=1&size=10")
        self.assertEqual(st, 401)
        st, r = Client(self.env).post(
            "/api/auth/login",
            {"username": "tr_a", "password": TR_PASSWORD}, csrf=False)
        self.assertEqual(st, 401)
        self.assertEqual(r["message"], "账号未绑定店铺，请联系管理员")

    # ---------- 备份凭证与审计日志（spec 4.5.3 / 5.6.1.7） ----------
    def test_06_backup_and_audit_log(self):
        dirs = [d for d in os.listdir(self.env.data_dir)
                if d.startswith("ownership_migration_")]
        self.assertTrue(dirs, "归属变更后应存在带时间戳的备份目录")
        latest = os.path.join(self.env.data_dir, sorted(dirs)[-1])
        self.assertTrue(os.path.isfile(os.path.join(latest, "users.json")))
        self.assertTrue(os.path.isfile(os.path.join(latest, "stores.json")))
        self.assertTrue(os.path.isdir(os.path.join(latest, "store_snapshot")))
        with open(os.path.join(latest, "users.json"), encoding="utf-8") as f:
            users = json.load(f)["users"]
        self.assertTrue(any(u["username"] == "tr_b" for u in users))
        # overview 报告最近一次备份目录
        st, r = self.admin.get("/api/admin/migration/overview")
        self.assertEqual(st, 200)
        self.assertIsNotNone(r["data"]["backup_dir"])
        # 安全日志六要素
        entries = _read_security_log(self.env)
        hits = [e for e in entries
                if e["event"] == "store.ownership_change"]
        self.assertTrue(hits)
        transfer_hits = [e for e in hits
                         if e["detail"].get("new_owner", {})
                         and isinstance(e["detail"]["new_owner"], dict)
                         and e["detail"]["new_owner"].get("username") == "tr_b"]
        self.assertTrue(transfer_hits)
        e = transfer_hits[0]
        detail = e["detail"]
        self.assertEqual(detail["store_id"], self.a_store)
        self.assertEqual(detail["previous_owner"]["username"], "tr_a")
        self.assertEqual(e.get("username"), self.env.admin_user)
        self.assertTrue(detail["backup_dir"])
        self.assertGreaterEqual(detail["revoked_sessions"], 1)

    # ---------- 待分配店不计入业务口径，但管理视图可见 ----------
    def test_07_unassigned_excluded_from_business_views(self):
        st, r = self.admin.get("/api/stores")
        self.assertEqual(st, 200)
        ids = {s["store_id"] for s in r["data"]["stores"]}
        self.assertNotIn(self.b_placeholder, ids)
        self.assertIn(self.a_store, ids)  # 现属 B
        by_id = self._all_stores()
        self.assertIn(self.b_placeholder, by_id)
        self.assertEqual(by_id[self.b_placeholder]["assignment_status"],
                         "unassigned")
        self.assertIn("orders", by_id[self.b_placeholder]["record_counts"])

    # ---------- 孤儿店铺自动归类 unassigned 且可转交 ----------
    def test_08_orphan_store_transferable(self):
        stores_path = os.path.join(self.env.data_dir, "stores.json")
        with open(stores_path, encoding="utf-8") as f:
            data = json.load(f)
        for s in data["stores"]:
            if s["store_id"] == self.b_placeholder:
                s["owner_user_id"] = "u_deleted99"
                s.pop("assignment_status", None)
        with open(stores_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        type(self).orphan_id = self.b_placeholder
        by_id = self._all_stores()
        orphan = by_id[self.orphan_id]
        self.assertEqual(orphan["assignment_status"], "unassigned")
        self.assertTrue(orphan.get("orphan"))
        self.assertIsNone(orphan["owner_username"])
        # 孤儿店铺经转店接口转交给无店成员 tr_a
        st, r = self.admin.post(
            f"/api/stores/{self.orphan_id}/transfer-ownership",
            {"target_user_id": self.a_user_id})
        self.assertEqual(st, 200, r)
        self.assertIsNone(r["data"]["previous_owner"])
        self.assertEqual(r["data"]["new_owner"]["username"], "tr_a")
        a = self.env.login_client("tr_a", TR_PASSWORD)
        st, r = a.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["store_id"], self.orphan_id)

    # ---------- 原子性：写入失败完全回滚、可重试成功（验收 12） ----------
    def test_09_atomic_rollback_and_retry(self):
        # 腾位：新成员 tr_e 的占位店标记待分配
        e = self.env.create_user("tr_e", TR_PASSWORD, store_name="原子性占位店")
        st, r = self.admin.post(
            f"/api/stores/{e['store_id']}/mark-unassigned", {})
        self.assertEqual(st, 200, r)
        # 打桩：写 users 表时模拟磁盘故障 → 服务层快照回滚
        with mock.patch.object(
                UserManager, "_mutate_users",
                side_effect=OSError("模拟磁盘写入失败")):
            st, r = self.admin.post(
                f"/api/stores/{self.orphan_id}/transfer-ownership",
                {"target_user_id": e["user_id"]})
        self.assertEqual(st, 500, r)
        self.assertIn("变更失败，请重试", r["message"])
        # 归属完全回滚：店铺记录仍指向 tr_a
        with open(os.path.join(self.env.data_dir, "stores.json"),
                  encoding="utf-8") as f:
            store = next(s for s in json.load(f)["stores"]
                         if s["store_id"] == self.orphan_id)
        self.assertEqual(store["owner_user_id"], self.a_user_id)
        a = self.env.login_client("tr_a", TR_PASSWORD)
        st, r = a.get("/api/auth/me")
        self.assertEqual(r["data"]["store_id"], self.orphan_id)
        # 可重试成功（无中间态）
        st, r = self.admin.post(
            f"/api/stores/{self.orphan_id}/transfer-ownership",
            {"target_user_id": e["user_id"]})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["new_owner"]["username"], "tr_e")


MF_PASSWORD = "Mf012345678"


def _xlsx_row_count(content):
    """解析导出 xlsx 字节 → 数据行数（去除表头）"""
    wb = load_workbook(io.BytesIO(content))
    return wb.active.max_row - 1


class TestMemberFilter(unittest.TestCase):
    """验收 8/11：成员+店铺双维筛选范围正确、三口径一致、member 参数越权 403

    环境布局：biz_carrier（空店）+ mf_a（甲店，MFA 数据）+ mf_b（乙店，MFB 数据）
    均为 2026-03 同构数据；管理员无店，三种范围切换仅经查询参数表达。
    """

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.mf_a = cls.env.create_user_client(
            "mf_a", MF_PASSWORD, store_name="筛选甲店")
        cls.mf_b = cls.env.create_user_client(
            "mf_b", MF_PASSWORD, store_name="筛选乙店")
        cls.a_user_id = cls.env.get_user_id("mf_a")
        cls.b_user_id = cls.env.get_user_id("mf_b")
        cls.admin_user_id = cls.env.get_user_id(cls.env.admin_user)
        cls.a_store = cls.mf_a.get("/api/auth/me")[1]["data"]["store_id"]
        cls.b_store = cls.mf_b.get("/api/auth/me")[1]["data"]["store_id"]
        for client, prefix in ((cls.mf_a, "MFA"), (cls.mf_b, "MFB")):
            st, _ = client.upload_confirm(
                "/api/import/wechat-orders", f"{prefix}订单.xlsx",
                make_wx_xlsx(prefix))
            assert st == 200
            st, _ = client.upload_confirm(
                "/api/import/supplier-costs", f"{prefix}成本.xlsx",
                make_sup_xlsx(prefix))
            assert st == 200
            st, _ = client.post("/api/reconciliation/run", {"month": "2026-03"})
            assert st == 200
            st, _ = client.post("/api/profit/recalculate", {"month": "2026-03"})
            assert st == 200
        # 基线：成员自身视角的订单/利润行数（carrier 店铺无业务数据）
        st, r = cls.mf_a.get("/api/orders?page=1&size=100&month=2026-03")
        cls.a_orders = r["data"]["total"]
        assert cls.a_orders > 0
        st, r = cls.mf_b.get("/api/orders?page=1&size=100&month=2026-03")
        cls.b_orders = r["data"]["total"]
        st, r = cls.mf_a.get("/api/profit?page=1&size=100&month=2026-03")
        cls.a_profit = r["data"]["total"]
        st, r = cls.mf_b.get("/api/profit?page=1&size=100&month=2026-03")
        cls.b_profit = r["data"]["total"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    @staticmethod
    def _with_member(api, member_id):
        sep = "&" if "?" in api else "?"
        return f"{api}{sep}member={member_id}"

    # ---------- 验收 8：admin 三种范围请求订单列表 ----------
    def test_01_admin_scope_orders(self):
        # 全部聚合 = 甲 + 乙（carrier 空店计入店铺数不计订单）
        st, r = self.admin.get("/api/orders?page=1&size=100&month=2026-03")
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["total"], self.a_orders + self.b_orders)
        self.assertEqual(r["data"]["store_count"], 3)
        self.assertEqual(r["data"]["store_scope"], "all")
        # member 收窄 = 该成员名下店铺
        st, r = self.admin.get(self._with_member(
            "/api/orders?page=1&size=100&month=2026-03", self.a_user_id))
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["total"], self.a_orders)
        self.assertEqual(r["data"]["store_count"], 1)
        self.assertEqual({o["store_id"] for o in r["data"]["orders"]},
                         {self.a_store})
        # store_id 单店 = 指定店铺
        st, r = self.admin.get(
            f"/api/orders?page=1&size=100&month=2026-03&store_id={self.b_store}")
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["total"], self.b_orders)
        self.assertNotIn("store_count", r["data"])
        self.assertEqual({o["store_id"] for o in r["data"]["orders"]},
                         {self.b_store})

    # ---------- 验收 8：总览/批次/日志三种范围 ----------
    def test_02_admin_scope_summary_batches_logs(self):
        # 总览：member 维度结果 == 该成员自身视角（同店同月）
        st, admin_r = self.admin.get(self._with_member(
            "/api/summary?month=2026-03", self.a_user_id))
        self.assertEqual(st, 200)
        self.assertEqual(admin_r["data"]["store_count"], 1)
        st, own_r = self.mf_a.get("/api/summary?month=2026-03")
        self.assertEqual(st, 200)
        for key in ("store_scope", "store_count"):
            admin_r["data"].pop(key, None)
        self.assertEqual(admin_r["data"], own_r["data"])
        # 总览全部聚合覆盖 3 个成员店铺
        st, r = self.admin.get("/api/summary?month=2026-03")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["store_count"], 3)
        # 批次：甲 2 + 乙 2；member 收窄后仅该店批次
        st, r = self.admin.get("/api/import/batches?page=1&size=50")
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["total"], 4)
        self.assertEqual(r["data"]["store_count"], 3)
        st, r = self.admin.get(self._with_member(
            "/api/import/batches?page=1&size=50", self.a_user_id))
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["total"], 2)
        self.assertEqual({b["store_id"] for b in r["data"]["batches"]},
                         {self.a_store})
        # 操作日志：member 收窄后条目全部来自该成员店铺
        st, r = self.admin.get(self._with_member(
            "/api/logs?page=1&size=100", self.a_user_id))
        self.assertEqual(st, 200, r)
        self.assertGreater(r["data"]["total"], 0)
        self.assertEqual({e["store_id"] for e in r["data"]["entries"]},
                         {self.a_store})
        # 日志全部聚合覆盖甲乙两店
        st, r = self.admin.get("/api/logs?page=1&size=100")
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["store_count"], 3)
        seen = {e["store_id"] for e in r["data"]["entries"]}
        self.assertIn(self.a_store, seen)
        self.assertIn(self.b_store, seen)

    # ---------- 验收 11：订单列表/利润结果/导出三口径一致（双身份） ----------
    def test_03_filter_caliber_consistency(self):
        def caliber(client, query):
            st, r = client.get(f"/api/orders?page=1&size=100&{query}")
            self.assertEqual(st, 200, r)
            orders_total = r["data"]["total"]
            st, r = client.get(f"/api/profit?page=1&size=100&{query}")
            self.assertEqual(st, 200, r)
            profit_total = r["data"]["total"]
            st, content, _ = client.get_binary(f"/api/orders/export?{query}")
            self.assertEqual(st, 200)
            self.assertEqual(_xlsx_row_count(content), orders_total)
            st, content, _ = client.get_binary(f"/api/profit/export?{query}")
            self.assertEqual(st, 200)
            self.assertEqual(_xlsx_row_count(content), profit_total)
            return orders_total, profit_total

        # admin 指定店铺 == 该店成员自身视角（同店同月三口径均一致）
        self.assertEqual(
            caliber(self.admin, f"store_id={self.a_store}&month=2026-03"),
            caliber(self.mf_a, "month=2026-03"))
        self.assertEqual(
            caliber(self.admin, f"store_id={self.b_store}&month=2026-03"),
            caliber(self.mf_b, "month=2026-03"))
        self.assertEqual(
            caliber(self.mf_a, "month=2026-03"), (self.a_orders, self.a_profit))
        self.assertEqual(
            caliber(self.mf_b, "month=2026-03"), (self.b_orders, self.b_profit))

    # ---------- member 携带 member 参数 → 全业务 API 403（验收 8 安全面） ----------
    def test_04_member_param_forbidden_all_apis(self):
        apis = [
            "/api/orders?page=1&size=20",
            "/api/orders/export?month=2026-03",
            "/api/refund-orders?page=1&size=20",
            "/api/refund-orders/export?month=2026-03",
            "/api/refund-details?page=1&size=20",
            "/api/refund-details/export?month=2026-03",
            "/api/supplier-costs?page=1&size=20",
            "/api/supplier-costs/export?month=2026-03",
            "/api/anomalies?page=1&size=20",
            "/api/anomalies/export?month=2026-03",
            "/api/anomalies/options",
            "/api/reconciliation?page=1&size=20",
            "/api/reconciliation/export?month=2026-03",
            "/api/reconciliation/options",
            "/api/summary?month=2026-03",
            "/api/profit?page=1&size=20",
            "/api/profit/export?month=2026-03",
            "/api/profit/options",
            "/api/ad-expenses?month=2026-03",
            "/api/data/orders",
            "/api/backups?name=orders",
            "/api/import/batches?page=1&size=20",
            "/api/logs?page=1&size=20",
        ]
        for api in apis:
            st, r = self.mf_a.get(self._with_member(api, self.b_user_id))
            self.assertEqual(st, 403, api)
            self.assertIn("无权使用成员筛选", r["message"])
        # 自身 user_id 同样拒绝（member 参数本身即越权）
        st, _ = self.mf_a.get(self._with_member(
            "/api/orders?page=1&size=20", self.a_user_id))
        self.assertEqual(st, 403)
        # 篡改 store_id 指向他店 → 403 兜底
        st, _ = self.mf_a.get(
            f"/api/orders?page=1&size=20&store_id={self.b_store}")
        self.assertEqual(st, 403)
        # 安全日志：member_filter_forbidden 且不泄露成员清单
        entries = _read_security_log(self.env)
        denied = [e for e in entries
                  if e["event"] == "access_denied"
                  and e.get("username") == "mf_a"
                  and e["detail"].get("reason") == "member_filter_forbidden"]
        self.assertGreaterEqual(len(denied), len(apis) + 1)
        for e in denied:
            self.assertNotIn("members", e["detail"])

    # ---------- admin member 参数校验链与兜底 ----------
    def test_05_admin_member_param_validation(self):
        # 目标成员不存在 → 404（成员校验先于店铺）
        st, r = self.admin.get(self._with_member(
            "/api/orders?page=1&size=20", "u_nope999"))
        self.assertEqual(st, 404)
        self.assertIn("目标用户不存在", r["message"])
        st, _ = self.admin.get(
            "/api/orders?page=1&size=20&member=u_nope999"
            f"&store_id={self.a_store}")
        self.assertEqual(st, 404)
        # 目标为 admin 角色 → 400
        st, r = self.admin.get(self._with_member(
            "/api/orders?page=1&size=20", self.admin_user_id))
        self.assertEqual(st, 400)
        self.assertIn("管理员不拥有业务店铺", r["message"])
        # member 与 store_id 同带但店铺不属于该成员 → 400
        st, r = self.admin.get(
            f"/api/orders?page=1&size=20&member={self.a_user_id}"
            f"&store_id={self.b_store}")
        self.assertEqual(st, 400)
        self.assertIn("目标店铺不属于该成员", r["message"])
        # 归属一致 → 200 且等于单店结果
        st, r = self.admin.get(
            f"/api/orders?page=1&size=100&month=2026-03"
            f"&member={self.a_user_id}&store_id={self.a_store}")
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["total"], self.a_orders)
        self.assertEqual({o["store_id"] for o in r["data"]["orders"]},
                         {self.a_store})

    # ---------- 性能抽测：三种范围切换劣化 ≤ 50ms（spec 4.1.1） ----------
    def test_06_scope_switch_performance(self):
        def timed(path, runs=3):
            best = None
            for _ in range(runs):
                t0 = time.perf_counter()
                st, _ = self.admin.get(path)
                self.assertEqual(st, 200, path)
                cost = time.perf_counter() - t0
                best = cost if best is None else min(best, cost)
            return best

        for base, member_q, store_q in (
                ("/api/orders?page=1&size=50&month=2026-03",
                 f"/api/orders?page=1&size=50&month=2026-03"
                 f"&member={self.a_user_id}",
                 f"/api/orders?page=1&size=50&month=2026-03"
                 f"&store_id={self.a_store}"),
                ("/api/summary?month=2026-03",
                 f"/api/summary?month=2026-03&member={self.a_user_id}",
                 f"/api/summary?month=2026-03&store_id={self.a_store}")):
            t_all = timed(base)
            t_member = timed(member_q)
            t_store = timed(store_q)
            self.assertLessEqual(t_member, t_all + 0.05)
            self.assertLessEqual(t_store, t_all + 0.05)


    # ---------- 写操作 member 维度（writeTarget 前端契约的后端支撑） ----------
    def test_07_write_targets_member_dimension(self):
        # 批量写操作：member 收窄 → 仅对该成员店铺执行（响应含 store_id 归属）
        st, r = self.admin.post("/api/reconciliation/run",
                                {"month": "2026-03", "member": self.a_user_id})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["store_id"], self.a_store)
        st, r = self.admin.post("/api/profit/recalculate",
                                {"month": "2026-03", "member": self.b_user_id})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["store_id"], self.b_store)
        # 单店写操作：仅携带 member → 定位该成员唯一店铺（广告费保存）
        st, r = self.admin.post("/api/ad-expenses", {
            "month": "2026-03", "amount": "12.34", "note": "成员维度代记",
            "member": self.a_user_id})
        self.assertEqual(st, 200, r)
        st, r = self.mf_a.get("/api/ad-expenses?month=2026-03")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 1)
        self.assertEqual(r["data"]["records"][0]["amount"], "12.34")
        st, r = self.mf_b.get("/api/ad-expenses?month=2026-03")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 0)
        # member 会话携带 member 键（body）→ 403
        st, _ = self.mf_a.post("/api/reconciliation/run",
                               {"month": "2026-03", "member": self.a_user_id})
        self.assertEqual(st, 403)
        st, _ = self.mf_a.post("/api/ad-expenses",
                               {"month": "2026-03", "amount": "1",
                                "member": self.b_user_id})
        self.assertEqual(st, 403)
        # admin 携带 member 与他店 target_store_id → 400
        st, r = self.admin.post("/api/reconciliation/run", {
            "month": "2026-03", "member": self.a_user_id,
            "target_store_id": self.b_store})
        self.assertEqual(st, 400)
        self.assertIn("目标店铺不属于该成员", r["message"])
        # 安全日志补录 body 侧 member_filter_forbidden
        entries = _read_security_log(self.env)
        denied = [e for e in entries
                  if e["event"] == "access_denied"
                  and e.get("username") == "mf_a"
                  and e["detail"].get("reason") == "member_filter_forbidden"]
        self.assertGreaterEqual(len(denied), 2)


MIG_PASSWORD = "Mig012345678"
PX_PASSWORD = "Px012345678"

STATIC_DIR = os.path.join(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))), "static")


def _hash_tree(root):
    """目录树内容指纹（相对路径 → 文件字节 hash）"""
    items = {}
    for dirpath, _dirs, files in os.walk(root):
        for fn in files:
            fp = os.path.join(dirpath, fn)
            with open(fp, "rb") as f:
                items[os.path.relpath(fp, root)] = \
                    hashlib.sha256(f.read()).hexdigest()
    return items


def _store_records(env, store_id, repo_name):
    """直读店铺业务文件记录（清点对比的文件侧真值）"""
    fp = os.path.join(env.data_dir, "stores", store_id, "data",
                      BUSINESS_FILES[repo_name]["file"])
    if not os.path.exists(fp):
        return []
    with open(fp, encoding="utf-8") as f:
        return json.load(f).get("records", [])


def _sum_field(records, field):
    total = Decimal("0")
    for r in records:
        total += Decimal(str(r.get(field) or "0"))
    return str(total.quantize(Decimal("0.01")))


class TestMigration(unittest.TestCase):
    """spec 5.6 历史数据迁移全规则：七步流程 + 解绑 + 失败回滚 + 禁删

    样本（隔离临时目录，模拟历史遗留状态，spec 5.6.1.12）：
    - mig_m1/mig_m2：灌入业务数据后手工改造为 admin 归属（admin_legacy）
    - mig_orphan：店铺 owner 指向已删用户（孤儿 → unassigned）
    - biz_carrier：正常成员店铺；mig_target：无店成员（接收方）
    """

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        m1 = cls.env.create_user_client("mig_m1", MIG_PASSWORD,
                                        store_name="管理员遗留一店")
        m2 = cls.env.create_user_client("mig_m2", MIG_PASSWORD,
                                        store_name="管理员遗留二店")
        orphan = cls.env.create_user_client("mig_orphan", MIG_PASSWORD,
                                            store_name="孤儿样本店")
        # m1 完整业务数据（订单/成本/对账/利润/广告费）
        st, _ = m1.upload_confirm("/api/import/wechat-orders",
                                  "M1订单.xlsx", make_wx_xlsx("M1A"))
        assert st == 200
        st, _ = m1.upload_confirm("/api/import/supplier-costs",
                                  "M1成本.xlsx", make_sup_xlsx("M1A"))
        assert st == 200
        st, _ = m1.post("/api/reconciliation/run", {"month": "2026-03"})
        assert st == 200
        st, _ = m1.post("/api/profit/recalculate", {"month": "2026-03"})
        assert st == 200
        st, _ = m1.post("/api/ad-expenses", {
            "month": "2026-03", "amount": "88.50", "note": "遗留广告费"})
        assert st == 200
        # m2 / 孤儿店铺少量订单
        st, _ = m2.upload_confirm("/api/import/wechat-orders",
                                  "M2订单.xlsx", make_wx_xlsx("M2A"))
        assert st == 200
        st, _ = orphan.upload_confirm("/api/import/wechat-orders",
                                      "M3订单.xlsx", make_wx_xlsx("M3A"))
        assert st == 200
        cls.m1_store = m1.get("/api/auth/me")[1]["data"]["store_id"]
        cls.m2_store = m2.get("/api/auth/me")[1]["data"]["store_id"]
        cls.orphan_store = orphan.get("/api/auth/me")[1]["data"]["store_id"]
        # 迁移前业务基线（从原归属成员视角）
        cls.m1_orders = m1.get(
            "/api/orders?page=1&size=100&month=2026-03")[1]["data"]["total"]
        cls.m1_costs = m1.get(
            "/api/supplier-costs?page=1&size=100&month=2026-03")[1]["data"]["total"]
        cls.m1_profit = m1.get(
            "/api/profit?page=1&size=100&month=2026-03")[1]["data"]["total"]
        cls.m1_batches = m1.get(
            "/api/import/batches?page=1&size=50")[1]["data"]["total"]
        cls.m1_ads = m1.get(
            "/api/ad-expenses?month=2026-03")[1]["data"]["total"]
        cls.m1_summary = m1.get(
            "/api/summary?month=2026-03")[1]["data"]
        # 接收方成员：占位店标记待分配 → mig_target 无店
        target = cls.env.create_user("mig_target", MIG_PASSWORD,
                                     store_name="接收占位店")
        st, r = cls.admin.post(
            f"/api/stores/{target['store_id']}/mark-unassigned", {})
        assert st == 200, r
        cls.target_user_id = target["user_id"]
        # 手工改造历史遗留状态：admin 绑 m1/m2 两店、孤儿店铺 owner 悬空
        cls.admin_user_id = cls.env.get_user_id(cls.env.admin_user)
        users_path = os.path.join(cls.env.data_dir, "users.json")
        stores_path = os.path.join(cls.env.data_dir, "stores.json")
        with open(users_path, encoding="utf-8") as f:
            users_data = json.load(f)
        for u in users_data["users"]:
            if u["username"] == cls.env.admin_user:
                u["store_ids"] = [cls.m1_store, cls.m2_store]
            elif u["username"] in ("mig_m1", "mig_m2"):
                u["store_ids"] = []
        with open(users_path, "w", encoding="utf-8") as f:
            json.dump(users_data, f, ensure_ascii=False, indent=2)
        with open(stores_path, encoding="utf-8") as f:
            stores_data = json.load(f)
        for s in stores_data["stores"]:
            if s["store_id"] in (cls.m1_store, cls.m2_store):
                s["owner_user_id"] = cls.admin_user_id
            elif s["store_id"] == cls.orphan_store:
                s["owner_user_id"] = "u_deleted_ghost"
        with open(stores_path, "w", encoding="utf-8") as f:
            json.dump(stores_data, f, ensure_ascii=False, indent=2)
        cls.env.restart()
        cls.admin = cls.env.login_client(cls.env.admin_user,
                                         cls.env.admin_password)
        # 文件指纹与店铺目录清单基线（迁移前后数据不变/禁删断言）
        cls.m1_files = _hash_tree(cls._business_root(cls.m1_store))
        cls.m2_files = _hash_tree(cls._business_root(cls.m2_store))
        cls.orphan_files = _hash_tree(cls._business_root(cls.orphan_store))
        cls.stores_snapshot = sorted(os.listdir(
            os.path.join(cls.env.data_dir, "stores")))

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    @classmethod
    def _store_root(cls, store_id):
        return os.path.join(cls.env.data_dir, "stores", store_id)

    @classmethod
    def _business_root(cls, store_id):
        """业务数据目录（数据不变的验收口径：仅 data/ 业务文件，
        不含 op_log 审计日志/备份等运维产物）"""
        return os.path.join(cls._store_root(store_id), "data")

    def _ensure_admin(self):
        """归属变更会吊销原归属（admin）会话 → 按需重建"""
        st, _ = self.admin.get("/api/auth/me")
        if st != 200:
            type(self).admin = self.env.login_client(
                self.env.admin_user, self.env.admin_password)

    def _migration_dirs(self):
        return sorted(d for d in os.listdir(self.env.data_dir)
                      if d.startswith("ownership_migration_"))

    # ---------- 步骤一：发现（admin_legacy / 孤儿自动归类） ----------
    def test_01_discovery_admin_legacy_and_orphan(self):
        st, r = self.admin.get("/api/admin/all-stores")
        self.assertEqual(st, 200)
        by_id = {s["store_id"]: s for s in r["data"]["stores"]}
        for sid in (self.m1_store, self.m2_store):
            self.assertEqual(by_id[sid]["assignment_status"], "admin_legacy",
                             sid)
            self.assertEqual(by_id[sid]["owner_username"],
                             self.env.admin_user, sid)
        o = by_id[self.orphan_store]
        self.assertEqual(o["assignment_status"], "unassigned")
        self.assertTrue(o["orphan"])
        self.assertEqual(by_id[self.env.member_store]["assignment_status"],
                         "assigned")
        # overview：admin 遗留清单 + 待分配/孤儿清单 + 成员店计数
        st, r = self.admin.get("/api/admin/migration/overview")
        self.assertEqual(st, 200)
        d = r["data"]
        self.assertEqual({s["store_id"] for s in d["admin_stores"]},
                         {self.m1_store, self.m2_store})
        for s in d["admin_stores"]:
            self.assertEqual(len(s["inventory"]), 10, s["store_id"])
        un = {s["store_id"] for s in d["unassigned_stores"]}
        self.assertIn(self.orphan_store, un)
        # admin 遗留店与待分配店不计入成员店铺（验收 7：admin 不入统计）
        self.assertEqual(d["member_stores_count"], 1)
        self.assertTrue(os.path.isdir(d["backup_dir"]))

    # ---------- 步骤二：清点（十项指标与店铺文件逐项一致） ----------
    def test_02_inventory_matches_store_files(self):
        st, r = self.admin.get("/api/admin/migration/overview")
        self.assertEqual(st, 200)
        inv = next(s["inventory"] for s in r["data"]["admin_stores"]
                   if s["store_id"] == self.m1_store)
        refund_details = _store_records(self.env, self.m1_store,
                                        "refund_details")
        costs = _store_records(self.env, self.m1_store, "supplier_costs")
        self.assertEqual(inv["orders"], len(
            _store_records(self.env, self.m1_store, "orders")))
        self.assertEqual(inv["refund_orders"], len(
            _store_records(self.env, self.m1_store, "refund_orders")))
        self.assertEqual(inv["refund_details"], len(refund_details))
        self.assertEqual(inv["refund_amount"],
                         _sum_field(refund_details, "refund_amount"))
        self.assertEqual(inv["supplier_costs"], len(costs))
        self.assertEqual(inv["cost_amount"],
                         _sum_field(costs, "total_amount"))
        self.assertEqual(inv["ad_expenses"], len(
            _store_records(self.env, self.m1_store, "ad_expenses")))
        self.assertEqual(inv["profit_results"], len(
            _store_records(self.env, self.m1_store, "profit_statements")))
        self.assertEqual(inv["import_records"], len(
            _store_records(self.env, self.m1_store, "import_records")))
        self.assertEqual(inv["anomalies"], len(
            _store_records(self.env, self.m1_store, "anomalies")))
        # 孤儿店铺清点同样与文件一致
        oi = next(s["inventory"] for s in r["data"]["unassigned_stores"]
                  if s["store_id"] == self.orphan_store)
        self.assertGreater(oi["orders"], 0)
        self.assertEqual(oi["orders"], len(
            _store_records(self.env, self.orphan_store, "orders")))

    # ---------- 步骤三/四：只读操作零归属变更（无自动分配路径） ----------
    def test_03_readonly_operations_no_ownership_change(self):
        users_fp = os.path.join(self.env.data_dir, "users.json")
        stores_fp = os.path.join(self.env.data_dir, "stores.json")
        with open(users_fp, "rb") as f:
            users_bytes = f.read()
        with open(stores_fp, "rb") as f:
            stores_bytes = f.read()
        for st, r in (self.admin.get("/api/auth/me"),
                      self.admin.get("/api/admin/all-stores"),
                      self.admin.get("/api/admin/migration/overview"),
                      self.admin.get("/api/orders?page=1&size=20"),
                      self.admin.get(f"/api/summary?month=2026-03"
                                     f"&store_id={self.m1_store}"),
                      self.admin.get("/api/stores")):
            self.assertEqual(st, 200)
        # 服务重启亦不触发任何自动迁移/自动分配（会话持久化，无需重登：
        # 重登会写 users.json 的登录时间戳，与归属变更无关但干扰字节对比）
        self.env.restart()
        st, _ = self.admin.get("/api/admin/migration/overview")
        self.assertEqual(st, 200)
        with open(users_fp, "rb") as f:
            self.assertEqual(f.read(), users_bytes)
        with open(stores_fp, "rb") as f:
            self.assertEqual(f.read(), stores_bytes)
        self.assertEqual(_hash_tree(self._business_root(self.m1_store)),
                         self.m1_files)

    # ---------- 失败回滚一：备份失败 → 中止且零归属变更（验收 12） ----------
    def test_04_backup_failure_aborts_with_zero_change(self):
        self._ensure_admin()
        before = self._migration_dirs()
        os.chmod(self.env.data_dir, 0o555)
        try:
            st, r = self.admin.post(
                f"/api/stores/{self.m1_store}/transfer-ownership",
                {"target_user_id": self.target_user_id})
        finally:
            os.chmod(self.env.data_dir, 0o755)
        self.assertEqual(st, 500, r)
        self.assertIn("迁移前备份失败，已中止且未做任何变更", r["message"])
        # 归属零变更：店铺仍归 admin、接收方仍无店、无新增备份目录
        self.assertEqual(_find_store(self.env, self.m1_store)
                         ["owner_user_id"], self.admin_user_id)
        self.assertEqual(_find_user(self.env, "mig_target")["store_ids"], [])
        self.assertEqual(_find_user(self.env, self.env.admin_user)
                         ["store_ids"], [self.m1_store, self.m2_store])
        self.assertEqual(self._migration_dirs(), before)

    # ---------- 失败回滚二：核对不一致 → 自动回滚两表 + 差异项（验收 12） ----------
    def test_05_verification_mismatch_rolls_back_both_tables(self):
        self._ensure_admin()
        real = self.env.server._store_inventory
        calls = {"n": 0}

        def fake(store_id):
            calls["n"] += 1
            inv = real(store_id)
            if calls["n"] == 2:  # 第二次 = 迁移后核对 → 模拟数据不一致
                inv = dict(inv, orders=inv["orders"] + 999)
            return inv

        with mock.patch.object(self.env.server, "_store_inventory",
                               side_effect=fake):
            st, r = self.admin.post(
                f"/api/stores/{self.m1_store}/transfer-ownership",
                {"target_user_id": self.target_user_id})
        self.assertEqual(st, 500, r)
        self.assertIn("迁移核对失败，已从备份恢复店铺归属", r["message"])
        self.assertEqual(calls["n"], 2)
        # 两表回滚：店铺仍归 admin、接收方仍无店
        self.assertEqual(_find_store(self.env, self.m1_store)
                         ["owner_user_id"], self.admin_user_id)
        self.assertEqual(_find_user(self.env, "mig_target")["store_ids"], [])
        self.assertEqual(_find_user(self.env, self.env.admin_user)
                         ["store_ids"], [self.m1_store, self.m2_store])
        # 审计：result=error + 差异项 + 备份位置（供人工核对）
        entries = _read_security_log(self.env)
        hits = [e for e in entries
                if e["event"] == "store.ownership_change"
                and e.get("result") == "error"
                and e["detail"].get("store_id") == self.m1_store]
        self.assertTrue(hits)
        detail = hits[-1]["detail"]
        self.assertTrue(detail["recovered"])
        self.assertEqual(detail["diff"]["orders"][1],
                         detail["diff"]["orders"][0] + 999)
        self.assertTrue(detail["backup_dir"])
        self.assertTrue(os.path.isdir(detail["backup_dir"]))

    # ---------- 步骤五/六：执行 A1 整店转交 + 核对（验收 3/4/5） ----------
    def test_06_execute_transfer_a1_full_store_handover(self):
        self._ensure_admin()
        st, r = self.admin.post(
            f"/api/stores/{self.m1_store}/transfer-ownership",
            {"target_user_id": self.target_user_id})
        self.assertEqual(st, 200, r)
        d = r["data"]
        self.assertEqual(d["store_id"], self.m1_store)
        self.assertEqual(d["previous_owner"]["username"], self.env.admin_user)
        self.assertEqual(d["new_owner"]["username"], "mig_target")
        self.assertGreater(d["record_counts"]["orders"], 0)
        self.assertTrue(d["data_preserved"])
        # 数据不变：店铺目录字节级一致、store_id 与目录不变（验收 4）
        self.assertEqual(_hash_tree(self._business_root(self.m1_store)),
                         self.m1_files)
        # 接收方得权且数据完整（验收 3/5）
        target = self.env.login_client("mig_target", MIG_PASSWORD)
        st, r = target.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["store_id"], self.m1_store)
        st, r = target.get("/api/orders?page=1&size=100&month=2026-03")
        self.assertEqual(r["data"]["total"], self.m1_orders)
        st, r = target.get("/api/summary?month=2026-03")
        self.assertEqual(r["data"]["order_count"],
                         self.m1_summary["order_count"])
        self.assertEqual(r["data"]["actual_paid_amount"],
                         self.m1_summary["actual_paid_amount"])

    # ---------- 迁移后范围：admin 遗留店不入聚合（验收 7） ----------
    def test_07_after_transfer_scope_and_partial_unbind(self):
        self._ensure_admin()
        # admin 仅剩 m2；m1 已归成员
        self.assertEqual(_find_user(self.env, self.env.admin_user)
                         ["store_ids"], [self.m2_store])
        st, r = self.admin.get("/api/admin/migration/overview")
        self.assertEqual(st, 200)
        self.assertEqual({s["store_id"] for s in r["data"]["admin_stores"]},
                         {self.m2_store})
        self.assertEqual(r["data"]["member_stores_count"], 2)
        # 聚合仅成员店铺：carrier + m1；不含 m2（admin 遗留）与待分配/孤儿
        st, r = self.admin.get("/api/orders?page=1&size=100&month=2026-03")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["store_count"], 2)
        # carrier 为空店（无订单记录），实际出现的店铺仅 m1
        self.assertEqual({o["store_id"] for o in r["data"]["orders"]},
                         {self.m1_store})
        self.assertEqual(r["data"]["total"], self.m1_orders)

    # ---------- 步骤七：审计日志 + 备份凭证（spec 5.6.1.7） ----------
    def test_08_audit_log_and_backup_evidence(self):
        self._ensure_admin()
        entries = _read_security_log(self.env)
        hits = [e for e in entries
                if e["event"] == "store.ownership_change"
                and e.get("result") != "error"
                and e["detail"].get("store_id") == self.m1_store]
        self.assertTrue(hits)
        e = hits[-1]
        detail = e["detail"]
        self.assertEqual(e.get("username"), self.env.admin_user)
        self.assertEqual(detail["previous_owner"]["username"],
                         self.env.admin_user)
        self.assertEqual(detail["new_owner"]["username"], "mig_target")
        self.assertGreater(detail["record_counts"]["orders"], 0)
        self.assertTrue(detail["backup_dir"])
        self.assertGreaterEqual(detail["revoked_sessions"], 1)
        # 备份目录：两表副本（迁移前快照）+ 整店快照（含全部业务文件）
        backup_dir = detail["backup_dir"]
        with open(os.path.join(backup_dir, "users.json"),
                  encoding="utf-8") as f:
            snap = json.load(f)["users"]
        admin_snap = next(u for u in snap
                          if u["username"] == self.env.admin_user)
        self.assertEqual(admin_snap["store_ids"],
                         [self.m1_store, self.m2_store])
        snapshot_dir = os.path.join(backup_dir, "store_snapshot")
        self.assertTrue(os.path.isdir(snapshot_dir))
        for rel in self.m1_files:
            self.assertTrue(
                os.path.isfile(os.path.join(snapshot_dir, "data", rel)), rel)

    # ---------- 步骤五（A2）：标记待分配 + admin 完全解绑 ----------
    def test_09_execute_a2_mark_unassigned_full_unbind(self):
        self._ensure_admin()
        st, r = self.admin.post(
            f"/api/stores/{self.m2_store}/mark-unassigned", {})
        self.assertEqual(st, 200, r)
        d = r["data"]
        self.assertEqual(d["assignment_status"], "unassigned")
        self.assertEqual(d["previous_owner"]["username"], self.env.admin_user)
        self.assertGreater(d["record_counts"]["orders"], 0)
        # m2 数据完整保留（禁删规则）
        self.assertEqual(_hash_tree(self._business_root(self.m2_store)),
                         self.m2_files)
        store = _find_store(self.env, self.m2_store)
        self.assertIsNone(store["owner_user_id"])
        self.assertEqual(store["assignment_status"], "unassigned")
        self.assertEqual(store["previous_owner_user_id"], self.admin_user_id)
        # admin 解绑：store_ids 清空、无店登录与管理正常（验收 1/2 复核）
        self.assertEqual(_find_user(self.env, self.env.admin_user)
                         ["store_ids"], [])
        type(self).admin = self.env.login_client(self.env.admin_user,
                                                 self.env.admin_password)
        st, r = self.admin.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.assertIsNone(r["data"]["store_id"])
        # 待迁移标识消失：全库无 admin_legacy
        st, r = self.admin.get("/api/admin/all-stores")
        self.assertEqual(st, 200)
        self.assertFalse([s for s in r["data"]["stores"]
                          if s["assignment_status"] == "admin_legacy"])
        st, r = self.admin.get("/api/admin/migration/overview")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["admin_stores"], [])
        un = {s["store_id"]: s for s in r["data"]["unassigned_stores"]}
        self.assertIn(self.m2_store, un)
        self.assertEqual(un[self.m2_store]["inventory"]["orders"], len(
            _store_records(self.env, self.m2_store, "orders")))
        # 聚合仅成员店铺（carrier 空店无订单，实际仅 m1 出现）
        st, r = self.admin.get("/api/orders?page=1&size=100&month=2026-03")
        self.assertEqual(st, 200)
        self.assertEqual({o["store_id"] for o in r["data"]["orders"]},
                         {self.m1_store})

    # ---------- 全流程收口：三类店铺历史数据完整存在（禁删规则） ----------
    def test_10_all_records_preserved_after_full_flow(self):
        target = self.env.login_client("mig_target", MIG_PASSWORD)
        for api, expected in (
                ("/api/orders?page=1&size=100&month=2026-03",
                 self.m1_orders),
                ("/api/supplier-costs?page=1&size=100&month=2026-03",
                 self.m1_costs),
                ("/api/profit?page=1&size=100&month=2026-03",
                 self.m1_profit),
                ("/api/import/batches?page=1&size=50", self.m1_batches),
                ("/api/ad-expenses?month=2026-03", self.m1_ads)):
            st, r = target.get(api)
            self.assertEqual(st, 200, api)
            self.assertEqual(r["data"]["total"], expected, api)
        # m2（待分配）/ 孤儿店铺：文件字节级完整保留
        self.assertEqual(_hash_tree(self._business_root(self.m2_store)),
                         self.m2_files)
        self.assertEqual(_hash_tree(self._business_root(self.orphan_store)),
                         self.orphan_files)
        # 店铺目录无删除
        self.assertEqual(sorted(os.listdir(
            os.path.join(self.env.data_dir, "stores"))),
            self.stores_snapshot)


class TestPermissionMatrix(unittest.TestCase):
    """spec 5.4 权限矩阵：member 越权全扫描 / admin 正向矩阵 / 接口防护"""

    READ_APIS = [
        "/api/orders?page=1&size=20",
        "/api/orders/export?month=2026-03",
        "/api/refund-orders?page=1&size=20",
        "/api/refund-orders/export?month=2026-03",
        "/api/refund-details?page=1&size=20",
        "/api/refund-details/export?month=2026-03",
        "/api/supplier-costs?page=1&size=20",
        "/api/supplier-costs/export?month=2026-03",
        "/api/anomalies?page=1&size=20",
        "/api/anomalies/export?month=2026-03",
        "/api/anomalies/options",
        "/api/reconciliation?page=1&size=20",
        "/api/reconciliation/export?month=2026-03",
        "/api/reconciliation/options",
        "/api/summary?month=2026-03",
        "/api/profit?page=1&size=20",
        "/api/profit/export?month=2026-03",
        "/api/profit/options",
        "/api/ad-expenses?month=2026-03",
        "/api/data/orders",
        "/api/backups?name=orders",
        "/api/import/batches?page=1&size=20",
        "/api/logs?page=1&size=20",
    ]

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.a = cls.env.create_user_client("px_a", PX_PASSWORD,
                                           store_name="矩阵甲店")
        cls.b = cls.env.create_user_client("px_b", PX_PASSWORD,
                                           store_name="矩阵乙店")
        cls.a_store = cls.a.get("/api/auth/me")[1]["data"]["store_id"]
        cls.b_store = cls.b.get("/api/auth/me")[1]["data"]["store_id"]
        cls.a_user_id = cls.env.get_user_id("px_a")
        cls.b_user_id = cls.env.get_user_id("px_b")
        cls.admin_user_id = cls.env.get_user_id(cls.env.admin_user)
        for client, prefix in ((cls.a, "PXA"), (cls.b, "PXB")):
            st, _ = client.upload_confirm(
                "/api/import/wechat-orders", f"{prefix}订单.xlsx",
                make_wx_xlsx(prefix))
            assert st == 200
        st, r = cls.admin.get(
            f"/api/import/batches?store_id={cls.b_store}"
            "&source_type=wechat_orders&page=1&page_size=20")
        assert st == 200
        cls.b_batch_id = r["data"]["batches"][0]["batch_id"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- member 越权扫描：全部读 API 携带他店 store_id（验收 9） ----------
    def test_01_member_cross_store_reads_forbidden(self):
        for api in self.READ_APIS:
            sep = "&" if "?" in api else "?"
            st, r = self.a.get(f"{api}{sep}store_id={self.b_store}")
            self.assertEqual(st, 403, api)
            self.assertIn("无权访问该店铺数据", r["message"], api)
            self.assertFalse(r.get("data"), api)  # 无任何数据泄露
        entries = _read_security_log(self.env)
        denied = [e for e in entries if e["event"] == "access_denied"
                  and e.get("username") == "px_a"
                  and e["detail"].get("reason") == "store_scope_forbidden"]
        self.assertGreaterEqual(len(denied), len(self.READ_APIS))

    # ---------- member 越权扫描：写操作他店目标/记录 ID（验收 9） ----------
    def test_02_member_cross_store_writes_forbidden(self):
        for path, body in (
                ("/api/ad-expenses", {"month": "2026-03",
                                      "amount": "1.00",
                                      "target_store_id": self.b_store}),
                ("/api/reconciliation/run", {"month": "2026-03",
                                             "target_store_id": self.b_store}),
                ("/api/profit/recalculate", {"month": "2026-03",
                                             "target_store_id": self.b_store}),
                ("/api/data/orders", {"records": [],
                                      "target_store_id": self.b_store}),
                ("/api/reconciliation/manual-match",
                 {"target_store_id": self.b_store}),
                ("/api/profit/config", {"target_store_id": self.b_store}),
                ("/api/backup", {"target_store_id": self.b_store}),
                ("/api/restore", {"target_store_id": self.b_store})):
            st, r = self.a.post(path, body)
            self.assertEqual(st, 403, path)
            self.assertIn("无权操作该店铺", r["message"], path)
        # 导入确认指向他店：_import_target 契约专用文案（无权上传）
        st, r = self.a.post("/api/import/confirm", {
            "upload_id": "any", "target_store_id": self.b_store})
        self.assertEqual(st, 403)
        self.assertIn("无权上传到该店铺", r["message"])
        # 上传指向他店
        st, r = self.a.upload(
            f"/api/import/wechat-orders?store_id={self.b_store}",
            "越权上传.xlsx", make_wx_xlsx("PXH"))
        self.assertEqual(st, 403)
        self.assertIn("无权上传到该店铺", r["message"])
        # 管理专属操作：批次撤销/删除历史、归属变更 → member 一律 403
        for path in (f"/api/import/batches/{self.b_batch_id}/revoke",
                     f"/api/import/batches/{self.b_batch_id}/delete-history"):
            st, r = self.a.post(path, {"confirm": True})
            self.assertEqual(st, 403, path)
        for path in (f"/api/stores/{self.a_store}/transfer-ownership",
                     f"/api/stores/{self.a_store}/mark-unassigned"):
            st, _ = self.a.post(path, {"target_user_id": self.b_user_id})
            self.assertEqual(st, 403, path)
        # 被攻击店铺数据零变化
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        self.assertGreater(r["data"]["total"], 0)

    # ---------- 身份伪造无效：伪造 role/user_id/member/操作人（spec 5.4.2 规则 4） ----------
    def test_03_member_forged_identity_ineffective(self):
        # 查询串伪造 role=admin + 他店 store_id → 按会话身份 → 403
        st, _ = self.a.get(
            f"/api/orders?page=1&size=20&role=admin&store_id={self.b_store}")
        self.assertEqual(st, 403)
        # body 伪造 role/user_id + 他店目标 → 403
        st, _ = self.a.post("/api/reconciliation/run", {
            "month": "2026-03", "role": "admin",
            "user_id": self.admin_user_id,
            "target_store_id": self.b_store})
        self.assertEqual(st, 403)
        st, _ = self.a.post("/api/ad-expenses", {
            "month": "2026-03", "amount": "1.00", "role": "admin",
            "user_id": self.admin_user_id,
            "target_store_id": self.b_store})
        self.assertEqual(st, 403)
        # 伪造 member 参数（含指向管理员本人）→ 403
        st, _ = self.a.get(
            f"/api/orders?page=1&size=20&member={self.admin_user_id}")
        self.assertEqual(st, 403)
        # 登录请求伪造 role 无效：会话身份仍为 member 本店
        c = Client(self.env)
        c.login("px_a", PX_PASSWORD, extra={"role": "admin"})
        st, r = c.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["role"], "member")
        self.assertEqual(r["data"]["store_id"], self.a_store)
        # 上传伪造 X-Operator 头：批次操作人仍按会话身份落账
        st, r = self.a.post_binary(
            f"/api/import/wechat-orders?store_id={self.a.store_id}",
            make_wx_xlsx("PXF"),
            headers={"X-Filename": quote("伪造操作人.xlsx"),
                     "X-Operator": self.env.admin_user})
        self.assertEqual(st, 200, r)
        st, r = self.a.post("/api/import/confirm", {
            "upload_id": r["data"]["upload_id"],
            "target_store_id": self.a.store_id})
        self.assertEqual(st, 200)
        batch = next(b for b in _store_records(
            self.env, self.a_store, "import_records")
            if b.get("file_name") == "伪造操作人.xlsx")
        self.assertEqual(batch["uploaded_by_user_id"], self.a_user_id)

    # ---------- admin 正向操作矩阵（spec 5.4.2 规则 1） ----------
    def test_04_admin_positive_operations_matrix(self):
        # 1. 查全店聚合 == 各成员店铺之和
        st, ra = self.a.get("/api/orders?page=1&size=100")
        st, rb = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        st, r = self.admin.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"],
                         ra["data"]["total"] + rb["data"]["total"])
        # 2. 双维筛选：member 维度与单店维度
        st, r = self.admin.get(
            f"/api/orders?page=1&size=100&member={self.a_user_id}")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], ra["data"]["total"])
        st, r = self.admin.get(
            f"/api/orders?page=1&size=100&store_id={self.b_store}")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], rb["data"]["total"])
        # 3. 代导入（B 店）+ 批次三要素 + 代传审计日志
        st, r = self.admin.upload_confirm(
            f"/api/import/wechat-orders?store_id={self.b_store}",
            "矩阵代传.xlsx", make_wx_xlsx("PXM"),
            target_store_id=self.b_store)
        self.assertEqual(st, 200, r)
        batch = next(b for b in _store_records(
            self.env, self.b_store, "import_records")
            if b.get("file_name") == "矩阵代传.xlsx")
        self.assertEqual(batch["uploaded_by_user_id"], self.admin_user_id)
        self.assertEqual(batch["target_store_id"], self.b_store)
        self.assertEqual(batch["target_user_id"], self.b_user_id)
        entries = _read_security_log(self.env)
        self.assertTrue([e for e in entries if e["event"] == "import.proxy"])
        # 4. 广告费代记（A 店）
        st, r = self.admin.post("/api/ad-expenses", {
            "month": "2026-03", "amount": "5.00", "note": "矩阵代记",
            "target_store_id": self.a_store})
        self.assertEqual(st, 200, r)
        st, r = self.a.get("/api/ad-expenses?month=2026-03")
        self.assertEqual(r["data"]["total"], 1)
        # 5. 重命名店铺（store_id 与数据归属不变）
        st, r = self.admin.post(f"/api/stores/{self.b_store}/rename",
                                {"name": "矩阵乙店新名"})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["store_id"], self.b_store)
        # 6. 批次撤销（代传批次）→ B 店订单恢复
        st, r = self.b.get("/api/orders?page=1&size=100")
        b_with_proxy = r["data"]["total"]
        self.assertGreater(b_with_proxy, rb["data"]["total"])
        st, r = self.admin.post(
            f"/api/import/batches/{batch['batch_id']}/revoke",
            {"confirm": True})
        self.assertEqual(st, 200, r)
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], rb["data"]["total"])
        # 7. 删除导入历史（仅历史记录; revoke 已连带删除被撤批次, 故作用于未撤销批次）
        st, r = self.admin.post(
            f"/api/import/batches/{self.b_batch_id}/delete-history",
            {"confirm": True})
        self.assertEqual(st, 200, r)
        # 8. 用户管理 + 店铺归属变更（停用/启用/重置密码/转店）
        c = self.env.create_user("px_c", PX_PASSWORD,
                                 store_name="矩阵丙占位店")
        st, _ = self.admin.post(f"/api/users/{c['user_id']}/status",
                                {"action": "disable"})
        self.assertEqual(st, 200)
        st, _ = self.admin.post(f"/api/users/{c['user_id']}/status",
                                {"action": "enable"})
        self.assertEqual(st, 200)
        st, _ = self.admin.post(f"/api/users/{c['user_id']}/reset-password",
                                {"new_password": "Px012345679"})
        self.assertEqual(st, 200)
        st, _ = self.admin.post(
            f"/api/stores/{c['store_id']}/mark-unassigned", {})
        self.assertEqual(st, 200)
        st, r = self.admin.post(
            f"/api/stores/{self.b_store}/transfer-ownership",
            {"target_user_id": c["user_id"]})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["new_owner"]["username"], "px_c")
        # 转店后 B 店数据仍完整
        px_c = self.env.login_client("px_c", "Px012345679")
        st, r = px_c.get("/api/orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], rb["data"]["total"])

    # ---------- 管理接口越权与未登录拦截（design 2.2.2 C 红线 3） ----------
    def test_05_admin_apis_member_and_anonymous_forbidden(self):
        checks = [
            ("GET", "/api/admin/all-stores"),
            ("GET", "/api/admin/migration/overview"),
            ("POST", f"/api/stores/{self.a_store}/transfer-ownership"),
            ("POST", f"/api/stores/{self.a_store}/mark-unassigned"),
        ]
        for method, path in checks:
            st, r = (self.a.get(path) if method == "GET" else
                     self.a.post(path, {"target_user_id": self.b_user_id}))
            self.assertEqual(st, 403, path)
        entries = _read_security_log(self.env)
        denied = [e for e in entries if e["event"] == "access_denied"
                  and e.get("username") == "px_a"
                  and e["detail"].get("reason") == "admin_only"]
        self.assertGreaterEqual(len(denied), len(checks))
        # 未登录：管理接口与业务接口一律 401
        anon = Client(self.env)
        for method, path in checks + [("GET", "/api/orders?page=1&size=20")]:
            st, r = (anon.get(path) if method == "GET" else
                     anon.post(path, {"target_user_id": self.b_user_id}))
            self.assertEqual(st, 401, path)


class TestFrontendAssets(unittest.TestCase):
    """spec 5.5 前端资产回归：双维筛选/归属管理/目标级联/nav/验收映射"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.member
        cls.assets = {}
        for name in ("store_scope.js", "users.html", "stores.html",
                     "index.html"):
            with open(os.path.join(STATIC_DIR, name),
                      encoding="utf-8") as f:
                cls.assets[name] = f.read()

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 10 个业务页面均接入双维筛选组件（spec 5.5.2 规则 1） ----------
    def test_01_business_pages_dual_scope_assets(self):
        for page in PAGE_FEATURES:
            st, html, _ = self.admin.get_page(page)
            self.assertEqual(st, 200, page)
            self.assertIn('id="store-scope"', html, page)
            self.assertIn("store_scope.js", html, page)

    # ---------- 组件逻辑特征：admin 双维级联 + member 固定本店 ----------
    def test_02_store_scope_dual_dimension_logic(self):
        js = self.assets["store_scope.js"]
        # admin 双维：成员+店铺两级下拉、级联收窄、全部入口
        for token in ("ss-member", "ss-store", "admin-dual", "全部成员"):
            self.assertIn(token, js)
        self.assertIn("_visibleStores", js)
        # member 固定本店：徽章展示、成员筛选强制清空、请求不携带 member
        self.assertIn("ss-badge", js)
        self.assertIn("this.memberId = null", js)
        # member 参数仅 admin 分支写入查询（member 视图无成员筛选特征）
        self.assertIn("p.member = this.memberId", js)
        # 写操作目标：admin 成员筛选下同步携带 member
        self.assertIn("t.member = this.memberId", js)

    # ---------- 店铺归属管理区块（nav_user_mgmt 拆分后落位 stores.html） ----------
    def test_03_users_page_ownership_management(self):
        html = self.assets["stores.html"]
        for token in ("openTransferOwnership", "markUnassigned",
                      "showMigrationOverview", "btn-migration-overview",
                      "/api/admin/all-stores", "assignment_status"):
            self.assertIn(token, html)
        st, _, _ = self.admin.get_page("/stores")
        self.assertEqual(st, 200)
        # 店铺管理页为 admin 专属：member 访问被拒
        st, _, _ = self.member.get_page("/stores")
        self.assertNotEqual(st, 200)
        # users.html 改造为纯用户管理页：用户操作仍在, 店铺区块已迁出
        users_html = self.assets["users.html"]
        for token in ("openResetPassword", "openClearData", "openDeleteUser",
                      "openRenameUser", "/api/users"):
            self.assertIn(token, users_html)
        self.assertNotIn("openTransferOwnership", users_html)
        self.assertNotIn("store-table", users_html)

    # ---------- index.html 店铺选择弹窗（import_store_picker 改造后形态） ----------
    def test_04_index_target_cascade(self):
        html = self.assets["index.html"]
        for token in ("openStorePicker", "uploadTask",
                      "IMPORT_API_PATHS", "选择导入店铺",
                      "请先选择目标店铺", "/static/modal.js"):
            self.assertIn(token, html)
        # 旧顶部目标级联已按需求移除
        for token in ("upload-member", "upload-store",
                      "uploadTargetReady", "请先选择目标成员和目标店铺"):
            self.assertNotIn(token, html)
        st, _, _ = self.admin.get_page("/data-import")
        self.assertEqual(st, 200)

    # ---------- nav 10 业务项 + 3 管理组项（nav_user_mgmt 拆分口径） ----------
    def test_05_nav_ten_items_on_every_page(self):
        for page, _ in NAV_ORDER:
            st, html, _ = self.admin.get_page(page)
            self.assertEqual(st, 200, page)
            nav = re.search(r"<nav>(.*?)</nav>", html, re.S)
            self.assertIsNotNone(nav, page)
            links = re.findall(
                r'<a href="([^"]+)"(?: class="active")?>([^<]+)</a>',
                nav.group(1))
            self.assertEqual(links[:len(NAV_ORDER)], NAV_ORDER, page)
            self.assertEqual(links[len(NAV_ORDER):], MGMT_ORDER, page)

    # ---------- spec 5.7.2 验收 1-12 ↔ 用例映射表（机器可校验） ----------
    def test_06_acceptance_checklist_mapping(self):
        """验收清单映射（缺项即失败）：

        验收 1  管理员不拥有业务店铺（双创建路径无店）
                → TestAdminNoStore.test_01 / test_02（对照 test_04）
        验收 2  无店管理员正常登录、管理、逐页全功能可用
                → TestAdminNoStore.test_03 / test_06
        验收 3  A→B 整店转交、store_id 与目录不变
                → TestTransferOwnership.test_05 /
                  TestMigration.test_06
        验收 4  转店后数据不变（文件字节级一致）
                → TestTransferOwnership.test_05 /
                  TestMigration.test_06 / test_10
        验收 5  A 失权（会话吊销+重登拦截）、B 得权
                → TestTransferOwnership.test_05 /
                  TestMigration.test_06
        验收 6  管理员代传归属正确（上传人=admin、数据落目标店铺、三要素）
                → TestMultiStoreAdmin.test_02 /
                  TestPermissionMatrix.test_04
        验收 7  管理员不入个人统计（admin 遗留店不入聚合/清点）
                → TestMigration.test_01 / test_07 / test_09
        验收 8  管理员可查看全部成员店铺数据（三种范围）
                → TestMemberFilter.test_01 / test_02 /
                  TestPermissionMatrix.test_04
        验收 9  member 越权一律拦截（member 参数/他店 store_id/
                他店记录 ID/伪造身份，全部 403/404 无泄露）
                → TestMemberFilter.test_04 /
                  TestPermissionMatrix.test_01 / test_02 / test_03 / test_05
        验收 10 禁止分给 admin（admin 不能自留/接收店铺）
                → TestTransferOwnership.test_04 /
                  TestAdminNoStore.test_02
        验收 11 订单列表/利润结果/导出 Excel 三口径一致
                → TestMemberFilter.test_03
        验收 12 迁移失败不丢数据（备份失败中止/核对回滚/写入回滚/禁删）
                → TestTransferOwnership.test_09 /
                  TestMigration.test_04 / test_05 / test_10
        """
        mapping = {
            1: [(TestAdminNoStore, "test_01_setup_init_admin_has_no_store"),
                (TestAdminNoStore, "test_02_create_admin_role_no_store"),
                (TestAdminNoStore, "test_04_member_contrast_with_store")],
            2: [(TestAdminNoStore,
                 "test_03_admin_no_store_all_pages_accessible"),
                (TestAdminNoStore, "test_06_restart_keeps_admin_session")],
            3: [(TestTransferOwnership, "test_05_transfer_positive_flow"),
                (TestMigration,
                 "test_06_execute_transfer_a1_full_store_handover")],
            4: [(TestTransferOwnership, "test_05_transfer_positive_flow"),
                (TestMigration,
                 "test_06_execute_transfer_a1_full_store_handover"),
                (TestMigration, "test_10_all_records_preserved_after_full_flow")],
            5: [(TestTransferOwnership, "test_05_transfer_positive_flow"),
                (TestMigration,
                 "test_06_execute_transfer_a1_full_store_handover")],
            6: [(TestMultiStoreAdmin, "test_02_admin_import_to_target_store"),
                (TestPermissionMatrix,
                 "test_04_admin_positive_operations_matrix")],
            7: [(TestMigration, "test_01_discovery_admin_legacy_and_orphan"),
                (TestMigration, "test_07_after_transfer_scope_and_partial_unbind"),
                (TestMigration,
                 "test_09_execute_a2_mark_unassigned_full_unbind")],
            8: [(TestMemberFilter, "test_01_admin_scope_orders"),
                (TestMemberFilter,
                 "test_02_admin_scope_summary_batches_logs"),
                (TestPermissionMatrix,
                 "test_04_admin_positive_operations_matrix")],
            9: [(TestMemberFilter, "test_04_member_param_forbidden_all_apis"),
                (TestPermissionMatrix,
                 "test_01_member_cross_store_reads_forbidden"),
                (TestPermissionMatrix,
                 "test_02_member_cross_store_writes_forbidden"),
                (TestPermissionMatrix,
                 "test_03_member_forged_identity_ineffective"),
                (TestPermissionMatrix,
                 "test_05_admin_apis_member_and_anonymous_forbidden")],
            10: [(TestTransferOwnership, "test_04_validation_chain"),
                 (TestAdminNoStore, "test_02_create_admin_role_no_store")],
            11: [(TestMemberFilter, "test_03_filter_caliber_consistency")],
            12: [(TestTransferOwnership, "test_09_atomic_rollback_and_retry"),
                 (TestMigration,
                  "test_04_backup_failure_aborts_with_zero_change"),
                 (TestMigration,
                  "test_05_verification_mismatch_rolls_back_both_tables"),
                 (TestMigration,
                  "test_10_all_records_preserved_after_full_flow")],
        }
        self.assertEqual(sorted(mapping), list(range(1, 13)))
        for no, cases in mapping.items():
            for klass, method in cases:
                self.assertTrue(
                    hasattr(klass, method),
                    f"验收 {no} 缺少用例 {klass.__name__}.{method}")


if __name__ == "__main__":
    unittest.main()

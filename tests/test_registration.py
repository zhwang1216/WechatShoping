"""公开自主注册验收与安全测试（spec 4.2/5.6、design 第三章）

- TestRegistrationAcceptance：15 项验收（spec 5.6.1.2），宽松限流环境
- TestRegistrationSecurity：安全补充 S1/S3-S6/S8，宽松限流环境
- TestRegistrationRateLimit：限流专项 S2/S2b（真实默认阈值 5 次/60 分钟）
- TestRegistrationConcurrency：并发同名注册 C1（恰一成功零残留）
- TestRegisterPageAccess：/register 三态访问矩阵 S7 + 未初始化拒绝 S9

五个类各自独立 TestEnv（独立临时目录 + 独立服务实例 = 独立限流计数与
安全日志），全部断言经真实 HTTP 请求驱动；测试仅运行于临时目录，
严禁读写真实 data/。A、B 用户一律经 POST /api/auth/register 公开创建。
"""
import json
import os
import threading
import unittest
from unittest import mock

from app.auth import RegistrationRateLimiter
from tests.http_client import (ADMIN_PASSWORD, ADMIN_USER, Client, TestEnv)
from tests.test_auth_isolation import make_sup_xlsx, make_wx_xlsx

REG_A = "rega"
REG_A_PASSWORD = "Rega123456"
REG_B = "regb"
REG_B_PASSWORD = "Regb123456"
A_PREFIX = "RGA"
MONTH = "2026-03"


def _dump(obj):
    return json.dumps(obj, ensure_ascii=False)


def _parse_body(raw):
    try:
        return json.loads(raw.decode("utf-8")) if raw else {}
    except (ValueError, UnicodeDecodeError):
        return {}


def _read_json(env, name):
    path = os.path.join(env.data_dir, name)
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _users(env):
    return (_read_json(env, "users.json") or {}).get("users", [])


def _stores(env):
    return (_read_json(env, "stores.json") or {}).get("stores", [])


def _find_user(env, username):
    return next((u for u in _users(env) if u.get("username") == username),
                None)


def _find_store(env, store_id):
    return next((s for s in _stores(env) if s.get("store_id") == store_id),
                None)


def _store_dirs(env):
    root = os.path.join(env.data_dir, "stores")
    return set(os.listdir(root)) if os.path.isdir(root) else set()


def register_raw(env, body):
    """匿名注册原始请求 → (status, headers, 响应字典)，供响应头契约断言"""
    client = Client(env)
    payload = json.dumps(body).encode("utf-8")
    status, headers, raw = client._request(
        "POST", "/api/auth/register", data=payload,
        headers={"Content-Type": "application/json"})
    return status, headers, _parse_body(raw)


def _list_totals(client):
    """各业务列表 total 汇总（空店铺初始状态断言用）"""
    cases = {
        "orders": "/api/orders?page=1&size=100",
        "refund_orders": "/api/refund-orders?page=1&size=100",
        "supplier_costs": "/api/supplier-costs?page=1&size=100",
        "reconciliation": "/api/reconciliation?page=1&size=100",
        "profit": "/api/profit?page=1&size=100",
        "import_batches": "/api/import/batches?page=1&size=100",
    }
    totals = {}
    for key, path in cases.items():
        st, r = client.get(path)
        assert st == 200, (key, r)
        totals[key] = r["data"]["total"]
    st, r = client.get("/api/ad-expenses")
    assert st == 200
    totals["ad_expenses"] = len(r["data"]["records"])
    return totals


# ==================== 15 项验收 ====================
class TestRegistrationAcceptance(unittest.TestCase):
    """spec 5.6.1.2 全部 15 项验收：A/B 经公开注册创建并复验隔离与业务机制"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv(register_rate_max=200)
        cls.admin = cls.env.admin
        # 验收 14 基线：注册任何用户前的管理员快照（管理员无业务店铺）
        cls.admin_before = _find_user(cls.env, ADMIN_USER)
        assert cls.admin_before["store_ids"] == [], "管理员不应拥有业务店铺"
        cls.store_dirs_before = _store_dirs(cls.env)
        # 验收 1/2：A、B 一律经公开注册创建（捕获原始响应用于响应契约断言）
        cls.a_register = register_raw(cls.env, {
            "username": REG_A, "password": REG_A_PASSWORD,
            "confirm_password": REG_A_PASSWORD, "store_name": "A注册店铺"})
        assert cls.a_register[0] == 200, cls.a_register[2]
        cls.b_register = register_raw(cls.env, {
            "username": REG_B, "password": REG_B_PASSWORD,
            "confirm_password": REG_B_PASSWORD, "store_name": "B注册店铺"})
        assert cls.b_register[0] == 200, cls.b_register[2]
        cls.a_data = cls.a_register[2]["data"]
        cls.b_data = cls.b_register[2]["data"]
        cls.a_user_id = cls.a_data["user_id"]
        cls.a_store_id = cls.a_data["store_id"]
        cls.b_user_id = cls.b_data["user_id"]
        cls.b_store_id = cls.b_data["store_id"]
        # 登录并捕获空店铺初始状态（验收 3/4 断言依据）
        cls.a = cls.env.login_client(REG_A, REG_A_PASSWORD)
        cls.b = cls.env.login_client(REG_B, REG_B_PASSWORD)
        cls.a_empty_totals = _list_totals(cls.a)
        # 验收 5/6/15 前置：A 导入业务数据并完成对账与利润
        st, _ = cls.a.upload_confirm("/api/import/wechat-orders",
                                     "A订单.xlsx", make_wx_xlsx(A_PREFIX))
        assert st == 200
        st, _ = cls.a.upload_confirm("/api/import/supplier-costs",
                                     "A成本.xlsx", make_sup_xlsx(A_PREFIX))
        assert st == 200
        st, _ = cls.a.post("/api/ad-expenses",
                           {"month": MONTH, "amount": "66.60"})
        assert st == 200
        st, _ = cls.a.post("/api/reconciliation/run", {"month": MONTH})
        assert st == 200
        st, _ = cls.a.post("/api/profit/recalculate", {"month": MONTH})
        assert st == 200

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 验收 1：A 可成功注册并创建店铺 ----------
    def test_01_register_a_success(self):
        st, headers, resp = self.a_register
        self.assertEqual(st, 200)
        self.assertEqual(resp["code"], 0)
        self.assertEqual(resp["message"], "注册成功，请前往登录")
        data = resp["data"]
        self.assertEqual(set(data),
                         {"user_id", "username", "role", "store_id",
                          "store_name"})
        self.assertEqual(data["username"], REG_A)
        self.assertEqual(data["role"], "member")
        self.assertEqual(data["store_name"], "A注册店铺")
        self.assertTrue(data["user_id"].startswith("u_"))
        self.assertTrue(data["store_id"].startswith("s_"))
        # 注册成功不建立会话（决策 D4：不下发 Set-Cookie）
        self.assertIsNone(headers.get("Set-Cookie"))
        # 管理员复核：A 在用户列表且角色为 member
        st, r = self.admin.get("/api/users")
        self.assertEqual(st, 200)
        row = next(u for u in r["data"]["users"] if u["username"] == REG_A)
        self.assertEqual(row["role"], "member")
        self.assertEqual(row["user_id"], data["user_id"])
        self.assertIn(data["store_id"], row["store_ids"])

    # ---------- 验收 2：B 可成功注册并创建店铺 ----------
    def test_02_register_b_success(self):
        st, _, resp = self.b_register
        self.assertEqual(st, 200)
        self.assertEqual(resp["code"], 0)
        data = resp["data"]
        self.assertEqual(data["username"], REG_B)
        self.assertEqual(data["role"], "member")
        self.assertTrue(data["user_id"].startswith("u_"))
        self.assertTrue(data["store_id"].startswith("s_"))
        self.assertNotEqual(data["user_id"], self.a_user_id)
        self.assertNotEqual(data["store_id"], self.a_store_id)
        # users.json 落盘记录：恒 member + active + 一对一店铺
        record = _find_user(self.env, REG_B)
        self.assertEqual(record["role"], "member")
        self.assertEqual(record["status"], "active")
        self.assertEqual(record["store_ids"], [data["store_id"]])
        store = _find_store(self.env, data["store_id"])
        self.assertEqual(store["name"], "B注册店铺")
        self.assertEqual(store["owner_user_id"], data["user_id"])

    # ---------- 验收 3：A 登录后只看到 A 的店铺 ----------
    def test_03_a_sees_own_store_only(self):
        st, r = self.a.get("/api/auth/me")
        self.assertEqual(st, 200)
        me = r["data"]
        self.assertEqual(me["user_id"], self.a_user_id)
        self.assertEqual(me["store_id"], self.a_store_id)
        self.assertEqual(me["role"], "member")
        self.assertEqual(me["store_name"], "A注册店铺")
        # 空店铺初始状态：导入前捕获的全部列表均为空
        for key, total in self.a_empty_totals.items():
            self.assertEqual(total, 0, key)

    # ---------- 验收 4：B 登录后只看到 B 的店铺 ----------
    def test_04_b_sees_own_store_only(self):
        st, r = self.b.get("/api/auth/me")
        self.assertEqual(st, 200)
        me = r["data"]
        self.assertEqual(me["user_id"], self.b_user_id)
        self.assertEqual(me["store_id"], self.b_store_id)
        self.assertNotEqual(me["store_id"], self.a_store_id)
        # B 的任何列表中无 A 的数据痕迹
        cases = [
            "/api/orders?page=1&size=100",
            "/api/refund-orders?page=1&size=100",
            "/api/supplier-costs?page=1&size=100",
            "/api/reconciliation?page=1&size=100",
            "/api/profit?page=1&size=100",
            "/api/import/batches?page=1&size=100",
            "/api/summary?month=" + MONTH,
            "/api/ad-expenses",
        ]
        for path in cases:
            with self.subTest(path=path):
                st, r = self.b.get(path)
                self.assertEqual(st, 200, path)
                self.assertNotIn(A_PREFIX, _dump(r), path)
                self.assertNotIn(self.a_store_id, _dump(r), path)
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], 0)

    # ---------- 验收 5：A 导入的订单不出现在 B 列表 ----------
    def test_05_import_isolation(self):
        st, r = self.a.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        self.assertGreaterEqual(r["data"]["total"], 1)
        self.assertTrue(all(o["order_id"].startswith(A_PREFIX)
                            for o in r["data"]["orders"]))
        # B 同接口 total=0 且搜索查不到 A 的订单号
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], 0)
        st, r = self.b.get(f"/api/orders?keyword={A_PREFIX}0001&page=1&size=100")
        self.assertEqual(r["data"]["total"], 0)
        # A 的导入批次对 B 不可见：列表不含、详情 404
        st, r = self.a.get("/api/import/batches?page=1&size=100")
        a_batches = {x["batch_id"] for x in r["data"]["batches"]}
        self.assertEqual(len(a_batches), 2)
        st, r = self.b.get("/api/import/batches?page=1&size=100")
        self.assertEqual(r["data"]["total"], 0)
        for batch_id in a_batches:
            st, r = self.b.get(f"/api/import/batches/{batch_id}")
            self.assertEqual(st, 404)

    # ---------- 验收 6：A 的退款/成本/广告费/利润不出现在 B 页面 ----------
    def test_06_financial_data_isolation(self):
        # A 侧各财务数据均存在
        st, r = self.a.get("/api/refund-orders?page=1&size=100")
        self.assertGreaterEqual(r["data"]["total"], 1)
        st, r = self.a.get("/api/supplier-costs?page=1&size=100")
        self.assertGreaterEqual(r["data"]["total"], 1)
        st, r = self.a.get("/api/ad-expenses")
        self.assertTrue(any(x.get("month") == MONTH
                            and x.get("amount") == "66.60"
                            for x in r["data"]["records"]))
        st, r = self.a.get("/api/reconciliation?page=1&size=100")
        self.assertGreaterEqual(r["data"]["total"], 1)
        st, r = self.a.get("/api/profit?page=1&size=100")
        self.assertGreaterEqual(r["data"]["total"], 1)
        # B 侧对应接口全部为空/金额为 0
        st, r = self.b.get("/api/refund-orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], 0)
        st, r = self.b.get("/api/supplier-costs?page=1&size=100")
        self.assertEqual(r["data"]["total"], 0)
        st, r = self.b.get("/api/ad-expenses")
        self.assertEqual(r["data"]["records"], [])
        st, r = self.b.get("/api/reconciliation?page=1&size=100")
        self.assertEqual(r["data"]["total"], 0)
        st, r = self.b.get("/api/profit?page=1&size=100")
        self.assertEqual(r["data"]["total"], 0)
        st, r = self.b.get("/api/summary?month=" + MONTH)
        self.assertFalse(r["data"]["has_orders"])

    # ---------- 验收 7：伪造 ID 查不到 A 数据 ----------
    def test_07_forged_ids_denied(self):
        # B 携带 A 的 store_id 查询参数 → 403 拒绝（多店铺隔离强化）
        st, r = self.b.get(
            f"/api/orders?page=1&size=100&store_id={self.a_store_id}"
            f"&user_id={self.a_user_id}")
        self.assertEqual(st, 403)
        # B 正常查询仅返回自己的数据（空）
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 0)
        self.assertNotIn(A_PREFIX, _dump(r))
        # B 用 A 的订单号查详情 → 404"不存在"
        st, r = self.b.get(f"/api/orders/{A_PREFIX}0001")
        self.assertEqual(st, 404)
        self.assertIn("不存在", r["message"])
        # B 调 A 的 user_id 用户管理接口 → 403（member 无任何用户管理权）
        st, r = self.b.post(f"/api/users/{self.a_user_id}/status",
                            {"action": "disable"})
        self.assertEqual(st, 403)
        st, r = self.b.get("/api/users")
        self.assertEqual(st, 403)

    # ---------- 验收 8：注册接口不能创建管理员账号 ----------
    def test_08_no_admin_via_register(self):
        st, r = self.env.register_user(
            "regc", "Regc123456", store_name="C注入店铺",
            extra={"role": "admin", "is_admin": True})
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["role"], "member")
        self.assertEqual(_find_user(self.env, "regc")["role"], "member")
        # 该账号登录后不具备任何管理员能力
        c = self.env.login_client("regc", "Regc123456")
        st, _, headers = c.get_page("/users")
        self.assertEqual(st, 302)
        self.assertEqual(headers.get("Location"), "/")
        st, r = c.post("/api/users",
                       {"username": "regx", "password": "Regx123456"})
        self.assertEqual(st, 403)
        st, r = c.get("/api/users")
        self.assertEqual(st, 403)

    # ---------- 验收 9：重复用户名被服务端拒绝 ----------
    def test_09_duplicate_username_rejected(self):
        users_before = len(_users(self.env))
        stores_before = len(_stores(self.env))
        st, r = self.env.register_user(REG_A, REG_A_PASSWORD,
                                       store_name="重复店铺")
        self.assertEqual(st, 400)
        self.assertIn("用户名已存在", r["message"])
        self.assertEqual(len(_users(self.env)), users_before)
        self.assertEqual(len(_stores(self.env)), stores_before)

    # ---------- 验收 10：密码不匹配时不能注册 ----------
    def test_10_password_mismatch_rejected(self):
        st, r = self.env.register_user(
            "regm", "Regm123456", store_name="M店铺",
            confirm_password="Regm654321")
        self.assertEqual(st, 400)
        self.assertIn("两次输入的密码不一致", r["message"])
        self.assertIsNone(_find_user(self.env, "regm"))
        self.assertFalse(any(s.get("name") == "M店铺"
                             for s in _stores(self.env)))

    # ---------- 验收 11：重复提交不创建重复店铺 ----------
    def test_11_duplicate_submission_no_dup_store(self):
        users_before = len(_users(self.env))
        dirs_before = _store_dirs(self.env)
        st1, r1 = self.env.register_user("regd", "Regd123456",
                                         store_name="D店铺")
        st2, r2 = self.env.register_user("regd", "Regd123456",
                                         store_name="D店铺")
        self.assertEqual(st1, 200)
        self.assertEqual(st2, 400)
        self.assertIn("用户名已存在", r2["message"])
        record = _find_user(self.env, "regd")
        self.assertIsNotNone(record)
        # 恰 1 用户记录、1 店铺记录、1 个店铺目录
        self.assertEqual(len(_users(self.env)), users_before + 1)
        owned = [s for s in _stores(self.env)
                 if s.get("owner_user_id") == record["user_id"]]
        self.assertEqual(len(owned), 1)
        self.assertEqual(owned[0]["store_id"], record["store_ids"][0])
        self.assertEqual(_store_dirs(self.env) - dirs_before,
                         {record["store_ids"][0]})

    # ---------- 验收 12：未登录访问业务 API 被拒绝 ----------
    def test_12_unauthenticated_api_denied(self):
        anon = Client(self.env)
        st, r = anon.get("/api/orders?page=1&size=50")
        self.assertEqual(st, 401)
        self.assertEqual(r["code"], 1001)
        st, r = anon.get("/api/users")
        self.assertEqual(st, 401)
        st, r = anon.post("/api/import/confirm", {"upload_id": "nonexistent"})
        self.assertEqual(st, 401)

    # ---------- 验收 13：登录/刷新/退出/重登无重定向循环 ----------
    def test_13_login_refresh_logout_no_loop(self):
        c = self.env.login_client(REG_A, REG_A_PASSWORD)
        # 登录后访问首页与业务页均 200（刷新语义，无 302 循环）
        for page in ("/", "/dashboard", "/orders", "/profit"):
            st, _, headers = c.get_page(page)
            self.assertEqual(st, 200, page)
            self.assertIsNone(headers.get("Location"), page)
        old_cookie = c.cookie
        st, _ = c.logout()
        self.assertEqual(st, 200)
        # 旧 Cookie 调业务 API → 401（会话已注销）
        c.cookie = old_cookie
        st, r = c.get("/api/orders?page=1&size=50")
        self.assertEqual(st, 401)
        # 登出后登录页可直接访问（不回跳成环）
        st, body, _ = c.get_page("/login")
        self.assertEqual(st, 200)
        self.assertIn("登录", body)
        # 重新登录成功
        c.cookie = None
        c.login(REG_A, REG_A_PASSWORD)
        st, r = c.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["username"], REG_A)

    # ---------- 验收 14：原有数据在升级后仍然完整（独立环境） ----------
    def test_14_existing_data_intact(self):
        env = TestEnv(register_rate_max=200)
        try:
            admin = env.admin
            # 管理员无店；承载成员店铺作为既有业务数据基线
            admin_before = _find_user(env, ADMIN_USER)
            assert admin_before["store_ids"] == [], "管理员不应拥有业务店铺"
            carrier = _find_user(env, "biz_carrier")
            carrier_store_before = _find_store(env, carrier["store_ids"][0])
            dirs_before = _store_dirs(env)
            totals_before = {
                "orders": admin.get("/api/orders?page=1&size=20")[1]["data"]["total"],
                "refund_orders": admin.get(
                    "/api/refund-orders?page=1&size=20")[1]["data"]["total"],
                "supplier_costs": admin.get(
                    "/api/supplier-costs?page=1&size=20")[1]["data"]["total"],
            }
            env.register_user("rega", "Rega123456", store_name="A注册店铺")
            env.register_user("regb", "Regb123456", store_name="B注册店铺")
            # 管理员与承载成员的用户/店铺记录逐字段不变
            self.assertEqual(_find_user(env, ADMIN_USER), admin_before)
            self.assertEqual(_find_user(env, "biz_carrier"), carrier)
            self.assertEqual(_find_store(env, carrier["store_ids"][0]),
                             carrier_store_before)
            # 聚合业务数据 total 不变（新注册店铺均为空店）
            self.assertEqual(
                admin.get("/api/orders?page=1&size=20")[1]["data"]["total"],
                totals_before["orders"])
            self.assertEqual(
                admin.get("/api/refund-orders?page=1&size=20")[1]["data"]["total"],
                totals_before["refund_orders"])
            self.assertEqual(
                admin.get("/api/supplier-costs?page=1&size=20")[1]["data"]["total"],
                totals_before["supplier_costs"])
            # stores/ 下仅新增 A、B 两个新目录
            new_dirs = _store_dirs(env) - dirs_before
            self.assertEqual(len(new_dirs), 2)
            registered = [u for u in _users(env)
                          if u["username"] in ("rega", "regb")]
            self.assertEqual({u["store_ids"][0] for u in registered},
                             new_dirs)
        finally:
            env.cleanup()

    # ---------- 验收 15：月份筛选/财务计算/权限控制仍正常 ----------
    def test_15_biz_rules_still_work(self):
        # 月份筛选：A 的数据全部在 2026-03
        st, r = self.a.get(f"/api/orders?month={MONTH}&page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 10)
        st, r = self.a.get("/api/orders?month=2026-01&page=1&size=100")
        self.assertEqual(r["data"]["total"], 0)
        st, r = self.a.get(f"/api/summary?month={MONTH}")
        self.assertEqual(r["data"]["order_count"], 12)
        # 对账与利润核算可正常执行
        st, r = self.a.post("/api/reconciliation/run", {"month": MONTH})
        self.assertEqual(st, 200)
        st, r = self.a.get(f"/api/reconciliation?month={MONTH}&page=1&size=100")
        self.assertEqual(st, 200)
        self.assertGreaterEqual(r["data"]["total"], 1)
        st, r = self.a.post("/api/profit/recalculate", {"month": MONTH})
        self.assertEqual(st, 200)
        st, r = self.a.get(f"/api/profit?month={MONTH}&page=1&size=100")
        self.assertEqual(st, 200)
        self.assertGreaterEqual(r["data"]["total"], 1)
        # 广告费可编辑
        st, r = self.a.post("/api/ad-expenses",
                            {"month": MONTH, "amount": "88.80"})
        self.assertEqual(st, 200)
        st, r = self.a.get("/api/ad-expenses")
        self.assertTrue(any(x.get("month") == MONTH
                            and x.get("amount") == "88.80"
                            for x in r["data"]["records"]))
        # 权限控制仍正常：member 不能访问用户管理
        st, _, headers = self.a.get_page("/users")
        self.assertEqual(st, 302)
        self.assertEqual(headers.get("Location"), "/")
        st, r = self.a.post("/api/users",
                            {"username": "regy", "password": "Regy123456"})
        self.assertEqual(st, 403)


# ==================== 安全补充场景 ====================
class TestRegistrationSecurity(unittest.TestCase):
    """spec 5.6.1.3 安全补充 S1/S3-S6/S8（限流专项 S2 归独立类）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv(register_rate_max=200)
        cls.admin = cls.env.admin

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _security_entries(self):
        path = os.path.join(self.env.log_dir, "security_log.jsonl")
        entries = []
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    entries.append(json.loads(line))
        return entries

    # ---------- S1：权限字段注入无效 ----------
    def test_s01_privileged_fields_ignored(self):
        forged = {"role": "admin", "is_admin": True, "user_id": "u_evil",
                  "store_id": "s_evil", "store_ids": ["s_evil2"],
                  "password_version": 99, "status": "disabled"}
        st, r = self.env.register_user(
            "seca", "Seca123456", store_name="安全A店铺", extra=forged)
        self.assertEqual(st, 200)
        data = r["data"]
        self.assertEqual(data["role"], "member")
        self.assertTrue(data["user_id"].startswith("u_"))
        self.assertTrue(data["store_id"].startswith("s_"))
        self.assertNotEqual(data["user_id"], "u_evil")
        self.assertNotEqual(data["store_id"], "s_evil")
        # users.json 落盘记录无任何注入痕迹
        record = _find_user(self.env, "seca")
        self.assertEqual(record["user_id"], data["user_id"])
        self.assertEqual(record["role"], "member")
        self.assertEqual(record["status"], "active")
        self.assertEqual(record["password_version"], 1)
        self.assertEqual(record["store_ids"], [data["store_id"]])
        self.assertNotIn("is_admin", record)
        store = _find_store(self.env, data["store_id"])
        self.assertEqual(store["owner_user_id"], data["user_id"])
        self.assertNotEqual(store["store_id"], "s_evil")

    # ---------- S3：CSRF 头缺失 403 ----------
    def test_s03_csrf_required(self):
        users_before = len(_users(self.env))
        stores_before = len(_stores(self.env))
        client = Client(self.env)
        st, r = client.post("/api/auth/register", {
            "username": "secc", "password": "Secc123456",
            "confirm_password": "Secc123456", "store_name": "跨站店铺",
        }, csrf=False)
        self.assertEqual(st, 403)
        self.assertEqual(r["code"], 1002)
        self.assertIn("防跨站校验头", r["message"])
        self.assertEqual(len(_users(self.env)), users_before)
        self.assertEqual(len(_stores(self.env)), stores_before)

    # ---------- S4：字段校验逐项拒绝 ----------
    def test_s04_field_validation(self):
        client = Client(self.env)
        valid_password = "Field123"
        missing_cases = [
            ("缺失用户名", {"username": None}),
            ("缺失密码", {"password": None}),
            ("缺失确认密码", {"confirm_password": None}),
            ("缺失店铺名", {"store_name": None}),
        ]
        for name, patch in missing_cases:
            body = {"username": "fieldmiss", "password": valid_password,
                    "confirm_password": valid_password, "store_name": "校验店铺"}
            body.update(patch)
            with self.subTest(case=name):
                st, r = client.post("/api/auth/register", body)
                self.assertEqual(st, 400, name)
                self.assertIsNotNone(r.get("message"))
        message_cases = [
            ("超长用户名", "a" * 33,
             "用户名长度需在 3-32 个字符之间"),
            ("非法字符用户名", "bad@user#1",
             "用户名只能包含字母、数字、下划线、连字符和点"),
            ("纯数字密码", "12345678",
             "密码必须同时包含字母和数字"),
            ("纯字母密码", "abcdefgh",
             "密码必须同时包含字母和数字"),
            ("7位短密码", "Abc1234",
             "密码至少 8 位"),
            ("空白店铺名", "   ",
             "请输入店铺名称"),
            ("超长店铺名", "铺" * 65,
             "店铺名称不能超过 64 个字符"),
        ]
        for name, value, message in message_cases:
            if name in ("超长用户名", "非法字符用户名"):
                body = {"username": value, "password": valid_password,
                        "confirm_password": valid_password,
                        "store_name": "校验店铺"}
            elif name in ("纯数字密码", "纯字母密码", "7位短密码"):
                body = {"username": "fielduser", "password": value,
                        "confirm_password": value, "store_name": "校验店铺"}
            else:
                body = {"username": "fielduser", "password": valid_password,
                        "confirm_password": valid_password,
                        "store_name": value}
            with self.subTest(case=name):
                st, r = client.post("/api/auth/register", body)
                self.assertEqual(st, 400, name)
                self.assertIn(message, r["message"], name)
        # 全部被拒：无任何记录创建
        self.assertFalse(any(u["username"].startswith("field")
                             for u in _users(self.env)))
        self.assertFalse(any(s.get("name") in ("校验店铺",)
                             for s in _stores(self.env)))

    # ---------- S5：注册失败回滚零残留 ----------
    def test_s05_rollback_no_residue(self):
        users_before = len(_users(self.env))
        stores_before = len(_stores(self.env))
        dirs_before = _store_dirs(self.env)
        with mock.patch("app.server.StoreServiceRegistry.get",
                        side_effect=OSError("模拟店铺空间初始化失败")):
            st, r = self.env.register_user("secb", "Secb123456",
                                           store_name="回滚店铺")
        self.assertEqual(st, 500)
        self.assertIn("注册失败，请稍后重试", r["message"])
        # 回滚零残留：用户/店铺记录与店铺目录均未增加
        self.assertEqual(len(_users(self.env)), users_before)
        self.assertEqual(len(_stores(self.env)), stores_before)
        self.assertIsNone(_find_user(self.env, "secb"))
        self.assertFalse(any(s.get("name") == "回滚店铺"
                             for s in _stores(self.env)))
        self.assertEqual(_store_dirs(self.env), dirs_before)
        # 取消 patch 后同一用户名可重试成功（可重试性）
        st, r = self.env.register_user("secb", "Secb123456",
                                       store_name="回滚店铺")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["username"], "secb")

    # ---------- S6：无 Set-Cookie（注册不自动登录） ----------
    def test_s06_no_set_cookie(self):
        st, headers, resp = register_raw(self.env, {
            "username": "secd", "password": "Secd123456",
            "confirm_password": "Secd123456", "store_name": "无Cookie店铺"})
        self.assertEqual(st, 200)
        self.assertIsNone(headers.get("Set-Cookie"))
        # 携带注册前任意旧 Cookie 调业务 API → 与未登录一致 401
        client = Client(self.env)
        client.cookie = "forged-session-token-xyz"
        st, r = client.get("/api/orders?page=1&size=50")
        self.assertEqual(st, 401)
        self.assertEqual(r["code"], 1001)
        # 随后正常登录：Set-Cookie 属性与既有一致
        payload = json.dumps({"username": "secd",
                              "password": "Secd123456"}).encode("utf-8")
        st, headers, raw = client._request(
            "POST", "/api/auth/login", data=payload, csrf=False,
            headers={"Content-Type": "application/json"})
        self.assertEqual(st, 200)
        cookie_header = headers.get("Set-Cookie") or ""
        for attr in ("wxr_session=", "Path=/", "HttpOnly",
                     "SameSite=Lax", "Max-Age=43200"):
            self.assertIn(attr, cookie_header)

    # ---------- S8：安全日志事件齐备且无秘密 ----------
    def test_s08_security_log_events(self):
        st, _ = self.env.register_user("sece", "Sece123456",
                                       store_name="日志店铺")
        self.assertEqual(st, 200)
        st, _ = self.env.register_user("sece", "Sece123456",
                                       store_name="日志店铺")
        self.assertEqual(st, 400)
        st, _ = self.env.register_user("x", "Sece123456", store_name="日志店铺")
        self.assertEqual(st, 400)
        entries = self._security_entries()
        by_event = {}
        for e in entries:
            by_event.setdefault(e["event"], []).append(e)
        # register.success：含 username/ip/store_id
        self.assertIn("register.success", by_event)
        success = by_event["register.success"][-1]
        self.assertEqual(success["username"], "sece")
        self.assertIsNotNone(success["ip"])
        self.assertTrue(success["detail"]["store_id"].startswith("s_"))
        # register.failed：含失败原因
        self.assertIn("register.failed", by_event)
        reasons = {e["detail"]["reason"] for e in by_event["register.failed"]}
        self.assertTrue(any("用户名已存在" in x for x in reasons))
        self.assertTrue(any("用户名长度需在" in x for x in reasons))
        # 全文不含任何秘密：注册密码/哈希/令牌
        text = _dump(entries)
        for pwd in ("Seca123456", "Secb123456", "Secd123456", "Sece123456"):
            self.assertNotIn(pwd, text)
        self.assertNotIn("pbkdf2", text)
        self.assertNotIn("salt", text.lower())
        self.assertNotIn("wxr_session", text)
        # 既有事件类型字段不变：每条均有 time/event/result
        for e in entries:
            self.assertIn("time", e)
            self.assertIn("event", e)
            self.assertIn("result", e)


# ==================== 限流专项 ====================
class TestRegistrationRateLimit(unittest.TestCase):
    """S2/S2b：真实默认阈值（5 次/60 分钟）；额度一次性消费，置于独立类"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()  # 不传参 → 生产默认阈值
        cls.admin = cls.env.admin

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_s02_rate_limit_429_and_independent(self):
        # 限流触发前：正常账密登录成功
        c = Client(self.env)
        c.login(ADMIN_USER, ADMIN_PASSWORD)
        # 1 次成功 + 4 次字段错误 = 5 次（成败均计）
        st, r = self.env.register_user("rluser", "Rluser123456",
                                       store_name="限流店铺")
        self.assertEqual(st, 200)
        for _ in range(4):
            st, r = self.env.register_user("x", "Rluser123456",
                                           store_name="限流店铺")
            self.assertEqual(st, 400)
        users_before = len(_users(self.env))
        dirs_before = _store_dirs(self.env)
        # 第 6 次 → 429/1004，且无记录创建
        st, r = self.env.register_user("rlblocked", "Rluser123456",
                                       store_name="限流店铺")
        self.assertEqual(st, 429)
        self.assertEqual(r["code"], 1004)
        self.assertIn("注册过于频繁", r["message"])
        self.assertEqual(len(_users(self.env)), users_before)
        self.assertEqual(_store_dirs(self.env), dirs_before)
        # 注册限流不影响登录：正确账密登录成功
        c2 = Client(self.env)
        me = c2.login(ADMIN_USER, ADMIN_PASSWORD)
        self.assertEqual(me["username"], ADMIN_USER)
        # 错误密码登录仍走登录限流语义（401 而非 429，计数互不共享）
        for _ in range(5):
            st, r = Client(self.env).post(
                "/api/auth/login",
                {"username": "rluser", "password": "WrongPass1"},
                csrf=False)
            self.assertEqual(st, 401)
            self.assertEqual(r["message"], "账号或密码错误")
        # 正确登录仍成功（登录失败计数远未达 10 次阈值）
        c3 = Client(self.env)
        c3.login("rluser", "Rluser123456")
        st, r = c3.get("/api/auth/me")
        self.assertEqual(st, 200)
        # 安全日志含 register.rate_limited 事件
        path = os.path.join(self.env.log_dir, "security_log.jsonl")
        with open(path, encoding="utf-8") as f:
            self.assertIn("register.rate_limited", f.read())

    def test_s02b_limiter_multi_ip(self):
        limiter = RegistrationRateLimiter(window_seconds=60, max_requests=5)
        for _ in range(5):
            limiter.record_request("1.1.1.1")
        self.assertTrue(limiter.is_blocked("1.1.1.1"))
        self.assertFalse(limiter.is_blocked("2.2.2.2"))
        limiter.record_request("2.2.2.2")
        self.assertFalse(limiter.is_blocked("2.2.2.2"))


# ==================== 并发注册 ====================
class TestRegistrationConcurrency(unittest.TestCase):
    """C1：两线程同时提交同一用户名 → 恰一成功、零残留（spec 4.2.2）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv(register_rate_max=200)
        cls.admin = cls.env.admin

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_c01_concurrent_same_username(self):
        body = {"username": "conc1", "password": "Conc112345",
                "confirm_password": "Conc112345", "store_name": "并发店铺"}
        results = []
        errors = []
        barrier = threading.Barrier(2)

        def worker():
            try:
                client = Client(self.env)
                barrier.wait(timeout=10)
                results.append(client.post("/api/auth/register", body))
            except Exception as e:
                errors.append(e)

        threads = [threading.Thread(target=worker) for _ in range(2)]
        for t in threads:
            t.start()
        for t in threads:
            t.join(timeout=30)
        self.assertEqual(errors, [])
        # 恰一个 200、一个 400"用户名已存在"
        statuses = sorted(r[0] for r in results)
        self.assertEqual(statuses, [200, 400])
        for st, r in results:
            if st == 400:
                self.assertIn("用户名已存在", r["message"])
        # users/stores 恰 1 用户 1 店铺、店铺目录恰 1 个
        record = _find_user(self.env, "conc1")
        self.assertIsNotNone(record)
        self.assertEqual(len([u for u in _users(self.env)
                              if u["username"] == "conc1"]), 1)
        owned = [s for s in _stores(self.env)
                 if s.get("owner_user_id") == record["user_id"]]
        self.assertEqual(len(owned), 1)
        store_dir = os.path.join(self.env.data_dir, "stores",
                                 record["store_ids"][0])
        self.assertTrue(os.path.isdir(store_dir))
        # users 恰 3 个（admin + 承载成员 biz_carrier + conc1）
        self.assertEqual(len(_users(self.env)), 3)
        self.assertEqual(len(_stores(self.env)), 2)
        self.assertEqual(len(_store_dirs(self.env)), 2)


# ==================== 页面访问与未初始化 ====================
class TestRegisterPageAccess(unittest.TestCase):
    """S7：/register 三态访问矩阵 + S9：未初始化注册拒绝"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv(initialize=False)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_s07_register_page_access_matrix(self):
        anon = Client(self.env)
        # 状态 1：未初始化 → /register 与静态直访均引导 /setup
        for path in ("/register", "/static/register.html"):
            st, _, headers = anon.get_page(path)
            self.assertEqual(st, 302, path)
            self.assertEqual(headers.get("Location"), "/setup", path)
        # 完成初始化
        st, _ = self.env.setup_admin()
        self.assertEqual(st, 200)
        # 状态 2：已初始化 + 匿名 → 200 且含注册表单四字段与返回登录入口
        st, body, _ = anon.get_page("/register")
        self.assertEqual(st, 200)
        for field in ('id="username"', 'id="password"', 'id="confirm"',
                      'id="store-name"'):
            self.assertIn(field, body)
        self.assertIn("返回登录", body)
        st, body2, _ = anon.get_page("/static/register.html")
        self.assertEqual(st, 200)
        self.assertEqual(body, body2)
        # 登录页含注册入口（任务 3.2 断言锚点）
        st, body, _ = anon.get_page("/login")
        self.assertEqual(st, 200)
        self.assertIn('href="/register"', body)
        # 状态 3：已登录 → /register 与静态直访均 302 回首页
        admin = self.env.admin
        for path in ("/register", "/static/register.html"):
            st, _, headers = admin.get_page(path)
            self.assertEqual(st, 302, path)
            self.assertEqual(headers.get("Location"), "/", path)
        # 重定向后首页正常放行（无循环）
        st, _, _ = admin.get_page("/")
        self.assertEqual(st, 200)

    def test_s09_uninitialized_register_rejected(self):
        env = TestEnv(initialize=False)
        try:
            st, r = env.register_user("earlybird", "Early123456",
                                      store_name="抢先店铺")
            self.assertEqual(st, 400)
            self.assertIn("系统尚未初始化", r["message"])
            # users.json 无任何用户（is_initialized 仍为 False）
            self.assertEqual(_users(env), [])
            # 随后 /setup 初始化仍可成功（决策 D8：注册不能锁死系统）
            st, _ = env.setup_admin()
            self.assertEqual(st, 200)
        finally:
            env.cleanup()


if __name__ == "__main__":
    unittest.main()

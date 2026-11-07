"""store_ownership 特性验收测试：一对多店铺归属 + 管理员删除店铺权限

测试类与任务组对应关系（tasks.md）:
- TestMemberCreateOptionalStore：创建 member 初始店铺可选（组 1）
- TestMemberMultiStoreScope：member 多店会话与查询范围三态（组 2）
- TestCreateStore / TestStoreStatus / TestStoreDetail /
  TestAdminAllStoresExtension / TestOneToManyOwnership（组 3）
- TestStoreDelete / TestStoreClearData：删除分级与清除数据（组 4）
- TestMigration：店铺数据模型迁移（组 5）
- TestFrontendAssets：前端资产静态断言（组 6）
- TestAuditAndSecurity / TestAcceptance12：审计安全与验收汇总（组 7）

全部用例经 TestEnv 隔离临时目录运行，严禁读写真实 data/。
"""
import json
import os
import re
import shutil
import time
import unittest
from decimal import Decimal, InvalidOperation
from unittest import mock
from urllib.parse import quote

from openpyxl import load_workbook

from app.auth import AuthError
from app.data_schemas import BUSINESS_FILES
from tests.http_client import Client, TestEnv
from tests.test_auth_isolation import make_sup_xlsx, make_wx_xlsx
from tests.test_delivery_e2e import WX_ROWS

PASSWORD = "So123456789"
MONTH = "2026-03"
# 一次订单 Excel 导入的行级记录数（orders.json 记录 ≠ API 订单级聚合）
WX_ROW_COUNT = len(WX_ROWS)


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


def _read_security_log(env):
    path = os.path.join(env.log_dir, "security_log.jsonl")
    entries = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                entries.append(json.loads(line))
    return entries


def _read_business_records(env, store_id, name):
    """读取店铺业务文件 records（与 _store_inventory 同口径数据源）"""
    spec = BUSINESS_FILES[name]
    path = os.path.join(env.data_dir, "stores", store_id, "data",
                        spec["file"])
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return json.load(f).get("records", [])


def _sum_records(records, field):
    total = Decimal("0")
    for r in records:
        try:
            total += Decimal(str(r.get(field) or "0"))
        except (InvalidOperation, ArithmeticError, TypeError, ValueError):
            continue
    return str(total.quantize(Decimal("0.01")))


class TestMemberCreateOptionalStore(unittest.TestCase):
    """创建 member 初始店铺可选（spec 5.3.1 规则 1/8）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_default_creates_initial_store(self):
        st, r = self.admin.post("/api/users", {
            "username": "so_default", "password": PASSWORD})
        self.assertEqual(st, 200, r)
        self.assertIsNotNone(r["data"]["store_id"])
        user = _find_user(self.env, "so_default")
        self.assertEqual(len(user["store_ids"]), 1)
        store = _find_store(self.env, r["data"]["store_id"])
        self.assertEqual(store["owner_user_id"], user["user_id"])
        # 六类必备字段完整（store_id/name/owner/status/created_at/updated_at）
        self.assertEqual(store.get("status"), "active")
        self.assertEqual(store.get("assignment_status"), "assigned")
        self.assertIsNotNone(store.get("created_at"))
        self.assertIsNotNone(store.get("updated_at"))

    def test_02_explicit_true_creates_store(self):
        st, r = self.admin.post("/api/users", {
            "username": "so_true", "password": PASSWORD,
            "create_store": True, "store_name": "显式建店"})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["store_name"], "显式建店")
        store = _find_store(self.env, r["data"]["store_id"])
        self.assertEqual(store.get("status"), "active")
        self.assertEqual(store.get("assignment_status"), "assigned")

    def test_03_false_skips_store(self):
        st, r = self.admin.post("/api/users", {
            "username": "so_nostore", "password": PASSWORD,
            "create_store": False})
        self.assertEqual(st, 200, r)
        self.assertIsNone(r["data"]["store_id"])
        self.assertIsNone(r["data"]["store_name"])
        user = _find_user(self.env, "so_nostore")
        self.assertEqual(user["store_ids"], [])
        self.assertIsNone(
            next((s for s in _read_stores(self.env)
                  if s.get("owner_user_id") == user["user_id"]), None))

    def test_04_member_without_store_login_denied(self):
        self.admin.post("/api/users", {
            "username": "so_locked", "password": PASSWORD,
            "create_store": False})
        st, r = Client(self.env).post("/api/auth/login", {
            "username": "so_locked", "password": PASSWORD}, csrf=False)
        self.assertEqual(st, 401)
        self.assertEqual(r["message"], "账号未绑定店铺，请联系管理员")

    def test_05_admin_role_never_creates_store(self):
        stores_before = len(_read_stores(self.env))
        st, r = self.admin.post("/api/users", {
            "username": "so_admin", "password": PASSWORD,
            "role": "admin", "create_store": True})
        self.assertEqual(st, 200, r)
        self.assertIsNone(r["data"]["store_id"])
        self.assertIsNone(r["data"]["store_name"])
        self.assertEqual(_find_user(self.env, "so_admin")["store_ids"], [])
        self.assertEqual(len(_read_stores(self.env)), stores_before)

    def test_06_register_still_creates_store(self):
        st, r = self.env.register_user("so_reg", PASSWORD,
                                       store_name="注册店铺")
        self.assertEqual(st, 200, r)
        user = _find_user(self.env, "so_reg")
        self.assertEqual(len(user["store_ids"]), 1)
        store = _find_store(self.env, user["store_ids"][0])
        self.assertEqual(store.get("status"), "active")
        self.assertEqual(store.get("assignment_status"), "assigned")
        self.assertIsNotNone(store.get("updated_at"))

    def test_07_invalid_create_store_type_rejected(self):
        st, r = self.admin.post("/api/users", {
            "username": "so_bad", "password": PASSWORD,
            "create_store": "yes"})
        self.assertEqual(st, 400)
        self.assertIn("create_store", r["message"])


class TestMemberMultiStoreScope(unittest.TestCase):
    """member 多店会话与查询范围三态（spec 5.7 / design 2.1.3(2)）

    构造: ms_a 拥有 S1（初始店）+ S2（转入店）各含不同订单数据;
    验证聚合态/单店态/拒绝态三态解析、写入与上传目标扩展、越权口径。
    """

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            "ms_a", PASSWORD, store_name="多店A一店")
        cls.helper = cls.env.create_user_client(
            "ms_b", PASSWORD, store_name="待转入店")
        cls.s1 = cls.member.get("/api/auth/me")[1]["data"]["store_id"]
        cls.s2 = cls.helper.get("/api/auth/me")[1]["data"]["store_id"]
        # S2 转给 ms_a → 一对多追加绑定（helper 会话被吊销, 成员会话保持）
        st, r = cls.admin.post(
            f"/api/stores/{cls.s2}/transfer-ownership",
            {"target_user_id": cls.env.get_user_id("ms_a")})
        assert st == 200, r
        # S1 灌入 MS1 前缀订单（member 上传默认落会话店铺）
        st, r = cls.member.upload_confirm(
            "/api/import/wechat-orders", "S1订单.xlsx", make_wx_xlsx("MS1"))
        assert st == 200, r
        # S2 灌入 MS2 前缀订单（member 携带本人第二店铺上传 → 组2 改动验证）
        st, r = cls.member.upload(
            f"/api/import/wechat-orders?store_id={cls.s2}",
            "S2订单.xlsx", make_wx_xlsx("MS2"))
        assert st == 200, r
        st, r = cls.member.post("/api/import/confirm", {
            "upload_id": r["data"]["upload_id"],
            "target_store_id": cls.s2})
        assert st == 200, r
        # 批量对账/重算: member 传 all → 本人两店逐店执行
        st, r = cls.member.post("/api/reconciliation/run",
                                {"target_store_id": "all", "month": MONTH})
        assert st == 200, r
        assert r["data"].get("store_count") == 2, r
        st, r = cls.member.post("/api/profit/recalculate",
                                {"target_store_id": "all", "month": MONTH})
        assert st == 200, r
        assert r["data"].get("store_count") == 2, r

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 登录绑定首有效店 ----------
    def test_01_login_binds_first_active_store(self):
        c = self.env.login_client("ms_a", PASSWORD)
        st, r = c.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["store_id"], self.s1)
        user = _find_user(self.env, "ms_a")
        self.assertEqual(len(user["store_ids"]), 2)

    # ---------- /api/stores 返回本人全部有效店铺 ----------
    def test_02_stores_returns_all_owned(self):
        st, r = self.member.get("/api/stores")
        self.assertEqual(st, 200)
        ids = {s["store_id"] for s in r["data"]["stores"]}
        self.assertEqual(ids, {self.s1, self.s2})
        for s in r["data"]["stores"]:
            self.assertEqual(s["owner_username"], "ms_a")
            self.assertEqual(s["owner_user_id"],
                             self.env.get_user_id("ms_a"))

    # ---------- 单店态: 多店间切换数据正确 ----------
    def test_03_switch_single_store(self):
        for sid, prefix in ((self.s1, "MS1"), (self.s2, "MS2")):
            st, r = self.member.get(
                f"/api/orders?store_id={sid}&page=1&size=100")
            self.assertEqual(st, 200)
            self.assertEqual(r["data"]["total"], 10)
            for o in r["data"]["orders"]:
                self.assertTrue(o["order_id"].startswith(prefix))
            # 经营总览与利润核算同步切换
            st, r = self.member.get(
                f"/api/summary?store_id={sid}&month={MONTH}")
            self.assertEqual(st, 200)
            st, r = self.member.get(
                f"/api/profit?store_id={sid}&page=1&size=20")
            self.assertEqual(st, 200)

    # ---------- 聚合态: 全部店铺 = S1+S2 之和 ----------
    def test_04_aggregate_all_stores(self):
        st, r = self.member.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 20)
        prefixes = {o["order_id"][:3] for o in r["data"]["orders"]}
        self.assertEqual(prefixes, {"MS1", "MS2"})
        self.assertEqual(r["data"].get("store_count"), 2)
        # 操作日志聚合
        st, r = self.member.get("/api/logs?page=1&page_size=100")
        self.assertEqual(st, 200)
        log_sids = {e.get("store_id") for e in r["data"]["entries"]}
        self.assertIn(self.s1, log_sids)
        self.assertIn(self.s2, log_sids)

    # ---------- 导出随店铺切换 ----------
    def test_05_export_follows_store(self):
        st, body, headers = self.member.get_binary(
            f"/api/orders/export?store_id={self.s2}")
        self.assertEqual(st, 200)
        wb = load_workbook(__import__("io").BytesIO(body))
        order_ids = [str(r[0]) for r in wb.active.iter_rows(
            min_row=2, values_only=True)]
        self.assertTrue(order_ids)
        self.assertTrue(all(a.startswith("MS2") for a in order_ids))

    # ---------- 拒绝态: 他店/不存在 → 403 + 安全日志 ----------
    def test_06_forbidden_scopes(self):
        other = self.env.member_store  # 其他成员的店铺
        for bad in (other, "s_nope999", "all_x"):
            st, r = self.member.get(
                f"/api/orders?store_id={bad}&page=1&size=20")
            self.assertEqual(st, 403, bad)
        # member 参数 → 403
        st, r = self.member.get(
            f"/api/orders?member={self.env.get_user_id('ms_a')}")
        self.assertEqual(st, 403)
        entries = _read_security_log(self.env)
        denied = [e for e in entries
                  if e["event"] == "access_denied"
                  and e.get("username") == "ms_a"]
        reasons = {e["detail"].get("reason") for e in denied}
        self.assertIn("store_scope_forbidden", reasons)
        self.assertIn("member_filter_forbidden", reasons)

    # ---------- 他店记录 ID 详情 → 404 ----------
    def test_07_other_store_detail_404(self):
        st, r = self.member.get(
            f"/api/orders?store_id={self.s2}&page=1&size=20")
        other_order = r["data"]["orders"][0]["order_id"]
        st, r = self.member.get(f"/api/orders/{other_order}")
        # 该订单属于本人 S2 → 应可见（单店态放行）
        self.assertEqual(st, 200)
        # 他店（biz_carrier）的订单号不存在于本人范围 → 404
        st, r = self.member.get("/api/orders/E2E0001")
        self.assertEqual(st, 404)

    # ---------- 写操作三态: 携带本人第二店铺 → 落该店 ----------
    def test_08_write_to_second_store(self):
        st, r = self.member.post("/api/ad-expenses", {
            "month": MONTH, "amount": "66.60",
            "note": "S2广告费", "target_store_id": self.s2})
        self.assertEqual(st, 200, r)
        st, r = self.member.get(f"/api/ad-expenses?store_id={self.s2}")
        self.assertEqual(st, 200)
        amounts = [a.get("amount") for a in r["data"].get("records", [])]
        self.assertIn("66.60", [str(a) for a in amounts])
        # S1 无该广告费
        st, r = self.member.get(f"/api/ad-expenses?store_id={self.s1}")
        self.assertEqual(st, 200)
        self.assertNotIn(
            "66.60", [str(a.get("amount"))
                      for a in r["data"].get("records", [])])
        # 写向他店 → 403
        st, r = self.member.post("/api/ad-expenses", {
            "month": MONTH, "amount": "1.00",
            "target_store_id": self.env.member_store})
        self.assertEqual(st, 403)

    # ---------- 上传目标: 携带本人第二店铺 → 落该店 ----------
    def test_09_upload_lands_second_store(self):
        st, r = self.member.get(
            f"/api/orders?store_id={self.s2}&page=1&size=100")
        self.assertEqual(r["data"]["total"], 10)  # setUpClass 已灌入 MS2

    # ---------- 批量写: member 传 all → 两店逐店执行 ----------
    def test_10_batch_all_covers_both(self):
        st, r = self.member.post("/api/profit/recalculate",
                                 {"target_store_id": "all", "month": MONTH})
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["store_count"], 2)
        sids = {s["store_id"] for s in r["data"]["stores"]}
        self.assertEqual(sids, {self.s1, self.s2})

    # ---------- 停用店铺: 400 + 从业务口径消失 ----------
    def test_11_disabled_store_rejected(self):
        stores_path = os.path.join(self.env.data_dir, "stores.json")
        with open(stores_path, encoding="utf-8") as f:
            data = json.load(f)
        for s in data["stores"]:
            if s["store_id"] == self.s2:
                s["status"] = "disabled"
        with open(stores_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        try:
            # 携带本人已停用店铺 → 400
            st, r = self.member.get(
                f"/api/orders?store_id={self.s2}&page=1&size=20")
            self.assertEqual(st, 400)
            self.assertIn("停用", r["message"])
            # 停用店不出现在 /api/stores
            st, r = self.member.get("/api/stores")
            ids = {s["store_id"] for s in r["data"]["stores"]}
            self.assertEqual(ids, {self.s1})
            # 聚合仅剩 S1
            st, r = self.member.get("/api/orders?page=1&size=100")
            self.assertEqual(r["data"]["total"], 10)
            # 恢复 active 后可再次访问
            with open(stores_path, encoding="utf-8") as f:
                data = json.load(f)
            for s in data["stores"]:
                if s["store_id"] == self.s2:
                    s["status"] = "active"
            with open(stores_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
            st, r = self.member.get(
                f"/api/orders?store_id={self.s2}&page=1&size=20")
            self.assertEqual(st, 200)
        finally:
            # 兜底恢复, 避免影响其他用例
            with open(stores_path, encoding="utf-8") as f:
                data = json.load(f)
            for s in data["stores"]:
                if s["store_id"] == self.s2:
                    s["status"] = "active"
            with open(stores_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)

    # ---------- 性能抽测: 切换劣化 ≤50ms ----------
    def test_12_switch_performance(self):
        def _best(sid):
            best = None
            for _ in range(3):
                t0 = time.monotonic()
                st, _r = self.member.get(
                    f"/api/orders?store_id={sid}&page=1&size=20")
                self.assertEqual(st, 200)
                dt = time.monotonic() - t0
                best = dt if best is None else min(best, dt)
            return best

        t1 = _best(self.s1)
        t2 = _best(self.s2)
        # 单店切换请求耗时差（同构数据同构查询）≤ 50ms
        self.assertLessEqual(abs(t2 - t1), 0.05,
                             f"S1={t1:.4f}s S2={t2:.4f}s")


class TestCreateStore(unittest.TestCase):
    """POST /api/stores 创建店铺（spec 5.3.1 规则 1-7 / design 2.2.2 A1）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member_id = cls.env.get_user_id("biz_carrier")

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_create_success_full_fields(self):
        stores_before = len(_read_stores(self.env))
        st, r = self.admin.post("/api/stores", {
            "name": "创建店一号", "owner_user_id": self.member_id})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["message"], "店铺创建成功")
        data = r["data"]
        self.assertTrue(data["store_id"].startswith("s_"))
        self.assertEqual(data["name"], "创建店一号")
        self.assertEqual(data["owner_user_id"], self.member_id)
        self.assertEqual(data["owner_username"], "biz_carrier")
        self.assertEqual(data["status"], "active")
        self.assertEqual(data["assignment_status"], "assigned")
        self.assertIsNotNone(data["created_at"])
        self.assertIsNotNone(data["updated_at"])
        # 两表落盘: 六字段完整 + 目标成员 store_ids 追加
        store = _find_store(self.env, data["store_id"])
        self.assertIsNotNone(store)
        self.assertEqual(store["owner_user_id"], self.member_id)
        self.assertEqual(store["status"], "active")
        self.assertEqual(store["assignment_status"], "assigned")
        user = _find_user(self.env, "biz_carrier")
        self.assertIn(data["store_id"], user["store_ids"])
        self.assertEqual(len(_read_stores(self.env)), stores_before + 1)
        # 审计 store.create 五要素
        entries = [e for e in _read_security_log(self.env)
                   if e["event"] == "store.create"]
        self.assertTrue(any(
            e.get("username") == self.env.admin_user
            and e["detail"].get("store_id") == data["store_id"]
            for e in entries))

    def test_02_owner_required(self):
        st, r = self.admin.post("/api/stores", {"name": "无归属店"})
        self.assertEqual(st, 400)
        self.assertEqual(r["message"], "请指定归属成员")

    def test_03_owner_not_found(self):
        st, r = self.admin.post("/api/stores", {
            "name": "幽灵归属店", "owner_user_id": "u_nope0000"})
        self.assertEqual(st, 404)
        self.assertEqual(r["message"], "目标用户不存在: u_nope0000")

    def test_04_owner_disabled_rejected(self):
        st, r = self.admin.post("/api/users", {
            "username": "cs_disabled", "password": PASSWORD})
        self.assertEqual(st, 200, r)
        uid = _find_user(self.env, "cs_disabled")["user_id"]
        st, r = self.admin.post(f"/api/users/{uid}/status",
                                {"action": "disable"})
        self.assertEqual(st, 200, r)
        st, r = self.admin.post("/api/stores", {
            "name": "停用归属店", "owner_user_id": uid})
        self.assertEqual(st, 400)
        self.assertEqual(r["message"], "目标成员已停用，请先启用")

    def test_05_owner_admin_rejected(self):
        admin_id = self.env.get_user_id(self.env.admin_user)
        st, r = self.admin.post("/api/stores", {
            "name": "管理员店", "owner_user_id": admin_id})
        self.assertEqual(st, 400)
        self.assertEqual(r["message"], "不能将店铺分配给管理员账号")

    def test_06_name_validation(self):
        for bad in (None, "", "   ", "x" * 65):
            st, r = self.admin.post("/api/stores", {
                "name": bad, "owner_user_id": self.member_id})
            self.assertEqual(st, 400, bad)
            self.assertEqual(r["message"], "请提供店铺名称（1-64 字符）")
        st, r = self.admin.post("/api/stores", {
            "name": "y" * 64, "owner_user_id": self.member_id})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["name"], "y" * 64)

    def test_07_unique_store_ids(self):
        ids = []
        for i in (1, 2):
            st, r = self.admin.post("/api/stores", {
                "name": f"连续创建店{i}", "owner_user_id": self.member_id})
            self.assertEqual(st, 200, r)
            ids.append(r["data"]["store_id"])
        self.assertNotEqual(ids[0], ids[1])
        all_ids = [s["store_id"] for s in _read_stores(self.env)]
        self.assertEqual(len(all_ids), len(set(all_ids)))

    def test_08_member_forbidden(self):
        st, r = self.env.member.post("/api/stores", {
            "name": "越权店", "owner_user_id": self.member_id})
        self.assertEqual(st, 403)
        entries = [e for e in _read_security_log(self.env)
                   if e["event"] == "access_denied"
                   and e.get("username") == "biz_carrier"]
        self.assertTrue(entries)

    def test_09_anonymous_unauthorized(self):
        st, r = Client(self.env).post("/api/stores", {
            "name": "匿名店", "owner_user_id": self.member_id})
        self.assertEqual(st, 401)

    def test_10_failed_create_no_residue(self):
        stores_before = _read_stores(self.env)
        # 归属校验失败（停用成员）→ 两表零残留
        st, r = self.admin.post("/api/stores", {
            "name": "失败残留店",
            "owner_user_id": _find_user(self.env, "cs_disabled")["user_id"]})
        self.assertEqual(st, 400)
        self.assertEqual(len(_read_stores(self.env)), len(stores_before))


class TestStoreStatus(unittest.TestCase):
    """POST /api/stores/{id}/status 停用/启用/归档（spec 5.5.3 / design A3）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            "ss_a", PASSWORD, store_name="状态店S1")
        cls.s1 = cls.member.get("/api/auth/me")[1]["data"]["store_id"]
        st, r = cls.member.upload_confirm(
            "/api/import/wechat-orders", "S1订单.xlsx", make_wx_xlsx("SS1"))
        assert st == 200, r
        st, r = cls.member.upload_confirm(
            "/api/import/supplier-costs", "S1成本.xlsx", make_sup_xlsx("SS1"))
        assert st == 200, r
        st, r = cls.member.post("/api/reconciliation/run",
                                {"target_store_id": "all", "month": MONTH})
        assert st == 200, r
        st, r = cls.member.post("/api/profit/recalculate",
                                {"target_store_id": "all", "month": MONTH})
        assert st == 200, r
        # admin 为 ss_a 创建第二店（空店）
        st, r = cls.admin.post("/api/stores", {
            "name": "状态店S2", "owner_user_id": cls.env.get_user_id("ss_a")})
        assert st == 200, r
        cls.s2 = r["data"]["store_id"]
        # 仅一店成员（验证停用独店后登录 401）
        cls.single = cls.env.create_user_client(
            "ss_single", PASSWORD, store_name="独店")
        cls.s3 = cls.single.get("/api/auth/me")[1]["data"]["store_id"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _inventory(self, store_id):
        st, r = self.admin.get(f"/api/stores/{store_id}")
        assert st == 200, r
        return r["data"]["inventory"]

    def test_01_disable_removes_from_scope_data_preserved(self):
        inv_before = self._inventory(self.s1)
        self.assertGreater(inv_before["orders"], 0)
        st, r = self.admin.post(f"/api/stores/{self.s1}/status",
                                {"action": "disable"})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["message"], "店铺已停用")
        data = r["data"]
        self.assertEqual(data["status"], "disabled")
        self.assertEqual(data["previous_status"], "active")
        self.assertEqual(data["store_id"], self.s1)
        self.assertTrue(data["data_preserved"])
        self.assertIsNotNone(data["backup_dir"])
        self.assertGreaterEqual(data["revoked_sessions"], 1)
        self.assertIn("orders", data["record_counts"])
        # 停用店可经 all-stores status 筛选查到
        st, r = self.admin.get(
            f"/api/admin/all-stores?status=disabled&page=1&page_size=50")
        self.assertEqual(st, 200)
        self.assertIn(self.s1,
                      {s["store_id"] for s in r["data"]["stores"]})
        # member 会话被吊销
        st, r = self.member.get("/api/orders?page=1&size=20")
        self.assertEqual(st, 401)
        # 重登重绑剩余有效店铺（S2）, S1 从业务口径消失
        relogged = self.env.login_client("ss_a", PASSWORD)
        st, r = relogged.get("/api/auth/me")
        self.assertEqual(r["data"]["store_id"], self.s2)
        st, r = relogged.get("/api/stores")
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.s2})
        # 十项清点逐项不变（数据完整保留）
        self.assertEqual(self._inventory(self.s1), inv_before)

    def test_02_enable_restores_scope(self):
        st, r = self.admin.post(f"/api/stores/{self.s1}/status",
                                {"action": "enable"})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["message"], "店铺已启用")
        self.assertEqual(r["data"]["status"], "active")
        self.assertEqual(r["data"]["previous_status"], "disabled")
        member = self.env.login_client("ss_a", PASSWORD)
        st, r = member.get("/api/stores")
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.s1, self.s2})
        st, r = member.get(f"/api/orders?store_id={self.s1}&page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 10)

    def test_03_archive_rejected_and_traceable(self):
        inv_before = self._inventory(self.s1)
        st, r = self.admin.post(f"/api/stores/{self.s1}/status",
                                {"action": "archive"})
        self.assertEqual(st, 400, r)
        self.assertEqual(r["message"], "归档功能已下线，不支持该操作")
        self.assertIsNone(r["data"])
        # 状态零变化: 仍为 active, 数据完整保留
        st, r = self.admin.get(f"/api/stores/{self.s1}")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["status"], "active")
        self.assertEqual(r["data"]["inventory"], inv_before)
        # 审计留痕: store.archive_removed（denied）
        entries = [e for e in _read_security_log(self.env)
                   if e["event"] == "store.archive_removed"
                   and e["detail"].get("store_id") == self.s1]
        self.assertTrue(entries)
        self.assertEqual(entries[-1]["result"], "denied")

    def test_04_archive_rejected_again(self):
        st, r = self.admin.post(f"/api/stores/{self.s1}/status",
                                {"action": "archive"})
        self.assertEqual(st, 400, r)
        self.assertEqual(r["message"], "归档功能已下线，不支持该操作")
        # 归档下线后不存在幂等成功; 重复拒绝均审计
        entries = [e for e in _read_security_log(self.env)
                   if e["event"] == "store.archive_removed"
                   and e["detail"].get("store_id") == self.s1]
        self.assertGreaterEqual(len(entries), 2)

    def test_05_invalid_transitions_rejected(self):
        # s1 当前为 active: enable 幂等返回 200, disable 合法返回 200
        st, r = self.admin.post(f"/api/stores/{self.s1}/status",
                                {"action": "enable"})
        self.assertEqual(st, 200, r)
        self.assertTrue(r["data"]["idempotent"])
        # 非法 action 被拒
        st, r = self.admin.post(f"/api/stores/{self.s2}/status",
                                {"action": "foo"})
        self.assertEqual(st, 400)
        self.assertEqual(r["message"], "action 仅支持 disable / enable（归档功能已下线）")
        # archive 同样被拒（归档已下线）
        st, r = self.admin.post(f"/api/stores/{self.s2}/status",
                                {"action": "archive"})
        self.assertEqual(st, 400)
        self.assertEqual(r["message"], "归档功能已下线，不支持该操作")
        st, r = self.admin.post("/api/stores/s_nope999/status",
                                {"action": "disable"})
        self.assertEqual(st, 404)

    def test_06_archive_rejected_login_preserved(self):
        member = self.env.login_client("ss_a", PASSWORD)
        st, r = self.admin.post(f"/api/stores/{self.s2}/status",
                                {"action": "archive"})
        self.assertEqual(st, 400, r)
        self.assertEqual(r["message"], "归档功能已下线，不支持该操作")
        # 归档被拒, 会话不吊销, 登录仍正常, 两店均在
        st, r = member.get("/api/stores")
        self.assertEqual(st, 200)
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.s1, self.s2})

    def test_07_single_store_member_disable_login_denied(self):
        st, r = self.admin.post(f"/api/stores/{self.s3}/status",
                                {"action": "disable"})
        self.assertEqual(st, 200, r)
        st, r = Client(self.env).post("/api/auth/login", {
            "username": "ss_single", "password": PASSWORD}, csrf=False)
        self.assertEqual(st, 401)
        self.assertEqual(r["message"], "账号未绑定店铺，请联系管理员")

    def test_08_member_forbidden(self):
        st, r = self.env.member.post(f"/api/stores/{self.s3}/status",
                                     {"action": "disable"})
        self.assertEqual(st, 403)
        entries = [e for e in _read_security_log(self.env)
                   if e["event"] == "access_denied"
                   and e.get("username") == "biz_carrier"]
        self.assertTrue(entries)

    def test_09_audit_five_elements(self):
        entries = _read_security_log(self.env)
        pairs = {(e["event"], e["detail"].get("store_id"))
                 for e in entries
                 if e["event"].startswith("store.")}
        self.assertIn(("store.disable", self.s1), pairs)
        self.assertIn(("store.enable", self.s1), pairs)
        self.assertIn(("store.archive_removed", self.s1), pairs)
        self.assertIn(("store.archive_removed", self.s2), pairs)
        self.assertIn(("store.disable", self.s3), pairs)
        # 五要素: time/event/username/user_id/ip + detail
        for e in entries:
            if e["event"].startswith("store.") and e["event"] != "store.create":
                self.assertIn("time", e)
                self.assertEqual(e["username"], self.env.admin_user)
                self.assertIsNotNone(e["user_id"])
                self.assertIn("ip", e)
                self.assertIn("store_id", e["detail"])


class TestStoreDetail(unittest.TestCase):
    """GET /api/stores/{id} 详情与十项清点（spec 5.4.1 / design 2.2.2 A2）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            "sd_a", PASSWORD, store_name="详情数据店")
        cls.s1 = cls.member.get("/api/auth/me")[1]["data"]["store_id"]
        st, r = cls.member.upload_confirm(
            "/api/import/wechat-orders", "订单.xlsx", make_wx_xlsx("SD1"))
        assert st == 200, r
        st, r = cls.member.upload_confirm(
            "/api/import/supplier-costs", "成本.xlsx", make_sup_xlsx("SD1"))
        assert st == 200, r
        st, r = cls.member.post("/api/reconciliation/run",
                                {"target_store_id": "all", "month": MONTH})
        assert st == 200, r
        st, r = cls.member.post("/api/profit/recalculate",
                                {"target_store_id": "all", "month": MONTH})
        assert st == 200, r
        st, r = cls.member.post("/api/ad-expenses", {
            "month": MONTH, "amount": "88.80", "note": "详情广告费"})
        assert st == 200, r

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_detail_inventory_matches_files(self):
        st, r = self.admin.get(f"/api/stores/{self.s1}")
        self.assertEqual(st, 200)
        data = r["data"]
        self.assertEqual(data["store_id"], self.s1)
        self.assertEqual(data["name"], "详情数据店")
        self.assertEqual(data["owner_user_id"],
                         self.env.get_user_id("sd_a"))
        self.assertEqual(data["owner_username"], "sd_a")
        self.assertEqual(data["status"], "active")
        self.assertEqual(data["assignment_status"], "assigned")
        self.assertFalse(data["orphan"])
        self.assertIsNotNone(data["created_at"])
        inv = data["inventory"]
        # 十项清点与店铺业务文件实际内容一致（inventory 键 → 文件键）
        file_keys = {"profit_results": "profit_statements"}
        for key in ("orders", "refund_orders", "refund_details",
                    "supplier_costs", "ad_expenses", "profit_results",
                    "import_records", "anomalies"):
            file_key = file_keys.get(key, key)
            self.assertEqual(
                inv[key],
                len(_read_business_records(self.env, self.s1, file_key)),
                f"清点 {key} 与文件记录数不一致")
        refund_details = _read_business_records(self.env, self.s1,
                                                "refund_details")
        costs = _read_business_records(self.env, self.s1, "supplier_costs")
        self.assertEqual(inv["refund_amount"],
                         _sum_records(refund_details, "refund_amount"))
        self.assertEqual(inv["cost_amount"],
                         _sum_records(costs, "total_amount"))
        # orders 为行级记录（同单多商品行）, 与 API 订单级去重口径不同
        self.assertGreater(inv["orders"], 0)
        self.assertEqual(inv["ad_expenses"], 1)

    def test_02_store_not_found(self):
        st, r = self.admin.get("/api/stores/s_nope999")
        self.assertEqual(st, 404)

    def test_03_member_forbidden(self):
        st, r = self.member.get(f"/api/stores/{self.s1}")
        self.assertEqual(st, 403)
        entries = [e for e in _read_security_log(self.env)
                   if e["event"] == "access_denied"
                   and e.get("username") == "sd_a"]
        self.assertTrue(entries)

    def test_04_archive_rejected_store_traceable(self):
        st, r = self.admin.get(f"/api/stores/{self.s1}")
        inv_before = r["data"]["inventory"]
        st, r = self.admin.post(f"/api/stores/{self.s1}/status",
                                {"action": "archive"})
        self.assertEqual(st, 400, r)
        self.assertEqual(r["message"], "归档功能已下线，不支持该操作")
        # 归档被拒, 状态零变化, 仍可只读追溯
        st, r = self.admin.get(f"/api/stores/{self.s1}")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["status"], "active")
        self.assertEqual(r["data"]["inventory"], inv_before)


class TestAdminAllStoresExtension(unittest.TestCase):
    """all-stores 筛选扩展: store_id + status 两维（spec 5.4.1 规则 2）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.u1 = cls.env.create_user_client(
            "ax_a", PASSWORD, store_name="扩展A店")
        cls.s1 = cls.u1.get("/api/auth/me")[1]["data"]["store_id"]
        st, r = cls.admin.post("/api/stores", {
            "name": "扩展B店", "owner_user_id": cls.env.get_user_id("ax_a")})
        assert st == 200, r
        cls.s2 = r["data"]["store_id"]
        st, r = cls.admin.post(f"/api/stores/{cls.s2}/status",
                               {"action": "disable"})
        assert st == 200, r
        cls.u2 = cls.env.create_user_client(
            "ax_b", PASSWORD, store_name="扩展C店")
        cls.s3 = cls.u2.get("/api/auth/me")[1]["data"]["store_id"]
        st, r = cls.admin.post(f"/api/stores/{cls.s3}/status",
                               {"action": "disable"})
        assert st == 200, r

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_status_field_present(self):
        st, r = self.admin.get("/api/admin/all-stores")
        self.assertEqual(st, 200)
        valid = {"active", "disabled"}
        for s in r["data"]["stores"]:
            self.assertIn(s["status"], valid)
        by_id = {s["store_id"]: s["status"] for s in r["data"]["stores"]}
        self.assertEqual(by_id[self.s1], "active")
        self.assertEqual(by_id[self.s2], "disabled")
        self.assertEqual(by_id[self.s3], "disabled")

    def test_02_store_id_filter(self):
        st, r = self.admin.get(
            f"/api/admin/all-stores?store_id={self.s2}&page=1&page_size=50")
        self.assertEqual(st, 200)
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.s2})
        # 关键词包含匹配（前缀）
        kw = self.s2[:6]
        st, r = self.admin.get(
            f"/api/admin/all-stores?store_id={kw}&page=1&page_size=50")
        self.assertEqual(st, 200)
        ids = {s["store_id"] for s in r["data"]["stores"]}
        self.assertIn(self.s2, ids)
        self.assertNotIn(self.s1, ids)

    def test_03_status_filter(self):
        # disabled 筛选仍 200（保留）
        st, r = self.admin.get(
            "/api/admin/all-stores?status=disabled&page=1&page_size=100")
        self.assertEqual(st, 200)
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.s2, self.s3})
        # 归档已下线 → archived 筛选预期 400
        st, r = self.admin.get(
            "/api/admin/all-stores?status=archived&page=1&page_size=100")
        self.assertEqual(st, 400)
        self.assertEqual(r["message"], "店铺状态仅支持 active / disabled")
        st, r = self.admin.get(
            "/api/admin/all-stores?status=active&page=1&page_size=100")
        ids = {s["store_id"] for s in r["data"]["stores"]}
        self.assertIn(self.s1, ids)
        self.assertNotIn(self.s2, ids)
        self.assertNotIn(self.s3, ids)

    def test_04_combined_filter_intersection(self):
        # store_id 精确 + status → 交集为空（S2 disabled 非 active）
        st, r = self.admin.get(
            f"/api/admin/all-stores?store_id={self.s2}&status=active"
            f"&page=1&page_size=50")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 0)
        # owner + status 组合
        st, r = self.admin.get(
            f"/api/admin/all-stores?owner=ax_a&status=active"
            f"&page=1&page_size=50")
        self.assertEqual(st, 200)
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.s1})

    def test_05_pagination_kept(self):
        st, r = self.admin.get(
            "/api/admin/all-stores?status=active&page=1&page_size=20")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["page"], 1)
        self.assertEqual(r["data"]["page_size"], 20)
        self.assertEqual(r["data"]["total"],
                         len(r["data"]["stores"]))
        # 既有筛选维度零变化（URL 中文需 percent-encode）
        st, r = self.admin.get(
            f"/api/admin/all-stores?name={quote('扩展A店')}"
            f"&page=1&page_size=50")
        self.assertEqual(st, 200)
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.s1})

    def test_06_invalid_status_rejected(self):
        st, r = self.admin.get(
            "/api/admin/all-stores?status=foo&page=1&page_size=20")
        self.assertEqual(st, 400)
        self.assertEqual(r["message"],
                         "店铺状态仅支持 active / disabled")


class TestOneToManyOwnership(unittest.TestCase):
    """一对多归属验收（spec 5.1.3 规则 3/4 + 异常 2 / 5.2.1 规则 3）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        # 成员 A: 无初始店（create_store=False）, 后续创建 2 店 + 转入 1 店
        st, r = cls.admin.post("/api/users", {
            "username": "oto_a", "password": PASSWORD,
            "create_store": False})
        assert st == 200, r
        cls.owner_id = cls.env.get_user_id("oto_a")
        cls.created = []
        for name in ("A的创建一店", "A的创建二店"):
            st, r = cls.admin.post("/api/stores", {
                "name": name, "owner_user_id": cls.owner_id})
            assert st == 200, r
            cls.created.append(r["data"]["store_id"])
        cls.s1, cls.s2 = cls.created
        # 店铺就绪后方可登录（无店 member 登录被拒）
        cls.owner = cls.env.login_client("oto_a", PASSWORD)
        # 转入 helper 的初始店
        cls.helper = cls.env.create_user_client(
            "oto_b", PASSWORD, store_name="待转入店")
        cls.s3 = cls.helper.get("/api/auth/me")[1]["data"]["store_id"]
        st, r = cls.admin.post(
            f"/api/stores/{cls.s3}/transfer-ownership",
            {"target_user_id": cls.owner_id})
        assert st == 200, r

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_three_stores_accessible(self):
        st, r = self.owner.get("/api/stores")
        self.assertEqual(st, 200)
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.s1, self.s2, self.s3})
        for sid in (self.s1, self.s2, self.s3):
            st, r = self.owner.get(
                f"/api/orders?store_id={sid}&page=1&size=20")
            self.assertEqual(st, 200, sid)

    def test_02_owner_field_scalar_unique(self):
        for s in _read_stores(self.env):
            if s["store_id"] in (self.s1, self.s2, self.s3):
                self.assertEqual(s["owner_user_id"], self.owner_id)
                self.assertNotIsInstance(s["owner_user_id"], list)
        user = _find_user(self.env, "oto_a")
        self.assertEqual(sorted(user["store_ids"]),
                         sorted([self.s1, self.s2, self.s3]))

    def test_03_transfer_removes_old_owner_binding(self):
        helper = _find_user(self.env, "oto_b")
        self.assertNotIn(self.s3, helper["store_ids"])
        self.assertEqual(helper["store_ids"], [])

    def test_04_transfer_to_current_owner_idempotent(self):
        user_before = _find_user(self.env, "oto_a")["store_ids"]
        st, r = self.admin.post(
            f"/api/stores/{self.s3}/transfer-ownership",
            {"target_user_id": self.owner_id})
        self.assertEqual(st, 200, r)
        self.assertEqual(_find_user(self.env, "oto_a")["store_ids"],
                         user_before)

    def test_05_transfer_to_admin_rejected(self):
        admin_id = self.env.get_user_id(self.env.admin_user)
        st, r = self.admin.post(
            f"/api/stores/{self.s1}/transfer-ownership",
            {"target_user_id": admin_id})
        self.assertEqual(st, 400)
        # 失败后归属不变
        store = _find_store(self.env, self.s1)
        self.assertEqual(store["owner_user_id"], self.owner_id)

    def test_06_same_name_stores_independent(self):
        ids = []
        for _ in range(2):
            st, r = self.admin.post("/api/stores", {
                "name": "同名店", "owner_user_id": self.owner_id})
            self.assertEqual(st, 200, r)
            ids.append(r["data"]["store_id"])
        sa, sb = ids
        # 停用 sa → sb 不受影响
        st, r = self.admin.post(f"/api/stores/{sa}/status",
                                {"action": "disable"})
        self.assertEqual(st, 200, r)
        st, r = self.admin.get(f"/api/stores/{sa}")
        self.assertEqual(r["data"]["status"], "disabled")
        st, r = self.admin.get(f"/api/stores/{sb}")
        self.assertEqual(r["data"]["status"], "active")
        # 改名 sb → sa 不受影响
        st, r = self.admin.post(f"/api/stores/{sb}/rename",
                                {"name": "改名后的同名店"})
        self.assertEqual(st, 200, r)
        st, r = self.admin.get(f"/api/stores/{sa}")
        self.assertEqual(r["data"]["name"], "同名店")


def _store_dir_exists(env, store_id):
    return os.path.isdir(os.path.join(env.data_dir, "stores", store_id))


def _ownership_backup_dirs(env):
    return [d for d in os.listdir(env.data_dir)
            if d.startswith("ownership_migration_")]


class TestStoreDelete(unittest.TestCase):
    """POST /api/stores/{id}/delete 物理删除分级（spec 5.5.1 / design A4）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            "del_a", PASSWORD, store_name="删除有数据店")
        cls.s1 = cls.member.get("/api/auth/me")[1]["data"]["store_id"]
        st, r = cls.member.upload_confirm(
            "/api/import/wechat-orders", "订单.xlsx", make_wx_xlsx("DL1"))
        assert st == 200, r
        st, r = cls.member.upload_confirm(
            "/api/import/supplier-costs", "成本.xlsx", make_sup_xlsx("DL1"))
        assert st == 200, r
        st, r = cls.member.post("/api/reconciliation/run",
                                {"target_store_id": "all", "month": MONTH})
        assert st == 200, r
        st, r = cls.member.post("/api/profit/recalculate",
                                {"target_store_id": "all", "month": MONTH})
        assert st == 200, r
        cls.owner_id = cls.env.get_user_id("del_a")
        cls.s2, cls.s3 = [], []
        for name in ("删除空店", "删除备用店"):
            st, r = cls.admin.post("/api/stores", {
                "name": name, "owner_user_id": cls.owner_id})
            assert st == 200, r
            cls.s3.append(r["data"]["store_id"])
        cls.s2 = cls.s3[0]  # 空店: 全流程删除
        cls.s3 = cls.s3[1]  # 备用: 确认失败/越权/回滚用例

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_empty_store_delete_full_flow(self):
        store = _find_store(self.env, self.s2)
        self.assertIsNotNone(store)
        st, r = self.admin.post(f"/api/stores/{self.s2}/delete", {
            "confirm_store_name": "删除空店",
            "confirm_acknowledged": True})  # 强确认双通道
        self.assertEqual(st, 200, r)
        self.assertEqual(r["message"], "店铺已永久删除（含全部关联业务数据）")
        data = r["data"]
        self.assertEqual(data["store_id"], self.s2)
        self.assertEqual(data["name"], "删除空店")
        self.assertEqual(data["owner_username"], "del_a")
        self.assertIsNotNone(data["txn_id"])
        self.assertTrue(data["removed_store_dir"])
        # 两表移除 + 目录移除 + all-stores 不可见
        self.assertIsNone(_find_store(self.env, self.s2))
        self.assertNotIn(self.s2,
                         _find_user(self.env, "del_a")["store_ids"])
        self.assertFalse(_store_dir_exists(self.env, self.s2))
        st, r = self.admin.get(
            "/api/admin/all-stores?page=1&page_size=100")
        self.assertNotIn(self.s2,
                         {s["store_id"] for s in r["data"]["stores"]})

    def test_02_has_trades_blocked(self):
        # 删除前直读落盘记录数（用于 deleted_counts 比对）
        orders_before = len(_read_business_records(self.env, self.s1, "orders"))
        # 有数据店铺直接删除成功（归档下线 + 物理删除新契约）
        st, r = self.admin.post(f"/api/stores/{self.s1}/delete", {
            "confirm_store_name": "删除有数据店",
            "confirm_acknowledged": True})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["deleted_counts"]["orders"], orders_before)
        # 删除后店铺记录、目录、归属引用确不存在
        self.assertIsNone(_find_store(self.env, self.s1))
        self.assertFalse(_store_dir_exists(self.env, self.s1))
        self.assertNotIn(self.s1, _find_user(self.env, "del_a")["store_ids"])

    def test_03_two_step_full_delete(self):
        # test_02 已删除 self.s1，本用例新建临时店铺验证两步删除链路
        st, r = self.admin.post("/api/stores", {
            "name": "两步删除临时店", "owner_user_id": self.env.get_user_id("del_a")})
        self.assertEqual(st, 200, r)
        tmp_sid = r["data"]["store_id"]
        st, r = self.admin.post(f"/api/stores/{tmp_sid}/clear-data", {
            "confirm_store_name": "两步删除临时店",
            "confirm_acknowledged": True})
        self.assertEqual(st, 200, r)
        clear_backup = r["data"]["backup_dir"]
        snap_orders = os.path.join(clear_backup, "data", "orders.json")
        self.assertTrue(os.path.exists(snap_orders))
        st, r = self.admin.post(f"/api/stores/{tmp_sid}/delete", {
            "confirm_store_name": "两步删除临时店",
            "confirm_acknowledged": True})
        self.assertEqual(st, 200, r)
        self.assertIsNone(_find_store(self.env, tmp_sid))
        self.assertFalse(_store_dir_exists(self.env, tmp_sid))
        self.assertIn("deleted_counts", r["data"])
        self.assertNotIn(tmp_sid,
                         _find_user(self.env, "del_a")["store_ids"])

    def test_04_confirm_mismatch_rejected(self):
        for bad in ("错误名", "", None, "删除有数据店 "):
            st, r = self.admin.post(f"/api/stores/{self.s3}/delete", {
                "confirm_store_name": bad,
                "confirm_acknowledged": True})
            self.assertEqual(st, 400, bad)
        self.assertIsNotNone(_find_store(self.env, self.s3))

    def test_05_member_forbidden(self):
        # del_a 会话已在 test_01 删除其店铺时被吊销, 重登后验证越权
        member = self.env.login_client("del_a", PASSWORD)
        st, r = member.post(f"/api/stores/{self.s3}/delete", {
            "confirm_store_name": "删除备用店"})
        self.assertEqual(st, 403)
        # 伪造角色字段不改后端判定
        st, r = member.post(f"/api/stores/{self.s3}/delete", {
            "confirm_store_name": "删除备用店", "role": "admin"})
        self.assertEqual(st, 403)
        self.assertIsNotNone(_find_store(self.env, self.s3))
        entries = [e for e in _read_security_log(self.env)
                   if e["event"] == "access_denied"
                   and e.get("username") == "del_a"]
        self.assertTrue(entries)

    def test_06_anonymous_unauthorized(self):
        st, r = Client(self.env).post(f"/api/stores/{self.s3}/delete", {
            "confirm_store_name": "删除备用店"})
        self.assertEqual(st, 401)

    def test_07_service_failure_zero_change(self):
        backups_before = len(_ownership_backup_dirs(self.env))
        stores_before = len(_read_stores(self.env))
        with mock.patch.object(
                self.env.server.user_manager, "delete_store",
                side_effect=AuthError("删除失败，请重试",
                                      code=500, http_status=500)):
            st, r = self.admin.post(f"/api/stores/{self.s3}/delete", {
                "confirm_store_name": "删除备用店",
                "confirm_acknowledged": True})
        self.assertEqual(st, 500)
        # 两表与目录回到操作前, 备份目录不残留, 可重试
        self.assertEqual(len(_read_stores(self.env)), stores_before)
        self.assertIsNotNone(_find_store(self.env, self.s3))
        self.assertIn(self.s3, _find_user(self.env, "del_a")["store_ids"])
        self.assertTrue(_store_dir_exists(self.env, self.s3))
        self.assertEqual(len(_ownership_backup_dirs(self.env)),
                         backups_before)

    def test_08_rmtree_failure_safe_state(self):
        original_rmtree = shutil.rmtree
        def mock_rmtree(*args, **kwargs):
            if kwargs.get("ignore_errors"):
                return original_rmtree(*args, **kwargs)
            raise OSError("denied")
        with mock.patch("shutil.rmtree", side_effect=mock_rmtree):
            st, r = self.admin.post(f"/api/stores/{self.s3}/delete", {
                "confirm_store_name": "删除备用店",
                "confirm_acknowledged": True})
        self.assertEqual(st, 200, r)
        # 记录已删（两表成功）; 目录残留属安全状态, 审计如实记录
        self.assertFalse(r["data"]["removed_store_dir"])
        self.assertIsNone(_find_store(self.env, self.s3))
        self.assertTrue(_store_dir_exists(self.env, self.s3))
        events = [e for e in _read_security_log(self.env)
                  if e["event"] == "store.delete"
                  and e["detail"].get("store_id") == self.s3]
        self.assertTrue(events)
        self.assertFalse(events[-1]["detail"]["removed_store_dir"])

    def test_09_audit_five_elements(self):
        events = [e for e in _read_security_log(self.env)
                  if e["event"] == "store.delete"]
        self.assertGreaterEqual(len(events), 3)
        for e in events:
            self.assertIn("time", e)
            self.assertEqual(e["username"], self.env.admin_user)
            self.assertIsNotNone(e["user_id"])
            self.assertIn("ip", e)
            for key in ("store_id", "store_name", "txn_id",
                        "removed_store_dir", "revoked_sessions"):
                self.assertIn(key, e["detail"])

    def test_10_no_orphan_records(self):
        # 已删店铺无目录残留; 唯一残留为可识别的快照备份
        for sid in (self.s1, self.s2):
            self.assertFalse(_store_dir_exists(self.env, sid))
        for d in _ownership_backup_dirs(self.env):
            note = os.path.join(self.env.data_dir, d,
                                "DELETED_STORE_NOTE.json")
            if os.path.exists(note):
                with open(note, encoding="utf-8") as f:
                    self.assertIn("store_id", json.load(f))


class TestStoreClearData(unittest.TestCase):
    """POST /api/stores/{id}/clear-data 清除单店数据（spec 5.5.1 / design A5）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            "clr_a", PASSWORD, store_name="清理数据店")
        cls.s1 = cls.member.get("/api/auth/me")[1]["data"]["store_id"]
        st, r = cls.member.upload_confirm(
            "/api/import/wechat-orders", "订单.xlsx", make_wx_xlsx("CL1"))
        assert st == 200, r
        st, r = cls.member.upload_confirm(
            "/api/import/supplier-costs", "成本.xlsx", make_sup_xlsx("CL1"))
        assert st == 200, r
        st, r = cls.member.post("/api/reconciliation/run",
                                {"target_store_id": "all", "month": MONTH})
        assert st == 200, r
        st, r = cls.member.post("/api/profit/recalculate",
                                {"target_store_id": "all", "month": MONTH})
        assert st == 200, r
        st, r = cls.member.post("/api/ad-expenses", {
            "month": MONTH, "amount": "12.34", "note": "清理前广告费"})
        assert st == 200, r
        cls.expected = {
            name: len(_read_business_records(cls.env, cls.s1, name))
            for name in BUSINESS_FILES if name != "settings"}
        # 边界验证: 另一成员两店, 清 S3 不影响 S4
        cls.member2 = cls.env.create_user_client(
            "clr_b", PASSWORD, store_name="边界三店")
        cls.s3 = cls.member2.get("/api/auth/me")[1]["data"]["store_id"]
        st, r = cls.member2.upload_confirm(
            "/api/import/wechat-orders", "B订单.xlsx", make_wx_xlsx("CL3"))
        assert st == 200, r
        st, r = cls.admin.post("/api/stores", {
            "name": "边界四店", "owner_user_id": cls.env.get_user_id("clr_b")})
        assert st == 200, r
        cls.s4 = r["data"]["store_id"]
        st, r = cls.member2.upload(
            f"/api/import/wechat-orders?store_id={cls.s4}",
            "B2订单.xlsx", make_wx_xlsx("CL4"))
        assert st == 200, r
        st, r = cls.member2.post("/api/import/confirm", {
            "upload_id": r["data"]["upload_id"],
            "target_store_id": cls.s4})
        assert st == 200, r

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_clear_success_preserves_structure(self):
        # 构造 uploads 暂存（未确认上传）供清理验证
        st, r = self.member.upload(
            f"/api/import/wechat-orders?store_id={self.member.store_id}",
            "暂存.xlsx", make_wx_xlsx("TMP"))
        self.assertEqual(st, 200, r)
        st, r = self.admin.post(f"/api/stores/{self.s1}/clear-data", {
            "confirm_store_name": "清理数据店",
            "confirm_acknowledged": True})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["message"], "店铺业务数据已清除（已创建整店快照备份）")
        data = r["data"]
        self.assertEqual(data["store_id"], self.s1)
        # records_before 与清理前文件实际记录数一致
        self.assertEqual({k: v for k, v in data["records_before"].items()
                          if self.expected.get(k)},
                         {k: v for k, v in self.expected.items() if v})
        self.assertGreater(data["uploads_removed"], 0)
        # 12 类业务文件清空; settings/操作日志/店铺记录/归属保留
        store_data = os.path.join(self.env.data_dir, "stores",
                                  self.s1, "data")
        for name in BUSINESS_FILES:
            if name == "settings":
                with open(os.path.join(store_data, "settings.json"),
                          encoding="utf-8") as f:
                    self.assertIsInstance(json.load(f), dict)
            else:
                records = _read_business_records(self.env, self.s1, name)
                self.assertEqual(records, [], name)
        self.assertTrue(os.path.exists(
            os.path.join(store_data, "settings.json")))
        op_log = os.path.join(self.env.data_dir, "stores", self.s1,
                              "logs", "operation_log.jsonl")
        self.assertTrue(os.path.exists(op_log))
        with open(op_log, encoding="utf-8") as f:
            self.assertTrue(f.read().strip())
        self.assertIsNotNone(_find_store(self.env, self.s1))
        self.assertIn(self.s1, _find_user(self.env, "clr_a")["store_ids"])
        uploads = os.path.join(self.env.data_dir, "stores", self.s1,
                               "uploads")
        if os.path.isdir(uploads):
            self.assertEqual(os.listdir(uploads), [])
        # 快照保留原数据
        snap_orders = os.path.join(data["backup_dir"], "data", "orders.json")
        self.assertTrue(os.path.exists(snap_orders))
        with open(snap_orders, encoding="utf-8") as f:
            self.assertGreater(len(json.load(f)["records"]), 0)

    def test_02_inventory_zero_after_clear(self):
        st, r = self.admin.get(f"/api/stores/{self.s1}")
        self.assertEqual(st, 200)
        inv = r["data"]["inventory"]
        for key, value in inv.items():
            if key.endswith("_amount"):
                self.assertEqual(value, "0.00", key)
            else:
                self.assertEqual(value, 0, key)

    def test_03_member_session_revoked_relogin(self):
        st, r = self.member.get("/api/stores")
        self.assertEqual(st, 401)  # 清理时会话已吊销
        relogged = self.env.login_client("clr_a", PASSWORD)
        st, r = relogged.get("/api/stores")
        self.assertEqual(st, 200)
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.s1})

    def test_04_confirm_mismatch_rejected(self):
        st, r = self.admin.post(f"/api/stores/{self.s3}/clear-data", {
            "confirm_store_name": "错误名称",
            "confirm_acknowledged": True})
        self.assertEqual(st, 400)
        self.assertIn("确认信息不一致", r["message"])
        self.assertGreater(
            len(_read_business_records(self.env, self.s3, "orders")), 0)

    def test_05_member_forbidden(self):
        st, r = self.member2.post(f"/api/stores/{self.s3}/clear-data", {
            "confirm_store_name": "边界三店"})
        self.assertEqual(st, 403)
        entries = [e for e in _read_security_log(self.env)
                   if e["event"] == "access_denied"
                   and e.get("username") == "clr_b"]
        self.assertTrue(entries)

    def test_06_single_store_boundary(self):
        s4_orders_before = len(
            _read_business_records(self.env, self.s4, "orders"))
        self.assertGreater(s4_orders_before, 0)
        st, r = self.admin.post(f"/api/stores/{self.s3}/clear-data", {
            "confirm_store_name": "边界三店",
            "confirm_acknowledged": True})
        self.assertEqual(st, 200, r)
        # 单店接口只清 S3; 同成员 S4 数据不受影响
        self.assertEqual(
            len(_read_business_records(self.env, self.s4, "orders")),
            s4_orders_before)
        relogged = self.env.login_client("clr_b", PASSWORD)
        st, r = relogged.get("/api/stores")
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.s3, self.s4})

    def test_07_rollback_on_failure(self):
        s4_orders_before = len(
            _read_business_records(self.env, self.s4, "orders"))
        bundle = self.env.server.registry.get(self.s4)
        with mock.patch.object(bundle.repo, "write",
                               side_effect=OSError("disk full")):
            st, r = self.admin.post(f"/api/stores/{self.s4}/clear-data", {
                "confirm_store_name": "边界四店",
                "confirm_acknowledged": True})
        self.assertEqual(st, 500)
        self.assertIn("已从快照回滚", r["message"])
        # 数据完整恢复到清理前状态
        self.assertEqual(
            len(_read_business_records(self.env, self.s4, "orders")),
            s4_orders_before)

    def test_08_audit_single_store_dimension(self):
        entries = [e for e in _read_security_log(self.env)
                   if e["event"] == "store.clear_data"
                   and e["detail"].get("store_id") == self.s1]
        self.assertEqual(len(entries), 1)
        detail = entries[0]["detail"]
        self.assertNotIn("target_user_id", detail)
        self.assertIn("records_before", detail)
        self.assertIn("backup_dir", detail)
        self.assertEqual(entries[0]["username"], self.env.admin_user)

    def test_09_two_step_delete_after_clear(self):
        st, r = self.admin.post(f"/api/stores/{self.s1}/delete", {
            "confirm_store_name": "清理数据店",
            "confirm_acknowledged": True})
        self.assertEqual(st, 200, r)
        self.assertIsNone(_find_store(self.env, self.s1))
        self.assertFalse(_store_dir_exists(self.env, self.s1))


class TestMigration(unittest.TestCase):
    """店铺数据模型迁移工具（spec 5.8.1 / design 2.4 + A6/A7）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            "mig_a", PASSWORD, store_name="迁移正常店")
        cls.s_ok = cls.member.get("/api/auth/me")[1]["data"]["store_id"]
        st, r = cls.member.upload_confirm(
            "/api/import/wechat-orders", "订单.xlsx", make_wx_xlsx("MG1"))
        assert st == 200, r
        # 构造异常样本: 直接改写 stores.json（迁移对象为存量缺字段记录）
        cls.orphan_owner = "u_deleted9999"
        cls.s_missing = None   # 缺三字段 + 有效归属
        cls.s_orphan = None    # 缺三字段 + 孤儿归属
        cls.s_unassigned = None  # 待分配缺 status
        cls.s_nodir = None     # 无目录店铺
        st, r = cls.admin.post("/api/stores", {
            "name": "迁移缺字段店", "owner_user_id": cls.env.get_user_id("mig_a")})
        assert st == 200, r
        cls.s_missing = r["data"]["store_id"]
        st, r = cls.admin.post("/api/stores", {
            "name": "迁移孤儿店", "owner_user_id": cls.env.get_user_id("mig_a")})
        assert st == 200, r
        cls.s_orphan = r["data"]["store_id"]
        st, r = cls.admin.post("/api/stores", {
            "name": "迁移待分配店", "owner_user_id": cls.env.get_user_id("mig_a")})
        assert st == 200, r
        cls.s_unassigned = r["data"]["store_id"]
        st, r = cls.admin.post("/api/stores", {
            "name": "迁移无目录店", "owner_user_id": cls.env.get_user_id("mig_a")})
        assert st == 200, r
        cls.s_nodir = r["data"]["store_id"]
        cls._rewrite_samples()
        # 无目录店铺样本: 移除其业务数据目录
        shutil.rmtree(os.path.join(cls.env.data_dir, "stores", cls.s_nodir))
        # 游离业务目录（有目录无店铺记录）
        os.makedirs(os.path.join(cls.env.data_dir, "stores", "s_stray0001"))

    @classmethod
    def _rewrite_samples(cls):
        path = os.path.join(cls.env.data_dir, "stores.json")
        with open(path, encoding="utf-8") as f:
            doc = json.load(f)
        for s in doc["stores"]:
            sid = s["store_id"]
            if sid == cls.s_missing:
                for key in ("status", "updated_at", "assignment_status"):
                    s.pop(key, None)
            elif sid == cls.s_orphan:
                for key in ("status", "updated_at", "assignment_status"):
                    s.pop(key, None)
                s["owner_user_id"] = cls.orphan_owner
            elif sid == cls.s_unassigned:
                s.pop("status", None)
                s["owner_user_id"] = None
                s["assignment_status"] = "unassigned"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_check_report_classification(self):
        st, r = self.admin.get("/api/admin/migration/model-check")
        self.assertEqual(st, 200)
        data = r["data"]
        # 全部店铺: biz_carrier 默认店 + 正常店 + 4 个构造样本
        self.assertEqual(data["total_stores"], 6)
        by_id = {it["store_id"]: it for it in data["items"]}
        # 正常店: 零 issue
        self.assertEqual(by_id[self.s_ok]["issues"], [])
        self.assertFalse(by_id[self.s_ok]["auto_fixable"])
        # 缺三字段 + 有效归属 → fill_fields, 可自动修复
        self.assertEqual(sorted(by_id[self.s_missing]["issues"]),
                         ["missing_assignment_status", "missing_status",
                          "missing_updated_at"])
        self.assertEqual(by_id[self.s_missing]["action"], "fill_fields")
        self.assertTrue(by_id[self.s_missing]["auto_fixable"])
        # 孤儿店 → fill_fields+unassign_owner
        self.assertEqual(by_id[self.s_orphan]["action"],
                         "fill_fields+unassign_owner")
        self.assertTrue(by_id[self.s_orphan]["auto_fixable"])
        # 待分配店 → 仅补 status
        self.assertEqual(by_id[self.s_unassigned]["issues"],
                         ["missing_status"])
        # 无目录店 → 歧义待核对, 不自动处理
        self.assertFalse(by_id[self.s_nodir]["auto_fixable"])
        self.assertIn("no_store_dir", by_id[self.s_nodir]["issues"])
        ambiguous_ids = {a["store_id"] for a in data["ambiguous"]}
        self.assertIn(self.s_nodir, ambiguous_ids)
        # 游离目录 → business_data_orphan
        stray = {b["store_id"] for b in data["business_data_orphan"]}
        self.assertIn("s_stray0001", stray)

    def test_02_migrate_fills_and_clears(self):
        st, r = self.admin.post("/api/admin/migration/model-migrate", {})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["message"], "迁移完成")
        data = r["data"]
        migrated_by_id = {m["store_id"]: m for m in data["migrated"]}
        self.assertEqual(set(migrated_by_id),
                         {self.s_missing, self.s_orphan,
                          self.s_unassigned})
        # 歧义样本跳过且保留原样
        skipped_ids = {s["store_id"] for s in data["skipped"]}
        self.assertIn(self.s_nodir, skipped_ids)
        # 字段补齐且默认值符合口径
        store = _find_store(self.env, self.s_missing)
        self.assertEqual(store["status"], "active")
        self.assertEqual(store["assignment_status"], "assigned")
        self.assertIsNotNone(store["updated_at"])
        self.assertEqual(store["owner_user_id"],
                         self.env.get_user_id("mig_a"))
        # 孤儿店铺: owner 置空 + unassigned + 溯源保留
        store = _find_store(self.env, self.s_orphan)
        self.assertIsNone(store["owner_user_id"])
        self.assertEqual(store["assignment_status"], "unassigned")
        self.assertEqual(store["previous_owner_user_id"], self.orphan_owner)
        self.assertEqual(store["status"], "active")
        self.assertTrue(migrated_by_id[self.s_orphan]["owner_cleared"])
        # 待分配店仅补 status, 归属保持空
        store = _find_store(self.env, self.s_unassigned)
        self.assertEqual(store["status"], "active")
        self.assertIsNone(store["owner_user_id"])
        self.assertEqual(store["assignment_status"], "unassigned")
        # store_id 与业务数据记录数不变
        store = _find_store(self.env, self.s_ok)
        self.assertIsNotNone(store)
        self.assertEqual(
            len(_read_business_records(self.env, self.s_ok, "orders")), 13)
        # 带时间戳两表备份目录存在且内容完整
        backup_dir = data["backup_dir"]
        self.assertTrue(os.path.basename(backup_dir)
                        .startswith("store_model_migration_"))
        for name in ("users.json", "stores.json"):
            self.assertTrue(os.path.exists(
                os.path.join(backup_dir, name)))

    def test_03_verification_report_matches(self):
        st, r = self.admin.post("/api/admin/migration/model-migrate", {})
        self.assertEqual(st, 200, r)
        ver = r["data"]["verification"]
        self.assertTrue(ver["match"])
        self.assertEqual(ver["before"], ver["after"])
        self.assertIn(self.s_ok, ver["before"])
        # 无目录店铺豁免清点且显式列出
        self.assertIn(self.s_nodir, ver["excluded_no_dir"])

    def test_04_rollback_on_mismatch(self):
        self._rewrite_samples()  # 重新制造缺字段样本
        with open(os.path.join(self.env.data_dir, "stores.json"),
                  encoding="utf-8") as f:
            doc_before = json.load(f)
        original = self.env.server._store_inventory
        calls = {"n": 0}

        def flaky(store_id):
            if store_id == self.s_ok:
                calls["n"] += 1
            result = original(store_id)
            # after 基线阶段（第二次清点 s_ok）篡改结果 → 触发回滚
            if store_id == self.s_ok and calls["n"] > 1:
                result = dict(result, orders=result["orders"] + 999)
            return result

        with mock.patch.object(self.env.server, "_store_inventory",
                               side_effect=flaky):
            st, r = self.admin.post(
                "/api/admin/migration/model-migrate", {})
        self.assertEqual(st, 500)
        self.assertIn("已回滚", r["message"])
        # 两表回到迁移前状态（缺字段仍缺, 零半成品）
        with open(os.path.join(self.env.data_dir, "stores.json"),
                  encoding="utf-8") as f:
            doc_after = json.load(f)
        self.assertEqual(doc_after, doc_before)
        events = [e for e in _read_security_log(self.env)
                  if e["event"] == "store.model_migrate"
                  and e.get("result") == "error"]
        self.assertTrue(events)
        self.assertTrue(events[-1]["detail"]["recovered"])

    def test_05_idempotent_and_clean_check(self):
        # test_04 回滚后样本回到缺字段状态, 先完成迁移再验证幂等
        st, r = self.admin.post("/api/admin/migration/model-migrate", {})
        self.assertEqual(st, 200, r)
        st, r = self.admin.get("/api/admin/migration/model-check")
        self.assertEqual(st, 200)
        data = r["data"]
        fixable = [it for it in data["items"] if it["auto_fixable"]]
        self.assertEqual(fixable, [])  # 迁移后零 auto_fixable
        # 再次迁移 → 无可迁移项, 幂等
        st, r = self.admin.post("/api/admin/migration/model-migrate", {})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["migrated"], [])
        # 歧义样本保留原样（无自动分配/自动处置）
        store = _find_store(self.env, self.s_nodir)
        self.assertIsNotNone(store)
        self.assertIn("no_store_dir",
                      next(it["issues"] for it in data["items"]
                           if it["store_id"] == self.s_nodir))

    def test_06_member_forbidden(self):
        st, r = self.env.member.get("/api/admin/migration/model-check")
        self.assertEqual(st, 403)
        st, r = self.env.member.post(
            "/api/admin/migration/model-migrate", {})
        self.assertEqual(st, 403)

    def test_07_audit_event(self):
        entries = [e for e in _read_security_log(self.env)
                   if e["event"] == "store.model_migrate"
                   and e.get("result") == "success"]
        self.assertTrue(entries)
        e = entries[-1]
        self.assertEqual(e["username"], self.env.admin_user)
        self.assertIsNotNone(e["user_id"])
        self.assertIn("backup_dir", e["detail"])
        self.assertIn("migrated", e["detail"])


_STATIC_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "static")


def _read_static(name):
    with open(os.path.join(_STATIC_DIR, name), encoding="utf-8") as f:
        return f.read()


class TestFrontendAssets(unittest.TestCase):
    """前端资产静态断言（组 6）: 生命周期入口/五要素弹窗/筛选/多店上传/导航红线

    纯静态文件断言（零服务依赖）; nav/member 视图剔除逻辑的 HTTP 层断言
    由 tests/test_page_routes 承担, 此处静态断言其结构前提。
    """

    def test_01_stores_html_lifecycle_entries_and_dialogs(self):
        html = _read_static("stores.html")
        # 创建入口 + 弹窗（名称/归属成员搜索）
        self.assertIn('id="btn-create-store"', html)
        self.assertIn("openCreateStore", html)
        self.assertIn('API.post("/api/stores"', html)
        # 停用/启用入口 + 影响说明（数据保留 + 重新登录）
        self.assertIn("changeStoreStatus", html)
        # 归档入口已移除 → 仅断言 disable/enable
        for action in ("disable", "enable"):
            self.assertIn(f", '{action}')", html)
        self.assertNotIn(", 'archive')", html)
        self.assertIn("业务数据完整保留", html)
        self.assertIn("需重新登录", html)
        # 查看数据（十项清点只读）
        self.assertIn("viewStoreData", html)
        self.assertIn("inventory", html)
        # 删除: 五要素 + 强确认 + 两步删除 + 阻止预告
        self.assertIn("deleteStoreFlow", html)
        for frag in ("店铺名称", "店铺 ID", "归属成员",
                     "confirm_store_name", "confirm_acknowledged",
                     "永久删除店铺及全部数据", "我确认永久删除该店铺及其全部关联业务数据"):
            self.assertIn(frag, html)
        self.assertNotIn("物理删除将被阻止", html)
        # 筛选扩展（店铺 ID / 店铺状态）
        self.assertIn('id="f-store-id"', html)
        self.assertIn('id="f-status"', html)

    def test_02_create_user_html_initial_store_choice(self):
        html = _read_static("create_user.html")
        for frag in ("初始店铺", "创建初始店铺", "暂不创建，稍后分配",
                     'name="init-store"', "店铺管理"):
            self.assertIn(frag, html)
        # 提交体显式携带 create_store（未经选择不默认创建）
        self.assertIn("payload.create_store", html)
        # 结果文案区分建店/未建店
        self.assertIn("已创建初始店铺", html)
        self.assertIn("未创建店铺（稍后分配）", html)

    def test_03_index_html_member_multi_store_upload(self):
        html = _read_static("index.html")
        # import_store_picker 改造后：上传统一经店铺选择弹窗（member 多店
        # 在弹窗列表中选择，uploadTask 携带目标 store_id）
        for frag in ("openStorePicker", "uploadTask", "SS.multi",
                     "target_store_id", "选择导入店铺",
                     "sp-list", "/static/modal.js"):
            self.assertIn(frag, html)
        # 旧顶部目标级联（member 多店下拉）已按需求移除
        for frag in ("memberUploadStore", "upload-member-item",
                     "upload-store-row", "upload-row-hint"):
            self.assertNotIn(frag, html)

    def test_04_store_scope_js_member_multi_render(self):
        js = _read_static("store_scope.js")
        # member 多店判定 + 渲染分支（单店徽章互斥）
        self.assertIn("this.stores.length > 1", js)
        self.assertIn("store-scope single", js)
        # 查询参数与写目标随店铺切换
        self.assertIn("store_id", js)
        self.assertIn("target_store_id", js)

    def test_05_nav_invariants_member_no_mgmt_entry(self):
        for page in ("stores.html", "create_user.html", "index.html",
                     "users.html"):
            html = _read_static(page)
            m = re.search(
                r"<!--admin-nav-start-->(.*?)<!--admin-nav-end-->",
                html, re.S)
            self.assertIsNotNone(m, f"{page} 缺少 admin 导航标记块")
            admin_block = m.group(1)
            # 管理入口全部位于 admin 块内（服务端为 member 渲染时剔除）
            self.assertIn('<a href="/stores"', admin_block, page)
            self.assertIn('<a href="/users"', admin_block, page)
            # 剔除 admin 块后 nav 恰 10 项业务菜单
            nav = re.search(r"<nav>(.*?)</nav>", html, re.S)
            self.assertIsNotNone(nav, page)
            nav_wo_admin = re.sub(
                r"<!--admin-nav-start-->.*?<!--admin-nav-end-->", "",
                nav.group(1), flags=re.S)
            links = re.findall(r'<a href="(/[a-z-]+)"', nav_wo_admin)
            self.assertEqual(len(links), 10, f"{page} nav 业务菜单数")
            self.assertNotIn("/stores", links, page)
            self.assertNotIn("/users", links, page)


class TestAuditAndSecurity(unittest.TestCase):
    """审计与安全用例组（组 7.1, spec 6.5/4.3.5/5.6.1 规则 3）

    7 类新事件五要素 / 审计内容红线 / 伪造身份无效 / 按操作者可追溯。
    """

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            "aud_a", PASSWORD, store_name="审计主店")
        cls.m1 = cls.member.get("/api/auth/me")[1]["data"]["store_id"]
        # 主店导入订单 → 有历史交易（clear-data 有量可清）
        st, r = cls.member.upload_confirm(
            "/api/import/wechat-orders", "订单.xlsx", make_wx_xlsx("AD1"))
        assert st == 200, r
        cls.operator_name = cls.env.admin_user
        cls.operator_id = cls.env.get_user_id(cls.env.admin_user)
        # 触发 7 类事件: create → disable → enable → archive
        st, r = cls.admin.post("/api/stores", {
            "name": "审计空店", "owner_user_id": cls.env.get_user_id("aud_a")})
        assert st == 200, r
        cls.s_created = r["data"]["store_id"]
        for action in ("disable", "enable"):
            st, r = cls.admin.post(
                f"/api/stores/{cls.s_created}/status", {"action": action})
            assert st == 200, r
        # 归档已下线 → archive 预期 400
        st, r = cls.admin.post(
            f"/api/stores/{cls.s_created}/status", {"action": "archive"})
        assert st == 400, r
        # clear_data + delete（主店: 先清空再删除）
        st, r = cls.admin.post(f"/api/stores/{cls.m1}/clear-data", {
            "confirm_store_name": "审计主店",
            "confirm_acknowledged": True})
        assert st == 200, r
        st, r = cls.admin.post(f"/api/stores/{cls.m1}/delete", {
            "confirm_store_name": "审计主店",
            "confirm_acknowledged": True})
        assert st == 200, r
        # model_migrate: 构造缺字段样本后迁移
        st, r = cls.admin.post("/api/stores", {
            "name": "审计缺字段店",
            "owner_user_id": cls.env.get_user_id("aud_a")})
        assert st == 200, r
        cls.s_legacy = r["data"]["store_id"]
        path = os.path.join(cls.env.data_dir, "stores.json")
        with open(path, encoding="utf-8") as f:
            doc = json.load(f)
        for s in doc["stores"]:
            if s["store_id"] == cls.s_legacy:
                for key in ("status", "updated_at", "assignment_status"):
                    s.pop(key, None)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)
        st, r = cls.admin.post("/api/admin/migration/model-migrate", {})
        assert st == 200, r

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_all_seven_events_five_elements(self):
        """6 类 success 事件 + store.archive_removed denied 事件，逐条五要素"""
        entries = _read_security_log(self.env)
        # 6 类 success 事件（store.archive success 不再存在）
        for event in ("store.create", "store.disable", "store.enable",
                      "store.delete", "store.clear_data", "store.model_migrate"):
            hits = [e for e in entries if e["event"] == event
                    and e.get("result") == "success"]
            self.assertTrue(hits, f"缺少成功事件 {event}")
            e = hits[-1]
            self.assertEqual(e["username"], self.operator_name, event)
            self.assertEqual(e["user_id"], self.operator_id, event)
            self.assertTrue(e.get("time"), event)
            self.assertIsInstance(e.get("detail"), dict, event)
            if event == "store.model_migrate":
                self.assertTrue(e["detail"].get("migrated"), event)
            else:
                self.assertIn("store_id", e["detail"], event)
        # store.archive_removed denied 事件存在且含五要素
        arch_removed = [e for e in entries if e["event"] == "store.archive_removed"
                        and e.get("result") == "denied"]
        self.assertTrue(arch_removed, "缺少 store.archive_removed denied 事件")
        e = arch_removed[-1]
        self.assertEqual(e["username"], self.operator_name)
        self.assertEqual(e["user_id"], self.operator_id)
        self.assertTrue(e.get("time"))
        self.assertIn("store_id", e["detail"])

    def test_02_no_secrets_in_audit_logs(self):
        """审计红线: 新事件日志不含密码/哈希/令牌等敏感内容"""
        entries = [e for e in _read_security_log(self.env)
                   if e["event"].startswith("store.")]
        self.assertGreaterEqual(len(entries), 8)
        for e in entries:
            blob = json.dumps(e, ensure_ascii=False).lower()
            for keyword in ("password", "passwd", '"hash"', "salt",
                            "token", "secret", "authorization", "cookie"):
                self.assertNotIn(keyword, blob, f"{e['event']} 泄漏 {keyword}")
            self.assertNotIn(PASSWORD.lower(), blob)

    def test_03_member_forged_identity_rejected(self):
        """member 会话伪造 role/user_id 请求全部新接口 → 403 + access_denied"""
        member = self.env.login_client("aud_a", PASSWORD)
        stores_before = _read_stores(self.env)
        forged_extra = {"role": "admin", "user_id": self.operator_id}
        checks = [
            ("POST", "/api/stores",
             {"name": "伪造店", "owner_user_id": self.operator_id,
              **forged_extra}),
            ("POST", f"/api/stores/{self.s_created}/status",
             {"action": "disable", **forged_extra}),
            ("POST", f"/api/stores/{self.s_created}/delete",
             {"confirm_store_name": "审计空店", **forged_extra}),
            ("POST", f"/api/stores/{self.s_created}/clear-data",
             {"confirm_store_name": "审计空店", **forged_extra}),
            ("POST", f"/api/stores/{self.s_created}/rename",
             {"name": "伪造改名", **forged_extra}),
            ("POST", f"/api/stores/{self.s_created}/transfer-ownership",
             {"target_user_id": self.operator_id, **forged_extra}),
            ("GET", f"/api/stores/{self.s_created}", None),
            ("GET", "/api/admin/migration/model-check", None),
            ("POST", "/api/admin/migration/model-migrate",
             {**forged_extra}),
        ]
        for method, path, body in checks:
            if method == "GET":
                st, r = member.get(path)
            else:
                st, r = member.post(path, body)
            self.assertEqual(st, 403, f"{method} {path} 未拒绝")
        # 零变更 + access_denied 逐次留痕（detail 含 reason/path）
        self.assertEqual(_read_stores(self.env), stores_before)
        denied = [e for e in _read_security_log(self.env)
                  if e["event"] == "access_denied"
                  and e.get("username") == "aud_a"
                  and e.get("user_id") == self.env.get_user_id("aud_a")]
        self.assertGreaterEqual(len(denied), len(checks))
        for e in denied[-len(checks):]:
            self.assertEqual(e["detail"].get("reason"), "admin_only")
            self.assertTrue(e["detail"].get("path"))

    def test_04_events_traceable_by_operator(self):
        """按操作管理员检索 → 定位其全部店铺管理事件（可追溯）"""
        entries = _read_security_log(self.env)
        mine = [e for e in entries
                if e.get("username") == self.operator_name
                and e.get("user_id") == self.operator_id
                and e["event"].startswith("store.")]
        events = {e["event"] for e in mine}
        self.assertTrue({"store.create", "store.disable", "store.enable",
                         "store.archive_removed", "store.delete", "store.clear_data",
                         "store.model_migrate"} <= events)
        for e in mine:
            self.assertTrue(e.get("time"))


class TestAcceptance12(unittest.TestCase):
    """spec 5.9.2 验收 1-12 端到端串联核对（组 7.2）

    验收 ↔ 用例映射表（既有覆盖 + 本类串联核对）:
    | 验收 | 内容 | 映射用例 |
    |------|------|----------|
    | 1 | 成员多店 | TestOneToManyOwnership / TestMemberMultiStoreScope + test_01 |
    | 2 | 唯一归属 | TestCreateStore + test_02 |
    | 3 | 管理员店铺管理全链路 | TestCreateStore/TestStoreStatus/TestStoreDelete + test_03 |
    | 4 | 成员禁止跨成员管理 | TestStoreStatus/TestStoreDelete 越权用例 + test_04 |
    | 5 | 删除 API 权限拒绝 | TestStoreDelete.test_05 + test_05 |
    | 6 | 管理员无店铺 | TestMemberCreateOptionalStore(admin 建号) + test_06 |
    | 7 | 代导入准确指定 | TestMemberMultiStoreScope 上传用例 + test_07 |
    | 8 | 成员数据范围 | TestMemberMultiStoreScope 三态 + test_08 |
    | 9 | 数据隔离 | tests/test_auth_isolation + test_09 |
    | 10 | 转移后数据保留 | TestOneToManyOwnership 转移用例 + test_10 |
    | 11 | 删除不伤历史财务 | TestStoreDelete.test_02/03 + test_11 |
    | 12 | 页面功能正常 | tests/test_page_routes + test_12 |
    """

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        # 成员 A: 两店（A1 初始店 + A2 管理员创建）, 各有订单数据
        cls.a = cls.env.create_user_client(
            "acc_a", PASSWORD, store_name="验收A一店")
        cls.a1 = cls.a.get("/api/auth/me")[1]["data"]["store_id"]
        cls.a_id = cls.env.get_user_id("acc_a")
        st, r = cls.admin.post("/api/stores", {
            "name": "验收A二店", "owner_user_id": cls.a_id})
        assert st == 200, r
        cls.a2 = r["data"]["store_id"]
        st, r = cls.a.upload_confirm(
            "/api/import/wechat-orders", "A1.xlsx", make_wx_xlsx("AC1"))
        assert st == 200, r
        # member 多店上传: URL store_id + confirm target_store_id 双携带
        st, r = cls.a.upload_confirm(
            "/api/import/wechat-orders?store_id=" + cls.a2,
            "A2.xlsx", make_wx_xlsx("AC2"), target_store_id=cls.a2)
        assert st == 200, r
        # 成员 B: 一店有数据（数据隔离对照）
        cls.b = cls.env.create_user_client(
            "acc_b", PASSWORD, store_name="验收B一店")
        cls.b1 = cls.b.get("/api/auth/me")[1]["data"]["store_id"]
        cls.b_id = cls.env.get_user_id("acc_b")
        st, r = cls.b.upload_confirm(
            "/api/import/wechat-orders", "B1.xlsx", make_wx_xlsx("BC1"))
        assert st == 200, r

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_member_multi_store(self):
        st, r = self.a.get("/api/stores")
        self.assertEqual(st, 200)
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.a1, self.a2})
        st, r = self.a.get("/api/auth/me")
        self.assertEqual(r["data"]["store_count"], 2)

    def test_02_single_ownership(self):
        by_id = {s["store_id"]: s for s in _read_stores(self.env)}
        self.assertEqual(by_id[self.a1]["owner_user_id"], self.a_id)
        self.assertEqual(by_id[self.a2]["owner_user_id"], self.a_id)
        self.assertEqual(by_id[self.b1]["owner_user_id"], self.b_id)
        for s in by_id.values():
            self.assertIsInstance(s["owner_user_id"], str)

    def test_03_admin_lifecycle_full_chain(self):
        # 牺牲成员承接 transfer/状态变更/清除/删除的会话吊销副作用, A/B 不受影响
        for name in ("acc_c", "acc_d"):
            if self.env.get_user_id(name) is None:
                self.env.create_user(name, PASSWORD)
        c_id = self.env.get_user_id("acc_c")
        d_id = self.env.get_user_id("acc_d")
        st, r = self.admin.post("/api/stores", {
            "name": "验收链路店", "owner_user_id": c_id})
        self.assertEqual(st, 200, r)
        sid = r["data"]["store_id"]
        st, r = self.admin.post(f"/api/stores/{sid}/rename",
                                {"name": "验收链路店改"})
        self.assertEqual(st, 200, r)
        st, r = self.admin.post(f"/api/stores/{sid}/transfer-ownership",
                                {"target_user_id": d_id})
        self.assertEqual(st, 200, r)
        for action, expect in (("disable", "disabled"), ("enable", "active")):
            st, r = self.admin.post(f"/api/stores/{sid}/status",
                                    {"action": action})
            self.assertEqual(st, 200, r)
            self.assertEqual(r["data"]["status"], expect)
        st, r = self.admin.post(f"/api/stores/{sid}/clear-data", {
            "confirm_store_name": "验收链路店改",
            "confirm_acknowledged": True})
        self.assertEqual(st, 200, r)
        st, r = self.admin.post(f"/api/stores/{sid}/delete", {
            "confirm_store_name": "验收链路店改",
            "confirm_acknowledged": True})
        self.assertEqual(st, 200, r)
        self.assertIsNone(_find_store(self.env, sid))
        # 全链路后两牺牲成员归属关系复原
        self.assertNotIn(sid, _find_user(self.env, "acc_c")["store_ids"])
        self.assertNotIn(sid, _find_user(self.env, "acc_d")["store_ids"])

    def test_04_member_cannot_cross_manage(self):
        stores_before = _read_stores(self.env)
        checks = [
            ("GET", f"/api/stores/{self.b1}", None),
            ("POST", f"/api/stores/{self.b1}/status", {"action": "disable"}),
            ("POST", f"/api/stores/{self.b1}/rename", {"name": "越权重命名"}),
            ("POST", f"/api/stores/{self.b1}/transfer-ownership",
             {"target_user_id": self.a_id}),
            ("POST", f"/api/stores/{self.b1}/clear-data",
             {"confirm_store_name": "验收B一店"}),
        ]
        for method, path, body in checks:
            if method == "GET":
                st, r = self.a.get(path)
            else:
                st, r = self.a.post(path, body)
            self.assertEqual(st, 403, f"{method} {path} 未拒绝")
        self.assertEqual(_read_stores(self.env), stores_before)

    def test_05_delete_api_permission_denied(self):
        st, r = self.a.post(f"/api/stores/{self.b1}/delete", {
            "confirm_store_name": "验收B一店", "role": "admin"})
        self.assertEqual(st, 403)
        self.assertIsNotNone(_find_store(self.env, self.b1))
        denied = [e for e in _read_security_log(self.env)
                  if e["event"] == "access_denied"
                  and e.get("username") == "acc_a"]
        self.assertTrue(denied)

    def test_06_admin_has_no_store(self):
        st, r = self.admin.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.assertIsNone(r["data"]["store_id"])
        self.assertEqual(r["data"]["store_count"], 0)
        st, r = self.admin.get("/api/admin/all-stores?page=1&page_size=100")
        self.assertEqual(st, 200)
        ids = {s["store_id"] for s in r["data"]["stores"]}
        self.assertTrue({self.a1, self.a2, self.b1} <= ids)

    def test_07_admin_proxy_import_to_specified_store(self):
        before_a1 = len(_read_business_records(self.env, self.a1, "orders"))
        before_a2 = len(_read_business_records(self.env, self.a2, "orders"))
        st, r = self.admin.upload_confirm(
            "/api/import/wechat-orders?store_id=" + self.a2,
            "代传.xlsx", make_wx_xlsx("PX7"), target_store_id=self.a2)
        self.assertEqual(st, 200, r)
        self.assertEqual(
            len(_read_business_records(self.env, self.a2, "orders")),
            before_a2 + WX_ROW_COUNT)
        self.assertEqual(
            len(_read_business_records(self.env, self.a1, "orders")),
            before_a1)

    def test_08_member_data_scope(self):
        # 聚合: A1(10, AC1) + A2(20, AC2+test_07 代传 PX7)
        st, r = self.a.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 30)
        prefixes = {o["order_id"][:3] for o in r["data"]["orders"]}
        self.assertEqual(prefixes, {"AC1", "AC2", "PX7"})
        # 单店: 仅 A2（test_07 代传后 20 条）
        st, r = self.a.get(f"/api/orders?store_id={self.a2}&page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 20)
        # B: 仅本店
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 10)
        self.assertTrue(all(o["order_id"].startswith("BC1")
                            for o in r["data"]["orders"]))

    def test_09_data_isolation_between_members(self):
        st, r = self.a.get("/api/stores")
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.a1, self.a2})
        st, r = self.b.get("/api/stores")
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.b1})
        st, r = self.a.get(f"/api/orders?store_id={self.b1}&page=1&size=20")
        self.assertEqual(st, 403)
        st, r = self.b.get(f"/api/orders?store_id={self.a1}&page=1&size=20")
        self.assertEqual(st, 403)

    def test_10_transfer_preserves_data(self):
        before = len(_read_business_records(self.env, self.a1, "orders"))
        st, r = self.admin.post(f"/api/stores/{self.a1}/transfer-ownership", {
            "target_user_id": self.b_id})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["data_preserved"], True)
        after = len(_read_business_records(self.env, self.a1, "orders"))
        self.assertEqual(after, before)
        store = _find_store(self.env, self.a1)
        self.assertEqual(store["owner_user_id"], self.b_id)
        self.assertEqual(_find_user(self.env, "acc_a")["store_ids"],
                         [self.a2])
        self.assertIn(self.a1, _find_user(self.env, "acc_b")["store_ids"])

    def test_11_delete_not_hurt_historical_finance(self):
        # 有数据店铺直接删除成功（归档下线 + 物理删除新契约）
        before = len(_read_business_records(self.env, self.a2, "orders"))
        st, r = self.admin.post(f"/api/stores/{self.a2}/delete", {
            "confirm_store_name": "验收A二店",
            "confirm_acknowledged": True})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["deleted_counts"]["orders"], before)
        # 删除后店铺不存在
        self.assertIsNone(_find_store(self.env, self.a2))
        self.assertFalse(_store_dir_exists(self.env, self.a2))
        # 归档已下线 → archive 请求预期 400（归档下线校验先于店铺存在性）
        st, r = self.admin.post(f"/api/stores/{self.a2}/status", {
            "action": "archive"})
        self.assertEqual(st, 400)

    def test_12_pages_functional_dual_identity(self):
        pages = ["/dashboard", "/data-import", "/orders", "/refund-orders",
                 "/supplier-costs", "/reconciliation",
                 "/profit", "/anomalies", "/logs"]
        for path in pages + ["/users", "/stores", "/users/create"]:
            st, body, _ = self.admin.get_page(path)
            self.assertEqual(st, 200, path)
            self.assertNotIn("服务器错误", body, path)
        # B（转移后两店）member 视角: 业务页正常且无管理入口;
        # 异常处理/操作日志仅 admin 可见, member 访问被拒
        member_pages = [p for p in pages if p not in ("/anomalies", "/logs")]
        for path in member_pages:
            st, body, headers = self.b.get_page(path)
            self.assertEqual(st, 200, path)
            self.assertNotIn("admin-nav-start", body, path)
            self.assertNotIn("admin-only-nav-start", body, path)
            self.assertNotIn('href="/stores"', body, path)
        for path in ("/anomalies", "/logs"):
            st, _, _ = self.b.get_page(path)
            self.assertNotEqual(st, 200, f"{path} member 应被拒绝")
        st, r = self.b.get("/api/auth/me")
        self.assertEqual(r["data"]["store_count"], 2)
        st, r = self.b.get(f"/api/summary?month={MONTH}")
        self.assertEqual(st, 200)
        st, r = self.b.get(f"/api/profit?page=1&size=20")
        self.assertEqual(st, 200)


class TestOwnershipTransferEnhancements(unittest.TestCase):
    """修改店铺归属者增强验收: 幂等提示 / 失败审计 / 成功审计 / 权限翻转 / 导入兼容 / 持久化

    需求映射:
    | 验收 | 内容 | 用例 |
    |------|------|------|
    | 幂等 | 目标=当前归属 → "归属成员未发生变化", 零写入零副作用 | test_01 |
    | 成功 | 审计含操作人角色/店铺名/原新归属 ID; A 失权 B 得权 | test_02 |
    | 失败 | 转给 admin/停用/不存在成员/不存在店铺/缺参 → 4xx + 审计含原因 | test_03 |
    | 导入 | 转移后 B 可导入、A 被拒、admin 代传正常、历史批次不变 | test_04 |
    | 持久 | 服务重启后归属关系仍正确 | test_05 |
    | 前端 | 弹窗三要素/仅启用普通成员/无成员提示/幂等文案 | test_06 |
    """

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        # 成员 A: 两店（A1 初始店 + A2 管理员创建）, A1 有订单数据
        cls.a = cls.env.create_user_client(
            "tra_a", PASSWORD, store_name="归属A一店")
        cls.a1 = cls.a.get("/api/auth/me")[1]["data"]["store_id"]
        cls.a_id = cls.env.get_user_id("tra_a")
        st, r = cls.admin.post("/api/stores", {
            "name": "归属A二店", "owner_user_id": cls.a_id})
        assert st == 200, r
        cls.a2 = r["data"]["store_id"]
        st, r = cls.a.upload_confirm(
            "/api/import/wechat-orders", "TRA1.xlsx", make_wx_xlsx("TRA1"))
        assert st == 200, r
        st, r = cls.a.upload_confirm(
            "/api/import/wechat-orders?store_id=" + cls.a2,
            "TRA2.xlsx", make_wx_xlsx("TRA2"), target_store_id=cls.a2)
        assert st == 200, r
        # 成员 B: 一店（转移目标, "其他店铺不受影响"对照）
        cls.b = cls.env.create_user_client(
            "tra_b", PASSWORD, store_name="归属B一店")
        cls.b1 = cls.b.get("/api/auth/me")[1]["data"]["store_id"]
        cls.b_id = cls.env.get_user_id("tra_b")
        # 成员 C: 停用成员（失败场景）
        cls.c = cls.env.create_user_client(
            "tra_c", PASSWORD, store_name="归属C店")
        cls.c_id = cls.env.get_user_id("tra_c")
        st, r = cls.admin.post(f"/api/users/{cls.c_id}/status",
                               {"action": "disable"})
        assert st == 200, r

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 辅助 ----------
    def _ownership_events(self):
        return [e for e in _read_security_log(self.env)
                if e.get("event") == "store.ownership_change"]

    def _migration_dirs(self):
        root = self.env.data_dir
        return sorted(d for d in os.listdir(root)
                      if d.startswith("ownership_migration_"))

    # ---------- 用例 ----------
    def test_01_idempotent_unchanged_no_side_effects(self):
        """目标=当前归属 → 200 + unchanged + 提示语; 不备份/不写盘/不吊销会话"""
        store_before = _find_store(self.env, self.a1)
        user_before = _find_user(self.env, "tra_a")
        events_before = len(self._ownership_events())
        self.assertEqual(self._migration_dirs(), [])
        st, r = self.admin.post(f"/api/stores/{self.a1}/transfer-ownership",
                                {"target_user_id": self.a_id})
        self.assertEqual(st, 200, r)
        self.assertEqual(r["message"], "归属成员未发生变化")
        data = r["data"]
        self.assertIs(data["unchanged"], True)
        self.assertEqual(data["store_id"], self.a1)
        self.assertEqual(data["previous_owner"]["user_id"], self.a_id)
        self.assertEqual(data["new_owner"]["user_id"], self.a_id)
        # 零写入: 两表记录（含 updated_at/store_ids）不变, 无迁移备份目录
        self.assertEqual(_find_store(self.env, self.a1), store_before)
        self.assertEqual(_find_user(self.env, "tra_a"), user_before)
        self.assertEqual(self._migration_dirs(), [])
        # 当前归属成员会话未被吊销
        st, me = self.a.get("/api/auth/me")
        self.assertEqual(st, 200)
        self.assertEqual(me["data"]["username"], "tra_a")
        # 审计: 恰一条 noop 记录, 标记 unchanged 与原因
        events = self._ownership_events()[events_before:]
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0]["result"], "noop")
        d = events[0]["detail"]
        self.assertIs(d["unchanged"], True)
        self.assertEqual(d["reason"], "归属成员未发生变化")
        self.assertEqual(d["store_name"], "归属A一店")
        self.assertEqual(d["operator_role"], "admin")

    def test_02_transfer_success_audit_and_permission_flip(self):
        """A1 从 A 转给 B: 数据保留 + 审计字段齐全 + A 失权/B 得权/其余店不受影响"""
        orders_before = _read_business_records(self.env, self.a1, "orders")
        self.assertGreater(len(orders_before), 0)
        a2_before = _find_store(self.env, self.a2)
        b1_before = _find_store(self.env, self.b1)
        st, r = self.admin.post(f"/api/stores/{self.a1}/transfer-ownership",
                                {"target_user_id": self.b_id})
        self.assertEqual(st, 200, r)
        data = r["data"]
        self.assertIs(data["unchanged"], False)
        self.assertEqual(data["store_id"], self.a1)
        self.assertEqual(data["store_name"], "归属A一店")
        self.assertEqual(data["previous_owner"]["user_id"], self.a_id)
        self.assertEqual(data["new_owner"]["user_id"], self.b_id)
        self.assertIs(data["data_preserved"], True)
        # 订单数据完整保留（记录级字节不变）
        self.assertEqual(
            _read_business_records(self.env, self.a1, "orders"), orders_before)
        # 归属两表同步更新, 店铺 ID 不变
        store = _find_store(self.env, self.a1)
        self.assertEqual(store["owner_user_id"], self.b_id)
        self.assertNotIn(self.a1, _find_user(self.env, "tra_a")["store_ids"])
        self.assertIn(self.a1, _find_user(self.env, "tra_b")["store_ids"])
        # A 的其他店铺（A2）与 B 的其他店铺（B1）不受影响
        self.assertEqual(_find_store(self.env, self.a2), a2_before)
        self.assertEqual(_find_store(self.env, self.b1), b1_before)
        # 审计字段齐全: 操作人/角色/店铺/原新归属/记录数/备份/吊销数
        ev = [e for e in self._ownership_events()
              if e["result"] == "success"][-1]
        self.assertEqual(ev["username"], "admin")
        self.assertEqual(ev["user_id"], self.env.get_user_id("admin"))
        self.assertTrue(ev["time"])
        d = ev["detail"]
        self.assertEqual(d["store_id"], self.a1)
        self.assertEqual(d["store_name"], "归属A一店")
        self.assertEqual(d["operator_role"], "admin")
        self.assertEqual(d["target_user_id"], self.b_id)
        self.assertEqual(d["previous_owner"]["user_id"], self.a_id)
        self.assertEqual(d["new_owner"]["user_id"], self.b_id)
        self.assertGreaterEqual(d["record_counts"]["orders"], 1)
        self.assertGreaterEqual(d["revoked_sessions"], 1)
        self.assertTrue(d["backup_dir"])
        # 权限翻转: A 旧会话被吊销(401), 重登后访问 A1 → 403; A2 仍可访问
        st, _ = self.a.get("/api/auth/me")
        self.assertEqual(st, 401)
        a_relogin = self.env.login_client("tra_a", PASSWORD)
        st, _ = a_relogin.get(f"/api/orders?store_id={self.a1}&page=1&size=20")
        self.assertEqual(st, 403)
        st, r = a_relogin.get(f"/api/orders?store_id={self.a2}&page=1&size=20")
        self.assertEqual(st, 200)
        # B 得权: 可见 A1 原有业务数据; B1 数据不受影响
        st, r = self.b.get(f"/api/orders?store_id={self.a1}&page=1&size=20")
        self.assertEqual(st, 200)
        self.assertGreater(r["data"]["total"], 0)
        st, r = self.b.get("/api/stores")
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.a1, self.b1})

    def test_03_failures_rejected_and_audited(self):
        """转给 admin/停用成员/不存在成员/不存在店铺/缺参 → 4xx, 零变更, 审计含失败原因"""
        admin_id = self.env.get_user_id(self.env.admin_user)
        events_before = len(self._ownership_events())
        stores_before = _read_stores(self.env)
        dirs_before = self._migration_dirs()
        cases = [
            (f"/api/stores/{self.a1}/transfer-ownership",
             {"target_user_id": admin_id}, 400, "管理员账号"),
            (f"/api/stores/{self.a1}/transfer-ownership",
             {"target_user_id": self.c_id}, 400, "已停用"),
            (f"/api/stores/{self.a1}/transfer-ownership",
             {"target_user_id": "u_nope999"}, 404, "目标用户不存在"),
            ("/api/stores/s_nope999/transfer-ownership",
             {"target_user_id": self.b_id}, 404, "店铺不存在"),
            (f"/api/stores/{self.a1}/transfer-ownership",
             {}, 400, "请提供目标成员"),
        ]
        for path, body, want_st, want_reason in cases:
            st, r = self.admin.post(path, body)
            self.assertEqual(st, want_st, f"{path} {body}: {r}")
            self.assertIn(want_reason, r["message"])
        # 失败零变更: 店铺两表不变, 无新增迁移备份目录
        self.assertEqual(_read_stores(self.env), stores_before)
        self.assertEqual(self._migration_dirs(), dirs_before)
        # 审计: 每次失败一条 error 记录, 含失败原因/操作人角色
        events = self._ownership_events()[events_before:]
        self.assertEqual(len(events), len(cases))
        self.assertTrue(all(e["result"] == "error" for e in events))
        for e in events:
            self.assertEqual(e["detail"]["operator_role"], "admin")
            self.assertTrue(e["detail"].get("reason"))
        reasons = " / ".join(e["detail"]["reason"] for e in events)
        for frag in ("管理员账号", "已停用", "不存在", "请提供目标成员"):
            self.assertIn(frag, reasons)

    def test_04_import_flow_after_transfer(self):
        """转移后: B 可向转入店导入, A 被拒, admin 代传正常, 历史批次记录不变"""
        batches_before = _read_business_records(self.env, self.a1,
                                                 "import_records")
        self.assertGreater(len(batches_before), 0)
        st, r = self.b.upload_confirm(
            "/api/import/wechat-orders", "TRB.xlsx", make_wx_xlsx("TRB"),
            target_store_id=self.a1)
        self.assertEqual(st, 200, r)
        a_relogin = self.env.login_client("tra_a", PASSWORD)
        st, r = a_relogin.upload_confirm(
            "/api/import/wechat-orders", "TRA3.xlsx", make_wx_xlsx("TRA3"),
            target_store_id=self.a1)
        self.assertEqual(st, 403, r)
        st, r = self.admin.upload_confirm(
            "/api/import/wechat-orders", "TRADM.xlsx", make_wx_xlsx("TRADM"),
            target_store_id=self.a1)
        self.assertEqual(st, 200, r)
        # 历史导入批次记录（含原始操作人身份）逐条不变, 新批次仅追加
        batches_after = _read_business_records(self.env, self.a1,
                                                "import_records")
        self.assertEqual(batches_after[:len(batches_before)], batches_before)
        self.assertEqual(len(batches_after), len(batches_before) + 2)

    def test_05_ownership_survives_restart(self):
        """JSON 写盘后重启服务: 归属关系仍为 B, 成员视角一致"""
        self.env.restart()
        admin = self.env.login_client(self.env.admin_user,
                                      self.env.admin_password)
        st, r = admin.get(f"/api/stores/{self.a1}")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["owner_user_id"], self.b_id)
        self.assertEqual(r["data"]["owner_username"], "tra_b")
        b = self.env.login_client("tra_b", PASSWORD)
        st, r = b.get("/api/stores")
        self.assertEqual({s["store_id"] for s in r["data"]["stores"]},
                         {self.a1, self.b1})
        st, r = b.get(f"/api/orders?store_id={self.a1}&page=1&size=20")
        self.assertEqual(st, 200)
        self.assertGreater(r["data"]["total"], 0)

    def test_06_frontend_transfer_dialog_assets(self):
        """前端静态断言: 专用归属弹窗已移除，归属修改经"编辑"弹窗完成"""
        html = _read_static("stores.html")
        # 专用归属弹窗的标题与入口函数均已删除
        for removed in ("修改店铺归属", "openTransferOwnership",
                        "confirmTransferOwnership", "showTransferResult"):
            self.assertNotIn(removed, html)
        # 归属修改改由 editStore 弹窗承担：标题 + 归属者下拉框
        self.assertIn("编辑店铺", html)
        self.assertIn('id="rn-owner"', html)
        self.assertIn("店铺归属者", html)
        # 成员列表仅显示启用状态的普通成员（过滤管理员与停用成员）
        self.assertIn('u.role === "member" && u.status === "active"', html)
        # 无可用成员提示
        self.assertIn("系统中没有可用的普通成员，请先创建成员账号", html)
        # 保存时调用归属转移接口
        self.assertIn("/transfer-ownership", html)


if __name__ == "__main__":
    unittest.main()

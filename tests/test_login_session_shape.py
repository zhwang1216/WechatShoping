"""登录时 sessions/users/stores 的 list 与 dict 混用。

复现生产故障：sessions.json 的 sessions 字段是列表，登录在
SessionStore.create 里对它调用 .items()，接口返回
'list' object has no attribute 'items'。
"""
import json
import os
import unittest

from tests.http_client import (
    ADMIN_PASSWORD, ADMIN_USER, Client, TestEnv, USER_A, USER_A_PASSWORD,
)


def _read_bytes(path):
    with open(path, "rb") as f:
        return f.read()


def _read(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _write(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _identities(users_doc):
    """不含 last_login 的账号身份，用来确认迁移和登录没有改密码或删账号。"""
    rows = []
    for user in users_doc.get("users", []):
        rows.append({
            "user_id": user.get("user_id"),
            "username": user.get("username"),
            "password_hash": user.get("password_hash"),
            "role": user.get("role"),
            "status": user.get("status"),
            "store_ids": list(user.get("store_ids") or []),
        })
    return rows


class LoginSessionShapeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.data = cls.env.data_dir
        cls.users_path = os.path.join(cls.data, "users.json")
        cls.stores_path = os.path.join(cls.data, "stores.json")
        cls.sessions_path = os.path.join(cls.data, "sessions.json")
        cls.admin = cls.env.admin
        cls.member = cls.env.member

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def setUp(self):
        self._users_bytes = _read_bytes(self.users_path)
        self._stores_bytes = _read_bytes(self.stores_path)
        self._sessions_bytes = _read_bytes(self.sessions_path)

    def tearDown(self):
        for path, blob in (
                (self.users_path, self._users_bytes),
                (self.stores_path, self._stores_bytes),
                (self.sessions_path, self._sessions_bytes)):
            with open(path, "wb") as f:
                f.write(blob)

    def _plant_sessions_list(self, items):
        _write(self.sessions_path, {
            "version": 1,
            "updated_at": "2026-11-07T21:17:22",
            "sessions": items,
        })

    def test_admin_and_member_login_when_sessions_is_list(self):
        identities = _identities(_read(self.users_path))
        store_ids = [s["store_id"] for s in _read(self.stores_path)["stores"]]
        self._plant_sessions_list([])

        admin = Client(self.env)
        admin_info = admin.login(ADMIN_USER, ADMIN_PASSWORD)
        self.assertEqual(admin_info["role"], "admin")
        self.assertIsNone(admin_info["store_id"])

        member = Client(self.env)
        member_info = member.login("biz_carrier", "Carrier123456")
        self.assertEqual(member_info["role"], "member")
        self.assertIn(member_info["store_id"], store_ids)

        sessions = _read(self.sessions_path)
        self.assertIsInstance(sessions["sessions"], dict)
        self.assertGreaterEqual(len(sessions["sessions"]), 1)
        backups = os.path.join(self.data, "backups")
        self.assertTrue(any(
            name.startswith("sessions.json.pre_migrate.")
            for name in os.listdir(backups)))

        after = {row["user_id"]: row for row in _identities(_read(self.users_path))}
        for row in identities:
            got = after[row["user_id"]]
            self.assertEqual(got["password_hash"], row["password_hash"])
            self.assertEqual(got["role"], row["role"])
            self.assertEqual(got["store_ids"], row["store_ids"])
            self.assertEqual(got["username"], row["username"])
        after_store_ids = [s["store_id"] for s in _read(self.stores_path)["stores"]]
        self.assertEqual(after_store_ids, store_ids)

    def test_wrong_password_does_not_500_or_rewrite_sessions_list(self):
        self._plant_sessions_list([])
        before = _read_bytes(self.sessions_path)
        status, resp = Client(self.env).post("/api/auth/login", {
            "username": ADMIN_USER, "password": "WrongPass123",
        }, csrf=False)
        self.assertEqual(status, 401)
        self.assertEqual(resp["message"], "账号或密码错误")
        self.assertNotIn("list", resp["message"])
        self.assertNotIn("服务器内部错误", resp["message"])
        self.assertEqual(_read_bytes(self.sessions_path), before)
        self.assertEqual(_read_bytes(self.users_path), self._users_bytes)

    def test_legacy_session_records_in_list_are_kept(self):
        users = _read(self.users_path)["users"]
        admin = next(u for u in users if u["username"] == ADMIN_USER)
        self._plant_sessions_list([{
            "token": "legacy-token-keep",
            "user_id": admin["user_id"],
            "username": ADMIN_USER,
            "store_id": None,
            "password_version": admin.get("password_version", 1),
            "created_at": "2026-11-08T00:00:00",
            "expires_at": "2099-01-01T00:00:00",
        }])
        Client(self.env).login(ADMIN_USER, ADMIN_PASSWORD)
        sessions = _read(self.sessions_path)["sessions"]
        self.assertIsInstance(sessions, dict)
        self.assertIn("legacy-token-keep", sessions)
        self.assertEqual(sessions["legacy-token-keep"]["user_id"], admin["user_id"])

    def test_users_toplevel_list_is_readable_and_preserved(self):
        original = _read(self.users_path)
        identities = _identities(original)
        _write(self.users_path, original["users"])
        client = Client(self.env)
        info = client.login(ADMIN_USER, ADMIN_PASSWORD)
        self.assertEqual(info["username"], ADMIN_USER)
        migrated = _read(self.users_path)
        self.assertIsInstance(migrated["users"], list)
        self.assertEqual(_identities(migrated), [
            {**row} for row in identities
        ])
        after = {row["username"]: row for row in _identities(migrated)}
        for row in identities:
            self.assertEqual(after[row["username"]]["password_hash"],
                             row["password_hash"])
        status, me = client.get("/api/auth/me")
        self.assertEqual(status, 200)
        self.assertEqual(me["data"]["role"], "admin")

    def test_stores_toplevel_list_is_readable(self):
        original = _read(self.stores_path)
        store_ids = [s["store_id"] for s in original["stores"]]
        _write(self.stores_path, original["stores"])
        member = Client(self.env)
        info = member.login("biz_carrier", "Carrier123456")
        self.assertIn(info["store_id"], store_ids)
        migrated = _read(self.stores_path)
        self.assertIsInstance(migrated["stores"], list)
        self.assertEqual(
            [s["store_id"] for s in migrated["stores"]], store_ids)
        self.assertEqual(
            [s["owner_user_id"] for s in migrated["stores"]],
            [s["owner_user_id"] for s in original["stores"]])

    def test_member_without_store_is_explicit_error(self):
        status, resp = self.admin.post("/api/users", {
            "username": "nostore",
            "password": "Nostore123",
            "role": "member",
            "create_store": False,
        })
        self.assertEqual(status, 200, resp)
        status, resp = Client(self.env).post("/api/auth/login", {
            "username": "nostore", "password": "Nostore123",
        }, csrf=False)
        self.assertEqual(status, 401)
        self.assertEqual(resp["message"], "账号未绑定店铺，请联系管理员")
        self.assertNotIn("服务器内部错误", resp["message"])

    def test_multiple_stores_and_admin_sees_all(self):
        status, me = self.member.get("/api/auth/me")
        self.assertEqual(status, 200)
        owner_id = me["data"]["user_id"]
        status, created = self.admin.post("/api/stores", {
            "name": "第二家店铺", "owner_user_id": owner_id,
        })
        self.assertEqual(status, 200, created)
        second = created["data"]["store_id"]
        member = Client(self.env)
        member.login("biz_carrier", "Carrier123456")
        status, owned = member.get("/api/stores")
        self.assertEqual(status, 200, owned)
        owned_ids = [s["store_id"] for s in owned["data"]["stores"]]
        self.assertIn(me["data"]["store_id"], owned_ids)
        self.assertIn(second, owned_ids)
        self.assertGreaterEqual(len(owned_ids), 2)
        status, me2 = member.get("/api/auth/me")
        self.assertEqual(status, 200)
        self.assertGreaterEqual(me2["data"]["store_count"], 2)

        status, all_stores = self.admin.get("/api/admin/all-stores")
        self.assertEqual(status, 200, all_stores)
        all_ids = [s["store_id"] for s in all_stores["data"]["stores"]]
        for store_id in owned_ids:
            self.assertIn(store_id, all_ids)

        status, page, _headers = member.get_page("/")
        self.assertEqual(status, 200)
        self.assertIn("app-sidebar", page)
        status, page, headers = member.get_page("/login")
        self.assertEqual(status, 302)
        self.assertEqual(headers.get("Location"), "/")

    def test_refresh_and_restart_keep_login(self):
        self._plant_sessions_list([])
        member = Client(self.env)
        member.login("biz_carrier", "Carrier123456")
        status, me = member.get("/api/auth/me")
        self.assertEqual(status, 200)
        self.assertEqual(me["data"]["username"], "biz_carrier")
        self.env.restart()
        status, me = member.get("/api/auth/me")
        self.assertEqual(status, 200)
        self.assertEqual(me["data"]["username"], "biz_carrier")
        status, resp = member.post("/api/auth/logout")
        self.assertEqual(status, 200, resp)
        status, me = member.get("/api/auth/me")
        self.assertEqual(status, 401)

    def test_unrecognized_sessions_list_is_not_wiped(self):
        self._plant_sessions_list([1, "not-a-session"])
        before = _read_bytes(self.sessions_path)
        status, resp = Client(self.env).post("/api/auth/login", {
            "username": ADMIN_USER, "password": ADMIN_PASSWORD,
        }, csrf=False)
        self.assertEqual(status, 500)
        self.assertIn("已保留原文件", resp["message"])
        self.assertNotIn("password", resp["message"].lower())
        self.assertEqual(_read_bytes(self.sessions_path), before)
        log_path = os.path.join(self.env.log_dir, "security_log.jsonl")
        with open(log_path, encoding="utf-8") as f:
            entries = [json.loads(line) for line in f if line.strip()]
        logged = [e for e in entries if e.get("event") == "api.error"]
        self.assertTrue(logged)
        detail = logged[-1]["detail"]
        self.assertEqual(detail["path"], "/api/auth/login")
        self.assertEqual(detail["data_file"], "sessions.json")
        self.assertEqual(detail["value_type"], "int")
        self.assertEqual(detail["file"], "auth.py")
        self.assertTrue(detail["function"])
        self.assertIn("Traceback", detail["traceback"])
        self.assertNotIn(ADMIN_PASSWORD, detail["traceback"])

    def test_logout_user_admin_and_register_still_work(self):
        status, users = self.admin.get("/api/users")
        self.assertEqual(status, 200)
        names = [u["username"] for u in users["data"]["users"]]
        self.assertIn(ADMIN_USER, names)
        self.assertIn("biz_carrier", names)
        status, resp = self.env.register_user(
            USER_A, USER_A_PASSWORD, store_name="注册店")
        self.assertEqual(status, 200, resp)
        registered = Client(self.env)
        info = registered.login(USER_A, USER_A_PASSWORD)
        self.assertEqual(info["role"], "member")
        self.assertTrue(info["store_id"])
        status, resp = registered.post("/api/auth/logout")
        self.assertEqual(status, 200, resp)


if __name__ == "__main__":
    unittest.main()

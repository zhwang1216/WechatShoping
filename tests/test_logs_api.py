"""操作日志查询接口测试（/api/logs，spec oplog_query_fix 第 4~5 节验收）

覆盖（11 项验收 + 辅助用例）：
1. 多条记录正常显示（五字段完整）
2. 空结果正常空状态
3. 相同操作人/店铺多记录无丢失
4. dict 形式 operator 脏数据不触发集合错误且不外泄敏感字段
5. 月份筛选（含 from/to 优先级、左闭右开）
6. 成员/店铺筛选（admin 聚合、member 收窄、store_id 单店）
7. 动作/操作人筛选（含 operator=admin 命中 dict 脏记录）
8. 分页/总数一致性与参数校验
9. 普通成员越权 403
10. 查询不改写磁盘历史日志（逐字节一致）
11. 共享查询函数页面无回归（unhashable 不复现）

全部基于 TestEnv 临时数据根；脏数据经直接追加 JSONL 行注入（模拟磁盘
历史，append 归一化不影响注入形态）。
"""
import io
import json
import os
import unittest
from contextlib import redirect_stderr
from unittest import mock

from tests.http_client import TestEnv

MARKER = "oplogfix"


def _line(operator, action, time, result="success", detail=None):
    entry = {
        "time": time,
        "operator": operator,
        "action": action,
        "result": result,
        "detail": detail if detail is not None else {"marker": MARKER},
    }
    if detail is None:
        entry["detail"] = {"marker": MARKER}
    return json.dumps(entry, ensure_ascii=False)


def _dirty_user_dict():
    """模拟旧版 store.rename 写入的完整用户 dict（含敏感字段）"""
    return {
        "user_id": "u_legacy",
        "username": "admin",
        "password_hash": "pbkdf2$sha256$fake$salt",
        "role": "admin",
    }


class TestLogsApi(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.member
        cls.member_store = cls.env.member_store
        cls.other = cls.env.create_user_client("otherguy", "Other123456")
        cls.other_user_id = cls.env.get_user_id("otherguy")
        _, me = cls.other.get("/api/auth/me")
        cls.other_store = me["data"]["store_id"]
        cls.log_file = os.path.join(
            cls.env.data_dir, "stores", cls.member_store,
            "logs", "operation_log.jsonl")
        cls._inject()
        with open(cls.log_file, "rb") as f:
            cls.baseline = f.read()

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    @classmethod
    def _inject(cls):
        """注入历史数据：1 条 dict operator 脏数据 + 正常记录若干"""
        lines = [
            # 脏数据：operator 为完整用户 dict（根因复现样本）
            _line(_dirty_user_dict(), "store.rename", "2026-10-21T23:16:17",
                  detail={"marker": MARKER, "note": "历史脏数据"}),
            # 同 username 字符串形态（双形态应计为同一操作人）
            _line("admin", "store.rename", "2026-10-20T09:00:00"),
            # 25 条分页测试记录（时间互异，验证倒序与翻页）
            *[_line("admin", "log.page_test",
                    f"2026-10-01T08:{i:02d}:00",
                    detail={"marker": MARKER, "seq": i})
              for i in range(25)],
            # 1 条 error 结果记录（statistics.error_count）
            _line("admin", "data.error", "2026-10-18T00:00:00",
                  result="error"),
            # 未来记录（空状态/月份排除）
            _line("admin", "future.entry", "2099-01-01T00:00:00"),
            # 时间格式异常记录（月份筛选下安全排除，不崩整表）
            _line("admin", "bad.time", "not-a-date"),
        ]
        with open(cls.log_file, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")

    def tearDown(self):
        # 项 10：全部查询操作不得改写磁盘日志（逐字节一致）
        with open(self.log_file, "rb") as f:
            self.assertEqual(f.read(), self.baseline)

    def _get(self, client, query):
        status, resp = client.get(f"/api/logs?{query}")
        return status, resp

    # ---------- 项 1：多条记录正常显示 ----------
    def test_list_multi_entries(self):
        status, resp = self._get(
            self.admin, f"page=1&page_size=50&month=2026-10&keyword={MARKER}")
        self.assertEqual(status, 200)
        data = resp["data"]
        self.assertEqual(data["total"], 28)
        self.assertGreaterEqual(len(data["entries"]), 28)
        for e in data["entries"]:
            for field in ("time", "action", "operator", "result", "detail"):
                self.assertIn(field, e)
            self.assertEqual(e["store_id"], self.member_store)
            self.assertTrue(e["store_name"])

    # ---------- 项 2：空结果正常空状态 ----------
    def test_empty_scope(self):
        status, resp = self._get(
            self.admin, f"page=1&page_size=50&month=2026-09&keyword={MARKER}")
        self.assertEqual(status, 200)
        data = resp["data"]
        self.assertEqual(data["total"], 0)
        self.assertEqual(data["entries"], [])
        stats = data["statistics"]
        self.assertEqual(stats["entry_count"], 0)
        self.assertEqual(stats["operator_count"], 0)
        self.assertEqual(stats["error_count"], 0)

    # ---------- 项 3：相同操作人/店铺多记录 ----------
    def test_same_operator_multiple(self):
        status, resp = self._get(
            self.admin,
            f"page=1&page_size=50&month=2026-10&keyword={MARKER}"
            f"&operator=admin")
        self.assertEqual(status, 200)
        data = resp["data"]
        self.assertEqual(data["total"], 28)
        self.assertEqual(len(data["entries"]), 28)
        self.assertTrue(all(e["operator"] == "admin" for e in data["entries"]))
        self.assertEqual(data["statistics"]["operator_count"], 1)

    # ---------- 项 4：dict 脏数据不触发集合错误、不外泄敏感字段 ----------
    def test_dirty_operator_store(self):
        status, resp = self._get(
            self.admin, f"page=1&page_size=50&keyword={MARKER}")
        self.assertEqual(status, 200)
        body_text = json.dumps(resp, ensure_ascii=False)
        self.assertNotIn("password_hash", body_text)
        self.assertNotIn("user_id", body_text)
        dirty = [e for e in resp["data"]["entries"]
                 if e.get("action") == "store.rename"
                 and e.get("time") == "2026-10-21T23:16:17"]
        self.assertEqual(len(dirty), 1)
        self.assertEqual(dirty[0]["operator"], "admin")
        stats = resp["data"]["statistics"]
        self.assertEqual(stats["operator_count"], 1)
        self.assertEqual(stats["error_count"], 1)
        self.assertEqual(resp["data"]["total"], 30)

    # ---------- 项 5：月份筛选与 from/to 优先级 ----------
    def test_month_filter(self):
        status, resp = self._get(
            self.admin, f"page=1&page_size=50&month=2026-10&keyword={MARKER}")
        self.assertEqual(status, 200)
        for e in resp["data"]["entries"]:
            self.assertTrue(e["time"].startswith("2026-10"))
        # from/to 与 month 同传 → from/to 优先（区间收窄至 [10-20, 10-22)）
        status, resp = self._get(
            self.admin,
            f"page=1&page_size=50&month=2026-10&from=2026-10-20"
            f"&to=2026-10-22&keyword={MARKER}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["total"], 2)
        times = sorted(e["time"] for e in resp["data"]["entries"])
        self.assertEqual(times, ["2026-10-20T09:00:00", "2026-10-21T23:16:17"])

    def test_bad_time_record_safe_with_month(self):
        # 时间格式异常记录：带月份筛选安全排除，不崩整表
        status, resp = self._get(
            self.admin,
            f"page=1&page_size=50&month=2026-10&keyword={MARKER}")
        self.assertEqual(status, 200)
        actions = {e["action"] for e in resp["data"]["entries"]}
        self.assertNotIn("bad.time", actions)
        self.assertNotIn("future.entry", actions)

    # ---------- 项 6：成员与店铺筛选 ----------
    def test_member_store_filter(self):
        # admin 无参 → 全部成员店铺聚合、全局倒序
        status, resp = self._get(
            self.admin, f"page=1&page_size=100&keyword={MARKER}")
        self.assertEqual(status, 200)
        data = resp["data"]
        self.assertEqual(data.get("store_scope"), "all")
        self.assertEqual(data.get("store_count"), 2)
        self.assertEqual(data["total"], 30)
        times = [e.get("time") or "" for e in data["entries"]]
        self.assertEqual(times, sorted(times, reverse=True))
        # member=<user_id> → 仅该成员店铺
        status, resp = self._get(
            self.admin,
            f"page=1&page_size=50&member={self.other_user_id}"
            f"&keyword={MARKER}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["total"], 0)
        # store_id=<某店> → 仅该店记录
        status, resp = self._get(
            self.admin,
            f"page=1&page_size=50&store_id={self.member_store}"
            f"&keyword={MARKER}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["total"], 30)
        self.assertTrue(all(e["store_id"] == self.member_store
                            for e in resp["data"]["entries"]))

    def test_member_default_scope_own_store(self):
        # 普通成员不带参数 → 仅本店日志（单店，无 store_scope 字段）
        status, resp = self._get(self.member, f"page=1&page_size=50&keyword={MARKER}")
        self.assertEqual(status, 200)
        data = resp["data"]
        self.assertNotIn("store_scope", data)
        self.assertEqual(data["total"], 30)
        self.assertTrue(all(e["store_id"] == self.member_store
                            for e in data["entries"]))

    # ---------- 项 7：动作与操作人筛选 ----------
    def test_action_operator_filter(self):
        status, resp = self._get(
            self.admin,
            "page=1&page_size=50&month=2026-10&action=store.rename")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["total"], 2)
        # operator=admin 命中 dict 脏记录（回归：归一化先于筛选）
        status, resp = self._get(
            self.admin,
            "page=1&page_size=50&month=2026-10&action=store.rename"
            "&operator=admin")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["total"], 2)
        self.assertIn("2026-10-21T23:16:17",
                      [e["time"] for e in resp["data"]["entries"]])
        # 组合交集：无命中操作人
        status, resp = self._get(
            self.admin,
            "page=1&page_size=50&month=2026-10&action=store.rename"
            "&operator=nobody")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["total"], 0)

    # ---------- 项 8：分页一致性与参数校验 ----------
    def test_pagination_consistency(self):
        base = (f"page=1&page_size=100&month=2026-10"
                f"&action=log.page_test")
        _, full = self._get(self.admin, base)
        entries_full = full["data"]["entries"]
        self.assertEqual(full["data"]["total"], 25)
        self.assertEqual(len(entries_full), 25)
        # 倒序（时间互异）
        times = [e["time"] for e in entries_full]
        self.assertEqual(times, sorted(times, reverse=True))
        # page_size=20 → 2 页，拼接与全量一一对应
        _, p1 = self._get(self.admin,
                          "page=1&page_size=20&month=2026-10"
                          "&action=log.page_test")
        _, p2 = self._get(self.admin,
                          "page=2&page_size=20&month=2026-10"
                          "&action=log.page_test")
        self.assertEqual(p1["data"]["total"], 25)
        self.assertEqual(p1["data"]["pages"], 2)
        self.assertEqual(len(p1["data"]["entries"]), 20)
        self.assertEqual(len(p2["data"]["entries"]), 5)
        merged = p1["data"]["entries"] + p2["data"]["entries"]
        self.assertEqual([e["time"] for e in merged], times)
        # 越界页码 → 200 空页 + 真实 total/pages
        status, p99 = self._get(self.admin,
                                "page=99&page_size=20&month=2026-10"
                                "&action=log.page_test")
        self.assertEqual(status, 200)
        self.assertEqual(p99["data"]["total"], 25)
        self.assertEqual(p99["data"]["entries"], [])

    def test_invalid_pagination_params(self):
        for query in ("page=abc&page_size=50",
                      "page=1&page_size=999",
                      "page=0&page_size=50"):
            status, resp = self._get(self.admin, query)
            self.assertEqual(status, 400, query)
            self.assertIn("message", resp)

    # ---------- 项 9：普通成员越权 403 ----------
    def test_member_forbidden(self):
        status, resp = self._get(
            self.member, f"page=1&page_size=50&member={self.other_user_id}")
        self.assertEqual(status, 403)
        status, resp = self._get(
            self.member, f"page=1&page_size=50&store_id={self.other_store}")
        self.assertEqual(status, 403)

    # ---------- 辅助：旧 limit 模式兼容 ----------
    def test_limit_mode(self):
        status, resp = self._get(self.admin, "action=log.page_test&limit=10")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["total"], 10)
        self.assertEqual(len(resp["data"]["entries"]), 10)
        self.assertNotIn("pages", resp["data"])
        # limit 截断至 LOG_QUERY_MAX=1000（全量 25 条 < 1000）
        status, resp = self._get(self.admin, "action=log.page_test&limit=99999")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["total"], 25)
        # limit 非法 → 400
        status, resp = self._get(self.admin, "limit=abc")
        self.assertEqual(status, 400)

    # ---------- 辅助：服务端 traceback 落日志且不外泄 ----------
    def test_traceback_logged_not_leaked(self):
        bundle = self.env.server.registry.get(self.member_store)
        buf = io.StringIO()
        with mock.patch.object(
                bundle.op_log, "query",
                side_effect=TypeError(
                    "cannot use 'dict' as a set element "
                    "(unhashable type: 'dict')")):
            with redirect_stderr(buf):
                status, resp = self._get(
                    self.member, "page=1&page_size=50&month=2026-10")
        self.assertEqual(status, 500)
        body_text = json.dumps(resp, ensure_ascii=False)
        self.assertIn("服务器内部错误", body_text)
        self.assertNotIn("Traceback", body_text)
        self.assertNotIn("File \"", body_text)
        err = buf.getvalue()
        self.assertIn("Traceback", err)
        self.assertIn("File", err)
        self.assertIn("TypeError", err)

    # ---------- 项 11：共享查询函数页面无回归 ----------
    def test_shared_pages_regression(self):
        member_paths = [
            "/api/orders?page=1&page_size=20&month=2026-10",
            "/api/refund-orders?page=1&page_size=20&month=2026-10",
            "/api/refund-details?page=1&page_size=20&month=2026-10",
            "/api/supplier-costs?page=1&page_size=20&month=2026-10",
            "/api/anomalies?page=1&page_size=20&month=2026-10",
            "/api/reconciliation?page=1&page_size=20&month=2026-10",
            "/api/profit?page=1&page_size=20&month=2026-10",
            "/api/import/batches?page=1&page_size=20",
        ]
        for path in member_paths:
            status, resp = self.member.get(path)
            self.assertEqual(status, 200, path)
            body_text = json.dumps(resp, ensure_ascii=False)
            self.assertNotIn("unhashable type", body_text)
            self.assertNotIn("服务器内部错误", body_text)
        admin_paths = ["/api/users", "/api/stores"]
        for path in admin_paths:
            status, resp = self.admin.get(path)
            self.assertEqual(status, 200, path)


if __name__ == "__main__":
    unittest.main()

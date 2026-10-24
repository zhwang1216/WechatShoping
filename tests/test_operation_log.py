import json
import shutil
import tempfile
import unittest

from app.operation_log import OperationLog
from app.time_range import month_bounds


class TestOperationLog(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.log = OperationLog(log_dir=self.tmp)

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def test_append_and_query_newest_first(self):
        self.log.append("server.start", detail={"port": 1})
        self.log.append("data.write.orders", detail={"count": 3}, operator="admin")
        entries = self.log.query()
        self.assertEqual(len(entries), 2)
        self.assertEqual(entries[0]["action"], "data.write.orders")
        self.assertEqual(entries[1]["action"], "server.start")
        self.assertEqual(entries[0]["operator"], "admin")

    def test_query_empty(self):
        self.assertEqual(self.log.query(), [])

    def test_query_limit(self):
        for i in range(10):
            self.log.append(f"act.{i}")
        self.assertEqual(len(self.log.query(limit=3)), 3)

    def test_query_by_action_and_operator(self):
        self.log.append("a.x", operator="u1")
        self.log.append("a.x", operator="u2")
        self.log.append("a.y", operator="u1")
        self.assertEqual(len(self.log.query(action="a.x")), 2)
        self.assertEqual(len(self.log.query(operator="u1")), 2)
        self.assertEqual(len(self.log.query(action="a.x", operator="u1")), 1)

    def _inject_line(self, entry):
        with open(self.log.log_file, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")

    def test_operator_filter_hits_dict(self):
        # spec 5.4.1-3：历史脏数据 operator 为用户 dict（旧版 store.rename 写入），
        # query(operator="admin") 必须命中该记录（归一化先于筛选）。
        self._inject_line({
            "time": "2026-10-21T23:16:17",
            "operator": {"user_id": "u_f7c06a4c", "username": "admin",
                         "password_hash": "pbkdf2_sha256$...", "role": "admin"},
            "action": "store.rename",
            "result": "success",
            "detail": {"store_id": "s_x", "old_name": "a", "new_name": "b"},
        })
        entries = self.log.query(operator="admin")
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["action"], "store.rename")
        self.assertEqual(entries[0]["operator"], "admin")

    def test_file_is_valid_jsonl(self):
        self.log.append("t", detail={"k": "中文"})
        with open(self.log.log_file, encoding="utf-8") as f:
            lines = [line for line in f.read().splitlines() if line]
        self.assertEqual(len(lines), 1)
        entry = json.loads(lines[0])
        self.assertEqual(entry["action"], "t")
        self.assertEqual(entry["detail"]["k"], "中文")

    def test_dict_operator_normalized(self):
        # 读取契约：dict 脏数据归一化为 username 字符串返回；
        # 磁盘 JSONL 该行原样保留（operator 仍为 dict）。
        self._inject_line({
            "time": "2026-10-21T23:16:17",
            "operator": {"user_id": "u_f7c06a4c", "username": "admin",
                         "password_hash": "pbkdf2_sha256$..."},
            "action": "store.rename", "result": "success",
            "detail": {"store_id": "s_x"},
        })
        entries = self.log.query()
        self.assertEqual(entries[0]["operator"], "admin")
        self.assertNotIn("password_hash", json.dumps(entries))
        with open(self.log.log_file, encoding="utf-8") as f:
            disk = json.loads(f.read().splitlines()[0])
        self.assertIsInstance(disk["operator"], dict)

    def test_operator_fallback_chain(self):
        cases = [
            ({"user_id": "u_1", "password_hash": "x"}, "u_1"),
            ({"user_id": "", "username": ""}, "unknown"),
            ({"role": "admin"}, "unknown"),
            (None, None),
            (12345, "12345"),
            (["a", "b"], "unknown"),
        ]
        for i, (raw, expected) in enumerate(cases):
            self._inject_line({"time": f"2026-10-01T00:00:{i:02d}",
                               "operator": raw, "action": f"a.{i}",
                               "result": "success", "detail": None})
        entries = self.log.query()
        got = [e["operator"] for e in reversed(entries)]
        self.assertEqual(got, [expected for _, expected in cases])

    def test_keyword_excludes_sensitive_fields(self):
        # 归一化后匹配文本不含 password_hash，敏感字段不可被检索。
        self._inject_line({
            "time": "2026-10-21T23:16:17",
            "operator": {"user_id": "u_x", "username": "admin",
                         "password_hash": "pbkdf2_sha256$secret"},
            "action": "store.rename", "result": "success",
            "detail": {"store_id": "s_x"},
        })
        self.assertEqual(self.log.query(keyword="password_hash"), [])
        self.assertEqual(self.log.query(keyword="secret"), [])
        self.assertEqual(len(self.log.query(keyword="admin")), 1)

    def test_corrupt_line_skipped(self):
        self.log.append("ok.1")
        with open(self.log.log_file, "a", encoding="utf-8") as f:
            f.write("{not json\n\n")
        self.log.append("ok.2")
        entries = self.log.query()
        self.assertEqual([e["action"] for e in entries], ["ok.2", "ok.1"])

    def test_month_bounds(self):
        # 左闭右开：当月第一天 00:00:00 命中，下月第一天 00:00:00 排他，
        # 当月最后一天 23:59:59.5（毫秒精度）必须命中。
        rows = [
            ("2026-10-01T00:00:00", "in.first"),
            ("2026-10-31T23:59:59.5", "in.last"),
            ("2026-11-01T00:00:00", "out.next"),
            ("2026-09-30T23:59:59", "out.prev"),
        ]
        for t, action in rows:
            self._inject_line({"time": t, "operator": "admin",
                               "action": action, "result": "success",
                               "detail": None})
        start, end = month_bounds("2026-10")
        entries = self.log.query(start=start, end=end)
        self.assertEqual([e["action"] for e in entries],
                         ["in.last", "in.first"])

    def test_disk_unchanged(self):
        self.log.append("a", operator="admin")
        self._inject_line({"time": "2026-10-21T23:16:17",
                           "operator": {"username": "admin",
                                        "password_hash": "x"},
                           "action": "store.rename", "result": "success",
                           "detail": {}})
        with open(self.log.log_file, "rb") as f:
            before = f.read()
        self.log.query()
        self.log.query(operator="admin")
        self.log.query(keyword="rename")
        with open(self.log.log_file, "rb") as f:
            after = f.read()
        self.assertEqual(before, after)

    def test_append_defensive_normalization(self):
        user = {"user_id": "u_1", "username": "admin",
                "password_hash": "pbkdf2_sha256$secret", "role": "admin"}
        entry = self.log.append("store.rename", detail={"k": 1},
                                operator=user)
        self.assertEqual(entry["operator"], "admin")
        with open(self.log.log_file, encoding="utf-8") as f:
            disk = json.loads(f.read().splitlines()[0])
        self.assertEqual(disk["operator"], "admin")
        self.assertNotIn("password_hash", json.dumps(disk))
        entry2 = self.log.append("x", operator=None)
        self.assertEqual(entry2["operator"], "system")


if __name__ == "__main__":
    unittest.main()

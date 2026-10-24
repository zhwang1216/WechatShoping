import unittest

from tests.http_client import TestEnv


class TestServer(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        # 管理员无业务店铺: 数据文件/日志/备份等业务断言经成员客户端承载
        cls.member = cls.env.member

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _request(self, path, body=None):
        if body is None:
            return self.member.get(path)
        return self.member.post(path, body)

    def test_health_ok(self):
        status, resp = self._request("/api/health")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        self.assertEqual(resp["data"]["status"], "ok")
        self.assertTrue(resp["data"]["python"])
        for name, info in resp["data"]["data_files"].items():
            self.assertTrue(info["exists"], f"{name} 应存在")
            self.assertTrue(info["valid"], f"{name} 应合法")

    def test_get_settings(self):
        status, resp = self._request("/api/data/settings")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        self.assertIn("settings", resp["data"])

    def test_get_unknown_data_404(self):
        status, resp = self._request("/api/data/unknown")
        self.assertEqual(status, 404)
        self.assertEqual(resp["code"], 404)

    def test_write_and_read_data(self):
        status, resp = self._request("/api/data/orders", {"records": [{"order_id": "TEST-1"}]})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        self.assertEqual(resp["message"], "写入成功")
        status, resp = self._request("/api/data/orders")
        self.assertEqual(resp["data"]["records"][0]["order_id"], "TEST-1")

    def test_write_invalid_body_rejected_and_data_intact(self):
        status, resp = self._request("/api/data/orders", {"wrong": 1})
        self.assertEqual(status, 400)
        self.assertIn("records", resp["message"])
        status, resp = self._request("/api/data/orders")
        self.assertEqual(resp["code"], 0)

    def test_write_non_object_collection_rejected(self):
        status, resp = self._request("/api/data/orders", {"records": "not-a-list"})
        self.assertEqual(status, 400)

    def test_backup_and_restore_api(self):
        self._request("/api/data/orders", {"records": [{"order_id": "B1"}]})
        status, resp = self._request("/api/backup", {"name": "orders"})
        self.assertEqual(resp["code"], 0)
        backup_file = resp["data"]["backup_file"]

        self._request("/api/data/orders", {"records": [{"order_id": "B2"}]})
        status, resp = self._request(
            "/api/restore", {"name": "orders", "backup_file": backup_file})
        self.assertEqual(resp["code"], 0)
        self.assertEqual(resp["message"], "恢复成功")

        status, resp = self._request("/api/data/orders")
        self.assertEqual(resp["data"]["records"][0]["order_id"], "B1")

    def test_list_backups_api(self):
        status, resp = self._request("/api/backups?name=orders")
        self.assertEqual(resp["code"], 0)
        self.assertGreaterEqual(resp["data"]["total"], 0)

    def test_logs_api(self):
        # server.start 属于系统级生命周期日志（全局安全日志，见 test_auth_isolation），
        # 店铺操作日志仅记录本店铺的业务操作，故此处断言刚触发的写入操作
        self._request("/api/data/orders", {"records": [{"order_id": "LOG-1"}]})
        status, resp = self._request("/api/logs?limit=5")
        self.assertEqual(resp["code"], 0)
        self.assertGreater(resp["data"]["total"], 0)
        actions = {e["action"] for e in resp["data"]["entries"]}
        self.assertIn("data.write.orders", actions)

    def test_unknown_path_404(self):
        status, resp = self._request("/api/not_exist")
        self.assertEqual(status, 404)
        self.assertEqual(resp["code"], 404)


if __name__ == "__main__":
    unittest.main()

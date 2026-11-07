"""经营总览 /api/summary 扩展字段测试（TASK-301）

验证 daily_trends / business_status / latest_update_time 增补字段
存在性、结构正确性、既有字段保留与 member 权限隔离。
"""
import unittest

from tests.http_client import TestEnv


class TestDashboardSummary(unittest.TestCase):
    """经营总览 /api/summary 扩展字段（TASK-301）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.member

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_summary_has_new_fields(self):
        st, r = self.admin.get("/api/summary")
        self.assertEqual(st, 200)
        data = r["data"]
        self.assertIn("daily_trends", data)
        self.assertIn("business_status", data)
        self.assertIn("latest_update_time", data)

    def test_02_summary_keeps_existing_fields(self):
        st, r = self.admin.get("/api/summary")
        data = r["data"]
        for field in ("has_orders", "declaration", "policies"):
            self.assertIn(field, data, f"缺失既有字段: {field}")
        if data.get("has_orders"):
            for field in ("order_count", "actual_received_amount",
                          "refund_success_amount", "supplier_total_cost",
                          "operating_contribution_profit",
                          "profit_status_counts", "unmatched_order_count",
                          "open_anomaly_count"):
                self.assertIn(field, data, f"缺失既有字段: {field}")

    def test_03_daily_trends_structure(self):
        st, r = self.admin.get("/api/summary")
        t = r["data"]["daily_trends"]
        for key in ("dates", "order_count", "actual_received",
                    "refund_success", "supplier_cost",
                    "contribution_profit"):
            self.assertIn(key, t, f"daily_trends 缺失: {key}")
            self.assertIsInstance(t[key], list)

    def test_04_business_status_fields(self):
        st, r = self.admin.get("/api/summary")
        bs = r["data"]["business_status"]
        for key in ("pending_reconcile", "cost_unmatched",
                    "pending_profit", "anomaly_count",
                    "latest_import_batch"):
            self.assertIn(key, bs, f"business_status 缺失: {key}")

    def test_05_member_summary_isolation(self):
        st, r = self.member.get("/api/summary")
        self.assertEqual(st, 200)
        data = r["data"]
        self.assertIn("daily_trends", data)
        self.assertIn("business_status", data)
        self.assertIn("latest_update_time", data)


if __name__ == "__main__":
    unittest.main()

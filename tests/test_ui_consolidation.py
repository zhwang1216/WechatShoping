"""前端页面复用整合回归测试

覆盖本次页面整合的关键约束：
- 应用外壳（侧边栏/顶栏）由服务端按路由注入，页面不再各自复制；
- 列表页统一使用 ListPage + CommonTable，订单/退款详情复用同一抽屉视图；
- 登录/注册/初始化页复用同一套认证卡片样式；
- 各页面导航仍只有一份，且不残留占位符。
"""
import re
import unittest

from pathlib import Path

from tests.http_client import Client, TestEnv

SHELL_PAGES = [
    "/dashboard", "/data-import", "/orders", "/refund-orders",
    "/supplier-costs", "/reconciliation", "/profit", "/monthly-profit",
    "/anomalies", "/logs", "/users", "/users/create", "/stores",
]

LIST_PAGES = ["/orders", "/refund-orders", "/supplier-costs",
              "/anomalies", "/logs"]


class TestUIConsolidation(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _page(self, path):
        st, body, _ = self.admin.get_page(path)
        self.assertEqual(st, 200, path)
        return body

    def _static(self, path):
        st, body, _ = self.admin.get_binary("/static/" + path)
        self.assertEqual(st, 200, path)
        return body.decode("utf-8")

    def test_01_shared_assets_served(self):
        self.assertIn("function ListPage", self._static("list_page.js"))
        self.assertIn("function openOrderDetail", self._static("order_detail.js"))
        self.assertIn("function openDrawer", self._static("modal.js"))
        style = self._static("style.css")
        self.assertIn("responsive.css", style)
        resp = self._static("responsive.css")
        self.assertIn("overflow-x: clip", resp)
        self.assertIn("modal-drawer", resp)

    def test_02_list_pages_use_shared_components(self):
        for path in LIST_PAGES:
            html = self._page(path)
            self.assertIn("/static/list_page.js", html, path)
            self.assertIn("/static/common_table.js", html, path)
            self.assertIn('id="store-scope"', html, path)
            self.assertIn('id="table-state"', html, path)
            self.assertIn('id="pager"', html, path)

    def test_03_app_shell_injected_once(self):
        for path in SHELL_PAGES:
            html = self._page(path)
            self.assertNotIn("<!--@app-shell-->", html, path)
            self.assertEqual(html.count('id="app-sidebar"'), 1, path)
            self.assertEqual(html.count("<nav>"), 1, path)
            self.assertEqual(html.count('class="topbar"'), 1, path)
            nav = re.search(r"<nav>(.*?)</nav>", html, re.S).group(1)
            self.assertNotIn("/static/", nav, path)

    def test_04_order_pages_share_detail_drawer(self):
        for path in ("/orders", "/refund-orders"):
            html = self._page(path)
            self.assertIn("/static/modal.js", html, path)
            self.assertIn("/static/order_detail.js", html, path)
            # 旧页面内嵌详情模态已移除，统一走 openDrawer
            self.assertNotIn('id="detail-modal"', html, path)

    def test_05_auth_pages_share_card_styles(self):
        anon = Client(self.env)
        for path in ("/login", "/register"):
            st, html, _ = anon.get_page(path)
            self.assertEqual(st, 200, path)
            self.assertIn('class="auth-card', html, path)
            self.assertIn('class="auth-form"', html, path)
            self.assertNotIn("<style>", html, path)
        # /setup 在已初始化环境会重定向，直接校验静态源文件同一套样式类
        setup_html = (Path(__file__).resolve().parents[1] / "static" /
                      "setup.html").read_text(encoding="utf-8")
        self.assertIn('class="auth-card', setup_html)
        self.assertIn('class="auth-form"', setup_html)
        self.assertNotIn("<style>", setup_html)


if __name__ == "__main__":
    unittest.main()

"""页面导航路由测试：业务页面固定顺序、语义化 URL、默认页、高亮、旧路径兼容

需求口径：
- 菜单固定顺序：经营总览/数据导入/每月盈利/对账成本/利润核算/
  普通订单/退款订单/供应商成本/用户管理(组)/异常处理/操作日志
- 根路径 / 默认进入"经营总览"；刷新按 URL 打开对应页面
- 每页导航含自身 active 高亮；"对账成本"为标准菜单名
- admin 视角追加"用户管理"菜单组（组标题"用户管理"，二级项创建用户/店铺管理/
  用户管理，仅 admin 输出；member 视角由服务端剔除标记块，nav 仅含业务项）
"""
import re
import unittest

from tests.http_client import TestEnv

NAV_ORDER = [
    ("/dashboard", "经营总览"),
    ("/data-import", "数据导入"),
    ("/monthly-profit", "每月盈利"),
    ("/reconciliation", "对账成本"),
    ("/profit", "利润核算"),
    ("/orders", "普通订单"),
    ("/refund-orders", "退款订单"),
    ("/supplier-costs", "供应商成本"),
    ("/anomalies", "异常处理"),
    ("/logs", "操作日志"),
]

MGMT_ORDER = [
    ("/users/create", "创建用户"),
    ("/stores", "店铺管理"),
    ("/users", "用户管理"),
]

EXPECTED_LINKS = NAV_ORDER[:8] + MGMT_ORDER + NAV_ORDER[8:]

PAGE_FEATURES = {
    "/dashboard": "经营总览",
    "/data-import": "数据导入",
    "/orders": "普通订单",
    "/refund-orders": "退款订单",
    "/supplier-costs": "供应商成本",
    "/reconciliation": "对账成本",
    "/profit": "利润核算",
    "/monthly-profit": "每月盈利",
    "/anomalies": "异常处理",
    "/logs": "操作日志",
}


class TestPageRoutes(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _get(self, path, expect_status=200):
        status, body, headers = self.admin.get_page(path)
        if status != expect_status:
            self.fail(f"GET {path} 期望 {expect_status} 实际 {status}: {body[:200]}")
        return status, body, headers

    def test_01_root_defaults_to_dashboard(self):
        """访问根路径且未指定路由时，默认进入经营总览"""
        status, body, headers = self._get("/")
        self.assertEqual(status, 200)
        self.assertIn("经营总览", body)
        self.assertIn('href="/dashboard" class="active"', body)
        self.assertTrue(headers.get("Content-Type", "").startswith("text/html"))

    def test_02_all_page_routes_serve_correct_page(self):
        """各语义路由均返回 200 且内容与页面职责匹配"""
        for path, feature in PAGE_FEATURES.items():
            with self.subTest(path=path):
                status, body, headers = self._get(path)
                self.assertEqual(status, 200)
                self.assertIn(feature, body)
                self.assertNotIn("成本对账", body)
                self.assertTrue(
                    headers.get("Content-Type", "").startswith("text/html"))

    def test_03_nav_order_and_names_on_every_page(self):
        """每个页面导航均为固定顺序、固定名称（业务项 + 管理菜单组 3 项）"""
        for path, _ in NAV_ORDER:
            with self.subTest(path=path):
                _, body, _ = self._get(path)
                nav = re.search(r"<nav>(.*?)</nav>", body, re.S)
                self.assertIsNotNone(nav, f"{path} 缺少导航")
                links = re.findall(
                    r'<a href="([^"]+)"(?: class="active")?>([^<]+)</a>',
                    nav.group(1))
                self.assertEqual(
                    links, EXPECTED_LINKS,
                    f"{path} 导航顺序/名称不符合固定顺序")
                self.assertIn(
                    '<button type="button" class="nav-group" '
                    'data-nav-group="user-mgmt"',
                    nav.group(1), f"{path} 缺少用户管理组标题")
                self.assertIn(
                    ">用户管理</button>", nav.group(1),
                    f"{path} 组标题文案应为「用户管理」")

    def test_04_active_highlight_on_every_page(self):
        """每个页面的导航中自身菜单项有 active 高亮，且仅一项"""
        for path, _ in NAV_ORDER:
            with self.subTest(path=path):
                _, body, _ = self._get(path)
                self.assertIn(f'href="{path}" class="active"', body)
                self.assertEqual(body.count('class="active"'), 1,
                                 f"{path} active 高亮数量应为 1")
        for path, _ in MGMT_ORDER:
            with self.subTest(path=path):
                _, body, _ = self._get(path)
                self.assertIn(f'href="{path}" class="active"', body)
                self.assertEqual(body.count('class="active"'), 1,
                                 f"{path} active 高亮数量应为 1")
                self.assertIn("nav-group-current", body,
                              f"{path} 组标题缺少 nav-group-current 当前态")
                self.assertIn(">用户管理</button>", body,
                              f"{path} 组标题文案应为「用户管理」")

    def test_05_legacy_static_paths_still_work(self):
        """旧 /static/xxx.html 路径保持兼容（历史书签可用）"""
        for path, feature in [
            ("/static/dashboard.html", "经营总览"),
            ("/static/index.html", "数据导入"),
            ("/static/reconciliation.html", "对账成本"),
            ("/static/profit.html", "利润核算"),
        ]:
            with self.subTest(path=path):
                status, body, _ = self._get(path)
                self.assertEqual(status, 200)
                self.assertIn(feature, body)

    def test_06_unknown_page_returns_404(self):
        """未知页面路径返回 404，不误路由到其他页面"""
        status, body, _ = self._get("/no-such-page", expect_status=404)
        self.assertIn("不存在", body)

    def test_07_nav_links_use_semantic_routes_only(self):
        """导航不使用 /static/*.html 旧链接；数据导入不再挂在根路径 /"""
        for path, _ in NAV_ORDER:
            with self.subTest(path=path):
                _, body, _ = self._get(path)
                nav = re.search(r"<nav>(.*?)</nav>", body, re.S).group(1)
                self.assertNotIn("/static/", nav)
                self.assertNotIn('href="/"', nav)

    def test_08_member_nav_excludes_admin_group(self):
        """member 视角导航经服务端剔除: 仅业务项, 无管理菜单组结构,
        且不含异常处理/操作日志（仅 admin 可见）"""
        member = self.env.create_user_client(
            "route_member", "Member123456", store_name="路由测试店铺")
        member_nav = NAV_ORDER[:8]
        for path, _ in member_nav:
            with self.subTest(path=path):
                status, body, _ = member.get_page(path)
                self.assertEqual(status, 200)
                nav = re.search(r"<nav>(.*?)</nav>", body, re.S).group(1)
                links = re.findall(
                    r'<a href="([^"]+)"(?: class="active")?>([^<]+)</a>',
                    nav)
                self.assertEqual(links, member_nav,
                                 f"{path} member 视角应与业务菜单一致")
                self.assertNotIn("admin-nav-start", body)
                self.assertNotIn("admin-only-nav-start", body)
                self.assertNotIn("nav-group", nav)
                self.assertNotIn("<button", nav)
                self.assertNotIn("用户管理", nav)
                self.assertNotIn('href="/users/create"', nav)
                self.assertNotIn('href="/stores"', nav)
                self.assertNotIn('href="/users"', nav)
                self.assertNotIn('href="/anomalies"', nav)
                self.assertNotIn('href="/logs"', nav)
                self.assertNotIn("异常处理", nav)
                self.assertNotIn("操作日志", nav)
        for path in ("/anomalies", "/logs"):
            with self.subTest(path=path):
                status, _, _ = member.get_page(path)
                self.assertNotEqual(status, 200,
                                    f"{path} member 应被拒绝访问")


if __name__ == "__main__":
    unittest.main()

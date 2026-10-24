"""密码输入体验统一验收测试：小眼睛公共组件 + 两次密码一致性校验

覆盖任务验收标准（真实 HTTP 驱动, 隔离临时数据环境）:
- TestPasswordUxPages: 登录/注册/初始化/修改密码弹窗/重置密码弹窗的静态结构
  （默认 type=password、小眼睛注入调用、实时一致性提示、提交带 confirm_password）
- TestChangePasswordConsistency: 后端 change-password 一致性校验
  （缺失/不一致拒绝; 一致成功后旧密码拒登新密码可登; 业务数据零影响）
- 注册接口两次密码不一致的拒绝已由 tests/test_registration.py
  S4/test_09 用例覆盖（"两次输入的密码不一致"）, 此处不重复。
"""
import unittest

from tests.http_client import Client, TestEnv
from tests.test_auth_isolation import make_wx_xlsx
from tests.test_member_management import order_count

PASSWORD = "Uxuser123456"
NEW_PASSWORD = "Uxnew123456"


class TestPasswordUxPages(unittest.TestCase):
    """五个密码表单均接入统一公共组件（小眼睛 + 实时一致性校验）"""

    @classmethod
    def setUpClass(cls):
        # 未初始化环境: /setup 页面可直接访问
        cls.fresh_env = TestEnv(initialize=False)
        # 已初始化环境: /users 页面需管理员登录访问
        cls.env = TestEnv()
        cls.admin = cls.env.admin

    @classmethod
    def tearDownClass(cls):
        cls.fresh_env.cleanup()
        cls.env.cleanup()

    # ---------- 公共组件与样式 ----------
    def test_01_api_js_has_toggle_and_match(self):
        status, js, _ = Client(self.env).get_page("/static/api.js")
        self.assertEqual(status, 200)
        self.assertIn("function setupPasswordToggle", js)
        self.assertIn("function bindPasswordMatch", js)
        self.assertIn("input.type = \"password\"", js)  # 默认隐藏
        self.assertIn("btn.type = \"button\"", js)      # 不触发表单提交
        self.assertIn("aria-label", js)
        self.assertIn("aria-pressed", js)
        self.assertIn("确认密码与密码不一致", js)
        self.assertIn("两次输入一致", js)
        self.assertNotIn("localStorage", js.split("setupPasswordToggle")[1]
                         .split("function bindPasswordMatch")[0])  # 不持久化密码

    def test_02_style_css_has_toggle_styles(self):
        status, css, _ = Client(self.env).get_page("/static/style.css")
        self.assertEqual(status, 200)
        self.assertIn(".pwd-field", css)
        self.assertIn(".pwd-toggle", css)
        self.assertIn(".pwd-match-hint", css)

    # ---------- 登录页 ----------
    def test_03_login_page_toggle(self):
        status, html, _ = Client(self.env).get_page("/login")
        self.assertEqual(status, 200)
        self.assertIn('id="password"', html)
        self.assertIn('type="password"', html)  # 默认隐藏
        self.assertIn("setupPasswordToggle", html)  # 小眼睛注入
        # Enter 登录: 输入框在 form 内 + 按钮 type=submit, 天然支持
        self.assertIn('id="login-form"', html)
        self.assertIn('type="submit"', html)
        self.assertIn('id="login-error"', html)  # 登录失败提示容器

    # ---------- 注册页 ----------
    def test_04_register_page_toggle_and_match(self):
        status, html, _ = Client(self.env).get_page("/register")
        self.assertEqual(status, 200)
        self.assertEqual(html.count('type="password"'), 2)  # 密码+确认密码默认隐藏
        self.assertIn('id="confirm"', html)
        self.assertEqual(html.count("setupPasswordToggle"), 2)  # 两个独立小眼睛
        self.assertIn("bindPasswordMatch", html)  # 实时一致性校验
        self.assertIn('id="pwd-match"', html)     # 确认密码框附近提示容器
        # 提交时仍保留最终拦截（后端之外的双保险）
        self.assertIn("password !== confirm", html)

    # ---------- 初始化页 ----------
    def test_05_setup_page_toggle_and_match(self):
        status, html, _ = Client(self.fresh_env).get_page("/setup")
        self.assertEqual(status, 200)
        self.assertEqual(html.count('type="password"'), 2)
        self.assertEqual(html.count("setupPasswordToggle"), 2)
        self.assertIn("bindPasswordMatch", html)
        self.assertIn('id="pwd-match"', html)

    # ---------- 修改密码弹窗（auth_nav.js, 全业务页共用） ----------
    def test_06_change_password_modal(self):
        status, js, _ = Client(self.env).get_page("/static/auth_nav.js")
        self.assertEqual(status, 200)
        # 三个输入框 + 三个独立小眼睛
        self.assertIn('id="pwd-old"', js)
        self.assertIn('id="pwd-new"', js)
        self.assertIn('id="pwd-confirm"', js)
        self.assertEqual(js.count("setupPasswordToggle"), 3)
        # 实时一致性 + 不一致禁止提交
        self.assertIn("bindPasswordMatch", js)
        self.assertIn("pwd-submit", js)
        # 提交时向服务端发送两次密码, 后端独立校验
        self.assertIn("confirm_password", js)
        self.assertIn("newPwd !== confirmPwd", js)  # 提交时最终拦截

    # ---------- 管理员重置密码弹窗（users.html, 单输入框） ----------
    def test_07_reset_password_modal_toggle(self):
        status, html, _ = self.admin.get_page("/users")
        self.assertEqual(status, 200)
        self.assertIn('id="rp-new"', html)
        self.assertIn("setupPasswordToggle", html)

    # ---------- 后端注册一致性校验既有覆盖核验 ----------
    def test_08_register_api_rejects_mismatch(self):
        """直接调注册 API 传不一致的两次密码: 后端必须拒绝（验收第二节.7）"""
        anon = Client(self.env)
        status, resp = anon.post("/api/auth/register", {
            "username": "uxmismatch1", "password": "Uxuser123456",
            "confirm_password": "Different123", "store_name": "不一致店铺"})
        self.assertEqual(status, 400)
        self.assertIn("不一致", resp["message"])


class TestChangePasswordConsistency(unittest.TestCase):
    """后端 change-password: 两次新密码一致性 + 身份与业务数据保护"""

    USER = "uxuser"
    STORE = "密码体验测试店"

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.create_user_client(
            cls.USER, PASSWORD, store_name=cls.STORE)
        status, resp = cls.member.upload_confirm(
            "/api/import/wechat-orders", "UX订单.xlsx", make_wx_xlsx("UXP"))
        assert status == 200, f"导入失败: {resp}"
        cls.orders_before = order_count(cls.member)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def test_01_missing_confirm_rejected(self):
        """直接调 API 不带 confirm_password: 拒绝, 不允许绕过一致性校验"""
        status, resp = self.member.post("/api/auth/change-password", {
            "old_password": PASSWORD, "new_password": NEW_PASSWORD})
        self.assertEqual(status, 400)
        self.assertIn("请再次输入新密码", resp["message"])
        # 修改未生效: 原密码仍可登录
        client = Client(self.env)
        status, _ = client.post("/api/auth/login",
                                {"username": self.USER, "password": PASSWORD},
                                csrf=False)
        self.assertEqual(status, 200)

    def test_02_mismatch_rejected(self):
        """两次新密码不一致: 拒绝且提示明确"""
        status, resp = self.member.post("/api/auth/change-password", {
            "old_password": PASSWORD, "new_password": NEW_PASSWORD,
            "confirm_password": "Other123456"})
        self.assertEqual(status, 400)
        self.assertIn("两次输入的新密码不一致", resp["message"])
        client = Client(self.env)
        status, _ = client.post("/api/auth/login",
                                {"username": self.USER, "password": PASSWORD},
                                csrf=False)
        self.assertEqual(status, 200)

    def test_03_wrong_old_password_rejected(self):
        """原密码错误: 明确拒绝"""
        status, resp = self.member.post("/api/auth/change-password", {
            "old_password": "WrongOld123", "new_password": NEW_PASSWORD,
            "confirm_password": NEW_PASSWORD})
        self.assertEqual(status, 400)
        self.assertIn("原密码不正确", resp["message"])

    def test_04_success_flows(self):
        """一致时成功: 旧密码不能登录、新密码可以、旧会话失效、业务数据不变"""
        status, resp = self.member.post("/api/auth/change-password", {
            "old_password": PASSWORD, "new_password": NEW_PASSWORD,
            "confirm_password": NEW_PASSWORD})
        self.assertEqual(status, 200)
        # 旧会话失效
        status, _ = self.member.get("/api/auth/me")
        self.assertEqual(status, 401)
        # 旧密码不能登录
        client = Client(self.env)
        status, _ = client.post("/api/auth/login",
                                {"username": self.USER, "password": PASSWORD},
                                csrf=False)
        self.assertEqual(status, 401)
        # 新密码可以登录
        self.member = self.env.login_client(self.USER, NEW_PASSWORD)
        # 业务数据不受任何影响（订单/店铺归属不变）
        self.assertEqual(order_count(self.member), self.orders_before)
        status, resp = self.member.get("/api/auth/me")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["store_name"], self.STORE)

    def test_05_audit_log_no_plaintext_password(self):
        """安全日志不落明文密码"""
        import json
        import os
        path = os.path.join(self.env.log_dir, "security_log.jsonl")
        with open(path, encoding="utf-8") as f:
            content = f.read().lower()
        self.assertNotIn("uxuser123456", content)
        self.assertNotIn("uxnew123456", content)


if __name__ == "__main__":
    unittest.main()

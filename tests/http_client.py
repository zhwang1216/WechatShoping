"""公共 HTTP 测试辅助：真实服务 + 登录态客户端（多身份）+ 重启支持

- 以 tempfile.mkdtemp() 构造独立数据根，启动真实 ReconcileServer（非 mock）
- 经 POST /api/setup/init 一次性初始化测试管理员（可选延迟初始化）
- Client 封装登录态请求：自动携带会话 Cookie 与 X-Requested-With 头的
  get / post / upload（X-Filename 头）/ post_binary / get_binary
- 支持经管理员创建用户 A、B 并各自独立登录（会话互不干扰）、指定角色
- 支持公开注册辅助：register_user（匿名注册，不自动登录）与
  register_user_client（注册后登录）；注册限流阈值可经构造参数注入
- 提供基于同一 data_dir 的服务重启辅助（验证会话持久化）
- 测试账号密码仅存在于本文件且仅在临时目录环境使用；严禁读写真实 data/
"""
import json
import os
import shutil
import tempfile
import urllib.error
import urllib.request
from http.cookies import SimpleCookie
from urllib.parse import quote

from app.server import ReconcileServer

ADMIN_USER = "admin"
ADMIN_PASSWORD = "Admin123456"
USER_A = "usera"
USER_A_PASSWORD = "Usera123456"
USER_B = "userb"
USER_B_PASSWORD = "Userb123456"

SESSION_COOKIE = "wxr_session"
CSRF_VALUE = "XMLHttpRequest"


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """禁止自动跟随 302，便于断言重定向行为与目标地址"""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def _parse_json(body):
    if not body:
        return {}
    try:
        return json.loads(body.decode("utf-8"))
    except (ValueError, UnicodeDecodeError):
        return {}


class Client:
    """登录态 HTTP 客户端：自动携带会话 Cookie 与 CSRF 头

    经 TestEnv 创建；base 动态取自所属环境（重启换端口后仍可用）。
    """

    def __init__(self, env):
        self.env = env
        self.cookie = None
        self.username = None
        self.store_id = None
        self._opener = urllib.request.build_opener(_NoRedirect)

    @property
    def base(self):
        return self.env.base

    # ---------- 底层请求 ----------
    def _request(self, method, path, data=None, headers=None, csrf=True):
        req_headers = dict(headers or {})
        if csrf:
            req_headers.setdefault("X-Requested-With", CSRF_VALUE)
        if self.cookie:
            req_headers["Cookie"] = f"{SESSION_COOKIE}={self.cookie}"
        req = urllib.request.Request(self.base + path, data=data,
                                     method=method, headers=req_headers)
        try:
            with self._opener.open(req, timeout=30) as resp:
                return resp.status, resp.headers, resp.read()
        except urllib.error.HTTPError as e:
            body = e.read()
            hdrs = e.headers
            e.close()
            return e.code, hdrs, body

    # ---------- 会话 ----------
    def login(self, username, password, extra=None):
        """登录并保存会话 Cookie，返回登录信息（user_id/store_id 等）

        extra: 向登录体注入额外字段（如伪造 role）用于安全测试，
        服务端应忽略未知字段、以存储身份为准。
        """
        payload = {"username": username, "password": password}
        if extra:
            payload.update(extra)
        payload = json.dumps(payload).encode("utf-8")
        status, headers, body = self._request(
            "POST", "/api/auth/login", data=payload, csrf=False,
            headers={"Content-Type": "application/json"})
        resp = _parse_json(body)
        if status != 200:
            raise AssertionError(f"用户 {username} 登录失败({status}): {resp}")
        jar = SimpleCookie()
        try:
            jar.load(headers.get("Set-Cookie") or "")
        except Exception:
            pass
        morsel = jar.get(SESSION_COOKIE)
        if not morsel or not morsel.value:
            raise AssertionError(f"用户 {username} 登录响应未设置会话 Cookie")
        self.cookie = morsel.value
        self.username = username
        self.store_id = (resp.get("data") or {}).get("store_id")
        return resp.get("data") or {}

    def logout(self):
        """退出登录并清除本地 Cookie，返回 (status, 响应体)"""
        status, resp = self.post("/api/auth/logout")
        self.cookie = None
        return status, resp

    # ---------- 请求辅助 ----------
    def get(self, path):
        """GET JSON API → (status, 响应字典)"""
        status, _, body = self._request("GET", path)
        return status, _parse_json(body)

    def get_page(self, path):
        """GET 页面/静态资源（不跟随 302）→ (status, 文本, headers)"""
        status, headers, body = self._request("GET", path)
        return status, body.decode("utf-8", "replace"), headers

    def get_binary(self, path):
        """GET 导出下载 → (status, 字节内容, headers)"""
        status, headers, body = self._request("GET", path)
        return status, body, headers

    def post(self, path, body=None, csrf=True):
        """POST JSON → (status, 响应字典)；csrf=False 可模拟跨站攻击请求"""
        payload = json.dumps(body or {}).encode("utf-8")
        status, _, raw = self._request(
            "POST", path, data=payload, csrf=csrf,
            headers={"Content-Type": "application/json"})
        return status, _parse_json(raw)

    def upload(self, path, filename, content, csrf=True):
        """POST Excel 二进制上传（X-Filename 头）→ (status, 响应字典)"""
        status, _, raw = self._request(
            "POST", path, data=content, csrf=csrf,
            headers={"X-Filename": quote(filename),
                     "Content-Type": "application/octet-stream"})
        return status, _parse_json(raw)

    def post_binary(self, path, content, headers=None, csrf=True):
        """POST 原始二进制体（自定义头）→ (status, 响应字典)"""
        status, _, raw = self._request("POST", path, data=content,
                                       headers=headers, csrf=csrf)
        return status, _parse_json(raw)

    def upload_confirm(self, path, filename, content, force=False,
                       target_store_id=None):
        """上传并确认导入 → 确认接口的 (status, 响应字典)

        target_store_id: 显式目标店铺（管理员代传或多店指定）。
        未显式指定时自动携带登录本店（登录响应 data.store_id），
        与旧"member 默认本店"行为等价；admin 无店为 None 时保持
        不带参数（保留"缺参必须 400"的可测性）。
        """
        sid = target_store_id or self.store_id
        upload_path = path
        if sid:
            sep = "&" if "?" in path else "?"
            upload_path = f"{path}{sep}store_id={quote(str(sid))}"
        status, resp = self.upload(upload_path, filename, content)
        if status != 200:
            return status, resp
        upload_id = (resp.get("data") or {}).get("upload_id")
        body = {"upload_id": upload_id, "force": force}
        if sid:
            body["target_store_id"] = sid
        return self.post("/api/import/confirm", body)


class TestEnv:
    """测试环境：临时数据根 + 真实 HTTP 服务 + 一次性初始化 + 多身份客户端

    用法（setUpClass）::

        cls.env = TestEnv()
        cls.admin = cls.env.admin

    tearDownClass 调用 ``cls.env.cleanup()``。
    """

    def __init__(self, initialize=True, admin_user=ADMIN_USER,
                 admin_password=ADMIN_PASSWORD,
                 register_rate_window=None, register_rate_max=None):
        self.tmp = tempfile.mkdtemp(prefix="wxr_test_")
        self.data_dir = os.path.join(self.tmp, "data")
        self.backup_dir = os.path.join(self.tmp, "backups")
        self.log_dir = os.path.join(self.tmp, "logs")
        self.admin_user = admin_user
        self.admin_password = admin_password
        self.register_rate_window = register_rate_window
        self.register_rate_max = register_rate_max
        self.admin = None
        self.member = None
        self.member_store = None
        self.server = None
        self._start()
        if initialize:
            self.setup_admin()
            # 管理员无业务店铺（admin_role_sep）: 自动创建业务承载成员
            self.member = self.ensure_member()
            st, r = self.member.get("/api/auth/me")
            self.member_store = r["data"]["store_id"]

    # ---------- 服务生命周期 ----------
    def _start(self):
        self.server = ReconcileServer(
            host="127.0.0.1", port=0,
            data_dir=self.data_dir, backup_dir=self.backup_dir,
            log_dir=self.log_dir,
            register_rate_window=self.register_rate_window,
            register_rate_max=self.register_rate_max)
        self.server.start(blocking=False)

    @property
    def base(self):
        return f"http://127.0.0.1:{self.server.port}"

    def restart(self):
        """同 data_dir 以新实例重启服务（换端口），用于验证会话持久化"""
        if self.server is not None:
            self.server.shutdown()
            self.server = None
        self._start()

    def shutdown(self):
        if self.server is not None:
            self.server.shutdown()
            self.server = None

    def cleanup(self):
        self.shutdown()
        shutil.rmtree(self.tmp, ignore_errors=True)

    # ---------- 初始化与账号 ----------
    def setup_admin(self, store_name=None, acknowledge_legacy=True):
        """POST /api/setup/init 一次性初始化管理员（无业务店铺）；成功后自动登录

        store_name 参数仅为兼容既有调用保留（管理员不再创建业务店铺）。
        可重复调用场景：迁移失败回滚后修复环境再试（幂等由服务端保证）。
        返回 (status, 响应字典)。
        """
        client = Client(self)
        status, resp = client.post("/api/setup/init", {
            "username": self.admin_user,
            "password": self.admin_password,
            "confirm_password": self.admin_password,
            "acknowledge_legacy": acknowledge_legacy,
        }, csrf=False)
        if status == 200 and self.admin is None:
            self.admin = self.login_client(self.admin_user, self.admin_password)
        return status, resp

    def login_client(self, username, password):
        """创建独立客户端并登录（多客户端会话互不干扰）"""
        client = Client(self)
        client.login(username, password)
        return client

    def create_user(self, username, password, role="member", store_name=None):
        """经管理员会话创建用户（默认 member，各自独立店铺），返回响应 data"""
        status, resp = self.admin.post("/api/users", {
            "username": username, "password": password,
            "role": role, "store_name": store_name,
        })
        if status != 200:
            raise AssertionError(f"创建用户 {username} 失败({status}): {resp}")
        return resp["data"]

    def create_user_client(self, username, password, role="member",
                           store_name=None):
        """创建用户并返回其已登录客户端"""
        self.create_user(username, password, role=role, store_name=store_name)
        return self.login_client(username, password)

    def ensure_member(self, username="biz_carrier",
                      password="Carrier123456", store_name=None):
        """创建（或复用）业务承载成员客户端：业务数据测试的标准载体

        管理员无业务店铺后, 既有测试中经 admin 直传/直写的业务数据
        统一改由该成员店铺承载; 幂等, 同名用户已存在时直接复用登录。
        """
        if self.get_user_id(username) is None:
            self.create_user(username, password, role="member",
                             store_name=store_name)
        return self.login_client(username, password)

    def get_user_id(self, username):
        """经管理员查询用户列表, 返回 username 对应 user_id（无则 None）"""
        status, resp = self.admin.get("/api/users")
        if status != 200:
            return None
        for u in resp.get("data", {}).get("users", []):
            if u.get("username") == username:
                return u.get("user_id")
        return None

    # ---------- 公开注册 ----------
    def register_user(self, username, password, store_name=None,
                      confirm_password=None, extra=None):
        """匿名调用 POST /api/auth/register（不自动登录）

        返回 (status, 响应字典)；extra 可注入任意伪造字段（安全测试用）。
        """
        body = {
            "username": username,
            "password": password,
            "confirm_password": confirm_password
            if confirm_password is not None else password,
            "store_name": store_name,
        }
        if extra:
            body.update(extra)
        client = Client(self)
        return client.post("/api/auth/register", body)

    def register_user_client(self, username, password, store_name=None):
        """公开注册成功后登录，返回该用户已登录客户端"""
        status, resp = self.register_user(username, password,
                                          store_name=store_name)
        if status != 200:
            raise AssertionError(f"注册用户 {username} 失败({status}): {resp}")
        return self.login_client(username, password)

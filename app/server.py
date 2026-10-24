"""本地 HTTP 服务：多用户认证、店铺数据隔离、路由、导入、备份恢复、静态页面

认证与隔离架构:
1. 用户/店铺/会话存于数据根目录(users.json/stores.json/sessions.json), 业务数据
   存于 data/stores/{store_id}/ 实现物理隔离;
2. 所有业务 API 从服务端会话获取身份与店铺范围, 绝不信任请求中的 user_id/store_id;
3. 未登录访问业务页面返回 302 到 /login, 访问业务 API 返回 401;
4. 系统首次使用需经 /setup 一次性初始化管理员, 并确认旧数据归属;
5. CSRF 防护: 写请求必须携带 X-Requested-With 自定义头(跨站表单无法伪造)。
"""
import errno
import json
import os
import platform
import re
import shutil
import sys
import traceback
from datetime import datetime
from decimal import Decimal
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from http.cookies import SimpleCookie
from urllib.parse import parse_qs, quote, unquote, urlparse

from . import config
from .api_response import ApiResponse
from .auth import (AuthError, LoginRateLimiter, SecurityLog, SessionStore,
                   UserManager, REGISTER_MAX_REQUESTS,
                   REGISTER_WINDOW_SECONDS, RegistrationRateLimiter)
from .data_schemas import BUSINESS_FILES, SERVICE_NAME, VERSION, now_iso
from .errors import (ImportServiceError, ProfitServiceError,
                     ReconcileServiceError)
from .json_repository import FileLock, RepositoryError
from .list_queries import PAGE_SIZES
from .order_queries import NORMAL_SORT_FIELDS, REFUND_SORT_FIELDS
from .refund_view import SUCCESS_STATUSES
from .store_services import (StoreServiceRegistry, StoreServices,
                             migrate_legacy_store_data)
from .time_range import TimeRangeError, range_meta, resolve_range

UPLOAD_PATHS = {
    "/api/import/wechat-orders": "wechat_orders",
    "/api/import/supplier-costs": "supplier_cost",
}

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# 无需登录即可访问的页面与 API
PUBLIC_PAGES = {"/login", "/register", "/setup"}
PUBLIC_APIS_GET = {"/api/health", "/api/setup/status"}
PUBLIC_APIS_POST = {"/api/auth/login", "/api/setup/init"}

# 页面语义路由：URL -> static/ 下的页面文件
# 根路径默认进入"经营总览"；各菜单项使用语义化 URL，刷新/前进/后退均按 URL 打开对应页面
PAGE_ROUTES = {
    "/": "dashboard.html",
    "/dashboard": "dashboard.html",
    "/data-import": "index.html",
    "/orders": "orders.html",
    "/refund-orders": "refund_orders.html",
    "/refund-details": "refund_details.html",
    "/supplier-costs": "supplier_costs.html",
    "/reconciliation": "reconciliation.html",
    "/profit": "profit.html",
    "/anomalies": "anomalies.html",
    "/logs": "logs.html",
    "/users": "users.html",
    "/users/create": "create_user.html",
    "/stores": "stores.html",
    "/index.html": "index.html",
    "/login": "login.html",
    "/setup": "setup.html",
    "/register": "register.html",
}

# 页面文件 -> 需要的最低角色
PAGE_ROLE = {
    "users.html": "admin",
    "create_user.html": "admin",
    "stores.html": "admin",
}

# 侧边栏"用户管理"菜单组标记块（nav_user_mgmt）: member 会话输出页面前剔除,
# 导航结构不向未授权角色输出; 页面级访问控制仍由 _page_access 兜底
ADMIN_NAV_BLOCK_RE = re.compile(
    rb"<!--admin-nav-start-->.*?<!--admin-nav-end-->", re.S)

SESSION_COOKIE = "wxr_session"

# 店铺数据清理范围: 全部业务数据文件（12 类），但保留 settings（系统设置不属于业务数据）
CLEARABLE_DATA_FILES = tuple(name for name in BUSINESS_FILES if name != "settings")

DELETE_TXN_DIR_PREFIX = "delete_txn_"
DELETE_TXN_STATE_FILE = "txn.json"
TOMBSTONE_FILE = "deleted_store_tombstones.json"

STATIC_MIME = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".png": "image/png",
    ".svg": "image/svg+xml",
    ".ico": "image/x-icon",
}


class ApiError(Exception):
    def __init__(self, message, code=400, http_status=400):
        super().__init__(message)
        self.code = code
        self.http_status = http_status


class ServerStartupError(Exception):
    """服务启动失败（端口占用/无法绑定等），message 为面向用户的中文提示"""

    def __init__(self, message, host=None, port=None, occupied_by_self=False):
        super().__init__(message)
        self.host = host
        self.port = port
        self.occupied_by_self = occupied_by_self


class ReconcileServer:
    def __init__(self, host=None, port=None, data_dir=None, backup_dir=None,
                 log_dir=None, register_rate_window=None, register_rate_max=None):
        self.host = host or config.HOST
        self.port = port if port is not None else config.PORT
        # 数据根目录：用户/店铺/会话等全局文件；业务数据在 data/stores/{store_id}/
        self.data_root = data_dir or config.DATA_DIR
        os.makedirs(self.data_root, exist_ok=True)
        self.stores_root = os.path.join(self.data_root, "stores")
        os.makedirs(self.stores_root, exist_ok=True)
        # 全局安全日志目录（与业务操作日志隔离）
        self.security_log_dir = log_dir or os.path.join(self.data_root, "security")
        self.user_manager = UserManager(self.data_root)
        self.sessions = SessionStore(os.path.join(self.data_root, "sessions.json"))
        self.security_log = SecurityLog(self.security_log_dir)
        self.rate_limiter = LoginRateLimiter()
        # 注册限流: 默认 60 分钟 / 5 次（成败均计）; 参数化便于测试注入宽松阈值
        self.register_limiter = RegistrationRateLimiter(
            window_seconds=register_rate_window or REGISTER_WINDOW_SECONDS,
            max_requests=register_rate_max or REGISTER_MAX_REQUESTS)
        self.registry = StoreServiceRegistry(self.stores_root)
        # 全局备份目录: 店铺数据清理/删除账号前的整店快照存放处（独立于店铺自身 backups/）
        self.backup_dir = backup_dir or config.BACKUP_DIR
        os.makedirs(self.backup_dir, exist_ok=True)
        self.started_at = now_iso()
        self.httpd = None

    # ---------- 认证上下文 ----------
    def _session_from_cookie(self, handler):
        """解析 Cookie 中的会话令牌 -> (token, session) 或 (None, None)"""
        raw = handler.headers.get("Cookie") or ""
        try:
            cookie = SimpleCookie()
            cookie.load(raw)
            morsel = cookie.get(SESSION_COOKIE)
            token = morsel.value if morsel else None
        except Exception:
            token = None
        if not token:
            return None, None
        return token, self.sessions.get(token)

    def _resolve_session(self, handler):
        """校验会话有效性: 存在 + 未过期 + 用户启用 + 密码版本一致 + 店铺在授权范围

        返回 (token, session, user); 无效时抛出 401。
        """
        token, session = self._session_from_cookie(handler)
        if session is None:
            raise ApiError("未登录或会话已失效，请重新登录",
                           code=1001, http_status=401)
        user = self.user_manager.get_user(session["user_id"])
        if user is None or user.get("status") != "active":
            raise ApiError("账号不可用或已停用", code=1001, http_status=401)
        if session.get("password_version", 1) != user.get("password_version", 1):
            raise ApiError("密码已变更，请重新登录", code=1001, http_status=401)
        store_id = session.get("store_id")
        if user.get("role") == "admin":
            # 管理员会话 store_id 为空（无店）合法; 非空时仍需在授权范围内
            if store_id and store_id not in (user.get("store_ids") or []):
                raise ApiError("店铺授权已变更，请重新登录",
                               code=1001, http_status=401)
        elif store_id not in (user.get("store_ids") or []):
            raise ApiError("店铺授权已变更，请重新登录", code=1001, http_status=401)
        return token, session, user

    def _services(self, handler):
        """从已验证会话获取当前店铺的全套服务; 未登录抛 401

        管理员一律返回 None: 业务范围经 _query_scope 解析, 不依赖本人 bundle。
        """
        _, session, user = self._resolve_session(handler)
        if user.get("role") == "admin":
            return None
        return self.registry.get(session["store_id"])

    # ---------- 店铺查询范围（多店铺隔离核心） ----------
    class _Scope:
        """查询范围: targets=[(bundle, store_id, store_name)]; multi=全部店铺聚合"""

        def __init__(self, targets, multi):
            self.targets = targets
            self.multi = multi

        @property
        def bundle(self):
            return self.targets[0][0]

        @property
        def store_id(self):
            return self.targets[0][1]

        @property
        def store_name(self):
            return self.targets[0][2]

    def _stores_of_member(self, member_id):
        """member 参数唯一口径: 校验目标成员身份并返回其有效店铺列表"""
        target = self.user_manager.get_user(member_id)
        if target is None:
            raise ApiError(f"目标用户不存在: {member_id}",
                           code=404, http_status=404)
        if target.get("role") == "admin":
            raise ApiError("管理员不拥有业务店铺，无法按成员筛选",
                           code=400, http_status=400)
        return [s for s in self.user_manager.member_stores()
                if s.get("owner_user_id") == member_id]

    def _reject_member_filter(self, handler, user, raw_member):
        """member 会话携带 member 参数（查询或 body）→ 403 + access_denied"""
        if not raw_member:
            return
        self.security_log.append(
            "access_denied", username=user.get("username"),
            user_id=user.get("user_id"), ip=handler.client_address[0],
            result="denied",
            detail={"reason": "member_filter_forbidden",
                    "requested_member": raw_member,
                    "path": handler.path})
        raise ApiError("无权使用成员筛选", code=1003, http_status=403)

    def _member_active_stores(self, user):
        """member 本人全部有效店铺（member_stores 口径: 含 status=active 过滤）"""
        uid = user.get("user_id")
        return [s for s in self.user_manager.member_stores()
                if s.get("owner_user_id") == uid]

    def _member_store_target(self, handler, user, store_id, reason, message):
        """member 写/上传目标店铺三态校验 → (bundle, store_id, store_name)

        - store_id 不在本人 store_ids 内（他店/不存在/已转走）→ 403 + 安全日志
        - store_id 在本人 store_ids 内但店铺 disabled/archived → 400
        - 本人有效店铺 → 放行（member_stores 口径, 不自行过滤）
        """
        if store_id not in (user.get("store_ids") or []):
            self.security_log.append(
                "access_denied", username=user.get("username"),
                user_id=user.get("user_id"), ip=handler.client_address[0],
                result="denied",
                detail={"reason": reason, "requested_store": store_id,
                        "path": handler.path})
            raise ApiError(message, code=1003, http_status=403)
        store = self.user_manager.get_store(store_id) or {}
        if store.get("status", "active") != "active":
            status = store.get("status")
            raise ApiError("该店铺已停用，暂时无法操作" if status == "disabled"
                           else "该店铺已归档，暂时无法操作",
                           code=400, http_status=400)
        return (self.registry.get(store_id), store_id,
                store.get("name") or store_id)

    def _query_scope(self, handler, parsed):
        """解析业务查询的店铺范围（绝不信任请求中的身份字段）

        - 普通成员（一对多归属三态）:
          * 聚合态: store_id 缺省或 "all" → 本人全部有效店铺聚合
            （单店成员退化为单店直查, 响应结构与既有行为零变化）;
            无任何有效店铺 → 401
          * 单店态: store_id 在本人 store_ids 内且店铺 active → 单店放行
            （本人名下其他店铺不再被 403 误伤）
          * 拒绝态: 本人 store_ids 内但店铺 disabled/archived → 400;
            他店/不存在/已删除 → 403 + store_scope_forbidden 安全日志
          * member 参数一律 403
        - 管理员: 不传/传 all → 全部成员店铺聚合; member=<user_id> 先收窄为该
          成员店铺; 传具体 store_id → 校验后单店（与 member 同带时须归属该成员）
        """
        _, session, user = self._resolve_session(handler)
        own = session["store_id"]
        qs = parse_qs(parsed.query) if parsed else {}
        req_store = qs.get("store_id", [None])[0] or None
        req_member = qs.get("member", [None])[0] or None
        if user.get("role") != "admin":
            self._reject_member_filter(handler, user, req_member)
            if req_store in (None, "", "all"):
                stores = self._member_active_stores(user)
                if not stores:
                    raise ApiError("账号未绑定店铺，请联系管理员",
                                   code=1001, http_status=401)
                if len(stores) == 1:
                    s = stores[0]
                    return ReconcileServer._Scope(
                        [(self.registry.get(s["store_id"]), s["store_id"],
                          s.get("name") or s["store_id"])], False)
                targets = [(self.registry.get(s["store_id"]), s["store_id"],
                            s.get("name") or s["store_id"]) for s in stores]
                return ReconcileServer._Scope(targets, True)
            if req_store not in (user.get("store_ids") or []):
                self.security_log.append(
                    "access_denied", username=user.get("username"),
                    user_id=user.get("user_id"), ip=handler.client_address[0],
                    result="denied",
                    detail={"reason": "store_scope_forbidden",
                            "requested_store": req_store,
                            "path": handler.path})
                raise ApiError("无权访问该店铺数据", code=1003, http_status=403)
            store = self.user_manager.get_store(req_store) or {}
            if store.get("status", "active") != "active":
                status = store.get("status")
                raise ApiError(
                    "该店铺已停用，暂时无法访问" if status == "disabled"
                    else "该店铺已归档，暂时无法访问",
                    code=400, http_status=400)
            return ReconcileServer._Scope(
                [(self.registry.get(req_store), req_store,
                  store.get("name") or req_store)], False)
        stores = (self._stores_of_member(req_member)
                  if req_member else self.user_manager.member_stores())
        if req_store in (None, "", "all"):
            targets = [(self.registry.get(s["store_id"]), s["store_id"],
                        s.get("name") or s["store_id"]) for s in stores]
            return ReconcileServer._Scope(targets, True)
        store = self.user_manager.get_store(req_store)
        if store is None:
            raise ApiError(f"店铺不存在: {req_store}", code=404, http_status=404)
        if req_member and store.get("owner_user_id") != req_member:
            raise ApiError("目标店铺不属于该成员", code=400, http_status=400)
        return ReconcileServer._Scope(
            [(self.registry.get(req_store), req_store,
              store.get("name") or req_store)], False)

    def _write_target(self, handler, body):
        """写操作目标店铺: 管理员可经 body.target_store_id/member 指定; 成员三态

        成员请求: target_store_id/store_id 指向他店或 member 键 → 一律 403;
        携带本人名下其他有效店铺 → 放行该店（多店写入）; 缺省/"all" → 会话默认店铺。
        管理员未指定目标: 有店（存量过渡期）默认本店, 无店 → 400 要求显式选择;
        仅携带 member（未选店）→ 定位该成员唯一店铺（多店/无店 → 400 要求选店）。
        """
        _, session, user = self._resolve_session(handler)
        own = session["store_id"]
        target = None
        member = None
        if isinstance(body, dict):
            target = body.get("target_store_id") or body.get("store_id") or None
            member = body.get("member") or None
        if user.get("role") != "admin":
            self._reject_member_filter(handler, user, member)
            if target in (None, "", "all"):
                return self.registry.get(own)
            return self._member_store_target(
                handler, user, target,
                "write_store_forbidden", "无权操作该店铺")[0]
        if target in (None, "", "all"):
            if member:
                stores = self._stores_of_member(member)
                if len(stores) != 1:
                    raise ApiError("请先选择目标店铺",
                                   code=400, http_status=400)
                return self.registry.get(stores[0]["store_id"])
            if own:
                return self.registry.get(own)
            raise ApiError("请先选择目标店铺", code=400, http_status=400)
        store = self.user_manager.get_store(target)
        if store is None:
            raise ApiError(f"店铺不存在: {target}", code=404, http_status=404)
        if member and store.get("owner_user_id") != member:
            raise ApiError("目标店铺不属于该成员", code=400, http_status=400)
        return self.registry.get(target)

    def _find_batch_store(self, scope, batch_id):
        """在范围内逐店查找导入批次 → (bundle, store_id, store_name)"""
        for bundle, sid, sname in scope.targets:
            try:
                bundle.importer.get_batch(batch_id)
                return bundle, sid, sname
            except ImportServiceError:
                continue
        raise ApiError(f"导入批次不存在: {batch_id}", code=404, http_status=404)

    # ---------- 多店铺聚合 ----------
    @staticmethod
    def _try_decimal(value):
        """含小数点且可解析为 Decimal 的字符串（金额格式）→ Decimal; 否则 None"""
        if not isinstance(value, str) or "." not in value:
            return None
        try:
            from decimal import Decimal
            return Decimal(value)
        except Exception:
            return None

    @staticmethod
    def _merge_dicts(base, extra):
        """通用字典合并: 数值相加 / 列表拼接 / 字典递归 / 金额字符串相加 / 其他取首值"""
        for k, v in extra.items():
            if k not in base:
                base[k] = v
            elif isinstance(base[k], (int, float)) and isinstance(v, (int, float)) \
                    and not isinstance(base[k], bool) and not isinstance(v, bool):
                base[k] = base[k] + v
            elif isinstance(base[k], list) and isinstance(v, list):
                base[k] = base[k] + v
            elif isinstance(base[k], dict) and isinstance(v, dict):
                ReconcileServer._merge_dicts(base[k], v)
            elif isinstance(base[k], str) and isinstance(v, str):
                a = ReconcileServer._try_decimal(base[k])
                b = ReconcileServer._try_decimal(v)
                if a is not None and b is not None:
                    base[k] = str(a + b)

    def _merge_lists(self, parts, records_key, page, size):
        """跨店铺分页合并: 每条记录附 store_id/store_name 后整体重分页"""
        merged = []
        for (sid, sname), recs in parts:
            for r in recs:
                if isinstance(r, dict):
                    r = dict(r)
                    r.setdefault("store_id", sid)
                    r.setdefault("store_name", sname)
                merged.append(r)
        total = len(merged)
        return {"total": total, "page": page, "size": size,
                "pages": max(1, (total + size - 1) // size),
                records_key: merged[(page - 1) * size:page * size]}

    def _fetch_all_pages(self, data_fn, records_key):
        """data_fn(page, size) → 分页 dict; 逐页拉取全量 records 与首页元数据"""
        out, page, first = [], 1, None
        while True:
            d = data_fn(page, 100)
            if first is None:
                first = d
            out.extend(d.get(records_key) or [])
            if page >= d.get("pages", 1):
                break
            page += 1
        return out, first

    def _scoped_list(self, handler, parsed, data_fn, records_key):
        """列表类查询统一分发: 单店直查; 全部店铺逐店拉全量合并重分页

        多店时各店返回的非分页键（statistics 等）按通用规则合并。
        """
        scope = self._query_scope(handler, parsed)
        qs = parse_qs(parsed.query)
        page, size = self._page_params(qs)
        if not scope.multi:
            result = data_fn(scope.bundle, scope.store_id, scope.store_name,
                             page, size)
            for r in result.get(records_key) or []:
                if isinstance(r, dict):
                    r.setdefault("store_id", scope.store_id)
                    r.setdefault("store_name", scope.store_name)
            result.setdefault("store_scope", scope.store_name)
            result["time_range"] = self._time_meta(qs)
            return handler._send(ApiResponse.ok(result))
        merged, extras = [], {}
        skip = {records_key, "total", "page", "size", "pages", "time_range"}
        for b, sid, sname in scope.targets:
            recs, first = self._fetch_all_pages(
                lambda p, s, _b=b, _sid=sid, _sn=sname:
                data_fn(_b, _sid, _sn, p, s), records_key)
            for r in recs:
                if isinstance(r, dict):
                    r = dict(r)
                    r.setdefault("store_id", sid)
                    r.setdefault("store_name", sname)
                merged.append(r)
            if first:
                ReconcileServer._merge_dicts(
                    extras, {k: v for k, v in first.items() if k not in skip})
        total = len(merged)
        result = {"total": total, "page": page, "size": size,
                  "pages": max(1, (total + size - 1) // size),
                  records_key: merged[(page - 1) * size:page * size]}
        result.update(extras)
        result["store_scope"] = "all"
        result["store_count"] = len(scope.targets)
        result["time_range"] = self._time_meta(qs)
        handler._send(ApiResponse.ok(result))

    def _scoped_options(self, handler, parsed, fn):
        """选项/统计类查询统一分发: 多店时数值求和、列表拼接、字典递归合并"""
        scope = self._query_scope(handler, parsed)
        if not scope.multi:
            return handler._send(ApiResponse.ok(fn(scope.bundle)))
        merged = {}
        for b, _sid, _sn in scope.targets:
            ReconcileServer._merge_dicts(merged, fn(b))
        handler._send(ApiResponse.ok(merged))

    def _scoped_find(self, handler, parsed, fn, message=None):
        """详情类查询统一分发: 多店时逐店查找, 未找到抛 404"""
        scope = self._query_scope(handler, parsed)
        last_err = None
        for b, sid, sname in scope.targets:
            try:
                data = fn(b)
                if isinstance(data, dict):
                    data.setdefault("store_id", sid)
                    data.setdefault("store_name", sname)
                return handler._send(ApiResponse.ok(data))
            except (KeyError, ImportServiceError) as e:
                last_err = e
                continue
        raise ApiError(message or f"记录不存在: {last_err}",
                       code=404, http_status=404)

    @staticmethod
    def _merge_excel_buffers(buffers):
        """合并多个同构 Excel: 首文件保留表头, 其余逐行追加数据"""
        import io
        from openpyxl import load_workbook
        if len(buffers) == 1:
            return buffers[0]
        wb = load_workbook(io.BytesIO(buffers[0]))
        ws = wb.active
        for buf in buffers[1:]:
            wb2 = load_workbook(io.BytesIO(buf))
            for row in wb2.active.iter_rows(min_row=2, values_only=True):
                ws.append(list(row))
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

    def _scoped_export(self, handler, parsed, export_fn, post_merge=None):
        """导出类统一分发: 单店直出; 全部店铺逐店导出后合并（无数据店铺跳过）

        post_merge: 可选钩子，接收合并前各店 bytes 列表、返回最终 bytes
        （利润导出用于在合并文件追加唯一聚合合计行）；缺省走普通合并。
        """
        scope = self._query_scope(handler, parsed)
        if not scope.multi:
            try:
                filename, content = export_fn(scope.bundle)
            except ValueError as e:
                raise ApiError(str(e), code=400, http_status=400)
            return handler._send_bytes(EXCEL_MIME, content, filename)
        filename, contents = None, []
        for b, _sid, _sn in scope.targets:
            try:
                fn, content = export_fn(b)
            except ValueError:
                continue
            filename = filename or fn
            contents.append(content)
        if not contents:
            raise ApiError("当前筛选无数据可导出", code=400, http_status=400)
        content = post_merge(contents) if post_merge \
            else self._merge_excel_buffers(contents)
        handler._send_bytes(EXCEL_MIME, content, filename)

    def _append_profit_summary_row(self, merged_bytes, stats_list):
        """多店导出合并文件追加唯一聚合合计行

        聚合数值 = 各店 statistics 依次经 _merge_dicts 合并
        （与页面聚合统计卡同一合并函数，保证数学恒等）；
        金额以 Decimal 写入并设 0.00 格式，防止存为字符串。
        """
        import io
        from openpyxl import load_workbook
        merged = {}
        for s in stats_list:
            ReconcileServer._merge_dicts(merged, s)
        wb = load_workbook(io.BytesIO(merged_bytes))
        ws = wb.active
        header = list(next(ws.iter_rows(min_row=1, max_row=1,
                                        values_only=True)))
        col = {name: i for i, name in enumerate(header)}
        notes = []
        pc_missing = merged.get("product_cost_missing_count")
        cf_missing = merged.get("courier_fee_missing_count")
        if isinstance(pc_missing, int) and pc_missing > 0:
            notes.append(f"产品成本待核对 {pc_missing} 单")
        if isinstance(cf_missing, int) and cf_missing > 0:
            notes.append(f"快递费合计待核对 {cf_missing} 单")
        row = [None] * len(header)
        row[col["订单号"]] = "合计"
        row[col["下单时间"]] = f"共 {merged.get('total', 0)} 单"
        if notes:
            row[col["状态原因"]] = "；".join(notes)
        for name, key in (("产品成本", "product_cost_total"),
                          ("快递费合计", "courier_fee_total"),
                          ("综合成本", "composite_cost_total")):
            v = ReconcileServer._try_decimal(merged.get(key))
            row[col[name]] = v
        ws.append(row)
        summary_idx = ws.max_row
        for name in ("产品成本", "快递费合计", "综合成本"):
            c = ws.cell(row=summary_idx, column=col[name] + 1)
            if c.value is not None:
                c.number_format = "0.00"
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

    def _write_targets(self, handler, body):
        """批量写操作目标店铺: 管理员传 all → 全部店铺; member 键收窄为该成员店铺

        成员传 "all"/缺省 → 本人全部有效店铺（批量对账/重算逐店执行）;
        携带本人名下具体店铺 → 单店执行; 他店 → 403。
        """
        _, session, user = self._resolve_session(handler)
        own = session["store_id"]
        target = None
        member = None
        if isinstance(body, dict):
            target = body.get("target_store_id") or body.get("store_id") or None
            member = body.get("member") or None
        if user.get("role") != "admin":
            self._reject_member_filter(handler, user, member)
            if target in (None, "", "all"):
                stores = self._member_active_stores(user)
                if not stores:
                    raise ApiError("账号未绑定店铺，请联系管理员",
                                   code=1001, http_status=401)
                return [(self.registry.get(s["store_id"]), s["store_id"],
                         s.get("name") or s["store_id"]) for s in stores]
            return [self._member_store_target(
                handler, user, target,
                "write_store_forbidden", "无权操作该店铺")]

        def wrap(store_ids):
            out = []
            for sid in store_ids:
                s = self.user_manager.get_store(sid) or {}
                out.append((self.registry.get(sid), sid,
                            s.get("name") or sid))
            return out

        if target == "all":
            if member:
                stores = self._stores_of_member(member)
                if not stores:
                    raise ApiError("该成员暂无有效店铺", code=400, http_status=400)
                return wrap([s["store_id"] for s in stores])
            return wrap([s["store_id"]
                         for s in self.user_manager.member_stores()])
        if not target and member:
            stores = self._stores_of_member(member)
            if not stores:
                raise ApiError("该成员暂无有效店铺", code=400, http_status=400)
            return wrap([s["store_id"] for s in stores])
        if target == own and own:
            store = self.user_manager.get_store(own) or {}
            return [(self.registry.get(own), own, store.get("name") or own)]
        if not target:
            if not own:
                raise ApiError("请先选择目标店铺，或传入 all 操作全部成员店铺",
                               code=400, http_status=400)
            store = self.user_manager.get_store(own) or {}
            return [(self.registry.get(own), own, store.get("name") or own)]
        store = self.user_manager.get_store(target)
        if store is None:
            raise ApiError(f"店铺不存在: {target}", code=404, http_status=404)
        if member and store.get("owner_user_id") != member:
            raise ApiError("目标店铺不属于该成员", code=400, http_status=400)
        return [(self.registry.get(target), target,
                 store.get("name") or target)]

    def _upload_target(self, handler, parsed):
        """上传目标店铺: 管理员必须显式指定目标且目标须为有效成员店铺; 成员强制本店

        管理员上传的数据不得默认进入管理员或当前登录用户名下（spec 5.3.2 规则 9）。
        """
        _, session, user = self._resolve_session(handler)
        own = session["store_id"]
        qs = parse_qs(parsed.query)
        target = (handler.headers.get("X-Target-Store")
                  or qs.get("store_id", [None])[0] or None)
        if user.get("role") != "admin":
            if target in (None, "", "all"):
                return self.registry.get(own)
            return self._member_store_target(
                handler, user, target,
                "upload_store_forbidden", "无权上传到该店铺")[0]
        if not target or target == "all":
            raise ApiError("请先选择目标成员和目标店铺", code=400, http_status=400)
        store = self.user_manager.get_store(target)
        if store is None:
            raise ApiError(f"店铺不存在: {target}", code=404, http_status=404)
        owner = (self.user_manager.get_user(store.get("owner_user_id"))
                 if store.get("owner_user_id") else None)
        if (owner is None or owner.get("role") != "member"
                or owner.get("status") != "active"):
            raise ApiError("目标店铺不属于有效成员，请先为其创建或分配店铺",
                           code=400, http_status=400)
        return self.registry.get(target)

    def _import_target(self, handler, raw_store_id, action):
        """导入目标店铺校验: store_id 一律必填 + 会话身份权威复核

        返回 (bundle, store_id, store_name, owner_user_dict)。
        - 缺失/空串/"all" → 400 请先选择目标店铺（文件不解析、不暂存、不建批次）
        - member: 复用 _member_store_target 三态（他店 403 + import_store_forbidden
          安全日志; 停用/归档 400），通过后由店铺归属反查 owner（member 即本人）
        - admin: 店铺不存在 404; 停用/归档 400（补齐旧 _upload_target 缺失的
          状态校验）; owner 非有效成员 400
        - action 仅用于错误上下文区分（upload/confirm）
        """
        if raw_store_id in (None, "", "all"):
            raise ApiError("请先选择目标店铺", code=400, http_status=400)
        _, session, user = self._resolve_session(handler)
        if user.get("role") != "admin":
            bundle, store_id, store_name = self._member_store_target(
                handler, user, raw_store_id,
                "import_store_forbidden", "无权上传到该店铺")
            store = self.user_manager.get_store(store_id) or {}
            owner = (self.user_manager.get_user(store.get("owner_user_id"))
                     if store.get("owner_user_id") else None)
            return (bundle, store_id, store_name, owner)
        store = self.user_manager.get_store(raw_store_id)
        if store is None:
            raise ApiError(f"店铺不存在: {raw_store_id}", code=404, http_status=404)
        if store.get("status", "active") != "active":
            status = store.get("status")
            raise ApiError("该店铺已停用，暂时无法导入" if status == "disabled"
                           else "该店铺已归档，暂时无法导入",
                           code=400, http_status=400)
        owner = (self.user_manager.get_user(store.get("owner_user_id"))
                 if store.get("owner_user_id") else None)
        if (owner is None or owner.get("role") != "member"
                or owner.get("status") != "active"):
            raise ApiError("目标店铺不属于有效成员，请先为其创建或分配店铺",
                           code=400, http_status=400)
        return (self.registry.get(raw_store_id), raw_store_id,
                store.get("name") or raw_store_id, owner)

    def _require_admin(self, handler):
        _, session, user = self._resolve_session(handler)
        if user.get("role") != "admin":
            self.security_log.append(
                "access_denied", username=user.get("username"),
                user_id=user.get("user_id"), ip=handler.client_address[0],
                result="denied", detail={"reason": "admin_only",
                                         "path": handler.path})
            raise ApiError("需要管理员权限", code=1003, http_status=403)
        return session, user

    @staticmethod
    def _check_csrf(handler):
        """写请求 CSRF 防护: 必须携带浏览器跨站表单无法伪造的自定义头"""
        if (handler.headers.get("X-Requested-With") or "") != "XMLHttpRequest":
            raise ApiError("请求缺少防跨站校验头，请从系统页面发起操作",
                           code=1002, http_status=403)

    # ---------- 生命周期 ----------
    def _log_lifecycle(self, event, detail=None, result="success"):
        """系统级生命周期事件容错写入: 失败时告警到 stderr, 不改变主流程结果"""
        try:
            self.security_log.append(event, detail=detail, result=result)
        except (OSError, Exception) as e:
            print(f"[警告] 生命周期日志写入失败（事件: {event}）: {e!r}",
                  file=sys.stderr)

    def start(self, blocking=True):
        if self.httpd is not None:
            raise ServerStartupError(
                f"服务已在 {self.host}:{self.port} 运行（本实例），请勿重复调用 start()",
                host=self.host, port=self.port, occupied_by_self=True)
        self._log_lifecycle(
            "server.init",
            detail={"initialized": self.user_manager.is_initialized(),
                    "store_count": len(os.listdir(self.stores_root))})
        try:
            migrated = self.user_manager.migrate_archived_stores()
            if migrated:
                self.security_log.append(
                    "store.archive_migration", username="system",
                    detail={"migrated": migrated, "count": len(migrated)},
                    result="success")
        except Exception as e:
            self._log_lifecycle(
                "store.archive_migration",
                detail={"error": repr(e)}, result="error")
        self._recover_pending_delete_txns()
        try:
            self.httpd = ThreadingHTTPServer(
                (self.host, self.port), self._make_handler())
        except OSError as e:
            # ThreadingHTTPServer 构造失败时内部已 server_close()，不遗留 socket
            self.httpd = None
            message, by_self = self._startup_error_message(e)
            self._log_lifecycle(
                "server.start",
                detail={"host": self.host, "port": self.port,
                        "error": repr(e), "errno": e.errno},
                result="error")
            raise ServerStartupError(
                message, host=self.host, port=self.port,
                occupied_by_self=by_self) from e
        self.httpd.app = self
        self.port = self.httpd.server_address[1]
        self.started_at = now_iso()
        self._log_lifecycle("server.start", detail={"host": self.host, "port": self.port})
        if blocking:
            try:
                self.httpd.serve_forever()
            except KeyboardInterrupt:
                pass
            finally:
                self.shutdown()
        else:
            import threading
            t = threading.Thread(target=self.httpd.serve_forever, daemon=True)
            t.start()
            return t

    def _startup_error_message(self, error):
        """生成启动失败的中文提示；区分本系统实例占用与其他进程占用"""
        addr = f"{self.host}:{self.port}"
        if error.errno == errno.EADDRINUSE:
            existing = self._probe_existing_service()
            if existing:
                return (
                    f"端口 {addr} 已被本系统正在运行的服务实例占用"
                    f"（服务：{existing.get('service')}，版本：{existing.get('version')}，"
                    f"启动时间：{existing.get('started_at') or '未知'}）。\n"
                    f"请直接访问 http://{addr}/ 使用现有实例；"
                    f"如需重启，请先停止原实例（在原终端按 Ctrl+C，"
                    f"或执行 lsof -nP -iTCP:{self.port} -sTCP:LISTEN 查看进程后手动停止），"
                    f"再重新启动。本次启动已取消，未影响正在运行的服务。"), True
            return (
                f"端口 {addr} 已被其他进程占用，无法启动服务。\n"
                f"排查命令：lsof -nP -iTCP:{self.port} -sTCP:LISTEN\n"
                f"请确认占用进程可安全停止后重试，或使用 --port 参数指定其他端口"
                f"（例如：python main.py --port 8766）。"), False
        hint = {errno.EACCES: "端口无权限（可能需要 1024 以上端口）",
                errno.EADDRNOTAVAIL: "监听地址不可用（请检查 --host 参数）"}.get(error.errno)
        text = f"服务绑定 {addr} 失败：{error}"
        if hint:
            text += f"（{hint}）"
        return text, False

    def _probe_existing_service(self, timeout=2.0):
        """探测占用端口的是否为本系统已启动的服务实例，返回 health 数据或 None"""
        import urllib.request
        url = f"http://{self.host}:{self.port}/api/health"
        try:
            with urllib.request.urlopen(url, timeout=timeout) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except Exception:
            return None
        payload = data.get("data") or {}
        if data.get("code") == 0 and SERVICE_NAME in str(payload.get("service", "")):
            return payload
        return None

    def shutdown(self):
        if self.httpd is not None:
            self._log_lifecycle("server.stop", detail={"port": self.port})
            self.httpd.shutdown()
            self.httpd.server_close()
            self.httpd = None

    def init_all_store_files(self):
        """为全部已注册店铺初始化业务数据文件(幂等)

        经用户注册表收集全部 store_id 并去重, 逐店铺经 registry.get()
        完成目录创建与业务文件 ensure_all(); 单店铺初始化异常向上抛出。
        返回 {"store_count": int, "created": int, "repaired": int}
        """
        store_ids = []
        for user in self.user_manager.list_users():
            for store_id in user.get("store_ids") or []:
                if store_id not in store_ids:
                    store_ids.append(store_id)
        report = {"store_count": len(store_ids), "created": 0, "repaired": 0}
        for store_id in store_ids:
            # registry.get() 首次创建时会内部 ensure_all 并丢弃报告,
            # 故先用独立实例取准确报告, 再经 registry.get() 预热缓存(幂等)
            store_report = StoreServices(store_id, self.stores_root).ensure_all()
            self.registry.get(store_id)
            report["created"] += len(store_report.get("created") or [])
            report["repaired"] += len(store_report.get("repaired") or [])
        return report

    def _make_handler(self):
        server_app = self

        class Handler(BaseHTTPRequestHandler):
            server_version = f"WxReconcile/{VERSION}"

            def do_GET(self):
                self._dispatch("GET")

            def do_POST(self):
                self._dispatch("POST")

            def _dispatch(self, method):
                parsed = urlparse(self.path)
                app = server_app
                try:
                    # 统一 CSRF 校验: 除登录/初始化外的全部写请求
                    if method == "POST" and parsed.path not in PUBLIC_APIS_POST:
                        app._check_csrf(self)
                    if method == "POST" and parsed.path in UPLOAD_PATHS:
                        app.route_upload(self, parsed)
                    elif method == "GET":
                        app.route_get(self, parsed)
                    else:
                        body = self._read_body()
                        app.route_post(self, parsed, body)
                except ImportServiceError as e:
                    self._send(ApiResponse.error(str(e), code=e.code, data=e.data),
                               e.http_status)
                except ReconcileServiceError as e:
                    self._send(ApiResponse.error(str(e), code=e.code, data=e.data),
                               e.http_status)
                except ProfitServiceError as e:
                    self._send(ApiResponse.error(str(e), code=e.code, data=e.data),
                               e.http_status)
                except AuthError as e:
                    self._send(ApiResponse.error(str(e), code=e.code), e.http_status)
                except TimeRangeError as e:
                    self._send(ApiResponse.error(f"时间筛选参数错误: {e}", code=400),
                               400)
                except ApiError as e:
                    self._send(ApiResponse.error(str(e), code=e.code), e.http_status)
                except RepositoryError as e:
                    self._send(ApiResponse.error(f"数据错误: {e}", code=1001), 400)
                except json.JSONDecodeError:
                    self._send(ApiResponse.error("请求体不是合法的 JSON", code=400), 400)
                except Exception as e:
                    # 完整 traceback 仅落服务端日志（stderr 经 start.sh
                    # 重定向至 logs/server.out），便于检索报错文件与行号；
                    # 响应体保持简短消息，不外泄堆栈与磁盘路径。
                    traceback.print_exc()
                    app.security_log.append(
                        "api.error",
                        detail={"path": self.path, "error": repr(e)},
                        result="error",
                    )
                    self._send(ApiResponse.error(f"服务器内部错误: {e}", code=500), 500)

            def _read_body(self):
                length = int(self.headers.get("Content-Length") or 0)
                raw = self.rfile.read(length) if length > 0 else b""
                if not raw:
                    return {}
                body = json.loads(raw.decode("utf-8"))
                if not isinstance(body, dict):
                    raise ApiError("请求体必须是 JSON 对象", code=400, http_status=400)
                return body

            def _send(self, resp, status=200):
                payload = ApiResponse.to_bytes(resp)
                self.send_response(status)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)

            def _redirect(self, location):
                self.send_response(302)
                self.send_header("Location", location)
                self.send_header("Content-Length", "0")
                self.end_headers()

            def _send_bytes(self, content_type, data, download_name=None):
                self.send_response(200)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Length", str(len(data)))
                # 页面资源禁止浏览器缓存：避免代码更新后浏览器仍显示旧页面
                self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
                if download_name:
                    self.send_header(
                        "Content-Disposition",
                        f"attachment; filename*=UTF-8''{quote(download_name)}")
                self.end_headers()
                self.wfile.write(data)

            def _send_with_cookie(self, resp, cookie_value, status=200):
                payload = ApiResponse.to_bytes(resp)
                self.send_response(status)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(payload)))
                if cookie_value is None:
                    self.send_header(
                        "Set-Cookie",
                        f"{SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0")
                else:
                    self.send_header(
                        "Set-Cookie",
                        f"{SESSION_COOKIE}={cookie_value}; Path=/; HttpOnly; "
                        f"SameSite=Lax; Max-Age={12 * 3600}")
                self.end_headers()
                self.wfile.write(payload)

            def log_message(self, fmt, *args):
                pass

        return Handler

    # ---------- 路由 ----------
    def _page_access(self, handler, page_file):
        """页面访问控制: 返回 True 表示放行; 否则已发送 302 重定向

        规则:
        - 系统未初始化: 除 /setup 与登录页外全部引导到 /setup
        - 已初始化: 先按统一标准校验会话（存在 + 用户 active + password_version 一致）
          - 未登录访问 /login、/register: 放行（公开页必须可直接访问, 严禁回跳形成循环）
          - 未登录访问 /setup: 引导到 /login（系统已初始化, 无需再初始化）
          - 未登录访问其余页面: 引导到 /login
          - 已登录访问 /login、/register、/setup: 引导回首页
        - users.html 仅管理员可访问
        """
        initialized = self.user_manager.is_initialized()
        is_login = page_file == "login.html"
        is_setup = page_file == "setup.html"
        is_register = page_file == "register.html"
        if not initialized:
            if is_setup or is_login:
                return True
            handler._redirect("/setup")
            return False
        # 已初始化: 统一会话校验（登录页与业务页同一标准, 单点实现避免规则冲突）
        token, session = self._session_from_cookie(handler)
        session_user = None
        if session is not None:
            user = self.user_manager.get_user(session.get("user_id"))
            if (user is not None and user.get("status") == "active"
                    and session.get("password_version", 1) == user.get("password_version", 1)):
                session_user = user
        if is_login or is_setup or is_register:
            if session_user is not None:
                handler._redirect("/")
                return False
            if is_setup:
                handler._redirect("/login")
                return False
            return True
        if session_user is None:
            handler._redirect("/login")
            return False
        need_role = PAGE_ROLE.get(page_file)
        if need_role == "admin" and session_user.get("role") != "admin":
            handler._redirect("/")
            return False
        return True

    def route_get(self, handler, parsed):
        path = parsed.path
        if path in PAGE_ROUTES:
            if not self._page_access(handler, PAGE_ROUTES[path]):
                return None
            return self._handle_static(handler, PAGE_ROUTES[path])
        if path.startswith("/static/"):
            rel = path[len("/static/"):]
            # 直接访问 HTML 静态文件同样受页面访问控制（防止绕过语义路由）
            if rel.endswith(".html") and rel in PAGE_ROUTES.values():
                if not self._page_access(handler, rel):
                    return None
            elif rel.endswith(".html"):
                raise ApiError("未找到页面资源", code=404, http_status=404)
            return self._handle_static(handler, rel)
        if path == "/api/health":
            return self._handle_health(handler)
        if path == "/api/setup/status":
            return self._handle_setup_status(handler)
        if path == "/api/auth/me":
            return self._handle_auth_me(handler)
        if path == "/api/users":
            return self._handle_user_list(handler, parsed)
        m = re.fullmatch(r"/api/users/([A-Za-z0-9_\-]+)/store-data-summary", path)
        if m:
            return self._handle_store_data_summary(handler, m.group(1))
        # ---------- 管理接口: 不依赖业务 bundle（admin 可无店铺） ----------
        if path == "/api/stores":
            return self._handle_stores(handler)
        m = re.fullmatch(r"/api/stores/([A-Za-z0-9_\-]+)", path)
        if m:
            return self._handle_store_detail(handler, m.group(1))
        if path == "/api/admin/all-stores":
            return self._handle_admin_all_stores(handler, parsed)
        if path == "/api/admin/migration/overview":
            return self._handle_admin_migration_overview(handler)
        if path == "/api/admin/migration/model-check":
            return self._handle_migration_model_check(handler)
        # ---------- 以下全部为业务 API: 必须登录 ----------
        bundle = self._services(handler)
        if path == "/api/ad-expenses":
            return self._handle_ad_expenses(handler, parsed, bundle)
        m = re.fullmatch(r"/api/data/([a-z_]+)", path)
        if m:
            return self._handle_get_data(handler, parsed, m.group(1), bundle)
        if path == "/api/logs":
            return self._handle_logs(handler, parsed, bundle)
        if path == "/api/backups":
            return self._handle_list_backups(handler, parsed, bundle)
        if path == "/api/orders":
            return self._handle_orders(handler, parsed, kind="normal", bundle=bundle)
        if path == "/api/orders/export":
            return self._handle_orders_export(handler, parsed, kind="normal", bundle=bundle)
        if path == "/api/refund-orders":
            return self._handle_orders(handler, parsed, kind="refund", bundle=bundle)
        if path == "/api/refund-orders/export":
            return self._handle_orders_export(handler, parsed, kind="refund", bundle=bundle)
        if path == "/api/refund-details":
            return self._handle_refund_details(handler, parsed, bundle)
        if path == "/api/refund-details/export":
            return self._handle_refund_details_export(handler, parsed, bundle)
        if path == "/api/supplier-costs":
            return self._handle_supplier_costs(handler, parsed, bundle)
        if path == "/api/supplier-costs/export":
            return self._handle_supplier_costs_export(handler, parsed, bundle)
        if path == "/api/anomalies":
            return self._handle_anomalies(handler, parsed, bundle)
        if path == "/api/anomalies/export":
            return self._handle_anomalies_export(handler, parsed, bundle)
        if path == "/api/anomalies/options":
            return self._handle_anomalies_options(handler, parsed, bundle)
        if path == "/api/reconciliation":
            return self._handle_reconciliation_list(handler, parsed, bundle)
        if path == "/api/reconciliation/export":
            return self._handle_reconciliation_export(handler, parsed, bundle)
        if path == "/api/reconciliation/options":
            return self._handle_reconciliation_options(handler, parsed, bundle)
        if path == "/api/summary":
            return self._handle_summary(handler, parsed, bundle)
        if path == "/api/profit":
            return self._handle_profit_list(handler, parsed, bundle)
        if path == "/api/profit/export":
            return self._handle_profit_export(handler, parsed, bundle)
        if path == "/api/profit/options":
            return self._handle_profit_options(handler, parsed, bundle)
        m = re.fullmatch(r"/api/orders/([A-Za-z0-9_\-]+)/refunds", path)
        if m:
            return self._handle_order_refunds(handler, m.group(1), bundle)
        m = re.fullmatch(r"/api/orders/([A-Za-z0-9_\-]+)", path)
        if m:
            return self._handle_order_detail(handler, m.group(1), bundle)
        if path == "/api/import/batches":
            return self._handle_import_batches(handler, parsed, bundle)
        m = re.fullmatch(r"/api/import/batches/([A-Za-z0-9_\-]+)/revoke-preview", path)
        if m:
            return self._handle_batch_revoke_preview(handler, parsed, m.group(1))
        m = re.fullmatch(r"/api/import/batches/([A-Za-z0-9_\-]+)/rows", path)
        if m:
            return self._handle_import_batch_rows(handler, parsed, m.group(1), bundle)
        m = re.fullmatch(r"/api/import/batches/([A-Za-z0-9_\-]+)/errors/export", path)
        if m:
            return self._handle_import_errors_export(handler, parsed, m.group(1), bundle)
        m = re.fullmatch(r"/api/import/batches/([A-Za-z0-9_\-]+)", path)
        if m:
            return self._handle_import_batch_detail(handler, parsed, m.group(1), bundle)
        raise ApiError(f"接口不存在: {path}", code=404, http_status=404)

    def route_post(self, handler, parsed, body):
        path = parsed.path
        if path == "/api/auth/login":
            return self._handle_login(handler, body)
        if path == "/api/setup/init":
            return self._handle_setup_init(handler, body)
        if path == "/api/auth/register":
            return self._handle_auth_register(handler, body)
        # ---------- 以下全部为业务/账号 API: 必须登录 ----------
        if path == "/api/auth/logout":
            return self._handle_logout(handler)
        if path == "/api/auth/change-password":
            return self._handle_change_password(handler, body)
        if path == "/api/users":
            return self._handle_user_create(handler, body)
        if path == "/api/stores":
            return self._handle_store_create(handler, body)
        m = re.fullmatch(r"/api/users/([A-Za-z0-9_\-]+)/status", path)
        if m:
            return self._handle_user_status(handler, body, m.group(1))
        m = re.fullmatch(r"/api/users/([A-Za-z0-9_\-]+)/rename", path)
        if m:
            return self._handle_user_rename(handler, body, m.group(1))
        m = re.fullmatch(r"/api/users/([A-Za-z0-9_\-]+)/reset-password", path)
        if m:
            return self._handle_user_reset_password(handler, body, m.group(1))
        m = re.fullmatch(r"/api/users/([A-Za-z0-9_\-]+)/delete", path)
        if m:
            return self._handle_user_delete(handler, body, m.group(1))
        m = re.fullmatch(r"/api/users/([A-Za-z0-9_\-]+)/store-data-clear", path)
        if m:
            return self._handle_store_data_clear(handler, body, m.group(1))
        m = re.fullmatch(r"/api/stores/([A-Za-z0-9_\-]+)/rename", path)
        if m:
            return self._handle_store_rename(handler, body, m.group(1))
        m = re.fullmatch(
            r"/api/stores/([A-Za-z0-9_\-]+)/transfer-ownership", path)
        if m:
            return self._handle_store_transfer_ownership(
                handler, body, m.group(1))
        m = re.fullmatch(
            r"/api/stores/([A-Za-z0-9_\-]+)/mark-unassigned", path)
        if m:
            return self._handle_store_mark_unassigned(
                handler, body, m.group(1))
        m = re.fullmatch(r"/api/stores/([A-Za-z0-9_\-]+)/status", path)
        if m:
            return self._handle_store_status(handler, body, m.group(1))
        m = re.fullmatch(r"/api/stores/([A-Za-z0-9_\-]+)/delete", path)
        if m:
            return self._handle_store_delete(handler, body, m.group(1))
        m = re.fullmatch(r"/api/stores/([A-Za-z0-9_\-]+)/clear-data", path)
        if m:
            return self._handle_store_clear_data(handler, body, m.group(1))
        if path == "/api/admin/migration/model-migrate":
            return self._handle_migration_model_migrate(handler, body)
        m = re.fullmatch(r"/api/import/batches/([A-Za-z0-9_\-]+)/delete-history", path)
        if m:
            return self._handle_batch_delete_history(handler, parsed, body, m.group(1))
        m = re.fullmatch(r"/api/import/batches/([A-Za-z0-9_\-]+)/revoke", path)
        if m:
            return self._handle_batch_revoke(handler, parsed, body, m.group(1))
        # ---------- 以下业务写操作: 按会话/目标店铺路由 ----------
        # 批量路由自行展开 _write_targets（admin 无店也允许 all），不经 _write_target
        if path == "/api/reconciliation/run":
            return self._handle_reconciliation_run(handler, body, None)
        if path == "/api/profit/recalculate":
            return self._handle_profit_recalculate(handler, body, None)
        if path == "/api/import/confirm":
            return self._handle_import_confirm(handler, body)
        bundle = self._write_target(handler, body)
        if path == "/api/ad-expenses":
            return self._handle_ad_expense_save(handler, body, bundle)
        if path == "/api/reconciliation/manual-match":
            return self._handle_reconciliation_manual_match(handler, body, bundle)
        if path == "/api/profit/config":
            return self._handle_profit_config(handler, body, bundle)
        if path == "/api/backup":
            return self._handle_backup(handler, body, bundle)
        if path == "/api/restore":
            return self._handle_restore(handler, body, bundle)
        m = re.fullmatch(r"/api/data/([a-z_]+)", path)
        if m:
            return self._handle_write_data(handler, m.group(1), body, bundle)
        raise ApiError(f"接口不存在: {path}", code=404, http_status=404)

    def route_upload(self, handler, parsed):
        """Excel 上传预览（原始二进制 body + X-Filename 头）

        目标店铺: store_id 一律必填（X-Target-Store 头或 query store_id）;
        member 限本人有效店铺, admin 限有效成员店铺（_import_target 复核）。
        """
        qs = parse_qs(parsed.query)
        raw_store = (handler.headers.get("X-Target-Store")
                     or qs.get("store_id", [None])[0] or None)
        bundle, store_id, store_name, owner = self._import_target(
            handler, raw_store, "import_upload")
        length = int(handler.headers.get("Content-Length") or 0)
        if length <= 0:
            raise ApiError("请上传文件内容（请求体为空）", code=400, http_status=400)
        if length > config.MAX_UPLOAD_MB * 1024 * 1024:
            raise ApiError(f"文件超过大小限制 {config.MAX_UPLOAD_MB}MB",
                           code=413, http_status=413)
        filename = unquote(handler.headers.get("X-Filename") or "upload.xlsx")
        data = handler.rfile.read(length)
        source_type = UPLOAD_PATHS[parsed.path]
        result = bundle.importer.preview_upload(source_type, filename, data,
                                                operator=self._operator(handler),
                                                target_store_id=store_id)
        result["target_store_id"] = store_id
        result["target_store_name"] = store_name
        result["target_owner_username"] = (owner or {}).get("username")
        handler._send(ApiResponse.ok(result, message="预览成功"))

    # ---------- 处理器 ----------
    def _operator(self, handler):
        """操作人: 一律取服务端会话身份, 不信任前端 X-Operator 头"""
        try:
            _, session, user = self._resolve_session(handler)
            return user.get("username") or session.get("username") or "unknown"
        except ApiError:
            return "anonymous"

    # ---------- 认证 API ----------
    def _client_ip(self, handler):
        return handler.client_address[0] if handler.client_address else "unknown"

    def _handle_login(self, handler, body):
        username = body.get("username")
        password = body.get("password")
        if not isinstance(username, str) or not isinstance(password, str) or not username:
            raise ApiError("请输入账号和密码", code=400, http_status=400)
        ip = self._client_ip(handler)
        if self.rate_limiter.is_blocked(ip, username):
            self.security_log.append("login.rate_limited", username=username,
                                     ip=ip, result="blocked")
            raise ApiError("失败次数过多，请 15 分钟后再试", code=1004, http_status=429)
        user, reason = self.user_manager.authenticate(username, password)
        if reason == "invalid" or user is None:
            self.rate_limiter.record_failure(ip, username)
            self.security_log.append("login.failed", username=username, ip=ip,
                                     result="failed")
            # 统一提示, 不泄露账号是否存在
            raise ApiError("账号或密码错误", code=1001, http_status=401)
        if reason == "disabled":
            self.security_log.append("login.disabled", username=username,
                                     user_id=user.get("user_id"), ip=ip,
                                     result="denied")
            raise ApiError("账号或密码错误", code=1001, http_status=401)
        # 一对多归属: member 绑定本人第一个有效（active）店铺为会话默认店铺;
        # 无任何有效店铺（含全部停用/归档/被转走）→ 401, 与无店成员口径一致
        store_id = None
        if user.get("role") == "admin":
            pass  # 管理员无业务店铺, 合法放行（全店视角）
        else:
            active = {s.get("store_id") for s
                      in self.user_manager.member_stores()
                      if s.get("owner_user_id") == user["user_id"]}
            for sid in user.get("store_ids") or []:
                if sid in active:
                    store_id = sid
                    break
            if not store_id:
                raise ApiError("账号未绑定店铺，请联系管理员",
                               code=1001, http_status=401)
        self.rate_limiter.clear(ip, username)
        self.user_manager.mark_login(user["user_id"])
        token, session = self.sessions.create(user, store_id)
        self.security_log.append("login.success", username=username,
                                 user_id=user["user_id"], ip=ip)
        store = self.user_manager.get_store(store_id) or {} if store_id else {}
        resp = ApiResponse.ok({
            "user_id": user["user_id"],
            "username": user["username"],
            "role": user.get("role", "member"),
            "store_id": store_id,
            "store_name": store.get("name"),
        }, message="登录成功")
        handler._send_with_cookie(resp, token)

    def _handle_logout(self, handler):
        token, _ = self._session_from_cookie(handler)
        if token:
            session = self.sessions.get(token)
            self.sessions.delete(token)
            if session:
                self.security_log.append(
                    "logout", username=session.get("username"),
                    user_id=session.get("user_id"), ip=self._client_ip(handler))
        handler._send_with_cookie(ApiResponse.ok({"ok": True}, message="已退出登录"), None)

    def _handle_auth_me(self, handler):
        _, session, user = self._resolve_session(handler)
        store = (self.user_manager.get_store(session["store_id"]) or {}
                 if session.get("store_id") else {})
        handler._send(ApiResponse.ok({
            "user_id": user["user_id"],
            "username": user["username"],
            "role": user.get("role", "member"),
            "store_id": session.get("store_id"),
            "store_name": store.get("name"),
            # 一对多归属派生字段（design 2.5.5）: 服务端直出避免前端补发请求
            "store_count": len(user.get("store_ids") or []),
        }))

    def _handle_change_password(self, handler, body):
        _, _, user = self._resolve_session(handler)
        old_password = body.get("old_password")
        new_password = body.get("new_password")
        confirm_password = body.get("confirm_password")
        if not isinstance(old_password, str) or not isinstance(new_password, str):
            raise ApiError("请提供 old_password 和 new_password", code=400, http_status=400)
        updated = self.user_manager.change_password(user["user_id"], old_password,
                                                    new_password,
                                                    confirm_password)
        # 安全设计: 修改密码后该用户全部会话失效（含当前会话）, 需用新密码重新登录
        revoked = self.sessions.revoke_user(user["user_id"])
        self.security_log.append(
            "password.change", username=user.get("username"),
            user_id=user["user_id"], ip=self._client_ip(handler),
            detail={"revoked_sessions": revoked,
                    "new_version": updated.get("password_version")})
        handler._send_with_cookie(
            ApiResponse.ok({"ok": True, "revoked_sessions": revoked,
                            "relogin": True},
                           message="密码已修改，请使用新密码重新登录"), None)

    # ---------- 自主注册 API（公开） ----------
    def _handle_auth_register(self, handler, body):
        """公开注册: 创建普通用户（恒为 member）与默认空店铺

        防护链顺序: 环境前置 → 限流判定 → 限流计数 → 字段校验 → 原子创建
        → 空数据空间初始化（失败回滚）。注册成功不建立会话、不下发 Cookie。
        """
        # ① 环境前置: 系统未初始化时拒绝注册（防止注册出首个 member 用户
        #    导致 is_initialized 为真、/setup 永久禁用、系统永远无管理员）
        if not self.user_manager.is_initialized():
            self.security_log.append(
                "register.failed", ip=self._client_ip(handler),
                result="failed", detail={"reason": "not_initialized"})
            raise ApiError("系统尚未初始化，请先完成管理员初始化",
                           code=400, http_status=400)
        # ② 限流判定（未命中才继续; 环境异常不消耗用户限流额度）
        ip = self._client_ip(handler)
        if self.register_limiter.is_blocked(ip):
            self.security_log.append(
                "register.rate_limited", ip=ip, result="blocked",
                detail={"window_seconds": self.register_limiter.window,
                        "max_requests": self.register_limiter.max_requests})
            raise ApiError("注册过于频繁，请稍后再试",
                           code=1004, http_status=429)
        # ③ 限流计数: 判定通过后立即记录, 此后无论成败均计
        self.register_limiter.record_request(ip)
        # ④ 字段读取: 仅四个白名单键, 无任何权限字段读取路径
        username = body.get("username")
        password = body.get("password")
        confirm_password = body.get("confirm_password")
        store_name = body.get("store_name")
        # ⑤ 原子注册（持锁复检用户名唯一性, 并发同名恰一成功）
        try:
            user, store = self.user_manager.register_user(
                username, password, confirm_password, store_name)
        except AuthError as e:
            self.security_log.append(
                "register.failed", username=username, ip=ip,
                result="failed", detail={"reason": str(e)})
            raise
        # ⑥ 空数据空间初始化（建 data/backups/logs/uploads + 14 类业务文件,
        #    不复制任何数据）; 失败整体回滚: 删用户/店铺记录 + 删店铺目录
        try:
            self.registry.get(store["store_id"])
        except Exception as e:
            rollback_error = None
            try:
                self.user_manager.rollback_user_store(
                    user["user_id"], store["store_id"])
                store_dir = os.path.join(self.stores_root, store["store_id"])
                if os.path.isdir(store_dir):
                    shutil.rmtree(store_dir, ignore_errors=True)
            except Exception as re:
                rollback_error = repr(re)
            self.security_log.append(
                "register.failed", username=username, user_id=user["user_id"],
                ip=ip, result="error",
                detail={"reason": repr(e), "rolled_back": rollback_error is None,
                        "rollback_error": rollback_error})
            raise ApiError("注册失败，请稍后重试", code=500, http_status=500)
        # ⑦ 成功: 记日志 + 五键白名单响应; 不建会话、不下发 Set-Cookie
        self.security_log.append(
            "register.success", username=user["username"],
            user_id=user["user_id"], ip=ip,
            detail={"store_id": store["store_id"]})
        handler._send(ApiResponse.ok({
            "user_id": user["user_id"],
            "username": user["username"],
            "role": user.get("role", "member"),
            "store_id": store["store_id"],
            "store_name": store.get("name"),
        }, message="注册成功，请前往登录"))

    # ---------- 初始化 API ----------
    def _handle_setup_status(self, handler):
        initialized = self.user_manager.is_initialized()
        legacy = []
        from .store_services import _legacy_files
        for _, filename, _ in _legacy_files(self.data_root):
            legacy.append(filename)
        handler._send(ApiResponse.ok({
            "initialized": initialized,
            "legacy_files": legacy,
        }))

    def _handle_setup_init(self, handler, body):
        """一次性初始化: 创建管理员（无业务店铺）+ 旧数据迁入待分配店铺"""
        if self.user_manager.is_initialized():
            raise ApiError("系统已初始化，如需新增用户请登录后由管理员创建",
                           code=400, http_status=400)
        username = body.get("username")
        password = body.get("password")
        confirm = body.get("confirm_password")
        acknowledge = bool(body.get("acknowledge_legacy"))
        if password != confirm:
            raise ApiError("两次输入的密码不一致", code=400, http_status=400)
        from .store_services import _legacy_files
        has_legacy = bool(_legacy_files(self.data_root))
        if has_legacy and not acknowledge:
            raise ApiError(
                "检测到系统已有历史业务数据，必须确认迁移方式后才能初始化"
                "（旧数据将迁入待分配店铺，登录后由管理员分配归属）",
                code=400, http_status=400)
        user, store = self.user_manager.initialize_admin(username, password)
        self.security_log.append(
            "setup.admin_created", username=user["username"],
            user_id=user["user_id"], ip=self._client_ip(handler),
            detail={"store_id": None, "note": "管理员账号（无业务店铺）"})
        migration = {"migrated": [], "backup_dir": None, "uploads": 0}
        legacy_store = None
        if has_legacy:
            legacy_store = self.user_manager.create_unassigned_store(
                "待分配数据（历史迁移）")
            try:
                migration = migrate_legacy_store_data(
                    self.data_root, self.stores_root, legacy_store["store_id"],
                    security_log=self.security_log)
            except Exception as e:
                # 迁移失败: 回滚管理员账号与待分配店铺记录, 保持未初始化状态可重试
                rollback_error = None
                try:
                    self.user_manager.rollback_user_store(
                        user["user_id"], legacy_store["store_id"])
                    store_dir = os.path.join(self.stores_root,
                                             legacy_store["store_id"])
                    if os.path.isdir(store_dir):
                        shutil.rmtree(store_dir, ignore_errors=True)
                except Exception as rb_err:
                    rollback_error = repr(rb_err)
                backup_dir = None
                try:
                    candidates = sorted(
                        (d for d in os.listdir(self.data_root)
                         if d.startswith("pre_migration_backup_")), reverse=True)
                    if candidates:
                        backup_dir = os.path.join(self.data_root, candidates[0])
                except OSError:
                    pass
                self.security_log.append(
                    "setup.rollback", username=user.get("username"),
                    user_id=user.get("user_id"), ip=self._client_ip(handler),
                    result="error",
                    detail={"unassigned_store_id": legacy_store["store_id"],
                            "error": repr(e), "backup_dir": backup_dir,
                            "rollback_error": rollback_error})
                message = (f"旧数据迁移失败，已回滚本次初始化，可直接重试。"
                           f"失败原因: {e}")
                if backup_dir:
                    message += f"；迁移前备份目录: {backup_dir}"
                if rollback_error:
                    message += "（注意：账号回滚未完全成功，请检查用户数据文件后重试）"
                raise ApiError(message, code=500, http_status=500) from e
        if legacy_store is not None:
            # 触发待分配店铺目录初始化并记录迁移报告
            self.registry.get(legacy_store["store_id"])
            store_dir = os.path.join(self.stores_root, legacy_store["store_id"])
            with open(os.path.join(store_dir, "MIGRATION_REPORT.json"), "w",
                      encoding="utf-8") as f:
                json.dump({"admin_user_id": user["user_id"],
                           "unassigned_store_id": legacy_store["store_id"],
                           "assignment_status": "unassigned",
                           "migration": migration,
                           "created_at": now_iso()}, f, ensure_ascii=False,
                          indent=2)
        handler._send(ApiResponse.ok({
            "user_id": user["user_id"],
            "username": user["username"],
            "store_id": None,
            "store_name": None,
            "migration": migration,
        }, message="初始化完成，请使用管理员账号登录"))

    # ---------- 用户管理 API（仅管理员） ----------
    def _handle_user_list(self, handler, parsed=None):
        """用户列表（仅 admin）: 复用 list_users 投影, 不含任何密码字段

        可选筛选分页（nav_user_mgmt）: username 模糊 / status / role /
        page / page_size; 不传任何参数时维持既有全量返回（向后兼容）。
        """
        self._require_admin(handler)
        users = self.user_manager.list_users()
        qs = parse_qs(parsed.query) if parsed is not None else {}
        has_filter = any(k in qs for k in ("username", "status", "role",
                                           "page", "page_size", "size"))
        if not has_filter:
            handler._send(ApiResponse.ok({"users": users,
                                          "total": len(users)}))
            return
        username_kw = (qs.get("username", [None])[0] or "").strip().lower()
        status = qs.get("status", [None])[0]
        role = qs.get("role", [None])[0]
        if status is not None and status not in ("active", "disabled"):
            raise ApiError("账号状态仅支持 active / disabled",
                           code=400, http_status=400)
        if role is not None and role not in ("admin", "member"):
            raise ApiError("角色仅支持 admin / member",
                           code=400, http_status=400)
        filtered = [
            u for u in users
            if (not username_kw
                or username_kw in (u.get("username") or "").lower())
            and (status is None or u.get("status") == status)
            and (role is None or u.get("role") == role)
        ]
        page, size = self._page_params(qs)
        start = (page - 1) * size
        handler._send(ApiResponse.ok({
            "users": filtered[start:start + size],
            "total": len(filtered),
            "page": page,
            "page_size": size,
        }))

    def _handle_user_create(self, handler, body):
        _, operator = self._require_admin(handler)
        username = body.get("username")
        password = body.get("password")
        role = body.get("role", "member")
        store_name = body.get("store_name")
        status = body.get("status", "active")
        create_store = body.get("create_store", True)
        if not isinstance(create_store, bool):
            raise ApiError("create_store 仅支持 true / false",
                           code=400, http_status=400)
        if role == "admin" and create_store is False:
            create_store = True  # admin 本就不建店, 参数对 admin 无意义
        user, store = self.user_manager.create_user(
            operator, username, password, role=role, store_name=store_name,
            status=status, create_store=create_store)
        if store is not None:
            self.registry.get(store["store_id"])  # 初始化新店铺目录
        self.security_log.append(
            "user.create", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            detail={"new_username": user["username"],
                    "new_user_id": user["user_id"],
                    "store_id": store["store_id"] if store else None,
                    "role": role,
                    "status": user["status"],
                    "note": (None if store
                             else ("管理员账号（无业务店铺）" if role == "admin"
                                   else "未创建初始店铺（稍后分配）"))})
        handler._send(ApiResponse.ok({
            "user_id": user["user_id"], "username": user["username"],
            "role": user["role"],
            "status": user["status"],
            "store_id": store["store_id"] if store else None,
            "store_name": store["name"] if store else None,
        }, message="用户创建成功"))

    def _handle_user_rename(self, handler, body, user_id):
        """修改用户名（仅管理员）: 登录账号变更, user_id 与权限数据不变"""
        _, operator = self._require_admin(handler)
        new_username = body.get("new_username")
        if not isinstance(new_username, str) or not new_username.strip():
            raise ApiError("请提供新用户名（new_username）",
                           code=400, http_status=400)
        old_username, updated = self.user_manager.rename_user(
            operator, user_id, new_username)
        self.security_log.append(
            "user.rename", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            detail={"target_user_id": user_id,
                    "old_username": old_username,
                    "new_username": updated["username"]})
        handler._send(ApiResponse.ok({
            "user_id": updated["user_id"], "username": updated["username"],
        }, message="用户名已更新"))

    def _handle_user_status(self, handler, body, user_id):
        _, operator = self._require_admin(handler)
        action = body.get("action")
        target = self.user_manager.set_status(operator, user_id, action)
        if action == "disable":
            revoked = self.sessions.revoke_user(user_id)
            self.security_log.append(
                "user.disable", username=operator.get("username"),
                user_id=operator["user_id"], ip=self._client_ip(handler),
                detail={"target_user_id": user_id,
                        "target_username": target.get("username"),
                        "revoked_sessions": revoked})
        else:
            self.security_log.append(
                "user.enable", username=operator.get("username"),
                user_id=operator["user_id"], ip=self._client_ip(handler),
                detail={"target_user_id": user_id,
                        "target_username": target.get("username")})
        handler._send(ApiResponse.ok({
            "user_id": user_id, "status": target["status"],
        }, message="账号已停用" if action == "disable" else "账号已启用"))

    def _handle_user_reset_password(self, handler, body, user_id):
        _, operator = self._require_admin(handler)
        new_password = body.get("new_password")
        target = self.user_manager.reset_password(operator, user_id, new_password)
        revoked = self.sessions.revoke_user(user_id)
        self.security_log.append(
            "user.password_reset", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            detail={"target_user_id": user_id,
                    "target_username": target.get("username"),
                    "revoked_sessions": revoked})
        handler._send(ApiResponse.ok({"user_id": user_id},
                                     message="密码已重置，该用户的所有会话已失效"))

    # ---------- 成员管理: 数据清理与账号删除（仅管理员） ----------
    def _target_stores(self, target_user):
        """解析目标用户名下的全部有效店铺记录（店铺记录缺失的跳过）"""
        stores = []
        for store_id in target_user.get("store_ids") or []:
            store = self.user_manager.get_store(store_id)
            if store is not None:
                stores.append(store)
        return stores

    def _snapshot_store(self, store_id):
        """整店快照备份到全局备份目录, 并校验文件清单一致后返回 (快照目录, 文件数)

        校验失败时删除半成品快照并抛 500, 不做任何数据变更。
        """
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        store_dir = os.path.join(self.stores_root, store_id)
        snapshot_dir = os.path.join(self.backup_dir, f"store_snapshot_{store_id}_{ts}")
        shutil.copytree(store_dir, snapshot_dir)

        def _inventory(root):
            items = {}
            for dirpath, _dirs, files in os.walk(root):
                for fn in files:
                    fp = os.path.join(dirpath, fn)
                    items[os.path.relpath(fp, root)] = os.path.getsize(fp)
            return items

        source, copy = _inventory(store_dir), _inventory(snapshot_dir)
        if source != copy:
            shutil.rmtree(snapshot_dir, ignore_errors=True)
            raise ApiError("快照备份校验失败（文件清单与源目录不一致），已中止且未做任何变更",
                           code=500, http_status=500)
        return snapshot_dir, len(copy)

    def _restore_store_from_snapshot(self, store_id, snapshot_dir):
        """从快照恢复店铺业务数据文件与上传文件（清理失败回滚用, 幂等）"""
        bundle = self.registry.get(store_id)
        snap_data = os.path.join(snapshot_dir, "data")
        for name in BUSINESS_FILES:
            src = os.path.join(snap_data, BUSINESS_FILES[name]["file"])
            if os.path.exists(src):
                shutil.copy2(src, os.path.join(bundle.data_dir,
                                               BUSINESS_FILES[name]["file"]))
        snap_uploads = os.path.join(snapshot_dir, "uploads")
        if os.path.isdir(snap_uploads):
            for dirpath, _dirs, files in os.walk(snap_uploads):
                rel = os.path.relpath(dirpath, snap_uploads)
                dst_dir = (bundle.upload_dir if rel == "."
                           else os.path.join(bundle.upload_dir, rel))
                os.makedirs(dst_dir, exist_ok=True)
                for fn in files:
                    shutil.copy2(os.path.join(dirpath, fn), os.path.join(dst_dir, fn))
        self.registry.evict(store_id)

    def _clear_store_data_files(self, store_id, snapshot_dir):
        """清空单个店铺的 12 类业务数据文件与 uploads 暂存; 失败时从快照回滚

        保留 settings/操作日志/店铺级备份; 清理完成后驱逐服务缓存,
        防止总览等页面继续读到内存中的旧数据。返回清理前统计。
        """
        bundle = self.registry.get(store_id)
        records_before = {}
        for name in CLEARABLE_DATA_FILES:
            try:
                records_before[name] = len(bundle.repo.read(name).get("records", []))
            except (RepositoryError, OSError, ValueError):
                records_before[name] = None
        upload_files = [os.path.join(dirpath, fn)
                        for dirpath, _dirs, files in os.walk(bundle.upload_dir)
                        for fn in files]
        try:
            for name in CLEARABLE_DATA_FILES:
                bundle.repo.write(name, BUSINESS_FILES[name]["initial"](),
                                  backup=False)
            for fp in upload_files:
                os.remove(fp)
        except Exception as e:
            self._restore_store_from_snapshot(store_id, snapshot_dir)
            raise ApiError(f"清理过程中发生错误，已从快照回滚，数据保持原状: {e}",
                           code=500, http_status=500)
        self.registry.evict(store_id)
        return {"records_before": records_before, "uploads_removed": len(upload_files)}

    # ---------- 店铺物理删除事务（store_physical_delete） ----------
    @staticmethod
    def _dir_inventory(root):
        """目录文件清单（相对路径 → 字节数）, 快照校验用"""
        items = {}
        for dirpath, _dirs, files in os.walk(root):
            for fn in files:
                fp = os.path.join(dirpath, fn)
                items[os.path.relpath(fp, root)] = os.path.getsize(fp)
        return items

    def _write_delete_txn_state(self, txn_dir, state):
        """txn.json 状态文件原子写入（复用 UserManager._atomic_write）"""
        path = os.path.join(txn_dir, DELETE_TXN_STATE_FILE)
        UserManager._atomic_write(path, state)
        return path

    def _advance_delete_txn(self, txn_dir, state, **extra):
        """原子推进事务状态: prepared → committing → committed / rolled_back

        状态文件损坏时以最小结构重建, 保证状态机可继续推进。
        """
        path = os.path.join(txn_dir, DELETE_TXN_STATE_FILE)
        data = UserManager._read_json(path) or {}
        data["state"] = state
        data["updated_at"] = now_iso()
        data.update(extra)
        UserManager._atomic_write(path, data)

    def _find_pending_delete_txn(self, store_id):
        """扫描目标店铺未完成（prepared/committing）的删除事务目录

        返回事务目录路径或 None; txn.json 无法判读时保守视为执行中。
        """
        prefix = f"{DELETE_TXN_DIR_PREFIX}{store_id}_"
        try:
            names = os.listdir(self.backup_dir)
        except OSError:
            return None
        for name in sorted(names):
            if not name.startswith(prefix):
                continue
            state_file = os.path.join(self.backup_dir, name,
                                      DELETE_TXN_STATE_FILE)
            if not os.path.isfile(state_file):
                continue
            try:
                with open(state_file, encoding="utf-8") as f:
                    state = json.load(f).get("state")
            except (OSError, ValueError):
                return os.path.join(self.backup_dir, name)
            if state in ("prepared", "committing"):
                return os.path.join(self.backup_dir, name)
        return None

    def _create_delete_txn(self, store_id, store, operator, inventory,
                           client_ip=None):
        """创建物理删除事务: 同店互斥 + txn.json + 两表副本 + 整店快照

        同店并发删除在 FileLock 临界区内扫描未完成事务 → 409;
        快照/副本校验失败 → 删除半成品事务目录并抛错（零数据变更）。
        """
        lock_path = os.path.join(
            self.backup_dir, f"{DELETE_TXN_DIR_PREFIX}{store_id}.lock")
        with FileLock(lock_path):
            pending = self._find_pending_delete_txn(store_id)
            if pending is not None:
                self.security_log.append(
                    "store.delete_conflict",
                    username=operator.get("username"),
                    user_id=operator["user_id"], ip=client_ip,
                    result="denied",
                    detail={"store_id": store_id,
                            "txn_id": os.path.basename(pending)})
                raise ApiError("该店铺删除操作正在执行，请勿重复提交",
                               code=409, http_status=409)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            txn_dir = os.path.join(
                self.backup_dir, f"{DELETE_TXN_DIR_PREFIX}{store_id}_{ts}")
            os.mkdir(txn_dir)
            try:
                self._write_delete_txn_state(txn_dir, {
                    "version": 1, "state": "prepared",
                    "store_id": store_id,
                    "store_name": store.get("name"),
                    "owner_user_id": store.get("owner_user_id"),
                    "operator": operator.get("username"),
                    "expected_counts": inventory,
                    "created_at": now_iso(), "updated_at": now_iso(),
                })
                for name, path in (
                        ("users.json.bak", self.user_manager.users_path),
                        ("stores.json.bak", self.user_manager.stores_path)):
                    if os.path.exists(path):
                        shutil.copy2(path, os.path.join(txn_dir, name))
                store_dir = os.path.join(self.stores_root, store_id)
                snapshot_dir = os.path.join(txn_dir, "store_snapshot")
                shutil.copytree(store_dir, snapshot_dir)
                if (self._dir_inventory(store_dir)
                        != self._dir_inventory(snapshot_dir)):
                    raise OSError("整店快照与源目录文件清单不一致")
            except Exception:
                shutil.rmtree(txn_dir, ignore_errors=True)
                raise ApiError("备份创建失败，已中止且未做任何变更",
                               code=500, http_status=500)
            return txn_dir

    def _rollback_delete_txn(self, txn_dir, store_id):
        """从事务快照与两表副本回滚到删除前一致状态

        两表副本原子回写 → 整店目录从快照重建 → 驱逐缓存;
        返回 "ok"; 回滚失败返回 "failed"（事务标记
        rolled_back_pending_manual, 保留现场待人工处置）。
        """
        try:
            stores_bak = os.path.join(txn_dir, "stores.json.bak")
            users_bak = os.path.join(txn_dir, "users.json.bak")
            snapshot_dir = os.path.join(txn_dir, "store_snapshot")
            if os.path.exists(stores_bak):
                data = UserManager._read_json(stores_bak)
                if data is not None:
                    UserManager._atomic_write(
                        self.user_manager.stores_path, data)
            if os.path.exists(users_bak):
                data = UserManager._read_json(users_bak)
                if data is not None:
                    UserManager._atomic_write(
                        self.user_manager.users_path, data)
            if os.path.isdir(snapshot_dir):
                store_dir = os.path.join(self.stores_root, store_id)
                if os.path.exists(store_dir):
                    shutil.rmtree(store_dir)
                shutil.copytree(snapshot_dir, store_dir)
            self.registry.evict(store_id)
            self._advance_delete_txn(txn_dir, "rolled_back")
            return "ok"
        except Exception:
            try:
                self._advance_delete_txn(txn_dir, "rolled_back_pending_manual")
            except (OSError, ValueError):
                pass
            return "failed"

    def _verify_store_deleted(self, store_id):
        """删除事务提交前回读核对: 两表无记录且无残留引用"""
        if self.user_manager.get_store(store_id) is not None:
            raise RuntimeError(
                f"店铺记录仍存在于 stores.json: {store_id}")
        for u in self.user_manager.list_users():
            if store_id in (u.get("store_ids") or []):
                raise RuntimeError(
                    f"用户 {u.get('user_id')} 仍引用店铺: {store_id}")

    def _run_delete_txn(self, store_id, txn_dir, operator, store):
        """执行删除事务: 固定顺序 业务文件 → 两表 → 目录, 全程可回滚

        ① 吊销归属成员会话（防删除期间并发写入）
        ② 清空 12 类业务文件与 uploads（复用 _clear_store_data_files）
        ③ 删除两表记录（UserManager.delete_store 自带两表补偿）
        ④ 移除店铺目录（含 settings/logs/店铺级备份; OSError 如实上报）
        ⑤ 统计回读核对（两表无记录无引用）→ ⑥ committed。
        任一环节异常 → 事务回滚 → 500（已恢复原状）; 回滚失败 →
        rolled_back_pending_manual 并保留现场, 绝不部分删除。
        """
        snapshot_dir = os.path.join(txn_dir, "store_snapshot")
        store_dir = os.path.join(self.stores_root, store_id)
        owner_id = store.get("owner_user_id")
        records_before, uploads_removed, revoked = {}, 0, 0
        deleted_store, removed_dir = None, False
        try:
            self._advance_delete_txn(txn_dir, "committing")
            revoked = self.sessions.revoke_user(owner_id) if owner_id else 0
            stats = self._clear_store_data_files(store_id, snapshot_dir)
            records_before = stats["records_before"]
            uploads_removed = stats["uploads_removed"]
            deleted_store = self.user_manager.delete_store(operator, store_id)
            try:
                self.registry.evict(store_id)
                shutil.rmtree(store_dir)
                removed_dir = not os.path.exists(store_dir)
            except OSError:
                removed_dir = False
            self._verify_store_deleted(store_id)
            self._advance_delete_txn(txn_dir, "committed")
        except Exception as e:
            rollback = self._rollback_delete_txn(txn_dir, store_id)
            if rollback != "ok":
                raise ApiError(
                    f"删除回滚失败，数据现场已保留（{txn_dir}），"
                    f"请联系运维: {e}", code=500, http_status=500) from e
            raise ApiError(
                f"清理过程中发生错误，已从快照回滚，数据保持原状: {e}",
                code=500, http_status=500) from e
        return {
            "deleted_store": deleted_store or store,
            "records_before": records_before,
            "uploads_removed": uploads_removed,
            "revoked": revoked,
            "removed_dir": removed_dir,
        }

    def _append_store_tombstone(self, record):
        """墓碑文件追加已删除店铺记录（FileLock + 原子写, 只增不删）"""
        path = os.path.join(self.backup_dir, TOMBSTONE_FILE)
        with FileLock(path + ".lock"):
            data = UserManager._read_json(path) or {
                "version": 1, "updated_at": now_iso(), "tombstones": []}
            data["tombstones"] = [
                t for t in data.get("tombstones", [])
                if t.get("store_id") != record["store_id"]]
            data["tombstones"].append(record)
            data["updated_at"] = now_iso()
            UserManager._atomic_write(path, data)
        return path

    def _govern_store_backups(self, store_id, txn_dir):
        """删除成功后的备份治理: 全局快照清理 + 墓碑写入 + 事务目录删除

        治理失败不阻断删除主结果: issues 如实返回（记入审计与
        backup_treatment, 不掩盖、不谎报）。
        """
        issues = []
        snapshot_prefix = f"store_snapshot_{store_id}_"
        try:
            for name in os.listdir(self.backup_dir):
                path = os.path.join(self.backup_dir, name)
                if name.startswith(snapshot_prefix) and os.path.isdir(path):
                    shutil.rmtree(path, ignore_errors=True)
        except OSError as e:
            issues.append(f"全局快照扫描失败: {e}")
        txn = UserManager._read_json(
            os.path.join(txn_dir, DELETE_TXN_STATE_FILE)) or {}
        try:
            self._append_store_tombstone({
                "store_id": store_id,
                "store_name": txn.get("store_name"),
                "deleted_at": now_iso(),
                "operator_username": txn.get("operator"),
                "txn_id": os.path.basename(txn_dir),
                "counts_summary": (txn.get("deleted_counts")
                                   or txn.get("expected_counts")),
            })
        except (OSError, ValueError) as e:
            issues.append(f"墓碑写入失败: {e}")
        shutil.rmtree(txn_dir, ignore_errors=True)
        if os.path.isdir(txn_dir):
            issues.append(f"事务目录删除失败，残留: {txn_dir}")
        return issues

    def _recover_pending_delete_txns(self):
        """启动扫描器: 未完成的物理删除事务保守恢复到一致状态

        prepared/committing（中断于删除中）→ 快照+两表副本回滚;
        committed 残留（清理中断）→ 幂等完成备份治理与目录清理;
        rolled_back 残留 → 清理事务目录; 无状态文件半成品 → 清理
        （未进入 prepared, 删除动作未开始）; rolled_back_pending_manual
        → 仅告警审计, 不自动破坏现场。
        """
        try:
            names = os.listdir(self.backup_dir)
        except OSError:
            return
        for name in sorted(names):
            txn_dir = os.path.join(self.backup_dir, name)
            if (not name.startswith(DELETE_TXN_DIR_PREFIX)
                    or not os.path.isdir(txn_dir)):
                continue
            state_file = os.path.join(txn_dir, DELETE_TXN_STATE_FILE)
            if not os.path.isfile(state_file):
                shutil.rmtree(txn_dir, ignore_errors=True)
                continue
            try:
                with open(state_file, encoding="utf-8") as f:
                    txn = json.load(f)
            except (OSError, ValueError) as e:
                self._log_lifecycle(
                    "store.delete_recovered",
                    detail={"txn_id": name,
                            "error": f"txn.json 不可读: {e!r}"},
                    result="error")
                continue
            state, store_id = txn.get("state"), txn.get("store_id")
            if state in ("prepared", "committing"):
                rollback = self._rollback_delete_txn(txn_dir, store_id)
                if rollback == "ok":
                    shutil.rmtree(txn_dir, ignore_errors=True)
                self._log_lifecycle(
                    "store.delete_recovered",
                    detail={"txn_id": name, "store_id": store_id,
                            "recovered_state": (
                                "rolled_back" if rollback == "ok"
                                else "rolled_back_pending_manual")},
                    result="success" if rollback == "ok" else "error")
            elif state == "committed":
                issues = self._govern_store_backups(store_id, txn_dir)
                self._log_lifecycle(
                    "store.delete_recovered",
                    detail={"txn_id": name, "store_id": store_id,
                            "recovered_state": "cleanup_completed",
                            "governance_issues": issues})
            elif state == "rolled_back":
                shutil.rmtree(txn_dir, ignore_errors=True)
                self._log_lifecycle(
                    "store.delete_recovered",
                    detail={"txn_id": name, "store_id": store_id,
                            "recovered_state": "rolled_back_txn_cleaned"})
            elif state == "rolled_back_pending_manual":
                self._log_lifecycle(
                    "store.delete_recovered",
                    detail={"txn_id": name, "store_id": store_id,
                            "recovered_state": "rolled_back_pending_manual",
                            "note": "人工处置后可删除事务目录"},
                    result="error")

    def _snapshot_all_stores(self, stores):
        """对多个店铺依次创建快照; 任一失败则清理全部已生成快照并抛错（无副作用）"""
        snapshots = []
        try:
            for store in stores:
                snap_dir, file_count = self._snapshot_store(store["store_id"])
                snapshots.append((store, snap_dir, file_count))
        except (ApiError, OSError):
            for _store, snap_dir, _n in snapshots:
                shutil.rmtree(snap_dir, ignore_errors=True)
            raise
        return snapshots

    @staticmethod
    def _backup_info(snapshots):
        return [{"store_id": store["store_id"], "store_name": store.get("name"),
                 "snapshot": os.path.basename(snap_dir), "files": file_count}
                for store, snap_dir, file_count in snapshots]

    def _handle_store_data_summary(self, handler, user_id):
        """数据清理预览: 目标成员各店铺的数据类别/记录数/上传文件数（仅 admin）"""
        self._require_admin(handler)
        target = self.user_manager.get_user(user_id)
        if target is None:
            raise ApiError(f"用户不存在: {user_id}", code=404, http_status=404)
        summary = []
        for store in self._target_stores(target):
            bundle = self.registry.get(store["store_id"])
            categories = {}
            for name in CLEARABLE_DATA_FILES:
                try:
                    categories[name] = len(bundle.repo.read(name).get("records", []))
                except (RepositoryError, OSError, ValueError):
                    categories[name] = None
            upload_count = sum(len(files) for _d, _dirs, files
                               in os.walk(bundle.upload_dir))
            summary.append({"store_id": store["store_id"],
                            "store_name": store.get("name"),
                            "categories": categories,
                            "upload_files": upload_count})
        handler._send(ApiResponse.ok({
            "user_id": user_id, "username": target.get("username"),
            "stores": summary,
        }))

    def _handle_store_data_clear(self, handler, body, user_id):
        """清除成员店铺全部业务数据（场景 C: 保留账号/店铺记录/settings/日志; 仅 admin）

        流程: 权限与归属校验 -> 店铺名二次确认 -> 整店快照(校验) -> 吊销目标会话
        -> 逐类原子清理(失败从快照回滚) -> 驱逐缓存 -> 审计日志。
        """
        _, operator = self._require_admin(handler)
        target = self.user_manager.get_user(user_id)
        if target is None:
            raise ApiError(f"用户不存在: {user_id}", code=404, http_status=404)
        if target["user_id"] == operator["user_id"]:
            raise ApiError("不能清除自己店铺的业务数据", code=400, http_status=400)
        stores = self._target_stores(target)
        if not stores:
            raise ApiError("该成员名下没有店铺，无需清理", code=400, http_status=400)
        confirm_name = body.get("confirm_store_name")
        if (not isinstance(confirm_name, str)
                or confirm_name not in [s.get("name") for s in stores]):
            raise ApiError("店铺名称确认不一致，请输入目标成员的店铺名称后重试",
                           code=400, http_status=400)
        snapshots = self._snapshot_all_stores(stores)
        # 快照就绪后吊销目标会话, 避免清理期间该成员并发写入
        revoked = self.sessions.revoke_user(user_id)
        cleared = []
        try:
            for store, snap_dir, _n in snapshots:
                stats = self._clear_store_data_files(store["store_id"], snap_dir)
                cleared.append({"store_id": store["store_id"],
                                "store_name": store.get("name"), **stats})
        except ApiError:
            # 跨店铺回滚: 恢复全部涉及店铺（含清理中断的那家）, 保证不出现半删状态
            for store, snap_dir, _n in snapshots:
                self._restore_store_from_snapshot(store["store_id"], snap_dir)
            raise
        backup_info = self._backup_info(snapshots)
        self.security_log.append(
            "store.clear_data", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            detail={"target_user_id": user_id,
                    "target_username": target.get("username"),
                    "stores": [{"store_id": c["store_id"],
                                "store_name": c["store_name"],
                                "records_before": c["records_before"],
                                "uploads_removed": c["uploads_removed"]}
                               for c in cleared],
                    "backup": backup_info, "revoked_sessions": revoked})
        handler._send(ApiResponse.ok({
            "user_id": user_id, "username": target.get("username"),
            "cleared": cleared, "backup": backup_info,
            "revoked_sessions": revoked,
        }, message="店铺业务数据已清除（已创建整店快照备份）"))

    def _handle_user_delete(self, handler, body, user_id):
        """删除成员账号（仅 admin）

        场景 A（clear_store_data=false）: 仅删除账号, 保留店铺记录与全部业务数据,
        历史数据仍可追溯; 场景 B（true）: 整店快照 -> 清空业务数据 -> 删除账号与
        店铺记录并移除店铺目录, 快照保留在全局备份目录。两种场景均吊销全部会话。
        """
        _, operator = self._require_admin(handler)
        target = self.user_manager.get_user(user_id)
        if target is None:
            raise ApiError(f"用户不存在: {user_id}", code=404, http_status=404)
        if target["user_id"] == operator["user_id"]:
            raise ApiError("不能删除自己的账号", code=400, http_status=400)
        confirm_username = body.get("confirm_username")
        if (not isinstance(confirm_username, str)
                or confirm_username != target.get("username")):
            raise ApiError("用户名确认不一致，请输入目标成员的用户名后重试",
                           code=400, http_status=400)
        clear_store_data = body.get("clear_store_data", False)
        if not isinstance(clear_store_data, bool):
            raise ApiError("clear_store_data 必须是布尔值", code=400, http_status=400)
        stores = self._target_stores(target)
        revoked = self.sessions.revoke_user(user_id)
        backup_info = None
        cleared = None
        if not clear_store_data:
            # 场景 A: 只删账号, 店铺记录与业务数据完整保留（可追溯, 不转移归属）
            self.user_manager.delete_user(operator, user_id)
            message = "账号已删除，业务数据已完整保留（可追溯）"
        else:
            # 场景 B: 快照 -> 清数据 -> 删账号与店铺记录 -> 移除店铺目录
            snapshots = self._snapshot_all_stores(stores)
            backup_info = self._backup_info(snapshots)
            cleared = []
            try:
                for store, snap_dir, _n in snapshots:
                    stats = self._clear_store_data_files(store["store_id"], snap_dir)
                    cleared.append({"store_id": store["store_id"],
                                    "store_name": store.get("name"), **stats})
            except ApiError:
                for store, snap_dir, _n in snapshots:
                    self._restore_store_from_snapshot(store["store_id"], snap_dir)
                raise
            try:
                self.user_manager.delete_user(operator, user_id)
            except Exception as e:
                # 删除账号失败: 从快照恢复业务数据, 店铺与账号回到操作前状态
                for store, snap_dir, _n in snapshots:
                    self._restore_store_from_snapshot(store["store_id"], snap_dir)
                raise ApiError(
                    f"删除账号失败，已从快照恢复店铺业务数据（账号未删除）: {e}",
                    code=500, http_status=500)
            for store in stores:
                self.user_manager.rollback_user_store(user_id, store["store_id"])
            for store in stores:
                self.registry.evict(store["store_id"])
                shutil.rmtree(os.path.join(self.stores_root, store["store_id"]),
                              ignore_errors=True)
            message = "账号与店铺数据已删除（整店快照已保留在备份目录）"
        self.security_log.append(
            "user.delete", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            detail={"target_user_id": user_id,
                    "target_username": target.get("username"),
                    "clear_store_data": clear_store_data,
                    "store_ids": [s["store_id"] for s in stores],
                    "cleared": [{"store_id": c["store_id"],
                                 "store_name": c["store_name"],
                                 "records_before": c["records_before"],
                                 "uploads_removed": c["uploads_removed"]}
                                for c in cleared] if cleared else None,
                    "backup": backup_info,
                    "revoked_sessions": revoked})
        handler._send(ApiResponse.ok({
            "user_id": user_id, "username": target.get("username"),
            "clear_store_data": clear_store_data,
            "store_ids": [s["store_id"] for s in stores],
            "backup": backup_info, "cleared": cleared,
            "revoked_sessions": revoked,
        }, message=message))

    # ---------- 广告费用 API（唯一可人工编辑的财务金额, 仅限本人店铺） ----------
    def _handle_ad_expenses(self, handler, parsed, bundle):
        qs = parse_qs(parsed.query)
        month = qs.get("month", [None])[0] or None
        self._scoped_options(
            handler, parsed, lambda b: b.ad_expenses.list(month=month))

    def _handle_ad_expense_save(self, handler, body, bundle):
        month = body.get("month")
        amount = body.get("amount")
        note = body.get("note")
        if not month or not isinstance(month, str):
            raise ApiError("请提供 month（格式 YYYY-MM）", code=400, http_status=400)
        if amount is None or isinstance(amount, bool) or not isinstance(amount, (int, float, str)):
            raise ApiError("请提供数字型 amount", code=400, http_status=400)
        record = bundle.ad_expenses.save(month, str(amount), note=note,
                                         operator=self._operator(handler))
        handler._send(ApiResponse.ok(record, message="广告花费已保存"))

    def _handle_static(self, handler, rel):
        static_root = os.path.realpath(config.STATIC_DIR)
        full = os.path.realpath(os.path.join(static_root, rel))
        if full != static_root and not full.startswith(static_root + os.sep):
            raise ApiError("未找到页面资源", code=404, http_status=404)
        if not os.path.isfile(full):
            raise ApiError("未找到页面资源", code=404, http_status=404)
        ext = os.path.splitext(full)[1].lower()
        mime = STATIC_MIME.get(ext, "application/octet-stream")
        with open(full, "rb") as f:
            content = f.read()
        # 侧边栏管理菜单组仅对 admin 输出: 非 admin 会话剔除 admin-nav 标记块
        if ext == ".html" and b"admin-nav-start" in content:
            _, session = self._session_from_cookie(handler)
            is_admin = False
            if session is not None:
                user = self.user_manager.get_user(session.get("user_id"))
                if (user is not None and user.get("status") == "active"
                        and session.get("password_version", 1)
                        == user.get("password_version", 1)
                        and user.get("role") == "admin"):
                    is_admin = True
            if not is_admin:
                content = ADMIN_NAV_BLOCK_RE.sub(b"", content)
        handler._send_bytes(mime, content)

    def _handle_health(self, handler):
        """健康检查: 未登录仅返回基本状态; 登录后返回本人店铺数据文件详情"""
        token, session = self._session_from_cookie(handler)
        bundle = None
        authenticated = False
        if session is not None:
            try:
                user = self.user_manager.get_user(session["user_id"])
                if (user is not None and user.get("status") == "active"
                        and session.get("password_version", 1)
                        == user.get("password_version", 1)):
                    if user.get("role") == "admin":
                        # 管理员会话即认证有效（无店铺也合法）
                        authenticated = True
                        sid = session.get("store_id")
                        if sid and sid in (user.get("store_ids") or []):
                            bundle = self.registry.get(sid)
                    elif session.get("store_id") in (user.get("store_ids") or []):
                        bundle = self.registry.get(session["store_id"])
                        authenticated = True
            except Exception:
                bundle = None
                authenticated = False
        payload = {
            "status": "ok",
            "service": SERVICE_NAME,
            "version": VERSION,
            "started_at": self.started_at,
            "host": self.host,
            "port": self.port,
            "python": platform.python_version(),
            "platform": f"{platform.system()} {platform.release()} {platform.machine()}",
            "authenticated": authenticated,
        }
        if bundle is not None:
            files_status = {}
            for name, spec in BUSINESS_FILES.items():
                entry = {
                    "file": spec["file"],
                    "description": spec["description"],
                    "exists": os.path.exists(os.path.join(bundle.repo.data_dir, spec["file"])),
                    "valid": False,
                }
                if entry["exists"]:
                    try:
                        data = bundle.repo.read(name)
                        entry["valid"] = True
                        coll = data.get(spec["collection_key"])
                        entry["count"] = len(coll) if isinstance(coll, (list, dict)) else None
                    except Exception as e:
                        entry["error"] = str(e)
                files_status[name] = entry
            all_valid = all(v["valid"] for v in files_status.values())
            payload["status"] = "ok" if all_valid else "degraded"
            payload["data_dir"] = bundle.repo.data_dir
            payload["backup_dir"] = bundle.repo.backup_dir
            payload["log_dir"] = bundle.op_log.log_dir
            payload["data_files"] = files_status
        handler._send(ApiResponse.ok(payload))

    def _handle_get_data(self, handler, parsed, name, bundle):
        if name not in BUSINESS_FILES:
            raise ApiError(f"未知数据: {name}", code=404, http_status=404)
        # 统一经查询范围解析: member 携带越权参数（member/store_id）与其他业务 API 同口径 403
        scope = self._query_scope(handler, parsed)
        if not scope.targets:
            raise ApiError("暂无成员店铺数据", code=404, http_status=404)
        bundle = scope.bundle
        handler._send(ApiResponse.ok(bundle.repo.read(name)))

    def _handle_write_data(self, handler, name, body, bundle):
        if name not in BUSINESS_FILES:
            raise ApiError(f"未知数据: {name}", code=404, http_status=404)
        spec = BUSINESS_FILES[name]
        ckey = spec["collection_key"]
        if ckey not in body:
            raise ApiError(f"缺少必需字段: {ckey}", code=400, http_status=400)
        operator = self._operator(handler)

        def mutator(data):
            data[ckey] = body[ckey]
            data["updated_at"] = now_iso()
            return data

        data = bundle.repo.update(name, mutator)
        count = len(data[ckey]) if isinstance(data[ckey], (list, dict)) else None
        bundle.op_log.append(
            f"data.write.{name}", detail={"count": count}, operator=operator)
        handler._send(ApiResponse.ok(data, message="写入成功"))

    def _handle_import_confirm(self, handler, body):
        upload_id = body.get("upload_id")
        if not upload_id or not isinstance(upload_id, str):
            raise ApiError("请提供 upload_id（预览成功后返回）",
                           code=400, http_status=400)
        force = bool(body.get("force", False))
        _, session, user = self._resolve_session(handler)
        # 确认时按会话身份复核目标店铺（缺参/越权/无效一律拒绝, 不写任何数据）
        bundle, target_store_id, store_name, owner = self._import_target(
            handler, body.get("target_store_id"), "import_confirm")
        # 目标成员一律由服务端从店铺归属反查, 不信任前端提交的 user_id
        if owner is None or owner.get("role") != "member" \
                or owner.get("status") != "active":
            raise ApiError("目标店铺不属于有效成员，请先为其创建或分配店铺",
                           code=400, http_status=400)
        target_user_id = owner.get("user_id")
        if user.get("role") == "admin" and user.get("user_id") != target_user_id:
            self.security_log.append(
                "import.proxy", username=user.get("username"),
                user_id=user.get("user_id"), ip=handler.client_address[0],
                detail={"upload_id": upload_id,
                        "target_user_id": target_user_id,
                        "target_store_id": target_store_id})
        result = bundle.importer.confirm_import(
            upload_id, force=force, operator=self._operator(handler),
            target_store_id=target_store_id, target_user_id=target_user_id,
            uploaded_by_user_id=user.get("user_id"),
            target_store_name=store_name)
        handler._send(ApiResponse.ok(result, message="导入成功"))

    # ---------- 统一月份筛选（month/from/to -> 左闭右开区间） ----------
    @staticmethod
    def _time_filter(qs):
        """解析时间筛选：优先级 from/to > month > start/end（旧参数）"""
        return resolve_range(
            month=qs.get("month", [None])[0] or None,
            date_from=qs.get("from", [None])[0] or None,
            date_to=qs.get("to", [None])[0] or None,
            start=qs.get("start", [None])[0] or None,
            end=qs.get("end", [None])[0] or None,
        )

    @staticmethod
    def _time_meta(qs):
        """区间元信息（响应附加，供前端展示当前筛选条件）"""
        return range_meta(
            month=qs.get("month", [None])[0] or None,
            date_from=qs.get("from", [None])[0] or None,
            date_to=qs.get("to", [None])[0] or None,
            start=qs.get("start", [None])[0] or None,
            end=qs.get("end", [None])[0] or None,
        )

    @staticmethod
    def _page_params(qs, default_size=50):
        try:
            page = int(qs.get("page", ["1"])[0] or 1)
            size_raw = (qs.get("page_size", [None])[0]
                        or qs.get("size", [None])[0] or str(default_size))
            size = int(size_raw)
        except ValueError:
            raise ApiError("page/page_size 必须是整数", code=400, http_status=400)
        if page < 1:
            raise ApiError("page 必须 >= 1", code=400, http_status=400)
        if size not in PAGE_SIZES:
            raise ApiError(f"page_size 仅支持 {list(PAGE_SIZES)}",
                           code=400, http_status=400)
        return page, size

    # ---------- 订单与成本对账 ----------
    @staticmethod
    def _date_range(body):
        return resolve_range(
            month=body.get("month") or None,
            date_from=body.get("from") or None,
            date_to=body.get("to") or None,
            start=body.get("start") or None,
            end=body.get("end") or None,
        )

    def _handle_reconciliation_run(self, handler, body, bundle):
        start, end = self._date_range(body)
        operator = self._operator(handler)
        results = []
        for b, sid, sname in self._write_targets(handler, body):
            r = dict(b.reconciler.run_reconciliation(
                start=start, end=end, operator=operator))
            r["store_id"] = sid
            r["store_name"] = sname
            results.append(r)
        if len(results) == 1:
            return handler._send(ApiResponse.ok(results[0], message="对账完成"))
        handler._send(ApiResponse.ok(
            {"store_count": len(results), "stores": results}, message="对账完成"))

    def _reconciliation_filters(self, qs):
        start, end = self._time_filter(qs)
        return {
            "start": start, "end": end,
            "order_id": qs.get("order_id", [None])[0] or None,
            "match_status": qs.get("match_status", [None])[0] or None,
            "anomaly_type": qs.get("anomaly_type", [None])[0] or None,
        }

    def _handle_reconciliation_list(self, handler, parsed, bundle):
        filters = self._reconciliation_filters(parse_qs(parsed.query))

        def data_fn(b, _sid, _sname, page, size):
            return b.reconciler.list_results(page=page, size=size, **filters)

        self._scoped_list(handler, parsed, data_fn, "results")

    def _handle_reconciliation_export(self, handler, parsed, bundle):
        filters = self._reconciliation_filters(parse_qs(parsed.query))
        operator = self._operator(handler)

        def export_fn(b):
            return b.reconciler.export_excel(operator=operator, **filters)

        self._scoped_export(handler, parsed, export_fn)

    def _handle_reconciliation_options(self, handler, parsed, bundle):
        self._scoped_options(handler, parsed, lambda b: b.reconciler.get_options())

    def _handle_reconciliation_manual_match(self, handler, body, bundle):
        order_id = body.get("order_id")
        if not order_id or not isinstance(order_id, str):
            raise ApiError("请提供 order_id（微信订单号）", code=400, http_status=400)
        action = body.get("action", "confirm")
        if action not in ("confirm", "unbind"):
            raise ApiError("action 仅支持 confirm / unbind",
                           code=400, http_status=400)
        sids = body.get("supplier_record_ids")
        if sids is not None and (not isinstance(sids, list)
                                 or not all(isinstance(s, str) for s in sids)):
            raise ApiError("supplier_record_ids 必须是字符串数组",
                           code=400, http_status=400)
        gran = body.get("cost_granularity") or None
        if gran not in (None, "order_level", "line_level"):
            raise ApiError("cost_granularity 仅支持 order_level / line_level",
                           code=400, http_status=400)
        record = bundle.reconciler.manual_match(
            order_id, sids, action=action, cost_granularity=gran,
            note=body.get("note"), operator=self._operator(handler))
        handler._send(ApiResponse.ok(record, message="人工确认已保存"))

    # ---------- 利润核算与经营总览 ----------
    @staticmethod
    def _profit_filters(qs):
        start, end = ReconcileServer._time_filter(qs)
        return {
            "start": start, "end": end,
            "order_id": qs.get("order_id", [None])[0] or None,
            "profit_status": qs.get("profit_status", [None])[0] or None,
            "calc_id": qs.get("calc_id", [None])[0] or None,
        }

    def _handle_profit_recalculate(self, handler, body, bundle):
        start, end = self._date_range(body)
        operator = self._operator(handler)
        results = []
        for b, sid, sname in self._write_targets(handler, body):
            r = dict(b.profits.run_calculation(
                start=start, end=end, operator=operator))
            r["store_id"] = sid
            r["store_name"] = sname
            results.append(r)
        if len(results) == 1:
            return handler._send(ApiResponse.ok(results[0], message="利润核算完成"))
        handler._send(ApiResponse.ok(
            {"store_count": len(results), "stores": results},
            message="利润核算完成"))

    def _handle_profit_list(self, handler, parsed, bundle):
        filters = self._profit_filters(parse_qs(parsed.query))

        def data_fn(b, _sid, _sname, page, size):
            return b.profits.list_results(page=page, size=size, **filters)

        self._scoped_list(handler, parsed, data_fn, "results")

    def _handle_profit_export(self, handler, parsed, bundle):
        filters = self._profit_filters(parse_qs(parsed.query))
        operator = self._operator(handler)
        scope = self._query_scope(handler, parsed)

        if scope.multi:
            stats_sink = []

            def export_fn(b):
                try:
                    return b.profits.export_excel(
                        operator=operator, include_summary_row=False,
                        stats_sink=stats_sink, **filters)
                except ProfitServiceError:
                    # 空数据店转 ValueError 交 _scoped_export 既有跳过逻辑
                    # （ProfitServiceError 非 ValueError 子类，直接传播会 400）
                    raise ValueError("当前筛选无数据可导出")

            def post_merge(contents):
                return self._append_profit_summary_row(
                    self._merge_excel_buffers(contents), stats_sink)

            return self._scoped_export(handler, parsed, export_fn,
                                       post_merge=post_merge)

        def export_fn(b):
            return b.profits.export_excel(operator=operator, **filters)

        self._scoped_export(handler, parsed, export_fn)

    def _handle_profit_options(self, handler, parsed, bundle):
        self._scoped_options(handler, parsed, lambda b: b.profits.get_options())

    def _handle_profit_config(self, handler, body, bundle):
        refund_status = body.get("refund_deduction_status")
        subsidy = body.get("platform_subsidy_included")
        if refund_status is not None and not isinstance(refund_status, str):
            raise ApiError("refund_deduction_status 必须是字符串",
                           code=400, http_status=400)
        if subsidy is not None and not isinstance(subsidy, str):
            raise ApiError("platform_subsidy_included 必须是字符串",
                           code=400, http_status=400)
        result = bundle.profits.update_config(
            refund_deduction_status=refund_status,
            platform_subsidy_included=subsidy,
            operator=self._operator(handler))
        handler._send(ApiResponse.ok(result, message="口径配置已保存"))

    def _handle_summary(self, handler, parsed, bundle):
        qs = parse_qs(parsed.query)
        start, end = self._time_filter(qs)
        scope = self._query_scope(handler, parsed)
        if not scope.multi:
            result = scope.bundle.profits.get_summary(start, end)
        else:
            result = {}
            for b, _sid, _sn in scope.targets:
                ReconcileServer._merge_dicts(
                    result, b.profits.get_summary(start, end))
        result["time_range"] = self._time_meta(qs)
        if scope.multi:
            result["store_scope"] = "all"
            result["store_count"] = len(scope.targets)
        handler._send(ApiResponse.ok(result))

    # ---------- 订单查询 ----------
    def _parse_list_params(self, qs, kind):
        page, size = self._page_params(qs)
        order_by = qs.get("order_by", ["order_time"])[0]
        allowed = NORMAL_SORT_FIELDS if kind == "normal" else REFUND_SORT_FIELDS
        if order_by not in allowed:
            raise ApiError(f"order_by 仅支持 {list(allowed)}", code=400, http_status=400)
        order_dir = qs.get("order_dir", ["desc"])[0]
        if order_dir not in ("asc", "desc"):
            raise ApiError("order_dir 仅支持 asc/desc", code=400, http_status=400)
        time_field = qs.get("time_field", ["apply_time"])[0]
        if time_field not in ("apply_time", "success_time", "order_time"):
            raise ApiError("time_field 仅支持 apply_time/success_time/order_time",
                           code=400, http_status=400)

        def money_param(key):
            raw = qs.get(key, [None])[0]
            if raw in (None, ""):
                return None
            try:
                from decimal import Decimal
                return Decimal(raw)
            except Exception:
                raise ApiError(f"{key} 必须是数字", code=400, http_status=400)

        start, end = self._time_filter(qs)
        return {
            "order_id": qs.get("order_id", [None])[0] or None,
            "status": qs.get("status", [None])[0] or None,
            "keyword": qs.get("keyword", [None])[0] or None,
            "start": start, "end": end,
            "min_amount": money_param("min_amount"),
            "max_amount": money_param("max_amount"),
            "time_field": time_field,
            "order_by": order_by, "order_dir": order_dir,
            "page": page, "size": size,
        }

    def _handle_orders(self, handler, parsed, kind, bundle):
        params = self._parse_list_params(parse_qs(parsed.query), kind)

        def data_fn(b, _sid, _sname, page, size):
            p = dict(params, page=page, size=size)
            if kind == "normal":
                p = {k: v for k, v in p.items()
                     if k not in ("min_amount", "max_amount", "time_field")}
                return b.queries.list_normal_orders(**p)
            return b.queries.list_refund_orders(**p)

        self._scoped_list(handler, parsed, data_fn, "orders")

    def _handle_orders_export(self, handler, parsed, kind, bundle):
        params = self._parse_list_params(parse_qs(parsed.query), kind)
        if kind == "normal":
            params = {k: v for k, v in params.items()
                      if k not in ("min_amount", "max_amount", "time_field")}
        else:
            params = {k: v for k, v in params.items() if k not in ("page", "size")}
        operator = self._operator(handler)

        def export_fn(b):
            return b.queries.export_excel(kind, operator=operator, **params)

        self._scoped_export(handler, parsed, export_fn)

    # ---------- 退款明细（逐笔） ----------
    def _refund_detail_filters(self, qs):
        start, end = self._time_filter(qs)
        time_field = qs.get("time_field", ["apply_time"])[0]
        if time_field not in ("apply_time", "success_time"):
            raise ApiError("time_field 仅支持 apply_time/success_time",
                           code=400, http_status=400)
        return {
            "time_field": time_field, "start": start, "end": end,
            "order_id": qs.get("order_id", [None])[0] or None,
            "refund_status": qs.get("status", [None])[0] or None,
            "keyword": qs.get("keyword", [None])[0] or None,
        }

    def _handle_refund_details(self, handler, parsed, bundle):
        filters = self._refund_detail_filters(parse_qs(parsed.query))

        def data_fn(b, _sid, _sname, page, size):
            return b.queries.list_refund_details(
                page=page, size=size, **filters)

        self._scoped_list(handler, parsed, data_fn, "details")

    def _handle_refund_details_export(self, handler, parsed, bundle):
        filters = self._refund_detail_filters(parse_qs(parsed.query))
        operator = self._operator(handler)

        def export_fn(b):
            return b.queries.export_refund_details_excel(
                operator=operator, **filters)

        self._scoped_export(handler, parsed, export_fn)

    # ---------- 供应商成本 ----------
    def _supplier_cost_filters(self, qs):
        start, end = self._time_filter(qs)
        time_field = qs.get("time_field", ["created_at"])[0]
        if time_field not in ("created_at", "ship_time"):
            raise ApiError("time_field 仅支持 created_at/ship_time",
                           code=400, http_status=400)
        return {
            "time_field": time_field, "start": start, "end": end,
            "keyword": qs.get("keyword", [None])[0] or None,
            "match_status": qs.get("match_status", [None])[0] or None,
        }

    def _handle_supplier_costs(self, handler, parsed, bundle):
        filters = self._supplier_cost_filters(parse_qs(parsed.query))

        def data_fn(b, _sid, _sname, page, size):
            return b.supplier_queries.list_costs(
                page=page, size=size, **filters)

        self._scoped_list(handler, parsed, data_fn, "costs")

    def _handle_supplier_costs_export(self, handler, parsed, bundle):
        filters = self._supplier_cost_filters(parse_qs(parsed.query))
        operator = self._operator(handler)

        def export_fn(b):
            return b.supplier_queries.export_excel(operator=operator, **filters)

        self._scoped_export(handler, parsed, export_fn)

    # ---------- 异常处理 ----------
    def _anomaly_filters(self, qs):
        start, end = self._time_filter(qs)
        return {
            "start": start, "end": end,
            "status": qs.get("status", [None])[0] or None,
            "exception_type": qs.get("exception_type", [None])[0] or None,
            "severity": qs.get("severity", [None])[0] or None,
            "order_id": qs.get("order_id", [None])[0] or None,
            "keyword": qs.get("keyword", [None])[0] or None,
        }

    def _handle_anomalies(self, handler, parsed, bundle):
        filters = self._anomaly_filters(parse_qs(parsed.query))

        def data_fn(b, _sid, _sname, page, size):
            return b.anomaly_queries.list_anomalies(
                page=page, size=size, **filters)

        self._scoped_list(handler, parsed, data_fn, "anomalies")

    def _handle_anomalies_export(self, handler, parsed, bundle):
        filters = self._anomaly_filters(parse_qs(parsed.query))
        operator = self._operator(handler)

        def export_fn(b):
            return b.anomaly_queries.export_excel(operator=operator, **filters)

        self._scoped_export(handler, parsed, export_fn)

    def _handle_anomalies_options(self, handler, parsed, bundle):
        self._scoped_options(
            handler, parsed, lambda b: b.anomaly_queries.get_options())

    def _handle_order_detail(self, handler, order_id, bundle):
        def fn(b):
            return b.queries.get_order(order_id)
        self._scoped_find(handler, None, fn,
                          message=f"订单不存在: {order_id}")

    def _handle_order_refunds(self, handler, order_id, bundle):
        def fn(b):
            refunds = b.queries.get_order_refunds(order_id)
            return {"total": len(refunds), "refunds": refunds}
        self._scoped_find(handler, None, fn,
                          message=f"订单不存在: {order_id}")

    def _handle_import_batches(self, handler, parsed, bundle):
        qs = parse_qs(parsed.query)
        start, end = self._time_filter(qs)
        source_type = qs.get("source_type", [None])[0]
        keyword = qs.get("keyword", [None])[0] or None

        def data_fn(b, _sid, _sname, page, size):
            result = b.importer.list_batches(
                source_type=source_type, start=start, end=end, keyword=keyword,
                page=page, size=size)
            # store_name 快照优先（导入时 target_store_name）, 存量批次回退当前店铺名
            for rec in result.get("batches") or []:
                if isinstance(rec, dict):
                    rec.setdefault("store_name",
                                   rec.get("target_store_name") or _sname)
            return result

        self._scoped_list(handler, parsed, data_fn, "batches")

    def _handle_import_batch_detail(self, handler, parsed, batch_id, bundle):
        scope = self._query_scope(handler, parsed)
        b, sid, sname = self._find_batch_store(scope, batch_id)
        data = b.importer.get_batch(batch_id)
        # store_name 快照优先（导入时 target_store_name）, 存量批次回退当前店铺名
        data.setdefault("store_name",
                        data.get("target_store_name") or sname)
        data.setdefault("store_id", sid)
        handler._send(ApiResponse.ok(data))

    def _handle_import_batch_rows(self, handler, parsed, batch_id, bundle):
        qs = parse_qs(parsed.query)
        try:
            page = int(qs.get("page", ["1"])[0] or 1)
            size = int(qs.get("size", ["20"])[0] or 20)
        except ValueError:
            raise ApiError("page/size 必须是整数", code=400, http_status=400)
        status = qs.get("status", [None])[0]
        scope = self._query_scope(handler, parsed)
        b, _sid, _sn = self._find_batch_store(scope, batch_id)
        result = b.importer.get_batch_rows(batch_id, status=status,
                                           page=page, size=size)
        handler._send(ApiResponse.ok(result))

    def _handle_import_errors_export(self, handler, parsed, batch_id, bundle):
        scope = self._query_scope(handler, parsed)
        b, _sid, _sn = self._find_batch_store(scope, batch_id)
        filename, content = b.importer.export_errors(batch_id)
        b.op_log.append("import.export_errors", detail={"batch_id": batch_id},
                        operator=self._operator(handler))
        handler._send_bytes(EXCEL_MIME, content, filename)

    # ---------- 多店铺管理与导入批次撤销（管理员） ----------
    def _handle_stores(self, handler):
        """店铺列表: 管理员返回全部成员店铺（含 owner 字段）, 成员返回本人全部有效店铺"""
        _, session, user = self._resolve_session(handler)
        if user.get("role") == "admin":
            records = self.user_manager.member_stores()
        else:
            records = self._member_active_stores(user)
        users = {u.get("user_id"): u for u in self.user_manager.list_users()}
        stores = []
        for s in records:
            sid = s.get("store_id")
            if not sid:
                continue
            b = self.registry.get(sid)
            counts = {}
            for name in ("orders", "supplier_costs", "refund_details",
                         "import_records"):
                try:
                    counts[name] = len(b.repo.read(name)["records"])
                except (RepositoryError, OSError, ValueError):
                    counts[name] = 0
            owner = users.get(s.get("owner_user_id"))
            stores.append({
                "store_id": sid,
                "name": s.get("name") or sid,
                "owner_user_id": s.get("owner_user_id"),
                "owner_username": owner.get("username") if owner else None,
                "created_at": s.get("created_at"),
                "updated_at": s.get("updated_at"),
                "record_counts": counts,
            })
        handler._send(ApiResponse.ok({"total": len(stores), "stores": stores}))

    def _handle_store_rename(self, handler, body, store_id):
        """店铺重命名（仅管理员）: store_id 与数据归属不变"""
        _, operator = self._require_admin(handler)
        new_name = body.get("name")
        if not isinstance(new_name, str):
            raise ApiError("请提供店铺名称（name）", code=400, http_status=400)
        old_name, applied = self.user_manager.rename_store(operator, store_id, new_name)
        detail = {"store_id": store_id, "old_name": old_name,
                  "new_name": applied}
        target_bundle = self.registry.get(store_id)
        # operator 此处为用户 dict（_require_admin 返回 user），写入日志必须
        # 落 username 字符串：否则 JSONL 中 operator 成为整个用户对象
        # （含 password_hash），且 /api/logs 操作人去重集合抛 TypeError。
        target_bundle.op_log.append(
            "store.rename", detail=detail,
            operator=operator.get("username"))
        handler._send(ApiResponse.ok({
            "store_id": store_id, "old_name": old_name, "name": applied,
        }, message="店铺已重命名"))

    # ---------- 店铺归属管理（仅管理员） ----------
    @staticmethod
    def _store_record_counts(bundle):
        """店铺四类业务记录数（与 /api/stores 口径一致）"""
        counts = {}
        for name in ("orders", "supplier_costs", "refund_details",
                     "import_records"):
            try:
                counts[name] = len(bundle.repo.read(name)["records"])
            except (RepositoryError, OSError, ValueError):
                counts[name] = 0
        return counts

    def _store_inventory(self, store_id):
        """单店十项清点（spec 5.6.1.4）: 各业务文件记录数与金额聚合, 只读"""
        b = self.registry.get(store_id)

        def _records(name):
            try:
                return b.repo.read(name).get("records", [])
            except (RepositoryError, OSError, ValueError):
                return []

        def _sum_field(records, field):
            total = Decimal("0")
            for r in records:
                try:
                    total += Decimal(str(r.get(field) or "0"))
                except (ArithmeticError, ValueError, TypeError):
                    continue
            return str(total.quantize(Decimal("0.01")))

        refund_details = _records("refund_details")
        costs = _records("supplier_costs")
        return {
            "orders": len(_records("orders")),
            "refund_orders": len(_records("refund_orders")),
            "refund_details": len(refund_details),
            "refund_amount": _sum_field(refund_details, "refund_amount"),
            "supplier_costs": len(costs),
            "cost_amount": _sum_field(costs, "total_amount"),
            "ad_expenses": len(_records("ad_expenses")),
            "profit_results": len(_records("profit_statements")),
            "import_records": len(_records("import_records")),
            "anomalies": len(_records("anomalies")),
        }

    def _store_delete_inventory(self, store_id):
        """物理删除范围完整清点（design 2.2 接口三）: 十项记录数
        + 4 类补齐（raw_records/export_records/supplier_cost_items/
        reconciliation_results）+ uploads/店铺级备份/操作日志行数,
        只读无副作用; other_total 供删除弹窗"其他关联数据"展示。"""
        b = self.registry.get(store_id)

        def _count(name):
            try:
                return len(b.repo.read(name).get("records", []))
            except (RepositoryError, OSError, ValueError):
                return 0

        def _walk_count(root):
            if not os.path.isdir(root):
                return 0
            return sum(len(files) for _dp, _dc, files in os.walk(root))

        op_log_file = os.path.join(b.log_dir, "operation_log.jsonl")
        op_lines = 0
        if os.path.isfile(op_log_file):
            with open(op_log_file, encoding="utf-8") as f:
                op_lines = sum(1 for line in f if line.strip())
        inventory = {
            "orders": _count("orders"),
            "refund_orders": _count("refund_orders"),
            "refund_details": _count("refund_details"),
            "supplier_costs": _count("supplier_costs"),
            "reconciliation_results": _count("reconciliation_results"),
            "profit_results": _count("profit_statements"),
            "import_records": _count("import_records"),
            "ad_expenses": _count("ad_expenses"),
            "anomalies": _count("anomalies"),
            "raw_records": _count("raw_records"),
            "export_records": _count("export_records"),
            "supplier_cost_items": _count("supplier_cost_items"),
            "upload_files": _walk_count(b.upload_dir),
            "store_backup_files": _walk_count(b.backup_dir),
            "operation_log_lines": op_lines,
        }
        inventory["other_total"] = sum(
            inventory[key] for key in (
                "raw_records", "export_records", "supplier_cost_items",
                "upload_files", "store_backup_files", "operation_log_lines"))
        return inventory

    def _store_extra_metrics(self, store_id):
        """店铺详情追加只读指标（不进 inventory, 保持既有校验协议不变）:
        success_refund_amount 成功退款金额（partial/full 口径对齐 SUCCESS_STATUSES）;
        pending_refund_amount 处理中退款金额（refund_status == "pending"）;
        last_import_at 最近导入批次时间; 均向后兼容, 无记录时为 None。"""
        b = self.registry.get(store_id)

        def _records(name):
            try:
                return b.repo.read(name).get("records", [])
            except (RepositoryError, OSError, ValueError):
                return []

        def _sum_field(records, field):
            total = Decimal("0")
            for r in records:
                try:
                    total += Decimal(str(r.get(field) or "0"))
                except (ArithmeticError, ValueError, TypeError):
                    continue
            return str(total.quantize(Decimal("0.01")))

        refund_details = _records("refund_details")
        import_records = _records("import_records")
        success = [r for r in refund_details
                   if r.get("refund_status") in SUCCESS_STATUSES]
        pending = [r for r in refund_details
                   if r.get("refund_status") == "pending"]
        return {
            "success_refund_amount": _sum_field(success, "refund_amount"),
            "pending_refund_amount": _sum_field(pending, "refund_amount"),
            "last_import_at": (import_records[-1].get("created_at")
                               if import_records else None),
        }

    def _backup_ownership_change(self, store_id):
        """归属变更前备份: data/ 下 ownership_migration_<时间戳>/

        目录含 users/stores 两表副本与整店快照; 校验快照文件清单与源一致,
        备份失败立即中止（零归属变更）。
        """
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        backup_dir = os.path.join(self.data_root, f"ownership_migration_{ts}")
        snapshot_dir = os.path.join(backup_dir, "store_snapshot")
        store_dir = os.path.join(self.stores_root, store_id)
        if not os.path.isdir(store_dir):
            raise ApiError(f"店铺目录不存在: {store_id}",
                           code=404, http_status=404)

        def _inventory(root):
            items = {}
            for dirpath, _dirs, files in os.walk(root):
                for fn in files:
                    fp = os.path.join(dirpath, fn)
                    items[os.path.relpath(fp, root)] = os.path.getsize(fp)
            return items

        try:
            os.makedirs(backup_dir)
            for name, path in (("users.json", self.user_manager.users_path),
                               ("stores.json", self.user_manager.stores_path)):
                if os.path.exists(path):
                    shutil.copy2(path, os.path.join(backup_dir, name))
            shutil.copytree(store_dir, snapshot_dir)
            if _inventory(store_dir) != _inventory(snapshot_dir):
                raise OSError("整店快照与源目录文件清单不一致")
        except Exception as e:
            shutil.rmtree(backup_dir, ignore_errors=True)
            raise ApiError(f"迁移前备份失败，已中止且未做任何变更: {e}",
                           code=500, http_status=500) from e
        return {"backup_dir": backup_dir, "snapshot_dir": snapshot_dir}

    def _audit_ownership_change(self, handler, operator, store_id,
                                store_name, target_user_id, result,
                                detail_extra=None):
        """店铺归属变更审计公共写入（安全日志, 只追加不覆盖）

        固定字段: 操作人（username/user_id 来自服务端会话, 不可前端伪造）、
        操作人角色、店铺 ID、店铺名称、目标成员 ID; 业务结果与失败原因
        由 detail_extra 补充。返回写入的 detail 便于复用。
        """
        detail = {
            "store_id": store_id,
            "store_name": store_name,
            "operator_role": operator.get("role"),
            "target_user_id": target_user_id,
        }
        if detail_extra:
            detail.update(detail_extra)
        self.security_log.append(
            "store.ownership_change", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            result=result, detail=detail)
        return detail

    def _handle_store_transfer_ownership(self, handler, body, store_id):
        """店铺归属变更（仅管理员）: 幂等短路 → 备份 → 变更 → 吊销原用户会话 → 审计 → 数据核对

        幂等: 新归属成员与当前归属一致且为有效普通成员时, 直接返回
        "归属成员未发生变化"（unchanged=true）, 不备份、不写盘、不吊销会话。
        失败: 业务校验拒绝时归属与业务数据零变更, 审计留痕含失败原因。
        """
        _, operator = self._require_admin(handler)
        target_user_id = body.get("target_user_id")
        store = self.user_manager.get_store(store_id)
        store_name = store.get("name") if store else None
        if not isinstance(target_user_id, str) or not target_user_id:
            self._audit_ownership_change(
                handler, operator, store_id, store_name, None, "error",
                {"reason": "请提供目标成员（target_user_id）"})
            raise ApiError("请提供目标成员（target_user_id）",
                           code=400, http_status=400)
        if store is None:
            self._audit_ownership_change(
                handler, operator, store_id, None, target_user_id, "error",
                {"reason": f"店铺不存在: {store_id}"})
            raise ApiError(f"店铺不存在: {store_id}", code=404, http_status=404)
        prev_owner_id = store.get("owner_user_id")
        # 幂等短路: 目标即当前归属且为有效普通成员 → 零写入零副作用
        if prev_owner_id == target_user_id:
            target = self.user_manager.get_user(target_user_id)
            if (target is not None and target.get("role") == "member"
                    and target.get("status") == "active"):
                owner_label = {"user_id": target.get("user_id"),
                               "username": target.get("username")}
                result = {
                    "store_id": store_id,
                    "store_name": store_name,
                    "previous_owner": dict(owner_label),
                    "new_owner": dict(owner_label),
                    "record_counts": self._store_record_counts(
                        self.registry.get(store_id)),
                    "data_preserved": True,
                    "unchanged": True,
                }
                self._audit_ownership_change(
                    handler, operator, store_id, store_name, target_user_id,
                    "noop", {
                        "previous_owner": result["previous_owner"],
                        "new_owner": result["new_owner"],
                        "unchanged": True,
                        "reason": "归属成员未发生变化",
                        "record_counts": result["record_counts"],
                    })
                self.registry.get(store_id).op_log.append(
                    "store.ownership_change",
                    detail={"previous_owner": result["previous_owner"],
                            "new_owner": result["new_owner"],
                            "operator_role": operator.get("role"),
                            "unchanged": True},
                    operator=operator.get("username"))
                handler._send(ApiResponse.ok(
                    result, message="归属成员未发生变化"))
                return
        inventory_before = self._store_inventory(store_id)
        backup = None
        try:
            backup = self._backup_ownership_change(store_id)
            result = self.user_manager.transfer_store_ownership(
                operator, store_id, target_user_id)
        except (AuthError, ApiError) as e:
            if backup is not None:
                shutil.rmtree(backup["backup_dir"], ignore_errors=True)
            self._audit_ownership_change(
                handler, operator, store_id, store_name, target_user_id,
                "error", {"reason": str(e),
                          "previous_owner_id": prev_owner_id})
            raise
        # 原归属用户绑定该店的活跃会话立即失效
        revoked = (self.sessions.revoke_user(prev_owner_id)
                   if prev_owner_id else 0)
        # 迁移核对: 十项清点与备份前逐项对照, 不一致 → 从备份恢复两表
        inventory_after = self._store_inventory(store_id)
        if inventory_after != inventory_before:
            for name, path in (("users.json", self.user_manager.users_path),
                               ("stores.json", self.user_manager.stores_path)):
                src = os.path.join(backup["backup_dir"], name)
                if os.path.exists(src):
                    shutil.copy2(src, path)
            self.registry.evict(store_id)
            diff = {k: [inventory_before[k], inventory_after.get(k)]
                    for k in inventory_before
                    if inventory_before[k] != inventory_after.get(k)}
            self._audit_ownership_change(
                handler, operator, store_id, store_name, target_user_id,
                "error", {
                    "reason": "迁移核对失败，已从备份恢复店铺归属",
                    "previous_owner": result.get("previous_owner"),
                    "new_owner": result.get("new_owner"),
                    "backup_dir": backup["backup_dir"],
                    "recovered": True, "diff": diff})
            raise ApiError("迁移核对失败，已从备份恢复店铺归属，请重试",
                           code=500, http_status=500)
        counts = self._store_record_counts(self.registry.get(store_id))
        result["record_counts"] = counts
        result["data_preserved"] = True
        result["unchanged"] = False
        self._audit_ownership_change(
            handler, operator, store_id, store_name, target_user_id,
            "success", {
                "previous_owner": result.get("previous_owner"),
                "new_owner": result.get("new_owner"),
                "record_counts": counts,
                "backup_dir": backup["backup_dir"],
                "revoked_sessions": revoked})
        self.registry.get(store_id).op_log.append(
            "store.ownership_change",
            detail={"store_id": store_id,
                    "store_name": store_name,
                    "operator_role": operator.get("role"),
                    "previous_owner": result.get("previous_owner"),
                    "new_owner": result.get("new_owner"),
                    "backup_dir": backup["backup_dir"]},
            operator=operator.get("username"))
        handler._send(ApiResponse.ok(result, message="店铺归属已变更"))

    def _handle_store_mark_unassigned(self, handler, body, store_id):
        """标记店铺待分配（仅管理员）: 备份 → 标记 → 吊销原归属会话 → 审计"""
        _, operator = self._require_admin(handler)
        store = self.user_manager.get_store(store_id)
        if store is None:
            raise ApiError(f"店铺不存在: {store_id}", code=404, http_status=404)
        prev_owner_id = store.get("owner_user_id")
        already = (store.get("assignment_status") == "unassigned"
                   and not prev_owner_id)
        backup = None
        try:
            if not already:
                backup = self._backup_ownership_change(store_id)
            result = self.user_manager.mark_store_unassigned(
                operator, store_id)
        except AuthError:
            if backup is not None:
                shutil.rmtree(backup["backup_dir"], ignore_errors=True)
            raise
        revoked = self.sessions.revoke_user(prev_owner_id) \
            if prev_owner_id else 0
        result["record_counts"] = self._store_record_counts(
            self.registry.get(store_id))
        self.security_log.append(
            "store.ownership_change", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            detail={"store_id": store_id,
                    "previous_owner": result.get("previous_owner"),
                    "new_owner": "待分配",
                    "record_counts": result["record_counts"],
                    "backup_dir": backup["backup_dir"] if backup else None,
                    "revoked_sessions": revoked})
        self.registry.get(store_id).op_log.append(
            "store.ownership_change",
            detail={"new_owner": "待分配",
                    "previous_owner": result.get("previous_owner")},
            operator=operator.get("username"))
        handler._send(ApiResponse.ok(result, message="店铺已标记为待分配"))

    # ---------- 店铺生命周期管理（仅管理员） ----------
    _STORE_STATUS_MESSAGES = {
        "disable": "店铺已停用",
        "enable": "店铺已启用",
    }

    def _handle_store_create(self, handler, body):
        """创建新店铺并绑定归属成员（仅管理员; 一对多追加绑定）"""
        _, operator = self._require_admin(handler)
        store = self.user_manager.create_store(
            operator, body.get("name"), body.get("owner_user_id"))
        owner = self.user_manager.get_user(store["owner_user_id"])
        detail = {
            "store_id": store["store_id"],
            "name": store["name"],
            "owner_user_id": store["owner_user_id"],
            "owner_username": owner.get("username") if owner else None,
        }
        self.security_log.append(
            "store.create", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            detail=detail)
        # registry.get 惰性创建店铺目录结构并初始化业务文件
        self.registry.get(store["store_id"]).op_log.append(
            "store.create", detail=detail,
            operator=operator.get("username"))
        handler._send(ApiResponse.ok({
            "store_id": store["store_id"],
            "name": store["name"],
            "owner_user_id": store["owner_user_id"],
            "owner_username": owner.get("username") if owner else None,
            "status": store["status"],
            "assignment_status": store["assignment_status"],
            "created_at": store["created_at"],
            "updated_at": store["updated_at"],
        }, message="店铺创建成功"))

    def _handle_store_detail(self, handler, store_id):
        """店铺详情与数据概况（仅管理员, 只读）

        十项清点复用 _store_inventory: 删除确认弹窗要素、
        店铺管理页"查看数据"概况、停用店铺历史数据追溯;
        delete_inventory 为物理删除范围权威清点（只读, 无副作用）。
        """
        self._require_admin(handler)
        store = self.user_manager.get_store(store_id)
        if store is None:
            raise ApiError(f"店铺不存在: {store_id}",
                           code=404, http_status=404)
        owner_id = store.get("owner_user_id")
        owner = self.user_manager.get_user(owner_id) if owner_id else None
        orphan = bool(owner_id) and owner is None
        handler._send(ApiResponse.ok({
            "store_id": store_id,
            "name": store.get("name") or store_id,
            "owner_user_id": owner_id,
            "owner_username": owner.get("username") if owner else None,
            "status": store.get("status") or "active",
            "assignment_status":
                store.get("assignment_status") or "assigned",
            "orphan": orphan,
            "created_at": store.get("created_at"),
            "updated_at": store.get("updated_at"),
            "delete_inventory": self._store_delete_inventory(store_id),
            "inventory": self._store_inventory(store_id),
            **self._store_extra_metrics(store_id),
        }))

    def _handle_store_status(self, handler, body, store_id):
        """店铺停用/启用（仅管理员; 归档功能已下线）

        幂等先行（目标状态已达成 → 直接返回, 不备份不重复审计）
        → 备份两表 → 变更 → 吊销归属成员会话（disable）→ 审计。
        """
        _, operator = self._require_admin(handler)
        action = body.get("action")
        if action == "archive":
            # 归档功能已下线: 显式拒绝并审计留痕, 零数据变更
            self.security_log.append(
                "store.archive_removed", username=operator.get("username"),
                user_id=operator["user_id"], ip=self._client_ip(handler),
                result="denied", detail={"store_id": store_id,
                                         "action": action})
            raise ApiError("归档功能已下线，不支持该操作",
                           code=400, http_status=400)
        if action not in self._STORE_STATUS_MESSAGES:
            raise ApiError("action 仅支持 disable / enable（归档功能已下线）",
                           code=400, http_status=400)
        store = self.user_manager.get_store(store_id)
        if store is None:
            raise ApiError(f"店铺不存在: {store_id}",
                           code=404, http_status=404)
        current = store.get("status") or "active"
        target = {"disable": "disabled", "enable": "active"}[action]
        owner_id = store.get("owner_user_id")
        owner = self.user_manager.get_user(owner_id) if owner_id else None
        data = {
            "store_id": store_id,
            "name": store.get("name") or store_id,
            "status": current,
            "previous_status": current,
            "owner_user_id": owner_id,
            "owner_username": owner.get("username") if owner else None,
            "record_counts": self._store_record_counts(
                self.registry.get(store_id)),
            "data_preserved": True,
            "revoked_sessions": 0,
            "backup_dir": None,
        }
        if current == target:
            data["idempotent"] = True
            handler._send(ApiResponse.ok(
                data, message=self._STORE_STATUS_MESSAGES[action]))
            return
        backup = None
        try:
            backup = self._backup_ownership_change(store_id)
            result = self.user_manager.set_store_status(
                operator, store_id, action)
        except AuthError:
            if backup is not None:
                shutil.rmtree(backup["backup_dir"], ignore_errors=True)
            raise
        # disable → 归属成员绑定该店的活跃会话立即失效; enable 不吊销
        revoked = (self.sessions.revoke_user(owner_id)
                   if owner_id and action == "disable" else 0)
        data.update({
            "status": result["status"],
            "previous_status": result["previous_status"],
            "updated_at": result["updated_at"],
            "revoked_sessions": revoked,
            "backup_dir": backup["backup_dir"],
        })
        audit_detail = {
            "store_id": store_id,
            "store_name": data["name"],
            "status": result["status"],
            "previous_status": result["previous_status"],
            "owner_user_id": owner_id,
            "record_counts": data["record_counts"],
            "data_preserved": True,
            "revoked_sessions": revoked,
            "backup_dir": backup["backup_dir"],
        }
        self.security_log.append(
            f"store.{action}", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            detail=audit_detail)
        self.registry.get(store_id).op_log.append(
            f"store.{action}",
            detail={"status": result["status"],
                    "previous_status": result["previous_status"]},
            operator=operator.get("username"))
        handler._send(ApiResponse.ok(
            data, message=self._STORE_STATUS_MESSAGES[action]))

    def _handle_store_delete(self, handler, body, store_id):
        """物理删除店铺及全部关联业务数据（仅管理员; store_physical_delete）

        编排: 鉴权 → 存在性 → 强确认双通道（完整名称或 store_id + 勾选）
        → 服务端权威清点 → 归属完整性检查 → 复验 → 删除事务（互斥创建
        → 吊销会话 → 清业务文件 → 删两表 → 移除目录 → 回读核对 →
        committed）→ 备份治理（墓碑）→ 审计 → 响应真实删除统计。
        任一环节失败自动回滚到删除前状态, 绝不部分删除、绝不虚报成功。
        """
        _, operator = self._require_admin(handler)
        store = self.user_manager.get_store(store_id)
        if store is None:
            raise ApiError(f"店铺不存在: {store_id}",
                           code=404, http_status=404)
        # 强确认双通道: 完整店铺名称或店铺 ID + 勾选确认
        # （严格布尔判定, 字符串 "true" 拒绝; 零数据变更）
        confirm = body.get("confirm_store_name")
        if (not isinstance(confirm, str)
                or (confirm != store.get("name") and confirm != store_id)
                or body.get("confirm_acknowledged") is not True):
            raise ApiError("确认信息不一致，请核对后重试",
                           code=400, http_status=400)
        # 服务端权威清点（后端口径固定, 不信任前端传入的任何计数）
        inventory = self._store_delete_inventory(store_id)
        # 归属完整性检查: 记录存在但目录缺失 → 阻断（不得删除归属不明数据）
        store_dir = os.path.join(self.stores_root, store_id)
        if not os.path.isdir(store_dir):
            raise ApiError(f"店铺数据目录缺失: {store_id}",
                           code=404, http_status=404)
        # 删除前复验: 权限已验、店铺仍存在且名称未变（重读重校）
        store_recheck = self.user_manager.get_store(store_id)
        if (store_recheck is None
                or store_recheck.get("name") != store.get("name")):
            raise ApiError(f"店铺不存在: {store_id}",
                           code=404, http_status=404)
        owner_id = store.get("owner_user_id")
        owner = self.user_manager.get_user(owner_id) if owner_id else None
        txn_dir = None
        try:
            txn_dir = self._create_delete_txn(
                store_id, store, operator, inventory,
                client_ip=self._client_ip(handler))
            result = self._run_delete_txn(
                store_id, txn_dir, operator, store)
        except ApiError as e:
            # 409 冲突已在事务创建处审计; 其余失败路径统一留痕
            if e.http_status != 409:
                self.security_log.append(
                    "store.delete_failed",
                    username=operator.get("username"),
                    user_id=operator["user_id"],
                    ip=self._client_ip(handler), result="error",
                    detail={
                        "store_id": store_id,
                        "store_name": store.get("name"),
                        "owner_user_id": owner_id,
                        "error": str(e),
                        "txn_id": (os.path.basename(txn_dir)
                                   if txn_dir else None)})
            raise
        # committed → 终态统计 → 备份治理（墓碑）→ 审计 → 响应
        counts = self._deleted_counts(inventory, result)
        txn_id = os.path.basename(txn_dir)
        self._advance_delete_txn(txn_dir, "committed",
                                 deleted_counts=counts)
        issues = self._govern_store_backups(store_id, txn_dir)
        if issues:
            backup_treatment = ("备份清理失败: " + "; ".join(issues))
        else:
            backup_treatment = (
                "业务备份已治理：全局快照目录已清理，墓碑已写入；"
                "灾难恢复级历史目录（ownership_migration_*）按策略保留，"
                "不承诺从中清除")
        self.security_log.append(
            "store.delete", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            detail={
                "store_id": store_id,
                "store_name": store.get("name"),
                "owner_user_id": owner_id,
                "owner_username": (owner.get("username")
                                   if owner else None),
                "deleted_counts": counts,
                "revoked_sessions": result["revoked"],
                "removed_store_dir": result["removed_dir"],
                "backup_treatment": backup_treatment,
                "txn_id": txn_id,
            })
        handler._send(ApiResponse.ok({
            "store_id": store_id,
            "name": store.get("name"),
            "owner_user_id": owner_id,
            "owner_username": owner.get("username") if owner else None,
            "deleted_counts": counts,
            "revoked_sessions": result["revoked"],
            "removed_store_dir": result["removed_dir"],
            "backup_treatment": backup_treatment,
            "txn_id": txn_id,
        }, message="店铺已永久删除（含全部关联业务数据）"))

    def _deleted_counts(self, inventory, result):
        """删除统计（实际删除动作回读 + 清点口径, 不用前端计数）"""
        records = result["records_before"]
        counts = {}
        for name in CLEARABLE_DATA_FILES:
            key = ("profit_results" if name == "profit_statements"
                   else name)
            value = records.get(name)
            counts[key] = inventory.get(key, 0) if value is None else value
        counts["upload_files"] = result["uploads_removed"]
        counts["store_backup_files"] = inventory.get("store_backup_files", 0)
        counts["operation_log_lines"] = inventory.get("operation_log_lines", 0)
        counts["settings"] = 1
        return counts

    def _handle_store_clear_data(self, handler, body, store_id):
        """清除单店铺业务数据（仅管理员; 独立清数据入口）

        单店维度: 与按用户维度的 /api/users/{id}/store-data-clear 并存
        （后者清目标成员名下全部店铺, 职责边界见 design 2.7.3）。
        保留 settings/操作日志/店铺级备份与店铺记录本身。
        强确认口径与物理删除一致（名称或 store_id 双通道 + 勾选）。
        """
        _, operator = self._require_admin(handler)
        store = self.user_manager.get_store(store_id)
        if store is None:
            raise ApiError(f"店铺不存在: {store_id}",
                           code=404, http_status=404)
        confirm = body.get("confirm_store_name")
        if (not isinstance(confirm, str)
                or (confirm != store.get("name") and confirm != store_id)
                or body.get("confirm_acknowledged") is not True):
            raise ApiError("确认信息不一致，请核对后重试",
                           code=400, http_status=400)
        owner_id = store.get("owner_user_id")
        owner = self.user_manager.get_user(owner_id) if owner_id else None
        snapshot_dir, _files = self._snapshot_store(store_id)
        # 快照就绪后吊销归属成员会话, 避免清理期间并发写入
        revoked = self.sessions.revoke_user(owner_id) if owner_id else 0
        try:
            stats = self._clear_store_data_files(store_id, snapshot_dir)
        except ApiError:
            raise  # _clear_store_data_files 已从快照回滚, 数据保持原状
        self.security_log.append(
            "store.clear_data", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            detail={"store_id": store_id,
                    "store_name": store.get("name"),
                    "owner_user_id": owner_id,
                    "records_before": stats["records_before"],
                    "uploads_removed": stats["uploads_removed"],
                    "backup_dir": snapshot_dir,
                    "revoked_sessions": revoked})
        # 店铺记录保留, 操作日志留痕（缓存已驱逐, get 重建后追加）
        self.registry.get(store_id).op_log.append(
            "store.clear_data",
            detail={"records_before": stats["records_before"],
                    "backup_dir": snapshot_dir},
            operator=operator.get("username"))
        handler._send(ApiResponse.ok({
            "store_id": store_id,
            "store_name": store.get("name"),
            "records_before": stats["records_before"],
            "uploads_removed": stats["uploads_removed"],
            "backup_dir": snapshot_dir,
            "revoked_sessions": revoked,
        }, message="店铺业务数据已清除（已创建整店快照备份）"))

    def _handle_admin_all_stores(self, handler, parsed=None):
        """店铺管理全量视图（仅管理员）: 全部店铺 + 派生归属状态, 只读

        可选筛选分页（nav_user_mgmt）: name 模糊 / owner（用户名或 user_id）/
        assignment_status / page / page_size; 不传参数时维持全量返回。
        """
        self._require_admin(handler)
        users = {u.get("user_id"): u for u in self.user_manager.list_users()}
        stores = []
        for s in self.user_manager.list_stores():
            sid = s.get("store_id")
            if not sid:
                continue
            owner_id = s.get("owner_user_id")
            owner = users.get(owner_id) if owner_id else None
            if (s.get("assignment_status") == "unassigned"
                    or not owner_id or owner is None):
                status = "unassigned"
            elif owner.get("role") == "admin":
                status = "admin_legacy"
            else:
                status = "assigned"
            entry = {
                "store_id": sid,
                "name": s.get("name") or sid,
                "owner_user_id": owner_id,
                "owner_username": owner.get("username") if owner else None,
                "assignment_status": status,
                # 店铺运营状态（存量缺字段按 active 派生, 迁移完成前兼容）
                "status": s.get("status") or "active",
                "created_at": s.get("created_at"),
                "updated_at": s.get("updated_at"),
                "record_counts": self._store_record_counts(
                    self.registry.get(sid)),
            }
            if status == "unassigned" and owner_id and owner is None:
                entry["orphan"] = True
            stores.append(entry)
        qs = parse_qs(parsed.query) if parsed is not None else {}
        has_filter = any(k in qs for k in ("name", "owner",
                                           "assignment_status", "store_id",
                                           "status",
                                           "page", "page_size", "size"))
        if has_filter:
            name_kw = (qs.get("name", [None])[0] or "").strip().lower()
            owner_kw = (qs.get("owner", [None])[0] or "").strip()
            assignment_kw = qs.get("assignment_status", [None])[0]
            if assignment_kw is not None and assignment_kw not in (
                    "assigned", "unassigned", "admin_legacy"):
                raise ApiError("归属状态仅支持 assigned / unassigned / admin_legacy",
                               code=400, http_status=400)
            sid_kw = (qs.get("store_id", [None])[0] or "").strip()
            status_kw = qs.get("status", [None])[0]
            if status_kw is not None and status_kw not in (
                    "active", "disabled"):
                raise ApiError("店铺状态仅支持 active / disabled",
                               code=400, http_status=400)
            filtered = [
                e for e in stores
                if (not name_kw or name_kw in (e["name"] or "").lower())
                and (not owner_kw or e["owner_username"] == owner_kw
                     or e["owner_user_id"] == owner_kw)
                and (assignment_kw is None
                     or e["assignment_status"] == assignment_kw)
                and (not sid_kw or sid_kw in e["store_id"])
                and (status_kw is None or e["status"] == status_kw)
            ]
            page, size = self._page_params(qs)
            start = (page - 1) * size
            stores = filtered[start:start + size]
            handler._send(ApiResponse.ok({
                "total": len(filtered), "stores": stores,
                "page": page, "page_size": size}))
            return
        handler._send(ApiResponse.ok({"total": len(stores), "stores": stores}))

    def _handle_admin_migration_overview(self, handler):
        """迁移清点（仅管理员）: admin 遗留与待分配/孤儿店铺十项清点, 只读"""
        self._require_admin(handler)
        users = {u.get("user_id"): u for u in self.user_manager.list_users()}
        admin_stores, unassigned_stores = [], []
        member_stores_count = 0
        for s in self.user_manager.list_stores():
            sid = s.get("store_id")
            if not sid:
                continue
            owner_id = s.get("owner_user_id")
            owner = users.get(owner_id) if owner_id else None
            if (s.get("assignment_status") == "unassigned"
                    or not owner_id or owner is None):
                entry = {
                    "store_id": sid, "name": s.get("name") or sid,
                    "owner_user_id": owner_id,
                    "owner_username": owner.get("username") if owner else None,
                    "assignment_status": "unassigned",
                    "inventory": self._store_inventory(sid),
                }
                if owner_id and owner is None:
                    entry["orphan"] = True
                unassigned_stores.append(entry)
            elif owner.get("role") == "admin":
                admin_stores.append({
                    "store_id": sid, "name": s.get("name") or sid,
                    "owner_user_id": owner_id,
                    "owner_username": owner.get("username"),
                    "assignment_status": "admin_legacy",
                    "inventory": self._store_inventory(sid),
                })
            else:
                member_stores_count += 1
        backup_dir = None
        try:
            candidates = [d for d in os.listdir(self.data_root)
                          if d.startswith(("pre_migration_backup_",
                                           "ownership_migration_"))]
            if candidates:
                latest = max(
                    candidates,
                    key=lambda d: os.path.getmtime(
                        os.path.join(self.data_root, d)))
                backup_dir = os.path.join(self.data_root, latest)
        except OSError:
            pass
        handler._send(ApiResponse.ok({
            "admin_stores": admin_stores,
            "unassigned_stores": unassigned_stores,
            "member_stores_count": member_stores_count,
            "backup_dir": backup_dir,
        }))

    def _migration_model_scan(self):
        """店铺数据模型一致性检查（只读, spec 5.8.1 规则 2）

        逐店铺: 字段缺失清单 / 归属指向有效性（孤儿/admin 遗留）/
        业务数据目录存在性; 歧义（无目录/admin 遗留）与游离业务目录
        列入待核对清单, 不自动处理。check 与 migrate 共用本口径。
        """
        users = {u.get("user_id"): u for u in self.user_manager.list_users()}
        items, ambiguous, business_data_orphan = [], [], []
        known_ids = set()
        for s in self.user_manager.list_stores():
            sid = s.get("store_id")
            if not sid:
                continue
            known_ids.add(sid)
            issues = []
            if not s.get("status"):
                issues.append("missing_status")
            if not s.get("updated_at"):
                issues.append("missing_updated_at")
            if not s.get("assignment_status"):
                issues.append("missing_assignment_status")
            owner_id = s.get("owner_user_id")
            owner = users.get(owner_id) if owner_id else None
            orphan_owner = bool(owner_id) and owner is None
            admin_legacy = bool(owner) and owner.get("role") == "admin"
            has_dir = os.path.isdir(os.path.join(self.stores_root, sid))
            if orphan_owner:
                issues.append("orphan_owner")
            if admin_legacy:
                issues.append("admin_legacy")
            if not has_dir:
                issues.append("no_store_dir")
            fillable = {"missing_status", "missing_updated_at",
                        "missing_assignment_status", "orphan_owner"}
            auto_fixable = (bool(issues) and has_dir and not admin_legacy
                            and all(i in fillable for i in issues))
            if not has_dir or admin_legacy:
                ambiguous.append({
                    "store_id": sid, "name": s.get("name") or sid,
                    "reason": "no_store_dir" if not has_dir else "admin_legacy",
                })
            actions = []
            if any(i.startswith("missing_") for i in issues):
                actions.append("fill_fields")
            if orphan_owner:
                actions.append("unassign_owner")
            items.append({
                "store_id": sid,
                "name": s.get("name") or sid,
                "issues": issues,
                "action": "+".join(actions) if actions else "none",
                "auto_fixable": auto_fixable,
            })
        try:
            extra_dirs = [d for d in os.listdir(self.stores_root)
                          if d not in known_ids
                          and os.path.isdir(
                              os.path.join(self.stores_root, d))]
        except OSError:
            extra_dirs = []
        for d in extra_dirs:
            business_data_orphan.append({"store_id": d, "name": None})
        return {
            "total_stores": len(items),
            "items": items,
            "ambiguous": ambiguous,
            "business_data_orphan": business_data_orphan,
        }

    def _handle_migration_model_check(self, handler):
        """店铺模型迁移前检查（仅管理员, 只读）: design 2.2.2 A6"""
        self._require_admin(handler)
        handler._send(ApiResponse.ok(self._migration_model_scan()))

    def _handle_migration_model_migrate(self, handler, body):
        """店铺模型迁移执行（仅管理员）: design 2.2.2 A7 / 2.4.2

        备份两表 → 字段补齐 + 孤儿归属置空（不改 store_id/其余归属/
        业务数据）→ 十项清点与归属前后核对（不一致自动回滚两表）→ 审计。
        业务数据目录 data/stores/ 全程只读。
        """
        _, operator = self._require_admin(handler)
        report = self._migration_model_scan()
        fixable = [it for it in report["items"] if it["auto_fixable"]]
        skipped = [{"store_id": it["store_id"], "name": it["name"],
                    "issues": it["issues"]}
                   for it in report["items"]
                   if it["issues"] and not it["auto_fixable"]]
        # ① 两表备份（带时间戳目录, 与业务目录分离存放）
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        backup_dir = os.path.join(self.data_root,
                                  f"store_model_migration_{ts}")
        table_paths = (("users.json", self.user_manager.users_path),
                       ("stores.json", self.user_manager.stores_path))
        try:
            os.makedirs(backup_dir)
            for name, path in table_paths:
                if os.path.exists(path):
                    shutil.copy2(path, os.path.join(backup_dir, name))
        except OSError as e:
            shutil.rmtree(backup_dir, ignore_errors=True)
            raise ApiError(f"已中止且未做任何变更: 迁移前备份失败: {e}",
                           code=500, http_status=500) from e

        def _restore_tables():
            for name, path in table_paths:
                src = os.path.join(backup_dir, name)
                if os.path.exists(src):
                    shutil.copy2(src, path)

        # ② 迁移前核对基线（十项清点 + 归属快照; 无目录店铺豁免清点）
        stores_before = {s["store_id"]: dict(s)
                         for s in self.user_manager.list_stores()
                         if s.get("store_id")}
        ownership_before = {sid: s.get("owner_user_id")
                            for sid, s in stores_before.items()}
        verifiable = [sid for sid in stores_before
                      if os.path.isdir(os.path.join(self.stores_root, sid))]
        inventory_before = {sid: self._store_inventory(sid)
                            for sid in verifiable}
        # ③ 字段补齐 + 孤儿归属置空（内存态修改, 一次原子写回）
        now = now_iso()
        migrated = []
        for it in fixable:
            store = stores_before[it["store_id"]]
            orphan_owner = "orphan_owner" in it["issues"]
            filled = {}
            if not store.get("status"):
                filled["status"] = "active"
            if not store.get("updated_at"):
                filled["updated_at"] = now
            if orphan_owner:
                filled["owner_user_id"] = None
                filled["previous_owner_user_id"] = store.get("owner_user_id")
                filled["assignment_status"] = "unassigned"
            elif not store.get("assignment_status"):
                filled["assignment_status"] = (
                    "unassigned" if not store.get("owner_user_id")
                    else "assigned")
            store.update(filled)
            migrated.append({
                "store_id": it["store_id"],
                "filled": filled,
                "owner_cleared": orphan_owner,
                "previous_owner_user_id":
                    filled.get("previous_owner_user_id"),
            })
        if migrated:
            try:
                with open(self.user_manager.stores_path,
                          encoding="utf-8") as f:
                    doc = json.load(f)
                doc["stores"] = list(stores_before.values())
                doc["updated_at"] = now
                self.user_manager._atomic_write(
                    self.user_manager.stores_path, doc)
            except (OSError, ValueError, json.JSONDecodeError) as e:
                _restore_tables()
                shutil.rmtree(backup_dir, ignore_errors=True)
                raise ApiError(f"已中止且未做任何变更: 字段写入失败: {e}",
                               code=500, http_status=500) from e
        # ④ 全店铺十项清点 + 归属关系前后核对（孤儿置空项豁免且显式列出）
        stores_after = {s["store_id"]: s
                        for s in self.user_manager.list_stores()
                        if s.get("store_id")}
        ownership_after = {sid: s.get("owner_user_id")
                           for sid, s in stores_after.items()}
        inventory_after = {sid: self._store_inventory(sid)
                           for sid in verifiable}
        orphan_cleared = {m["store_id"] for m in migrated
                          if m["owner_cleared"]}
        ownership_diff = {
            sid: [ownership_before[sid], ownership_after.get(sid)]
            for sid in ownership_before
            if ownership_before[sid] != ownership_after.get(sid)
            and sid not in orphan_cleared}
        inventory_diff = {
            sid: [inventory_before[sid], inventory_after.get(sid)]
            for sid in inventory_before
            if inventory_before[sid] != inventory_after.get(sid)}
        if ownership_diff or inventory_diff:
            _restore_tables()
            diff_summary = sorted(set(ownership_diff) | set(inventory_diff))
            self.security_log.append(
                "store.model_migrate",
                username=operator.get("username"),
                user_id=operator["user_id"], ip=self._client_ip(handler),
                result="error",
                detail={"backup_dir": backup_dir,
                        "recovered": True,
                        "ownership_diff": ownership_diff,
                        "inventory_diff": inventory_diff})
            raise ApiError(
                "已回滚，差异项见 detail: "
                + ", ".join(diff_summary or ownership_diff),
                code=500, http_status=500)
        # ⑤ 审计 + 响应（响应即迁移报告）
        self.security_log.append(
            "store.model_migrate", username=operator.get("username"),
            user_id=operator["user_id"], ip=self._client_ip(handler),
            detail={"migrated": [m["store_id"] for m in migrated],
                    "owner_cleared": sorted(orphan_cleared),
                    "skipped": [s["store_id"] for s in skipped],
                    "backup_dir": backup_dir,
                    "verification_match": True})
        handler._send(ApiResponse.ok({
            "migrated": migrated,
            "skipped": skipped,
            "backup_dir": backup_dir,
            "verification": {
                "before": inventory_before,
                "after": inventory_after,
                "match": True,
                "ownership_cleared": sorted(orphan_cleared),
                "excluded_no_dir": sorted(
                    set(stores_before) - set(verifiable)),
            },
        }, message="迁移完成"))

    def _handle_batch_revoke_preview(self, handler, parsed, batch_id):
        """撤销批次影响预览（仅管理员）: 各业务文件记录数/受影响订单/月份"""
        self._require_admin(handler)
        scope = self._query_scope(handler, parsed)
        b, sid, sname = self._find_batch_store(scope, batch_id)
        data = b.importer.revoke_preview(batch_id)
        data["store_id"] = sid
        data["store_name"] = sname
        handler._send(ApiResponse.ok(data))

    def _handle_batch_delete_history(self, handler, parsed, body, batch_id):
        """仅删除导入历史记录（不动业务数据; 仅管理员, 需二次确认）"""
        self._require_admin(handler)
        if body.get("confirm") is not True:
            raise ApiError("危险操作：请传 confirm=true 二次确认",
                           code=400, http_status=400)
        scope = self._query_scope(handler, parsed)
        b, sid, sname = self._find_batch_store(scope, batch_id)
        result = b.importer.delete_batch_history(
            batch_id, operator=self._operator(handler))
        result["store_id"] = sid
        result["store_name"] = sname
        handler._send(ApiResponse.ok(result, message="导入历史已删除"))

    def _handle_batch_revoke(self, handler, parsed, body, batch_id):
        """撤销导入批次（仅管理员, 需二次确认）

        流程: 整店快照 → 删除该批次业务记录与关联结果 → 按月重算对账/利润
        → 驱逐缓存。任一步失败自动从快照回滚, 数据保持原状。
        """
        self._require_admin(handler)
        if body.get("confirm") is not True:
            raise ApiError("危险操作：请先预览影响并传 confirm=true 二次确认",
                           code=400, http_status=400)
        scope = self._query_scope(handler, parsed)
        b, sid, sname = self._find_batch_store(scope, batch_id)
        operator = self._operator(handler)
        snapshot_dir, file_count = self._snapshot_store(sid)
        store_record = self.user_manager.get_store(sid) or {"store_id": sid}
        try:
            result = b.importer.revoke_batch(batch_id, operator=operator)
            recalc = []
            for month in result.get("recalc_months") or []:
                start, end = resolve_range(month=month)
                recon = b.reconciler.run_reconciliation(
                    start=start, end=end, operator=operator)
                profit = b.profits.run_calculation(
                    start=start, end=end, operator=operator)
                recalc.append({
                    "month": month,
                    "reconciliation_id": recon.get("reconciliation_id"),
                    "calculation_id": profit.get("calculation_id"),
                })
        except Exception:
            self._restore_store_from_snapshot(sid, snapshot_dir)
            raise
        result["recalculated"] = recalc
        result["store_id"] = sid
        result["store_name"] = sname
        result["backup"] = self._backup_info(
            [(store_record, snapshot_dir, file_count)])
        self.registry.evict(sid)
        b.op_log.append("store.batch_revoke", detail={
            "batch_id": batch_id, "removed": result.get("removed"),
            "recalc_months": result.get("recalc_months"),
            "backup": result["backup"][0] if result["backup"] else None,
        }, operator=operator)
        handler._send(ApiResponse.ok(result, message="批次已撤销并重算完成"))

    def _handle_backup(self, handler, body, bundle):
        name = body.get("name")
        if not name or name not in BUSINESS_FILES:
            raise ApiError("请提供有效的 name（业务数据名称）", code=400, http_status=400)
        filename = bundle.repo.backup(name)
        bundle.op_log.append(
            "data.backup",
            detail={"name": name, "backup_file": filename},
            operator=self._operator(handler),
        )
        handler._send(ApiResponse.ok({"name": name, "backup_file": filename}, message="备份成功"))

    def _handle_restore(self, handler, body, bundle):
        name = body.get("name")
        backup_file = body.get("backup_file")
        if not name or name not in BUSINESS_FILES:
            raise ApiError("请提供有效的 name（业务数据名称）", code=400, http_status=400)
        if not backup_file:
            raise ApiError("请提供 backup_file（备份文件名）", code=400, http_status=400)
        data = bundle.repo.restore(name, backup_file)
        bundle.op_log.append(
            "data.restore",
            detail={"name": name, "backup_file": backup_file},
            operator=self._operator(handler),
        )
        handler._send(ApiResponse.ok(data, message="恢复成功"))

    def _handle_logs(self, handler, parsed, bundle):
        qs = parse_qs(parsed.query)
        action = qs.get("action", [None])[0]
        operator = qs.get("operator", [None])[0]
        keyword = qs.get("keyword", [None])[0] or None
        start, end = self._time_filter(qs)
        scope = self._query_scope(handler, parsed)

        def query_all(b, sid, sname, limit=None):
            entries = b.op_log.query(
                limit=limit, action=action, operator=operator,
                start=start, end=end, keyword=keyword)
            out = []
            for e in entries:
                if isinstance(e, dict):
                    # operator 归一化已下沉到读取层（op_log.query 逐行
                    # 归一化，先于筛选），此处仅补充店铺归属并复制条目，
                    # 防止上层修改 query 返回对象。
                    e = dict(e)
                    e.setdefault("store_id", sid)
                    e.setdefault("store_name", sname)
                out.append(e)
            return out

        if any(k in qs for k in ("page", "page_size", "size")):
            page, size = self._page_params(qs)
            if not scope.multi:
                entries = query_all(scope.bundle, scope.store_id,
                                    scope.store_name)
            else:
                entries = []
                for b, sid, sname in scope.targets:
                    entries.extend(query_all(b, sid, sname))
                entries.sort(key=lambda e: (e.get("time") or ""), reverse=True)
            total = len(entries)
            operators = {e.get("operator") for e in entries if e.get("operator")}
            errors = sum(1 for e in entries if e.get("result") == "error")
            handler._send(ApiResponse.ok({
                "total": total, "page": page, "size": size,
                "pages": (total + size - 1) // size,
                "entries": entries[(page - 1) * size:page * size],
                "statistics": {
                    "entry_count": total,
                    "operator_count": len(operators),
                    "error_count": errors,
                    "time_field": "time",
                },
                "time_range": self._time_meta(qs),
                **({"store_scope": "all", "store_count": len(scope.targets)}
                   if scope.multi else {}),
            }))
            return
        # 兼容旧 limit 模式（无分页参数）
        try:
            limit = int(qs.get("limit", ["100"])[0] or 100)
        except ValueError:
            raise ApiError("limit 必须是整数", code=400, http_status=400)
        limit = max(0, min(limit, config.LOG_QUERY_MAX))
        if not scope.multi:
            entries = query_all(scope.bundle, scope.store_id, scope.store_name,
                                limit=limit)
        else:
            entries = []
            for b, sid, sname in scope.targets:
                entries.extend(query_all(b, sid, sname, limit=limit))
            entries.sort(key=lambda e: (e.get("time") or ""), reverse=True)
            entries = entries[:limit]
        handler._send(ApiResponse.ok({"total": len(entries), "entries": entries}))

    def _handle_list_backups(self, handler, parsed, bundle):
        qs = parse_qs(parsed.query)
        name = qs.get("name", [None])[0]
        if not name or name not in BUSINESS_FILES:
            raise ApiError("请提供有效的 name（业务数据名称）", code=400, http_status=400)
        # 统一经查询范围解析: member 携带越权参数（member/store_id）与其他业务 API 同口径 403
        scope = self._query_scope(handler, parsed)
        if not scope.targets:
            raise ApiError("暂无成员店铺数据", code=404, http_status=404)
        bundle = scope.bundle
        backups = bundle.repo.list_backups(name)
        handler._send(ApiResponse.ok({"name": name, "total": len(backups), "backups": backups}))

"""全站统一应用外壳（侧边栏 + 顶栏）的唯一来源。

各业务页面不再复制侧边栏/顶栏 HTML，只保留一行 ``<!--@app-shell-->`` 占位符，
由 ``app.server`` 在输出 HTML 前按当前路由注入本模块生成的片段。
这样导航顺序、当前项高亮、管理员菜单标记块只有一处维护。

管理员标记块（admin-nav / admin-only-nav）保留在生成结果中，
服务端仍按既有逻辑在 member 会话下剔除，安全口径不变。
"""

SHELL_TOKEN = b"<!--@app-shell-->"

# 页面文件 -> 侧边栏高亮路由（与 app.server.PAGE_ROUTES 的语义路由一致）
PAGE_FILE_ACTIVE_ROUTE = {
    "dashboard.html": "/dashboard",
    "index.html": "/data-import",
    "orders.html": "/orders",
    "refund_orders.html": "/refund-orders",
    "supplier_costs.html": "/supplier-costs",
    "reconciliation.html": "/reconciliation",
    "profit.html": "/profit",
    "monthly_profit.html": "/monthly-profit",
    "anomalies.html": "/anomalies",
    "logs.html": "/logs",
    "users.html": "/users",
    "create_user.html": "/users/create",
    "stores.html": "/stores",
}

# 业务菜单：语义路由、文案（顺序即全站固定顺序）
NAV_ITEMS = [
    ("/dashboard", "经营总览"),
    ("/data-import", "数据导入"),
    ("/monthly-profit", "每月盈利"),
    ("/reconciliation", "对账成本"),
    ("/profit", "利润核算"),
    ("/orders", "普通订单"),
    ("/refund-orders", "退款订单"),
    ("/supplier-costs", "供应商成本"),
]

# 管理菜单组（仅 admin 输出，服务端按标记块剔除）
MGMT_ITEMS = [
    ("/users/create", "创建用户"),
    ("/stores", "店铺管理"),
    ("/users", "用户管理"),
]

# 管理员专属菜单（仅 admin 输出）
ADMIN_ONLY_ITEMS = [
    ("/anomalies", "异常处理"),
    ("/logs", "操作日志"),
]

_ROUTE_LABEL = {href: label for href, label in (
    NAV_ITEMS + MGMT_ITEMS + ADMIN_ONLY_ITEMS)}
_MGMT_PATHS = {href for href, _ in MGMT_ITEMS}


def _link(href, label, active):
    cls = ' class="active"' if active else ""
    return f'    <a href="{href}"{cls}>{label}</a>'


def _breadcrumb(active_path):
    """顶栏面包屑：首页 > (用户管理) > 当前页"""
    path = (active_path or "").rstrip("/") or "/"
    if path == "/":
        path = "/dashboard"
    parts = ['<nav class="app-breadcrumb" aria-label="面包屑">']
    parts.append('<a href="/dashboard" class="bc-link">首页</a>')
    label = _ROUTE_LABEL.get(path)
    if path in _MGMT_PATHS:
        parts.append('<span class="bc-sep" aria-hidden="true">/</span>')
        parts.append('<span class="bc-muted">用户管理</span>')
    if path != "/dashboard" and label:
        parts.append('<span class="bc-sep" aria-hidden="true">/</span>')
        parts.append(f'<span class="bc-current" aria-current="page">{label}</span>')
    parts.append("</nav>")
    return "".join(parts)


def render(active_path):
    """生成侧边栏 + 顶栏 + 内容容器起始标签。"""
    active_path = (active_path or "").rstrip("/") or "/"
    if active_path == "/":
        active_path = "/dashboard"
    group_active = active_path in _MGMT_PATHS
    lines = [
        '<aside class="app-sidebar" id="app-sidebar">',
        '  <div class="sidebar-brand"><span class="logo">微</span><span class="brand-text">微信小店对账与利润核算系统</span></div>',
        '  <nav>',
        '',
    ]
    lines += [_link(href, label, active_path == href)
              for href, label in NAV_ITEMS]
    lines += [
        '    <!--admin-nav-start-->',
        f'    <button type="button" class="nav-group{" nav-group-current" if group_active else ""}" data-nav-group="user-mgmt"',
        '            aria-expanded="true" aria-controls="nav-user-mgmt">用户管理</button>',
        '    <div class="nav-submenu" id="nav-user-mgmt">',
    ]
    lines += [_link(href, label, active_path == href)
              for href, label in MGMT_ITEMS]
    lines += [
        '    </div>',
        '    <!--admin-nav-end-->',
        '    <!--admin-only-nav-start-->',
    ]
    lines += [_link(href, label, active_path == href)
              for href, label in ADMIN_ONLY_ITEMS]
    lines += [
        '    <!--admin-only-nav-end-->',
        '  </nav>',
        '  <button class="sidebar-collapse-btn" id="sidebar-collapse" type="button" aria-label="收起或展开侧边菜单"><span class="toggle-arrow">«</span><span class="toggle-text">收起菜单</span></button>',
        '</aside>',
        '<div class="sidebar-backdrop" id="sidebar-backdrop"></div>',
        '<div class="app-main">',
        '<header class="topbar">',
        '  <button class="topbar-menu-btn" id="sidebar-open" type="button" aria-label="打开菜单">☰</button>',
        '  <span class="topbar-brand"><span class="logo">微</span>微信小店对账与利润核算系统</span>',
        f'  {_breadcrumb(active_path)}',
        '  <span class="meta" id="nav-meta"></span>',
        '  <div id="auth-area" class="auth-area"></div>',
        '</header>',
        '<div class="container app-content">',
        '<script src="/static/app_chrome.js"></script>',
    ]
    return "\n".join(lines) + "\n"

/* 侧边导航交互: 桌面折叠/展开（localStorage 记忆）+ 移动端抽屉 + 管理用户菜单组
   布局与图标由 style.css 静态渲染, 本脚本仅负责交互:
   - #sidebar-collapse: 折叠/展开（body.sidebar-collapsed, 220px <-> 64px）
   - #sidebar-open + #sidebar-backdrop: 移动端抽屉开合（body.sidebar-open）
   - .nav-group: 管理用户菜单组展开/收起（localStorage 记忆, 组内路由强制展开）
   - 折叠态悬停提示: 动态为菜单项设置 title（不写入静态 HTML, 保持导航断言口径） */
(function () {
  var KEY = "wxr_sidebar_collapsed";
  var NAV_GROUP_KEY = "wxr_nav_user_mgmt";
  var GROUP_ROUTES = ["/users/create", "/stores", "/users"];
  var body = document.body;

  function isMobile() {
    return window.matchMedia("(max-width: 768px)").matches;
  }

  function init() {
    try {
      if (localStorage.getItem(KEY) === "1") {
        body.classList.add("sidebar-collapsed");
      }
    } catch (e) { /* 隐私模式等场景忽略 */ }

    var links = document.querySelectorAll(".app-sidebar nav a");
    for (var i = 0; i < links.length; i++) {
      links[i].title = (links[i].textContent || "").trim();
    }

    var collapseBtn = document.getElementById("sidebar-collapse");
    if (collapseBtn) {
      collapseBtn.addEventListener("click", function () {
        var collapsed = body.classList.toggle("sidebar-collapsed");
        try {
          localStorage.setItem(KEY, collapsed ? "1" : "0");
        } catch (e) { /* 忽略存储失败 */ }
        /* 箭头方向与文案由 CSS 按 body.sidebar-collapsed 状态统一控制
           （rotate 180°; 折叠态文案本身 display:none）, 此处不改 DOM 避免双重翻转 */
      });
    }

    var openBtn = document.getElementById("sidebar-open");
    if (openBtn) {
      openBtn.addEventListener("click", function () {
        body.classList.add("sidebar-open");
      });
    }
    var backdrop = document.getElementById("sidebar-backdrop");
    if (backdrop) {
      backdrop.addEventListener("click", closeDrawer);
    }
    for (var j = 0; j < links.length; j++) {
      (function (a) {
        a.addEventListener("click", function () {
          if (isMobile()) closeDrawer();
        });
      })(links[j]);
    }
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape") closeDrawer();
    });
    window.addEventListener("resize", function () {
      if (!isMobile()) closeDrawer();
    });

    initNavGroup();
  }

  function closeDrawer() {
    body.classList.remove("sidebar-open");
  }

  /* "管理用户"菜单组: 展开状态 localStorage 记忆; 组内路由强制展开;
     桌面折叠态点击组图标先展开侧边栏（折叠态子菜单被 CSS 隐藏, 需先恢复侧栏宽度）;
     member 视角（服务端已剔除标记块）querySelector 为空时安全返回 */
  function initNavGroup() {
    var btn = document.querySelector('.app-sidebar nav .nav-group[data-nav-group="user-mgmt"]');
    var submenu = document.getElementById("nav-user-mgmt");
    if (!btn || !submenu) return;
    btn.title = (btn.textContent || "").trim();
    var path = location.pathname.replace(/\/+$/, "") || "/";
    var collapsed = false;
    try {
      collapsed = localStorage.getItem(NAV_GROUP_KEY) === "1";
    } catch (e) { /* 隐私模式等场景忽略 */ }
    if (GROUP_ROUTES.indexOf(path) !== -1) collapsed = false;
    apply(collapsed);
    btn.addEventListener("click", function () {
      /* 折叠态（桌面）下子菜单不可见: 点击组图标 = 展开侧边栏 + 展开子菜单 */
      if (!isMobile() && body.classList.contains("sidebar-collapsed")) {
        body.classList.remove("sidebar-collapsed");
        try { localStorage.setItem(KEY, "0"); } catch (e) { /* 忽略存储失败 */ }
        apply(false);
        return;
      }
      apply(!submenu.classList.contains("collapsed"));
    });
    function apply(c) {
      submenu.classList.toggle("collapsed", c);
      btn.setAttribute("aria-expanded", c ? "false" : "true");
      try { localStorage.setItem(NAV_GROUP_KEY, c ? "1" : "0"); } catch (e) { /* 忽略存储失败 */ }
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();

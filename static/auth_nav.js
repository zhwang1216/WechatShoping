/* 账号区公共脚本: 当前用户信息、修改密码、退出登录
   渲染目标为各业务页面 </nav> 之后的 #auth-area 容器（nav 之外, 不影响业务菜单） */
(function () {
  function render() {
    const area = document.getElementById("auth-area");
    if (!area) return;
    API.get("/api/auth/me").then(me => {
      const isAdmin = me.role === "admin";
      /* 管理员无业务店铺（store_id=null）→ 全店视角文案；
         成员单店显示本店名；成员多店（store_count>1）显示店铺数量 */
      let whoName;
      if (isAdmin) {
        whoName = me.store_name ? `（${esc(me.store_name)}）` : "（全店视角）";
      } else if (me.store_count > 1) {
        whoName = `（${me.store_count} 家店铺）`;
      } else {
        whoName = `（${esc(me.store_name || "本店")}）`;
      }
      area.innerHTML = `
        <span class="who">
          <span>${esc(me.username)}${whoName}</span>
          <span class="role-badge ${isAdmin ? "admin" : "member"}">${isAdmin ? "管理员" : "成员"}</span>
        </span>
        ${isAdmin ? '<a class="auth-link" href="/users">用户管理</a>' : ""}
        <button class="auth-link" id="auth-change-pwd" type="button">修改密码</button>
        <button class="btn-logout" id="auth-logout" type="button">退出登录</button>`;
      document.getElementById("auth-logout").onclick = async () => {
        try {
          await API.post("/api/auth/logout", {});
        } catch (e) { /* 会话已失效也可继续回登录页 */ }
        location.href = "/login";
      };
      document.getElementById("auth-change-pwd").onclick = openChangePassword;
    }).catch(() => {
      /* 会话失效: api.js 已统一引导跳转登录页, 此处不重复处理 */
    });
  }

  function openChangePassword() {
    /* 修改密码弹窗独立于 modal.js 的 #um-modal：先清理旧弹窗，避免重复点击叠加。
       未引入 modal.js 的页面（多数列表页）不应因 closeModal 未定义而中断。 */
    if (typeof closeModal === "function") closeModal();
    const oldPwd = document.getElementById("pwd-modal");
    if (oldPwd) oldPwd.remove();
    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";
    overlay.id = "pwd-modal";
    overlay.innerHTML = `
      <div class="modal-card">
        <h3>修改密码</h3>
      <div class="modal-body">
        <div><label>原密码</label><input id="pwd-old" type="password" autocomplete="current-password"></div>
        <div><label>新密码</label><input id="pwd-new" type="password" autocomplete="new-password" placeholder="至少 8 位，且同时包含字母和数字"></div>
        <div><label>确认新密码</label><input id="pwd-confirm" type="password" autocomplete="new-password" placeholder="再次输入新密码"></div>
        <div class="pwd-match-hint" id="pwd-match" aria-live="polite"></div>
        <div class="modal-msg" id="pwd-msg"></div>
        <div class="modal-actions">
          <button class="btn secondary" id="pwd-cancel" type="button">取消</button>
          <button class="btn" id="pwd-submit" type="button">确认修改</button>
        </div>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.addEventListener("click", e => { if (e.target === overlay) closeModal(); });
  document.getElementById("pwd-cancel").onclick = closeModal;
  setupPasswordToggle(document.getElementById("pwd-old"), "显示或隐藏原密码");
  setupPasswordToggle(document.getElementById("pwd-new"), "显示或隐藏新密码");
  setupPasswordToggle(document.getElementById("pwd-confirm"), "显示或隐藏确认新密码");
  bindPasswordMatch({
    newInput: document.getElementById("pwd-new"),
    confirmInput: document.getElementById("pwd-confirm"),
    hintEl: document.getElementById("pwd-match"),
    submitBtn: document.getElementById("pwd-submit"),
  });
  document.getElementById("pwd-submit").onclick = async () => {
    const msg = document.getElementById("pwd-msg");
    const oldPwd = document.getElementById("pwd-old").value;
    const newPwd = document.getElementById("pwd-new").value;
    const confirmPwd = document.getElementById("pwd-confirm").value;
    msg.style.display = "none";
    if (!oldPwd || !newPwd) {
      showPwdMsg("请填写原密码和新密码");
      return;
    }
    if (newPwd !== confirmPwd) {
      showPwdMsg("两次输入的新密码不一致");
      return;
    }
    const btn = document.getElementById("pwd-submit");
    btn.disabled = true;
    btn.textContent = "提交中…";
    try {
      await API.post("/api/auth/change-password",
                     { old_password: oldPwd, new_password: newPwd,
                       confirm_password: confirmPwd });
        sessionStorage.setItem("login_notice", "请使用新密码重新登录");
        location.href = "/login";
      } catch (e) {
        showPwdMsg(e.message || "修改失败，请稍后重试");
        btn.disabled = false;
        btn.textContent = "确认修改";
      }
    };
    function showPwdMsg(text) {
      const msg = document.getElementById("pwd-msg");
      msg.textContent = text;
      msg.style.display = "block";
    }
  }

  function closeModal() {
    const m = document.getElementById("pwd-modal");
    if (m) m.remove();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", render);
  } else {
    render();
  }
})();

/* 公共工具：API 请求封装、格式化、状态徽章、toast */
const API = {
  async request(path, options = {}) {
    // 统一注入 CSRF 请求头（写操作必需; GET 携带无害）; 调用方显式同名头优先
    const opts = Object.assign({}, options, {
      headers: Object.assign({ "X-Requested-With": "XMLHttpRequest" },
                             options.headers || {}),
    });
    const resp = await fetch(path, opts);
    // 公开认证页（/login、/setup）豁免会话失效跳转, 避免登录失败触发跳转循环
    const onPublicAuthPage = ["/login", "/register", "/setup"]
      .some(p => location.pathname.startsWith(p));
    const sessionExpired = () => {
      if (onPublicAuthPage) return;
      location.href = "/login";
    };
    let body;
    try {
      body = await resp.json();
    } catch (e) {
      if (resp.status === 401) sessionExpired();
      throw { code: resp.status, message: `接口响应异常（HTTP ${resp.status}）` };
    }
    if (resp.status === 401 || body.code === 1001) {
      sessionExpired();
    }
    if (body.code !== 0) {
      throw { code: body.code, message: body.message || "请求失败", data: body.data };
    }
    return body.data;
  },
  get(path) {
    return this.request(path);
  },
  post(path, body) {
    return this.request(path, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body || {}),
    });
  },
  async upload(path, file) {
    return this.request(path, {
      method: "POST",
      headers: { "X-Filename": encodeURIComponent(file.name) },
      body: file,
    });
  },
};

function fmtMoney(v) {
  if (v === null || v === undefined || v === "") return "—";
  const n = Number(v);
  if (Number.isNaN(n)) return String(v);
  return n.toFixed(2);
}
function fmtMoneyYuan(v) {
  if (v === null || v === undefined || v === "") return "—";
  const n = Number(v);
  if (Number.isNaN(n)) return String(v);
  return `¥${n.toFixed(2)}`;
}
function fmtTime(v) {
  if (!v) return "—";
  return String(v).replace("T", " ");
}
function esc(s) {
  if (s === null || s === undefined) return "";
  return String(s).replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[c]);
}

const ORDER_STATUS_META = {
  completed: ["已完成", "green"], shipped: ["已发货", "blue"],
  pending_shipment: ["待发货", "cyan"], paid: ["已支付", "blue"],
  pending_payment: ["待付款", "gray"], closed: ["已关闭", "gray"],
  refund_processing: ["退款处理中", "orange"], partial_refunded: ["部分退款", "orange"],
  full_refunded: ["全额退款", "red"], unknown: ["未知状态", "gray"],
};
const REFUND_STATUS_META = {
  none: ["无退款", "gray"], pending: ["退款处理中", "orange"],
  partial: ["部分退款", "blue"], full: ["全额退款", "red"],
  closed: ["退款关闭", "gray"], unknown: ["未知/待核对", "orange"],
};
function orderStatusBadge(s) {
  const [label, color] = ORDER_STATUS_META[s] || [s || "—", "gray"];
  return `<span class="badge ${color}">${esc(label)}</span>`;
}
function refundStatusBadge(s) {
  const [label, color] = REFUND_STATUS_META[s] || [s || "—", "gray"];
  return `<span class="badge ${color}">${esc(label)}</span>`;
}
function statusSourceBadge(s) {
  if (s === "derived") return `<span class="badge orange">推导</span>`;
  if (s === "platform") return `<span class="badge gray">平台原始</span>`;
  return "—";
}

function toast(message, type = "info", ms = 3200) {
  let root = document.getElementById("toast-root");
  if (!root) {
    root = document.createElement("div");
    root.id = "toast-root";
    document.body.appendChild(root);
  }
  const el = document.createElement("div");
  el.className = `toast ${type}`;
  el.textContent = message;
  root.appendChild(el);
  setTimeout(() => el.remove(), ms);
}

/* 简易分页控件 */
function renderPager(el, data, { onPage, onSize }) {
  if (!data || data.total === 0) { el.innerHTML = ""; return; }
  const { page, pages, total, size } = data;
  const btns = [];
  btns.push(`<button ${page <= 1 ? "disabled" : ""} data-p="1">首页</button>`);
  btns.push(`<button ${page <= 1 ? "disabled" : ""} data-p="${page - 1}">上一页</button>`);
  for (let p = 1; p <= pages; p++) {
    if (pages > 7 && p > 2 && p < pages - 1 && Math.abs(p - page) > 1) continue;
    btns.push(`<button class="${p === page ? "cur" : ""}" data-p="${p}">${p}</button>`);
  }
  btns.push(`<button ${page >= pages ? "disabled" : ""} data-p="${page + 1}">下一页</button>`);
  btns.push(`<button ${page >= pages ? "disabled" : ""} data-p="${pages}">末页</button>`);
  el.innerHTML = `
    <span class="info">共 ${total} 条 / ${pages} 页</span>
    <span class="controls">
      <select id="page-size">
        ${[20, 50, 100].map(s => `<option value="${s}" ${s === size ? "selected" : ""}>${s} 条/页</option>`).join("")}
      </select>
      ${btns.join("")}
    </span>`;
  el.querySelectorAll("button[data-p]").forEach(b => {
    b.onclick = () => onPage(Number(b.dataset.p));
  });
  el.querySelector("#page-size").onchange = e => onSize(Number(e.target.value));
}

/* ==================== 密码输入公共组件 ====================
   登录/注册/初始化/修改密码/重置密码各表单复用, 保证行为一致:
   - setupPasswordToggle(input): 输入框右侧注入小眼睛按钮, password/text 切换;
     仅改变显示方式, 不触碰输入内容/焦点/光标, 不持久化密码到任何存储;
   - bindPasswordMatch({...}): 新密码与确认密码实时一致性校验;
     两者均已填写且不一致时显示提示并禁用提交按钮, 任一为空不显示误导性提示。
   注意: 前端校验仅用于交互反馈, 后端仍独立校验一致性。 */
const PWD_EYE_SVG = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>';
const PWD_EYE_OFF_SVG = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></svg>';

function setupPasswordToggle(input, ariaLabel) {
  if (!input || input.dataset.pwdToggle === "on") return input;
  input.dataset.pwdToggle = "on";
  input.type = "password"; // 默认隐藏
  const wrap = document.createElement("span");
  wrap.className = "pwd-field";
  input.parentNode.insertBefore(wrap, input);
  wrap.appendChild(input);
  input.style.paddingRight = "40px"; // 为右侧眼睛按钮留位（各页面输入框均为 border-box）
  const btn = document.createElement("button");
  btn.type = "button"; // 独立按钮, 不触发任何表单提交
  btn.className = "pwd-toggle";
  btn.setAttribute("aria-label", ariaLabel || "显示或隐藏密码");
  btn.setAttribute("aria-pressed", "false");
  btn.title = "显示或隐藏密码";
  btn.innerHTML = PWD_EYE_SVG;
  btn.addEventListener("click", () => {
    const wasFocused = document.activeElement === input;
    const start = input.selectionStart, end = input.selectionEnd;
    const show = input.type === "password";
    input.type = show ? "text" : "password";
    btn.innerHTML = show ? PWD_EYE_OFF_SVG : PWD_EYE_SVG;
    btn.setAttribute("aria-pressed", show ? "true" : "false");
    if (wasFocused) { // 切换后交还焦点并恢复光标位置, 避免打断输入
      input.focus();
      try { input.setSelectionRange(start, end); } catch (e) { /* type 不支持选择时忽略 */ }
    }
  });
  wrap.appendChild(btn);
  return input;
}

function bindPasswordMatch({ newInput, confirmInput, hintEl, submitBtn }) {
  const sync = () => {
    const nv = newInput ? newInput.value : "";
    const cv = confirmInput ? confirmInput.value : "";
    let state = ""; // ""=不提示 / "ok"=一致 / "err"=不一致
    if (nv !== "" && cv !== "") {
      state = nv === cv ? "ok" : "err";
    }
    if (hintEl) {
      hintEl.className = "pwd-match-hint" + (state ? " " + state : "");
      hintEl.textContent = state === "ok" ? "两次输入一致"
        : state === "err" ? "确认密码与密码不一致" : "";
    }
    if (submitBtn) {
      submitBtn.disabled = state === "err"; // 不一致时禁止提交
    }
    return state !== "err";
  };
  if (newInput) newInput.addEventListener("input", sync);
  if (confirmInput) confirmInput.addEventListener("input", sync);
  sync();
  return sync;
}

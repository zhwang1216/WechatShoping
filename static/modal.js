/* 公共弹窗组件（管理页表单 + 订单/成本/利润/对账等只读详情共用）
   依赖: api.js 的 esc()/toast()（modal.js 须在 api.js 之后引用）
   约定: overlay id 统一为 um-modal, 同屏互斥单弹窗;
         openModal 为新统一入口; confirmModal/inputModal 为向后兼容的薄封装;
         openDrawer 为右侧抽屉形态（详情查看场景） */

const _escCloseHooks = [];
function onEscClose(fn) { if (typeof fn === "function") _escCloseHooks.push(fn); }

/* 统一弹窗封装: { title, bodyHtml, width, actions, onClose }
   actions: [{ text, kind: "primary"|"secondary"|"danger"|"main", onClick, disabled }]
   - onClick 可返回 Promise; 请求进行中按钮自动禁用并显示"处理中…";
   - 失败（抛错）时弹窗不关闭, 错误写入 .modal-msg 区域;
   - 点击 × / 遮罩 / Esc / kind=secondary 均只关闭不提交。 */
function openModal({ title, bodyHtml, width, actions, onClose }) {
  closeModal();
  const overlay = document.createElement("div");
  overlay.className = "modal-overlay";
  overlay.id = "um-modal";
  const cardW = width || 420;
  const actBtns = (actions || []).map((a, i) => {
    const kind = a.kind === "danger" ? " danger"
      : a.kind === "secondary" ? " secondary" : "";
    return `<button class="btn${kind}" id="um-act-${i}" type="button"
      ${a.disabled ? "disabled" : ""}>${esc(a.text || "确认")}</button>`;
  }).join("");
  overlay.innerHTML = `
    <div class="modal-card um-card" style="max-width:${cardW}px">
      <div class="um-head">
        <h3>${esc(title)}</h3>
        <button class="modal-close" id="um-x" type="button" aria-label="关闭">×</button>
      </div>
      <div class="modal-body um-body">
        ${bodyHtml || ""}
        <div class="modal-msg" id="um-msg"></div>
        <div class="modal-actions">
          ${actBtns}
        </div>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  const doClose = () => {
    closeModal();
    if (onClose) onClose();
  };
  overlay.addEventListener("click", e => { if (e.target === overlay) doClose(); });
  document.getElementById("um-x").onclick = doClose;
  (actions || []).forEach((a, i) => {
    const btn = document.getElementById(`um-act-${i}`);
    if (!btn) return;
    btn.onclick = async () => {
      btn.disabled = true;
      btn.textContent = "处理中…";
      const msg = document.getElementById("um-msg");
      if (msg) { msg.style.display = "none"; }
      try {
        const ret = await a.onClick(btn);
        if (ret !== false) doClose();
        else { btn.disabled = false; btn.textContent = a.text || "确认"; }
      } catch (e) {
        if (msg) {
          msg.textContent = (e && e.message) || "操作失败，请重试";
          msg.style.display = "block";
        }
        btn.disabled = false;
        btn.textContent = a.text || "确认";
      }
    };
  });
  if (document.getElementById("um-msg")) {
    document.getElementById("um-msg").style.display = "none";
  }
}

function confirmModal(title, text, onOk) {
  openModal({
    title,
    bodyHtml: `<div style="font-size:13px;line-height:1.7">${esc(text)}</div>`,
    actions: [
      { text: "取消", kind: "secondary", onClick: () => {} },
      { text: "确认", kind: "primary", onClick: onOk },
    ],
  });
}

/* 带输入校验的确认弹窗: 输入值与 expected（字符串或数组, 数组时任一匹配）一致时才可提交 */
function inputModal({ title, bodyHtml, expected, expectedLabel, confirmText, onOk }) {
  let inputEl;
  const html = `
    <div style="font-size:13px;line-height:1.9;border:1px solid var(--border);border-radius:8px;padding:10px 14px;margin-bottom:10px">${bodyHtml || ""}</div>
    <div><label>${esc(expectedLabel)}</label><input id="im-input" autocomplete="off"></div>`;
  const expectedList = Array.isArray(expected) ? expected : [expected];
  openModal({
    title,
    bodyHtml: html,
    actions: [
      { text: "取消", kind: "secondary", onClick: () => {} },
      { text: confirmText || "确认", kind: "danger", disabled: true,
        onClick: () => onOk(inputEl.value.trim()) },
    ],
  });
  inputEl = document.getElementById("im-input");
  const okBtn = document.getElementById("um-act-1");
  const sync = () => { okBtn.disabled = !expectedList.includes(inputEl.value.trim()); };
  inputEl.oninput = sync;
  inputEl.onkeydown = e => { if (e.key === "Enter" && !okBtn.disabled) okBtn.click(); };
  inputEl.focus();
}

function closeModal() {
  const m = document.getElementById("um-modal");
  if (m) m.remove();
}

/* 右侧抽屉: 复用 openModal 逻辑, 仅给 overlay 追加 .modal-drawer 类以切换为右侧滑出布局。
   适用: 订单详情/退款汇总/成本匹配/日志详情等复杂只读展示场景。
   API 与 openModal 完全一致: { title, bodyHtml, width, actions, onClose } */
function openDrawer({ title, bodyHtml, width, actions, onClose }) {
  const vw = typeof window !== "undefined" ? window.innerWidth : 1024;
  const drawerW = width || 480;
  const cap = vw <= 768 ? vw : Math.min(drawerW, Math.max(320, vw - 48));
  openModal({ title, bodyHtml, width: cap, actions, onClose });
  const overlay = document.getElementById("um-modal");
  if (overlay) overlay.classList.add("modal-drawer");
}

/* 全局 Esc: 优先关闭展开的"更多"下拉菜单（各页面经 onEscClose 注册）, 再关闭弹窗;
   两者关闭均不触发任何业务提交 */
document.addEventListener("keydown", e => {
  if (e.key !== "Escape") return;
  for (const fn of _escCloseHooks.slice()) fn();
  if (document.getElementById("um-modal")) closeModal();
});
/* 全站 ERP 外壳：统一页面标题区（不改业务逻辑与接口） */
(function () {
  function enhancePageHeader() {
    var root = document.querySelector(".container.app-content");
    if (!root || root.querySelector(".erp-page-header")) return;
    var existing = root.querySelector(":scope > .page-head");
    if (existing) {
      existing.classList.add("erp-page-header");
      return;
    }
    var title = root.querySelector(":scope > .page-title");
    if (!title) return;
    var desc = title.nextElementSibling;
    var hasDesc = desc && desc.classList.contains("page-desc");
    var wrap = document.createElement("div");
    wrap.className = "erp-page-header page-head";
    title.parentNode.insertBefore(wrap, title);
    wrap.appendChild(title);
    if (hasDesc) wrap.appendChild(desc);
    var loneBtn = wrap.nextElementSibling;
    if (loneBtn && loneBtn.classList && loneBtn.classList.contains("btn")
        && loneBtn.id && loneBtn.id.indexOf("btn-") === 0) {
      var slot = document.createElement("div");
      slot.className = "erp-page-actions";
      wrap.appendChild(slot);
      slot.appendChild(loneBtn);
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", enhancePageHeader);
  } else {
    enhancePageHeader();
  }
})();

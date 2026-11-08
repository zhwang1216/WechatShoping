/* 统一双维筛选组件（所有业务页面共用：成员 + 店铺级联）
 *
 * 行为：
 * - 管理员：渲染成员下拉（全部成员 + 各成员）与店铺下拉两级级联；
 *   选定成员后店铺下拉收窄为该成员店铺，清除成员恢复全量（spec 5.5.2 规则 2）；
 * - 普通成员 / 单店铺：只读徽章显示本店名称，无切换入口（后端锁定本店）；
 * - 选择"全部"时请求不携带对应参数（后端对管理员默认聚合全部成员店铺）；
 * - 选择具体成员/店铺时列表/导出/统计请求自动携带 member/store_id；
 * - 写操作目标（对账/重算等）经 writeTarget() 携带 target_store_id 与 member；
 * - 筛选状态写入 URL（?member= 与 ?store_id=），刷新后恢复；非法值自动回退"全部"。
 *
 * 用法（三方法签名与既有调用约定零变更，10 个已接入页面无需改动）：
 *   const SS = new StoreScope({
 *     container: "store-scope",
 *     onChange: () => { STATE.page = 1; load(); },
 *   });
 *   const p = new URLSearchParams();
 *   Object.entries(MF.getParams()).forEach(([k, v]) => p.set(k, v));
 *   SS.appendTo(p);          // 或 Object.entries(SS.getParams())...
 *   // 写操作（对账/重算等）目标店铺：
 *   const body = Object.assign({month: "2026-03"}, SS.writeTarget());
 */
class StoreScope {
  constructor(opts) {
    this.el = typeof opts.container === "string"
      ? document.getElementById(opts.container) : opts.container;
    if (!this.el) throw new Error("StoreScope: 容器不存在");
    this.onChange = opts.onChange || (() => {});
    this.allLabel = opts.allLabel || "全部店铺";
    this.stores = [];
    this.members = [];      /* admin: [{user_id, username, stores: [...]}] */
    this.multi = false;
    this.admin = false;
    this.storeId = null;
    this.memberId = null;
    this._initFromUrl();
    /* 店铺管理页物理删除店铺后派发 store-deleted: 清理当前选择并重载 */
    this._onStoreDeleted = e => this._handleStoreDeleted(e);
    window.addEventListener("store-deleted", this._onStoreDeleted);
    this._onScopeRefresh = () => this._load().then(() => this.onChange());
    window.addEventListener("store-scope-refresh", this._onScopeRefresh);
    this.ready = this._load();
  }

  _initFromUrl() {
    const q = new URLSearchParams(location.search);
    const sid = q.get("store_id");
    this.storeId = (sid && sid !== "all") ? sid : null;
    const mid = q.get("member");
    this.memberId = (mid && mid !== "all") ? mid : null;
  }

  _syncUrl() {
    const q = new URLSearchParams(location.search);
    q.delete("store_id");
    q.delete("member");
    if (this.admin) {
      if (this.memberId) q.set("member", this.memberId);
      if (this.multi && this.storeId) q.set("store_id", this.storeId);
    } else if (this.multi && this.storeId) {
      q.set("store_id", this.storeId);
    }
    const s = q.toString();
    history.replaceState(null, "", location.pathname + (s ? `?${s}` : "") + location.hash);
  }

  async _load() {
    try {
      const [d, me] = await Promise.all([
        API.get("/api/stores"),
        API.get("/api/auth/me").catch(() => null),
      ]);
      this.stores = d.stores || [];
      this.admin = !!(me && me.role === "admin");
      /* 成员 → 店铺映射（数据源：/api/stores 的 owner 字段聚合） */
      const map = new Map();
      this.stores.forEach(s => {
        if (!s.owner_user_id) return;
        if (!map.has(s.owner_user_id)) {
          map.set(s.owner_user_id, {
            user_id: s.owner_user_id,
            username: s.owner_username || s.owner_user_id,
            stores: [],
          });
        }
        map.get(s.owner_user_id).stores.push(s);
      });
      this.members = Array.from(map.values())
        .sort((a, b) => a.username.localeCompare(b.username, "zh-Hans-CN"));
      if (this.admin) {
        this.multi = true;
        if (this.memberId
            && !this.members.some(m => m.user_id === this.memberId)) {
          this.memberId = null;
          this._syncUrl();
        }
        const vis = this._visibleStores();
        if (this.storeId && !vis.some(s => s.store_id === this.storeId)) {
          this.storeId = null;
          this._syncUrl();
        }
      } else {
        this.multi = this.stores.length > 1;
        this.memberId = null;
        if (this.multi) {
          if (this.storeId
              && !this.stores.some(s => s.store_id === this.storeId)) {
            this.storeId = null;
            this._syncUrl();
          }
        } else {
          this.storeId = this.stores.length === 1
            ? this.stores[0].store_id : null;
        }
      }
    } catch (e) {
      this.stores = [];
      this.members = [];
      this.multi = false;
      this.admin = false;
      this.storeId = null;
      this.memberId = null;
    }
    this._render();
    /* 管理员/多店视图：数据就绪后触发一次重载（表头/统计需按聚合模式渲染） */
    if (this.multi) this.onChange();
  }

  /* 店铺被物理删除（store-deleted 事件）: 清除当前选择并重新加载,
   * 已删店铺不再出现在下拉与请求参数中, 页面经 onChange 重载数据 */
  _handleStoreDeleted(e) {
    const sid = e && e.detail && e.detail.store_id;
    if (sid && this.storeId === sid) {
      this.storeId = null;
      this._syncUrl();
    }
    this._load();
  }

  _visibleStores() {
    if (!this.memberId) return this.stores;
    const m = this.members.find(x => x.user_id === this.memberId);
    return m ? m.stores : [];
  }

  _render() {
    if (!this.admin && !this.multi) {
      const name = this.stores.length ? this.stores[0].name : "";
      this.el.className = "store-scope single";
      this.el.innerHTML =
        `<span class="ss-badge" title="当前店铺（普通成员固定本店）">🏪 ${esc(name || "本店数据")}</span>`;
      return;
    }
    if (!this.admin) {
      /* 防御路径：成员多店（当前单店模型不触发），保持原单维渲染 */
      this.el.className = "store-scope multi";
      this.el.innerHTML = `
        <label class="ss-label">店铺</label>
        <select class="ss-select" title="按店铺筛选（全部店铺 = 跨店汇总）">
          <option value="">${esc(this.allLabel)}</option>
          ${this.stores.map(s =>
            `<option value="${esc(s.store_id)}" ${s.store_id === this.storeId ? "selected" : ""}>${esc(s.name)}</option>`).join("")}
        </select>
        <span class="ss-count" title="当前登录身份可管理的店铺数量">${this.stores.length} 家店铺</span>`;
      this.el.querySelector(".ss-select").onchange = e => {
        this.storeId = e.target.value || null;
        this._syncUrl();
        this.onChange();
      };
      return;
    }
    /* 管理员双维级联：成员下拉 + 店铺下拉 */
    const vis = this._visibleStores();
    this.el.className = "store-scope multi admin-dual";
    this.el.innerHTML = `
      <label class="ss-label">成员</label>
      <select class="ss-select ss-member" title="按成员筛选（全部成员 = 全部店铺汇总）">
        <option value="">全部成员</option>
        ${this.members.map(m =>
          `<option value="${esc(m.user_id)}" ${m.user_id === this.memberId ? "selected" : ""}>${esc(m.username)}</option>`).join("")}
      </select>
      <label class="ss-label">店铺</label>
      <select class="ss-select ss-store" title="按店铺筛选（全部店铺 = 当前成员范围汇总）">
        <option value="">${esc(this.allLabel)}</option>
        ${vis.map(s =>
          `<option value="${esc(s.store_id)}" ${s.store_id === this.storeId ? "selected" : ""}>${esc(s.name)}</option>`).join("")}
      </select>
      <span class="ss-count" title="当前登录身份可管理的成员与店铺数量">${this.members.length} 名成员 · ${this.stores.length} 家店铺</span>`;
    this.el.querySelector(".ss-member").onchange = e => {
      this.memberId = e.target.value || null;
      const now = this._visibleStores();
      if (this.storeId && !now.some(s => s.store_id === this.storeId)) {
        this.storeId = null;
      }
      this._syncUrl();
      this._render();
      this.onChange();
    };
    this.el.querySelector(".ss-store").onchange = e => {
      this.storeId = e.target.value || null;
      this._syncUrl();
      this.onChange();
    };
  }

  /* 查询参数：admin → {member?, store_id?}；成员单店 → {store_id} 或 {} */
  getParams() {
    if (this.admin) {
      const p = {};
      if (this.memberId) p.member = this.memberId;
      if (this.storeId) p.store_id = this.storeId;
      return p;
    }
    return this.storeId ? { store_id: this.storeId } : {};
  }

  /* 合并进现有 URLSearchParams 或普通对象 */
  appendTo(params) {
    Object.entries(this.getParams()).forEach(([k, v]) => {
      if (typeof params.set === "function") params.set(k, v);
      else params[k] = v;
    });
    return params;
  }

  /* 写操作（执行对账/利润重算）目标 body 字段：
     全部 → {target_store_id: "all"}; 具体店 → {target_store_id: sid};
     admin 成员筛选下同步携带 member（后端收窄为该成员店铺）;
     单店/成员 → {}（后端默认本店） */
  writeTarget() {
    if (!this.multi) return {};
    const t = { target_store_id: this.storeId || "all" };
    if (this.admin && this.memberId) t.member = this.memberId;
    return t;
  }
}

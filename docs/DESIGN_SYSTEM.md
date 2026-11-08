# 微信小店对账系统 · Design System

轻量企业 ERP / 财务后台视觉规范。实现文件：`static/design-system.css`（令牌与公共组件），由 `static/style.css` 通过 `@import` 全局引入。

## 1. 色彩

| 用途 | Token | 色值 |
|------|--------|------|
| 主色 | `--color-primary` | `#16877C` |
| 主色悬停 | `--color-primary-hover` | `#13796F` |
| 主色按下 | `--color-primary-active` | `#0F6B62` |
| 主色浅底 | `--color-primary-subtle` | `#E8F5F3` |
| 页面背景 | `--color-bg-page` | `#F5F7FA` |
| 卡片背景 | `--color-bg-card` | `#FFFFFF` |
| 主文字 | `--color-text-primary` | `#263445` |
| 辅助文字 | `--color-text-secondary` | `#758194` |
| 边框 | `--color-border` | `#E2E8F0` |
| 成功 | `--color-success` | `#16A34A` |
| 警告 | `--color-warning` | `#D97706` |
| 危险 | `--color-danger` | `#DC2626` |
| 信息 | `--color-info` | `#0891B2` |

遗留别名（勿在新代码中继续扩散）：`--primary`、`--text`、`--text-light`、`--bg`、`--card`、`--border` 等均映射到上表。

## 2. 字体与排版

| Token | 值 |
|--------|-----|
| `--font-family` | 系统字体 + PingFang SC / 微软雅黑 |
| `--font-size-xs` ~ `--font-size-2xl` | 12 / 13 / 14 / 16 / 20 / 24 px |
| `--line-height-normal` | 1.55 |
| `--font-weight-medium` | 500 |
| `--font-weight-semibold` | 600 |
| `--font-weight-bold` | 700 |

- 页面标题：`.page-title`（20px / 700）
- 说明文字：`.page-desc`（辅助色 / 1.55 行高）
- 区块小标题：`.kpi-section-title`、`.card h3`

## 3. 布局与间距

| Token | 值 | 说明 |
|--------|-----|------|
| `--layout-max-width` | 1680px | 主内容最大宽度（居中） |
| `--layout-gutter` | 24px | 页面左右内边距 |
| `--card-gap` | 16px | 卡片纵向间距 |
| `--card-padding` | 16px | 卡片内边距 |
| `--card-radius` | 10px | 卡片圆角 |
| `--space-1` … `--space-8` | 4–32px | 统一间距阶梯 |

## 4. 按钮

| 类名 | 高度 Token | 场景 |
|------|------------|------|
| `.btn`（默认） | `--btn-height-md`（36px） | 主操作 |
| `.btn.btn-sm` | 32px | 表格/紧凑区 |
| `.btn.btn-lg` | 40px | 强调提交 |
| `.btn.secondary` | 同上 | 次要操作，白底描边 |
| `.btn.danger` | 同上 | 破坏性操作 |
| `:disabled` | opacity 0.45 | 禁用 |

悬停：主色按钮使用 `--color-primary-hover`；次要按钮浅灰底。

## 5. 表单控件

统一高度 `--control-height-md`（36px）、圆角 `--control-radius`（8px）、边框 `--color-border`。

已统一样式的选择器（勿在页面 `<style>` 重复）：

- `.filter-item` / `.filter-grid` 内 `input`、`select`
- `.month-filter`（`.mf-year`、`.mf-month`、`.mf-from`、`.mf-to`）
- `.modal-card` 表单
- `.create-field`、`.fee-field`

聚焦：`border-color: primary` + `box-shadow: 0 0 0 3px var(--focus-ring-color)`。

可选扩展类：`.ds-control`、`.ds-control--search`。

## 6. 弹窗、抽屉与确认

| 形态 | 结构 | 入口 |
|------|------|------|
| 居中弹窗 | `.modal-mask` + `.modal` | 列表页内联模态 |
| 业务弹窗 | `.modal-overlay` + `.modal-card.um-card` | `modal.js` → `openModal` |
| 右侧抽屉 | `.modal-overlay.modal-drawer` + `.um-card` | `modal.js` → `openDrawer` |
| 确认/表单 | `.modal-overlay` + `.modal-card` | `auth_nav.js`、管理页 |
| 移动侧栏 | `.sidebar-backdrop` + `body.sidebar-open` | `sidebar.js` |

统一遮罩：`--overlay-scrim`；对话框：`--shadow-dialog`（轻量）、1px `--color-border`，无厚重渐变。

## 7. 状态标签与金额

- 标签：`.badge` + `.gray` / `.blue` / `.green` / `.orange` / `.red` / `.cyan` / `.outline`
- 金额：`.amount--default|success|warning|danger|muted`；表格数字列 `td.num` + `tabular-nums`

JS 渲染金额时优先使用 `amount--*` 类，避免内联 `style="color:..."`。

## 8. 反馈状态

| 类名 | 用途 |
|------|------|
| `.loading-state` + `.spinner` | 加载中 |
| `.empty-state` | 无数据 |
| `.error-state` | 错误 |
| `.success-state` | 成功块（表单提交等） |
| `.toast` / `.toast.success|error|warn` | 全局提示（`api.js`） |
| `.create-msg.ok` / `.err` | 创建用户结果 |

## 9. 页面模块（自各页 `<style>` 迁移）

均在 `design-system.css` 中维护：

- 经营总览：`.kpi-*`、`.trend-*`、`.bs-*`
- 数据导入：`.import-zone`、`.batch-filters`、`.pv-actions`
- 每月盈利：`.fee-form`、`.fee-summary`
- 利润核算：`.formula-box`、`.policy-select`
- 对账成本：`.candidate-list`、`.risk-yes`、`.blocked`
- 创建用户：`.create-*`、`.store-option`

## 10. 使用约定

1. 业务页 **仅** 引用 `/static/style.css`，禁止复制色值/按钮高度。
2. 新组件优先使用 Design Token（`var(--color-*)`、`var(--space-*)`）。
3. 装饰克制：默认 `--shadow-sm`；避免粗边框（优先 1px `--color-border`）、大阴影与渐变背景。
4. 侧栏壳层变量保留在 `style.css`（`--sidebar-w` 等），与 Design System 主色一致。

## 11. 修改文件清单（本次）

| 文件 | 变更 |
|------|------|
| `static/design-system.css` | **新建**：令牌 + 公共组件 + 页面模块样式 |
| `static/style.css` | `@import` DS；布局/壳层；去除重复色板与 badge 色块 |
| `docs/DESIGN_SYSTEM.md` | **新建**：本文档 |
| `static/dashboard.html` | 移除内联 `<style>` |
| `static/index.html` | 移除内联 `<style>` |
| `static/monthly_profit.html` | 移除内联 `<style>` |
| `static/profit.html` | 移除内联 `<style>` |
| `static/reconciliation.html` | 移除内联 `<style>` |
| `static/create_user.html` | 移除内联 `<style>` |

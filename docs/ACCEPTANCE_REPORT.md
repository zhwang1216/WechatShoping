# 微信小店对账系统 — 回归测试与验收报告

**执行时间**：2026-11-08  
**执行环境**：本机 `python -m pytest tests/`（临时 `TestEnv` 数据根，**未读写生产/用户真实数据目录**）  
**全量结果**（最近一轮 2026-11-08 复跑）：`670 passed`，`14 failed`，`1 skipped`，`98 subtests passed`，耗时约 **474s**（日志：项目根目录 `.pytest_full_run.log`）  
**跳过用例**：`tests/test_recon_anomaly_fix.py`（约 L266）— `skipTest("需要至少两个店铺测试隔离")`（夹具仅单店时跳过）

---

## 总体验收结论：**有条件通过（未达最终签收）**

| 维度 | 结论 | 说明 |
|------|------|------|
| 自动化回归 | **未通过** | 14 项失败需分类处置（见修复清单） |
| 业务红线（10 项） | **大部分有自动化覆盖且通过** | 单项见下文第三节 |
| 视觉 / 多分辨率 | **未在本轮人工执行** | 仅有静态资源与壳层结构自动化检查 |
| 万级性能 | **未执行** | 仓库内无 ≥10000 条订单基准用例 |

---

## 一、视觉测试

| 检查项 | 方式 | 结果 |
|--------|------|------|
| 设计系统统一（style + design-system + responsive） | `test_ui_consolidation`、静态资源可访问 | **通过（自动化）** |
| 壳层单例、ListPage/DataTable 迁移 | `test_common_table`、`test_page_routes` | **通过（自动化）** |
| 1920/1440/1280/1024/768/390 布局、错位、溢出、按钮遮挡、弹窗 | 浏览器逐分辨率人工走查 | **未执行** — **不得标为通过** |
| 弹窗/抽屉响应式规则 | `responsive.css` 存在性断言 | **通过（自动化）** |

**说明**：本轮未启动 Cursor 浏览器对上述分辨率做全页面目视验收；请在签收前由 QA 按业务菜单 13 页 + 登录/注册补测。

---

## 二、功能测试（自动化映射）

| 能力 | 主要测试模块 | 本轮 |
|------|----------------|------|
| 导航 / 路由 | `test_page_routes`、`test_ui_consolidation` | 通过 |
| 筛选 / 月份 / 店铺 / 成员 | `test_month_filter`、`test_multi_store_admin` | 通过 |
| 分页 / 搜索 | `test_logs_api`、`test_order_pages` | 通过 |
| 导入 / 预览 / 确认 | `test_import_flow_revamp`（2 失败见下）、`test_delivery_e2e` | **部分失败** |
| 导出 | `test_excel_export`、`test_logs_api::test_logs_export_structured` | 通过（利润导出见 `test_profit` 失败） |
| 对账编辑 / 人工匹配 | `test_manual_cost_entry`、`test_reconciliation` | 通过 |
| 删除 / 清数 / 店铺生命周期 | `test_store_ownership`、`test_member_management`（1 页断言失败） | **大部分通过** |
| 权限 | `test_auth_isolation`、`test_admin_role_sep`（2 失败） | **部分失败** |
| 备份 / 恢复 | `test_json_repository`、`test_store_ownership` 清数/删除快照 | 通过 |

---

## 三、业务测试（用户指定的 10 项）

| # | 业务点 | 验收测试（代表） | 结果 |
|---|--------|------------------|------|
| 1 | 普通订单与退款订单分离 | `test_month_filter`、`test_order_pages` | **通过** |
| 2 | 多笔退款不重复计算 | `test_month_filter::test_statistics_consistent_with_table_refund_orders` | **通过** |
| 3 | 跨月退款统计 | `test_month_filter` 跨月/回退时间字段 | **通过** |
| 4 | 未发货供应商数据排除 | `test_import_unshipped_skip`、`test_supplier_cost_ship_validity`、`test_profit_unshipped` | **通过** |
| 5 | 综合成本不重复扣减 | `test_reconciliation::test_04_duplicate_order_level_cost` 等 | **通过** |
| 6 | 未匹配成本人工补录 | `test_manual_cost_entry`、`test_unmatched_order_cost` | **通过** |
| 7 | 新导入覆盖冲突的人工修改 | `test_recon_anomaly_fix`、`manual_cost_entry` 覆盖场景 | **通过** |
| 8 | 广告费为唯一允许手动编辑财务金额 | `test_profit_cost_summary`（`ad_expense` API）；**无独立 E2E「禁止改其他金额」用例** | **部分覆盖** — API 通过，UI 禁令需人工或补测 |
| 9 | 用户与店铺数据权限隔离 | `test_auth_isolation`、`test_multi_store_admin` | **通过** |
| 10 | 店铺归属修改失败回滚 | `test_store_ownership -k rollback`（2 例）**通过**；`test_admin_role_sep::test_09_atomic_rollback_and_retry` **失败** | **不一致** — 核心回滚路径有覆盖，原子失败路径断言与现网行为不符 |

**业务专项批跑**（9 个模块，140 passed）：`test_month_filter`、`test_auth_isolation`、`test_reconciliation`、`test_manual_cost_entry`、`test_recon_anomaly_fix`、`test_json_repository`、`test_ui_consolidation`、`test_logs_api`、`test_profit_cost_summary`。

---

## 四、性能测试

| 要求 | 结果 |
|------|------|
| ≥10000 条订单：查询 / 筛选 / 分页 / 导入 | **未执行** — 仓库无对应基准测试；仅 `test_admin_role_sep` / `test_store_ownership` 有小规模切换耗时断言 |

**风险**：万级数据下 JSONL 全量合并、管理端多店日志/异常列表的内存与响应时间**未经实测**。

---

## 五、数据安全

| 项 | 测试 | 结果 |
|----|------|------|
| JSON 读写 / 版本 | `test_json_repository` | 通过 |
| 操作日志只读、不删审计 | `test_logs_api::test dirty + tearDown 字节一致` | 通过 |
| 备份目录与恢复 | `test_store_ownership` 清数/删除/迁移 | 通过 |
| 并发写入（乐观） | `test_store_edit_consistency` | 在本轮全量中（未单独重跑，全量 670 pass 包含此类用例大类） |
| 进程异常与 API 不外泄 traceback | `test_logs_api::test_traceback_logged_not_leaked` | 通过 |

**未覆盖**：真实磁盘断电、多进程同时写同一 JSONL 的混沌测试 — **未执行**。

---

## 六、失败用例清单（14）与修复建议

| 用例 | 失败摘要 | 建议归类 |
|------|----------|----------|
| `test_profit.py` ×5 | `rules_version` 期望 `P1.2` 实际 `P1.3`；导出列 `None` | **测试与产品版本漂移**，更新断言/导出列序或确认利润规则发布说明 |
| `test_delivery_e2e.py` ×2 | 对账场景索引空；导出 sheet 数 7≠6 | **E2E 数据或导出格式变更**，更新夹具或期望 |
| `test_import_flow_revamp.py` ×2 | 成员自建店返回 403「需要管理员权限」 | **产品策略与旧测不一致** — 确认是否 intentional，同步改测或恢复能力 |
| `test_member_management.py::test_g01` | 用户管理页无 `data-act="delete"` | **UI 改版** — 更新测或恢复删除入口 |
| `test_store_ownership.py` ×2 | 静态 `stores.html` 缺 admin 导航标记；页内含「修改店铺归属」与旧断言冲突 | **壳层注入 vs 静态页断言** — 测应查注入后 HTML 或改静态标记 |
| `test_admin_role_sep.py` ×2 | 回滚期望 500 得 404；成员筛选项数量 13≠12 | **API 错误码/导航项数变更** — 对齐 spec 后改实现或改测 |

---

## 七、未完成事项

1. 全分辨率人工视觉验收 + 关键路径录屏归档。  
2. 万级订单性能基准（建议：生成夹具、`/api/orders` 分页 P95、导入耗时上限）。  
3. 利润模块 5 个失败用例清零或书面确认「仅测例过时」。  
4. 广告费「唯一可编辑」的前端负向测试（尝试改对账/利润字段应 403 或只读）。  
5. `test_admin_role_sep` 归属原子失败路径与 `test_store_ownership` 回滚断言统一口径。

---

## 八、风险说明

1. **测试债务**：14 个失败中有相当比例为测例与现网行为/版本号未同步，但不排除真实回归；签收前必须逐条定性。  
2. **性能未知**：无 1 万单压测，生产大店可能出现列表超时或内存尖峰。  
3. **视觉未验**：CSS 已响应式加固，但无逐页目视，不能排除个别页（利润工作台、导入预览）窄屏遮挡。  
4. **Python 3.14**：当前运行在 3.14.6；若生产环境版本不同，需再跑一轮全量。  
5. **真实数据**：本轮全部在 `TestEnv` 临时目录；**未对你们线上 `data/` 做任何操作**。

---

## 九、建议签收前命令

```bash
# 全量回归（约 8–9 分钟）
python -m pytest tests/ -q --tb=line

# 业务红线快检（约 2 分钟）
python -m pytest tests/test_month_filter.py tests/test_auth_isolation.py \
  tests/test_reconciliation.py tests/test_manual_cost_entry.py \
  tests/test_recon_anomaly_fix.py tests/test_profit_cost_summary.py -q
```

---

*本报告仅陈述已执行的自动化结果与明确未执行项，未对未测项标注为通过。*

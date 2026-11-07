# 微信小店对账与利润核算系统

本地运行的微信小店对账与利润核算工具：后端为 Python 本地 HTTP 服务，数据以 JSON 文件存储（无需数据库），前端为 HTML5 页面（后续阶段提供）。

- 适用系统：macOS 12+（Intel / Apple Silicon）、Linux
- Python：3.10+（开发验证环境为 Python 3.14.6）

## 目录结构

```
WechatShopping/
├── main.py              # 服务入口
├── start.sh             # macOS 启动脚本
├── stop.sh              # 停止脚本
├── requirements.txt     # Python 依赖
├── app/
│   ├── config.py        # 全局配置（路径/端口/备份策略）
│   ├── api_response.py  # 统一 API 响应格式
│   ├── data_schemas.py  # 业务 JSON 文件定义与初始数据
│   ├── json_repository.py # JSON 数据仓库（原子写入/备份/恢复/文件锁）
│   ├── operation_log.py # 操作日志（JSONL）
│   ├── page_shell.py    # 统一应用外壳（侧边栏/顶栏）注入
│   └── server.py        # 本地 HTTP 服务
├── static/              # 页面与公共前端组件（list_page/order_detail/modal 等）
├── data/                # 业务 JSON 数据（自动生成）
├── backups/             # 自动/手动备份（自动生成）
├── logs/                # 服务输出与操作日志（自动生成）
└── tests/               # 单元测试（unittest）
```

## 安装

```bash
cd WechatShopping
python3 -m pip install -r requirements.txt   # 阶段1 仅用标准库，可跳过；阶段2 Excel 导入需要 openpyxl
```

## 启动

方式一（推荐）：

```bash
./start.sh
# 自定义端口: WECHAT_SHOP_PORT=9000 ./start.sh
```

方式二（前台运行，Ctrl+C 停止）：

```bash
python3 main.py            # 默认 127.0.0.1:8765
python3 main.py --port 9000
python3 main.py --init-only   # 仅初始化业务数据文件后退出
```

启动后首次访问任意接口会自动初始化全部业务 JSON 文件（已存在且合法的不会重建）。

## 停止

```bash
./stop.sh
```

前台运行时直接按 Ctrl+C。

## 健康检查

```bash
curl http://127.0.0.1:8765/api/health
```

返回 `data.status = "ok"` 表示服务与全部数据文件正常；`data_files` 中列出每个业务文件的存在性、合法性与记录数。

## API 一览（阶段1）

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET | /api/health | 健康检查 |
| GET | /api/data/{name} | 读取业务数据 |
| POST | /api/data/{name} | 写入业务数据，body: `{"records": [...]}`（settings 为 `{"settings": {...}}`） |
| GET | /api/logs?limit=100&action=&operator= | 查询操作日志 |
| GET | /api/backups?name=orders | 列出某业务数据的备份 |
| POST | /api/backup | 手动备份，body: `{"name": "orders"}` |
| POST | /api/restore | 从备份恢复，body: `{"name": "orders", "backup_file": "orders.xxx.json"}` |

`{name}` 取值：orders / refund_orders / refund_details / import_records / supplier_costs / supplier_cost_items / reconciliation_results / profit_statements / anomalies / export_records / settings

统一响应格式：`{"code": 0, "message": "success", "data": ...}`，失败时 code 非 0。

## 备份与恢复

- **自动备份**：每次写入数据前，旧文件会自动复制到 `backups/`，文件名形如 `orders.20261017_153000_123456.json`；每个业务数据默认保留最近 30 份（可在 `app/config.py` 调整 `BACKUP_KEEP_COUNT`）。
- **手动备份**：`curl -X POST http://127.0.0.1:8765/api/backup -d '{"name":"orders"}' -H 'Content-Type: application/json'`
- **查看备份**：`curl 'http://127.0.0.1:8765/api/backups?name=orders'`
- **恢复**：`curl -X POST http://127.0.0.1:8765/api/restore -d '{"name":"orders","backup_file":"orders.xxx.json"}' -H 'Content-Type: application/json'`（恢复前会自动再备份当前文件）
- **损坏自愈**：若数据文件意外损坏（如断电），下次启动或 `--init-only` 时会自动把损坏文件转存为 `backups/xxx.corrupt.时间戳.json` 并重建初始文件。
- 直接拷贝 `data/` 目录也可整体备份；恢复时整体拷回即可。

## 数据安全机制

- **原子写入**：先写同目录临时文件并 fsync，再 `os.replace` 原子替换。写入中途断电/失败时，旧文件保持完整，不会出现半截 JSON。
- **文件锁**：基于 `fcntl.flock` 的跨进程锁，读共享、写排他，多进程/多线程并发读写不会损坏文件。
- **写前校验**：所有写入先经过 schema 校验（必需字段、集合类型），非法数据直接拒绝，不会落盘。

## 运行测试

```bash
python3 -m unittest discover -s tests -v
```

测试覆盖：统一响应格式、操作日志、JSON 仓库读写/原子性/备份/恢复/文件锁/损坏自愈、HTTP 服务健康检查与各接口（含模拟写入失败时旧文件不损坏）。

## 前端页面复用

为避免各页面重复维护同一套结构，前端按职责拆为公共组件，页面只保留业务配置：

- **应用外壳**：侧边栏/顶栏由 `app/page_shell.py` 统一生成，页面内只写 `<!--@app-shell-->`，服务端按当前路由注入并高亮菜单；管理员菜单标记块仍在服务端按会话剔除。
- **列表页**：`static/list_page.js` 收敛月份/成员/店铺筛选、查询参数、分页、统计卡片、导出、重置、空/错误态；`static/common_table.js` 负责表体渲染与分页控件。订单、退款订单、供应商成本、异常处理、操作日志均已接入。
- **订单详情**：普通订单与退款订单共用 `static/order_detail.js`，统一通过 `static/modal.js` 的右侧抽屉展示，不再各自维护详情弹窗。
- **认证页**：登录、注册、系统初始化共用 `style.css` 中的 `.auth-*` 卡片/表单样式。
- **公共样式**：利润正负色、统计卡片副标题、统计口径说明等已抽到 `style.css`，页面不再内联重复定义。

修改导航、列表行为或详情结构时，优先改上述公共文件，避免再次出现多页复制。

## 打包发布

将系统打包为 Windows 独立可执行程序（.exe），无需安装 Python 即可双击运行：

```bash
python build/build_exe.py          # 默认单目录模式
python build/build_exe.py --mode onefile  # 单文件模式
```

产物位于 `dist/` 目录，部署与异常排查详见 [build/README.md](build/README.md)。

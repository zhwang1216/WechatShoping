# 打包发布说明

将微信小店对账系统打包为 Windows 独立可执行程序（.exe），无需安装 Python 即可双击运行。

## 构建前置条件

- **操作系统**：Windows 10/11 64 位
- **Python**：3.8 及以上
- **PyInstaller**：6.0 及以上（`pip install pyinstaller`）
- **项目文件**：`main.py`、`app/`、`static/`、`requirements.txt` 齐全

## 构建命令

```bash
# 默认单目录模式（推荐，启动快）
python build/build_exe.py

# 单文件模式（生成单个 .exe，启动稍慢）
python build/build_exe.py --mode onefile

# 无控制台窗口（不推荐，无法查看启动日志）
python build/build_exe.py --no-console

# 构建前强制清理旧产物
python build/build_exe.py --clean
```

## 产物位置

- **单目录模式**：`dist/微信小店对账系统/` 目录（含 .exe + _internal/ 运行时资源）
- **单文件模式**：`dist/微信小店对账系统.exe` 单个文件

产物体积上限 100 MB，超限时输出告警（不阻断构建）。

## 部署步骤（终端用户）

1. 将 `dist/微信小店对账系统/` 目录**整体复制**到目标 Windows 机器任意位置（如桌面、D 盘）
2. 双击 `微信小店对账系统.exe`
3. 控制台窗口弹出，显示启动信息与服务地址
4. 在浏览器中访问控制台输出的 URL（默认 `http://127.0.0.1:19222`）
5. 使用完毕后按 `Ctrl+C` 停止服务

## 命令行参数

| 参数 | 说明 | 示例 |
|------|------|------|
| `--host` | 指定监听地址 | `微信小店对账系统.exe --host 0.0.0.0` |
| `--port` | 指定监听端口 | `微信小店对账系统.exe --port 8080` |
| `--init-only` | 仅初始化数据后退出 | `微信小店对账系统.exe --init-only` |

## 运行时数据目录

程序运行时会在 **exe 同级目录**自动创建以下目录：

- `data/` — 业务数据（用户、店铺、订单等）
- `backups/` — 数据备份
- `logs/` — 运行日志
- `uploads/` — 上传文件

这些目录**不包含在打包产物中**，运行时动态创建。重启后数据保留。

## 异常排查

### 缺少 Visual C++ 运行库
- **现象**：启动时系统弹窗提示缺少 DLL 文件
- **解决**：安装 [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe)

### 防火墙拦截
- **现象**：浏览器无法访问页面
- **解决**：在 Windows 防火墙中放行 19222 端口，或使用 `--port` 指定其他端口

### 端口被占用
- **现象**：控制台提示"端口 xxx 已被占用"
- **解决**：使用 `--port` 参数指定其他端口，或停止占用端口的进程

### 数据目录无写权限
- **现象**：控制台提示"数据目录创建失败，请将程序复制到有写入权限的目录运行"
- **解决**：将程序从系统保护目录（如 `C:\Program Files\`）复制到用户目录（如桌面）

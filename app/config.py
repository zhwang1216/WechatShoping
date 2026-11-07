"""全局配置：路径、监听地址、备份策略"""
import os

from .runtime_paths import resolve_paths

_paths = resolve_paths()

BASE_DIR = _paths.runtime_root
DATA_DIR = os.path.join(_paths.runtime_root, "data")
BACKUP_DIR = os.path.join(_paths.runtime_root, "backups")
LOG_DIR = os.path.join(_paths.runtime_root, "logs")
UPLOAD_DIR = os.path.join(_paths.runtime_root, "uploads")
STATIC_DIR = os.path.join(_paths.bundle_root, "static")

HOST = "127.0.0.1"
PORT = 19222

OPERATION_LOG_FILE = "operation_log.jsonl"

# 每个业务文件保留的自动/手动备份份数
BACKUP_KEEP_COUNT = 30

# 操作日志单次查询上限
LOG_QUERY_MAX = 1000

# 上传文件大小限制（MB）与导入预览行数
MAX_UPLOAD_MB = 50
PREVIEW_ROWS = 20

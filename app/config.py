"""全局配置：路径、监听地址、备份策略"""
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
LOG_DIR = os.path.join(BASE_DIR, "logs")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
STATIC_DIR = os.path.join(BASE_DIR, "static")

HOST = "127.0.0.1"
PORT = 8765

OPERATION_LOG_FILE = "operation_log.jsonl"

# 每个业务文件保留的自动/手动备份份数
BACKUP_KEEP_COUNT = 30

# 操作日志单次查询上限
LOG_QUERY_MAX = 1000

# 上传文件大小限制（MB）与导入预览行数
MAX_UPLOAD_MB = 50
PREVIEW_ROWS = 20

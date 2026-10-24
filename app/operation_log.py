"""操作日志：JSONL 追加写入与查询（支持月份/日期区间左闭右开筛选）"""
import json
import os
import threading
from datetime import datetime

from . import config
from .time_range import in_range


def _normalize_operator(op):
    """operator 读取类型契约：返回值恒为 str 或 None（可哈希）。

    历史脏数据（旧版 store.rename）曾把整个用户 dict（含
    password_hash）写入 operator，穿透到调用方集合去重时抛
    TypeError（unhashable type: 'dict'），且 dict 形态会外泄敏感
    字段。归一化规则：
    - dict: username 非空取之，否则 user_id 非空取之，否则 "unknown"
    - None: 保持 None（统计集合排除、前端显示 —）
    - str:  原样
    - 其他（数字等标量）: str() 转换；list 等不可哈希容器: "unknown"
    纯函数，不修改传入对象。
    """
    if op is None:
        return None
    if isinstance(op, str):
        return op
    if isinstance(op, dict):
        username = op.get("username")
        if isinstance(username, str) and username:
            return username
        user_id = op.get("user_id")
        if isinstance(user_id, str) and user_id:
            return user_id
        return "unknown"
    if isinstance(op, (list, tuple, set)):
        return "unknown"
    return str(op)


class OperationLog:
    def __init__(self, log_dir=None, filename=None):
        self.log_dir = log_dir or config.LOG_DIR
        os.makedirs(self.log_dir, exist_ok=True)
        self.log_file = os.path.join(self.log_dir, filename or config.OPERATION_LOG_FILE)
        self._io_lock = threading.Lock()

    def append(self, action, detail=None, operator="system", result="success"):
        # 写入端防御性归一化：调用方误传用户 dict（重蹈旧版 store.rename
        # 覆辙）时只落 username 字符串；None 兜底为 "system"（与默认值一致）。
        op = _normalize_operator(operator)
        if op is None:
            op = "system"
        entry = {
            "time": datetime.now().isoformat(timespec="seconds"),
            "operator": op,
            "action": action,
            "result": result,
            "detail": detail,
        }
        line = json.dumps(entry, ensure_ascii=False)
        with self._io_lock:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(line + "\n")
                f.flush()
                os.fsync(f.fileno())
        return entry

    def query(self, limit=100, action=None, operator=None,
              start=None, end=None, keyword=None):
        """按操作发生时间（time 字段）左闭右开筛选；limit=None 返回全部匹配。

        读取类型契约：返回条目的 operator 恒为 str 或 None（含历史
        dict 脏数据也在解析后立即归一化，先于一切筛选），保证调用方
        集合去重安全且敏感字段不外泄。
        """
        if not os.path.exists(self.log_file):
            return []
        entries = []
        with self._io_lock:
            with open(self.log_file, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        e = json.loads(line)
                    except ValueError:
                        continue
                    if isinstance(e, dict):
                        e["operator"] = _normalize_operator(e.get("operator"))
                    entries.append(e)
        if action:
            entries = [e for e in entries if e.get("action") == action]
        if operator:
            entries = [e for e in entries if e.get("operator") == operator]
        if start is not None or end is not None:
            entries = [e for e in entries
                       if in_range(e.get("time"), start, end)]
        if keyword:
            kw = keyword.lower()
            entries = [e for e in entries
                       if kw in json.dumps(e, ensure_ascii=False).lower()]
        entries.reverse()
        if limit is None:
            return entries
        return entries[:max(0, limit)]

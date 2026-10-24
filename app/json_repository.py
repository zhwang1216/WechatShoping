"""JSON 数据仓库：读取、校验、原子写入、备份、恢复、文件锁

安全性设计:
1. 原子写入: 先写同目录临时文件并 fsync, 再 os.replace 原子替换,
   任何一步失败都不会损坏旧文件;
2. 文件锁: 基于 fcntl.flock 的跨进程排他/共享锁, 保护并发读写;
3. 写前备份: write/update 前自动将旧文件复制到 backups/;
4. 损坏自愈: ensure_all 发现损坏文件时, 先把损坏文件转存到 backups/
   供人工修复, 再重建初始文件。
"""
import fcntl
import json
import os
import shutil
import tempfile
from datetime import datetime

from . import config
from .data_schemas import BUSINESS_FILES


class RepositoryError(Exception):
    """数据仓库操作异常"""


class FileLock:
    """跨进程文件锁 (macOS/Linux 使用 fcntl.flock)"""

    def __init__(self, path, exclusive=True):
        self._path = path
        self._exclusive = exclusive
        self._fd = None

    def __enter__(self):
        self._fd = open(self._path, "a+", encoding="utf-8")
        mode = fcntl.LOCK_EX if self._exclusive else fcntl.LOCK_SH
        fcntl.flock(self._fd.fileno(), mode)
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._fd is not None:
            try:
                fcntl.flock(self._fd.fileno(), fcntl.LOCK_UN)
            finally:
                self._fd.close()
                self._fd = None
        return False


class JsonRepository:
    def __init__(self, data_dir=None, backup_dir=None):
        self.data_dir = data_dir or config.DATA_DIR
        self.backup_dir = backup_dir or config.BACKUP_DIR
        os.makedirs(self.data_dir, exist_ok=True)
        os.makedirs(self.backup_dir, exist_ok=True)

    # ---------- 基础工具 ----------
    @staticmethod
    def _ts():
        return datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    def _spec(self, name):
        if name not in BUSINESS_FILES:
            raise RepositoryError(f"未知的业务数据文件: {name}")
        return BUSINESS_FILES[name]

    def _path(self, name):
        return os.path.join(self.data_dir, self._spec(name)["file"])

    def _lock(self, name, exclusive=True):
        return FileLock(self._path(name) + ".lock", exclusive=exclusive)

    # ---------- 校验 ----------
    def validate(self, name, data):
        spec = self._spec(name)
        if not isinstance(data, dict):
            raise RepositoryError(f"{spec['file']}: 数据必须是 JSON 对象")
        for key in spec["required_keys"]:
            if key not in data:
                raise RepositoryError(f"{spec['file']}: 缺少必需字段 {key}")
        if not isinstance(data["version"], int):
            raise RepositoryError(f"{spec['file']}: version 必须是整数")
        if not isinstance(data["updated_at"], str):
            raise RepositoryError(f"{spec['file']}: updated_at 必须是字符串")
        ckey = spec["collection_key"]
        ctype = spec["collection_type"]
        if not isinstance(data[ckey], ctype):
            want = "数组" if ctype is list else "对象"
            raise RepositoryError(f"{spec['file']}: {ckey} 必须是{want}")
        return True

    # ---------- 初始化 ----------
    def ensure_all(self):
        """确保全部业务文件存在且合法, 幂等; 返回新建/修复清单"""
        created, repaired = [], []
        for name in BUSINESS_FILES:
            state = self._ensure_one(name)
            if state == "created":
                created.append(name)
            elif state == "repaired":
                repaired.append(name)
        return {"created": created, "repaired": repaired}

    def _ensure_one(self, name):
        spec = self._spec(name)
        path = self._path(name)
        with self._lock(name):
            if os.path.exists(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    self.validate(name, data)
                    return "ok"
                except (OSError, ValueError, RepositoryError):
                    corrupt = os.path.join(
                        self.backup_dir, f"{name}.corrupt.{self._ts()}.json")
                    try:
                        shutil.copy2(path, corrupt)
                    except OSError:
                        corrupt = None
                    self._atomic_write_json(path, spec["initial"]())
                    return "repaired"
            self._atomic_write_json(path, spec["initial"]())
            return "created"

    # ---------- 读取 ----------
    def read(self, name):
        spec = self._spec(name)
        path = self._path(name)
        with self._lock(name, exclusive=False):
            if not os.path.exists(path):
                raise RepositoryError(f"{spec['file']}: 文件不存在，请先初始化")
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        self.validate(name, data)
        return data

    # ---------- 原子写入 ----------
    @staticmethod
    def _atomic_write_json(path, data):
        directory = os.path.dirname(path)
        fd, tmp_path = tempfile.mkstemp(prefix=".tmp_", suffix=".json", dir=directory)
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_path, path)
        except BaseException:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
            raise

    def write(self, name, data, backup=True):
        """整体写入, 写前自动备份旧文件, 原子替换"""
        self.validate(name, data)
        path = self._path(name)
        with self._lock(name):
            if backup and os.path.exists(path):
                self._do_backup(name)
            self._atomic_write_json(path, data)
            self._prune_backups(name)
        return data

    def update(self, name, mutator, backup=True):
        """持锁的读-改-写: 在锁内读取、修改、写回, 避免并发丢失更新"""
        spec = self._spec(name)
        path = self._path(name)
        with self._lock(name):
            if not os.path.exists(path):
                raise RepositoryError(f"{spec['file']}: 文件不存在，请先初始化")
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.validate(name, data)
            result = mutator(data)
            data = result if result is not None else data
            self.validate(name, data)
            if backup and os.path.exists(path):
                self._do_backup(name)
            self._atomic_write_json(path, data)
            self._prune_backups(name)
        return data

    # ---------- 备份 ----------
    def _do_backup(self, name):
        src = self._path(name)
        dst = os.path.join(self.backup_dir, f"{name}.{self._ts()}.json")
        shutil.copy2(src, dst)
        return os.path.basename(dst)

    def backup(self, name):
        """手动备份指定业务文件"""
        self._spec(name)
        with self._lock(name):
            if not os.path.exists(self._path(name)):
                raise RepositoryError(f"数据文件不存在，无法备份: {name}")
            filename = self._do_backup(name)
            self._prune_backups(name)
        return filename

    def list_backups(self, name):
        self._spec(name)
        prefix = f"{name}."
        result = []
        if not os.path.isdir(self.backup_dir):
            return result
        for fn in os.listdir(self.backup_dir):
            if fn.startswith(prefix) and fn.endswith(".json") and ".corrupt." not in fn:
                full = os.path.join(self.backup_dir, fn)
                st = os.stat(full)
                result.append({
                    "file": fn,
                    "size": st.st_size,
                    "modified_at": datetime.fromtimestamp(
                        st.st_mtime).isoformat(timespec="seconds"),
                    "_mtime": st.st_mtime,
                })
        # 按修改时间倒序（亚秒精度）：同一秒内连续写入时也能稳定区分新旧
        result.sort(key=lambda x: x["_mtime"], reverse=True)
        for item in result:
            item.pop("_mtime", None)
        return result

    def _prune_backups(self, name, keep=None):
        keep = config.BACKUP_KEEP_COUNT if keep is None else keep
        removed = []
        for item in self.list_backups(name)[keep:]:
            try:
                os.remove(os.path.join(self.backup_dir, item["file"]))
                removed.append(item["file"])
            except OSError:
                pass
        return removed

    # ---------- 恢复 ----------
    def restore(self, name, backup_file):
        """从备份文件恢复, 恢复前会先备份当前文件"""
        self._spec(name)
        safe_name = os.path.basename(str(backup_file))
        src = os.path.join(self.backup_dir, safe_name)
        if not os.path.isfile(src):
            raise RepositoryError(f"备份文件不存在: {safe_name}")
        with open(src, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.validate(name, data)
        path = self._path(name)
        with self._lock(name):
            if os.path.exists(path):
                self._do_backup(name)
            self._atomic_write_json(path, data)
        return data

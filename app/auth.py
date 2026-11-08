"""认证与用户管理：密码哈希、用户/店铺注册表、会话管理、登录限流、安全日志

安全设计:
1. 密码使用 PBKDF2-HMAC-SHA256（NIST SP 800-132 认可、Django 默认方案）加盐哈希,
   格式: pbkdf2_sha256$<iterations>$<salt_hex>$<dk_hex>, 校验使用常量时间比较;
2. 用户表(users.json)/店铺表(stores.json)/会话表(sessions.json)存于数据根目录,
   原子写入 + 跨进程文件锁;
3. 会话令牌仅存在于 HttpOnly Cookie 与服务端, 安全日志不记录密码与令牌;
4. 登录失败统一提示"账号或密码错误", 不泄露账号是否存在;
5. 登录限流: 同一 (IP, 用户名) 15 分钟内最多 10 次失败。
"""
import hashlib
import hmac
import json
import os
import secrets
import shutil
import sys
import threading
import time
from datetime import datetime, timedelta

from .data_schemas import now_iso
from .json_repository import FileLock

PBKDF2_ITERATIONS = 600_000
DK_LEN = 32
SESSION_TTL_HOURS = 12
LOGIN_WINDOW_SECONDS = 15 * 60
LOGIN_MAX_FAILURES = 10
REGISTER_WINDOW_SECONDS = 60 * 60
REGISTER_MAX_REQUESTS = 5

USERNAME_MIN, USERNAME_MAX = 3, 32
PASSWORD_MIN = 8


class AuthError(Exception):
    """认证/授权错误; message 面向用户（中文）"""

    def __init__(self, message, code=400, http_status=400, extra=None):
        super().__init__(message)
        self.code = code
        self.http_status = http_status
        self.extra = extra or {}


# ==================== 密码哈希 ====================
def hash_password(password):
    if not isinstance(password, str) or not password:
        raise AuthError("密码不能为空")
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt,
                             PBKDF2_ITERATIONS, dklen=DK_LEN)
    return f"pbkdf2_sha256${PBKDF2_ITERATIONS}${salt.hex()}${dk.hex()}"


def verify_password(password, encoded):
    try:
        algo, iterations, salt_hex, dk_hex = encoded.split("$")
        if algo != "pbkdf2_sha256":
            return False
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"),
                                 bytes.fromhex(salt_hex), int(iterations),
                                 dklen=len(bytes.fromhex(dk_hex)))
        return hmac.compare_digest(dk.hex(), dk_hex)
    except (ValueError, AttributeError, TypeError):
        return False


def validate_username(username):
    if not isinstance(username, str):
        raise AuthError("用户名必须是字符串")
    username = username.strip()
    if not (USERNAME_MIN <= len(username) <= USERNAME_MAX):
        raise AuthError(f"用户名长度需在 {USERNAME_MIN}-{USERNAME_MAX} 个字符之间")
    if not all(c.isalnum() or c in "_-." for c in username):
        raise AuthError("用户名只能包含字母、数字、下划线、连字符和点")
    return username


def validate_password_strength(password):
    if not isinstance(password, str) or len(password) < PASSWORD_MIN:
        raise AuthError(f"密码至少 {PASSWORD_MIN} 位")
    if not any(c.isdigit() for c in password) or not any(c.isalpha() for c in password):
        raise AuthError("密码必须同时包含字母和数字")
    return True


# ==================== 安全日志（全局） ====================
class SecurityLog:
    """全局安全日志: 登录/登出/用户管理/权限拒绝等; 严禁记录密码与会话令牌"""

    def __init__(self, log_dir):
        self.log_dir = log_dir
        os.makedirs(self.log_dir, exist_ok=True)
        self.log_file = os.path.join(log_dir, "security_log.jsonl")
        self._lock = threading.Lock()

    def append(self, event, username=None, user_id=None, ip=None,
               result="success", detail=None):
        entry = {
            "time": now_iso(),
            "event": event,
            "username": username,
            "user_id": user_id,
            "ip": ip,
            "result": result,
            "detail": detail,
        }
        line = json.dumps(entry, ensure_ascii=False)
        with self._lock:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(line + "\n")
                f.flush()
                os.fsync(f.fileno())
        return entry


# ==================== 登录限流 ====================
class LoginRateLimiter:
    def __init__(self, window_seconds=LOGIN_WINDOW_SECONDS,
                 max_failures=LOGIN_MAX_FAILURES):
        self.window = window_seconds
        self.max_failures = max_failures
        self._failures = {}
        self._lock = threading.Lock()

    def _key(self, ip, username):
        return f"{ip}|{username}"

    def is_blocked(self, ip, username):
        key = self._key(ip, username)
        now = time.monotonic()
        with self._lock:
            stamps = [t for t in self._failures.get(key, []) if now - t < self.window]
            self._failures[key] = stamps
            return len(stamps) >= self.max_failures

    def record_failure(self, ip, username):
        key = self._key(ip, username)
        with self._lock:
            self._failures.setdefault(key, []).append(time.monotonic())

    def clear(self, ip, username):
        with self._lock:
            self._failures.pop(self._key(ip, username), None)


# ==================== 注册限流 ====================
class RegistrationRateLimiter:
    """公开注册接口限流: 同一 IP 滑动窗口内请求次数上限（成败均计）

    与 LoginRateLimiter 的语义差异:
    1. 键空间仅 IP（无用户名拼接）——注册的用户名可任意变化;
    2. 成败均计（无 clear 方法）——成功后不重置, 防止脚本循环注册不同账号;
    3. 窗口 60 分钟 / 5 次——注册为低频操作, 阈值远严于登录;
    4. 与登录限流实例互不共享任何计数。
    """

    def __init__(self, window_seconds=REGISTER_WINDOW_SECONDS,
                 max_requests=REGISTER_MAX_REQUESTS):
        self.window = window_seconds
        self.max_requests = max_requests
        self._requests = {}
        self._lock = threading.Lock()

    def is_blocked(self, ip):
        now = time.monotonic()
        with self._lock:
            stamps = [t for t in self._requests.get(ip, [])
                      if now - t < self.window]
            self._requests[ip] = stamps
            return len(stamps) >= self.max_requests

    def record_request(self, ip):
        """判定通过后由调用方调用一次; 此后无论注册成败均不再操作"""
        with self._lock:
            self._requests.setdefault(ip, []).append(time.monotonic())


# ==================== JSON 结构兼容（登录读路径） ====================
_DOC_META_KEYS = {"version", "updated_at"}


def _read_json_file(path):
    """读取 JSON 文件。不存在返回 None；内容损坏时抛出明确错误且不改写原文件。"""
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        raise AuthError(
            f"{os.path.basename(path)} 不是合法的 JSON，已保留原文件",
            code=500, http_status=500,
            extra={"data_file": os.path.basename(path),
                   "value_type": "invalid_json"}) from e


def _atomic_write_json(path, data):
    """同目录临时文件写入并 fsync 后原子替换。失败时删除临时文件，不改原文件。"""
    import tempfile
    directory = os.path.dirname(path)
    fd, tmp = tempfile.mkstemp(prefix=".tmp_", suffix=".json", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    except BaseException:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def _backup_json_file(path):
    """把原文件复制到同级 backups/ 并校验字节一致。校验失败不改原文件。"""
    backup_dir = os.path.join(os.path.dirname(path), "backups")
    os.makedirs(backup_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    dest = os.path.join(
        backup_dir, f"{os.path.basename(path)}.pre_migrate.{stamp}")
    shutil.copy2(path, dest)
    with open(path, "rb") as src, open(dest, "rb") as bak:
        if src.read() != bak.read():
            raise AuthError(
                f"{os.path.basename(path)} 备份校验失败，已中止且未修改原文件",
                code=500, http_status=500,
                extra={"data_file": os.path.basename(path),
                       "value_type": "backup_mismatch"})
    return dest


def _restore_json_file(backup, path):
    shutil.copy2(backup, path)
    with open(backup, "rb") as src, open(path, "rb") as dst:
        if src.read() != dst.read():
            raise AuthError(
                f"{os.path.basename(path)} 恢复原始文件失败",
                code=500, http_status=500,
                extra={"data_file": os.path.basename(path),
                       "value_type": "restore_failed",
                       "rollback_failed": True})


def _log_migration(data_file, field, old_type, new_type, kept):
    sys.stderr.write(
        f"[data.migrate] file={data_file} field={field} "
        f"from={old_type} to={new_type} kept={kept}\n")


def _json_records(records):
    return sorted(json.dumps(r, ensure_ascii=False, sort_keys=True)
                  for r in records)


def _records_equivalent(left, right):
    if not isinstance(left, list) or not isinstance(right, list):
        return False
    return _json_records(left) == _json_records(right)


def _shape_error(data_file, field, value, detail):
    return AuthError(
        detail,
        code=500, http_status=500,
        extra={"data_file": data_file, "field": field,
               "value_type": type(value).__name__})


def normalize_collection_document(data, collection_key, id_field, data_file):
    """把用户/店铺表整理为 {version, updated_at, collection: [记录...]}。

    返回 (document, changed)。合法的「对象包列表」原样返回且 changed=False，
    不触发写盘。顶层列表、或集合字段本身是对象时，转换成列表并保留全部记录。
    无法识别的条目直接报错，调用方不得写盘。
    """
    if isinstance(data, list):
        records = _records_from_list(data, collection_key, data_file)
        return {
            "version": 1,
            "updated_at": now_iso(),
            collection_key: records,
        }, True
    if not isinstance(data, dict):
        raise _shape_error(
            data_file, collection_key, data,
            f"{data_file} 顶层是 {type(data).__name__}，无法读取，已保留原文件")
    if collection_key not in data:
        if isinstance(data.get(id_field), str):
            return {
                "version": data.get("version") if isinstance(data.get("version"), int) else 1,
                "updated_at": data.get("updated_at") if isinstance(data.get("updated_at"), str) else now_iso(),
                collection_key: [dict(data)],
            }, True
        extra = set(data) - _DOC_META_KEYS
        if extra:
            raise _shape_error(
                data_file, collection_key, data,
                f"{data_file} 缺少 {collection_key} 字段，已保留原文件")
        doc = dict(data)
        doc[collection_key] = []
        if not isinstance(doc.get("version"), int):
            doc["version"] = 1
        if not isinstance(doc.get("updated_at"), str):
            doc["updated_at"] = now_iso()
        return doc, True
    collection = data[collection_key]
    if isinstance(collection, list):
        _records_from_list(collection, collection_key, data_file)
        return data, False
    if isinstance(collection, dict):
        records = []
        for index, (key, value) in enumerate(collection.items()):
            if not isinstance(value, dict):
                raise _shape_error(
                    data_file, collection_key, value,
                    f"{data_file} 的 {collection_key} 第 {index + 1} 条不是对象"
                    f"（实际类型 {type(value).__name__}），已保留原文件")
            record = value
            if id_field not in record and isinstance(key, str) and key:
                record = dict(value)
                record[id_field] = key
            records.append(record)
        doc = dict(data)
        doc[collection_key] = records
        if not isinstance(doc.get("version"), int):
            doc["version"] = 1
        if not isinstance(doc.get("updated_at"), str):
            doc["updated_at"] = now_iso()
        return doc, True
    raise _shape_error(
        data_file, collection_key, collection,
        f"{data_file} 的 {collection_key} 是 {type(collection).__name__}，"
        f"无法按列表读取，已保留原文件")


def _records_from_list(items, collection_key, data_file):
    records = []
    for index, item in enumerate(items):
        if not isinstance(item, dict):
            raise _shape_error(
                data_file, collection_key, item,
                f"{data_file} 的 {collection_key} 第 {index + 1} 条不是对象"
                f"（实际类型 {type(item).__name__}），已保留原文件")
        records.append(item)
    return records


def _commit_collection_migration(path, document, collection_key, id_field, old_type):
    """调用方已持有排他锁。备份 → 临时文件原子替换 → 读回校验；失败则恢复原文件。"""
    data_file = os.path.basename(path)
    backup = _backup_json_file(path)
    expected = document.get(collection_key)
    try:
        _atomic_write_json(path, document)
        written = _read_json_file(path)
        got = written.get(collection_key) if isinstance(written, dict) else None
        if not _records_equivalent(expected, got):
            raise _shape_error(
                data_file, collection_key, got,
                f"{data_file} 迁移后校验失败，已恢复原始文件")
    except Exception:
        try:
            _restore_json_file(backup, path)
        except Exception as restore_error:
            sys.stderr.write(
                f"[data.migrate] file={data_file} field={collection_key} "
                f"restore_failed={type(restore_error).__name__}\n")
            raise
        sys.stderr.write(
            f"[data.migrate] file={data_file} field={collection_key} "
            f"from={old_type} failed restored=yes\n")
        raise
    _log_migration(data_file, collection_key, old_type, "list", len(expected))


def _collection_value_type(raw, collection_key):
    if isinstance(raw, list):
        return "list"
    if isinstance(raw, dict):
        return type(raw.get(collection_key)).__name__
    return type(raw).__name__


def _session_entry(item, index):
    """从一条历史会话记录取出 (token, session)。无法识别时抛错，不丢弃。"""
    if isinstance(item, dict):
        token = item.get("token")
        if isinstance(token, str) and token:
            session = dict(item)
            session.pop("token", None)
            return token, session
        if len(item) == 1:
            key, value = next(iter(item.items()))
            if (isinstance(key, str) and key and isinstance(value, dict)
                    and value.get("user_id")):
                return key, dict(value)
        if item.get("user_id"):
            material = json.dumps(
                {k: item.get(k) for k in (
                    "user_id", "created_at", "expires_at",
                    "password_version", "store_id")},
                ensure_ascii=False, sort_keys=True)
            token = "recovered-" + hashlib.sha256(material.encode()).hexdigest()[:24]
            return token, dict(item)
        raise _shape_error(
            "sessions.json", "sessions", item,
            f"会话数据第 {index + 1} 条缺少用户标识，已保留原文件")
    if (isinstance(item, (list, tuple)) and len(item) == 2
            and isinstance(item[0], str) and item[0]
            and isinstance(item[1], dict)):
        return item[0], dict(item[1])
    raise _shape_error(
        "sessions.json", "sessions", item,
        f"会话数据第 {index + 1} 条无法识别"
        f"（实际类型 {type(item).__name__}），已保留原文件")


def _sessions_from_list(items):
    sessions = {}
    for index, item in enumerate(items):
        token, session = _session_entry(item, index)
        if token in sessions:
            raise _shape_error(
                "sessions.json", "sessions", item,
                f"会话数据第 {index + 1} 条与已有会话冲突，已保留原文件")
        sessions[token] = session
    return sessions


def _sessions_value_type(raw):
    if isinstance(raw, list):
        return "list"
    if isinstance(raw, dict):
        return type(raw.get("sessions")).__name__
    return type(raw).__name__


def normalize_sessions_document(data):
    """会话表规范形是 {sessions: {token: session}}。

    历史文件里 sessions 可能是列表。空列表变成空对象；列表中的会话记录
    按 token 收进对象，一条都不丢。已经是对象则 changed=False，不写盘。
    """
    if isinstance(data, list):
        return {
            "version": 1,
            "updated_at": now_iso(),
            "sessions": _sessions_from_list(data),
        }, True
    if not isinstance(data, dict):
        raise _shape_error(
            "sessions.json", "sessions", data,
            f"sessions.json 顶层是 {type(data).__name__}，无法读取，已保留原文件")
    if "sessions" not in data:
        if isinstance(data.get("token"), str) and data.get("user_id"):
            token = data["token"]
            session = dict(data)
            session.pop("token", None)
            return {
                "version": 1,
                "updated_at": now_iso(),
                "sessions": {token: session},
            }, True
        extra = set(data) - _DOC_META_KEYS
        if extra:
            raise _shape_error(
                "sessions.json", "sessions", data,
                "sessions.json 缺少 sessions 字段，已保留原文件")
        doc = dict(data)
        doc["sessions"] = {}
        if not isinstance(doc.get("version"), int):
            doc["version"] = 1
        if not isinstance(doc.get("updated_at"), str):
            doc["updated_at"] = now_iso()
        return doc, True
    sessions = data["sessions"]
    if isinstance(sessions, dict):
        for index, (key, value) in enumerate(sessions.items()):
            if not isinstance(key, str) or not isinstance(value, dict):
                raise _shape_error(
                    "sessions.json", "sessions", value,
                    f"会话数据第 {index + 1} 条不是对象"
                    f"（实际类型 {type(value).__name__}），已保留原文件")
        return data, False
    if isinstance(sessions, list):
        doc = dict(data)
        doc["sessions"] = _sessions_from_list(sessions)
        if not isinstance(doc.get("version"), int):
            doc["version"] = 1
        doc["updated_at"] = now_iso()
        return doc, True
    if sessions is None:
        doc = dict(data)
        doc["sessions"] = {}
        return doc, True
    raise _shape_error(
        "sessions.json", "sessions", sessions,
        f"sessions.json 的 sessions 是 {type(sessions).__name__}，"
        f"无法读取，已保留原文件")


def _sessions_equivalent(left, right):
    if not isinstance(left, dict) or not isinstance(right, dict):
        return False
    if set(left) != set(right):
        return False
    return all(left[key] == right[key] for key in left)


def _commit_sessions_migration(path, document, old_type):
    """调用方已持有排他锁。备份、原子写入、读回校验，失败恢复原文件。"""
    backup = _backup_json_file(path)
    expected = document.get("sessions")
    try:
        _atomic_write_json(path, document)
        written = _read_json_file(path)
        got = written.get("sessions") if isinstance(written, dict) else None
        if not _sessions_equivalent(expected, got):
            raise _shape_error(
                "sessions.json", "sessions", got,
                "sessions.json 迁移后校验失败，已恢复原始文件")
    except Exception:
        try:
            _restore_json_file(backup, path)
        except Exception as restore_error:
            sys.stderr.write(
                "[data.migrate] file=sessions.json field=sessions "
                f"restore_failed={type(restore_error).__name__}\n")
            raise
        sys.stderr.write(
            f"[data.migrate] file=sessions.json field=sessions "
            f"from={old_type} failed restored=yes\n")
        raise
    _log_migration("sessions.json", "sessions", old_type, "dict", len(expected))


# ==================== 会话存储 ====================
class SessionStore:
    """服务端会话: token -> {user_id, store_id, ...}; 持久化支持服务重启后保持登录"""

    def __init__(self, path):
        self.path = path
        self._lock_path = path + ".lock"
        directory = os.path.dirname(path)
        os.makedirs(directory, exist_ok=True)
        self._io_lock = threading.Lock()

    def _load(self):
        if not os.path.exists(self.path):
            return {"version": 1, "updated_at": now_iso(), "sessions": {}}
        with FileLock(self._lock_path, exclusive=False):
            raw = _read_json_file(self.path)
        document, changed = normalize_sessions_document(raw)
        if not changed:
            return document
        # 共享锁已释放后再取排他锁，避免同一进程 flock 自死锁
        with FileLock(self._lock_path):
            raw = _read_json_file(self.path)
            document, changed = normalize_sessions_document(raw)
            if changed:
                _commit_sessions_migration(
                    self.path, document, _sessions_value_type(raw))
            return document

    def _save(self, data):
        directory = os.path.dirname(self.path)
        fd = -1
        tmp = None
        with FileLock(self._lock_path):
            with self._io_lock:
                import tempfile
                fd, tmp = tempfile.mkstemp(prefix=".tmp_sessions_",
                                           suffix=".json", dir=directory)
                with os.fdopen(fd, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(tmp, self.path)

    def create(self, user, store_id):
        token = secrets.token_urlsafe(32)
        now = datetime.now()
        expires = now + timedelta(hours=SESSION_TTL_HOURS)
        data = self._load()
        sessions = data.get("sessions")
        if not isinstance(sessions, dict):
            raise AuthError(
                "会话数据异常：sessions 不是对象，已保留原文件",
                code=500, http_status=500,
                extra={"data_file": "sessions.json", "field": "sessions",
                       "value_type": type(sessions).__name__})
        # 顺手清理过期会话
        data["sessions"] = {t: s for t, s in sessions.items()
                            if isinstance(s, dict)
                            and s.get("expires_at", "") > now_iso()}
        data["sessions"][token] = {
            "user_id": user["user_id"],
            "username": user["username"],
            "store_id": store_id,
            "password_version": user.get("password_version", 1),
            "created_at": now.isoformat(timespec="seconds"),
            "expires_at": expires.isoformat(timespec="seconds"),
        }
        data["updated_at"] = now_iso()
        self._save(data)
        return token, data["sessions"][token]

    def get(self, token):
        """返回会话字典; 不校验用户状态/密码版本（由调用方结合用户表校验）"""
        if not token:
            return None
        data = self._load()
        sessions = data.get("sessions")
        if not isinstance(sessions, dict):
            raise AuthError(
                "会话数据异常：sessions 不是对象，已保留原文件",
                code=500, http_status=500,
                extra={"data_file": "sessions.json", "field": "sessions",
                       "value_type": type(sessions).__name__})
        session = sessions.get(token)
        if not session:
            return None
        if session.get("expires_at", "") <= now_iso():
            self.delete(token)
            return None
        return session

    def delete(self, token):
        data = self._load()
        sessions = data.get("sessions")
        if not isinstance(sessions, dict):
            raise AuthError(
                "会话数据异常：sessions 不是对象，已保留原文件",
                code=500, http_status=500,
                extra={"data_file": "sessions.json", "field": "sessions",
                       "value_type": type(sessions).__name__})
        if token in sessions:
            del sessions[token]
            data["updated_at"] = now_iso()
            self._save(data)

    def revoke_user(self, user_id, keep_token=None):
        """吊销某用户全部会话（停用账号/重置密码）; keep_token 保留当前会话"""
        data = self._load()
        sessions = data.get("sessions")
        if not isinstance(sessions, dict):
            raise AuthError(
                "会话数据异常：sessions 不是对象，已保留原文件",
                code=500, http_status=500,
                extra={"data_file": "sessions.json", "field": "sessions",
                       "value_type": type(sessions).__name__})
        removed = 0
        for t in list(sessions):
            if sessions[t].get("user_id") == user_id and t != keep_token:
                del sessions[t]
                removed += 1
        if removed:
            data["updated_at"] = now_iso()
            self._save(data)
        return removed


# ==================== 用户与店铺管理 ====================
class UserManager:
    """users.json + stores.json: 账号、角色、状态、用户-店铺授权关系

    - user_id / store_id 创建后永久不变（稳定唯一 ID）
    - 一个用户初始化为一对一店铺关系（store_ids 列表保留多店铺扩展能力）
    - 所有业务数据存放在 data/stores/{store_id}/ 下实现物理隔离
    """

    def __init__(self, data_dir):
        self.users_path = os.path.join(data_dir, "users.json")
        self.stores_path = os.path.join(data_dir, "stores.json")
        self._users_lock = self.users_path + ".lock"
        self._stores_lock = self.stores_path + ".lock"
        os.makedirs(data_dir, exist_ok=True)

    # ---------- 底层读写 ----------
    @staticmethod
    def _read_json(path):
        return _read_json_file(path)

    @staticmethod
    def _atomic_write(path, data):
        _atomic_write_json(path, data)

    def _canonical_collection_locked(self, path, collection_key, id_field, missing):
        """调用方已持有排他锁。结构不一致时备份并迁移，记录一条不丢。"""
        raw = self._read_json(path)
        if raw is None:
            return missing
        document, changed = normalize_collection_document(
            raw, collection_key, id_field, os.path.basename(path))
        if changed:
            _commit_collection_migration(
                path, document, collection_key, id_field,
                _collection_value_type(raw, collection_key))
        return document

    def _read_collection(self, path, lock_path, collection_key, id_field):
        if not os.path.exists(path):
            return None
        with FileLock(lock_path, exclusive=False):
            raw = self._read_json(path)
        if raw is None:
            return None
        document, changed = normalize_collection_document(
            raw, collection_key, id_field, os.path.basename(path))
        if not changed:
            return document
        with FileLock(lock_path):
            return self._canonical_collection_locked(
                path, collection_key, id_field, None)

    def _read_users_doc(self):
        return self._read_collection(
            self.users_path, self._users_lock, "users", "user_id")

    def _read_stores_doc(self):
        return self._read_collection(
            self.stores_path, self._stores_lock, "stores", "store_id")

    def _mutate_users(self, mutator):
        """持锁读-改-写用户表"""
        with FileLock(self._users_lock):
            data = self._canonical_collection_locked(
                self.users_path, "users", "user_id",
                {"version": 1, "updated_at": now_iso(), "users": []})
            result = mutator(data)
            data = result if result is not None else data
            data["updated_at"] = now_iso()
            self._atomic_write(self.users_path, data)
            return data

    def _mutate_stores(self, mutator):
        with FileLock(self._stores_lock):
            data = self._canonical_collection_locked(
                self.stores_path, "stores", "store_id",
                {"version": 1, "updated_at": now_iso(), "stores": []})
            result = mutator(data)
            data = result if result is not None else data
            data["updated_at"] = now_iso()
            self._atomic_write(self.stores_path, data)
            return data

    # ---------- 状态 ----------
    def is_initialized(self):
        data = self._read_users_doc()
        return bool(data and data.get("users"))

    # ---------- 初始化管理员（一次性） ----------
    def initialize_admin(self, username, password, store_name=None):
        """创建首个管理员账号; 管理员是系统管理角色, 不拥有业务店铺

        store_name 参数仅为兼容旧调用保留, 服务端不再据此创建店铺。
        """
        username = validate_username(username)
        validate_password_strength(password)
        if self.is_initialized():
            raise AuthError("系统已初始化，禁止重复执行初始化", code=400)
        user_id = "u_" + secrets.token_hex(4)
        user = {
            "user_id": user_id,
            "username": username,
            "password_hash": hash_password(password),
            "role": "admin",
            "status": "active",
            "store_ids": [],
            "password_version": 1,
            "created_at": now_iso(),
            "last_login_at": None,
        }
        self._mutate_users(lambda d: d["users"].append(user))
        return user, None

    def create_unassigned_store(self, name=None):
        """创建待分配店铺记录（owner 为空, 承接历史迁移数据）"""
        store_id = "s_" + secrets.token_hex(4)
        now = now_iso()
        store = {
            "store_id": store_id,
            "name": name or "待分配数据",
            "owner_user_id": None,
            "assignment_status": "unassigned",
            "status": "active",
            "created_at": now,
            "updated_at": now,
        }
        self._mutate_stores(lambda d: d["stores"].append(store))
        return store

    # ---------- 查询 ----------
    def get_user_by_username(self, username):
        data = self._read_users_doc()
        if not data:
            return None
        for u in data.get("users", []):
            if u.get("username") == username:
                return u
        return None

    def get_user(self, user_id):
        data = self._read_users_doc()
        if not data:
            return None
        for u in data.get("users", []):
            if u.get("user_id") == user_id:
                return u
        return None

    def get_store(self, store_id):
        data = self._read_stores_doc()
        if not data:
            return None
        for s in data.get("stores", []):
            if s.get("store_id") == store_id:
                return s
        return None

    def list_stores(self):
        """全部店铺记录（保持 stores.json 顺序）; 供跨店铺查询/管理使用"""
        data = self._read_stores_doc() or {"stores": []}
        return list(data.get("stores", []))

    def rename_store(self, operator, store_id, new_name,
                     expected_updated_at=None):
        """修改店铺名称: store_id 与数据归属保持不变, 仅更新名称与 updated_at

        并发保护: 若 expected_updated_at 与当前 updated_at 不一致, 拒绝覆盖
        （409 Conflict）, 要求前端刷新后重试。
        """
        if operator.get("role") != "admin":
            raise AuthError("仅管理员可以修改店铺信息", code=403, http_status=403)
        if not isinstance(new_name, str) or not new_name.strip():
            raise AuthError("店铺名称不能为空")
        new_name = new_name.strip()
        if len(new_name) > 64:
            raise AuthError("店铺名称不能超过 64 个字符")
        store = self.get_store(store_id)
        if store is None:
            raise AuthError(f"店铺不存在: {store_id}", code=404, http_status=404)
        if (expected_updated_at is not None
                and store.get("updated_at") != expected_updated_at):
            raise AuthError(
                "店铺信息已被其他操作修改，请刷新后重试",
                code=409, http_status=409)
        old_name = None
        _now = datetime.now().isoformat()

        def mutator(d):
            nonlocal old_name
            for s in d.get("stores", []):
                if s.get("store_id") == store_id:
                    old_name = s.get("name")
                    s["name"] = new_name
                    s["updated_at"] = _now
                    break

        self._mutate_stores(mutator)
        if old_name is None:
            raise AuthError(f"店铺不存在: {store_id}", code=404, http_status=404)
        return old_name, new_name

    def member_stores(self):
        """全部有效成员店铺: owner 存在、active 且 role=member

        待分配/孤儿店铺（owner 为空或指向不存在用户）与 admin 归属店铺一律排除;
        店铺记录缺 assignment_status 字段时默认按 assigned 处理（旧语义）;
        店铺记录缺 status 字段时默认按 active 处理（迁移完成前兼容存量记录）。
        停用/归档店铺从业务口径（页面筛选、聚合统计、导入目标、成员店铺列表）自动移除。
        该方法是"成员店铺"口径的唯一来源, 聚合查询/写入目标/前端下拉均以此为准。
        """
        users_data = self._read_users_doc() or {"users": []}
        users = {u.get("user_id"): u for u in users_data.get("users", [])}
        result = []
        for s in self.list_stores():
            owner = users.get(s.get("owner_user_id"))
            if owner is None or owner.get("role") != "member":
                continue
            if owner.get("status") != "active":
                continue
            if s.get("assignment_status", "assigned") != "assigned":
                continue
            if s.get("status", "active") != "active":
                continue
            result.append(s)
        return result

    def transfer_store_ownership(self, operator, store_id, target_user_id,
                                 expected_updated_at=None):
        """店铺归属变更: owner 改为目标成员, store_id 与业务数据归属不变

        校验链（顺序固定）: 店铺不存在 404 → 并发冲突 409 → 目标用户不存在 404
        → 目标已停用 400 → 目标为 admin 角色 400;
        一对多归属模型: 目标成员已拥有其他店铺时允许追加绑定;
        店铺待分配（owner 为空）或 admin 遗留 → 允许转交;
        幂等: 目标即当前归属成员 → 直接返回当前归属关系, 不写盘不刷新 updated_at。
        原子性: 变更前快照 users/stores 两表, 任一步写入异常用快照回滚两表;
                回滚本身失败时抛出带 rollback_failed 标志的异常。
        """
        if operator.get("role") != "admin":
            raise AuthError("仅管理员可以修改店铺归属", code=403, http_status=403)
        users_snap = self._read_users_doc()
        stores_snap = self._read_stores_doc()
        store = self.get_store(store_id)
        if store is None:
            raise AuthError(f"店铺不存在: {store_id}", code=404, http_status=404)
        if (expected_updated_at is not None
                and store.get("updated_at") != expected_updated_at):
            raise AuthError(
                "店铺信息已被其他操作修改，请刷新后重试",
                code=409, http_status=409)
        target = self.get_user(target_user_id)
        if target is None:
            raise AuthError(f"目标用户不存在: {target_user_id}",
                            code=404, http_status=404)
        if target.get("status") != "active":
            raise AuthError("目标成员已停用，请先启用")
        if target.get("role") != "member":
            raise AuthError("不能将店铺分配给管理员账号")
        prev_owner_id = store.get("owner_user_id")
        if prev_owner_id == target_user_id:
            return {
                "store_id": store_id,
                "store_name": store.get("name"),
                "previous_owner": {"user_id": target.get("user_id"),
                                   "username": target.get("username")},
                "new_owner": {"user_id": target.get("user_id"),
                              "username": target.get("username")},
                "record_counts": {},
                "data_preserved": True,
            }
        prev_owner = self.get_user(prev_owner_id) if prev_owner_id else None
        now = datetime.now().isoformat()

        def mutate_store(d):
            for s in d.get("stores", []):
                if s.get("store_id") == store_id:
                    s["owner_user_id"] = target_user_id
                    s["assignment_status"] = "assigned"
                    s.pop("previous_owner_user_id", None)
                    s["updated_at"] = now
                    break

        def mutate_target(d):
            for u in d.get("users", []):
                if u.get("user_id") == target_user_id:
                    sids = u.get("store_ids") or []
                    if store_id not in sids:
                        sids.append(store_id)
                    u["store_ids"] = sids
                    break

        def mutate_prev(d):
            if not prev_owner_id:
                return
            for u in d.get("users", []):
                if u.get("user_id") == prev_owner_id:
                    u["store_ids"] = [x for x in (u.get("store_ids") or [])
                                      if x != store_id]
                    break

        try:
            self._mutate_stores(mutate_store)
            self._mutate_users(mutate_target)
            self._mutate_users(mutate_prev)
        except Exception:
            rollback_failed = False
            rollback_errors = []
            for label, path, snap in (("stores", self.stores_path, stores_snap),
                                      ("users", self.users_path, users_snap)):
                if snap is not None:
                    try:
                        self._atomic_write(path, snap)
                    except Exception as re:
                        rollback_failed = True
                        rollback_errors.append(f"{label}: {re}")
            if rollback_failed:
                raise AuthError(
                    "变更失败且数据恢复异常，请检查系统日志: "
                    + "; ".join(rollback_errors),
                    code=500, http_status=500,
                    extra={"rollback_failed": True,
                           "rollback_errors": rollback_errors})
            raise AuthError("变更失败，已恢复原始数据", code=500, http_status=500,
                            extra={"rollback_failed": False})
        return {
            "store_id": store_id,
            "store_name": store.get("name"),
            "previous_owner": ({"user_id": prev_owner.get("user_id"),
                                "username": prev_owner.get("username")}
                               if prev_owner else None),
            "new_owner": {"user_id": target.get("user_id"),
                          "username": target.get("username")},
            "record_counts": {},
            "data_preserved": True,
        }

    def create_store(self, operator, store_name, owner_user_id):
        """创建新店铺并绑定归属成员（管理员代建或普通成员自建; 一对多追加绑定）

        校验链（顺序固定）: operator 非 admin/member 403 → 店铺名称 1-64 字符 400
        → member 分支: 强制 owner_user_id=operator 本人（忽略前端提交值）
        → admin 分支: 归属成员必选 400 → 目标用户不存在 404 → 目标已停用 400
        → 目标为 admin 角色 400（admin 永不拥有业务店铺）。
        原子性: 两表内存快照 + 失败补偿回滚; store_id 全局唯一（冲突重试）。
        """
        role = operator.get("role")
        if role not in ("admin", "member"):
            raise AuthError("仅管理员或普通成员可以创建店铺", code=403, http_status=403)
        if not isinstance(store_name, str) or not store_name.strip():
            raise AuthError("请提供店铺名称（1-64 字符）")
        store_name = store_name.strip()
        if len(store_name) > 64:
            raise AuthError("请提供店铺名称（1-64 字符）")
        if role == "member":
            owner_user_id = operator.get("user_id")
            target = operator
        else:
            if not isinstance(owner_user_id, str) or not owner_user_id:
                raise AuthError("请指定归属成员")
            target = self.get_user(owner_user_id)
            if target is None:
                raise AuthError(f"目标用户不存在: {owner_user_id}",
                                code=404, http_status=404)
            if target.get("status") != "active":
                raise AuthError("目标成员已停用，请先启用")
            if target.get("role") != "member":
                raise AuthError("不能将店铺分配给管理员账号")
        users_snap = self._read_users_doc()
        stores_snap = self._read_stores_doc()
        existing = {s.get("store_id")
                    for s in (stores_snap or {}).get("stores", [])}
        store_id = "s_" + secrets.token_hex(4)
        while store_id in existing:
            store_id = "s_" + secrets.token_hex(4)
        now = now_iso()
        store = {
            "store_id": store_id,
            "name": store_name,
            "owner_user_id": owner_user_id,
            "assignment_status": "assigned",
            "status": "active",
            "created_at": now,
            "updated_at": now,
        }

        def mutate_store(d):
            d["stores"].append(store)

        def mutate_owner(d):
            for u in d.get("users", []):
                if u.get("user_id") == owner_user_id:
                    sids = u.get("store_ids") or []
                    if store_id not in sids:
                        sids.append(store_id)
                    u["store_ids"] = sids
                    break

        try:
            self._mutate_stores(mutate_store)
            self._mutate_users(mutate_owner)
        except Exception:
            if users_snap is not None:
                self._atomic_write(self.users_path, users_snap)
            if stores_snap is not None:
                self._atomic_write(self.stores_path, stores_snap)
            raise AuthError("创建失败，请重试", code=500, http_status=500)
        return store

    def set_store_status(self, operator, store_id, action):
        """店铺生命周期状态变更（仅管理员）: disable/enable

        状态机: disable 仅对 active 有效; enable 仅对 disabled 有效;
        幂等: 目标状态已达成 → 直接返回, 不写盘不重复审计。
        归档功能已下线（store_physical_delete）: archive 动作拒绝。
        """
        if operator.get("role") != "admin":
            raise AuthError("仅管理员可以管理店铺状态", code=403, http_status=403)
        if action == "archive":
            raise AuthError("归档功能已下线，不支持该操作")
        if action not in ("disable", "enable"):
            raise AuthError("action 仅支持 disable / enable")
        store = self.get_store(store_id)
        if store is None:
            raise AuthError(f"店铺不存在: {store_id}", code=404, http_status=404)
        current = store.get("status", "active")
        target_status = {"disable": "disabled", "enable": "active"}[action]
        if current == target_status:
            return {"store_id": store_id, "status": current,
                    "previous_status": current,
                    "updated_at": store.get("updated_at"),
                    "store_name": store.get("name"), "idempotent": True}
        valid = ((action == "disable" and current == "active")
                 or (action == "enable" and current == "disabled"))
        if not valid:
            raise AuthError(f"店铺当前状态为 {current}，不支持 {action} 操作")
        stores_snap = self._read_stores_doc()
        now = now_iso()

        def mutator(d):
            for s in d.get("stores", []):
                if s.get("store_id") == store_id:
                    s["status"] = target_status
                    s["updated_at"] = now
                    break

        try:
            self._mutate_stores(mutator)
        except Exception:
            if stores_snap is not None:
                self._atomic_write(self.stores_path, stores_snap)
            raise AuthError("变更失败，请重试", code=500, http_status=500)
        return {"store_id": store_id, "status": target_status,
                "previous_status": current, "updated_at": now,
                "store_name": store.get("name")}

    def migrate_archived_stores(self):
        """存量 archived 店铺幂等迁移为 disabled（归档功能已下线）

        仅当存在 archived 状态店铺时写盘（无则零写入, 幂等）;
        返回迁移清单 [{store_id, name}] 供启动审计使用。
        """
        data = self._read_stores_doc()
        if not data:
            return []
        migrated = [{"store_id": s.get("store_id"), "name": s.get("name")}
                    for s in data.get("stores", [])
                    if s.get("status") == "archived" and s.get("store_id")]
        if not migrated:
            return []
        now = now_iso()

        def mutator(d):
            for s in d.get("stores", []):
                if s.get("status") == "archived":
                    s["status"] = "disabled"
                    s["updated_at"] = now

        self._mutate_stores(mutator)
        return migrated

    def delete_store(self, operator, store_id):
        """删除店铺记录与归属绑定（仅管理员; 业务数据前置校验与目录由调用方处理）

        原子性: 两表快照 + 失败补偿回滚; 返回被删除的店铺记录。
        """
        if operator.get("role") != "admin":
            raise AuthError("仅管理员可以删除店铺", code=403, http_status=403)
        store = self.get_store(store_id)
        if store is None:
            raise AuthError(f"店铺不存在: {store_id}", code=404, http_status=404)
        owner_id = store.get("owner_user_id")
        users_snap = self._read_users_doc()
        stores_snap = self._read_stores_doc()

        def mutate_store(d):
            d["stores"] = [s for s in d.get("stores", [])
                           if s.get("store_id") != store_id]

        def mutate_owner(d):
            if not owner_id:
                return
            for u in d.get("users", []):
                if u.get("user_id") == owner_id:
                    u["store_ids"] = [x for x in (u.get("store_ids") or [])
                                      if x != store_id]
                    break

        try:
            self._mutate_stores(mutate_store)
            self._mutate_users(mutate_owner)
        except Exception:
            if users_snap is not None:
                self._atomic_write(self.users_path, users_snap)
            if stores_snap is not None:
                self._atomic_write(self.stores_path, stores_snap)
            raise AuthError("删除失败，请重试", code=500, http_status=500)
        return store

    def mark_store_unassigned(self, operator, store_id):
        """标记店铺为待分配: owner 置空, store_id 与业务数据完整保留

        幂等: 已是待分配时直接返回当前状态; 原子性同 transfer_store_ownership。
        """
        if operator.get("role") != "admin":
            raise AuthError("仅管理员可以修改店铺归属", code=403, http_status=403)
        store = self.get_store(store_id)
        if store is None:
            raise AuthError(f"店铺不存在: {store_id}", code=404, http_status=404)
        prev_owner_id = store.get("owner_user_id")
        prev_owner = self.get_user(prev_owner_id) if prev_owner_id else None
        if store.get("assignment_status") == "unassigned" and not prev_owner_id:
            return {
                "store_id": store_id,
                "store_name": store.get("name"),
                "assignment_status": "unassigned",
                "previous_owner": ({"user_id": prev_owner.get("user_id"),
                                    "username": prev_owner.get("username")}
                                   if prev_owner else None),
                "record_counts": {},
            }
        users_snap = self._read_users_doc()
        stores_snap = self._read_stores_doc()
        now = now_iso()

        def mutate_store(d):
            for s in d.get("stores", []):
                if s.get("store_id") == store_id:
                    s["owner_user_id"] = None
                    s["assignment_status"] = "unassigned"
                    if prev_owner_id:
                        s["previous_owner_user_id"] = prev_owner_id
                    s["updated_at"] = now
                    break

        def mutate_prev(d):
            for u in d.get("users", []):
                if u.get("user_id") == prev_owner_id:
                    u["store_ids"] = [x for x in (u.get("store_ids") or [])
                                      if x != store_id]
                    break

        try:
            self._mutate_stores(mutate_store)
            if prev_owner_id:
                self._mutate_users(mutate_prev)
        except Exception:
            if users_snap is not None:
                self._atomic_write(self.users_path, users_snap)
            if stores_snap is not None:
                self._atomic_write(self.stores_path, stores_snap)
            raise AuthError("变更失败，请重试", code=500, http_status=500)
        return {
            "store_id": store_id,
            "store_name": store.get("name"),
            "assignment_status": "unassigned",
            "previous_owner": ({"user_id": prev_owner.get("user_id"),
                                "username": prev_owner.get("username")}
                               if prev_owner else None),
            "record_counts": {},
        }

    def list_users(self):
        data = self._read_users_doc() or {"users": []}
        stores = self._read_stores_doc() or {"stores": []}
        store_names = {s["store_id"]: s.get("name") for s in stores.get("stores", [])}
        result = []
        for u in data.get("users", []):
            result.append({
                "user_id": u["user_id"],
                "username": u["username"],
                "role": u.get("role", "member"),
                "status": u.get("status", "active"),
                "store_ids": u.get("store_ids", []),
                "store_names": [store_names.get(sid) for sid in u.get("store_ids", [])],
                "created_at": u.get("created_at"),
                "last_login_at": u.get("last_login_at"),
            })
        return result

    # ---------- 认证 ----------
    def authenticate(self, username, password):
        """校验账号密码; 统一返回 (user|None, reason)"""
        user = self.get_user_by_username(username) if username else None
        if user is None:
            # 对不存在账号执行一次 dummy 校验, 避免时序侧信道泄露账号存在性
            verify_password(password or "",
                            "pbkdf2_sha256$1$" + "0" * 32 + "$" + "0" * 64)
            return None, "invalid"
        if not verify_password(password or "", user.get("password_hash", "")):
            return None, "invalid"
        if user.get("status") != "active":
            return user, "disabled"
        return user, None

    def mark_login(self, user_id):
        def mutator(d):
            for u in d["users"]:
                if u["user_id"] == user_id:
                    u["last_login_at"] = now_iso()
                    break
        self._mutate_users(mutator)

    # ---------- 用户管理（管理员） ----------
    def create_user(self, operator, username, password, role="member",
                    store_name=None, status="active", create_store=True):
        """创建新用户: member 默认绑定一个全新独立店铺; admin 不创建业务店铺

        status: 账号初始状态（active/disabled, 默认 active）——停用创建的
        账号无法登录直至启用。
        create_store: member 是否创建初始店铺（默认 True 保持既有调用方兼容）;
        False 时跳过建店, 用户记录 store_ids=[], 可稍后经店铺管理页分配。
        """
        if operator.get("role") != "admin":
            raise AuthError("仅管理员可以创建用户", code=403, http_status=403)
        username = validate_username(username)
        validate_password_strength(password)
        if role not in ("admin", "member"):
            raise AuthError("角色仅支持 admin / member")
        if status not in ("active", "disabled"):
            raise AuthError("账号状态仅支持 active / disabled")
        if self.get_user_by_username(username) is not None:
            raise AuthError("用户名已存在")
        user_id = "u_" + secrets.token_hex(4)
        user = {
            "user_id": user_id,
            "username": username,
            "password_hash": hash_password(password),
            "role": role,
            "status": status,
            "store_ids": [],
            "password_version": 1,
            "created_at": now_iso(),
            "last_login_at": None,
        }
        if role == "admin":
            self._mutate_users(lambda d: d["users"].append(user))
            return user, None
        if not create_store:
            self._mutate_users(lambda d: d["users"].append(user))
            return user, None
        store_id = "s_" + secrets.token_hex(4)
        user["store_ids"] = [store_id]
        now = now_iso()
        store = {
            "store_id": store_id,
            "name": store_name or f"{username}的店铺",
            "owner_user_id": user_id,
            "assignment_status": "assigned",
            "status": "active",
            "created_at": now,
            "updated_at": now,
        }
        self._mutate_users(lambda d: d["users"].append(user))
        self._mutate_stores(lambda d: d["stores"].append(store))
        return user, store

    def rename_user(self, operator, user_id, new_username):
        """管理员修改用户名（登录账号）: user_id 与其余字段不变

        - 新用户名须满足格式校验且全局唯一; 与原名相同视为未变更幂等成功
        - 会话绑定 user_id: 重命名后既有会话仍有效, 后续以新用户名登录
        - 返回 (old_username, 更新后的用户记录)
        """
        if operator.get("role") != "admin":
            raise AuthError("仅管理员可以修改用户名", code=403, http_status=403)
        new_username = validate_username(new_username)
        target = self.get_user(user_id)
        if target is None:
            raise AuthError(f"用户不存在: {user_id}", code=404, http_status=404)
        old_username = target["username"]
        if old_username == new_username:
            return old_username, target
        if self.get_user_by_username(new_username) is not None:
            raise AuthError("用户名已存在")
        updated = None

        def mutator(d):
            nonlocal updated
            for u in d["users"]:
                if u["user_id"] == user_id:
                    u["username"] = new_username
                    updated = dict(u)
                    break
        self._mutate_users(mutator)
        if updated is None:
            raise AuthError(f"用户不存在: {user_id}", code=404, http_status=404)
        return old_username, updated

    def set_status(self, operator, user_id, action):
        """启用/停用账号; 停用后其会话全部吊销（由调用方执行）"""
        if operator.get("role") != "admin":
            raise AuthError("仅管理员可以管理用户状态", code=403, http_status=403)
        if action not in ("enable", "disable"):
            raise AuthError("action 仅支持 enable / disable")
        target = None

        def mutator(d):
            nonlocal target
            for u in d["users"]:
                if u["user_id"] == user_id:
                    if u["user_id"] == operator["user_id"]:
                        raise AuthError("不能启用或停用自己的账号")
                    u["status"] = "active" if action == "enable" else "disabled"
                    target = u
                    break
        self._mutate_users(mutator)
        if target is None:
            raise AuthError(f"用户不存在: {user_id}", code=404, http_status=404)
        return target

    def reset_password(self, operator, user_id, new_password):
        """管理员重置密码: 提升密码版本并吊销目标用户全部会话（由调用方执行）

        自我保护: 管理员重置自己的密码必须走 change_password（验证原密码）,
        不得经成员管理接口绕过自身密码验证。
        """
        if operator.get("role") != "admin":
            raise AuthError("仅管理员可以重置密码", code=403, http_status=403)
        if user_id == operator["user_id"]:
            raise AuthError("不能通过成员管理接口重置自己的密码，请使用修改密码功能")
        validate_password_strength(new_password)
        target = None

        def mutator(d):
            nonlocal target
            for u in d["users"]:
                if u["user_id"] == user_id:
                    u["password_hash"] = hash_password(new_password)
                    u["password_version"] = u.get("password_version", 1) + 1
                    target = u
                    break
        self._mutate_users(mutator)
        if target is None:
            raise AuthError(f"用户不存在: {user_id}", code=404, http_status=404)
        return target

    def delete_user(self, operator, user_id):
        """删除用户记录（仅用户记录; 店铺记录与业务数据由调用方按场景处理）

        自我保护: 管理员不能删除自己的账号。
        返回被删除的用户记录; 目标不存在时抛 404。
        """
        if operator.get("role") != "admin":
            raise AuthError("仅管理员可以删除用户", code=403, http_status=403)
        if user_id == operator["user_id"]:
            raise AuthError("不能删除自己的账号")
        removed = None

        def mutator(d):
            nonlocal removed
            kept = []
            for u in d.get("users", []):
                if u.get("user_id") == user_id:
                    removed = u
                else:
                    kept.append(u)
            d["users"] = kept
        self._mutate_users(mutator)
        if removed is None:
            raise AuthError(f"用户不存在: {user_id}", code=404, http_status=404)
        return removed

    def change_password(self, user_id, old_password, new_password,
                        confirm_password=None):
        """本人修改密码: 校验旧密码与两次新密码一致性, 提升密码版本（其他会话将失效）

        confirm_password 为必填的服务端校验项: 缺失或不一致均拒绝,
        防止绕过前端直接调 API 提交仅一次的新密码。
        """
        user = self.get_user(user_id)
        if user is None:
            raise AuthError("用户不存在", code=404, http_status=404)
        if not verify_password(old_password or "", user.get("password_hash", "")):
            raise AuthError("原密码不正确")
        if not isinstance(confirm_password, str):
            raise AuthError("请再次输入新密码")
        if new_password != confirm_password:
            raise AuthError("两次输入的新密码不一致")
        validate_password_strength(new_password)
        new_hash = hash_password(new_password)

        def mutator(d):
            for u in d["users"]:
                if u["user_id"] == user_id:
                    u["password_hash"] = new_hash
                    u["password_version"] = u.get("password_version", 1) + 1
                    break
        self._mutate_users(mutator)
        return self.get_user(user_id)

    def rollback_user_store(self, user_id, store_id):
        """按精确 ID 删除用户与店铺记录（幂等）; 仅供初始化迁移失败回滚调用

        复用既有持锁读-改-写与原子写入机制, 不触碰其他用户/店铺记录。
        返回被删除的记录数（用户记录数 + 店铺记录数, 0~2）。
        """
        removed = 0

        def remove_user(d):
            nonlocal removed
            before = len(d.get("users", []))
            d["users"] = [u for u in d.get("users", [])
                          if u.get("user_id") != user_id]
            if len(d["users"]) != before:
                removed += 1

        def remove_store(d):
            nonlocal removed
            before = len(d.get("stores", []))
            d["stores"] = [s for s in d.get("stores", [])
                           if s.get("store_id") != store_id]
            if len(d["stores"]) != before:
                removed += 1

        self._mutate_users(remove_user)
        self._mutate_stores(remove_store)
        return removed

    # ---------- 自主注册（公开注册接口） ----------
    def register_user(self, username, password, confirm_password, store_name):
        """公开自主注册: 创建普通用户（恒为 member）及其默认店铺

        与 create_user（管理员创建）的差异:
        - 签名仅四个值参数, 无 operator/role——权限字段在类型签名层面即无读取路径;
        - 用户名唯一性在 _mutate_users 持锁 mutator 内部复检, 消除并发注册竞态窗口;
        - 密码哈希锁外预计算, 避免持锁 PBKDF2 阻塞其他认证写操作。
        """
        username = validate_username(username)
        validate_password_strength(password)
        if password != confirm_password:
            raise AuthError("两次输入的密码不一致")
        if not isinstance(store_name, str) or not store_name.strip():
            raise AuthError("请输入店铺名称")
        store_name = store_name.strip()
        if len(store_name) > 64:
            raise AuthError("店铺名称不能超过 64 个字符")

        password_hash = hash_password(password)
        user_id = "u_" + secrets.token_hex(4)
        store_id = "s_" + secrets.token_hex(4)
        user = {
            "user_id": user_id,
            "username": username,
            "password_hash": password_hash,
            "role": "member",
            "status": "active",
            "store_ids": [store_id],
            "password_version": 1,
            "created_at": now_iso(),
            "last_login_at": None,
        }
        now = now_iso()
        store = {
            "store_id": store_id,
            "name": store_name,
            "owner_user_id": user_id,
            "assignment_status": "assigned",
            "status": "active",
            "created_at": now,
            "updated_at": now,
        }

        def add_user(d):
            for u in d.get("users", []):
                if u.get("username") == username:
                    raise AuthError("用户名已存在")
            d["users"].append(user)

        self._mutate_users(add_user)
        self._mutate_stores(lambda d: d["stores"].append(store))
        return user, store

    # ---------- 授权校验 ----------
    @staticmethod
    def check_store_authorized(user, store_id):
        """校验店铺在用户授权范围内; 防止伪造 store_id"""
        if store_id not in (user.get("store_ids") or []):
            raise AuthError("无权访问该店铺数据", code=403, http_status=403)
        return True

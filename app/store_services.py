"""按店铺隔离的服务工厂与旧数据迁移

隔离设计:
1. 每个店铺独立数据目录 data/stores/{store_id}/, 内含全部业务 JSON 文件、
   backups/、logs/（操作日志）、uploads/ —— 物理目录级完全隔离;
2. StoreServices 为单个店铺组装全套服务（repo/导入/查询/对账/利润/广告费）,
   StoreServiceRegistry 按 store_id 惰性创建并缓存（线程安全）;
3. 旧版单店铺数据迁移: 先整体备份到 data/pre_migration_backup_{ts}/,
   再复制到目标店铺目录并校验, 全部成功后才移除原文件; 任何一步失败自动回滚。
"""
import json
import os
import shutil
import threading
from datetime import datetime

from . import config
from .ad_expense_service import AdExpenseService
from .data_schemas import BUSINESS_FILES, now_iso
from .import_service import ImportService
from .json_repository import JsonRepository
from .list_queries import AnomalyQueryService, SupplierCostQueryService
from .operation_log import OperationLog
from .order_queries import OrderQueryService
from .profit_service import ProfitService
from .reconciliation_service import ReconciliationService


class StoreServices:
    """单个店铺的全套服务绑定; 所有组件仅读写本店铺数据目录"""

    def __init__(self, store_id, stores_root):
        self.store_id = store_id
        self.store_dir = os.path.join(stores_root, store_id)
        self.data_dir = os.path.join(self.store_dir, "data")
        self.backup_dir = os.path.join(self.store_dir, "backups")
        self.log_dir = os.path.join(self.store_dir, "logs")
        self.upload_dir = os.path.join(self.store_dir, "uploads")
        os.makedirs(self.data_dir, exist_ok=True)
        os.makedirs(self.backup_dir, exist_ok=True)
        os.makedirs(self.log_dir, exist_ok=True)
        os.makedirs(self.upload_dir, exist_ok=True)
        self.repo = JsonRepository(data_dir=self.data_dir,
                                   backup_dir=self.backup_dir)
        self.op_log = OperationLog(log_dir=self.log_dir)
        self.importer = ImportService(self.repo, self.op_log,
                                      upload_dir=self.upload_dir)
        self.queries = OrderQueryService(self.repo, self.op_log)
        self.reconciler = ReconciliationService(self.repo, self.op_log)
        self.profits = ProfitService(self.repo, self.op_log)
        self.supplier_queries = SupplierCostQueryService(self.repo, self.op_log)
        self.anomaly_queries = AnomalyQueryService(self.repo, self.op_log)
        self.ad_expenses = AdExpenseService(self.repo, self.op_log)

    def ensure_all(self):
        return self.repo.ensure_all()


class StoreServiceRegistry:
    """store_id -> StoreServices 的线程安全缓存"""

    def __init__(self, stores_root):
        self.stores_root = stores_root
        self._cache = {}
        self._lock = threading.Lock()

    def get(self, store_id):
        if not isinstance(store_id, str) or not store_id:
            raise ValueError("store_id 无效")
        with self._lock:
            bundle = self._cache.get(store_id)
            if bundle is None:
                bundle = StoreServices(store_id, self.stores_root)
                bundle.ensure_all()
                self._cache[store_id] = bundle
            return bundle

    def evict(self, store_id):
        """从缓存中驱逐指定店铺的服务实例（数据清理/删除账号后调用）

        清理或删除店铺后, 缓存中的旧 bundle 仍持有清理前的内存态与文件句柄,
        不驱逐会导致总览/利润等页面继续读到已清理的旧数据。
        驱逐后再次 get() 会基于磁盘最新状态重建全套服务。
        """
        with self._lock:
            return self._cache.pop(store_id, None)


# ==================== 旧数据迁移 ====================
def _legacy_files(data_root):
    """数据根目录下遗留的旧业务文件（不含用户/会话/店铺等全局文件）"""
    result = []
    for name, spec in BUSINESS_FILES.items():
        path = os.path.join(data_root, spec["file"])
        if os.path.exists(path):
            result.append((name, spec["file"], path))
    return result


def migrate_legacy_store_data(data_root, stores_root, store_id,
                              security_log=None):
    """将单店铺时代的 data/*.json 迁移到 data/stores/{store_id}/data/

    步骤: 整体备份 → 逐文件复制并校验 → 迁移 uploads → 移除原文件。
    任何一步失败: 回滚（删除已复制文件, 从备份恢复原位）, 抛出异常。
    返回迁移报告 dict。
    """
    legacy = _legacy_files(data_root)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    backup_dir = os.path.join(data_root, f"pre_migration_backup_{ts}")
    store_dir = os.path.join(stores_root, store_id)
    store_data_dir = os.path.join(store_dir, "data")
    store_upload_dir = os.path.join(store_dir, "uploads")

    def log(event, **kw):
        if security_log:
            security_log.append(event, detail={"store_id": store_id, **kw})

    if not legacy:
        return {"migrated": [], "backup_dir": None, "uploads": 0,
                "skipped_reason": "no_legacy_data"}

    # 1) 整体备份（复制, 不移动原文件）
    os.makedirs(backup_dir, exist_ok=True)
    copied = []
    try:
        for name, filename, path in legacy:
            shutil.copy2(path, os.path.join(backup_dir, filename))
            copied.append(filename)

        # 2) 复制到店铺目录并校验可解析
        os.makedirs(store_data_dir, exist_ok=True)
        for name, filename, path in legacy:
            dst = os.path.join(store_data_dir, filename)
            shutil.copy2(path, dst)
            with open(dst, "r", encoding="utf-8") as f:
                json.load(f)  # 校验 JSON 完整可读

        # 3) 迁移 uploads 目录（历史导入源文件）
        uploads_migrated = 0
        old_uploads = os.path.join(data_root, os.pardir, "uploads")
        old_uploads = os.path.normpath(old_uploads)
        legacy_uploads = config.UPLOAD_DIR if os.path.abspath(config.UPLOAD_DIR) == old_uploads else old_uploads
        if os.path.isdir(legacy_uploads):
            os.makedirs(store_upload_dir, exist_ok=True)
            for fn in os.listdir(legacy_uploads):
                src = os.path.join(legacy_uploads, fn)
                if os.path.isfile(src):
                    shutil.copy2(src, os.path.join(store_upload_dir, fn))
                    uploads_migrated += 1

        # 4) 全部校验通过后移除原文件（json + 同名 .lock）
        for name, filename, path in legacy:
            os.remove(path)
            lock = path + ".lock"
            if os.path.exists(lock):
                try:
                    os.remove(lock)
                except OSError:
                    pass

        # 迁移成功后移除旧 uploads 中已复制的文件（保留目录本身）
        if uploads_migrated and os.path.isdir(legacy_uploads):
            for fn in os.listdir(legacy_uploads):
                p = os.path.join(legacy_uploads, fn)
                if os.path.isfile(p) and os.path.exists(
                        os.path.join(store_upload_dir, fn)):
                    try:
                        os.remove(p)
                    except OSError:
                        pass

        report = {"migrated": [f for _, f, _ in legacy],
                  "backup_dir": backup_dir,
                  "uploads": uploads_migrated,
                  "migrated_at": now_iso()}
        log("migration.success", files=report["migrated"],
            backup_dir=backup_dir)
        return report
    except BaseException as e:
        # 回滚: 清理店铺目录中已复制文件, 从备份恢复原位
        errors = []
        try:
            for name, filename, path in legacy:
                dst = os.path.join(store_data_dir, filename)
                if os.path.exists(dst):
                    os.remove(dst)
                if os.path.exists(path):
                    continue  # 原文件还在（尚未删除阶段）
                bpath = os.path.join(backup_dir, filename)
                if os.path.exists(bpath):
                    shutil.copy2(bpath, path)
        except BaseException as rollback_err:
            errors.append(f"回滚异常: {rollback_err!r}")
        log("migration.failed", error=repr(e), rollback_errors=errors)
        raise RuntimeError(
            f"旧数据迁移失败, 已回滚到迁移前状态: {e}; 备份目录: {backup_dir}"
        ) from e

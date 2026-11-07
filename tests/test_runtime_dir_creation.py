"""运行时目录创建测试。

验证 ReconcileServer.__init__ 后 data/、data/stores/、backups/、logs/、uploads/
五个目录均被自动创建。通过 data_dir/backup_dir 参数注入临时目录，并 mock
config.LOG_DIR / UPLOAD_DIR 指向临时目录下的子目录，避免污染真实项目目录。
"""
import os
import tempfile
import unittest
from unittest.mock import patch

from app.server import ReconcileServer


class TestRuntimeDirCreation(unittest.TestCase):

    def setUp(self):
        self._tmp = tempfile.mkdtemp(prefix="wxs_test_dirs_")

    def tearDown(self):
        import shutil
        shutil.rmtree(self._tmp, ignore_errors=True)

    def _make_server(self):
        data_dir = os.path.join(self._tmp, "data")
        backup_dir = os.path.join(self._tmp, "backups")
        log_dir = os.path.join(self._tmp, "logs")
        upload_dir = os.path.join(self._tmp, "uploads")
        with patch("app.server.config.LOG_DIR", log_dir), \
             patch("app.server.config.UPLOAD_DIR", upload_dir):
            return ReconcileServer(
                host="127.0.0.1", port=19999,
                data_dir=data_dir, backup_dir=backup_dir)

    def test_data_dir_created(self):
        srv = self._make_server()
        self.assertTrue(os.path.isdir(srv.data_root))

    def test_stores_dir_created(self):
        srv = self._make_server()
        self.assertTrue(os.path.isdir(srv.stores_root))

    def test_backup_dir_created(self):
        srv = self._make_server()
        self.assertTrue(os.path.isdir(srv.backup_dir))

    def test_log_dir_created(self):
        log_dir = os.path.join(self._tmp, "logs")
        with patch("app.server.config.LOG_DIR", log_dir), \
             patch("app.server.config.UPLOAD_DIR",
                   os.path.join(self._tmp, "uploads")):
            ReconcileServer(
                host="127.0.0.1", port=19998,
                data_dir=os.path.join(self._tmp, "data"),
                backup_dir=os.path.join(self._tmp, "backups"))
        self.assertTrue(os.path.isdir(log_dir))

    def test_upload_dir_created(self):
        upload_dir = os.path.join(self._tmp, "uploads")
        with patch("app.server.config.LOG_DIR",
                   os.path.join(self._tmp, "logs")), \
             patch("app.server.config.UPLOAD_DIR", upload_dir):
            ReconcileServer(
                host="127.0.0.1", port=19997,
                data_dir=os.path.join(self._tmp, "data"),
                backup_dir=os.path.join(self._tmp, "backups"))
        self.assertTrue(os.path.isdir(upload_dir))

    def test_all_five_dirs_exist_after_init(self):
        srv = self._make_server()
        expected = [
            srv.data_root,
            srv.stores_root,
            srv.backup_dir,
            os.path.join(self._tmp, "logs"),
            os.path.join(self._tmp, "uploads"),
        ]
        for d in expected:
            self.assertTrue(os.path.isdir(d), f"目录未创建: {d}")


if __name__ == "__main__":
    unittest.main()

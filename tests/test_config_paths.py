"""config.py 路径适配回归测试。

验证开发环境下 config 六个路径常量取值与改造前一致，
且 STATIC_DIR 与 DATA_DIR 在开发环境共用同一基准。
"""
import os
import unittest
from unittest.mock import patch

from app import config

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class TestConfigPathsDevEnv(unittest.TestCase):
    """开发环境 config 路径常量。"""

    def test_base_dir_is_project_root(self):
        self.assertEqual(config.BASE_DIR, PROJECT_ROOT)

    def test_data_dir(self):
        self.assertEqual(config.DATA_DIR, os.path.join(PROJECT_ROOT, "data"))

    def test_backup_dir(self):
        self.assertEqual(config.BACKUP_DIR, os.path.join(PROJECT_ROOT, "backups"))

    def test_log_dir(self):
        self.assertEqual(config.LOG_DIR, os.path.join(PROJECT_ROOT, "logs"))

    def test_upload_dir(self):
        self.assertEqual(config.UPLOAD_DIR, os.path.join(PROJECT_ROOT, "uploads"))

    def test_static_dir(self):
        self.assertEqual(config.STATIC_DIR, os.path.join(PROJECT_ROOT, "static"))

    def test_static_and_data_share_same_base_in_dev(self):
        self.assertEqual(
            os.path.dirname(config.STATIC_DIR),
            os.path.dirname(config.DATA_DIR))


if __name__ == "__main__":
    unittest.main()

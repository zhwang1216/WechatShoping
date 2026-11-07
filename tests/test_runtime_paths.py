"""runtime_paths 模块单元测试。

覆盖：
- 开发环境：runtime_root == bundle_root == 项目根目录
- 打包环境（mock sys.frozen）：runtime_root 指向 exe 同级、bundle_root 指向 _MEIPASS
- RuntimePaths 为不可变对象（frozen=True）
"""
import os
import sys
import unittest
from unittest.mock import patch

from app.runtime_paths import RuntimePaths, resolve_paths

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class TestRuntimePathsDevEnv(unittest.TestCase):
    """开发环境路径解析。"""

    def test_dev_env_returns_project_root(self):
        paths = resolve_paths()
        self.assertEqual(paths.runtime_root, PROJECT_ROOT)
        self.assertEqual(paths.bundle_root, PROJECT_ROOT)

    def test_dev_env_runtime_equals_bundle(self):
        paths = resolve_paths()
        self.assertEqual(paths.runtime_root, paths.bundle_root)


class TestRuntimePathsFrozenEnv(unittest.TestCase):
    """打包环境（PyInstaller）路径解析。"""

    @patch.object(sys, "frozen", True, create=True)
    @patch.object(sys, "executable", "/opt/app/微信小店对账系统.exe")
    @patch.object(sys, "_MEIPASS", "/tmp/_MEI12345", create=True)
    def test_frozen_env_separates_paths(self):
        paths = resolve_paths()
        self.assertEqual(paths.runtime_root, "/opt/app")
        self.assertEqual(paths.bundle_root, "/tmp/_MEI12345")
        self.assertNotEqual(paths.runtime_root, paths.bundle_root)

    @patch.object(sys, "frozen", True, create=True)
    @patch.object(sys, "executable", "/opt/app/app.exe")
    def test_frozen_env_no_meipass_fallback(self):
        with patch.object(sys, "_MEIPASS", None, create=True):
            paths = resolve_paths()
            self.assertEqual(paths.runtime_root, "/opt/app")
            self.assertEqual(paths.bundle_root, "/opt/app")


class TestRuntimePathsImmutable(unittest.TestCase):
    """RuntimePaths 不可变值对象。"""

    def test_frozen_dataclass_cannot_mutate(self):
        paths = resolve_paths()
        with self.assertRaises(Exception):
            paths.runtime_root = "/changed"


if __name__ == "__main__":
    unittest.main()

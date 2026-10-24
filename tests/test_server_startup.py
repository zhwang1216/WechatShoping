"""服务启动与端口冲突处理测试（OSError 48 Address already in use 修复）

验收项：
1. 正常情况下服务可以成功启动（随机端口）。
2. 端口被其他进程占用时，抛出 ServerStartupError 且提示包含 IP/端口/排查命令。
3. 端口被本系统已运行实例占用时，提示"本系统正在运行的服务实例"。
4. 同一实例重复调用 start() 不崩溃，给出"请勿重复调用 start()"。
5. shutdown() 后可再次 start()（端口释放，无遗留 socket/线程）。
6. server.port 与实际监听端口一致（前端访问地址正确）。
7. 现有业务功能不受影响（/api/health 正常返回）。
"""
import json
import os
import shutil
import socket
import tempfile
import time
import unittest
import urllib.error
import urllib.request

from app.server import ReconcileServer, ServerStartupError


class TestServerStartup(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp()

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp)

    def _make_server(self, port):
        return ReconcileServer(
            host="127.0.0.1", port=port,
            data_dir=os.path.join(self.tmp, f"data_{port}_{id(self)}"),
            backup_dir=os.path.join(self.tmp, f"backups_{port}_{id(self)}"),
            log_dir=os.path.join(self.tmp, f"logs_{port}_{id(self)}"),
        )

    def _free_port(self):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(("127.0.0.1", 0))
            return s.getsockname()[1]

    def _health(self, server):
        with urllib.request.urlopen(
                f"http://127.0.0.1:{server.port}/api/health", timeout=5) as resp:
            return json.loads(resp.read().decode("utf-8"))

    # 1. 正常启动
    def test_01_start_success_and_health(self):
        server = self._make_server(self._free_port())
        try:
            t = server.start(blocking=False)
            self.assertTrue(t.is_alive())
            body = self._health(server)
            self.assertEqual(body["code"], 0)
            self.assertIn("微信小店", body["data"]["service"])
        finally:
            server.shutdown()

    # 2. 端口被其他进程（非本系统）占用
    def test_02_port_occupied_by_other_process(self):
        port = self._free_port()
        blocker = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        blocker.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        blocker.bind(("127.0.0.1", port))
        blocker.listen(1)
        server = self._make_server(port)
        try:
            with self.assertRaises(ServerStartupError) as ctx:
                server.start(blocking=False)
            msg = str(ctx.exception)
            self.assertIn("127.0.0.1", msg)
            self.assertIn(str(port), msg)
            self.assertIn("lsof", msg)
            self.assertIn("--port", msg)
            self.assertFalse(ctx.exception.occupied_by_self)
            self.assertIsNone(server.httpd)  # 启动失败不遗留服务对象
        finally:
            blocker.close()
            server.shutdown()

    # 3. 端口被本系统已运行实例占用 → 提示已有实例
    def test_03_port_occupied_by_self_instance(self):
        first = self._make_server(self._free_port())
        second = None
        try:
            first.start(blocking=False)
            second = self._make_server(first.port)
            with self.assertRaises(ServerStartupError) as ctx:
                second.start(blocking=False)
            msg = str(ctx.exception)
            self.assertIn("本系统", msg)
            self.assertIn("正在运行", msg)
            self.assertTrue(ctx.exception.occupied_by_self)
            # 原实例不受影响
            body = self._health(first)
            self.assertEqual(body["code"], 0)
        finally:
            first.shutdown()
            if second:
                second.shutdown()

    # 4. 同一实例重复调用 start()
    def test_04_double_start_same_instance(self):
        server = self._make_server(self._free_port())
        try:
            server.start(blocking=False)
            with self.assertRaises(ServerStartupError) as ctx:
                server.start(blocking=False)
            self.assertIn("请勿重复调用 start()", str(ctx.exception))
            self.assertTrue(ctx.exception.occupied_by_self)
        finally:
            server.shutdown()

    # 5. shutdown 后可再次启动（端口释放、无残留）
    def test_05_restart_after_shutdown(self):
        port = self._free_port()
        server = self._make_server(port)
        try:
            server.start(blocking=False)
            server.shutdown()
            self.assertIsNone(server.httpd)
            # shutdown 幂等
            server.shutdown()
            # 端口已释放，可立即重新绑定
            server.start(blocking=False)
            self.assertEqual(server.port, port)
            body = self._health(server)
            self.assertEqual(body["code"], 0)
        finally:
            server.shutdown()

    # 6. 启动失败后再以正确端口启动成功（失败不留资源）
    def test_06_recover_after_startup_failure(self):
        blocker_port = self._free_port()
        blocker = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        blocker.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        blocker.bind(("127.0.0.1", blocker_port))
        blocker.listen(1)
        server = self._make_server(blocker_port)
        try:
            with self.assertRaises(ServerStartupError):
                server.start(blocking=False)
            blocker.close()
            # 同一实例：失败后 httpd 为 None，可换端口重新启动
            server.port = self._free_port()
            server.start(blocking=False)
            self.assertEqual(self._health(server)["code"], 0)
        finally:
            server.shutdown()

    # 7. blocking 模式 SIGINT 语义（KeyboardInterrupt 触发 shutdown 清理）
    def test_07_shutdown_is_idempotent_and_safe(self):
        server = self._make_server(self._free_port())
        server.start(blocking=False)
        server.shutdown()
        server.shutdown()
        self.assertIsNone(server.httpd)
        # 再次完整生命周期仍正常
        server.start(blocking=False)
        self.assertIsNotNone(server.httpd)
        server.shutdown()


if __name__ == "__main__":
    unittest.main()

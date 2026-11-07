"""数据导入流程调整与弹窗内新建店铺 验收测试

阶段1: member 建店权限矩阵（TASK-103）
阶段2: precheck 契约测试（TASK-203）
阶段6: 17 项验收标准 VC-1~VC-17（TASK-603，后续补充）

复用 tests/test_import_store_picker.py 的 TestEnv/Client 基础设施。
"""
import io
import json
import os
import socket
import unittest
from unittest.mock import MagicMock, patch

import xlwt

from tests.http_client import ADMIN_USER, TestEnv
from tests.test_auth_isolation import make_wx_xlsx
from tests.test_delivery_e2e import WX_HEADERS, WX_ROWS


def make_wx_xls(prefix="E2E"):
    """用 xlwt 生成与 make_wx_xlsx 同构数据的 .xls (BIFF8) 文件，用于验证 .xls 导入支持"""
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("订单")
    for c, val in enumerate(WX_HEADERS):
        ws.write(0, c, val)
    for r, row in enumerate(WX_ROWS, start=1):
        rr = list(row)
        if isinstance(rr[0], str) and rr[0].startswith("E2E"):
            rr[0] = prefix + rr[0][3:]
        for c, val in enumerate(rr):
            if val is None:
                continue
            ws.write(r, c, val)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestImportFlowRevamp(unittest.TestCase):
    """数据导入流程调整与弹窗内新建店铺 验收测试"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.member
        st, r = cls.member.get("/api/auth/me")
        cls.member_user_id = r["data"]["user_id"]
        cls.member_store = r["data"].get("store_id")
        cls.other = cls.env.create_user_client(
            "flow_other", "Other123456", store_name="其他铺")
        cls.other_user_id = cls.env.get_user_id("flow_other")
        cls.admin_user_id = cls.env.get_user_id(ADMIN_USER)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 阶段1: member 建店权限矩阵 ----------

    def test_member_create_own_store(self):
        st, r = self.member.post("/api/stores", {"name": "我的新店铺"})
        self.assertEqual(st, 200, r)
        data = r["data"]
        self.assertEqual(data["owner_user_id"], self.member_user_id)
        self.assertEqual(data["status"], "active")

    def test_member_create_with_other_owner_ignored(self):
        st, r = self.member.post("/api/stores", {
            "name": "强制归属店铺",
            "owner_user_id": self.other_user_id,
        })
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["owner_user_id"], self.member_user_id)

    def test_admin_create_self_owner_rejected(self):
        st, r = self.admin.post("/api/stores", {
            "name": "管理员店铺",
            "owner_user_id": self.admin_user_id,
        })
        self.assertEqual(st, 400, r)
        self.assertIn("管理员", r.get("message", ""))

    def test_admin_create_for_member_ok(self):
        st, r = self.admin.post("/api/stores", {
            "name": "代建店铺",
            "owner_user_id": self.member_user_id,
        })
        self.assertEqual(st, 200, r)
        self.assertEqual(r["data"]["owner_user_id"], self.member_user_id)

    # ---------- 阶段2: precheck 契约测试 ----------

    def _precheck(self, client, source_type="wechat_orders",
                  filename="test.xlsx", content=None):
        if content is None:
            content = make_wx_xlsx("PC")
        return client.upload(
            f"/api/import/precheck?source_type={source_type}",
            filename, content)

    def test_precheck_no_store_id_required(self):
        st, r = self._precheck(self.member)
        self.assertEqual(st, 200, r)
        data = r["data"]
        for key in ("file_name", "file_size", "file_hash",
                    "source_type", "sheets", "readable"):
            self.assertIn(key, data, key)
        self.assertTrue(data["readable"])
        self.assertEqual(data["source_type"], "wechat_orders")
        self.assertIsInstance(data["sheets"], list)
        self.assertGreater(len(data["sheets"]), 0)

    def test_precheck_no_business_data_write(self):
        store_id = self.member_store
        data_dir = os.path.join(self.env.data_dir, "stores",
                                store_id, "data")
        before = self._snapshot(data_dir)
        st, r = self._precheck(self.member)
        self.assertEqual(st, 200, r)
        after = self._snapshot(data_dir)
        self.assertEqual(before, after, "precheck 不应写入业务数据")

    @staticmethod
    def _snapshot(directory):
        snap = {}
        if not os.path.isdir(directory):
            return snap
        for root, _dirs, files in os.walk(directory):
            for name in files:
                fp = os.path.join(root, name)
                try:
                    with open(fp, "rb") as f:
                        snap[os.path.relpath(fp, directory)] = f.read()
                except OSError:
                    snap[os.path.relpath(fp, directory)] = None
        return snap

    def test_precheck_invalid_xlsx(self):
        st, r = self._precheck(self.member, filename="test.csv",
                               content=b"a,b,c\n1,2,3\n")
        self.assertEqual(st, 400, r)
        self.assertIn(".xlsx", r.get("message", ""))

    def test_precheck_oversize(self):
        port = self.env.server.port
        cookie = self.member.cookie
        req = (
            "POST /api/import/precheck?source_type=wechat_orders "
            "HTTP/1.1\r\n"
            f"Host: 127.0.0.1:{port}\r\n"
            "Content-Length: 60000000\r\n"
            "X-Filename: big.xlsx\r\n"
            f"Cookie: wxr_session={cookie}\r\n"
            "X-Requested-With: XMLHttpRequest\r\n"
            "Connection: close\r\n\r\n"
        ).encode("utf-8")
        with socket.create_connection(("127.0.0.1", port),
                                       timeout=10) as s:
            s.sendall(req)
            chunks = []
            while True:
                chunk = s.recv(4096)
                if not chunk:
                    break
                chunks.append(chunk)
        raw = b"".join(chunks)
        status_line = raw[:raw.find(b"\r\n")].decode("utf-8")
        status = int(status_line.split()[1])
        body = raw[raw.find(b"\r\n\r\n") + 4:]
        try:
            r = json.loads(body.decode("utf-8"))
        except Exception:
            r = {}
        self.assertEqual(status, 413, r)
        self.assertIn("大小限制", r.get("message", ""))

    def test_precheck_corrupted(self):
        st, r = self._precheck(self.member, filename="bad.xlsx",
                               content=b"PK\x00\x01not a real xlsx")
        self.assertEqual(st, 400, r)

    def test_precheck_empty_sheets(self):
        wb_mock = MagicMock()
        wb_mock.sheetnames = []
        with patch("app.import_service.load_workbook_from_bytes",
                   return_value=wb_mock):
            st, r = self._precheck(self.member, filename="empty.xlsx",
                                   content=b"PK" + b"\x00" * 100)
        self.assertEqual(st, 400, r)
        self.assertIn("工作表", r.get("message", ""))

    # ---------- .xls 格式支持 ----------

    def test_precheck_xls_accepted(self):
        """上传 .xls 文件，precheck 应返回 200 且识别工作表"""
        st, r = self._precheck(self.member, filename="test.xls",
                               content=make_wx_xls("XLS"))
        self.assertEqual(st, 200, r)
        data = r["data"]
        self.assertTrue(data["readable"], r)
        self.assertGreater(len(data["sheets"]), 0, r)
        self.assertEqual(data["sheets"][0]["name"], "订单", r)

    def test_precheck_xls_invalid_magic(self):
        """扩展名 .xls 但内容不是 OLE2 魔数，应返回 400"""
        st, r = self._precheck(self.member, filename="fake.xls",
                               content=b"NOT_AN_EXCEL_FILE" + b"\x00" * 50)
        self.assertEqual(st, 400, r)

    def test_precheck_xls_corrupted_ole2(self):
        """OLE2 魔数但内容损坏，xlrd 解析失败应返回 400"""
        st, r = self._precheck(self.member, filename="bad.xls",
                               content=b"\xd0\xcf\x11\xe0" + b"\x00" * 100)
        self.assertEqual(st, 400, r)


class TestFrontendDomContract(unittest.TestCase):
    """前端 DOM 结构契约自动化断言（spec 4.4.7，design 2.7.5.3.B）

    通过 HTTP GET 拉取 index.html / style.css 源码，断言包含关键元素字符串，
    作为前端结构契约的自动化守护（结构断言而非行为断言）。
    """

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _get_static(self, path):
        st, body, _ = self.admin.get_page("/static/" + path)
        self.assertEqual(st, 200, f"GET /static/{path} 失败: status={st}")
        return body

    def _get_page(self, path):
        st, body, _ = self.admin.get_page(path)
        self.assertEqual(st, 200, f"GET {path} 失败: status={st}")
        return body

    def test_01_index_html_contains_file_info_bar_elements(self):
        html = self._get_page("/data-import")
        for token in ('id="file-info-bar"', 'class="file-info-bar"',
                      'class="fi-header"', 'class="fi-title"',
                      'class="fi-actions"', 'id="fi-meta"', 'id="fi-sheets"'):
            self.assertIn(token, html, f"index.html 缺少元素: {token}")

    def test_02_index_html_contains_handle_file_selected(self):
        html = self._get_page("/data-import")
        self.assertIn("handleFileSelected", html, "index.html 缺少 handleFileSelected 函数")

    def test_03_index_html_contains_open_store_picker_with_file_info(self):
        html = self._get_page("/data-import")
        self.assertIn("openStorePicker", html, "index.html 缺少 openStorePicker")
        self.assertIn("fileInfo", html, "index.html 缺少 fileInfo 参数")

    def test_04_index_html_contains_create_store_entry(self):
        html = self._get_page("/data-import")
        self.assertIn("新建店铺", html, "index.html 缺少新建店铺入口")
        self.assertIn("openCreateStoreForm", html, "index.html 缺少 openCreateStoreForm")

    def test_05_index_html_contains_confirm_select_button(self):
        html = self._get_page("/data-import")
        self.assertIn("确认选择", html, "index.html 缺少确认选择按钮")

    def test_06_index_html_contains_back_to_store_button(self):
        html = self._get_page("/data-import")
        self.assertIn("返回修改店铺", html, "index.html 缺少返回修改店铺按钮")

    def test_07_index_html_no_inline_margin_on_file_info_bar(self):
        import re
        html = self._get_page("/data-import")
        matches = re.findall(r'file-info-bar[^>]*style="[^"]*margin-bottom', html)
        self.assertEqual(matches, [], "file-info-bar 元素仍含内联 margin-bottom 样式")

    def test_08_style_css_contains_file_info_bar_definitions(self):
        css = self._get_static("style.css")
        for token in (".file-info-bar", ".fi-header", ".fi-title", ".fi-actions"):
            self.assertIn(token, css, f"style.css 缺少样式定义: {token}")

    def test_09_style_css_file_info_bar_uses_design_tokens(self):
        css = self._get_static("style.css")
        idx = css.find(".file-info-bar")
        self.assertGreater(idx, -1, "style.css 未找到 .file-info-bar 定义")
        segment = css[idx:idx + 300]
        self.assertIn("var(--card)", segment, ".file-info-bar 未使用 var(--card) 令牌")
        self.assertIn("var(--border)", segment, ".file-info-bar 未使用 var(--border) 令牌")

    def test_10_index_html_accepts_xls(self):
        """前端文件选择控件应同时接受 .xlsx 与 .xls，且提示文案包含 .xls"""
        html = self._get_page("/data-import")
        self.assertIn('accept=".xlsx,.xls"', html, "文件选择控件未接受 .xls")
        self.assertIn(".xls", html, "页面提示文案未提及 .xls 格式")


if __name__ == "__main__":
    unittest.main()

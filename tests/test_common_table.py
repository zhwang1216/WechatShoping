"""公共表格组件 CommonTable 接口契约验收测试

验证 spec 5.9 最终验收：
  test_01: common_table.js 可被静态服务正确返回
  test_02: 页面引入 common_table.js 后全局可访问 CommonTable
  test_03: CommonTable 挂载 NativeAdapter/ModalAdapter/DataAdapter/ColumnRenderer
  test_04: ColumnRenderer renderHead/renderBody 列对齐与 hidden 列
  test_05: TableDataAdapter buildRequest 字段映射 (sort_field→order_by)
  test_06: TableDataAdapter parseResponse 兼容 size/page_size
  test_07: TableDataAdapter parseResponse pages 缺失补全
  test_08: CommonTable 必填项校验 (columns/rowKey/el 缺失时 console.warn)
  test_09: 既有页面回归 — orders 页面表格结构未破坏
  test_10: 既有页面回归 — profit 页面 parts-table 例外保留
"""
import unittest

from tests.http_client import TestEnv


class TestCommonTable(unittest.TestCase):
    """公共表格组件 CommonTable 接口契约验收测试"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin

    def test_01_js_served(self):
        """common_table.js 可被静态服务正确返回"""
        st, body, _ = self.admin.get_binary("/static/common_table.js")
        self.assertEqual(st, 200, f"common_table.js 返回 {st}")
        self.assertIn(b"CommonTable", body, "应包含 CommonTable 构造函数")
        self.assertIn(b"ColumnRenderer", body, "应包含 ColumnRenderer")
        self.assertIn(b"NativeAdapter", body, "应包含 NativeAdapter")
        self.assertIn(b"ModalAdapter", body, "应包含 ModalAdapter")
        self.assertIn(b"DataAdapter", body, "应包含 DataAdapter")

    def test_02_global_accessible(self):
        """页面引入 common_table.js 后全局可访问 CommonTable"""
        _, html, _ = self.admin.get_page("/orders")
        self.assertIn("/static/common_table.js", html,
                       "orders 页面应引入 common_table.js")
        self.assertIn("/static/data_table.js", html,
                       "orders 页面应引入 data_table.js")

    def test_03_static_namespaces(self):
        """CommonTable 挂载 NativeAdapter/ModalAdapter/DataAdapter/ColumnRenderer"""
        st, body, _ = self.admin.get_binary("/static/common_table.js")
        self.assertIn(b"CommonTable.NativeAdapter", body)
        self.assertIn(b"CommonTable.ModalAdapter", body)
        self.assertIn(b"CommonTable.DataAdapter", body)
        self.assertIn(b"CommonTable.ColumnRenderer", body)

    def test_04_column_renderer_contract(self):
        """ColumnRenderer renderHead/renderBody 契约 — hidden 列跳过"""
        st, body, _ = self.admin.get_binary("/static/common_table.js")
        self.assertIn(b"renderHead", body, "应有 renderHead 方法")
        self.assertIn(b"renderBody", body, "应有 renderBody 方法")
        self.assertIn(b"c.hidden", body, "应处理 hidden 列跳过")
        self.assertIn(b"data-row-key", body, "应写入 data-row-key 属性")

    def test_05_data_adapter_build_request(self):
        """TableDataAdapter buildRequest 字段映射 sort_field→order_by"""
        st, body, _ = self.admin.get_binary("/static/common_table.js")
        self.assertIn(b"order_by", body, "buildRequest 应映射到 order_by")
        self.assertIn(b"order_dir", body, "buildRequest 应映射到 order_dir")
        self.assertIn(b"sortFieldMap", body, "应支持 sortFieldMap 映射")

    def test_06_data_adapter_parse_response_size(self):
        """TableDataAdapter parseResponse 兼容 size 字段"""
        st, body, _ = self.admin.get_binary("/static/common_table.js")
        self.assertIn(b"rawData.size", body, "应兼容 size 响应字段")
        self.assertIn(b"rawData.page_size", body, "应兼容 page_size 响应字段")

    def test_07_data_adapter_pages_fallback(self):
        """TableDataAdapter parseResponse pages 缺失补全"""
        st, body, _ = self.admin.get_binary("/static/common_table.js")
        self.assertIn(b"Math.ceil(total / pageSize)", body,
                      "pages 缺失时应 Math.ceil(total/pageSize) 补全")

    def test_08_required_validation(self):
        """CommonTable 必填项校验 — columns/rowKey 缺失时 console.warn"""
        st, body, _ = self.admin.get_binary("/static/common_table.js")
        self.assertIn(b"columns required", body, "缺 columns 应 warn")
        self.assertIn(b"rowKey required", body, "缺 rowKey 应 warn")
        self.assertIn(b"el or tableId required", body, "缺 el/tableId 应 warn")

    def test_09_orders_page_table_structure(self):
        """既有页面回归 — orders 页面表格结构未破坏"""
        _, html, _ = self.admin.get_page("/orders")
        self.assertIn('class="table-wrap"', html, "table-wrap 结构保留")
        self.assertIn("<table", html, "table 元素保留")
        self.assertIn("<thead>", html, "thead 元素保留")
        self.assertIn("useDataTable: true", html, "orders 应启用 DataTable")

    def test_11_data_table_served(self):
        """data_table.js 可被静态服务且导出 DataTable"""
        st, body, _ = self.admin.get_binary("/static/data_table.js")
        self.assertEqual(st, 200)
        self.assertIn(b"global.DataTable", body)
        self.assertIn(b"data-table__viewport", body)

    def test_12_migrated_pages_include_data_table(self):
        """已迁移列表页均引入 data_table.js"""
        for path in ("/refund-orders", "/supplier-costs", "/anomalies", "/logs"):
            with self.subTest(path=path):
                _, html, _ = self.admin.get_page(path)
                self.assertIn("/static/data_table.js", html, path)
                self.assertIn("useDataTable: true", html, path)

    def test_10_profit_parts_table_exception(self):
        """既有页面回归 — profit 页面 parts-table 例外保留"""
        _, html, _ = self.admin.get_page("/profit")
        self.assertIn("parts-table", html, "parts-table 例外保留")
        self.assertIn('class="table-wrap"', html, "profit 主表格结构保留")


if __name__ == "__main__":
    unittest.main()

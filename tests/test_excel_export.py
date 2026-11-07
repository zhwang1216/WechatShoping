"""导出单元格：以公式触发符开头的文本必须仍是字符串，不能写成公式。"""
import io
import unittest
import zipfile

from openpyxl import Workbook, load_workbook

from app.excel_export import append_row


class TestAppendRow(unittest.TestCase):
    def test_formula_prefixes_stay_text(self):
        wb = Workbook()
        ws = wb.active
        append_row(ws, ["订单号", "金额"])
        append_row(ws, ["=1+1", 12])
        append_row(ws, ["+cmd", "-2", "@SUM(A1)", "普通商品"])
        buf = io.BytesIO()
        wb.save(buf)
        xml = zipfile.ZipFile(io.BytesIO(buf.getvalue())).read(
            "xl/worksheets/sheet1.xml").decode()
        self.assertNotIn("<f>", xml)
        loaded = load_workbook(io.BytesIO(buf.getvalue())).active
        self.assertEqual(loaded["A2"].data_type, "s")
        self.assertEqual(loaded["A2"].value, "=1+1")
        self.assertEqual(loaded["A3"].value, "+cmd")
        self.assertEqual(loaded["B3"].value, "-2")
        self.assertEqual(loaded["C3"].value, "@SUM(A1)")
        self.assertEqual(loaded["D3"].value, "普通商品")
        self.assertEqual(loaded["B2"].value, 12)
        self.assertEqual(loaded["B2"].data_type, "n")


if __name__ == "__main__":
    unittest.main()

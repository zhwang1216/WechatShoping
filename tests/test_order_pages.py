"""阶段3测试：普通订单与退款订单隔离、分类、汇总、重复退款、状态处理、分页筛选排序、详情逐笔、导出

样例覆盖：无退款 / 部分退款 / 全额退款 / 多次退款（多行不同金额）/
退款处理中 / 退款关闭（0 元与有金额，历史有退款记录）/
订单级退款金额多行重复（粒度未知） / 空订单号错误行 / 金额非法错误行
"""
import io
import json
import os
import unittest

from openpyxl import Workbook, load_workbook

from tests.http_client import TestEnv

HEADERS = ["订单号", "订单下单时间", "订单状态", "订单实际支付金额", "订单实际收款金额",
           "订单运费", "技术服务费", "商品实际价格(总共)", "商品数量",
           "商品已退款金额", "商品售后", "商品标题", "sku属性", "sku编码"]

SAMPLE_ROWS = [
    # A1001 无退款
    ["A1001", "2026-03-01 10:00:00", "已完成", 199, 199, 0, 1, 199, 1, 0, "-", "测试商品A", "红色", "SKU-A"],
    # B1002 部分退款（平台状态）
    ["B1002", "2026-03-01 11:00:00", "已完成", 100, 100, 0, 0.5, 100, 1, 30, "部分退款", "商品B", "蓝色", "SKU-B"],
    # C1003 全额退款（实收 0）
    ["C1003", "2026-03-02 09:00:00", "已关闭", 50, 0, 0, 0, 50, 1, 50, "已全额退款", "商品C", "默认", "SKU-C"],
    # D1004 多次退款：两行不同退款金额（商品行级粒度）
    ["D1004", "2026-03-03 10:00:00", "已完成", 80, 80, 0, 0.4, 45, 1, 20, "部分退款", "商品D", "红色", "SKU-D1"],
    ["D1004", "2026-03-03 10:00:00", "已完成", 80, 80, 0, 0.4, 35, 1, 10, "部分退款", "商品D", "蓝色", "SKU-D2"],
    # E1005 退款处理中（金额不得计入成功退款）
    ["E1005", "2026-03-04 10:00:00", "已完成", 60, 60, 0, 0.3, 60, 1, 15, "退款处理中", "商品E", "默认", "SKU-E"],
    # F1006 退款关闭但历史有退款记录（金额 0）
    ["F1006", "2026-03-05 10:00:00", "已完成", 60, 60, 0, 0.3, 60, 1, 0, "退款关闭", "商品F", "默认", "SKU-F"],
    # G1007 退款关闭且存在退款金额
    ["G1007", "2026-03-03 12:00:00", "已完成", 60, 60, 0, 0.3, 60, 1, 25, "退款关闭", "商品G", "默认", "SKU-G"],
    # H1008 订单级退款金额多行重复（粒度未知，只计一次）
    ["H1008", "2026-03-04 15:00:00", "已完成", 120, 120, 0, 0.6, 70, 1, 40, "部分退款", "商品H", "红色", "SKU-H1"],
    ["H1008", "2026-03-04 15:00:00", "已完成", 120, 120, 0, 0.6, 50, 1, 40, "部分退款", "商品H", "蓝色", "SKU-H2"],
    # 错误行：订单号为空
    [None, "2026-03-04 16:00:00", "已完成", 10, 10, 0, 0, 10, 1, 0, "-", "错误行", "默认", "SKU-X"],
    # 错误行：实收金额非法
    ["J1010", "2026-03-05 11:00:00", "已完成", "abc", 30, 0, 0, 30, 1, 0, "-", "错误行", "默认", "SKU-J"],
    # K1011 无退款（数量 2）
    ["K1011", "2026-03-05 12:00:00", "已完成", 88, 88, 0, 0.4, 176, 2, 0, "-", "商品K", "绿色", "SKU-K"],
]


def make_wechat_xlsx(rows=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.append(HEADERS)
    for r in (rows if rows is not None else SAMPLE_ROWS):
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestOrderPages(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member_store = cls.env.member_store
        # 导入样例微信订单（管理员代成员店铺导入）
        content = make_wechat_xlsx()
        cls.uploaded_content = content  # 重复上传测试复用同一份字节（哈希必然一致）
        status, resp = cls.admin.upload(
            f"/api/import/wechat-orders?store_id={cls.member_store}",
            "样例订单.xlsx", content)
        assert status == 200 and resp["code"] == 0, resp
        cls.upload_id = resp["data"]["upload_id"]
        cls.preview = resp["data"]
        status, resp = cls.admin.post("/api/import/confirm",
                                      {"upload_id": cls.upload_id,
                                       "target_store_id": cls.member_store})
        assert status == 200 and resp["code"] == 0, resp
        cls.import_result = resp["data"]
        cls.batch_id = resp["data"]["batch_id"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 请求辅助 ----------
    def _get(self, path):
        return self.admin.get(path)

    def _post(self, path, body):
        body = dict(body or {})
        body.setdefault("target_store_id", self.member_store)
        return self.admin.post(path, body)

    def _upload_static(self, path, filename, content):
        sep = "&" if "?" in path else "?"
        return self.admin.upload(f"{path}{sep}store_id={self.member_store}",
                                 filename, content)

    _upload = _upload_static

    def _get_binary(self, path):
        return self.admin.get_binary(path)

    def _order_ids(self, data, key="orders"):
        return [r["order_id"] for r in data[key]]

    # ---------- 导入与分类 ----------
    def test_import_statistics(self):
        s = self.import_result["statistics"]
        self.assertEqual(s["total_rows"], 13)
        self.assertEqual(s["success_rows"], 11)
        self.assertEqual(s["error_rows"], 2)
        self.assertEqual(s["updated_rows"], 0)
        self.assertEqual(self.import_result["status"], "partial_success")

    def test_normal_orders_exclude_refunds(self):
        status, resp = self._get("/api/orders?page=1&size=100")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        ids = self._order_ids(resp["data"])
        self.assertEqual(sorted(ids), ["A1001", "K1011"])
        for r in resp["data"]["orders"]:
            self.assertFalse(r["has_refund"])

    def test_refund_orders_include_all_refund_kinds(self):
        """五类退款场景全部纳入：部分/全额/多次/处理中/关闭（含历史退款）"""
        status, resp = self._get("/api/refund-orders?page=1&size=100")
        self.assertEqual(status, 200)
        ids = self._order_ids(resp["data"])
        expected = ["B1002", "C1003", "D1004", "E1005", "F1006", "G1007", "H1008"]
        self.assertEqual(sorted(ids), sorted(expected))
        self.assertNotIn("A1001", ids)
        self.assertNotIn("K1011", ids)
        by_id = {r["order_id"]: r for r in resp["data"]["orders"]}
        self.assertEqual(by_id["B1002"]["refund"]["refund_status"], "partial")
        self.assertEqual(by_id["C1003"]["refund"]["refund_status"], "full")
        self.assertEqual(by_id["D1004"]["refund"]["refund_status"], "partial")
        self.assertEqual(by_id["E1005"]["refund"]["refund_status"], "pending")
        self.assertEqual(by_id["F1006"]["refund"]["refund_status"], "closed")
        self.assertEqual(by_id["G1007"]["refund"]["refund_status"], "closed")

    def test_one_row_per_order(self):
        """同一原订单（多行商品）在退款列表仅一条汇总记录"""
        status, resp = self._get("/api/refund-orders?page=1&size=100")
        ids = self._order_ids(resp["data"])
        self.assertEqual(len(ids), len(set(ids)))
        by_id = {r["order_id"]: r for r in resp["data"]["orders"]}
        self.assertEqual(by_id["D1004"]["product_line_count"], 2)
        self.assertEqual(by_id["H1008"]["product_line_count"], 2)

    def test_closed_with_history_not_in_normal(self):
        """退款关闭但历史有退款记录 → 退款页而非普通页"""
        status, resp = self._get("/api/orders?page=1&size=100")
        self.assertNotIn("F1006", self._order_ids(resp["data"]))
        self.assertNotIn("G1007", self._order_ids(resp["data"]))

    # ---------- 退款明细（逐笔）与详情 ----------
    def test_refund_details_line_by_line(self):
        status, resp = self._get("/api/orders/D1004/refunds")
        self.assertEqual(status, 200)
        refunds = resp["data"]["refunds"]
        self.assertEqual(len(refunds), 2)
        amounts = sorted(r["refund_amount"] for r in refunds)
        self.assertEqual(amounts, ["10.00", "20.00"])
        for r in refunds:
            self.assertEqual(r["refund_status"], "partial")
            self.assertEqual(r["status_source"], "platform")
            self.assertEqual(r["sheet_name"], "订单")
            self.assertTrue(r["row_number"] >= 2)
            self.assertEqual(r["batch_id"], self.batch_id)

    def test_refund_details_multiple_same_amount(self):
        status, resp = self._get("/api/orders/H1008/refunds")
        refunds = resp["data"]["refunds"]
        self.assertEqual(len(refunds), 2)
        self.assertEqual([r["refund_amount"] for r in refunds], ["40.00", "40.00"])

    def test_order_detail_tracing(self):
        status, resp = self._get("/api/orders/D1004")
        self.assertEqual(status, 200)
        d = resp["data"]
        self.assertTrue(d["has_refund"])
        self.assertEqual(len(d["lines"]), 2)
        self.assertEqual(len(d["refunds"]), 2)
        self.assertEqual(len(d["batches"]), 1)
        self.assertEqual(d["batches"][0]["file_name"], "样例订单.xlsx")
        self.assertEqual(d["batches"][0]["batch_id"], self.batch_id)
        self.assertEqual(d["amounts"]["actual_received_amount"], "80.00")
        self.assertFalse(d["amount_conflict"])
        self.assertEqual(d["refund_view"]["refund"]["refund_amount"], "30.00")

    def test_order_not_found(self):
        status, resp = self._get("/api/orders/X9999")
        self.assertEqual(status, 404)
        self.assertEqual(resp["code"], 404)
        self.assertIn("订单不存在", resp["message"])

    def test_refunds_empty_for_normal_order(self):
        status, resp = self._get("/api/orders/A1001/refunds")
        self.assertEqual(status, 200)
        self.assertEqual(resp["data"]["total"], 0)
        self.assertEqual(resp["data"]["refunds"], [])

    def test_refunds_order_not_found(self):
        status, resp = self._get("/api/orders/X9999/refunds")
        self.assertEqual(status, 404)

    # ---------- 汇总与重复退款 ----------
    def test_refund_statistics_by_status(self):
        status, resp = self._get("/api/refund-orders?page=1&size=100")
        st = resp["data"]["statistics"]
        self.assertEqual(st["refund_order_count"], 7)
        self.assertEqual(st["success_refund_amount"], "150.00")
        self.assertEqual(st["pending_refund_amount"], "15.00")
        self.assertEqual(st["closed_refund_amount"], "25.00")
        self.assertEqual(st["total_refund_amount"], "190.00")
        self.assertEqual(st["partial_refund_order_count"], 3)
        self.assertEqual(st["full_refund_order_count"], 1)
        self.assertEqual(st["pending_review_count"], 1)

    def test_pending_amount_not_counted_as_success(self):
        status, resp = self._get("/api/refund-orders?order_id=E1005")
        r = resp["data"]["orders"][0]
        self.assertEqual(r["refund_by_status"]["pending"], "15.00")
        self.assertEqual(r["refund_by_status"]["success"], "0.00")
        self.assertEqual(r["refund_by_status"]["closed"], "0.00")

    def test_duplicate_refund_amount_deduplicated(self):
        """订单级退款金额多行重复 → 粒度未知，总额只计一次"""
        status, resp = self._get("/api/refund-orders?order_id=H1008")
        r = resp["data"]["orders"][0]
        self.assertEqual(r["refund"]["refund_amount"], "40.00")
        self.assertEqual(r["refund"]["refund_amount_granularity"], "unknown")
        self.assertEqual(r["refund"]["refund_count"], 2)
        self.assertTrue(r["pending_review"])
        self.assertEqual(r["refund_by_status"]["success"], "40.00")

    def test_line_granularity_refund_summed(self):
        """商品行级退款（多行不同金额）→ 逐行求和"""
        status, resp = self._get("/api/refund-orders?order_id=D1004")
        r = resp["data"]["orders"][0]
        self.assertEqual(r["refund"]["refund_amount"], "30.00")
        self.assertEqual(r["refund"]["refund_amount_granularity"], "line")
        self.assertEqual(r["refund"]["refund_count"], 2)

    def test_order_level_amount_not_duplicated(self):
        """一单多行订单级实收金额只计一次"""
        status, resp = self._get("/api/refund-orders?order_id=D1004")
        r = resp["data"]["orders"][0]
        self.assertEqual(r["amounts"]["actual_received_amount"], "80.00")
        status, resp = self._get("/api/orders/D1004")
        self.assertEqual(resp["data"]["amounts"]["actual_received_amount"], "80.00")
        self.assertEqual(resp["data"]["amounts"]["technical_service_fee"], "0.40")

    def test_normal_order_statistics(self):
        status, resp = self._get("/api/orders?page=1&size=100")
        st = resp["data"]["statistics"]
        self.assertEqual(st["order_count"], 2)
        self.assertEqual(st["actual_paid_total"], "287.00")
        self.assertEqual(st["actual_received_total"], "287.00")
        self.assertEqual(st["technical_service_fee_total"], "1.40")
        self.assertEqual(st["product_total_sum"], "375.00")
        self.assertEqual(st["quantity_total"], 3)

    # ---------- 筛选、排序、分页 ----------
    def test_filter_by_order_id(self):
        status, resp = self._get("/api/orders?order_id=A1001")
        self.assertEqual(self._order_ids(resp["data"]), ["A1001"])
        status, resp = self._get("/api/refund-orders?order_id=B1002")
        self.assertEqual(self._order_ids(resp["data"]), ["B1002"])

    def test_filter_by_refund_status(self):
        status, resp = self._get("/api/refund-orders?status=pending")
        self.assertEqual(self._order_ids(resp["data"]), ["E1005"])
        status, resp = self._get("/api/refund-orders?status=closed")
        self.assertEqual(sorted(self._order_ids(resp["data"])), ["F1006", "G1007"])

    def test_filter_by_order_status(self):
        status, resp = self._get("/api/orders?status=completed")
        self.assertEqual(sorted(self._order_ids(resp["data"])), ["A1001", "K1011"])

    def test_filter_by_time_range(self):
        status, resp = self._get("/api/refund-orders?start=2026-03-03")
        self.assertEqual(sorted(self._order_ids(resp["data"])),
                         ["D1004", "E1005", "F1006", "G1007", "H1008"])
        status, resp = self._get("/api/refund-orders?end=2026-03-02")
        self.assertEqual(sorted(self._order_ids(resp["data"])), ["B1002", "C1003"])
        status, resp = self._get("/api/orders?end=2026-03-02")
        self.assertEqual(self._order_ids(resp["data"]), ["A1001"])
        # 闭区间：start=2026-03-03 & end=2026-03-03 应含当天订单
        status, resp = self._get("/api/refund-orders?start=2026-03-03&end=2026-03-03")
        self.assertEqual(sorted(self._order_ids(resp["data"])), ["D1004", "G1007"])

    def test_filter_by_amount_range(self):
        status, resp = self._get("/api/refund-orders?min_amount=30&max_amount=40")
        ids = sorted(self._order_ids(resp["data"]))
        self.assertEqual(ids, ["B1002", "D1004", "H1008"])
        status, resp = self._get("/api/refund-orders?min_amount=45")
        self.assertEqual(sorted(self._order_ids(resp["data"])), ["C1003"])

    def test_sort_by_refund_amount(self):
        status, resp = self._get(
            "/api/refund-orders?order_by=refund_amount&order_dir=asc&page=1&size=100")
        ids = self._order_ids(resp["data"])
        self.assertEqual(ids[0], "F1006")  # 0 元最低
        self.assertEqual(ids[-1], "C1003")  # 50 元最高

    def test_sort_desc_by_order_time(self):
        status, resp = self._get(
            "/api/orders?order_by=order_time&order_dir=desc&page=1&size=100")
        ids = self._order_ids(resp["data"])
        self.assertEqual(ids[:2], ["K1011", "A1001"])

    def test_pagination(self):
        """导入 45 行无退款订单验证分页（普通订单 47 条 → 3 页）"""
        rows = []
        for i in range(1, 46):
            rows.append([f"L20{i:02d}", f"2026-02-{(i % 27) + 1:02d} 10:00:00",
                         "已完成", 10, 10, 0, 0, 10, 1, 0, "-",
                         f"批量订单{i}", "默认", "SKU-L"])
        content = make_wechat_xlsx(rows)
        status, resp = self._upload("/api/import/wechat-orders",
                                    "批量订单.xlsx", content)
        self.assertEqual(status, 200)
        status, resp = self._post("/api/import/confirm",
                                  {"upload_id": resp["data"]["upload_id"]})
        self.assertEqual(resp["data"]["statistics"]["success_rows"], 45)

        status, resp = self._get("/api/orders?page=1&size=20")
        d = resp["data"]
        self.assertEqual(d["total"], 47)
        self.assertEqual(d["pages"], 3)
        self.assertEqual(len(d["orders"]), 20)
        status, resp = self._get("/api/orders?page=3&size=20")
        self.assertEqual(len(resp["data"]["orders"]), 7)
        status, resp = self._get("/api/orders?page=9&size=50")
        self.assertEqual(resp["data"]["total"], 47)
        self.assertEqual(len(resp["data"]["orders"]), 0)
        # 退款订单分页不受普通订单导入影响
        status, resp = self._get("/api/refund-orders?page=1&size=20")
        self.assertEqual(resp["data"]["total"], 7)
        self.assertEqual(len(resp["data"]["orders"]), 7)

    def test_invalid_pagination_params(self):
        for q in ("page=-1", "page=0", "size=9999", "size=30", "size=abc",
                  "order_by=xxx", "order_dir=up"):
            status, resp = self._get(f"/api/orders?{q}")
            self.assertEqual(status, 400, q)
            self.assertEqual(resp["code"], 400, q)
            status, resp = self._get(f"/api/refund-orders?{q}")
            self.assertEqual(status, 400, q)

    # ---------- 导出 ----------
    def test_export_refund_orders_excel(self):
        status, content, headers = self._get_binary("/api/refund-orders/export")
        self.assertEqual(status, 200)
        self.assertIn("spreadsheetml", headers.get("Content-Type", ""))
        wb = load_workbook(io.BytesIO(content))
        ws = wb.active
        self.assertEqual(ws.title, "退款订单")
        rows = list(ws.iter_rows(values_only=True))
        self.assertEqual(len(rows), 8)  # 表头 + 7 单
        header = rows[0]
        self.assertIn("订单号", header)
        self.assertIn("退款总额", header)
        amt_col = header.index("退款总额") + 1
        for r in list(ws.iter_rows(min_row=2))[0:7]:
            cell = r[amt_col - 1]
            self.assertEqual(cell.number_format, "0.00")
        # 金额两位小数（退款总额列）
        values = [r[amt_col - 1] for r in rows[1:]]
        self.assertEqual(sorted(str(v) for v in values),
                         sorted(["0", "15", "25", "30", "30", "40", "50"]))
        # H1008 行待核对
        col_pr = header.index("待核对")
        col_oid = header.index("订单号")
        for r in rows[1:]:
            if r[col_oid] == "H1008":
                self.assertEqual(r[col_pr], "是")

    def test_export_matches_filter(self):
        """导出与筛选结果一致（status=pending → 仅 1 行）"""
        status, content, _ = self._get_binary(
            "/api/refund-orders/export?status=pending")
        self.assertEqual(status, 200)
        wb = load_workbook(io.BytesIO(content))
        rows = list(wb.active.iter_rows(values_only=True))
        self.assertEqual(len(rows), 2)  # 表头 + E1005
        header = rows[0]
        self.assertEqual(rows[1][header.index("订单号")], "E1005")

    def test_export_empty_returns_400(self):
        status, resp = self._get("/api/refund-orders/export?order_id=X9999")
        self.assertEqual(status, 400)
        self.assertIn("无数据", resp["message"])

    def test_export_records_logged(self):
        # 自包含: 先执行一次导出, 不依赖其他用例的执行顺序
        self._get_binary("/api/orders/export")
        status, resp = self._get("/api/data/export_records")
        records = resp["data"]["records"]
        self.assertGreaterEqual(len(records), 1)
        # 多用户架构：操作日志位于数据归属店铺目录（admin 代操作写目标店铺）
        log_file = os.path.join(
            self.env.data_dir, "stores", self.member_store,
            "logs", "operation_log.jsonl")
        with open(log_file, encoding="utf-8") as f:
            actions = [json.loads(line)["action"]
                       for line in f if line.strip()]
        self.assertIn("data.export", actions)

    # ---------- 重复上传保护（数据不被破坏） ----------
    def test_duplicate_upload_blocked(self):
        # 重复上传必须使用与已导入批次完全相同的字节（openpyxl 会嵌入秒级时间戳，
        # 重新生成会因跨秒导致哈希不同，从而绕过重复检测）
        content = self.uploaded_content
        status, resp = self._upload("/api/import/wechat-orders",
                                    "样例订单.xlsx", content)
        self.assertEqual(status, 200)
        self.assertIsNotNone(resp["data"]["duplicate_file"])
        self.assertEqual(resp["data"]["duplicate_file"]["batch_id"], self.batch_id)
        upload_id = resp["data"]["upload_id"]
        status, resp = self._post("/api/import/confirm", {"upload_id": upload_id})
        self.assertEqual(status, 409)
        self.assertEqual(resp["code"], 2003)
        # 数据未重复累计
        status, resp = self._get("/api/refund-orders?page=1&size=100")
        self.assertEqual(resp["data"]["total"], 7)
        status, resp = self._get("/api/orders?page=1&size=100")
        self.assertEqual(resp["data"]["total"], 2)

    # ---------- 错误行留存 ----------
    def test_error_rows_kept_in_raw_records(self):
        """不得静默丢弃原始行：错误行仍留存并可查询"""
        status, resp = self._get(
            f"/api/import/batches/{self.batch_id}/rows?status=error&size=100")
        self.assertEqual(resp["data"]["total"], 2)
        for row in resp["data"]["rows"]:
            self.assertTrue(row["validation"]["errors"])
        # 原始字段保留
        first = resp["data"]["rows"][0]["raw_fields"]
        self.assertTrue(first)

    def test_health_ok_after_all(self):
        status, resp = self._get("/api/health")
        self.assertEqual(resp["data"]["status"], "ok")


if __name__ == "__main__":
    unittest.main()

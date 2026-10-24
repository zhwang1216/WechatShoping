"""阶段4测试：订单与供应商成本对账（一对一 / 一对多 / 重复成本 / 未匹配 / 人工确认）

样例覆盖：
- O1001 一对一精确匹配（主订单编号）；
- O1002 一对多：两包裹各 25 元、快递单号互不相同 → 物流单级分别计入 50.00；
- O1003 重复成本：订单级 48 元重复 3 行（无快递单号）→ 去重只扣 48.00，
  粒度未确认前不得用于最终利润；人工确认 order_level/line_level 后分别生效；
- O1004 未匹配订单 + S9999 未匹配成本（孤儿），人工绑定后异常关闭、重跑/重导不覆盖；
- O1005 疑似匹配：仅辅助线索（同日 + SKU）命中 → suspected，待人工确认；
- O1006 金额差异 / O1007 数量差异 / O1008 成本缺失；
- O2001/O2002 多对一：同一条成本记录被两个订单匹配 → 异常；
- 筛选（日期/订单号/匹配状态/异常类型）与导出一致性。
"""
import io
import json
import unittest

from openpyxl import Workbook, load_workbook

from tests.http_client import TestEnv

WX_HEADERS = ["订单号", "订单下单时间", "订单状态", "订单实际支付金额",
              "订单实际收款金额", "商品数量", "商品标题", "sku编码"]

SUP_HEADERS = ["主订单编号", "订单编号", "订单创建时间", "商品数量", "总金额",
               "订单实收款", "sku编码", "快递单号", "商品标题"]

# (订单号, 时间, 支付, 实收, 数量, SKU)
WX_ROWS = [
    ("O1001", "2026-03-01 10:00:00", 100, 100, 1, "SKU-A"),
    ("O1002", "2026-03-01 11:00:00", 120, 120, 1, "SKU-B1"),
    ("O1002", "2026-03-01 11:00:00", 120, 120, 1, "SKU-B2"),
    ("O1003", "2026-03-02 10:00:00", 200, 200, 1, "SKU-C"),
    ("O1004", "2026-03-02 11:00:00", 60, 60, 1, "SKU-D"),
    ("O1005", "2026-03-05 10:00:00", 80, 80, 2, "SKU-X"),
    ("O1006", "2026-03-03 10:00:00", 100, 100, 1, "SKU-F"),
    ("O1007", "2026-03-03 11:00:00", 90, 90, 2, "SKU-G"),
    ("O1008", "2026-03-03 12:00:00", 70, 70, 1, "SKU-H"),
    ("O2001", "2026-03-04 10:00:00", 50, 50, 1, "SKU-M"),
    ("O2002", "2026-03-04 11:00:00", 50, 50, 1, "SKU-N"),
]

# (主订单编号, 订单编号, 创建时间, 数量, 总金额, 实收款, SKU, 快递单号)
SUP_ROWS = [
    ("O1001", None, "2026-03-01 10:05:00", 1, 48, 100, "SKU-A", None),
    ("O1002", None, "2026-03-01 11:05:00", 1, 25, 120, "SKU-B1", "SF001"),
    ("O1002", None, "2026-03-01 11:05:00", 1, 25, 120, "SKU-B2", "SF002"),
    # 同一订单 3 个商品行，每行"总金额"均误填订单总额 48（订单级重复场景）
    ("O1003", None, "2026-03-02 10:05:00", 1, 48, 200, "SKU-C1", None),
    ("O1003", None, "2026-03-02 10:05:00", 1, 48, 200, "SKU-C2", None),
    ("O1003", None, "2026-03-02 10:05:00", 1, 48, 200, "SKU-C3", None),
    (None, "SUP-777", "2026-03-05 09:30:00", 2, 30, 80, "SKU-X", None),
    ("O1006", None, "2026-03-03 10:05:00", 1, 45, 99, "SKU-F", None),
    ("O1007", None, "2026-03-03 11:05:00", 1, 40, 90, "SKU-G", None),
    ("O1008", None, "2026-03-03 12:05:00", 1, None, 70, "SKU-H", None),
    ("S9999", None, "2026-03-02 12:00:00", 1, 20, None, "SKU-Z", None),
    ("O2001", "O2002", "2026-03-04 10:05:00", 1, 30, 50, "SKU-M", None),
]


def make_wechat_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.append(WX_HEADERS)
    for oid, t, paid, recv, qty, sku in WX_ROWS:
        ws.append([oid, t, "已完成", paid, recv, qty, f"商品{oid}", sku])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_supplier_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = "供应商成本"
    ws.append(SUP_HEADERS)
    for main, sub, t, qty, total, recv, sku, trk in SUP_ROWS:
        ws.append([main, sub, t, qty, total, recv, sku, trk, f"商品{sku}"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestReconciliation(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # admin_user="tester"：操作人取会话身份，保持既有断言 operator == "tester" 口径
        cls.env = TestEnv(admin_user="tester")
        cls.admin = cls.env.admin
        cls.member_store = cls.env.member_store

        status, resp = cls.admin.upload(
            f"/api/import/wechat-orders?store_id={cls.member_store}",
            "对账订单.xlsx", make_wechat_xlsx())
        assert status == 200 and resp["code"] == 0, resp
        status, resp = cls.admin.post(
            "/api/import/confirm",
            {"upload_id": resp["data"]["upload_id"],
             "target_store_id": cls.member_store})
        assert status == 200 and resp["code"] == 0, resp

        status, resp = cls.admin.upload(
            f"/api/import/supplier-costs?store_id={cls.member_store}",
            "对账成本.xlsx", make_supplier_xlsx())
        assert status == 200 and resp["code"] == 0, resp
        cls.sup_upload_id = resp["data"]["upload_id"]
        status, resp = cls.admin.post(
            "/api/import/confirm",
            {"upload_id": cls.sup_upload_id,
             "target_store_id": cls.member_store})
        assert status == 200 and resp["code"] == 0, resp

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

    def _get_binary(self, path):
        return self.admin.get_binary(path)

    # ---------- 查询辅助 ----------
    def _result_of(self, order_id, **params):
        qs = "&".join([f"order_id={order_id}"] + [f"{k}={v}" for k, v in params.items()])
        status, resp = self._get(f"/api/reconciliation?{qs}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        rows = resp["data"]["results"]
        self.assertEqual(len(rows), 1, f"订单 {order_id} 应有且仅有一条对账记录")
        return rows[0]

    def _anomalies(self):
        status, resp = self._get("/api/data/anomalies")
        self.assertEqual(status, 200)
        return resp["data"]["records"]

    def _find_anomaly(self, order_id, etype):
        return [a for a in self._anomalies()
                if a.get("order_id") == order_id and a.get("exception_type") == etype]

    def _supplier_id(self, main):
        status, resp = self._get("/api/data/supplier_costs")
        self.assertEqual(status, 200)
        for r in resp["data"]["records"]:
            if r.get("main_order_id") == main:
                return r["record_id"]
        raise AssertionError(f"未找到供应商记录: {main}")

    # ---------- 执行对账与总体摘要 ----------
    def test_01_run_summary(self):
        status, resp = self._post("/api/reconciliation/run", {})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        data = resp["data"]
        self.assertEqual(data["status"], "completed_with_exceptions")
        s = data["summary"]
        # 11 行微信明细中 O1002 占 2 行（一对多），唯一订单号共 10 个
        self.assertEqual(s["wechat_orders"], 10)
        self.assertEqual(s["matched"], 8)
        self.assertEqual(s["manual_confirmed"], 0)
        self.assertEqual(s["suspected"], 1)
        self.assertEqual(s["unmatched_order"], 1)
        self.assertEqual(s["unmatched_supplier"], 1)
        self.assertEqual(s["duplicate_risk"], 2)
        self.assertEqual(s["profit_blocked"], 2)

    # ---------- 一对一 ----------
    def test_02_one_to_one_exact_match(self):
        r = self._result_of("O1001")
        self.assertEqual(r["match_status"], "matched")
        self.assertEqual(r["match_method"], "main_order_id_exact")
        self.assertEqual(r["confidence"], "high")
        self.assertEqual(len(r["supplier_record_ids"]), 1)
        self.assertEqual(r["cost"]["line_count"], 1)
        self.assertEqual(r["cost"]["total_amount"], "48.00")
        self.assertEqual(r["cost"]["granularity"], "supplier_line")
        self.assertTrue(r["cost"]["usable_for_profit"])
        self.assertTrue(r["checks"]["amount_match"])
        self.assertTrue(r["checks"]["quantity_match"])
        self.assertEqual(r["exception_codes"], [])

    # ---------- 一对多（物流单级分别计入） ----------
    def test_03_one_to_many_packages(self):
        r = self._result_of("O1002")
        self.assertEqual(r["match_status"], "matched")
        self.assertEqual(r["match_method"], "main_order_id_exact")
        self.assertEqual(len(r["supplier_record_ids"]), 2)
        self.assertEqual(r["cost"]["granularity"], "package_level")
        self.assertEqual(r["cost"]["total_amount"], "50.00")
        self.assertEqual(r["cost"]["raw_line_sum"], "50.00")
        self.assertTrue(r["cost"]["duplicate_total_amount_risk"])
        self.assertFalse(r["cost"]["order_level_dedup_pending"])
        self.assertTrue(r["cost"]["usable_for_profit"])
        self.assertTrue(r["checks"]["amount_match"])

    # ---------- 重复成本（订单级去重 + 粒度未确认禁用） ----------
    def test_04_duplicate_order_level_cost(self):
        r = self._result_of("O1003")
        self.assertEqual(r["match_status"], "matched")
        self.assertEqual(len(r["supplier_record_ids"]), 3)
        self.assertEqual(r["cost"]["granularity"], "order_level")
        # 48 元重复 3 行：去重后只扣一次，禁止直接累加为 144
        self.assertEqual(r["cost"]["total_amount"], "48.00")
        self.assertEqual(r["cost"]["raw_line_sum"], "144.00")
        self.assertEqual(r["cost"]["dedup_amount"], "96.00")
        self.assertTrue(r["cost"]["duplicate_total_amount_risk"])
        self.assertTrue(r["cost"]["order_level_dedup_pending"])
        # 供应商总金额粒度未确认前，不得用于最终利润计算
        self.assertFalse(r["cost"]["usable_for_profit"])
        self.assertIn("DUPLICATE_TOTAL_AMOUNT_RISK", r["exception_codes"])
        self.assertEqual(len(self._find_anomaly("O1003", "DUPLICATE_TOTAL_AMOUNT_RISK")), 1)
        # 去重条目应标记被排除的重复记录
        entry = next(e for e in r["cost"]["lines"] if e["granularity"] == "order_level")
        self.assertEqual(len(entry["duplicated_record_ids"]), 2)

    # ---------- 未匹配订单 / 未匹配成本 ----------
    def test_05_unmatched_order_and_orphan_cost(self):
        r = self._result_of("O1004")
        self.assertEqual(r["match_status"], "unmatched_order")
        self.assertEqual(r["supplier_record_ids"], [])
        self.assertIn("UNMATCHED_ORDER", r["exception_codes"])
        self.assertEqual(len(self._find_anomaly("O1004", "UNMATCHED_ORDER")), 1)

        status, resp = self._get("/api/reconciliation?order_id=S9999&match_status=unmatched_supplier")
        self.assertEqual(status, 200)
        rows = resp["data"]["results"]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["supplier_order_key"], "S9999")
        self.assertEqual(rows[0]["match_status"], "unmatched_supplier")
        self.assertEqual(len(rows[0]["supplier_record_ids"]), 1)
        self.assertEqual(len(self._find_anomaly("S9999", "UNMATCHED_SUPPLIER_COST")), 1)

    # ---------- 疑似匹配（辅助候选，禁止自动认定） ----------
    def test_06_suspected_assisted_match(self):
        r = self._result_of("O1005")
        self.assertEqual(r["match_status"], "suspected")
        self.assertEqual(r["match_method"], "assisted_candidate")
        self.assertEqual(r["confidence"], "low")
        self.assertEqual(len(r["supplier_record_ids"]), 1)
        self.assertIn("SUSPECTED_MATCH_NEED_REVIEW", r["exception_codes"])
        # 疑似候选记录不再被判定为孤儿
        self.assertEqual(self._find_anomaly("SUP-777", "UNMATCHED_SUPPLIER_COST"), [])

    # ---------- 金额差异 / 数量差异 / 成本缺失 ----------
    def test_07_differences_and_missing(self):
        r6 = self._result_of("O1006")
        self.assertEqual(r6["checks"]["amount_diff"], "1.00")
        self.assertFalse(r6["checks"]["amount_match"])
        self.assertIn("AMOUNT_DIFF", r6["exception_codes"])

        r7 = self._result_of("O1007")
        self.assertEqual(r7["checks"]["quantity_diff"], 1)
        self.assertIn("QUANTITY_DIFF", r7["exception_codes"])

        r8 = self._result_of("O1008")
        self.assertEqual(r8["cost"]["missing_amount_lines"], 1)
        self.assertFalse(r8["cost"]["usable_for_profit"])
        self.assertIn("COST_MISSING", r8["exception_codes"])

    # ---------- 多对一异常 ----------
    def test_08_multi_order_one_cost(self):
        r1 = self._result_of("O2001")
        r2 = self._result_of("O2002")
        self.assertEqual(r1["match_status"], "matched")
        self.assertEqual(r2["match_status"], "matched")
        self.assertEqual(set(r1["supplier_record_ids"]),
                         set(r2["supplier_record_ids"]))
        found = [a for a in self._anomalies()
                 if a.get("exception_type") == "MULTI_ORDER_ONE_COST"
                 and "O2001" in a.get("message", "")]
        self.assertEqual(len(found), 1)

    # ---------- 人工确认匹配 ----------
    def test_09_manual_confirm_with_audit(self):
        sid = self._supplier_id("S9999")
        status, resp = self._post("/api/reconciliation/manual-match", {
            "order_id": "O1004",
            "supplier_record_ids": [sid],
            "note": "人工核对后确认绑定",
        })
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        rec = resp["data"]
        self.assertEqual(rec["match_status"], "manual_confirmed")
        self.assertEqual(rec["match_method"], "manual")
        self.assertEqual(rec["supplier_record_ids"], [sid])
        self.assertEqual(rec["cost"]["total_amount"], "20.00")
        history = rec["manual_review"]["history"]
        self.assertEqual(len(history), 1)
        h = history[0]
        # 记录操作人标签、时间、匹配规则与修改前后值
        self.assertEqual(h["operator"], "tester")
        self.assertTrue(h["operated_at"])
        self.assertEqual(h["match_rule"], "manual")
        self.assertEqual(h["before"]["match_status"], "unmatched_order")
        self.assertEqual(h["before"]["supplier_record_ids"], [])
        self.assertEqual(h["after"]["match_status"], "manual_confirmed")
        self.assertEqual(h["after"]["supplier_record_ids"], [sid])
        # 人工确认后相关异常关闭（留痕）
        a1 = self._find_anomaly("O1004", "UNMATCHED_ORDER")[0]
        self.assertEqual(a1["status"], "resolved")
        self.assertIn("人工确认", a1["resolution"]["text"])
        a2 = self._find_anomaly("S9999", "UNMATCHED_SUPPLIER_COST")[0]
        self.assertEqual(a2["status"], "resolved")

    # ---------- 人工确认不被重跑/重导覆盖 ----------
    def test_10_manual_match_preserved(self):
        sid = self._supplier_id("S9999")
        # 重跑对账：人工关系保留
        status, resp = self._post("/api/reconciliation/run", {})
        self.assertEqual(status, 200)
        r = self._result_of("O1004")
        self.assertEqual(r["match_status"], "manual_confirmed")
        self.assertEqual(r["match_method"], "manual")
        self.assertEqual(r["supplier_record_ids"], [sid])
        self.assertTrue(r["preserved_manual"])
        # 重新导入同一供应商文件（强制）：人工匹配关系仍不被覆盖
        status, resp = self._upload_static("/api/import/supplier-costs",
                                           "对账成本.xlsx", make_supplier_xlsx())
        self.assertEqual(status, 200)
        status, resp = self._post("/api/import/confirm",
                                  {"upload_id": resp["data"]["upload_id"], "force": True})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        r = self._result_of("O1004")
        self.assertEqual(r["match_status"], "manual_confirmed")
        self.assertEqual(r["supplier_record_ids"], [sid])
        self.assertEqual(len(r["manual_review"]["history"]), 1)

    # ---------- 人工匹配冲突与参数校验 ----------
    def test_11_manual_match_conflicts(self):
        sid = self._supplier_id("S9999")
        status, resp = self._post("/api/reconciliation/manual-match", {
            "order_id": "O1001", "supplier_record_ids": [sid]})
        self.assertEqual(status, 409)
        self.assertEqual(resp["code"], 3003)
        status, resp = self._post("/api/reconciliation/manual-match", {
            "order_id": "NOPE", "supplier_record_ids": [sid]})
        self.assertEqual(status, 404)
        status, resp = self._post("/api/reconciliation/manual-match", {
            "order_id": "O1001", "supplier_record_ids": ["sup_not_exist"]})
        self.assertEqual(status, 404)
        status, resp = self._post("/api/reconciliation/manual-match", {
            "order_id": "O1001", "supplier_record_ids": []})
        self.assertEqual(status, 400)

    # ---------- 成本粒度人工确认（订单级只扣一次 / 行级逐行累加） ----------
    def test_12_granularity_confirm(self):
        r0 = self._result_of("O1003")
        sids = r0["supplier_record_ids"]
        # 确认为订单级：同额多行只扣一次，可用于利润
        status, resp = self._post("/api/reconciliation/manual-match", {
            "order_id": "O1003", "supplier_record_ids": sids,
            "cost_granularity": "order_level", "note": "确认为订单级综合成本"})
        self.assertEqual(status, 200)
        rec = resp["data"]
        self.assertEqual(rec["cost"]["confirmed_granularity"], "order_level")
        self.assertEqual(rec["cost"]["total_amount"], "48.00")
        self.assertTrue(rec["cost"]["usable_for_profit"])
        self.assertFalse(rec["cost"]["order_level_dedup_pending"])
        # 改确认为行级：各明细行分别累加
        status, resp = self._post("/api/reconciliation/manual-match", {
            "order_id": "O1003", "supplier_record_ids": sids,
            "cost_granularity": "line_level", "note": "核实为三笔独立明细"})
        self.assertEqual(status, 200)
        rec = resp["data"]
        self.assertEqual(rec["cost"]["total_amount"], "144.00")
        self.assertTrue(rec["cost"]["usable_for_profit"])
        self.assertEqual(len(rec["manual_review"]["history"]), 2)
        # 重跑对账后确认粒度与结果保留
        status, resp = self._post("/api/reconciliation/run", {})
        self.assertEqual(status, 200)
        r = self._result_of("O1003")
        self.assertEqual(r["cost"]["confirmed_granularity"], "line_level")
        self.assertEqual(r["cost"]["total_amount"], "144.00")
        self.assertTrue(r["cost"]["usable_for_profit"])
        self.assertEqual(len(r["manual_review"]["history"]), 2)

    # ---------- 解除匹配 ----------
    def test_13_unbind(self):
        status, resp = self._post("/api/reconciliation/manual-match", {
            "order_id": "O1004", "action": "unbind", "note": "绑定错误，解除"})
        self.assertEqual(status, 200)
        rec = resp["data"]
        self.assertEqual(rec["match_status"], "unmatched_order")
        self.assertEqual(rec["supplier_record_ids"], [])
        h = rec["manual_review"]["history"][-1]
        self.assertEqual(h["action"], "unbind")
        self.assertEqual(h["before"]["match_status"], "manual_confirmed")
        self.assertEqual(h["after"]["match_status"], "unmatched_order")
        # 解除后重跑：订单回到未匹配，供应商记录回到孤儿
        status, resp = self._post("/api/reconciliation/run", {})
        self.assertEqual(status, 200)
        r = self._result_of("O1004")
        self.assertEqual(r["match_status"], "unmatched_order")
        status, resp = self._get("/api/reconciliation?order_id=S9999&match_status=unmatched_supplier")
        self.assertEqual(len(resp["data"]["results"]), 1)

    # ---------- 筛选与导出 ----------
    def test_14_filters_and_export(self):
        # 按匹配状态
        status, resp = self._get("/api/reconciliation?match_status=unmatched_order&size=100")
        self.assertEqual(status, 200)
        ids = {r["order_id"] for r in resp["data"]["results"]}
        self.assertEqual(ids, {"O1004"})
        # 按订单号
        r = self._result_of("O1001")
        self.assertEqual(r["order_id"], "O1001")
        # 按异常类型
        status, resp = self._get("/api/reconciliation?anomaly_type=DUPLICATE_TOTAL_AMOUNT_RISK&size=100")
        ids = {r["order_id"] for r in resp["data"]["results"]}
        self.assertEqual(ids, {"O1002", "O1003"})
        # 按日期
        status, resp = self._get("/api/reconciliation?start=2026-03-01&end=2026-03-01&size=100")
        ids = {r["order_id"] for r in resp["data"]["results"]}
        self.assertEqual(ids, {"O1001", "O1002"})
        # 组合筛选
        status, resp = self._get("/api/reconciliation?start=2026-03-03&end=2026-03-03&anomaly_type=AMOUNT_DIFF&size=100")
        ids = {r["order_id"] for r in resp["data"]["results"]}
        self.assertEqual(ids, {"O1006"})
        # 导出与筛选一致（重复风险 2 行数据 + 表头）
        status, content, headers = self._get_binary(
            "/api/reconciliation/export?anomaly_type=DUPLICATE_TOTAL_AMOUNT_RISK")
        self.assertEqual(status, 200)
        wb = load_workbook(io.BytesIO(content))
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0][0], "订单号")
        self.assertEqual({rows[1][0], rows[2][0]}, {"O1002", "O1003"})
        self.assertIn("attachment", headers.get("Content-Disposition", ""))
        # 导出记录与操作日志留痕
        status, resp = self._get("/api/data/export_records")
        self.assertTrue(any(e["kind"] == "reconciliation"
                            for e in resp["data"]["records"]))


if __name__ == "__main__":
    unittest.main()

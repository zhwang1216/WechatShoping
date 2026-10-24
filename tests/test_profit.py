"""阶段5测试：利润核算与经营总览（净收入 / 成本扣减 / 贡献利润 / 利润率 / 补贴 / 口径 / 追溯 / 总览 / 导出）

手工核算数据集（全部按 Decimal 精确到分，误差 0）：
- P1001 实收100 技术费3；供应商总金额48（成本40+快递5+泡沫1+加收1+燃油1，分项仅展示）+ 平台补贴6
  - 补贴口径 unknown：利润 = 100 + 0 − 48 − 3 = 49.00（与冒烟参考一致）
  - 补贴口径 not_included：利润 = 100 + 6 − 48 − 3 = 55.00，利润率 55.00%
  - 补贴口径 included：不重复计入，利润仍 49.00
- P1002 实收80，两行退款：部分退款10（成功）+ 退款处理中5；总金额30 技术费3
  - not_deducted：净收入 = 80 − 10 = 70.00（处理中5不计入）；利润 = 70 − 30 − 3 = 37.00
  - deducted：净收入 = 80.00（不重复扣）；利润 = 80 − 30 − 3 = 47.00
  - unknown：待确认 + REFUND_DEDUCTION_UNCONFIRMED
- P1003 实收60 全额退款60；总金额25 技术费2
  - not_deducted：净收入 0 → 利润率不计算（NET_REVENUE_NOT_POSITIVE）；利润 = −27.00（负利润异常）
  - deducted：净收入 60 → 利润 = 60 − 25 − 2 = 33.00
- P1004 无供应商成本 → 不可计算（UNMATCHED_NO_COST）
- P1005 供应商 3 行同额 48 无快递单号 → 去重只扣一次，粒度确认前待确认
  （DUPLICATE_RISK_UNRESOLVED）；确认 order_level 后利润 = 90 − 48 = 42.00
- P1006 供应商总金额为空 → 不可计算（SUPPLIER_COST_MISSING，不默认按 0）
- P1007 两行实收 100/98 冲突 → needs_review（RECEIVED_AMOUNT_CONFLICT），取首值100；
  利润 = 100 − 40 = 60.00
"""
import io
import json
import time
import unittest

from openpyxl import Workbook, load_workbook

from tests.http_client import TestEnv

WX_HEADERS = ["订单号", "订单下单时间", "订单状态", "订单实际支付金额",
              "订单实际收款金额", "商品数量", "商品标题", "sku编码",
              "商品实际价格(总共)", "技术服务费", "商品已退款金额", "商品售后"]

SUP_HEADERS = ["主订单编号", "订单编号", "订单创建时间", "商品数量", "总金额",
               "成本", "快递费", "泡沫耗材费", "加收", "燃油费", "平台补贴",
               "sku编码", "快递单号", "商品标题"]

# (订单号, 时间, 支付, 实收, 数量, 标题, SKU, 商品总价, 技术费, 退款金额, 售后状态)
WX_ROWS = [
    ("P1001", "2026-03-01 10:00:00", 100, 100, 1, "商品A", "SKU-A", 100, 3, None, None),
    ("P1002", "2026-03-01 11:00:00", 80, 80, 1, "商品B1", "SKU-B1", 40, 3, 10, "部分退款"),
    ("P1002", "2026-03-01 11:00:00", 80, 80, 1, "商品B2", "SKU-B2", 40, 3, 5, "退款处理中"),
    ("P1003", "2026-03-05 10:00:00", 60, 60, 1, "商品C", "SKU-C", 60, 2, 60, "全额退款"),
    ("P1004", "2026-03-02 10:00:00", 50, 50, 1, "商品D", "SKU-D", 50, None, None, None),
    ("P1005", "2026-03-02 11:00:00", 90, 90, 3, "商品E", "SKU-E", 90, None, None, None),
    ("P1006", "2026-03-02 12:00:00", 70, 70, 1, "商品F", "SKU-F", 70, None, None, None),
    ("P1007", "2026-03-03 10:00:00", 100, 100, 1, "商品G1", "SKU-G1", 50, None, None, None),
    ("P1007", "2026-03-03 10:00:00", 100, 98, 1, "商品G2", "SKU-G2", 50, None, None, None),
]

# (主订单编号, 创建时间, 数量, 总金额, 成本, 快递费, 泡沫耗材费, 加收, 燃油费, 补贴, SKU, 标题)
SUP_ROWS = [
    ("P1001", "2026-03-01 10:05:00", 1, 48, 40, 5, 1, 1, 1, 6, "SKU-A", "商品A"),
    ("P1002", "2026-03-01 11:05:00", 2, 30, None, None, None, None, None, None, "SKU-B", "商品B"),
    ("P1003", "2026-03-05 10:05:00", 1, 25, None, None, None, None, None, None, "SKU-C", "商品C"),
    ("P1005", "2026-03-02 11:05:00", 1, 48, None, None, None, None, None, None, "SKU-E1", "商品E1"),
    ("P1005", "2026-03-02 11:05:00", 1, 48, None, None, None, None, None, None, "SKU-E2", "商品E2"),
    ("P1005", "2026-03-02 11:05:00", 1, 48, None, None, None, None, None, None, "SKU-E3", "商品E3"),
    ("P1006", "2026-03-02 12:05:00", 1, None, None, None, None, None, None, None, "SKU-F", "商品F"),
    ("P1007", "2026-03-03 10:05:00", 2, 40, None, None, None, None, None, None, "SKU-G", "商品G"),
]


def make_wechat_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.append(WX_HEADERS)
    for oid, t, paid, recv, qty, title, sku, ptotal, tech, refund, after in WX_ROWS:
        ws.append([oid, t, "已完成", paid, recv, qty, title, sku,
                   ptotal, tech, refund, after])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_supplier_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = "供应商成本"
    ws.append(SUP_HEADERS)
    for main, t, qty, total, cost, courier, foam, surcharge, fuel, subsidy, sku, title in SUP_ROWS:
        ws.append([main, None, t, qty, total, cost, courier, foam, surcharge, fuel,
                   subsidy, sku, None, title])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestProfit(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member_store = cls.env.member_store
        cls.first_calc_id = None
        cls.latest_calc_id = None

        status, resp = cls.admin.upload(
            f"/api/import/wechat-orders?store_id={cls.member_store}",
            "利润订单.xlsx", make_wechat_xlsx())
        assert status == 200 and resp["code"] == 0, resp
        status, resp = cls.admin.post(
            "/api/import/confirm",
            {"upload_id": resp["data"]["upload_id"],
             "target_store_id": cls.member_store})
        assert status == 200 and resp["code"] == 0, resp

        status, resp = cls.admin.upload(
            f"/api/import/supplier-costs?store_id={cls.member_store}",
            "利润成本.xlsx", make_supplier_xlsx())
        assert status == 200 and resp["code"] == 0, resp
        status, resp = cls.admin.post(
            "/api/import/confirm",
            {"upload_id": resp["data"]["upload_id"],
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

    # ---------- 业务辅助 ----------
    def _run_recon(self):
        status, resp = self._post("/api/reconciliation/run", {})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        return resp["data"]

    def _recalc(self):
        # 核算记录 created_at 为秒级精度：保证相邻两次核算时间严格递增，
        # 避免"最新核算版本"在同一秒内按 calculation_id 字典序误判
        time.sleep(1.05)
        status, resp = self._post("/api/profit/recalculate", {})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        data = resp["data"]
        if self.first_calc_id is None:
            type(self).first_calc_id = data["calculation_id"]
        type(self).latest_calc_id = data["calculation_id"]
        return data

    @staticmethod
    def _stats_of(data):
        """recalculate 响应中的核算统计（平铺字段）"""
        return {k: data[k] for k in ("total", "calculated", "pending_confirmation",
                                     "not_calculable", "needs_review",
                                     "profit_total")}

    def _profit_of(self, order_id, **params):
        qs = "&".join([f"order_id={order_id}"]
                      + [f"{k}={v}" for k, v in params.items()])
        status, resp = self._get(f"/api/profit?{qs}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        rows = resp["data"]["results"]
        self.assertEqual(len(rows), 1, f"订单 {order_id} 应有且仅有一条利润记录")
        return rows[0]

    def _recon_of(self, order_id):
        status, resp = self._get(f"/api/reconciliation?order_id={order_id}")
        self.assertEqual(status, 200)
        rows = resp["data"]["results"]
        self.assertEqual(len(rows), 1)
        return rows[0]

    def _anomalies(self):
        status, resp = self._get("/api/data/anomalies")
        self.assertEqual(status, 200)
        return resp["data"]["records"]

    def _find_anomaly(self, order_id, etype):
        return [a for a in self._anomalies()
                if a.get("order_id") == order_id and a.get("exception_type") == etype]

    # ---------- 对账基础 ----------
    def test_01_reconciliation_base(self):
        """对账基线：7 个订单、6 匹配、1 未匹配、1 重复风险、2 利润阻断"""
        data = self._run_recon()
        self.assertEqual(data["status"], "completed_with_exceptions")
        s = data["summary"]
        self.assertEqual(s["total"], 7)
        self.assertEqual(s["wechat_orders"], 7)
        self.assertEqual(s["matched"], 6)
        self.assertEqual(s["manual_confirmed"], 0)
        self.assertEqual(s["suspected"], 0)
        self.assertEqual(s["unmatched_order"], 1)
        self.assertEqual(s["unmatched_supplier"], 0)
        self.assertEqual(s["duplicate_risk"], 1)
        # P1005 去重挂起 + P1006 成本缺失，均阻断利润
        self.assertEqual(s["profit_blocked"], 2)
        # P1001 成本明细：总金额 48，可用于利润
        r = self._recon_of("P1001")
        self.assertEqual(r["cost"]["total_amount"], "48.00")
        self.assertTrue(r["cost"]["usable_for_profit"])

    # ---------- 未知口径：基础利润公式 + 状态机 ----------
    def test_02_unknown_caliber_and_basic_profit(self):
        """unknown 口径：贡献利润=净收入+计入补贴−总金额−技术费；口径未确认/缺数据的状态"""
        data = self._recalc()

        # 核算元数据：规则版本 / 批次 / 时间
        self.assertTrue(data["calculation_id"].startswith("calc_"))
        self.assertEqual(data["rules_version"], "P1.1")
        self.assertTrue(data["calculated_at"])
        self.assertEqual(data["policies"], {
            "refund_deduction_status": "unknown",
            "platform_subsidy_included": "unknown"})
        self.assertEqual(len(data["source_batch_ids"]["wechat_orders"]), 1)
        self.assertEqual(len(data["source_batch_ids"]["supplier_costs"]), 1)

        # 统计：calculated=1 pending=3 not_calculable=2 needs_review=1
        # profit_total 仅含 calculated/needs_review：49 + 60 = 109.00
        self.assertEqual(self._stats_of(data), {
            "total": 7, "calculated": 1, "pending_confirmation": 3,
            "not_calculable": 2, "needs_review": 1, "profit_total": "109.00"})
        # 新增异常：退款口径未确认×2 + 负利润×1
        self.assertEqual(data["anomalies_created"], 3)
        self.assertEqual(data["anomalies_auto_resolved"], 0)

        # P1001：利润 49.00 = 100 + 0 − 48 − 3（成本/物流分项不重复扣减）
        r = self._profit_of("P1001")
        self.assertEqual(r["profit_status"], "calculated")
        self.assertEqual(r["status_reasons"], [])
        a = r["amounts"]
        self.assertEqual(a["actual_received_amount"], "100.00")
        self.assertEqual(a["net_revenue"], "100.00")
        self.assertEqual(a["platform_subsidy_total"], "6.00")
        self.assertEqual(a["platform_subsidy_credited"], "0.00")
        self.assertEqual(a["platform_subsidy_excluded"], "6.00")
        self.assertEqual(a["supplier_total_amount"], "48.00")
        self.assertEqual(a["technical_service_fee"], "3.00")
        self.assertEqual(a["operating_contribution_profit"], "49.00")
        self.assertEqual(r["margin"]["percent"], "49.00%")
        self.assertEqual(r["margin"]["numeric"], "49.00")
        self.assertIsNone(r["margin"]["blocked_reason"])
        # 分项仅展示：若重复扣减（48+40+5+1+1+1），利润将远小于 49
        self.assertEqual(r["cost_parts_display_only"], {
            "product_supply_cost": "40.00", "courier_fee": "5.00",
            "foam_material_fee": "1.00", "surcharge": "1.00",
            "fuel_or_tianjin_packing_fee": "1.00", "courier_fee_total": "0.00"})
        # 人工复核公式明细
        f = r["formula"]
        self.assertEqual(f["net_revenue"],
                         "口径=未确认（参考值）：净收入 = 实收 100.00"
                         " − 成功退款 0.00（处理中退款不计入）= 100.00")
        self.assertEqual(f["profit"],
                         "净收入 100.00 + 计入补贴 0.00"
                         " − 供应商综合总金额 48.00 − 技术服务费 3.00"
                         " − 其他经营费用 0.00 = 49.00")
        self.assertEqual(f["margin"], "49.00 ÷ 净收入 100.00 = 49.00%")
        self.assertEqual(f["cost_deduction_basis"],
                         "仅扣除供应商「总金额」（综合成本）一次：48.00")
        self.assertEqual(f["parts_note"],
                         "综合总金额中已包含的成本/物流分项仅用于展示核对，不重复扣减")

        # P1002：unknown + 有退款 → 待确认；净收入参考值 70（处理中 5 不计入）
        r = self._profit_of("P1002")
        self.assertEqual(r["profit_status"], "pending_confirmation")
        self.assertEqual(r["status_reasons"], ["REFUND_DEDUCTION_UNCONFIRMED"])
        self.assertEqual(r["amounts"]["refund_success_amount"], "10.00")
        self.assertEqual(r["amounts"]["refund_pending_amount"], "5.00")
        self.assertEqual(r["amounts"]["net_revenue"], "70.00")
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "37.00")
        self.assertEqual(r["formula"]["net_revenue"],
                         "口径=未确认（参考值）：净收入 = 实收 80.00"
                         " − 成功退款 10.00（处理中退款不计入）= 70.00")
        self.assertEqual(len(self._find_anomaly("P1002",
                                                "REFUND_DEDUCTION_UNCONFIRMED")), 1)

        # P1003：全额退款 + unknown → 待确认；净收入 0 → 不计算利润率
        r = self._profit_of("P1003")
        self.assertEqual(r["profit_status"], "pending_confirmation")
        self.assertEqual(r["amounts"]["net_revenue"], "0.00")
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "-27.00")
        self.assertIsNone(r["margin"]["percent"])
        self.assertEqual(r["margin"]["blocked_reason"], "NET_REVENUE_NOT_POSITIVE")
        # 负利润异常（RC-009）
        self.assertEqual(len(self._find_anomaly("P1003", "NEGATIVE_PROFIT")), 1)

        # P1004：未匹配 → 不可计算
        r = self._profit_of("P1004")
        self.assertEqual(r["profit_status"], "not_calculable")
        self.assertEqual(r["status_reasons"], ["UNMATCHED_NO_COST"])
        self.assertIsNone(r["amounts"]["supplier_total_amount"])
        self.assertIsNone(r["amounts"]["operating_contribution_profit"])
        self.assertIsNone(r["margin"]["percent"])
        self.assertEqual(r["margin"]["blocked_reason"], "NOT_CALCULABLE")
        self.assertEqual(r["formula"]["cost_deduction_basis"],
                         "供应商综合总金额缺失")

        # P1005：同额 3 行去重只扣一次，但粒度未确认 → 待确认
        r = self._profit_of("P1005")
        self.assertEqual(r["profit_status"], "pending_confirmation")
        self.assertEqual(r["status_reasons"], ["DUPLICATE_RISK_UNRESOLVED"])
        self.assertEqual(r["amounts"]["supplier_total_amount"], "48.00")

        # P1006：供应商总金额缺失 → 不可计算（禁止默认按 0）
        r = self._profit_of("P1006")
        self.assertEqual(r["profit_status"], "not_calculable")
        self.assertEqual(r["status_reasons"], ["SUPPLIER_COST_MISSING"])
        self.assertIsNone(r["amounts"]["supplier_total_amount"])
        self.assertIsNone(r["amounts"]["operating_contribution_profit"])

        # P1007：多行实收冲突 → 需复核，取首值 100
        r = self._profit_of("P1007")
        self.assertEqual(r["profit_status"], "needs_review")
        self.assertEqual(r["status_reasons"], ["RECEIVED_AMOUNT_CONFLICT"])
        self.assertEqual(r["amounts"]["actual_received_amount"], "100.00")
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "60.00")
        self.assertEqual(r["margin"]["percent"], "60.00%")

    # ---------- 口径配置接口 ----------
    def test_03_config_validation(self):
        """口径配置：非法枚举 400(code=4000)；合法保存后 config/options 反映新值"""
        status, resp = self._post("/api/profit/config",
                                  {"refund_deduction_status": "bogus"})
        self.assertEqual(status, 400)
        self.assertEqual(resp["code"], 4000)
        self.assertIsNone(resp["data"])

        status, resp = self._post("/api/profit/config",
                                  {"platform_subsidy_included": "bogus"})
        self.assertEqual(status, 400)
        self.assertEqual(resp["code"], 4000)

        status, resp = self._post("/api/profit/config", {})
        self.assertEqual(status, 400)
        self.assertEqual(resp["code"], 4000)

        # 合法保存：退款口径=未扣除
        status, resp = self._post("/api/profit/config",
                                  {"refund_deduction_status": "not_deducted"})
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        self.assertEqual(resp["data"]["policies"]["refund_deduction_status"],
                         "not_deducted")
        self.assertEqual(resp["data"]["policies"]["platform_subsidy_included"],
                         "unknown")
        self.assertEqual(resp["data"]["rules_version"], "P1.1")
        self.assertTrue(resp["data"]["declaration"])

        # options 反映新口径
        status, resp = self._get("/api/profit/options")
        self.assertEqual(status, 200)
        cfg = resp["data"]["config"]
        self.assertEqual(cfg["policies"]["refund_deduction_status"], "not_deducted")
        self.assertIn("未扣除", cfg["labels"]["refund_deduction_status"])
        self.assertEqual(resp["data"]["rules_version"], "P1.1")
        # 枚举选项完整
        self.assertEqual(set(resp["data"]["refund_deduction_status"]),
                         {"unknown", "not_deducted", "deducted"})
        self.assertEqual(set(resp["data"]["platform_subsidy_included"]),
                         {"unknown", "not_included", "included"})

    # ---------- 净收入公式：not_deducted ----------
    def test_04_net_revenue_not_deducted(self):
        """净收入 = 实收 − 成功退款（not_deducted）；处理中退款不计入；全额退款净收入为 0"""
        data = self._recalc()
        # 待确认异常（退款口径）自动关闭，留痕 system
        self.assertEqual(data["anomalies_auto_resolved"], 2)
        # profit_total = 49 + 37 − 27 + 60 = 119.00
        self.assertEqual(self._stats_of(data), {
            "total": 7, "calculated": 3, "pending_confirmation": 1,
            "not_calculable": 2, "needs_review": 1, "profit_total": "119.00"})

        # P1002：净收入 = 80 − 10 = 70（处理中 5 不计入成功退款）
        r = self._profit_of("P1002")
        self.assertEqual(r["profit_status"], "calculated")
        self.assertEqual(r["status_reasons"], [])
        self.assertEqual(r["amounts"]["net_revenue"], "70.00")
        self.assertEqual(r["amounts"]["refund_success_amount"], "10.00")
        self.assertEqual(r["amounts"]["refund_pending_amount"], "5.00")
        # 利润 = 70 − 30 − 3 = 37.00
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "37.00")
        self.assertEqual(r["formula"]["net_revenue"],
                         "口径=未扣除：净收入 = 实收 80.00"
                         " − 成功退款 10.00（处理中退款不计入）= 70.00")
        # 口径确认后异常关闭且留痕
        a = self._find_anomaly("P1002", "REFUND_DEDUCTION_UNCONFIRMED")[0]
        self.assertEqual(a["status"], "resolved")
        self.assertEqual(a["resolution"]["resolved_by"], "system")
        self.assertIn("未再出现", a["resolution"]["text"])

        # P1003：全额退款 → 净收入 0 → 利润率不计算并说明原因
        r = self._profit_of("P1003")
        self.assertEqual(r["profit_status"], "calculated")
        self.assertEqual(r["amounts"]["net_revenue"], "0.00")
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "-27.00")
        self.assertIsNone(r["margin"]["percent"])
        self.assertEqual(r["margin"]["blocked_reason"], "NET_REVENUE_NOT_POSITIVE")
        self.assertEqual(r["margin"]["blocked_reason_label"],
                         "净收入不大于零，不计算常规利润率")

        # P1001：无退款订单净收入 = 实收；补贴口径仍 unknown 不计入 → 利润 49
        r = self._profit_of("P1001")
        self.assertEqual(r["amounts"]["net_revenue"], "100.00")
        self.assertEqual(r["amounts"]["platform_subsidy_credited"], "0.00")
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "49.00")

    # ---------- 补贴规则：not_included 计入 ----------
    def test_05_subsidy_not_included(self):
        """补贴 not_included：计入利润 55.00、利润率 55.00%；计入/未计入分项"""
        status, resp = self._post("/api/profit/config",
                                  {"platform_subsidy_included": "not_included"})
        self.assertEqual(status, 200)
        data = self._recalc()

        # P1001：利润 = 100 + 6 − 48 − 3 = 55.00；利润率 = 55/100 = 55.00%
        r = self._profit_of("P1001")
        self.assertEqual(r["amounts"]["platform_subsidy_credited"], "6.00")
        self.assertEqual(r["amounts"]["platform_subsidy_excluded"], "0.00")
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "55.00")
        self.assertEqual(r["margin"]["percent"], "55.00%")
        self.assertEqual(r["formula"]["profit"],
                         "净收入 100.00 + 计入补贴 6.00"
                         " − 供应商综合总金额 48.00 − 技术服务费 3.00"
                         " − 其他经营费用 0.00 = 55.00")
        self.assertEqual(r["formula"]["margin"],
                         "55.00 ÷ 净收入 100.00 = 55.00%")
        # 负利润异常仍开（P1003 利润 −27）
        self.assertEqual(len(self._find_anomaly("P1003", "NEGATIVE_PROFIT")), 1)

        # P1002：利润率 = 37/70 = 52.86%（四舍五入两位）
        r = self._profit_of("P1002")
        self.assertEqual(r["margin"]["percent"], "52.86%")
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "37.00")

        # P1003：补贴为 0，全额退款净收入 0 仍阻断利润率
        r = self._profit_of("P1003")
        self.assertEqual(r["amounts"]["platform_subsidy_credited"], "0.00")
        self.assertIsNone(r["margin"]["percent"])

    # ---------- 经营总览 ----------
    def test_06_summary_dashboard(self):
        """经营总览：各指标 = 订单明细同口径求和；净收入≤0 利润率 null；空范围 has_orders=false"""
        # ---- 全量总览（not_deducted + not_included）----
        status, resp = self._get("/api/summary")
        self.assertEqual(status, 200)
        s = resp["data"]
        self.assertTrue(s["has_orders"])
        self.assertEqual(s["order_count"], 7)
        self.assertEqual(s["refund_order_count"], 2)
        # 550 = 100+40+40+60+50+90+70+50+50
        self.assertEqual(s["product_sales_amount"], "550.00")
        self.assertEqual(s["actual_paid_amount"], "550.00")
        self.assertEqual(s["actual_received_amount"], "550.00")
        self.assertEqual(s["refund_success_amount"], "70.00")
        self.assertEqual(s["refund_pending_amount"], "5.00")
        # 70/550 = 12.73%
        self.assertEqual(s["refund_rate"], "12.73%")
        # 191 = 48+30+25+48+40
        self.assertEqual(s["supplier_total_cost"], "191.00")
        self.assertEqual(s["technical_service_fee"], "8.00")
        # 125 = 55+37−27+60（calculated+needs_review 同口径求和）
        self.assertEqual(s["operating_contribution_profit"], "125.00")
        self.assertEqual(s["net_revenue_total"], "270.00")
        # 125/270 = 46.30%
        self.assertEqual(s["profit_margin"], "46.30%")
        self.assertIsNone(s["profit_margin_blocked_reason"])
        self.assertEqual(s["profit_status_counts"], {
            "calculated": 3, "pending_confirmation": 1,
            "not_calculable": 2, "needs_review": 1})
        self.assertEqual(s["calculable_order_count"], 4)
        # 4/7 = 57.14%
        self.assertEqual(s["profit_coverage"], "57.14%")
        self.assertEqual(s["unmatched_order_count"], 1)
        # 未决异常 5 个 = 导入/对账类 4（P1007 实收冲突组检查、P1004 未匹配、
        # P1005 重复成本、P1006 成本缺失）+ 利润类 1（P1003 负利润）
        self.assertEqual(s["open_anomaly_count"], 5)
        self.assertTrue(s["has_calculation"])
        self.assertEqual(s["latest_calculation"]["rules_version"], "P1.1")
        self.assertEqual(s["latest_calculation"]["calculation_id"],
                         self.latest_calc_id)
        self.assertEqual(s["policies"], {
            "refund_deduction_status": "not_deducted",
            "platform_subsidy_included": "not_included"})
        self.assertIn("口径声明", s["declaration"])

        # ---- 单订单日（仅 P1003）：净收入合计 0 → 利润率 null ----
        status, resp = self._get("/api/summary?start=2026-03-05&end=2026-03-05")
        self.assertEqual(status, 200)
        s = resp["data"]
        self.assertTrue(s["has_orders"])
        self.assertEqual(s["order_count"], 1)
        self.assertEqual(s["refund_order_count"], 1)
        self.assertEqual(s["actual_received_amount"], "60.00")
        self.assertEqual(s["refund_success_amount"], "60.00")
        self.assertEqual(s["refund_rate"], "100.00%")
        self.assertEqual(s["supplier_total_cost"], "25.00")
        self.assertEqual(s["operating_contribution_profit"], "-27.00")
        self.assertEqual(s["net_revenue_total"], "0.00")
        self.assertIsNone(s["profit_margin"])
        self.assertEqual(s["profit_margin_blocked_reason"],
                         "NET_REVENUE_NOT_POSITIVE")
        self.assertEqual(s["unmatched_order_count"], 0)

        # ---- 空日期范围：无订单 ----
        status, resp = self._get("/api/summary?start=2020-01-01&end=2020-01-02")
        self.assertEqual(status, 200)
        self.assertFalse(resp["data"]["has_orders"])

    # ---------- 净收入公式：deducted 不重复扣 ----------
    def test_07_refund_deducted_no_double(self):
        """deducted：净收入 = 实收（不重复扣成功退款）"""
        status, resp = self._post("/api/profit/config",
                                  {"refund_deduction_status": "deducted"})
        self.assertEqual(status, 200)
        data = self._recalc()

        # P1002：净收入 = 80（不再扣 10），利润 = 80 − 30 − 3 = 47.00
        r = self._profit_of("P1002")
        self.assertEqual(r["profit_status"], "calculated")
        self.assertEqual(r["amounts"]["net_revenue"], "80.00")
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "47.00")
        self.assertEqual(r["margin"]["percent"], "58.75%")  # 47/80
        self.assertEqual(r["formula"]["net_revenue"],
                         "口径=已扣除：净收入 = 实收 80.00（不重复扣减退款）= 80.00")
        # 成功退款金额仍如实展示
        self.assertEqual(r["amounts"]["refund_success_amount"], "10.00")

        # P1003：净收入 = 60，利润 = 60 − 25 − 2 = 33.00（负利润异常自动关闭）
        r = self._profit_of("P1003")
        self.assertEqual(r["amounts"]["net_revenue"], "60.00")
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "33.00")
        self.assertEqual(r["margin"]["percent"], "55.00%")  # 33/60
        # 负利润异常因不再出现而自动关闭
        a = self._find_anomaly("P1003", "NEGATIVE_PROFIT")[0]
        self.assertEqual(a["status"], "resolved")
        self.assertEqual(data["anomalies_auto_resolved"], 1)

        # profit_total = 55 + 47 + 33 + 60 = 195.00
        self.assertEqual(data["profit_total"], "195.00")

    # ---------- 补贴规则：included 不重复计入 ----------
    def test_08_subsidy_included_no_double(self):
        """included：补贴已含于实收，不重复加；单列展示"""
        status, resp = self._post("/api/profit/config",
                                  {"platform_subsidy_included": "included"})
        self.assertEqual(status, 200)
        self._recalc()

        r = self._profit_of("P1001")
        self.assertEqual(r["amounts"]["platform_subsidy_total"], "6.00")
        self.assertEqual(r["amounts"]["platform_subsidy_credited"], "0.00")
        self.assertEqual(r["amounts"]["platform_subsidy_excluded"], "6.00")
        # 利润 = 100 + 0 − 48 − 3 = 49.00（不重复加补贴）
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "49.00")
        self.assertEqual(r["margin"]["percent"], "49.00%")

    # ---------- 重复成本风险：人工确认粒度后可计算 ----------
    def test_09_duplicate_risk_manual_confirm(self):
        """同额多行 dedup_pending → 人工确认 order_level → 重跑对账 → 重算 calculated"""
        recon = self._recon_of("P1005")
        sids = recon["supplier_record_ids"]
        self.assertEqual(len(sids), 3)
        self.assertTrue(recon["cost"]["order_level_dedup_pending"])

        # 确认为订单级综合成本：只扣一次
        status, resp = self._post("/api/reconciliation/manual-match", {
            "order_id": "P1005", "supplier_record_ids": sids,
            "cost_granularity": "order_level", "note": "确认为订单级综合成本"})
        self.assertEqual(status, 200)
        rec = resp["data"]
        self.assertEqual(rec["cost"]["confirmed_granularity"], "order_level")
        self.assertEqual(rec["cost"]["total_amount"], "48.00")
        self.assertTrue(rec["cost"]["usable_for_profit"])
        self.assertFalse(rec["cost"]["order_level_dedup_pending"])

        # 重跑对账：人工确认保留
        self._run_recon()
        recon = self._recon_of("P1005")
        self.assertEqual(recon["match_status"], "manual_confirmed")
        self.assertEqual(recon["cost"]["total_amount"], "48.00")
        self.assertTrue(recon["cost"]["usable_for_profit"])

        # 重算利润：P1005 变为 calculated，利润 = 90 − 48 = 42.00
        data = self._recalc()
        r = self._profit_of("P1005")
        self.assertEqual(r["profit_status"], "calculated")
        self.assertEqual(r["status_reasons"], [])
        self.assertEqual(r["amounts"]["supplier_total_amount"], "48.00")
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "42.00")
        self.assertEqual(r["margin"]["percent"], "46.67%")  # 42/90
        self.assertEqual(r["confirmed_granularity"], "order_level")
        # 统计：calculated=4（P1001/P1002/P1003/P1005），无待确认
        self.assertEqual(self._stats_of(data), {
            "total": 7, "calculated": 4, "pending_confirmation": 0,
            "not_calculable": 2, "needs_review": 1, "profit_total": "231.00"})
        # 未决异常仅剩导入/对账类 4 个（利润类异常已全部自动关闭）
        status, resp = self._get("/api/summary")
        self.assertEqual(resp["data"]["open_anomaly_count"], 4)

    # ---------- 追溯与版本 ----------
    def test_10_trace_and_versioning(self):
        """核算记录规则版本/数据批次/时间；重算生成新版本且历史保留可查"""
        old_calc_id = self.latest_calc_id
        data = self._recalc()
        new_calc_id = data["calculation_id"]
        self.assertNotEqual(new_calc_id, old_calc_id)

        # 订单明细含 trace：微信批次 + 供应商批次 + 规则版本
        r = self._profit_of("P1001")
        self.assertEqual(r["rules_version"], "P1.1")
        self.assertEqual(r["calculation_id"], new_calc_id)
        self.assertTrue(r["created_at"])
        self.assertEqual(len(r["trace"]["batch_ids"]["wechat_orders"]), 1)
        self.assertEqual(len(r["trace"]["batch_ids"]["supplier_costs"]), 1)
        self.assertEqual(len(r["trace"]["order_record_ids"]), 1)
        self.assertEqual(len(r["trace"]["supplier_record_ids"]), 1)

        # 历史版本保留：核算运行记录数 = 累计 recalculate 次数（7 次）
        status, resp = self._get("/api/data/profit_statements")
        self.assertEqual(status, 200)
        records = resp["data"]["records"]
        runs = [x for x in records if x.get("record_type") == "calculation_run"]
        self.assertEqual(len(runs), 7)
        calc_ids = {x["record_id"] for x in runs}
        self.assertEqual(len(calc_ids), 7)
        self.assertIn(self.first_calc_id, calc_ids)
        # 运行记录含规则版本、口径与数据批次
        first_run = next(x for x in runs if x["record_id"] == self.first_calc_id)
        self.assertEqual(first_run["rules_version"], "P1.1")
        self.assertEqual(first_run["policies"]["refund_deduction_status"],
                         "unknown")

        # 历史版本明细仍可按 calc_id 查询（禁止覆盖历史）
        r = self._profit_of("P1001", calc_id=self.first_calc_id)
        self.assertEqual(r["calculation_id"], self.first_calc_id)
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "49.00")
        self.assertEqual(r["profit_status"], "calculated")
        # 历史版本保留当时的口径快照
        self.assertEqual(r["policies"]["refund_deduction_status"], "unknown")
        self.assertEqual(r["policies"]["platform_subsidy_included"], "unknown")

        # 默认查询返回最新版本
        status, resp = self._get("/api/profit?order_id=P1001")
        self.assertEqual(resp["data"]["calculation_id"], new_calc_id)

    # ---------- 导出 ----------
    def test_11_export(self):
        """导出与筛选一致；空数据 400；导出留痕；三新列与合计行"""
        # 全量：表头 + 7 行 + 合计行
        status, content, headers = self._get_binary("/api/profit/export")
        self.assertEqual(status, 200)
        wb = load_workbook(io.BytesIO(content))
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        self.assertEqual(len(rows), 9)
        self.assertEqual(rows[0][0], "订单号")
        self.assertIn("经营贡献利润", rows[0])
        self.assertIn("综合总金额(唯一扣减)", rows[0])
        # 三新列位于表尾（"供应商批次"之后），既有列零位移
        self.assertIn("产品成本", rows[0])
        self.assertIn("快递费合计", rows[0])
        self.assertIn("综合成本", rows[0])
        self.assertLess(rows[0].index("供应商批次"),
                        rows[0].index("产品成本"))
        self.assertLess(rows[0].index("产品成本"),
                        rows[0].index("快递费合计"))
        self.assertLess(rows[0].index("快递费合计"),
                        rows[0].index("综合成本"))
        self.assertEqual({rows[i][0] for i in range(1, 8)},
                         {"P1001", "P1002", "P1003", "P1004",
                          "P1005", "P1006", "P1007"})
        # P1001 行：实收 100 / 利润 49 / 利润率 49.00%
        p1 = next(r for r in rows[1:] if r[0] == "P1001")
        self.assertEqual(float(p1[4]), 100.00)
        self.assertEqual(float(p1[14]), 49.00)
        self.assertEqual(p1[15], "49.00%")
        self.assertIn("attachment", headers.get("Content-Disposition", ""))

        # 合计行三值 = 统计卡同源同值（数值类型，非字符串）
        st, pr = self._get("/api/profit")
        self.assertEqual(st, 200)
        stat = pr["data"]["statistics"]
        summary = rows[-1]
        self.assertEqual(summary[0], "合计")
        self.assertEqual(summary[1], "共 7 单")
        self.assertIsInstance(summary[22], (int, float))
        self.assertIsInstance(summary[23], (int, float))
        self.assertIsInstance(summary[24], (int, float))
        self.assertEqual(float(summary[22]),
                         float(stat["product_cost_total"]))
        self.assertEqual(float(summary[23]),
                         float(stat["courier_fee_total"]))
        self.assertEqual(float(summary[24]),
                         float(stat["composite_cost_total"]))

        # 按状态筛选：calculated 4 行 + 合计行
        status, content, _ = self._get_binary(
            "/api/profit/export?profit_status=calculated")
        self.assertEqual(status, 200)
        wb = load_workbook(io.BytesIO(content))
        rows = list(wb.active.iter_rows(values_only=True))
        self.assertEqual(len(rows), 6)
        self.assertEqual({rows[i][0] for i in range(1, 5)},
                         {"P1001", "P1002", "P1003", "P1005"})
        # 合计行随筛选同口径收窄
        st, pr = self._get("/api/profit?profit_status=calculated")
        self.assertEqual(st, 200)
        stat = pr["data"]["statistics"]
        summary = rows[-1]
        self.assertEqual(summary[0], "合计")
        self.assertEqual(summary[1], "共 4 单")
        self.assertEqual(float(summary[24]),
                         float(stat["composite_cost_total"]))

        # 空数据导出 → 400
        status, content, _ = self._get_binary("/api/profit/export?order_id=NOPE")
        self.assertEqual(status, 400)
        body = json.loads(content.decode("utf-8"))
        self.assertEqual(body["code"], 4000)
        self.assertIsNone(body["data"])

        # 导出留痕
        status, resp = self._get("/api/data/export_records")
        self.assertTrue(any(e["kind"] == "profit"
                            for e in resp["data"]["records"]))

    # ---------- 明细筛选与分页 ----------
    def test_12_list_filters(self):
        """利润明细筛选（状态/订单号/日期）与非法分页参数"""
        # 按状态
        status, resp = self._get("/api/profit?profit_status=not_calculable&size=100")
        self.assertEqual(status, 200)
        ids = {r["order_id"] for r in resp["data"]["results"]}
        self.assertEqual(ids, {"P1004", "P1006"})
        self.assertEqual(resp["data"]["statistics"]["total"], 2)

        # 按订单号
        r = self._profit_of("P1001")
        self.assertEqual(r["order_id"], "P1001")

        # 按日期（2026-03-02 仅 P1004/P1005/P1006）
        status, resp = self._get("/api/profit?start=2026-03-02&end=2026-03-02&size=100")
        self.assertEqual(status, 200)
        ids = {r["order_id"] for r in resp["data"]["results"]}
        self.assertEqual(ids, {"P1004", "P1005", "P1006"})

        # 默认返回最新核算版本
        status, resp = self._get("/api/profit?size=100")
        self.assertEqual(resp["data"]["calculation_id"], self.latest_calc_id)
        self.assertEqual(resp["data"]["total"], 7)

        # 非法分页
        status, resp = self._get("/api/profit?size=7")
        self.assertEqual(status, 400)
        status, resp = self._get("/api/profit?page=0")
        self.assertEqual(status, 400)

    # ---------- 缺失数据与日期范围 ----------
    def test_13_missing_data_and_date_scope(self):
        """实收缺失 → 不可计算（RECEIVED_AMOUNT_MISSING，禁止按 0）；日期范围排除无时间订单"""
        # 注入两条边界订单：P1008 实收缺失、P1009 无下单时间
        status, resp = self._get("/api/data/orders")
        self.assertEqual(status, 200)
        records = resp["data"]["records"]
        base = next(r for r in records if r["order_id"] == "P1004")
        p1008 = json.loads(json.dumps(base))
        p1008["record_id"] = "wx_test_p1008"
        p1008["order_id"] = "P1008"
        p1008["order_time"] = "2026-03-04 10:00:00"
        p1008["amounts"]["actual_received_amount"] = None
        p1009 = json.loads(json.dumps(base))
        p1009["record_id"] = "wx_test_p1009"
        p1009["order_id"] = "P1009"
        p1009["order_time"] = None
        p1009["amounts"]["actual_received_amount"] = "30.00"
        records.extend([p1008, p1009])
        status, resp = self._post("/api/data/orders", {"records": records})
        self.assertEqual(status, 200)

        # 未重跑对账：总览按对账结果统计未匹配（无对账记录的订单计入）
        status, resp = self._get("/api/summary")
        self.assertEqual(status, 200)
        s = resp["data"]
        self.assertEqual(s["order_count"], 9)
        # P1004（对账未匹配）+ P1008/P1009（尚无对账记录）
        self.assertEqual(s["unmatched_order_count"], 3)
        # 实收合计：550 + P1009 的 30（P1008 缺失不计，禁止按 0）
        self.assertEqual(s["actual_received_amount"], "580.00")

        # 重跑对账 + 全范围核算
        self._run_recon()
        data = self._recalc()
        # P1008：实收缺失 → 不可计算
        r = self._profit_of("P1008")
        self.assertEqual(r["profit_status"], "not_calculable")
        self.assertEqual(r["status_reasons"], ["RECEIVED_AMOUNT_MISSING"])
        self.assertIsNone(r["amounts"]["actual_received_amount"])
        self.assertIsNone(r["amounts"]["net_revenue"])
        self.assertIsNone(r["amounts"]["operating_contribution_profit"])
        self.assertEqual(r["margin"]["blocked_reason"], "NOT_CALCULABLE")
        self.assertEqual(r["formula"]["net_revenue"],
                         "实际收款金额缺失，净收入不可计算")
        # P1009：无时间订单全范围核算包含，未匹配 → 不可计算
        r = self._profit_of("P1009")
        self.assertEqual(r["profit_status"], "not_calculable")
        self.assertEqual(r["status_reasons"], ["UNMATCHED_NO_COST"])
        self.assertEqual(r["amounts"]["net_revenue"], "30.00")
        self.assertEqual(self._stats_of(data), {
            "total": 9, "calculated": 4, "pending_confirmation": 0,
            "not_calculable": 4, "needs_review": 1, "profit_total": "231.00"})

        # 日期范围核算：无时间订单（P1009）被排除
        time.sleep(1.05)
        status, resp = self._post("/api/profit/recalculate",
                                  {"start": "2026-03-01", "end": "2026-03-05"})
        self.assertEqual(status, 200)
        type(self).latest_calc_id = resp["data"]["calculation_id"]
        self.assertEqual(resp["data"]["total"], 8)
        status, resp = self._get("/api/profit?order_id=P1009&size=100")
        self.assertEqual(resp["data"]["total"], 0)
        # 总览同样排除无时间订单
        status, resp = self._get("/api/summary?start=2026-03-01&end=2026-03-05")
        self.assertEqual(resp["data"]["order_count"], 8)

    # ---------- 疑似匹配：利润待确认 ----------
    def test_14_suspected_match_pending(self):
        """疑似匹配（辅助线索）→ 利润待确认（SUSPECTED_MATCH_NOT_CONFIRMED）；end 范围排除无时间订单"""
        # 注入辅助候选成本：SKU+数量命中 P1004，但主订单编号不精确匹配
        status, resp = self._get("/api/data/supplier_costs")
        self.assertEqual(status, 200)
        records = resp["data"]["records"]
        base = next(r for r in records if r.get("main_order_id") == "P1006")
        sup = json.loads(json.dumps(base))
        sup["record_id"] = "sup_test_888"
        sup["main_order_id"] = "SUP-888"
        sup["supplier_order_id"] = None
        sup["created_at"] = "2026-03-02 10:05:00"
        sup["product"]["sku_code"] = "SKU-D"
        sup["product"]["quantity"] = 1
        sup["cost_breakdown"]["total_amount"] = "20.00"
        sup["cost_breakdown"]["product_supply_cost"] = None
        records.append(sup)
        status, resp = self._post("/api/data/supplier_costs",
                                  {"records": records})
        self.assertEqual(status, 200)

        # 重跑对账：P1004 通过辅助线索成为疑似匹配（禁止自动认定为已匹配）
        data = self._run_recon()
        self.assertEqual(data["summary"]["suspected"], 1)
        self.assertEqual(data["summary"]["matched"], 5)
        recon = self._recon_of("P1004")
        self.assertEqual(recon["match_status"], "suspected")
        self.assertEqual(recon["match_method"], "assisted_candidate")

        # 重算（仅 end 范围）：疑似匹配利润待确认；无时间订单 P1009 被排除
        time.sleep(1.05)
        status, resp = self._post("/api/profit/recalculate",
                                  {"end": "2026-03-05"})
        self.assertEqual(status, 200)
        data = resp["data"]
        type(self).latest_calc_id = data["calculation_id"]
        self.assertEqual(data["total"], 8)
        self.assertEqual(data["pending_confirmation"], 1)

        # P1004：待确认 + 疑似匹配原因；参考利润 = 50 − 20 − 0 = 30.00（不计入合计）
        r = self._profit_of("P1004")
        self.assertEqual(r["profit_status"], "pending_confirmation")
        self.assertEqual(r["status_reasons"], ["SUSPECTED_MATCH_NOT_CONFIRMED"])
        self.assertEqual(r["amounts"]["supplier_total_amount"], "20.00")
        self.assertEqual(r["amounts"]["operating_contribution_profit"], "30.00")
        # profit_total 仍为 49+47+33+42+60 = 231.00（待确认订单不计入）
        self.assertEqual(data["profit_total"], "231.00")


if __name__ == "__main__":
    unittest.main()

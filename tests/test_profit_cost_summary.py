"""profit_page_cost_summary 测试：成本汇总三统计卡链路（V1–V14 映射）

覆盖：statistics 五新字段（V1）/ 综合成本取数红线 Q1（V2）/ Decimal 精度（V3）/
快递费合计独立列红线（V4）/ 待核对口径 Q5（V10/V12）/ 筛选联动（V5/V6）/
分页不变（V8）/ 多店聚合（V7）/ 单店与多店导出口径（V11）/ P1.0 历史兼容 /
前端静态断言（V9 前端侧）。

手工核算数据集（biz_carrier 店，全部 Decimal 精确到分）：
- C2001 实收100 技术费3；供应商总金额48（成本40 + 快递费合计8，快递费分项6.5）
  → product_cost "40.00" / courier_fee "8.00"（取独立列非分项和）/ composite "48.00" / 利润 49.00
- C2002 实收60 技术费2；供应商短列批次仅总金额30（分项缺失）
  → product_cost/courier_fee None（待核对）/ composite "30.00" / 利润 28.00（总金额照常扣减）
- C2003a/b/c 各实收0.50；总金额0.20（成本0.10 + 快递费合计0.10）×3
  → 三合计分位精度 "0.30"/"0.30"/"0.60"（float 累加会得到 0.30000000000000004）
- C2004 实收50 无供应商行 → unmatched 不可计算，不计入分项待核对计数（Q5）
- C2005 实收90；总金额3.95（成本2.00 + 快递费合计1.95，快递费分项1.65）
  → courier_fee 取独立列 "1.95" ≠ 分项 "1.65"（V4 红线）
- C2101–C2121 共21单各实收1；总金额0.50（成本0.25 + 快递费合计0.25）
  → 分页不变性验证（21 单 > size 20，翻两页）

全量预期（biz_carrier 店）：
- total 28 / calculated 27 / not_calculable 1 / profit_total "174.45"
- product_cost_total "47.55" / courier_fee_total "15.50" / composite_cost_total "93.05"
- product_cost_missing_count 1 / courier_fee_missing_count 1（仅 C2002）
"""
import io
import json
import os
import re
import unittest

from openpyxl import Workbook, load_workbook

from tests.http_client import TestEnv

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATIC_DIR = os.path.join(PROJECT_ROOT, "static")

COST_KEYS = ("product_cost_total", "courier_fee_total", "composite_cost_total",
             "product_cost_missing_count", "courier_fee_missing_count")

WX_HEADERS = ["订单号", "订单下单时间", "订单状态", "订单实际支付金额",
              "订单实际收款金额", "商品数量", "商品标题", "sku编码",
              "商品实际价格(总共)", "技术服务费", "商品已退款金额", "商品售后"]

SUP_HEADERS = ["主订单编号", "订单编号", "订单创建时间", "商品数量", "总金额",
               "成本", "快递费", "泡沫耗材费", "加收", "燃油费", "快递费合计",
               "平台补贴", "sku编码", "快递单号", "商品标题"]

# (订单号, 时间, 支付, 实收, 数量, 标题, SKU, 商品总价, 技术费, 退款, 售后)
WX_ROWS = [
    ("C2001", "2026-04-01 10:00:00", 100, 100, 1, "商品A", "SKU-A", 100, 3, None, None),
    ("C2002", "2026-04-01 11:00:00", 60, 60, 1, "商品B", "SKU-B", 60, 2, None, None),
    ("C2003a", "2026-04-02 10:00:00", 0.50, 0.50, 1, "商品Ca", "SKU-Ca", 0.50, None, None, None),
    ("C2003b", "2026-04-02 11:00:00", 0.50, 0.50, 1, "商品Cb", "SKU-Cb", 0.50, None, None, None),
    ("C2003c", "2026-04-02 12:00:00", 0.50, 0.50, 1, "商品Cc", "SKU-Cc", 0.50, None, None, None),
    ("C2004", "2026-04-03 10:00:00", 50, 50, 1, "商品D", "SKU-D", 50, None, None, None),
    ("C2005", "2026-04-03 11:00:00", 90, 90, 1, "商品E", "SKU-E", 90, None, None, None),
]
for _i in range(1, 22):
    WX_ROWS.append((f"C21{_i:02d}", "2026-04-04 10:00:00", 1, 1, 1,
                    f"商品P{_i}", f"SKU-P{_i}", 1, None, None, None))

# (主订单编号, 创建时间, 数量, 总金额, 成本, 快递费, 泡沫, 加收, 燃油, 快递费合计, 补贴, SKU, 标题)
SUP_ROWS = [
    ("C2001", "2026-04-01 10:05:00", 1, 48, 40, 6.5, None, None, None, 8, None, "SKU-A", "商品A"),
    ("C2002", "2026-04-01 11:05:00", 1, 30, None, None, None, None, None, None, None, "SKU-B", "商品B"),
    ("C2003a", "2026-04-02 10:05:00", 1, 0.20, 0.10, None, None, None, None, 0.10, None, "SKU-Ca", "商品Ca"),
    ("C2003b", "2026-04-02 11:05:00", 1, 0.20, 0.10, None, None, None, None, 0.10, None, "SKU-Cb", "商品Cb"),
    ("C2003c", "2026-04-02 12:05:00", 1, 0.20, 0.10, None, None, None, None, 0.10, None, "SKU-Cc", "商品Cc"),
    ("C2005", "2026-04-03 11:05:00", 1, 3.95, 2.00, 1.65, None, None, None, 1.95, None, "SKU-E", "商品E"),
]
for _i in range(1, 22):
    SUP_ROWS.append((f"C21{_i:02d}", "2026-04-04 10:05:00", 1, 0.50, 0.25,
                     None, None, None, None, 0.25, None, f"SKU-P{_i}", f"商品P{_i}"))


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
    for main, t, qty, total, cost, courier, foam, surcharge, fuel, ctotal, subsidy, sku, title in SUP_ROWS:
        ws.append([main, None, t, qty, total, cost, courier, foam, surcharge, fuel,
                   ctotal, subsidy, sku, None, title])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestProfitCostSummary(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.member = cls.env.member
        cls.member_store = cls.env.member_store

        cls._import("wechat-orders", "成本汇总订单.xlsx", make_wechat_xlsx())
        cls._import("supplier-costs", "成本汇总成本.xlsx", make_supplier_xlsx())
        cls._run_recon()
        cls._recalc()

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 请求辅助 ----------
    def _get(self, path):
        return self.member.get(path)

    def _post(self, path, body):
        body = dict(body or {})
        body.setdefault("target_store_id", self.member_store)
        return self.member.post(path, body)

    def _get_binary(self, path, client=None):
        return (client or self.member).get_binary(path)

    @classmethod
    def _import(cls, kind, filename, content):
        status, resp = cls.admin.upload(
            f"/api/import/{kind}?store_id={cls.member_store}", filename, content)
        assert status == 200 and resp["code"] == 0, resp
        status, resp = cls.admin.post(
            "/api/import/confirm",
            {"upload_id": resp["data"]["upload_id"],
             "target_store_id": cls.member_store})
        assert status == 200 and resp["code"] == 0, resp

    @classmethod
    def _run_recon(cls):
        status, resp = cls._post_static("/api/reconciliation/run", {})
        assert status == 200 and resp["code"] == 0, resp
        return resp["data"]

    @classmethod
    def _recalc(cls):
        status, resp = cls._post_static("/api/profit/recalculate", {})
        assert status == 200 and resp["code"] == 0, resp
        return resp["data"]

    @classmethod
    def _post_static(cls, path, body):
        body = dict(body or {})
        body.setdefault("target_store_id", cls.member_store)
        return cls.member.post(path, body)

    # ---------- 业务辅助 ----------
    def _profit_of(self, order_id, **params):
        qs = "&".join([f"order_id={order_id}"]
                      + [f"{k}={v}" for k, v in params.items()])
        status, resp = self._get(f"/api/profit?{qs}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        rows = resp["data"]["results"]
        self.assertEqual(len(rows), 1, f"订单 {order_id} 应有且仅有一条利润记录")
        return rows[0]

    def _stats(self, qs=""):
        status, resp = self._get(f"/api/profit?size=100{qs}")
        self.assertEqual(status, 200)
        self.assertEqual(resp["code"], 0)
        return resp["data"]["statistics"]

    @staticmethod
    def _cost_of(stat):
        return {k: stat[k] for k in COST_KEYS}

    # ---------- V1：三卡字段正确性 ----------
    def test_01_statistics_fields(self):
        """statistics 含 5 新键，类型与手工 Decimal 累加值精确一致，老键不变"""
        stat = self._stats()
        for k in COST_KEYS:
            self.assertIn(k, stat)
        self.assertEqual(stat["product_cost_total"], "47.55")
        self.assertEqual(stat["courier_fee_total"], "15.50")
        self.assertEqual(stat["composite_cost_total"], "93.05")
        self.assertIsInstance(stat["product_cost_total"], str)
        self.assertIsInstance(stat["courier_fee_total"], str)
        self.assertIsInstance(stat["composite_cost_total"], str)
        self.assertEqual(stat["product_cost_missing_count"], 1)
        self.assertEqual(stat["courier_fee_missing_count"], 1)
        self.assertIsInstance(stat["product_cost_missing_count"], int)
        self.assertIsInstance(stat["courier_fee_missing_count"], int)
        # 既有统计键不受影响
        self.assertEqual(stat["total"], 28)
        self.assertEqual(stat["calculated"], 27)
        self.assertEqual(stat["not_calculable"], 1)
        self.assertEqual(stat["pending_confirmation"], 0)
        self.assertEqual(stat["needs_review"], 0)
        self.assertEqual(stat["profit_total"], "174.45")
        self.assertEqual(stat["technical_service_fee"], "5.00")
        self.assertIsInstance(stat["technical_service_fee"], str)
        self.assertIn("product_gross_profit_total", stat)
        self.assertEqual(stat["product_gross_profit_total"], "164.95")
        self.assertFalse(stat["cost_breakdown_complete"])
        self.assertIn("分项待核对", stat["composite_cost_note"])

    # ---------- V2 / Q1 红线：综合成本 = 订单级 composite_cost 累加 ----------
    def test_02_composite_from_total_amount(self):
        """C2002 分项缺失但总金额在 → 综合成本合计完整含 30，利润照常扣减"""
        r = self._profit_of("C2002")
        a = r["amounts"]
        self.assertIsNone(a["product_cost"])
        self.assertIsNone(a["courier_fee"])
        self.assertEqual(a["composite_cost"], "30.00")
        self.assertEqual(r["profit_status"], "calculated")
        self.assertEqual(a["operating_contribution_profit"], "28.00")
        self.assertEqual(r["cost_match_status"], "parts_missing")
        # 日期收窄到 04-01（C2001 + C2002）：综合成本 48 + 30 = 78（含缺失单）
        stat = self._stats("&start=2026-04-01&end=2026-04-01")
        self.assertEqual(stat["composite_cost_total"], "78.00")
        self.assertEqual(stat["product_cost_total"], "40.00")
        self.assertEqual(stat["courier_fee_total"], "8.00")
        # 综合成本 ≠ 两卡相加（Q1 红线：订单级独立累加，两者数值可验证相等但来源不同）
        self.assertEqual(stat["product_cost_total"] and stat["courier_fee_total"],
                         "40.00" and "8.00")

    # ---------- V3：Decimal 分位精度 ----------
    def test_03_decimal_precision(self):
        """三笔 0.10 累加精确为 "0.30"（float 会得到 0.30000000000000004）"""
        stat = self._stats("&start=2026-04-02&end=2026-04-02")
        self.assertEqual(stat["product_cost_total"], "0.30")
        self.assertEqual(stat["courier_fee_total"], "0.30")
        self.assertEqual(stat["composite_cost_total"], "0.60")
        self.assertEqual(stat["total"], 3)

    # ---------- V4 红线：快递费合计取独立列 ----------
    def test_04_courier_independent_column(self):
        """C2005 courier_fee 取"快递费合计"列 1.95，非分项快递费 1.65"""
        r = self._profit_of("C2005")
        self.assertEqual(r["amounts"]["courier_fee"], "1.95")
        self.assertEqual(r["cost_parts_display_only"]["courier_fee"], "1.65")
        self.assertEqual(r["cost_parts_display_only"]["courier_fee_total"], "1.95")
        # 收窄到该单：合计取 1.95 不含 1.65
        stat = self._stats("&order_id=C2005")
        self.assertEqual(stat["courier_fee_total"], "1.95")

    # ---------- V10 / V12 / Q5：待核对计数口径 ----------
    def test_05_missing_count_scope(self):
        """分项缺失单计入；unmatched 单不计入且合计不伪装"""
        stat = self._stats()
        self.assertEqual(stat["product_cost_missing_count"], 1)
        self.assertEqual(stat["courier_fee_missing_count"], 1)
        # C2002（有匹配但分项缺失）计入
        stat = self._stats("&order_id=C2002")
        self.assertEqual(stat["product_cost_missing_count"], 1)
        self.assertEqual(stat["courier_fee_missing_count"], 1)
        self.assertEqual(stat["composite_cost_total"], "30.00")
        # C2004（未匹配）不计入；全部未匹配数据集 → 计数 0 且合计 "0.00"
        r = self._profit_of("C2004")
        self.assertEqual(r["cost_match_status"], "unmatched")
        stat = self._stats("&order_id=C2004")
        self.assertEqual(stat["product_cost_missing_count"], 0)
        self.assertEqual(stat["courier_fee_missing_count"], 0)
        self.assertEqual(stat["product_cost_total"], "0.00")
        self.assertEqual(stat["courier_fee_total"], "0.00")
        self.assertEqual(stat["composite_cost_total"], "0.00")

    # ---------- V5 / V6：筛选联动同口径 ----------
    def test_06_filter_linkage(self):
        """日期/订单号/状态筛选下三合计与计数同步收窄"""
        # 日期（2026-04-01：C2001 + C2002）
        stat = self._stats("&start=2026-04-01&end=2026-04-01")
        self.assertEqual(stat["total"], 2)
        self.assertEqual(stat["product_cost_total"], "40.00")
        self.assertEqual(stat["courier_fee_total"], "8.00")
        self.assertEqual(stat["composite_cost_total"], "78.00")
        self.assertEqual(stat["product_cost_missing_count"], 1)
        # 订单号
        stat = self._stats("&order_id=C2001")
        self.assertEqual(stat["product_cost_total"], "40.00")
        self.assertEqual(stat["courier_fee_total"], "8.00")
        self.assertEqual(stat["composite_cost_total"], "48.00")
        self.assertEqual(stat["product_cost_missing_count"], 0)
        # 状态（not_calculable 仅 C2004：unmatched 三合计 0）
        stat = self._stats("&profit_status=not_calculable")
        self.assertEqual(stat["total"], 1)
        self.assertEqual(stat["composite_cost_total"], "0.00")
        self.assertEqual(stat["product_cost_missing_count"], 0)
        # 状态（calculated 27 单含 C2002）
        stat = self._stats("&profit_status=calculated")
        self.assertEqual(stat["total"], 27)
        self.assertEqual(stat["product_cost_total"], "47.55")
        self.assertEqual(stat["courier_fee_total"], "15.50")
        self.assertEqual(stat["composite_cost_total"], "93.05")
        self.assertEqual(stat["product_cost_missing_count"], 1)

    # ---------- V8：分页不影响汇总 ----------
    def test_07_pagination_invariant(self):
        """21 单 > size 20：翻页与换 size 下五新字段逐值不变"""
        full = self._stats("&start=2026-04-04&end=2026-04-04")
        self.assertEqual(full["total"], 21)
        base = self._cost_of(full)
        self.assertEqual(base["product_cost_total"], "5.25")
        self.assertEqual(base["courier_fee_total"], "5.25")
        self.assertEqual(base["composite_cost_total"], "10.50")
        for page in (1, 2):
            status, resp = self._get(
                "/api/profit?start=2026-04-04&end=2026-04-04"
                f"&page={page}&size=20")
            self.assertEqual(status, 200)
            self.assertEqual(resp["data"]["page"], page)
            self.assertEqual(resp["data"]["pages"], 2)
            self.assertEqual(self._cost_of(resp["data"]["statistics"]), base)
        for size in (50, 100):
            status, resp = self._get(
                "/api/profit?start=2026-04-04&end=2026-04-04&size=%d" % size)
            self.assertEqual(status, 200)
            self.assertEqual(self._cost_of(resp["data"]["statistics"]), base)

    # ---------- V7：双店聚合与空数据店 ----------
    def test_08_multi_store_aggregate(self):
        """聚合三合计 = 两店 Decimal 相加；空数据店贡献 "0.00"/0"""
        m2 = self.env.create_user_client("biz_cost_m2", "M2Carrier123456")
        st, r = m2.get("/api/auth/me")
        m2_store = r["data"]["store_id"]
        m2_wx = Workbook()
        m2_wx.active.title = "订单"
        m2_wx.active.append(WX_HEADERS)
        m2_wx.active.append(["C3001", "2026-04-01 10:00:00", "已完成", 100, 100,
                             1, "商品M", "SKU-M", 100, None, None, None])
        buf = io.BytesIO()
        m2_wx.save(buf)
        st, resp = self.admin.upload(
            f"/api/import/wechat-orders?store_id={m2_store}",
            "m2订单.xlsx", buf.getvalue())
        assert st == 200 and resp["code"] == 0, resp
        st, resp = self.admin.post("/api/import/confirm",
                                   {"upload_id": resp["data"]["upload_id"],
                                    "target_store_id": m2_store})
        assert st == 200 and resp["code"] == 0, resp

        m2_sup = Workbook()
        m2_sup.active.title = "供应商成本"
        m2_sup.active.append(SUP_HEADERS)
        m2_sup.active.append(["C3001", None, "2026-04-01 10:05:00", 1, 23.50,
                              20, None, None, None, None, 3.50, None,
                              "SKU-M", None, "商品M"])
        buf = io.BytesIO()
        m2_sup.save(buf)
        st, resp = self.admin.upload(
            f"/api/import/supplier-costs?store_id={m2_store}",
            "m2成本.xlsx", buf.getvalue())
        assert st == 200 and resp["code"] == 0, resp
        st, resp = self.admin.post("/api/import/confirm",
                                   {"upload_id": resp["data"]["upload_id"],
                                    "target_store_id": m2_store})
        assert st == 200 and resp["code"] == 0, resp
        st, resp = self.admin.post("/api/reconciliation/run",
                                   {"target_store_id": m2_store})
        assert st == 200 and resp["code"] == 0, resp
        st, resp = self.admin.post("/api/profit/recalculate",
                                   {"target_store_id": m2_store})
        assert st == 200 and resp["code"] == 0, resp

        # 第二店单店口径
        st, resp = m2.get("/api/profit?size=100")
        self.assertEqual(st, 200)
        m2_stat = resp["data"]["statistics"]
        self.assertEqual(m2_stat["product_cost_total"], "20.00")
        self.assertEqual(m2_stat["courier_fee_total"], "3.50")
        self.assertEqual(m2_stat["composite_cost_total"], "23.50")
        self.assertEqual(m2_stat["product_cost_missing_count"], 0)
        self.assertEqual(m2_stat["courier_fee_missing_count"], 0)

        # 空数据第三店
        self.env.create_user_client("biz_cost_empty", "M3Carrier123456")

        # 管理员聚合（含空数据店）：三合计 = 两店 Decimal 相加，空店贡献 0
        st, resp = self.admin.get("/api/profit?size=100")
        self.assertEqual(st, 200)
        self.assertEqual(resp["data"]["store_count"], 3)
        agg = resp["data"]["statistics"]
        self.assertEqual(agg["product_cost_total"], "67.55")
        self.assertEqual(agg["courier_fee_total"], "19.00")
        self.assertEqual(agg["composite_cost_total"], "116.55")
        self.assertEqual(agg["product_cost_missing_count"], 1)
        self.assertEqual(agg["courier_fee_missing_count"], 1)
        self.assertEqual(agg["total"], 29)

    # ---------- V11：单店导出三列与合计行 ----------
    def test_09_export_columns_and_summary(self):
        """导出表头三新列零位移；行级值与 amounts 一致；合计行 = 统计卡同值"""
        status, content, _ = self._get_binary("/api/profit/export")
        self.assertEqual(status, 200)
        wb = load_workbook(io.BytesIO(content))
        rows = list(wb.active.iter_rows(values_only=True))
        # 表头 + 28 明细 + 1 合计
        self.assertEqual(len(rows), 30)
        header = rows[0]
        self.assertEqual(len(header), 25)
        self.assertEqual(header[21], "供应商批次")
        self.assertEqual(header[22], "产品成本")
        self.assertEqual(header[23], "快递费合计")
        self.assertEqual(header[24], "综合成本")
        # 行级值与 amounts 一致（分项齐全 / 缺失留空 / unmatched 留空）
        c2001 = next(r for r in rows[1:] if r[0] == "C2001")
        self.assertEqual(float(c2001[22]), 40.00)
        self.assertEqual(float(c2001[23]), 8.00)
        self.assertEqual(float(c2001[24]), 48.00)
        c2002 = next(r for r in rows[1:] if r[0] == "C2002")
        self.assertIsNone(c2002[22])
        self.assertIsNone(c2002[23])
        self.assertEqual(float(c2002[24]), 30.00)
        c2004 = next(r for r in rows[1:] if r[0] == "C2004")
        self.assertIsNone(c2004[22])
        self.assertIsNone(c2004[23])
        self.assertIsNone(c2004[24])
        # 合计行与统计卡同源同值
        stat = self._stats()
        summary = rows[-1]
        self.assertEqual(summary[0], "合计")
        self.assertEqual(summary[1], "共 28 单")
        self.assertEqual(summary[3], "产品成本待核对 1 单；快递费合计待核对 1 单")
        self.assertIsInstance(summary[22], (int, float))
        self.assertIsInstance(summary[23], (int, float))
        self.assertIsInstance(summary[24], (int, float))
        self.assertEqual(float(summary[22]), float(stat["product_cost_total"]))
        self.assertEqual(float(summary[23]), float(stat["courier_fee_total"]))
        self.assertEqual(float(summary[24]), float(stat["composite_cost_total"]))
        self.assertEqual(float(summary[12]), float(stat["technical_service_fee"]))

    # ---------- V11 规则4 / V7：多店导出仅一个聚合合计行 ----------
    def test_10_multi_store_export_summary(self):
        """管理员聚合导出：仅一个合计行，数值 = 聚合 statistics（与页面聚合一致）"""
        status, content, _ = self._get_binary(
            "/api/profit/export", client=self.admin)
        self.assertEqual(status, 200)
        wb = load_workbook(io.BytesIO(content))
        rows = list(wb.active.iter_rows(values_only=True))
        # 表头 + 店1 28 明细 + 店2 1 明细 + 1 聚合合计（空数据店跳过）
        self.assertEqual(len(rows), 31)
        summary_rows = [r for r in rows if r[0] == "合计"]
        self.assertEqual(len(summary_rows), 1)
        summary = summary_rows[0]
        self.assertEqual(summary[1], "共 29 单")
        # 与页面聚合三卡双向一致
        st, resp = self.admin.get("/api/profit?size=100")
        self.assertEqual(st, 200)
        agg = resp["data"]["statistics"]
        self.assertEqual(float(summary[22]), float(agg["product_cost_total"]))
        self.assertEqual(float(summary[23]), float(agg["courier_fee_total"]))
        self.assertEqual(float(summary[24]), float(agg["composite_cost_total"]))
        self.assertEqual(float(summary[12]), float(agg["technical_service_fee"]))
        self.assertEqual(agg["technical_service_fee"], "5.00")
        self.assertIsInstance(summary[22], (int, float))
        # 待核对注记与聚合计数一致
        self.assertIn("产品成本待核对 1 单", summary[3])

    # ---------- P1.0 历史版本兼容 ----------
    def test_11_p10_history_compat(self):
        """P1.0 记录（amounts 无三键、无 cost_match_status）不报错且计入待核对"""
        st, resp = self._get("/api/profit?size=100")
        calc_id = resp["data"]["calculation_id"]
        st, resp = self._get("/api/data/profit_statements")
        self.assertEqual(st, 200)
        records = resp["data"]["records"]
        legacy = {
            "record_id": "op_p10_legacy_c2s",
            "record_type": "order_profit",
            "calculation_id": calc_id,
            "rules_version": "P1.0",
            "order_id": "P9LEGACY",
            "order_time": "2026-04-05 10:00:00",
            "profit_status": "calculated",
            "status_reasons": [],
            "amounts": {"operating_contribution_profit": "10.00"},
            "created_at": "2026-04-05 12:00:00",
        }
        records.append(legacy)
        st, resp = self._post("/api/data/profit_statements",
                              {"records": records})
        self.assertEqual(st, 200)

        # 单查不报错；三合计不含该单、missing 计入
        st, resp = self._get("/api/profit?order_id=P9LEGACY")
        self.assertEqual(st, 200)
        self.assertEqual(resp["data"]["statistics"]["total"], 1)
        self.assertEqual(resp["data"]["statistics"]["product_cost_total"], "0.00")
        self.assertEqual(resp["data"]["statistics"]["product_cost_missing_count"], 1)
        self.assertEqual(resp["data"]["statistics"]["courier_fee_missing_count"], 1)
        # 全量：total +1，三合计不变，missing +1
        stat = self._stats()
        self.assertEqual(stat["total"], 29)
        self.assertEqual(stat["product_cost_total"], "47.55")
        self.assertEqual(stat["courier_fee_total"], "15.50")
        self.assertEqual(stat["composite_cost_total"], "93.05")
        self.assertEqual(stat["product_cost_missing_count"], 2)
        self.assertEqual(stat["courier_fee_missing_count"], 2)
        # 导出容错
        status, content, _ = self._get_binary("/api/profit/export?order_id=P9LEGACY")
        self.assertEqual(status, 200)
        rows = list(load_workbook(io.BytesIO(content)).active.iter_rows(values_only=True))
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[2][0], "合计")
        self.assertIn("产品成本待核对 1 单", rows[2][3])

    def test_12b_analysis_payload(self):
        """analysis / ad_expense 与 statistics 同筛选全量聚合"""
        status, resp = self._get("/api/profit?size=20&page=2")
        self.assertEqual(status, 200)
        data = resp["data"]
        self.assertIn("analysis", data)
        a = data["analysis"]
        self.assertIn("structure", a)
        self.assertIn("monthly_trend", a)
        self.assertGreaterEqual(len(a["product_ranking"]), 1)
        self.assertIn("ad_expense", data)
        self.assertIn("total", data["ad_expense"])

    # ---------- V9 前端侧 / 红线静态断言 ----------
    def test_12_frontend_static(self):
        """profit.html 含千位分隔函数与三卡；数值仅来自响应不前端累加"""
        with open(os.path.join(STATIC_DIR, "profit.html"), encoding="utf-8") as f:
            html = f.read()
        with open(os.path.join(STATIC_DIR, "api.js"), encoding="utf-8") as f:
            api = f.read()
        self.assertIn("function fmtMoneyThousands", html)
        for label in ("产品成本合计", "快递费合计", "综合成本合计", "产品毛利"):
            self.assertIn(label, html)
        self.assertIn("chart-structure", html)
        self.assertIn("btn-edit-ad", html)
        for field in COST_KEYS:
            self.assertIn(f"s.{field}", html)
        # renderStats 内不做前端金额累加（三卡数值仅来自响应 statistics）
        m = re.search(r"function renderStats\(data\) \{[\s\S]*?\n\}", html)
        self.assertTrue(m, "renderStats 函数应存在")
        body = m.group(0)
        self.assertNotIn("currentRows", body)
        self.assertNotIn(".results", body)
        # Q1 红线：前端无两卡相加形态
        self.assertNotIn("product_cost_total + courier_fee_total", html)
        # 全局 api.js 未被修改（千位分隔仅页面级）
        self.assertNotIn("fmtMoneyThousands", api)


if __name__ == "__main__":
    unittest.main()

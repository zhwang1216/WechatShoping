"""多用户登录与店铺数据隔离验收测试（spec 5.9.1.1 十二项 + 5.9.1.3 认证行为九项）

全部断言经真实 HTTP 请求驱动（spec 5.9.1.2：禁止服务层直调替代 HTTP 层鉴权断言）。
- TestABIsolation：A/B 双店铺隔离 12 项验收（spec 5.9.1.1）
- TestAuthBehaviors：认证行为 9 类场景 + 安全日志汇总断言（spec 5.9.1.3）
"""
import io
import json
import os
import shutil
import threading
import unittest
from decimal import Decimal

from openpyxl import Workbook, load_workbook

from app.data_schemas import BUSINESS_FILES
from tests.http_client import (ADMIN_PASSWORD, Client, TestEnv, USER_A,
                               USER_A_PASSWORD, USER_B, USER_B_PASSWORD)
from tests.test_delivery_e2e import SUP_HEADERS, SUP_ROWS, WX_HEADERS, WX_ROWS

B_PREFIX = "B2E"
NEW_A_PASSWORD = "NewA123456"
RESET_B_PASSWORD = "NewB123456"
WRONG_PASSWORD = "WrongPass123"


def _dump(obj):
    return json.dumps(obj, ensure_ascii=False)


def make_wx_xlsx(prefix="E2E"):
    """E2E 订单数据；prefix 替换订单号前缀用于构造 B 店铺同构数据"""
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.append(WX_HEADERS)
    for row in WX_ROWS:
        r = list(row)
        if isinstance(r[0], str) and r[0].startswith("E2E"):
            r[0] = prefix + r[0][3:]
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_sup_xlsx(prefix="E2E"):
    """E2E 供应商成本数据；替换主/子订单号前缀"""
    wb = Workbook()
    ws = wb.active
    ws.title = "供应商"
    ws.append(SUP_HEADERS)
    for row in SUP_ROWS:
        r = list(row)
        for i in (0, 1):
            if isinstance(r[i], str) and r[i].startswith("E2E"):
                r[i] = prefix + r[i][3:]
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestABIsolation(unittest.TestCase):
    """A/B 双用户隔离验收（spec 5.9.1.1 全部 12 项）

    主环境：admin + A（E2E 数据）+ B（B2E 同构数据）各自登录并完成导入、对账、利润。
    验收 3/10/11 使用独立环境避免相互污染。
    """

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.a = cls.env.create_user_client(USER_A, USER_A_PASSWORD)
        cls.b = cls.env.create_user_client(USER_B, USER_B_PASSWORD)
        st, _ = cls.a.upload_confirm("/api/import/wechat-orders",
                                     "A订单.xlsx", make_wx_xlsx("E2E"))
        assert st == 200
        st, _ = cls.a.upload_confirm("/api/import/supplier-costs",
                                     "A成本.xlsx", make_sup_xlsx("E2E"))
        assert st == 200
        st, _ = cls.b.upload_confirm("/api/import/wechat-orders",
                                     "B订单.xlsx", make_wx_xlsx(B_PREFIX))
        assert st == 200
        st, _ = cls.b.upload_confirm("/api/import/supplier-costs",
                                     "B成本.xlsx", make_sup_xlsx(B_PREFIX))
        assert st == 200
        cls.a.post("/api/reconciliation/run", {"month": "2026-03"})
        cls.a.post("/api/profit/recalculate", {"month": "2026-03"})
        cls.b.post("/api/reconciliation/run", {"month": "2026-03"})
        cls.b.post("/api/profit/recalculate", {"month": "2026-03"})
        me_a = cls.a.get("/api/auth/me")[1]["data"]
        me_b = cls.b.get("/api/auth/me")[1]["data"]
        cls.a_store_id = me_a["store_id"]
        cls.b_store_id = me_b["store_id"]
        cls.b_user_id = me_b["user_id"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 辅助 ----------
    def _all_order_ids(self, client, kind):
        ids, page = set(), 1
        while True:
            st, r = client.get(f"/api/{kind}?page={page}&size=100")
            self.assertEqual(st, 200)
            for o in r["data"]["orders"]:
                ids.add(o["order_id"])
            if page >= r["data"]["pages"]:
                break
            page += 1
        return ids

    def _profit_of(self, client, order_id):
        st, r = client.get(f"/api/profit?order_id={order_id}")
        self.assertEqual(st, 200)
        rows = r["data"]["results"]
        self.assertEqual(len(rows), 1)
        return rows[0]

    # ---------- 验收 1：A 全部查询仅含 A 店铺数据 ----------
    def test_01_store_a_queries_contain_only_own_data(self):
        cases = [
            "/api/orders?page=1&size=100",
            "/api/refund-orders?page=1&size=100",
            "/api/supplier-costs?page=1&size=100",
            "/api/anomalies?page=1&size=100",
            "/api/reconciliation?page=1&size=100",
            "/api/profit?page=1&size=100",
            "/api/import/batches?page=1&size=100",
            "/api/summary?month=2026-03",
            "/api/ad-expenses",
        ]
        for path in cases:
            with self.subTest(path=path):
                st, r = self.a.get(path)
                self.assertEqual(st, 200, path)
                self.assertNotIn(B_PREFIX + "0", _dump(r), path)
        st, r = self.a.get("/api/orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], 10)
        self.assertTrue(all(x.startswith("E2E")
                            for x in self._all_order_ids(self.a, "orders")))
        st, r = self.a.get("/api/refund-orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], 2)
        st, r = self.a.get("/api/summary?month=2026-03")
        s = r["data"]
        self.assertEqual(s["order_count"], 12)
        self.assertEqual(s["actual_paid_amount"], "1105.00")

    # ---------- 验收 2：B 同样仅含 B 店铺数据 ----------
    def test_02_store_b_queries_contain_only_own_data(self):
        cases = [
            "/api/orders?page=1&size=100",
            "/api/refund-orders?page=1&size=100",
            "/api/supplier-costs?page=1&size=100",
            "/api/anomalies?page=1&size=100",
            "/api/reconciliation?page=1&size=100",
            "/api/profit?page=1&size=100",
            "/api/import/batches?page=1&size=100",
            "/api/summary?month=2026-03",
            "/api/ad-expenses",
        ]
        for path in cases:
            with self.subTest(path=path):
                st, r = self.b.get(path)
                self.assertEqual(st, 200, path)
                self.assertNotIn("E2E0", _dump(r), path)
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(r["data"]["total"], 10)
        self.assertTrue(all(x.startswith(B_PREFIX)
                            for x in self._all_order_ids(self.b, "orders")))
        st, r = self.b.get("/api/summary?month=2026-03")
        self.assertEqual(r["data"]["order_count"], 12)

    # ---------- 验收 3：A 导入后 B 查不到该批数据（独立环境） ----------
    def test_03_import_isolation_between_stores(self):
        env = TestEnv()
        try:
            a = env.create_user_client(USER_A, USER_A_PASSWORD)
            b = env.create_user_client(USER_B, USER_B_PASSWORD)
            st, r = a.upload_confirm("/api/import/wechat-orders",
                                     "A新批次.xlsx", make_wx_xlsx("E2E"))
            self.assertEqual(st, 200)
            # B 的订单列表与搜索均查不到 A 的订单
            st, r = b.get("/api/orders?page=1&size=100")
            self.assertEqual(st, 200)
            self.assertEqual(r["data"]["total"], 0)
            st, r = b.get("/api/orders?keyword=E2E0001&page=1&size=100")
            self.assertEqual(r["data"]["total"], 0)
            # A 的批次 B 不可见：列表不含、详情 404、明细 404
            st, r = a.get("/api/import/batches?page=1&size=100")
            a_batches = {x["batch_id"] for x in r["data"]["batches"]}
            self.assertEqual(len(a_batches), 1)
            st, r = b.get("/api/import/batches?page=1&size=100")
            self.assertEqual(r["data"]["total"], 0)
            for batch_id in a_batches:
                st, r = b.get(f"/api/import/batches/{batch_id}")
                self.assertEqual(st, 404)
                self.assertIn("不存在", r["message"])
                st, r = b.get(f"/api/import/batches/{batch_id}/rows?page=1")
                self.assertEqual(st, 404)
        finally:
            env.cleanup()

    # ---------- 验收 4：A 不能查看或修改 B 的数据 ----------
    def test_04_cross_store_access_and_write_denied(self):
        # 订单详情/退款明细：B 的订单对 A 统一 404
        st, r = self.a.get(f"/api/orders/{B_PREFIX}0001")
        self.assertEqual(st, 404)
        self.assertIn("不存在", r["message"])
        st, r = self.a.get(f"/api/orders/{B_PREFIX}0040/refunds")
        self.assertEqual(st, 404)
        # 成本写数据：A 提交含 B 记录的数据仅作用于 A 店铺
        st, r = self.a.get("/api/data/supplier_costs")
        original = r["data"]["records"]
        forged = original + [dict(original[0], record_id="FORGED-B-COST",
                                  main_order_id=f"{B_PREFIX}0001",
                                  total_amount="999.00")]
        st, r = self.a.post("/api/data/supplier_costs", {"records": forged})
        self.assertEqual(st, 200)
        st, r = self.b.get("/api/data/supplier_costs")
        b_records = r["data"]["records"]
        self.assertFalse(any(c.get("record_id") == "FORGED-B-COST"
                             for c in b_records))
        self.assertFalse(any(c.get("total_amount") == "999.00"
                             for c in b_records))
        # 恢复 A 的原始数据（避免影响后续汇总断言）
        st, r = self.a.post("/api/data/supplier_costs", {"records": original})
        self.assertEqual(st, 200)
        # 广告费保存：A 保存不影响 B 的广告费列表
        st, r = self.a.post("/api/ad-expenses",
                            {"month": "2026-03", "amount": "66.60"})
        self.assertEqual(st, 200)
        st, r = self.b.get("/api/ad-expenses")
        self.assertFalse(any(x.get("month") == "2026-03"
                             and x.get("amount") == "66.60"
                             for x in r["data"]["records"]))

    # ---------- 验收 5：伪造 user_id/store_id 参数无效 ----------
    def test_05_forged_identity_params_ignored(self):
        # 查询携带他人 store_id → 403 拒绝（多店铺隔离强化）
        st, r = self.a.get(
            f"/api/orders?page=1&size=100&store_id={self.b_store_id}"
            f"&user_id={self.b_user_id}")
        self.assertEqual(st, 403)
        # 不带越权参数时仅返回 A 数据
        st, r = self.a.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 10)
        self.assertNotIn(B_PREFIX + "0", _dump(r))
        # 写请求 body 携带他人 store_id → 403（成员不可写他店）
        st, r = self.a.get("/api/data/supplier_costs")
        original = r["data"]["records"]
        st, r = self.a.post("/api/data/supplier_costs",
                            {"store_id": self.b_store_id,
                             "user_id": self.b_user_id,
                             "records": original})
        self.assertEqual(st, 403)
        st, r = self.b.get("/api/data/supplier_costs")
        self.assertEqual(len(r["data"]["records"]), len(original))
        # 广告费保存携带他人 store_id → 403
        st, r = self.a.post("/api/ad-expenses",
                            {"month": "2026-03", "amount": "1.23",
                             "store_id": self.b_store_id})
        self.assertEqual(st, 403)
        st, r = self.b.get("/api/ad-expenses")
        self.assertFalse(any(x.get("amount") == "1.23"
                             for x in r["data"]["records"]))

    # ---------- 验收 6：跨店资源 ID 统一 404"不存在" ----------
    def test_06_cross_store_resource_ids_return_404(self):
        # B 的订单 ID / 退款订单 ID
        st, r = self.a.get(f"/api/orders/{B_PREFIX}0001")
        self.assertEqual(st, 404)
        self.assertIn("不存在", r["message"])
        st, r = self.a.get(f"/api/orders/{B_PREFIX}0040/refunds")
        self.assertEqual(st, 404)
        self.assertIn("不存在", r["message"])
        # B 的导入批次 ID
        st, r = self.b.get("/api/import/batches?page=1&size=100")
        b_batch = r["data"]["batches"][0]["batch_id"]
        st, r = self.a.get(f"/api/import/batches/{b_batch}")
        self.assertEqual(st, 404)
        self.assertIn("不存在", r["message"])
        # B 的成本/退款明细按关键词检索均无结果（不泄露存在性）
        st, r = self.a.get(f"/api/supplier-costs?keyword={B_PREFIX}0001"
                           "&page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 0)

    # ---------- 验收 7：指标/分页/月份/利润/导出按店铺独立 ----------
    def test_07_independent_metrics_pagination_and_exports(self):
        # A、B 各自汇总（同构数据同数值；若串户 order_count 将翻倍为 24）
        for client in (self.a, self.b):
            st, r = client.get("/api/summary?month=2026-03")
            s = r["data"]
            self.assertEqual(s["order_count"], 12)
            self.assertEqual(s["refund_order_count"], 2)
            self.assertEqual(s["actual_paid_amount"], "1105.00")
            self.assertEqual(s["supplier_total_cost"], "421.00")
            self.assertEqual(s["technical_service_fee"], "28.00")
            st, r = client.get("/api/summary?month=2026-04")
            self.assertFalse(r["data"]["has_orders"])
        # 分页总数各自独立
        for client in (self.a, self.b):
            st, r = client.get("/api/orders?page=1&size=50")
            self.assertEqual(r["data"]["total"], 10)
            self.assertEqual(r["data"]["pages"], 1)
        # 利润结果独立且同构一致
        rec_a = self._profit_of(self.a, "E2E0060")
        rec_b = self._profit_of(self.b, f"{B_PREFIX}0060")
        self.assertEqual(rec_a["amounts"]["operating_contribution_profit"],
                         rec_b["amounts"]["operating_contribution_profit"])
        st, r = self.a.get("/api/profit?page=1&size=100")
        self.assertEqual(r["data"]["total"], 12)
        st, r = self.b.get("/api/profit?page=1&size=100")
        self.assertEqual(r["data"]["total"], 12)
        # 导出文件内容仅含本店铺订单
        st, content, _ = self.a.get_binary("/api/orders/export")
        self.assertEqual(st, 200)
        text_a = "\n".join(
            str(c) for row in load_workbook(io.BytesIO(content))
            .active.iter_rows(values_only=True) for c in row if c)
        self.assertIn("E2E0001", text_a)
        self.assertNotIn(B_PREFIX + "0", text_a)
        st, content, _ = self.b.get_binary("/api/orders/export")
        self.assertEqual(st, 200)
        text_b = "\n".join(
            str(c) for row in load_workbook(io.BytesIO(content))
            .active.iter_rows(values_only=True) for c in row if c)
        self.assertIn(B_PREFIX + "0001", text_b)
        self.assertNotIn("E2E0", text_b)

    # ---------- 验收 12：广告费隔离 + 财务规则在隔离环境原样通过 ----------
    def test_08_ad_expense_isolation_and_financial_rules(self):
        # 广告费仅本人店铺可读写
        st, r = self.a.post("/api/ad-expenses",
                            {"month": "2026-03", "amount": "33.33"})
        self.assertEqual(st, 200)
        st, r = self.b.post("/api/ad-expenses",
                            {"month": "2026-03", "amount": "44.44"})
        self.assertEqual(st, 200)
        st, r = self.a.get("/api/ad-expenses")
        a_amounts = {x["month"]: x["amount"]
                     for x in r["data"]["records"]}
        self.assertEqual(a_amounts.get("2026-03"), "33.33")
        st, r = self.b.get("/api/ad-expenses")
        b_amounts = {x["month"]: x["amount"]
                     for x in r["data"]["records"]}
        self.assertEqual(b_amounts.get("2026-03"), "44.44")
        # 既有财务规则：排除退款订单、综合成本单次扣减、退款口径
        self.a.post("/api/profit/config",
                    {"refund_deduction_status": "unknown"})
        self.a.post("/api/profit/recalculate", {})
        rec = self._profit_of(self.a, "E2E0040")
        self.assertEqual(rec["profit_status"], "pending_confirmation")
        self.assertIn("REFUND_DEDUCTION_UNCONFIRMED", rec["status_reasons"])
        self.a.post("/api/profit/config",
                    {"refund_deduction_status": "deducted"})
        self.a.post("/api/profit/recalculate", {})
        rec = self._profit_of(self.a, "E2E0060")
        self.assertEqual(rec["amounts"]["supplier_total_amount"], "48.00")
        self.assertEqual(rec["amounts"]["operating_contribution_profit"],
                         "49.00")
        rec = self._profit_of(self.a, "E2E0050")
        self.assertEqual(rec["amounts"]["net_revenue"], "70.00")
        self.a.post("/api/profit/config",
                    {"refund_deduction_status": "not_deducted"})
        self.a.post("/api/profit/recalculate", {})
        rec = self._profit_of(self.a, "E2E0050")
        self.assertEqual(rec["amounts"]["net_revenue"], "0.00")
        self.assertIsNone(rec["margin"]["percent"])
        self.assertEqual(rec["margin"]["blocked_reason"],
                         "NET_REVENUE_NOT_POSITIVE")
        # 除广告花费外其他财务金额无前端写入口径（页面/脚本无 /api/data/ 调用）
        static_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "static")
        for fn in os.listdir(static_dir):
            if fn.endswith((".html", ".js")):
                with open(os.path.join(static_dir, fn),
                          encoding="utf-8") as f:
                    self.assertNotIn("/api/data/", f.read(), fn)

    # ---------- 验收 9：停用/启用用户与会话吊销 ----------
    def test_09_disable_enable_user_lifecycle(self):
        st, r = self.admin.post(
            f"/api/users/{self.b_user_id}/status", {"action": "disable"})
        self.assertEqual(st, 200)
        # B 的活动会话立即失效
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 401)
        # B 登录统一"账号或密码错误"
        st, r = Client(self.env).post(
            "/api/auth/login",
            {"username": USER_B, "password": USER_B_PASSWORD}, csrf=False)
        self.assertEqual(st, 401)
        self.assertEqual(r["message"], "账号或密码错误")
        # 重新启用后原密码登录成功且历史数据完整
        st, r = self.admin.post(
            f"/api/users/{self.b_user_id}/status", {"action": "enable"})
        self.assertEqual(st, 200)
        type(self).b = self.env.login_client(USER_B, USER_B_PASSWORD)
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], 10)

    # ---------- 验收 8：退出登录后旧会话全部失效 ----------
    def test_10_logout_revokes_session(self):
        st, _ = self.a.logout()
        self.assertEqual(st, 200)
        st, r = self.a.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 401)
        for page in ("/orders", "/dashboard", "/profit"):
            st, _, headers = self.a.get_page(page)
            self.assertEqual(st, 302, page)
            self.assertEqual(headers.get("Location"), "/login", page)

    # ---------- 验收 10：旧数据迁移零丢失零串户（独立环境） ----------
    def test_11_legacy_migration_zero_loss(self):
        src = TestEnv()
        new = None
        try:
            src_member = src.member  # 业务承载成员（管理员无业务店铺）
            st, _ = src_member.upload_confirm(
                "/api/import/wechat-orders", "源订单.xlsx", make_wx_xlsx("E2E"))
            self.assertEqual(st, 200)
            st, _ = src_member.upload_confirm(
                "/api/import/supplier-costs", "源成本.xlsx", make_sup_xlsx("E2E"))
            self.assertEqual(st, 200)
            st, _ = src_member.post("/api/reconciliation/run", {})
            self.assertEqual(st, 200)
            st, _ = src_member.post("/api/profit/recalculate", {})
            self.assertEqual(st, 200)
            st, _ = src_member.post("/api/ad-expenses",
                                    {"month": "2026-03", "amount": "55.55"})
            self.assertEqual(st, 200)
            st, _, _ = src_member.get_binary("/api/orders/export")
            self.assertEqual(st, 200)
            # 仅预览未确认的上传（历史遗留源文件形态: .xlsx + .meta.json）
            st, _ = src_member.upload(
                f"/api/import/wechat-orders?store_id={src.member_store}",
                "遗留暂存.xlsx", make_wx_xlsx("E2E"))
            self.assertEqual(st, 200)
            src_store = os.path.join(src.data_dir, "stores", src.member_store)
            expected = {}
            for name, spec in BUSINESS_FILES.items():
                with open(os.path.join(src_store, "data", spec["file"]),
                          encoding="utf-8") as f:
                    expected[name] = json.load(f)
            upload_files = set(os.listdir(os.path.join(src_store,
                                                       "uploads")))
            self.assertGreaterEqual(len(upload_files), 2)

            # 构造旧单店铺数据布局：业务 JSON 在数据根、uploads 在 data 同级
            new = TestEnv(initialize=False)
            for name, spec in BUSINESS_FILES.items():
                shutil.copy2(os.path.join(src_store, "data", spec["file"]),
                             os.path.join(new.data_dir, spec["file"]))
            old_uploads = os.path.join(new.tmp, "uploads")
            os.makedirs(old_uploads, exist_ok=True)
            for fn in upload_files:
                shutil.copy2(os.path.join(src_store, "uploads", fn),
                             os.path.join(old_uploads, fn))
            # 初始化并迁移
            st, r = new.setup_admin(acknowledge_legacy=True)
            self.assertEqual(st, 200, r)
            migration = r["data"]["migration"]
            self.assertEqual(
                set(migration["migrated"]),
                {spec["file"] for spec in BUSINESS_FILES.values()})
            self.assertTrue(os.path.isdir(migration["backup_dir"]))
            self.assertGreaterEqual(migration["uploads"],
                                    len(upload_files))
            # 管理员无店；旧数据迁入待分配店铺（owner 为空）
            self.assertIsNone(r["data"]["store_id"])
            with open(os.path.join(new.data_dir, "stores.json"),
                      encoding="utf-8") as f:
                unassigned = [s for s in json.load(f)["stores"]
                              if s.get("owner_user_id") is None]
            self.assertEqual(len(unassigned), 1)
            new_store = os.path.join(new.data_dir, "stores",
                                     unassigned[0]["store_id"])
            # 旧文件全部移出数据根；备份与店铺目录内容与源完全一致
            for name, spec in BUSINESS_FILES.items():
                self.assertFalse(
                    os.path.exists(os.path.join(new.data_dir, spec["file"])))
                for d in (migration["backup_dir"],
                          os.path.join(new_store, "data")):
                    with open(os.path.join(d, spec["file"]),
                              encoding="utf-8") as f:
                        self.assertEqual(json.load(f), expected[name],
                                         f"{d}/{spec['file']}")
            self.assertEqual(
                set(os.listdir(os.path.join(new_store, "uploads"))),
                upload_files)
            self.assertTrue(os.path.isfile(
                os.path.join(new_store, "MIGRATION_REPORT.json")))
            with open(os.path.join(new_store, "MIGRATION_REPORT.json"),
                      encoding="utf-8") as f:
                report = json.load(f)
            self.assertEqual(report["unassigned_store_id"],
                             unassigned[0]["store_id"])
            self.assertEqual(report["assignment_status"], "unassigned")
            # 关键金额 Decimal 精确比对（orders 为商品行级记录，按订单去重）
            by_order = {}
            for o in expected["orders"]["records"]:
                by_order[o["order_id"]] = Decimal(
                    o["amounts"]["actual_received_amount"])
            self.assertEqual(sum(by_order.values(), Decimal("0")),
                             Decimal("1105.00"))
            # 待分配店铺不计入管理员聚合与任何成员视图（分配归属前不可见）
            other = new.create_user_client(USER_B, USER_B_PASSWORD)
            st, r = new.admin.get("/api/orders?page=1&size=100")
            self.assertEqual(st, 200)
            self.assertEqual(r["data"]["total"], 0)
            # 无串户：新用户看不到迁移数据
            st, r = other.get("/api/orders?page=1&size=100")
            self.assertEqual(st, 200)
            self.assertEqual(r["data"]["total"], 0)
        finally:
            if new is not None:
                new.cleanup()
            src.cleanup()

    # ---------- 验收 11：A/B 并发上传导入互不覆盖（独立环境） ----------
    def test_12_concurrent_imports_isolated(self):
        env = TestEnv()
        try:
            a = env.create_user_client(USER_A, USER_A_PASSWORD)
            b = env.create_user_client(USER_B, USER_B_PASSWORD)
            barrier = threading.Barrier(2)
            results, errors = {}, {}

            def worker(client, tag, wx, sup, key):
                try:
                    barrier.wait(timeout=15)
                    st1, _ = client.upload_confirm(
                        "/api/import/wechat-orders", f"{tag}订单.xlsx", wx)
                    barrier.wait(timeout=15)
                    st2, _ = client.upload_confirm(
                        "/api/import/supplier-costs", f"{tag}成本.xlsx", sup)
                    results[key] = (st1, st2)
                except Exception as e:  # noqa: BLE001
                    errors[key] = repr(e)

            t1 = threading.Thread(
                target=worker, args=(a, "A", make_wx_xlsx("E2E"),
                                     make_sup_xlsx("E2E"), "a"))
            t2 = threading.Thread(
                target=worker, args=(b, "B", make_wx_xlsx(B_PREFIX),
                                     make_sup_xlsx(B_PREFIX), "b"))
            t1.start()
            t2.start()
            t1.join(timeout=60)
            t2.join(timeout=60)
            self.assertEqual(errors, {})
            self.assertEqual(results["a"], (200, 200))
            self.assertEqual(results["b"], (200, 200))
            # 与单独导入时完全一致
            for client, prefix in ((a, "E2E"), (b, B_PREFIX)):
                st, r = client.get("/api/orders?page=1&size=100")
                self.assertEqual(st, 200)
                self.assertEqual(r["data"]["total"], 10)
                ids = self._all_order_ids(client, "orders")
                self.assertTrue(all(x.startswith(prefix) for x in ids))
                self.assertEqual(len(ids), 10)
                st, r = client.get("/api/refund-orders?page=1&size=100")
                self.assertEqual(r["data"]["total"], 2)
                st, r = client.get("/api/import/batches?page=1&size=100")
                self.assertEqual(r["data"]["total"], 2)
                st, r = client.get("/api/summary?month=2026-03")
                self.assertEqual(r["data"]["order_count"], 12)
                self.assertEqual(r["data"]["actual_paid_amount"], "1105.00")
        finally:
            env.cleanup()


class TestAuthBehaviors(unittest.TestCase):
    """认证行为补充测试（spec 5.9.1.3 全部 9 类场景 + 安全日志汇总断言）"""

    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.a = cls.env.create_user_client(USER_A, USER_A_PASSWORD)
        cls.b = cls.env.create_user_client(USER_B, USER_B_PASSWORD)
        cls.a_password = USER_A_PASSWORD
        cls.b_password = USER_B_PASSWORD
        cls.b_user_id = cls.b.get("/api/auth/me")[1]["data"]["user_id"]

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    def _security_log_text(self, env=None):
        path = os.path.join((env or self.env).log_dir, "security_log.jsonl")
        with open(path, encoding="utf-8") as f:
            return f.read()

    # ---------- 1. 未登录/未初始化的访问引导 ----------
    def test_01_unauthenticated_and_uninitialized_redirects(self):
        anon = Client(self.env)
        # 已初始化环境：未登录
        st, r = anon.get("/api/orders?page=1&size=50")
        self.assertEqual(st, 401)
        self.assertEqual(r["code"], 1001)
        for page in ("/", "/orders", "/profit", "/users"):
            st, _, headers = anon.get_page(page)
            self.assertEqual(st, 302, page)
            self.assertEqual(headers.get("Location"), "/login", page)
        # 未初始化环境：业务页 → /setup、业务 API → 401、登录页可访问
        fresh = TestEnv(initialize=False)
        try:
            anon2 = Client(fresh)
            st, _, headers = anon2.get_page("/orders")
            self.assertEqual(st, 302)
            self.assertEqual(headers.get("Location"), "/setup")
            st, _, headers = anon2.get_page("/")
            self.assertEqual(st, 302)
            self.assertEqual(headers.get("Location"), "/setup")
            st, r = anon2.get("/api/orders?page=1&size=50")
            self.assertEqual(st, 401)
            st, _, _ = anon2.get_page("/setup")
            self.assertEqual(st, 200)
            st, _, _ = anon2.get_page("/login")
            self.assertEqual(st, 200)
        finally:
            fresh.cleanup()

    # ---------- 1b. 登录页/初始化页跳转无循环（重定向循环回归防护） ----------
    def test_01b_login_setup_pages_no_redirect_loop(self):
        # 已初始化 + 未登录：/login 必须放行（302 回 / 会与 / → /login 形成死循环）
        anon = Client(self.env)
        st, body, headers = anon.get_page("/login")
        self.assertEqual(st, 200)
        self.assertIn("登录", body)
        # 已初始化 + 未登录：/setup 一次受控跳转到 /login（不再经 / 中转）
        st, _, headers = anon.get_page("/setup")
        self.assertEqual(st, 302)
        self.assertEqual(headers.get("Location"), "/login")
        # 已登录：/login、/setup 引导回首页（一次性, 首页放行不再跳）
        for page in ("/login", "/setup"):
            st, _, headers = self.a.get_page(page)
            self.assertEqual(st, 302, page)
            self.assertEqual(headers.get("Location"), "/", page)
        st, _, _ = self.a.get_page("/")
        self.assertEqual(st, 200)

    # ---------- 2. CSRF 防护与数据无变化 ----------
    def test_02_csrf_header_required_for_writes(self):
        total_before = self.a.get(
            "/api/orders?page=1&size=100")[1]["data"]["total"]
        # 不带 X-Requested-With 的写请求全部 403
        st, r = self.a.post("/api/import/confirm",
                            {"upload_id": "nonexistent"}, csrf=False)
        self.assertEqual(st, 403)
        self.assertEqual(r["code"], 1002)
        st, r = self.a.post("/api/auth/logout", csrf=False)
        self.assertEqual(st, 403)
        self.assertEqual(r["code"], 1002)
        st, r = self.a.upload("/api/import/wechat-orders", "csrf.xlsx",
                              make_wx_xlsx("E2E"), csrf=False)
        self.assertEqual(st, 403)
        # 会话未被登出、数据无变化
        st, r = self.a.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        self.assertEqual(r["data"]["total"], total_before)
        # 普通用户访问用户管理 API/页面被拒
        st, r = self.a.get("/api/users")
        self.assertEqual(st, 403)
        st, _, headers = self.a.get_page("/users")
        self.assertEqual(st, 302)
        self.assertEqual(headers.get("Location"), "/")

    # ---------- 3. 停用账号登录统一文案 ----------
    def test_03_disabled_account_login_unified_message(self):
        st, r = self.admin.post(
            f"/api/users/{self.b_user_id}/status", {"action": "disable"})
        self.assertEqual(st, 200)
        st, r = Client(self.env).post(
            "/api/auth/login",
            {"username": USER_B, "password": self.b_password}, csrf=False)
        self.assertEqual(st, 401)
        self.assertEqual(r["message"], "账号或密码错误")
        self.assertIn("login.disabled", self._security_log_text())
        st, r = self.admin.post(
            f"/api/users/{self.b_user_id}/status", {"action": "enable"})
        self.assertEqual(st, 200)
        # 注意: 替换客户端必须写类属性(每个测试方法都是新实例)
        type(self).b = self.env.login_client(USER_B, self.b_password)

    # ---------- 4. 修改/重置密码后全部旧会话失效 ----------
    def test_04_password_change_revokes_all_sessions(self):
        # 本人修改密码
        st, r = self.a.post("/api/auth/change-password", {
            "old_password": self.a_password,
            "new_password": NEW_A_PASSWORD,
            "confirm_password": NEW_A_PASSWORD})
        self.assertEqual(st, 200)
        self.assertGreaterEqual(r["data"]["revoked_sessions"], 1)
        self.assertTrue(r["data"]["relogin"])
        st, r = self.a.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 401)
        st, r = Client(self.env).post(
            "/api/auth/login",
            {"username": USER_A, "password": self.a_password}, csrf=False)
        self.assertEqual(st, 401)
        type(self).a_password = NEW_A_PASSWORD
        type(self).a = self.env.login_client(USER_A, self.a_password)
        # 管理员重置密码
        st, r = self.admin.post(
            f"/api/users/{self.b_user_id}/reset-password",
            {"new_password": RESET_B_PASSWORD})
        self.assertEqual(st, 200)
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 401)
        type(self).b_password = RESET_B_PASSWORD
        type(self).b = self.env.login_client(USER_B, self.b_password)
        st, r = self.b.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)

    # ---------- 5. 服务重启后未过期会话仍有效 ----------
    def test_05_restart_keeps_sessions(self):
        st, _ = self.a.get("/api/orders?page=1&size=100")
        self.assertEqual(st, 200)
        self.env.restart()
        for client in (self.admin, self.a, self.b):
            st, r = client.get("/api/orders?page=1&size=100")
            self.assertEqual(st, 200)

    # ---------- 6. 重复初始化被拒 ----------
    def test_06_double_init_rejected(self):
        st, r = self.admin.post("/api/setup/init", {
            "username": "hackadmin", "password": "Hack12345678",
            "confirm_password": "Hack12345678", "store_name": "x"})
        self.assertEqual(st, 400)
        self.assertIn("系统已初始化", r["message"])
        st, _, headers = self.admin.get_page("/setup")
        self.assertEqual(st, 302)
        self.assertEqual(headers.get("Location"), "/")

    # ---------- 7. 旧数据未确认归属时阻止初始化 ----------
    def test_07_legacy_requires_acknowledgement(self):
        env = TestEnv(initialize=False)
        try:
            legacy = {"version": 1, "updated_at": "2026-01-01T00:00:00",
                      "records": []}
            with open(os.path.join(env.data_dir, "orders.json"), "w",
                      encoding="utf-8") as f:
                json.dump(legacy, f, ensure_ascii=False)
            st, r = env.setup_admin(acknowledge_legacy=False)
            self.assertEqual(st, 400)
            self.assertIn("确认迁移方式", r["message"])
            self.assertIn("待分配", r["message"])
            anon = Client(env)
            st, r = anon.get("/api/setup/status")
            self.assertEqual(st, 200)
            self.assertFalse(r["data"]["initialized"])
            self.assertIn("orders.json", r["data"]["legacy_files"])
            st, r = env.setup_admin(acknowledge_legacy=True)
            self.assertEqual(st, 200, r)
            self.assertIn("orders.json", r["data"]["migration"]["migrated"])
        finally:
            env.cleanup()

    # ---------- 8. 迁移失败回滚并可直接重试 ----------
    def test_08_migration_failure_rollback_and_retry(self):
        env = TestEnv(initialize=False)
        stores_root = os.path.join(env.data_dir, "stores")
        try:
            orders = {"version": 1, "updated_at": "2026-01-01T00:00:00",
                      "records": [
                          {"order_id": "RB0001",
                           "order_time": "2026-03-01 10:00:00"},
                          {"order_id": "RB0002",
                           "order_time": "2026-03-02 10:00:00"}]}
            costs = {"version": 1, "updated_at": "2026-01-01T00:00:00",
                     "records": [
                         {"record_id": "RBC0001", "main_order_id": "RB0001",
                          "total_amount": "30.00"}]}
            for name, payload in (("orders", orders),
                                  ("supplier_costs", costs)):
                with open(os.path.join(env.data_dir,
                                       BUSINESS_FILES[name]["file"]), "w",
                          encoding="utf-8") as f:
                    json.dump(payload, f, ensure_ascii=False)
            # 店铺根目录设为只读 → 迁移中断
            os.chmod(stores_root, 0o500)
            try:
                st, r = env.setup_admin(acknowledge_legacy=True)
                self.assertEqual(st, 500, r)
                self.assertIn("迁移失败", r["message"])
                self.assertIn("回滚", r["message"])
            finally:
                os.chmod(stores_root, 0o755)
            # 旧文件原样、系统回到未初始化状态
            with open(os.path.join(env.data_dir, "orders.json"),
                      encoding="utf-8") as f:
                self.assertEqual(json.load(f), orders)
            anon = Client(env)
            st, r = anon.get("/api/setup/status")
            self.assertEqual(st, 200)
            self.assertFalse(r["data"]["initialized"])
            self.assertIn("setup.rollback", self._security_log_text(env))
            # 修复后重试成功且数据完整（迁入待分配店铺，文件级核对）
            st, r = env.setup_admin(acknowledge_legacy=True)
            self.assertEqual(st, 200, r)
            with open(os.path.join(env.data_dir, "stores.json"),
                      encoding="utf-8") as f:
                unassigned = [s for s in json.load(f)["stores"]
                              if s.get("owner_user_id") is None]
            self.assertEqual(len(unassigned), 1)
            with open(os.path.join(env.data_dir, "stores",
                                   unassigned[0]["store_id"],
                                   "data", "orders.json"),
                      encoding="utf-8") as f:
                self.assertEqual(len(json.load(f)["records"]), 2)
            # 待分配店铺不计入管理员聚合视图（分配归属前不可见）
            st, r = env.admin.get("/api/data/orders")
            self.assertEqual(st, 404)
        finally:
            os.chmod(stores_root, 0o755)
            env.cleanup()

    # ---------- 9. 登录限流与统一错误文案 ----------
    def test_09_login_rate_limit_and_unified_message(self):
        # 同一账号连续 10 次密码错误
        for _ in range(10):
            st, r = Client(self.env).post(
                "/api/auth/login",
                {"username": USER_B, "password": WRONG_PASSWORD},
                csrf=False)
            self.assertEqual(st, 401)
            self.assertEqual(r["message"], "账号或密码错误")
        # 第 11 次即使密码正确也被限流
        st, r = Client(self.env).post(
            "/api/auth/login",
            {"username": USER_B, "password": self.b_password}, csrf=False)
        self.assertEqual(st, 429)
        self.assertEqual(r["code"], 1004)
        self.assertEqual(r["message"], "失败次数过多，请 15 分钟后再试")
        # 不存在账号与错误密码的提示完全一致（不泄露账号存在性）
        st, r1 = Client(self.env).post(
            "/api/auth/login",
            {"username": "ghost_user", "password": WRONG_PASSWORD},
            csrf=False)
        st, r2 = Client(self.env).post(
            "/api/auth/login",
            {"username": USER_A, "password": WRONG_PASSWORD}, csrf=False)
        self.assertEqual(r1["message"], r2["message"])

    # ---------- 10. 安全日志汇总：事件齐全且无任何秘密 ----------
    def test_10_security_log_complete_and_secret_free(self):
        # 主动触发 logout 事件补全记录
        st, _ = self.b.logout()
        self.assertEqual(st, 200)
        text = self._security_log_text()
        events = {json.loads(line)["event"]
                  for line in text.splitlines() if line.strip()}
        expected_events = {
            "setup.admin_created", "login.success", "login.failed",
            "login.rate_limited", "login.disabled", "logout", "user.create",
            "user.disable", "user.enable", "password.change",
            "user.password_reset", "access_denied"}
        self.assertTrue(expected_events <= events,
                        f"缺少事件: {expected_events - events}")
        # 不含任何密码、密码哈希与会话令牌
        secrets = [ADMIN_PASSWORD, USER_A_PASSWORD, USER_B_PASSWORD,
                   NEW_A_PASSWORD, RESET_B_PASSWORD, WRONG_PASSWORD,
                   "Hack12345678", self.admin.cookie, self.a.cookie]
        for secret in secrets:
            self.assertNotIn(secret, text)
        for marker in ("pbkdf2", "salt", "wxr_session"):
            self.assertNotIn(marker, text.lower())
        # 各事件条目结构完整
        for line in text.splitlines():
            if line.strip():
                entry = json.loads(line)
                self.assertIn(entry["event"], expected_events | {
                    "server.init", "server.start", "server.stop",
                    "migration.success", "setup.rollback", "api.error"})
                self.assertTrue(entry["time"])
                self.assertIn("result", entry)


if __name__ == "__main__":
    unittest.main()

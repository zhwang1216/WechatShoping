"""数据导入店铺选择弹窗（import_store_picker）验收测试 G1-G12

覆盖 spec.md 5.1~5.7 全部 17 项验收条件（后端契约已由店铺归属特性
组 1-6 交付：_import_target 必填校验 / 三接口收紧 / 预览回显三字段 /
meta 一致性 / 批次快照）；G11 验证 S2 前端弹窗改造后的页面形态。

全部断言经真实 HTTP 请求驱动 + 临时数据根业务文件直读双通道；
严禁读写真实 data/ 目录（TestEnv tempfile 隔离）。
"""
import json
import os
import unittest

from app.data_schemas import BUSINESS_FILES
from tests.http_client import ADMIN_PASSWORD, ADMIN_USER, Client, TestEnv
from tests.test_auth_isolation import make_sup_xlsx, make_wx_xlsx

PW_A = "PickA123456"
PW_B = "PickB123456"


def _store_records(env, store_id, repo_name):
    """直读店铺业务文件记录（文件侧真值）"""
    fp = os.path.join(env.data_dir, "stores", store_id, "data",
                      BUSINESS_FILES[repo_name]["file"])
    if not os.path.exists(fp):
        return []
    with open(fp, encoding="utf-8") as f:
        return json.load(f).get("records", [])


def _security_events(env):
    fp = os.path.join(env.log_dir, "security_log.jsonl")
    if not os.path.exists(fp):
        return []
    with open(fp, encoding="utf-8") as f:
        return [json.loads(line) for line in f if line.strip()]


class TestImportStorePicker(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = TestEnv()
        cls.admin = cls.env.admin
        cls.a = cls.env.create_user_client("px_pick_a", PW_A,
                                           store_name="选店A铺")
        cls.b = cls.env.create_user_client("px_pick_b", PW_B,
                                           store_name="选店B铺")
        cls.a_store = cls.a.store_id
        cls.b_store = cls.b.store_id
        cls.a_user_id = cls.env.get_user_id("px_pick_a")
        cls.b_user_id = cls.env.get_user_id("px_pick_b")
        cls.admin_user_id = cls.env.get_user_id(ADMIN_USER)

    @classmethod
    def tearDownClass(cls):
        cls.env.cleanup()

    # ---------- 辅助 ----------
    def _upload_only(self, client, path, store_id, prefix, filename=None):
        """仅预览（不确认）→ (status, resp)；成功时返回 upload_id"""
        st, r = client.upload(
            f"{path}?store_id={store_id}",
            filename or f"{prefix}.xlsx",
            make_wx_xlsx(prefix) if "wechat" in path else make_sup_xlsx(prefix))
        return st, r

    def _import_ok(self, client, path, store_id, prefix):
        """完整导入（上传 + 确认）→ (status, resp)"""
        st, r = self._upload_only(client, path, store_id, prefix)
        if st != 200:
            return st, r
        st, r = client.post("/api/import/confirm", {
            "upload_id": r["data"]["upload_id"],
            "target_store_id": store_id})
        return st, r

    # ---------- G1 必填强校验（验收项 11） ----------
    def test_g1_store_id_required(self):
        before_a = len(_store_records(self.env, self.a_store,
                                      "import_records"))
        for path in ("/api/import/wechat-orders",
                     "/api/import/supplier-costs"):
            for sid in ("", "all"):
                st, r = self.a.upload(f"{path}?store_id={sid}",
                                      "缺参.xlsx", make_wx_xlsx("G1"))
                self.assertEqual(st, 400, (path, sid, r))
                self.assertIn("请先选择目标店铺", r["message"])
            st, r = self.a.upload(path, "缺参.xlsx", make_wx_xlsx("G1"))
            self.assertEqual(st, 400, (path, r))
            self.assertIn("请先选择目标店铺", r["message"])
        # confirm 缺 target_store_id → 400（需先成功上传取得 upload_id）
        st, r = self._upload_only(self.a, "/api/import/wechat-orders",
                                  self.a_store, "G1")
        self.assertEqual(st, 200, r)
        st, r = self.a.post("/api/import/confirm", {
            "upload_id": r["data"]["upload_id"]})
        self.assertEqual(st, 400, r)
        self.assertIn("请先选择目标店铺", r["message"])
        # 无批次产生
        after_a = len(_store_records(self.env, self.a_store,
                                     "import_records"))
        self.assertEqual(after_a, before_a)

    # ---------- G2 越权与无效目标（验收项 9/12） ----------
    def test_g2_forbidden_and_invalid_targets(self):
        # member 传他店 → 403 + import_store_forbidden 审计
        st, r = self.a.upload(
            f"/api/import/wechat-orders?store_id={self.b_store}",
            "越权.xlsx", make_wx_xlsx("G2X"))
        self.assertEqual(st, 403, r)
        self.assertIn("无权上传到该店铺", r["message"])
        events = _security_events(self.env)
        self.assertTrue(any(
            e.get("event") == "access_denied"
            and (e.get("detail") or {}).get("reason") == "import_store_forbidden"
            for e in events))
        # member 传不存在 ID → 403（他店口径）
        st, r = self.a.upload(
            "/api/import/wechat-orders?store_id=s_not_exist",
            "不存在.xlsx", make_wx_xlsx("G2N"))
        self.assertEqual(st, 403, r)
        # admin 传不存在 ID → 404
        st, r = self.admin.upload(
            "/api/import/wechat-orders?store_id=s_not_exist",
            "不存在.xlsx", make_wx_xlsx("G2N"))
        self.assertEqual(st, 404, r)
        self.assertIn("店铺不存在", r["message"])
        # 停用店铺（admin 视角全覆盖）
        c = self.env.create_user_client("px_pick_c", "PickC123456",
                                        store_name="选店C铺")
        c_store = c.store_id
        st, r = self.admin.post(f"/api/stores/{c_store}/status",
                                {"action": "disable"})
        self.assertEqual(st, 200, r)
        st, r = self.admin.upload(
            f"/api/import/wechat-orders?store_id={c_store}",
            "停用.xlsx", make_wx_xlsx("G2D"))
        self.assertEqual(st, 400, r)
        self.assertIn("该店铺已停用", r["message"])
        # 归档子场景（2026-10-23 契约变更）: 归档功能已被
        # store_physical_delete 并行特性下线, 店铺无法再进入归档态,
        # _import_target 归档分支成为防御性代码; 此处断言归档构造
        # 请求被显式拒绝, 停用态上传拦截已由上文覆盖
        st, r = self.admin.post(f"/api/stores/{c_store}/status",
                                {"action": "archive"})
        self.assertEqual(st, 400, r)
        self.assertIn("归档功能已下线", r["message"])
        # 停用/归档/越权均未写入任何批次
        self.assertEqual(_store_records(self.env, c_store,
                                        "import_records"), [])
        self.assertEqual(_store_records(self.env, self.b_store,
                                        "import_records"), [])

    # ---------- G3 伪造身份无效（spec 5.5.1 规则 9） ----------
    def test_g3_forged_identity_ineffective(self):
        # 登录体伪造 role=admin：会话仍为 member，他店上传 → 403
        forged = Client(self.env)
        forged.login("px_pick_a", PW_A, extra={"role": "admin"})
        st, r = forged.upload(
            f"/api/import/wechat-orders?store_id={self.b_store}",
            "伪造角色.xlsx", make_wx_xlsx("G3F"))
        self.assertEqual(st, 403, r)
        # confirm body 携带他人 user_id / role：服务端按会话身份落账
        st, r = self._import_ok(self.a, "/api/import/wechat-orders",
                                self.a_store, "G3")
        self.assertEqual(st, 200, r)
        st, r = self.a.post("/api/import/confirm", {
            "upload_id": "any", "target_store_id": self.a_store,
            "user_id": self.admin_user_id, "role": "admin"})
        self.assertIn(st, (400, 404), r)  # 过期 upload_id 被拒，不写数据
        batch = next(b for b in _store_records(
            self.env, self.a_store, "import_records")
            if b.get("file_name") == "G3.xlsx")
        self.assertEqual(batch["uploaded_by_user_id"], self.a_user_id)

    # ---------- G4 预览-确认目标一致性（验收项 13） ----------
    def test_g4_confirm_store_consistency(self):
        st, r = self._upload_only(self.a, "/api/import/wechat-orders",
                                  self.a_store, "G4")
        self.assertEqual(st, 200, r)
        upload_id = r["data"]["upload_id"]
        # member 确认时改道他店 → 403（会话身份权威先行拒绝）
        st, r = self.a.post("/api/import/confirm", {
            "upload_id": upload_id, "target_store_id": self.b_store})
        self.assertEqual(st, 403, r)
        # admin 改道：目标店 bundle 内无此暂存 → 404 拒绝（改道不可达
        # 数据：upload_id 暂存于预览店铺，确认店铺不一致时文件不存在）
        st, r = self.admin.post("/api/import/confirm", {
            "upload_id": upload_id, "target_store_id": self.b_store})
        self.assertEqual(st, 404, r)
        self.assertIn("上传文件不存在或已过期", r["message"])
        # 两店均无新增批次/订单数据
        self.assertEqual(_store_records(self.env, self.b_store,
                                        "import_records"), [])
        a_orders = [o for o in _store_records(self.env, self.a_store,
                                              "orders")
                    if str(o.get("order_id", "")).startswith("G4")]
        self.assertEqual(a_orders, [])

    # ---------- G5 预览后店铺状态变化（spec 5.6.1 规则 5） ----------
    def test_g5_store_disabled_after_preview(self):
        st, r = self._upload_only(self.admin, "/api/import/wechat-orders",
                                  self.a_store, "G5")
        self.assertEqual(st, 200, r)
        upload_id = r["data"]["upload_id"]
        # 预览成功后停用目标店铺 → 确认被拒
        st, r = self.admin.post(f"/api/stores/{self.a_store}/status",
                                {"action": "disable"})
        self.assertEqual(st, 200, r)
        st, r = self.admin.post("/api/import/confirm", {
            "upload_id": upload_id, "target_store_id": self.a_store})
        self.assertEqual(st, 400, r)
        self.assertIn("该店铺已停用", r["message"])
        # 恢复现场：启用店铺 + 成员重登（停用已吊销其会话；
        # 更新类属性，避免实例属性随用例结束丢失）
        st, r = self.admin.post(f"/api/stores/{self.a_store}/status",
                                {"action": "enable"})
        self.assertEqual(st, 200, r)
        type(self).a = self.env.login_client("px_pick_a", PW_A)
        # 无批次写入
        names = [b.get("file_name") for b in _store_records(
            self.env, self.a_store, "import_records")]
        self.assertNotIn("G5.xlsx", names)

    # ---------- G6 预览回显三字段（验收项 10 的 admin 视角） ----------
    def test_g6_preview_echoes_target_fields(self):
        # member 本人视角
        st, r = self._upload_only(self.a, "/api/import/wechat-orders",
                                  self.a_store, "G6A")
        self.assertEqual(st, 200, r)
        d = r["data"]
        self.assertEqual(d["target_store_id"], self.a_store)
        self.assertEqual(d["target_store_name"], "选店A铺")
        self.assertEqual(d["target_owner_username"], "px_pick_a")
        # admin 代传视角
        st, r = self._upload_only(self.admin, "/api/import/supplier-costs",
                                  self.b_store, "G6B")
        self.assertEqual(st, 200, r)
        d = r["data"]
        self.assertEqual(d["target_store_id"], self.b_store)
        self.assertEqual(d["target_store_name"], "选店B铺")
        self.assertEqual(d["target_owner_username"], "px_pick_b")

    # ---------- G7 批次店铺名称快照（验收项 15） ----------
    def test_g7_batch_store_name_snapshot(self):
        st, r = self._import_ok(self.b, "/api/import/wechat-orders",
                                self.b_store, "G7")
        self.assertEqual(st, 200, r)
        # 改名后批次快照保持导入时名称
        st, r = self.admin.post(f"/api/stores/{self.b_store}/rename",
                                {"name": "选店B铺改名"})
        self.assertEqual(st, 200, r)
        batch = next(b for b in _store_records(
            self.env, self.b_store, "import_records")
            if b.get("file_name") == "G7.xlsx")
        self.assertEqual(batch["target_store_name"], "选店B铺")
        self.assertEqual(batch["target_store_id"], self.b_store)
        # 批次列表 API 回退/快照口径（b 视角本店直查）
        st, r = self.b.get("/api/import/batches?page=1&page_size=20")
        self.assertEqual(st, 200, r)
        listed = next(b for b in r["data"]["batches"]
                      if b["file_name"] == "G7.xlsx")
        self.assertIn(listed.get("store_name"), ("选店B铺", "选店B铺改名"))
        # 改回原名，避免影响后续用例
        st, r = self.admin.post(f"/api/stores/{self.b_store}/rename",
                                {"name": "选店B铺"})
        self.assertEqual(st, 200, r)

    # ---------- G8 admin 代导入留痕（验收项 10） ----------
    def test_g8_admin_proxy_audit_trail(self):
        st, r = self._import_ok(self.admin, "/api/import/supplier-costs",
                                self.b_store, "G8")
        self.assertEqual(st, 200, r)
        batch = next(b for b in _store_records(
            self.env, self.b_store, "import_records")
            if b.get("file_name") == "G8.xlsx")
        # 目标店铺 / 目标成员 / 实际操作人三字段独立且正确
        self.assertEqual(batch["target_store_id"], self.b_store)
        self.assertEqual(batch["target_user_id"], self.b_user_id)
        self.assertEqual(batch["uploaded_by_user_id"], self.admin_user_id)
        self.assertEqual(batch["target_store_name"], "选店B铺")
        # 安全日志含 import.proxy 事件
        events = _security_events(self.env)
        self.assertTrue(any(
            e.get("event") == "import.proxy"
            and (e.get("detail") or {}).get("target_store_id") == self.b_store
            for e in events))
        # 成本数据实际写入目标店铺
        costs = _store_records(self.env, self.b_store, "supplier_costs")
        self.assertGreater(len(costs), 0)

    # ---------- G9 存量批次兼容（spec 5.7.3 场景 1） ----------
    def test_g9_legacy_batch_without_snapshot(self):
        fp = os.path.join(self.env.data_dir, "stores", self.a_store,
                          "data", BUSINESS_FILES["import_records"]["file"])
        with open(fp, encoding="utf-8") as f:
            doc = json.load(f)
        legacy = {
            "batch_id": "B_LEGACY_G9", "source_type": "wechat_orders",
            "file_name": "legacy.xlsx", "file_hash": "legacyhash",
            "imported_at": "2026-01-01T00:00:00", "status": "success",
            "statistics": {"total_rows": 1, "success_rows": 1,
                           "updated_rows": 0, "skipped_rows": 0,
                           "duplicate_rows": 0, "error_rows": 0},
        }
        doc.setdefault("records", []).append(legacy)
        with open(fp, "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False)
        # 列表 API 正常返回且 store_name 回退当前店铺名
        st, r = self.a.get("/api/import/batches?page=1&page_size=50")
        self.assertEqual(st, 200, r)
        listed = next(b for b in r["data"]["batches"]
                      if b["batch_id"] == "B_LEGACY_G9")
        self.assertEqual(listed.get("store_name"), "选店A铺")

    # ---------- G10 失败不落库（验收项 14） ----------
    def test_g10_failures_write_nothing(self):
        files_before = {
            sid: (len(_store_records(self.env, sid, "import_records")),
                  len(_store_records(self.env, sid, "orders")),
                  len(_store_records(self.env, sid, "supplier_costs")))
            for sid in (self.a_store, self.b_store)}
        # 各失败形态：缺参 400 / 空串 400 / all 400 / member 他店 403 /
        # admin 不存在 404
        expectations = [
            (self.a, "/api/import/wechat-orders", 400),
            (self.a, "/api/import/wechat-orders?store_id=", 400),
            (self.a, "/api/import/wechat-orders?store_id=all", 400),
            (self.a, f"/api/import/wechat-orders?store_id={self.b_store}",
             403),
            (self.admin,
             "/api/import/wechat-orders?store_id=s_not_exist", 404),
        ]
        for client, path, expected in expectations:
            st, r = client.upload(path, "G10失败探针.xlsx",
                                  make_wx_xlsx("G10"))
            self.assertEqual(st, expected, (path, r))
        # 业务数据与批次零变化
        for sid in (self.a_store, self.b_store):
            self.assertEqual(
                (len(_store_records(self.env, sid, "import_records")),
                 len(_store_records(self.env, sid, "orders")),
                 len(_store_records(self.env, sid, "supplier_costs"))),
                files_before[sid], sid)
        # 未确认的暂存不产生成功批次：取消路径（上传后不确认）→ 仅暂存
        st, r = self._upload_only(self.a, "/api/import/wechat-orders",
                                  self.a_store, "G10C")
        self.assertEqual(st, 200, r)
        batches_now = [b for b in _store_records(
            self.env, self.a_store, "import_records")
            if b.get("file_name") == "G10C.xlsx"]
        self.assertEqual(batches_now, [])

    # ---------- G11 页面回归（验收项 1-5 页面侧，依赖 S2） ----------
    def test_g11_page_store_picker_ui(self):
        st, html, _ = self.admin.get_page("/data-import")
        self.assertEqual(st, 200)
        # 旧顶部目标选择区已删除
        for token in ("upload-store-row", "upload-member", "upload-store",
                      "upload-row-hint"):
            self.assertNotIn(token, html, token)
        # 新弹窗与状态机已就位
        for token in ("openStorePicker", "uploadTask", "选择导入店铺",
                      "请先选择目标店铺", "/static/modal.js",
                      "确认选择",              # 改：按钮文案从"确认并选择文件"改为"确认选择"
                      "file-info-bar",         # 新增：文件信息展示区
                      "handleFileSelected"):   # 新增：编排函数
            self.assertIn(token, html, token)
        # member 视角同样可访问
        st, html, _ = self.a.get_page("/data-import")
        self.assertEqual(st, 200)
        self.assertNotIn("upload-store-row", html)

    # ---------- G12 店铺口径（弹窗数据源，spec 5.2.1） ----------
    def test_g12_store_list_caliber(self):
        # member 视角：仅本人有效店铺，字段满足弹窗渲染
        st, r = self.a.get("/api/stores")
        self.assertEqual(st, 200, r)
        stores = r["data"]["stores"]
        self.assertEqual([s["store_id"] for s in stores], [self.a_store])
        for field in ("store_id", "name", "owner_user_id",
                      "owner_username"):
            self.assertIn(field, stores[0])
        # admin 视角：全部有效成员店铺（A/B 两店）
        st, r = self.admin.get("/api/stores")
        self.assertEqual(st, 200, r)
        admin_ids = {s["store_id"] for s in r["data"]["stores"]}
        self.assertIn(self.a_store, admin_ids)
        self.assertIn(self.b_store, admin_ids)
        # 停用店铺即时从两个视角消失
        st, r = self.admin.post("/api/stores", {
            "name": "选店A二铺", "owner_user_id": self.a_user_id})
        self.assertEqual(st, 200, r)
        store2 = r["data"]["store_id"]
        st, r = self.admin.post(f"/api/stores/{store2}/status",
                                {"action": "disable"})
        self.assertEqual(st, 200, r)
        st, r = self.admin.get("/api/stores")
        self.assertNotIn(store2, {s["store_id"]
                                  for s in r["data"]["stores"]})
        # a 会话已被吊销（停用其店铺触发），重登后本店列表不含停用店
        # （更新类属性，避免实例属性随用例结束丢失）
        type(self).a = self.env.login_client("px_pick_a", PW_A)
        st, r = self.a.get("/api/stores")
        self.assertEqual(st, 200, r)
        self.assertEqual([s["store_id"] for s in r["data"]["stores"]],
                         [self.a_store])


if __name__ == "__main__":
    unittest.main()

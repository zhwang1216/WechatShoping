import os
import shutil
import tempfile
import threading
import unittest
from unittest import mock

from app.data_schemas import BUSINESS_FILES
from app.json_repository import JsonRepository, RepositoryError


class TestJsonRepository(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.repo = JsonRepository(
            data_dir=os.path.join(self.tmp, "data"),
            backup_dir=os.path.join(self.tmp, "backups"),
        )

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def _read_file(self, name):
        with open(self.repo._path(name), encoding="utf-8") as f:
            return f.read()

    # ---------- 初始化 ----------
    def test_ensure_all_creates_all_business_files(self):
        report = self.repo.ensure_all()
        self.assertEqual(sorted(report["created"]), sorted(BUSINESS_FILES.keys()))
        for spec in BUSINESS_FILES.values():
            self.assertTrue(
                os.path.exists(os.path.join(self.repo.data_dir, spec["file"])),
                f"缺少文件 {spec['file']}",
            )
        report2 = self.repo.ensure_all()
        self.assertEqual(report2["created"], [])
        self.assertEqual(report2["repaired"], [])

    def test_read_initial(self):
        self.repo.ensure_all()
        data = self.repo.read("orders")
        self.assertEqual(data["version"], 1)
        self.assertEqual(data["records"], [])
        settings = self.repo.read("settings")
        self.assertIsInstance(settings["settings"], dict)

    def test_corrupted_file_repaired_and_kept_as_corrupt_backup(self):
        self.repo.ensure_all()
        with open(self.repo._path("orders"), "w", encoding="utf-8") as f:
            f.write("{broken json!!!")
        report = self.repo.ensure_all()
        self.assertIn("orders", report["repaired"])
        data = self.repo.read("orders")
        self.assertEqual(data["records"], [])
        corrupts = [f for f in os.listdir(self.repo.backup_dir) if ".corrupt." in f]
        self.assertEqual(len(corrupts), 1)

    # ---------- 读写 ----------
    def test_write_and_read_roundtrip(self):
        self.repo.ensure_all()
        data = self.repo.read("orders")
        data["records"] = [{"order_id": "A1", "pay_amount": 99.5}]
        self.repo.write("orders", data)
        result = self.repo.read("orders")
        self.assertEqual(result["records"][0]["order_id"], "A1")
        self.assertEqual(result["version"], 1)

    def test_write_invalid_data_rejected(self):
        self.repo.ensure_all()
        with self.assertRaises(RepositoryError):
            self.repo.write("orders", {"no_required_key": True})
        with self.assertRaises(RepositoryError):
            self.repo.write(
                "orders", {"version": 1, "updated_at": "x", "records": "not-list"})
        with self.assertRaises(RepositoryError):
            self.repo.write("unknown_file", {})

    # ---------- 原子性: 模拟写入失败, 旧文件不损坏 ----------
    def test_write_failure_keeps_old_file_on_dump_error(self):
        self.repo.ensure_all()
        data = self.repo.read("orders")
        data["records"] = [{"order_id": "KEEP-ME"}]
        self.repo.write("orders", data, backup=False)
        before = self._read_file("orders")

        broken = self.repo.read("orders")
        broken["records"] = [{"order_id": "BROKEN"}]
        with mock.patch("app.json_repository.json.dump",
                        side_effect=OSError("模拟磁盘写入失败")):
            with self.assertRaises(OSError):
                self.repo.write("orders", broken, backup=False)

        self.assertEqual(self._read_file("orders"), before)
        leftovers = [f for f in os.listdir(self.repo.data_dir) if f.startswith(".tmp_")]
        self.assertEqual(leftovers, [])

    def test_write_failure_keeps_old_file_on_replace_error(self):
        self.repo.ensure_all()
        data = self.repo.read("orders")
        data["records"] = [{"order_id": "KEEP-ME"}]
        self.repo.write("orders", data, backup=False)
        before = self._read_file("orders")

        broken = self.repo.read("orders")
        broken["records"] = [{"order_id": "BROKEN"}]
        with mock.patch("os.replace", side_effect=OSError("模拟替换失败")):
            with self.assertRaises(OSError):
                self.repo.write("orders", broken, backup=False)

        self.assertEqual(self._read_file("orders"), before)
        leftovers = [f for f in os.listdir(self.repo.data_dir) if f.startswith(".tmp_")]
        self.assertEqual(leftovers, [])

    # ---------- 备份与恢复 ----------
    def test_write_creates_backup(self):
        self.repo.ensure_all()
        d1 = self.repo.read("orders")
        d1["records"] = [{"order_id": "OLD"}]
        self.repo.write("orders", d1)
        self.assertEqual(len(self.repo.list_backups("orders")), 1)

        d2 = self.repo.read("orders")
        d2["records"] = [{"order_id": "NEW"}]
        self.repo.write("orders", d2)
        backups = self.repo.list_backups("orders")
        self.assertEqual(len(backups), 2)

        restored = self.repo.restore("orders", backups[0]["file"])
        self.assertEqual(restored["records"][0]["order_id"], "OLD")
        self.assertEqual(self.repo.read("orders")["records"][0]["order_id"], "OLD")

    def test_manual_backup(self):
        self.repo.ensure_all()
        filename = self.repo.backup("orders")
        self.assertTrue(filename.startswith("orders."))
        self.assertIn(filename, [b["file"] for b in self.repo.list_backups("orders")])

    def test_restore_missing_file(self):
        self.repo.ensure_all()
        with self.assertRaises(RepositoryError):
            self.repo.restore("orders", "not_exists.json")

    def test_restore_rejects_path_traversal(self):
        self.repo.ensure_all()
        with self.assertRaises(RepositoryError):
            self.repo.restore("orders", "../../etc/hosts")

    def test_prune_backups_keeps_limit(self):
        self.repo.ensure_all()
        for i in range(6):
            self.repo.write(
                "orders",
                {"version": 1, "updated_at": f"2026-01-0{i}", "records": [{"i": i}]},
                backup=True,
            )
        self.assertEqual(len(self.repo.list_backups("orders")), 6)
        removed = self.repo._prune_backups("orders", keep=2)
        self.assertEqual(len(removed), 4)
        self.assertEqual(len(self.repo.list_backups("orders")), 2)

    # ---------- 文件锁与并发 ----------
    def test_concurrent_update_no_lost_write_no_corruption(self):
        self.repo.ensure_all()

        def add_record(data):
            data["records"].append({"seq": len(data["records"]) + 1})

        threads = [
            threading.Thread(target=lambda: self.repo.update("orders", add_record, backup=False))
            for _ in range(10)
        ]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        data = self.repo.read("orders")
        self.assertEqual(len(data["records"]), 10)
        leftovers = [f for f in os.listdir(self.repo.data_dir) if f.startswith(".tmp_")]
        self.assertEqual(leftovers, [])

    def test_update_mutator_return_none_keeps_data(self):
        self.repo.ensure_all()

        def bad_mutator(data):
            data["records"] = [{"x": 1}]

        self.repo.update("orders", bad_mutator, backup=False)
        self.assertEqual(self.repo.read("orders")["records"], [{"x": 1}])

    def test_update_rejects_invalid_result(self):
        self.repo.ensure_all()
        with self.assertRaises(RepositoryError):
            self.repo.update("orders", lambda d: "invalid")


if __name__ == "__main__":
    unittest.main()

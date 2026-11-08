"""导入服务：预览、确认导入、批次查询、去重、统计、异常清单导出

流程（对应需求 5.5 / 5.6）：
1. preview_upload: 上传 -> SHA256 -> 模板/表头识别 -> 全量预检 -> 暂存文件 -> 返回预览
2. confirm_import: 用户确认后 -> 重复文件检测(hash+source_type) -> 生成批次 ->
   原始记录留存(含错误行) -> 标准记录 upsert(业务键去重) -> 退款识别 ->
   异常记录 -> 重建退款订单视图 -> 统计返回
3. 微信订单：商品发货=未发货（或无发货列且订单状态为待发货等）的行自动剔除，不写入订单库。
4. 供应商成本：导入时匹配微信订单发货状态，未发货标记无效，无法匹配/待核对标记 pending，均保留记录。
"""
import io
import json
import os
import uuid
from datetime import datetime
from decimal import Decimal

from openpyxl import Workbook

from . import config
from .data_schemas import now_iso
from .errors import (BatchNotFoundError, ImportConflictError, ImportServiceError,
                     ImportValidationError, UploadNotFoundError)
from .excel_export import append_row
from .excel_reader import (analyze_sheet, compute_sha256, iter_data_rows,
                           load_workbook_from_bytes, preview_rows, raw_fields_of)
from .import_schemas import (MAPPING_VERSION, SOURCE_TYPES, TEMPLATE_VERSION,
                             get_template)
from .normalizer import (build_supplier_cost_line, build_wechat_order_line,
                          money_str)
from .cost_validity import (
    apply_validity_to_cost,
    build_orders_lookup,
    evaluate_supplier_cost_validity,
    revalidate_supplier_costs_for_order_ids,
)
from .profit_service import is_unshipped_order
from .refund_view import build_refund_views
from .time_range import in_range

_EXCEL_EXTS = (".xlsx", ".xls")


def _is_excel_ext(filename):
    return filename.lower().endswith(_EXCEL_EXTS)


def _business_key(source_type, record):
    """业务唯一键：用于跨批次 upsert 与文件内去重（禁止重复累计）"""
    product = record.get("product") or {}
    if source_type == "wechat_orders":
        return "|".join([
            record.get("order_id") or "",
            product.get("sku_code") or "",
            product.get("title") or "",
        ])
    tracking = record.get("tracking") or {}
    return "|".join([
        record.get("main_order_id") or "",
        record.get("supplier_order_id") or "",
        product.get("sku_code") or "",
        product.get("product_code") or "",
        product.get("title") or "",
        tracking.get("tracking_number") or "",
    ])


def _data_file_of(source_type):
    return "orders" if source_type == "wechat_orders" else "supplier_costs"


_WECHAT_CELLS_TO_FIELDS = {
    "order_time": [("order_time", None)],
    "order_status": [("order_status", None), ("source_status", None)],
    "product_quantity": [("product", "quantity")],
    "product_title": [("product", "title")],
    "sku_attribute": [("product", "sku_attribute")],
    "sku_code": [("product", "sku_code")],
    "actual_received_amount": [("amounts", "actual_received_amount")],
    "actual_paid_amount": [("amounts", "actual_paid_amount")],
    "shipping_fee": [("amounts", "shipping_fee")],
    "product_total_price": [("amounts", "product_total"), ("product", "actual_total_price")],
    "technical_service_fee": [("amounts", "technical_service_fee")],
    "refund_amount": [("amounts", "refund_amount"), ("refund", "refund_amount")],
}
_SUPPLIER_CELLS_TO_FIELDS = {
    "main_order_id": [("main_order_id", None)],
    "supplier_order_id": [("supplier_order_id", None)],
    "created_at": [("created_at", None)],
    "product_quantity": [("product", "quantity")],
    "product_title": [("product", "title")],
    "product_code": [("product", "product_code")],
    "sku_code": [("product", "sku_code")],
    "sku_id": [("product", "sku_id")],
    "total_amount": [("cost_breakdown", "total_amount")],
    "product_supply_cost": [("cost_breakdown", "product_supply_cost")],
    "courier_fee": [("cost_breakdown", "courier_fee")],
    "foam_material_fee": [("cost_breakdown", "foam_material_fee")],
    "surcharge": [("cost_breakdown", "surcharge")],
    "fuel_or_tianjin_packing_fee": [("cost_breakdown", "fuel_or_tianjin_packing_fee")],
    "courier_fee_total": [("cost_breakdown", "courier_fee_total")],
    "platform_subsidy": [("platform_subsidy", None)],
    "supplier_actual_received": [("supplier_actual_received", None)],
    "shipping_fee": [("shipping_fee", None)],
    "carrier": [("tracking", "carrier")],
    "tracking_number": [("tracking", "tracking_number")],
}
_IMPORT_SYSTEM_FIELDS = {
    "record_id", "source_record_id", "batch_id", "business_key",
    "created_at", "updated_at", "match_status", "data_quality",
    "match_key_normalized", "order_level_amounts_duplicated",
    "refund_deduction_status", "total_amount_semantics",
    "cost_granularity", "duplicate_total_amount_risk", "created_at_ts",
    "refund", "tracking", "business_validity",
}


class ImportService:
    def __init__(self, repo, op_log, upload_dir=None):
        self.repo = repo
        self.op_log = op_log
        self.upload_dir = upload_dir or config.UPLOAD_DIR
        os.makedirs(self.upload_dir, exist_ok=True)

    def precheck(self, source_type, filename, data):
        """文件初步校验：仅识别结构，不暂存、不写业务数据（供导入流程前置校验）"""
        if source_type not in SOURCE_TYPES:
            raise ImportValidationError(f"未知数据源类型: {source_type}")
        if not isinstance(filename, str) or not _is_excel_ext(filename):
            raise ImportValidationError("仅支持 .xlsx 或 .xls 文件")
        file_hash = compute_sha256(data)[:16]
        try:
            wb = load_workbook_from_bytes(data)
        except ImportValidationError:
            raise
        except Exception as e:
            raise ImportValidationError(f"文件无法读取: {e}")
        sheets = wb.sheetnames
        if not sheets:
            raise ImportValidationError("文件不包含任何工作表")
        sheet_info = []
        for name in sheets:
            ws = wb[name]
            sheet_info.append({
                "name": name,
                "rows": ws.max_row or 0,
                "columns": ws.max_column or 0,
            })
        return {
            "file_name": filename,
            "file_size": len(data),
            "file_hash": file_hash,
            "source_type": source_type,
            "sheets": sheet_info,
            "readable": True,
        }

    # ==================== 预览 ====================
    def preview_upload(self, source_type, filename, data, operator="admin",
                       target_store_id=None):
        if source_type not in SOURCE_TYPES:
            raise ImportServiceError(f"未知数据源类型: {source_type}")
        if not filename or not _is_excel_ext(str(filename)):
            raise ImportValidationError("仅支持 .xlsx 或 .xls 文件")
        analysis = self._analyze(source_type, str(filename), data)

        upload_id = uuid.uuid4().hex
        with open(os.path.join(self.upload_dir, upload_id + ".xlsx"), "wb") as f:
            f.write(data)
        meta = {
            "upload_id": upload_id,
            "source_type": source_type,
            "filename": str(filename),
            "file_hash": analysis["file_hash"],
            "file_size": len(data),
            "created_at": now_iso(),
            "status": "previewed",
            "target_store_id": target_store_id,
        }
        with open(os.path.join(self.upload_dir, upload_id + ".meta.json"),
                  "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

        self.op_log.append("import.preview", detail={
            "upload_id": upload_id, "source_type": source_type,
            "filename": str(filename), "file_hash": analysis["file_hash"],
        }, operator=operator)

        preview = dict(analysis["preview"])
        preview.update({
            "upload_id": upload_id,
            "source_type": source_type,
            "file_name": str(filename),
            "file_size": len(data),
            "duplicate_file": self._find_duplicate_batch(source_type, analysis["file_hash"]),
            "confirm_hint": "数据 will only enter the official store after confirming the import",
        })
        return preview

    # ==================== 确认导入 ====================
    def confirm_import(self, upload_id, force=False, operator="admin",
                       target_store_id=None, target_user_id=None,
                       uploaded_by_user_id=None, target_store_name=None):
        data_path = os.path.join(self.upload_dir, upload_id + ".xlsx")
        meta_path = os.path.join(self.upload_dir, upload_id + ".meta.json")
        if not (os.path.exists(data_path) and os.path.exists(meta_path)):
            raise UploadNotFoundError("上传文件不存在或已过期，请重新上传")
        with open(meta_path, encoding="utf-8") as f:
            meta = json.load(f)
        with open(data_path, "rb") as f:
            data = f.read()
        source_type = meta["source_type"]

        # 预览-确认目标店铺一致性: 防止确认时改道他店（存量 meta 无该字段则跳过比对）
        meta_store = meta.get("target_store_id")
        if (meta_store and target_store_id
                and meta_store != target_store_id):
            raise ImportValidationError(
                "目标店铺与预览时不一致，请重新上传")

        analysis = self._analyze(source_type, meta["filename"], data)
        duplicate = self._find_duplicate_batch(source_type, analysis["file_hash"])
        if duplicate and not force:
            raise ImportConflictError(
                f"该文件已导入过（批次 {duplicate['batch_id']}），数据不会重复累计；"
                f"如需重新处理请选择强制重新导入", data=duplicate)

        batch_id = "batch_" + datetime.now().strftime("%Y%m%d%H%M%S") + "_" + uuid.uuid4().hex[:8]
        now = now_iso()
        stats = dict(analysis["statistics"])
        data_file = _data_file_of(source_type)

        existing = self.repo.read(data_file)["records"]
        key_index = {_business_key(source_type, rec): i for i, rec in enumerate(existing)}
        raw_all = self.repo.read("raw_records")["records"]
        refund_details = self.repo.read("refund_details")["records"]
        refund_index = {rd.get("order_line_record_id"): i
                        for i, rd in enumerate(refund_details) if rd.get("order_line_record_id")}
        anomalies = self.repo.read("anomalies")["records"]
        open_anomaly_keys = {(a.get("order_id"), a.get("exception_type"))
                             for a in anomalies if a.get("status") == "pending_review"}

        new_raw = []
        new_anomalies = []
        overwrite_log = []
        wechat_affected_order_ids = set()
        supplier_orders_by_id = (build_orders_lookup(self.repo.read("orders")["records"])
                                 if source_type == "supplier_cost" else None)

        def add_anomaly(order_id, etype, severity, message, related_ids):
            key = (order_id, etype)
            if key in open_anomaly_keys:
                return
            open_anomaly_keys.add(key)
            new_anomalies.append({
                "record_id": "exc_" + uuid.uuid4().hex,
                "batch_id": batch_id,
                "reconciliation_id": None,
                "order_id": order_id,
                "exception_type": etype,
                "severity": severity,
                "message": message,
                "related_record_ids": related_ids,
                "status": "pending_review",
                "resolution": None,
                "created_at": now,
            })

        for rec in analysis["row_results"]:
            if rec.get("empty"):
                continue
            raw_record_id = "raw_" + uuid.uuid4().hex
            raw = {
                "record_id": raw_record_id,
                "batch_id": batch_id,
                "source_type": source_type,
                "sheet_name": rec["sheet_name"],
                "row_number": rec["row_number"],
                "raw_fields": rec["raw_fields"],
                "validation": {"status": "passed", "errors": [], "warnings": []},
                "normalized_record_id": None,
                "created_at": now,
            }

            if rec["outcome"] == "error":
                raw["validation"]["status"] = "error"
                raw["validation"]["errors"] = rec["errors"]
                for err in rec["errors"]:
                    order_id = (rec.get("cells", {}).get("order_id")
                                or rec.get("cells", {}).get("main_order_id")
                                or rec.get("cells", {}).get("supplier_order_id"))
                    order_id = str(order_id).strip() if order_id else None
                    add_anomaly(order_id, err["type"], "high", err["message"],
                                [raw_record_id])
                self.op_log.append("import_validation_failed", detail={
                    "batch_id": batch_id,
                    "source_type": source_type,
                    "store_id": target_store_id,
                    "row_number": rec["row_number"],
                    "errors": rec["errors"],
                }, operator=operator)
                new_raw.append(raw)
                continue

            if rec["outcome"] == "duplicate":
                raw["validation"]["status"] = "warning"
                raw["validation"]["warnings"] = [{
                    "type": "DUPLICATE_ROW_IN_FILE",
                    "severity": "low",
                    "message": "文件内业务键重复，已跳过（保留首条）",
                }]
                new_raw.append(raw)
                continue

            if rec.get("outcome") == "skipped_unshipped":
                raw["validation"]["status"] = "skipped"
                raw["validation"]["warnings"] = [{
                    "type": "UNSHIPPED_ORDER_SKIPPED",
                    "severity": "low",
                    "message": "商品未发货，导入时已自动剔除，不写入订单库",
                }]
                new_raw.append(raw)
                continue

            line = rec["line"]
            line["batch_id"] = batch_id
            line["source_record_id"] = raw_record_id
            line["business_key"] = rec["business_key"]
            if rec["outcome"] == "update":
                old = existing[key_index[rec["business_key"]]]
                line, overwritten = self._merge_fields(
                    old, line, rec.get("cells", {}), source_type, batch_id, operator)
                if overwritten:
                    overwrite_log.append({
                        "business_key": rec["business_key"],
                        "fields": overwritten,
                    })
                line["record_id"] = old.get("record_id") or "rec_" + uuid.uuid4().hex
                line["business_key"] = rec["business_key"]
                line["batch_id"] = batch_id
                line["source_record_id"] = raw_record_id
                # created_at 保留业务时间语义：优先新解析值（文件修正），
                # 回退旧记录值，最后回退导入时刻（微信行无业务创建时间字段）
                line["created_at"] = (line.get("created_at")
                                      or old.get("created_at") or now)
                existing[key_index[rec["business_key"]]] = line
            else:
                prefix = "wx_" if source_type == "wechat_orders" else "sup_"
                line["record_id"] = prefix + uuid.uuid4().hex
                # 供应商行保留"订单创建时间"业务字段，禁止被导入时刻覆盖；
                # 微信行 created_at 为记录时间戳
                if not line.get("created_at"):
                    line["created_at"] = now
                existing.append(line)
            line["updated_at"] = now
            raw["normalized_record_id"] = line["record_id"]

            if source_type == "supplier_cost" and supplier_orders_by_id is not None:
                new_bv, changed, prev_bv = apply_validity_to_cost(
                    line, supplier_orders_by_id, batch_id=batch_id, evaluated_at=now)
                if changed or not prev_bv:
                    detail = {
                        "supplier_record_id": line.get("record_id"),
                        "business_key": rec["business_key"],
                        "batch_id": batch_id,
                        "store_id": target_store_id,
                        "before": prev_bv or None,
                        "after": new_bv,
                    }
                    evt = ("supplier_cost.validity_invalid"
                           if new_bv.get("validity_status") == "invalid"
                           else "supplier_cost.validity_revised")
                    self.op_log.append(evt, detail=detail, operator=operator)

            if source_type == "wechat_orders" and line.get("order_id"):
                wechat_affected_order_ids.add(line["order_id"])

            if rec["warnings"]:
                raw["validation"]["status"] = "warning"
                raw["validation"]["warnings"] = rec["warnings"]
                order_id = line.get("order_id") or line.get("main_order_id") \
                    or line.get("supplier_order_id")
                for w in rec["warnings"]:
                    add_anomaly(order_id, w["type"], w.get("severity", "low"),
                                w["message"], [raw_record_id, line["record_id"]])

            if source_type == "wechat_orders":
                rf = line["refund"]
                if rf["refund_status"] not in (None, "none"):
                    detail = {
                        "record_id": "rfd_" + uuid.uuid4().hex,
                        "order_line_record_id": line["record_id"],
                        "order_id": line["order_id"],
                        "batch_id": batch_id,
                        "source_record_id": raw_record_id,
                        "sheet_name": rec["sheet_name"],
                        "row_number": rec["row_number"],
                        "refund_amount": rf["refund_amount"],
                        "refund_status": rf["refund_status"],
                        "status_source": rf["status_source"],
                        "refund_amount_granularity": rf.get("granularity") or "line",
                        "refund_deduction_status": "unknown",
                        "raw_after_sale_text": rec["raw_fields"].get("商品售后", ""),
                        "refund_time": None,
                        "refund_reason": None,
                        "created_at": now,
                        "updated_at": now,
                    }
                    old_idx = refund_index.get(line["record_id"])
                    if old_idx is not None:
                        detail["record_id"] = refund_details[old_idx].get(
                            "record_id", detail["record_id"])
                        detail["created_at"] = refund_details[old_idx].get(
                            "created_at", now)
                        refund_details[old_idx] = detail
                    else:
                        refund_details.append(detail)
            new_raw.append(raw)

        for ga in analysis["group_anomalies"]:
            add_anomaly(ga["order_id"], ga["exception_type"], ga["severity"],
                        ga["message"], ga.get("related_record_ids", []))

        batch_status = "partial_success" if stats["error_rows"] else "success"
        batch = {
            "record_id": batch_id,
            "batch_id": batch_id,
            "source_type": source_type,
            "file_name": meta["filename"],
            "file_hash": analysis["file_hash"],
            "file_size": len(data),
            "imported_at": now,
            "imported_by": operator,
            "period": analysis["preview"]["period"],
            "sheet_name": analysis["sheet_analysis"].name,
            "template_version": TEMPLATE_VERSION,
            "mapping_version": MAPPING_VERSION,
            "status": batch_status,
            "statistics": stats,
            "warnings": analysis["hints"],
            "forced": bool(duplicate and force),
            "target_store_id": target_store_id,
            "target_store_name": target_store_name,
            "target_user_id": target_user_id,
            "uploaded_by_user_id": uploaded_by_user_id,
        }

        # ---- 持久化（内存构建完成后逐文件原子写入）----
        def replace_records(records):
            def mutator(d):
                d["records"] = records
                d["updated_at"] = now
            return mutator

        self.repo.update(data_file, replace_records(existing))
        self.repo.update("raw_records", replace_records(raw_all + new_raw))
        if source_type == "wechat_orders":
            self.repo.update("refund_details", replace_records(refund_details))
            revalidate_supplier_costs_for_order_ids(
                self.repo, wechat_affected_order_ids,
                op_log=self.op_log, operator=operator, batch_id=batch_id)
        if new_anomalies:
            self.repo.update("anomalies", replace_records(anomalies + new_anomalies))

        def add_batch(d):
            d["records"].append(batch)
            d["updated_at"] = now
        self.repo.update("import_records", add_batch)

        if source_type == "wechat_orders":
            self._rebuild_refund_order_views()

        os.remove(data_path)
        os.remove(meta_path)

        self.op_log.append("import.confirm", detail={
            "batch_id": batch_id, "source_type": source_type,
            "file_name": meta["filename"], "statistics": stats,
            "status": batch_status, "forced": bool(duplicate and force),
        }, operator=operator)

        if overwrite_log:
            self.op_log.append("import_overwrite", detail={
                "batch_id": batch_id, "source_type": source_type,
                "store_id": target_store_id,
                "overwritten_count": len(overwrite_log),
                "records": overwrite_log,
            }, operator=operator)

        return {
            "batch_id": batch_id,
            "source_type": source_type,
            "status": batch_status,
            "statistics": stats,
            "warnings": analysis["hints"],
            "errors_export_url": f"/api/import/batches/{batch_id}/errors/export",
            "affected_period": self._compute_period(
                source_type, analysis.get("row_results", [])),
        }

    # ==================== 批次查询 ====================
    def list_batches(self, source_type=None, start=None, end=None, keyword=None,
                     page=1, size=50):
        """导入批次列表：按导入时间（imported_at）左闭右开筛选，支持关键词与分页"""
        records = self.repo.read("import_records")["records"]
        rows = [r for r in records
                if not source_type or r.get("source_type") == source_type]
        if keyword:
            kw = keyword.lower()
            rows = [r for r in rows
                    if kw in (r.get("batch_id") or "").lower()
                    or kw in (r.get("file_name") or "").lower()]
        rows = [r for r in rows if in_range(r.get("imported_at"), start, end)]
        rows.sort(key=lambda r: r.get("imported_at") or "", reverse=True)

        stats = {"batch_count": len(rows), "success": 0,
                 "partial_success": 0, "failed": 0,
                 "total_rows": 0, "error_rows": 0}
        for r in rows:
            st = r.get("status")
            if st in stats:
                stats[st] += 1
            s = r.get("statistics") or {}
            stats["total_rows"] += s.get("total_rows") or 0
            stats["error_rows"] += s.get("error_rows") or 0

        total = len(rows)
        size = max(1, int(size))
        page = max(1, int(page))
        start_idx = (page - 1) * size
        return {
            "total": total, "page": page, "size": size,
            "pages": (total + size - 1) // size,
            "batches": rows[start_idx:start_idx + size],
            "statistics": stats,
        }

    def get_batch(self, batch_id):
        for r in self.repo.read("import_records")["records"]:
            if r.get("batch_id") == batch_id:
                return r
        raise BatchNotFoundError(f"导入批次不存在: {batch_id}")

    # ==================== 导入批次管理（管理员） ====================
    _REVOKE_TARGETS = ("orders", "supplier_costs", "refund_details",
                       "raw_records", "anomalies")

    def revoke_preview(self, batch_id):
        """撤销影响预览: 该批次在各业务文件中的记录数 + 受影响订单/月份"""
        batch = self.get_batch(batch_id)
        counts, affected_orders = {}, set()
        for name in self._REVOKE_TARGETS:
            hits = [r for r in self.repo.read(name)["records"]
                    if r.get("batch_id") == batch_id]
            counts[name] = len(hits)
            for r in hits:
                oid = r.get("order_id") or (r.get("raw_fields") or {}).get("订单号")
                if oid:
                    affected_orders.add(oid)
        months = set()
        for r in self.repo.read("orders")["records"]:
            if r.get("batch_id") == batch_id or r["order_id"] in affected_orders:
                t = (r.get("order_time") or "")[:7]
                if t:
                    months.add(t)
        linked = 0
        if affected_orders:
            linked = sum(
                1 for r in self.repo.read("reconciliation_results")["records"]
                if r.get("order_id") in affected_orders)
        return {
            "batch_id": batch_id,
            "source_type": batch.get("source_type"),
            "file_name": batch.get("file_name"),
            "imported_at": batch.get("imported_at"),
            "imported_by": batch.get("imported_by"),
            "counts": counts,
            "affected_order_count": len(affected_orders),
            "affected_orders": sorted(affected_orders),
            "affected_months": sorted(months),
            "linked_reconciliation_count": linked,
            "will_recalc": ["reconciliation_results", "profit_statements",
                            "refund_orders"],
        }

    def delete_batch_history(self, batch_id, operator="admin"):
        """仅删除导入历史记录（import_records）; 不触碰任何业务数据"""
        batch = self.get_batch(batch_id)
        removed = {"import_records": 0}

        def mutator(d):
            before = len(d["records"])
            d["records"] = [r for r in d["records"]
                            if r.get("batch_id") != batch_id]
            removed["import_records"] = before - len(d["records"])
            d["updated_at"] = now_iso()

        self.repo.update("import_records", mutator, backup=True)
        self.op_log.append("import.delete_history", detail={
            "batch_id": batch_id, "file_name": batch.get("file_name"),
            "removed": removed["import_records"],
        }, operator=operator)
        return {"batch_id": batch_id, "removed": removed}

    def revoke_batch(self, batch_id, operator="admin"):
        """撤销导入批次: 删除该批次实际创建的业务记录与关联结果, 重建派生视图

        返回删除统计与受影响月份（供上层触发对账/利润重算）。
        调用方负责先做整店备份并在失败时回滚。
        """
        batch = self.get_batch(batch_id)
        removed = {}
        affected_orders = set()
        months = set()
        for name in self._REVOKE_TARGETS:
            for r in self.repo.read(name)["records"]:
                if r.get("batch_id") == batch_id:
                    oid = r.get("order_id") or (r.get("raw_fields") or {}).get("订单号")
                    if oid:
                        affected_orders.add(oid)
        for r in self.repo.read("orders")["records"]:
            if r.get("batch_id") == batch_id or r.get("order_id") in affected_orders:
                t = (r.get("order_time") or "")[:7]
                if t:
                    months.add(t)

        def drop_from(name):
            def mutator(d):
                keep, drop = [], []
                for r in d["records"]:
                    if r.get("batch_id") == batch_id:
                        drop.append(r)
                    else:
                        keep.append(r)
                removed[name] = len(drop)
                for r in drop:
                    oid = r.get("order_id") or (r.get("raw_fields") or {}).get("订单号")
                    if oid:
                        affected_orders.add(oid)
                d["records"] = keep
                d["updated_at"] = now_iso()
            self.repo.update(name, mutator, backup=True)

        for name in self._REVOKE_TARGETS:
            drop_from(name)

        # 被删订单关联的对账/利润结果同步清除（稍后按月重算）
        if affected_orders:
            for name in ("reconciliation_results", "profit_statements"):
                def clear_linked(d, _name=name):
                    before = len(d["records"])
                    d["records"] = [r for r in d["records"]
                                    if r.get("order_id") not in affected_orders]
                    removed[_name] = before - len(d["records"])
                    d["updated_at"] = now_iso()
                self.repo.update(name, clear_linked, backup=True)

        def drop_batch(d):
            before = len(d["records"])
            d["records"] = [r for r in d["records"]
                            if r.get("batch_id") != batch_id]
            removed["import_records"] = before - len(d["records"])
            d["updated_at"] = now_iso()
        self.repo.update("import_records", drop_batch, backup=True)

        self._rebuild_refund_order_views()
        self.op_log.append("import.revoke", detail={
            "batch_id": batch_id, "file_name": batch.get("file_name"),
            "removed": removed, "affected_order_count": len(affected_orders),
        }, operator=operator)
        return {"batch_id": batch_id, "removed": removed,
                "affected_orders": sorted(affected_orders),
                "recalc_months": sorted(months)}

    def get_batch_rows(self, batch_id, status=None, page=1, size=20):
        self.get_batch(batch_id)
        size = max(1, min(int(size), 100))
        page = max(1, int(page))
        raws = [r for r in self.repo.read("raw_records")["records"]
                if r.get("batch_id") == batch_id]
        if status:
            raws = [r for r in raws
                    if r.get("validation", {}).get("status") == status]
        raws.sort(key=lambda r: r.get("row_number") or 0)
        total = len(raws)
        start = (page - 1) * size
        return {"total": total, "page": page, "size": size,
                "rows": raws[start:start + size]}

    def export_errors(self, batch_id):
        """导出导入异常清单（错误行 + 警告行 + 异常记录）为 Excel"""
        self.get_batch(batch_id)
        raws = [r for r in self.repo.read("raw_records")["records"]
                if r.get("batch_id") == batch_id
                and r.get("validation", {}).get("status") in ("error", "warning")]
        anomalies = [a for a in self.repo.read("anomalies")["records"]
                     if a.get("batch_id") == batch_id]

        wb = Workbook()
        ws = wb.active
        ws.title = "导入异常清单"
        append_row(ws, ["类别", "工作表", "行号", "级别", "异常类型", "详情",
                        "关联订单号", "原始数据"])
        for r in raws:
            v = r.get("validation", {})
            for e in v.get("errors", []):
                append_row(ws, ["原始记录行", r.get("sheet_name"), r.get("row_number"),
                                "错误", e.get("type", ""), e.get("message", ""),
                                (r.get("raw_fields") or {}).get("订单号", ""),
                                json.dumps(r.get("raw_fields") or {},
                                           ensure_ascii=False)[:500]])
            for w in v.get("warnings", []):
                append_row(ws, ["原始记录行", r.get("sheet_name"), r.get("row_number"),
                                "警告", w.get("type", ""), w.get("message", ""),
                                (r.get("raw_fields") or {}).get("订单号", ""),
                                json.dumps(r.get("raw_fields") or {},
                                           ensure_ascii=False)[:500]])
        for a in anomalies:
            append_row(ws, ["异常记录", "", "", "异常", a.get("exception_type", ""),
                            a.get("message", ""), a.get("order_id") or "", ""])
        buf = io.BytesIO()
        wb.save(buf)
        return f"导入异常清单_{batch_id}.xlsx", buf.getvalue()

    # ==================== 内部方法 ====================
    def _analyze(self, source_type, filename, data):
        """完整解析 + 全量预检（不落盘），返回分析结果"""
        template = get_template(source_type)
        wb = load_workbook_from_bytes(data)
        file_hash = compute_sha256(data)

        candidates, fallback = [], []
        for ws in wb.worksheets:
            sa = analyze_sheet(ws, template)
            if sa.header_row:
                fallback.append((ws, sa))
                if not sa.missing_required:
                    candidates.append((ws, sa))
        if not candidates:
            if not fallback:
                raise ImportValidationError(
                    "未找到有效表头：所有工作表前 10 行内无模板字段命中",
                    data={"missing_fields": [template.label(f)
                                             for f in template.required_columns]})
            _, best_sa = max(fallback, key=lambda x: x[1].matched_count)
            missing_labels = [template.label(f) for f in best_sa.missing_required]
            raise ImportValidationError(
                f"表头缺少必需字段: {', '.join(missing_labels)}，已阻止正式导入",
                data={"missing_fields": missing_labels, "sheet": best_sa.name})
        ws, sa = max(candidates, key=lambda x: x[1].matched_count)

        existing_keys = self._existing_business_keys(source_type)
        build = (build_wechat_order_line if source_type == "wechat_orders"
                 else build_supplier_cost_line)
        stats = {"total_rows": 0, "success_rows": 0, "error_rows": 0,
                 "duplicate_rows": 0, "updated_rows": 0, "skipped_rows": 0,
                 "excluded_unshipped_rows": 0, "warning_rows": 0,
                 "invalid_cost_rows": 0, "pending_cost_rows": 0}
        orders_by_id = None
        if source_type == "supplier_cost":
            orders_by_id = build_orders_lookup(
                self.repo.read("orders")["records"])
        row_results = []
        seen = set()
        for r, cells in iter_data_rows(ws, sa):
            if cells is None:
                stats["skipped_rows"] += 1
                row_results.append({"row_number": r, "empty": True,
                                    "outcome": "empty"})
                continue
            stats["total_rows"] += 1
            line, errors, warnings = build(cells)
            rec = {
                "row_number": r, "empty": False, "line": line,
                "errors": errors, "warnings": warnings, "cells": cells,
                "raw_fields": raw_fields_of(sa, cells), "sheet_name": sa.name,
            }
            if errors:
                stats["error_rows"] += 1
                rec["outcome"] = "error"
            elif (source_type == "wechat_orders"
                  and is_unshipped_order(line.get("order_status"),
                                         line.get("ship_status"))):
                stats["excluded_unshipped_rows"] += 1
                stats["skipped_rows"] += 1
                rec["outcome"] = "skipped_unshipped"
            elif (source_type == "supplier_cost" and orders_by_id is not None):
                preview_bv = evaluate_supplier_cost_validity(line, orders_by_id)
                line["business_validity_preview"] = preview_bv
                if preview_bv.get("validity_status") == "invalid":
                    stats["invalid_cost_rows"] += 1
                elif preview_bv.get("validity_status") == "pending":
                    stats["pending_cost_rows"] += 1
                key = _business_key(source_type, line)
                rec["business_key"] = key
                if key in seen:
                    stats["duplicate_rows"] += 1
                    rec["outcome"] = "duplicate"
                elif key in existing_keys:
                    stats["updated_rows"] += 1
                    rec["outcome"] = "update"
                else:
                    stats["success_rows"] += 1
                    rec["outcome"] = "insert"
                seen.add(key)
            else:
                key = _business_key(source_type, line)
                rec["business_key"] = key
                if key in seen:
                    stats["duplicate_rows"] += 1
                    rec["outcome"] = "duplicate"
                elif key in existing_keys:
                    stats["updated_rows"] += 1
                    rec["outcome"] = "update"
                else:
                    stats["success_rows"] += 1
                    rec["outcome"] = "insert"
                seen.add(key)
            if warnings:
                stats["warning_rows"] += 1
            row_results.append(rec)

        hints, group_anomalies = self._apply_group_checks(source_type, row_results)
        if source_type == "wechat_orders" and stats["excluded_unshipped_rows"]:
            hints.insert(
                0,
                f"已自动剔除未发货订单 {stats['excluded_unshipped_rows']} 行（不写入系统）",
            )
        if source_type == "supplier_cost" and (
                stats.get("invalid_cost_rows") or stats.get("pending_cost_rows")):
            hints.insert(
                0,
                f"供应商成本有效性预检：无效 {stats.get('invalid_cost_rows', 0)} 行"
                f"（未发货不参与统计），待核对 {stats.get('pending_cost_rows', 0)} 行",
            )
        period = self._compute_period(source_type, row_results)
        preview = {
            "file_hash": file_hash,
            "sheets": [{"name": s.title, "rows": s.max_row,
                        "columns": s.max_column} for s in wb.worksheets],
            "selected_sheet": sa.name,
            "header_row": sa.header_row,
            "columns": sa.columns,
            "mapping": {sa.header_by_col[c]: canonical
                        for canonical, c in sa.column_map.items()},
            "mapping_version": MAPPING_VERSION,
            "template_version": TEMPLATE_VERSION,
            "merged_cells": sa.merged_cells,
            "merged_cells_warning": (
                f"检测到合并单元格 {', '.join(sa.merged_cells)}，仅左上角单元格有值，请核对数据完整性"
                if sa.merged_cells else None),
            "preview_rows": preview_rows(ws, sa, limit=config.PREVIEW_ROWS),
            "precheck": stats,
            "granularity_hints": hints,
            "period": period,
            "sample_errors": [
                {"row_number": rec["row_number"],
                 "messages": [e["message"] for e in rec["errors"]]}
                for rec in row_results if rec.get("errors")][:10],
        }
        return {"file_hash": file_hash, "sheet_analysis": sa,
                "row_results": row_results, "statistics": stats,
                "hints": hints, "group_anomalies": group_anomalies,
                "preview": preview}

    @staticmethod
    def _apply_group_checks(source_type, row_results):
        """粒度/重复检查：订单级金额重复、退款粒度、供应商总金额重复风险"""
        hints, group_anomalies = [], []
        valid = [rec for rec in row_results
                 if rec.get("line")
                 and rec.get("outcome") not in ("skipped_unshipped", "empty")]
        if source_type == "wechat_orders":
            groups = {}
            for rec in valid:
                groups.setdefault(rec["line"]["order_id"], []).append(rec)
            for order_id, recs in groups.items():
                if len(recs) > 1:
                    received = {rec["line"]["amounts"]["actual_received_amount"]
                                for rec in recs}
                    if len(received) == 1:
                        for rec in recs:
                            rec["line"]["order_level_amounts_duplicated"] = True
                        hints.append(
                            f"订单 {order_id} 有 {len(recs)} 行商品明细，检测到订单级金额在多行重复，"
                            f"汇总时按订单去重只计一次")
                    else:
                        for rec in recs:
                            rec["line"]["order_level_amounts_duplicated"] = "conflict"
                        hints.append(
                            f"订单 {order_id} 的订单级金额在多行存在不同值，将生成异常待核对")
                        group_anomalies.append({
                            "exception_type": "DUPLICATE_ORDER_AMOUNT",
                            "severity": "high",
                            "order_id": order_id,
                            "message": f"订单 {order_id} 订单级金额（实收）在多行冲突: "
                                       f"{sorted(received)}，已按唯一值计入一次（RC-004）",
                        })
                    refunds = [rec["line"]["refund"] for rec in recs]
                    amts = {rf["refund_amount"] for rf in refunds
                            if rf["refund_amount"] not in (None, "0.00")}
                    if len(amts) == 1 and len(refunds) > 1:
                        for rf in refunds:
                            rf["granularity"] = "unknown"
                        hints.append(
                            f"订单 {order_id} 退款金额在多行相同（{list(amts)[0]}），"
                            f"无法判断订单级/商品行级粒度，已标记待核对")
                        group_anomalies.append({
                            "exception_type": "REFUND_GRANULARITY_UNKNOWN",
                            "severity": "medium",
                            "order_id": order_id,
                            "message": f"订单 {order_id} 退款金额 {list(amts)[0]} 在多行重复，"
                                       f"粒度未知（订单级/商品行级），不参与自动利润终算",
                        })
                    else:
                        for rf in refunds:
                            rf["granularity"] = "line"
                else:
                    recs[0]["line"]["refund"]["granularity"] = "line"
        else:
            groups = {}
            for rec in valid:
                key = (rec["line"].get("main_order_id")
                       or rec["line"].get("supplier_order_id"))
                groups.setdefault(key, []).append(rec)
            for order_key, recs in groups.items():
                if len(recs) > 1:
                    totals = {rec["line"]["cost_breakdown"]["total_amount"]
                              for rec in recs}
                    if len(totals) == 1 and None not in totals:
                        for rec in recs:
                            rec["line"]["duplicate_total_amount_risk"] = True
                            rec["line"]["cost_granularity"] = "order_level"
                        hints.append(
                            f"订单 {order_key} 总金额 {list(totals)[0]} 在 {len(recs)} 行重复，"
                            f"疑似订单级综合成本，已标记重复风险（对账去重后仅扣一次）")
                        group_anomalies.append({
                            "exception_type": "DUPLICATE_TOTAL_AMOUNT_RISK",
                            "severity": "high",
                            "order_id": order_key,
                            "message": f"订单 {order_key} 总金额 {list(totals)[0]} 在 "
                                       f"{len(recs)} 行重复，标记 duplicate_total_amount_risk，"
                                       f"对账时去重后仅扣减一次",
                        })
        return hints, group_anomalies

    def _merge_fields(self, old, new_line, cells, source_type, batch_id, operator):
        """字段级合并：仅覆盖本次 Excel 实际包含的字段，未包含字段保留历史值。
        被覆盖字段清除 manual_modified_fields 标记、置 data_source=import。
        返回 (merged_record, overwritten_fields_with_before_after)"""
        merged = json.loads(json.dumps(old))
        mapping = (_WECHAT_CELLS_TO_FIELDS if source_type == "wechat_orders"
                   else _SUPPLIER_CELLS_TO_FIELDS)
        overwritten = []
        for cells_key, field_paths in mapping.items():
            if cells_key not in cells:
                continue
            for top, sub in field_paths:
                new_val = (new_line.get(top) if sub is None
                           else (new_line.get(top) or {}).get(sub))
                if sub is None:
                    old_val = merged.get(top)
                    merged[top] = new_val
                    field_label = top
                else:
                    merged.setdefault(top, {})
                    old_val = merged[top].get(sub)
                    merged[top][sub] = new_val
                    field_label = f"{top}.{sub}"
                if old_val != new_val:
                    overwritten.append({
                        "field": field_label,
                        "before": old_val,
                        "after": new_val,
                    })
        merged["updated_at"] = now_iso()
        if batch_id:
            merged["last_import_batch_id"] = batch_id
        if source_type == "wechat_orders":
            if "refund_amount" in cells or "after_sale" in cells:
                merged["refund"] = json.loads(json.dumps(new_line.get("refund", {})))
                merged["refund_deduction_status"] = new_line.get("refund_deduction_status", "unknown")
                if "amounts" in merged and "refund_amount" in (new_line.get("amounts") or {}):
                    merged["amounts"]["refund_amount"] = new_line["amounts"]["refund_amount"]
        else:
            if "carrier" in cells or "tracking_number" in cells:
                merged["tracking"] = json.loads(json.dumps(new_line.get("tracking", {})))
        mf = merged.setdefault("manual_modified_fields", [])
        ds = merged.setdefault("data_source", {})
        for ow in overwritten:
            f = ow["field"]
            if f in mf:
                mf.remove(f)
            ds[f] = "import"
        if any(ow["field"] == "cost_breakdown.total_amount" for ow in overwritten):
            merged.pop("manual_cost_source", None)
            merged.pop("manual_cost_modified_at", None)
            merged.pop("manual_cost_modified_by", None)
        return merged, overwritten

    @staticmethod
    def _compute_period(source_type, row_results):
        field = "order_time" if source_type == "wechat_orders" else "created_at"
        times = [rec["line"].get(field) for rec in row_results
                 if rec.get("line")
                 and rec.get("outcome") not in ("skipped_unshipped", "empty")]
        times = [t for t in times if t]
        if not times:
            return {"start": None, "end": None}
        return {"start": min(times), "end": max(times)}

    def _existing_business_keys(self, source_type):
        keys = set()
        for rec in self.repo.read(_data_file_of(source_type))["records"]:
            keys.add(_business_key(source_type, rec))
        return keys

    def _find_duplicate_batch(self, source_type, file_hash):
        for r in self.repo.read("import_records")["records"]:
            if (r.get("file_hash") == file_hash
                    and r.get("source_type") == source_type
                    and r.get("status") in ("success", "partial_success")):
                return {"batch_id": r.get("batch_id"),
                        "imported_at": r.get("imported_at"),
                        "file_name": r.get("file_name")}
        return None

    def _rebuild_refund_order_views(self):
        """按订单去重重建退款订单汇总视图（派生数据，复用公共构建逻辑）"""
        views = build_refund_views(self.repo.read("orders")["records"])
        now = now_iso()

        def mutator(d):
            d["records"] = views
            d["updated_at"] = now
        self.repo.update("refund_orders", mutator)

"""操作日志结构化展示（只读查询，禁止改写审计历史）"""
import hashlib
import io
import json
from datetime import datetime

from openpyxl import Workbook

from .data_schemas import now_iso
from .excel_export import append_row

MODULE_LABELS = {
    "import": "数据导入",
    "reconciliation": "对账",
    "profit": "利润核算",
    "store": "店铺管理",
    "user": "用户管理",
    "data": "数据读写",
    "ad_expense": "广告费",
    "order": "订单",
    "backup": "备份恢复",
    "export": "导出",
    "auth": "认证",
    "security": "安全审计",
}

RESULT_LABELS = {"success": "成功", "error": "失败"}


def _detail_summary(detail):
    if detail is None:
        return "—"
    if isinstance(detail, str):
        t = detail.strip()
        return t if t else "—"
    if isinstance(detail, dict):
        parts = []
        for key in ("message", "note", "error", "order_id", "batch_id",
                    "store_id", "store_name", "file_name", "row_count",
                    "target_username", "month", "amount"):
            val = detail.get(key)
            if val is None or val == "":
                continue
            if isinstance(val, (dict, list)):
                continue
            parts.append(f"{key}={val}")
        if parts:
            return "；".join(parts[:8])
        text = json.dumps(detail, ensure_ascii=False)
        return text[:500] + ("…" if len(text) > 500 else "")
    return str(detail)[:500]


def _log_record_id(raw, store_id):
    base = "|".join([
        store_id or "",
        raw.get("time") or "",
        raw.get("action") or "",
        json.dumps(raw.get("detail"), ensure_ascii=False, sort_keys=True, default=str),
    ])
    return "log_" + hashlib.sha256(base.encode("utf-8")).hexdigest()[:16]


def structure_log_entry(raw, store_id=None, store_name=None):
    """API 出参契约：结构化字段 + 详情摘要（完整 detail 保留）。"""
    if not isinstance(raw, dict):
        return raw
    action = raw.get("action") or ""
    parts = action.split(".", 1)
    module_key = parts[0] if parts else ""
    type_tail = parts[1] if len(parts) > 1 else action
    detail = raw.get("detail")
    result = raw.get("result") or "success"
    return {
        "record_id": raw.get("record_id") or _log_record_id(raw, store_id),
        "time": raw.get("time"),
        "operator": raw.get("operator"),
        "store_id": store_id,
        "store_name": store_name,
        "action": action,
        "module": module_key,
        "module_label": MODULE_LABELS.get(module_key, module_key or "其他"),
        "operation_type": type_tail or action,
        "operation_type_label": type_tail or action,
        "result": result,
        "result_label": RESULT_LABELS.get(result, result),
        "detail": detail,
        "detail_summary": _detail_summary(detail),
    }


def export_logs_excel(entries, operator="admin"):
    if not entries:
        raise ValueError("当前筛选无日志可导出")
    headers = ["操作时间", "操作人", "所属店铺", "店铺 ID", "操作模块",
               "操作类型", "执行结果", "操作详情摘要"]
    wb = Workbook()
    ws = wb.active
    ws.title = "操作日志"
    append_row(ws, headers)
    for e in entries:
        append_row(ws, [
            e.get("time"),
            e.get("operator"),
            e.get("store_name"),
            e.get("store_id"),
            e.get("module_label"),
            e.get("operation_type_label") or e.get("action"),
            e.get("result_label"),
            e.get("detail_summary"),
        ])
    buf = io.BytesIO()
    wb.save(buf)
    filename = f"操作日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return filename, buf.getvalue()

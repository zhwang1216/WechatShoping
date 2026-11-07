"""业务 JSON 数据文件定义、初始数据与校验规则

核心业务数据文件（阶段2 起）：
- import_records      导入批次记录（import_batch）
- raw_records         原始导入记录（raw_import_record，原样留存，禁止覆盖）
- orders              微信订单标准记录（wechat_order_line，商品行级）
- supplier_costs      供应商成本标准记录（supplier_cost_line，成本行级）
- refund_details      退款识别明细（行级，含状态来源与粒度）
- refund_orders       退款订单汇总视图（派生，可重建）
- anomalies           异常记录（exception_record）
- supplier_cost_items 预留：成本拆分明细（后续阶段使用）
- export_records      数据导出记录
- ad_expenses         月度广告花费记录（允许人工编辑的财务金额）
- other_expenses      月度其他费用（允许人工编辑金额和说明）
- reconciliation_results / profit_statements 后续阶段使用
"""
from datetime import datetime

from . import config

SERVICE_NAME = "微信小店对账与利润核算系统"
VERSION = "0.2.0"


def now_iso():
    return datetime.now().isoformat(timespec="seconds")


def _records_file():
    """通用记录型文件的初始结构"""
    return {"version": 1, "updated_at": now_iso(), "records": []}


def _settings_file():
    return {
        "version": 1,
        "updated_at": now_iso(),
        "settings": {
            "operator_name": "店主",
            "refund_deduction_status": "unknown",
            "auto_save": True,
            "backup_on_write": True,
            "backup_keep_count": config.BACKUP_KEEP_COUNT,
            "amount_tolerance": "0.01",
            "page_size": 50,
            "max_upload_mb": config.MAX_UPLOAD_MB,
            "refund_import_strategy": "snapshot",
            "duplicate_file_policy": "block",
        },
    }


# 业务数据文件注册表
# required_keys: 顶层必需字段; collection_key/type: 核心数据集合及其类型
BUSINESS_FILES = {
    "import_records": {
        "file": "import_records.json",
        "description": "导入批次记录（批次号/文件摘要/统计/状态）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "raw_records": {
        "file": "raw_records.json",
        "description": "原始导入记录（原样留存，含源文件/工作表/行号/批次）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "orders": {
        "file": "orders.json",
        "description": "微信订单标准记录（商品行级 wechat_order_line）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "refund_orders": {
        "file": "refund_orders.json",
        "description": "退款订单汇总视图（派生数据，可重建）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "refund_details": {
        "file": "refund_details.json",
        "description": "退款识别明细（行级，含状态来源/粒度/追溯）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "supplier_costs": {
        "file": "supplier_costs.json",
        "description": "供应商成本标准记录（supplier_cost_line，成本行级）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "supplier_cost_items": {
        "file": "supplier_cost_items.json",
        "description": "预留：成本拆分明细（后续阶段使用）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "reconciliation_results": {
        "file": "reconciliation_results.json",
        "description": "订单与成本对账结果（后续阶段使用）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "profit_statements": {
        "file": "profit_statements.json",
        "description": "利润核算记录（后续阶段使用）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "anomalies": {
        "file": "anomalies.json",
        "description": "异常记录及处理状态（exception_record）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "export_records": {
        "file": "export_records.json",
        "description": "数据导出记录",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "settings": {
        "file": "settings.json",
        "description": "系统设置",
        "required_keys": ["version", "updated_at", "settings"],
        "collection_key": "settings",
        "collection_type": dict,
        "initial": _settings_file,
    },
    "ad_expenses": {
        "file": "ad_expenses.json",
        "description": "月度广告花费记录（允许人工编辑的财务金额）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
    "other_expenses": {
        "file": "other_expenses.json",
        "description": "月度其他费用（允许人工编辑金额和说明）",
        "required_keys": ["version", "updated_at", "records"],
        "collection_key": "records",
        "collection_type": list,
        "initial": _records_file,
    },
}

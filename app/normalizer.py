"""数据清洗与标准化

- 金额：Decimal 解析，两位小数十进制字符串存储（禁用浮点直接计算）
- 日期：支持 openpyxl datetime 及常见字符串格式，解析失败保留原值于 raw_fields 并出异常
- 数量：必须正整数，小数数量进入异常而非静默取整
- 状态：标准化枚举 + 原始状态保留（source_status）
- 退款：从"商品已退款金额 + 商品售后"推导 refund_status，推导结果标记 status_source=derived
"""
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

MONEY_QUANT = Decimal("0.01")

DATETIME_FORMATS = [
    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
    "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d",
    "%Y年%m月%d日 %H:%M:%S", "%Y年%m月%d日",
]

INVISIBLE_CHARS = dict.fromkeys(map(ord, "\u200b\u200c\u200d\ufeff\u00a0"), None)

ORDER_STATUS_MAP = {
    "已完成": "completed", "完成": "completed", "已收货": "completed", "已签收": "completed",
    "已发货": "shipped", "待收货": "shipped", "配送中": "shipped",
    "待发货": "pending_shipment", "待备货": "pending_shipment", "备货中": "pending_shipment",
    "已支付": "paid", "已付款": "paid", "待付款": "pending_payment", "待支付": "pending_payment",
    "已关闭": "closed", "交易关闭": "closed", "已取消": "closed", "订单关闭": "closed",
    "退款中": "refund_processing", "售后中": "refund_processing", "退款处理中": "refund_processing",
    "部分退款": "partial_refunded", "已部分退款": "partial_refunded",
    "已全额退款": "full_refunded", "全额退款": "full_refunded",
}


def clean_text(value):
    """去首尾空格与不可见字符；空值返回 None"""
    if value is None:
        return None
    s = str(value).translate(INVISIBLE_CHARS).strip()
    return s if s else None


def parse_money(value):
    """金额解析 -> (Decimal|None, error_message|None)。None 表示缺失（空单元格）"""
    if value is None:
        return None, None
    if isinstance(value, bool):
        return None, "不是有效数字"
    if isinstance(value, (int, float)):
        try:
            return Decimal(str(value)).quantize(MONEY_QUANT), None
        except InvalidOperation:
            return None, "不是有效数字"
    if isinstance(value, Decimal):
        return value.quantize(MONEY_QUANT), None
    s = str(value).strip().replace(",", "").replace("¥", "").replace("￥", "").replace(" ", "")
    if not s:
        return None, None
    try:
        return Decimal(s).quantize(MONEY_QUANT), None
    except InvalidOperation:
        return None, f"金额格式非法: {value!r}"


def money_str(d):
    """Decimal -> 两位小数字符串；None 保持 None（未知值禁用 0 冒充）"""
    if d is None:
        return None
    return str(d.quantize(MONEY_QUANT))


def parse_datetime(value):
    """日期时间解析 -> ISO 8601 字符串或 None（解析失败）"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.replace(microsecond=0, tzinfo=None).isoformat()
    if isinstance(value, date):
        return value.isoformat()
    s = str(value).strip()
    if not s:
        return None
    for fmt in DATETIME_FORMATS:
        try:
            return datetime.strptime(s, fmt).isoformat()
        except ValueError:
            continue
    return None


def parse_quantity(value, label="数量"):
    """数量解析 -> (int|None, error_message|None)；小数/非正数均报错"""
    if value is None or (isinstance(value, str) and not value.strip()):
        return None, f"{label}为空"
    try:
        d = Decimal(str(value).strip().replace(",", ""))
    except InvalidOperation:
        return None, f"{label}不是数字: {value!r}"
    if d != d.to_integral_value():
        return None, f"{label}含小数，需人工核对: {value!r}"
    q = int(d)
    if q <= 0:
        return None, f"{label}必须为正整数: {value!r}"
    return q, None


def normalize_order_status(raw):
    """订单状态标准化 -> (标准枚举, 是否已知)。原始状态由调用方保留"""
    s = clean_text(raw)
    if not s:
        return "unknown", False
    std = ORDER_STATUS_MAP.get(s)
    if std:
        return std, True
    return "unknown", False


def derive_refund_info(refund_amount, after_sale_text, actual_paid_amount):
    """退款状态推导

    返回 {"refund_status", "status_source", "warnings": [{"type","severity","message"}]}
    - 平台文案直接给出 partial/full/pending/closed -> platform
    - 仅有金额或需结合金额判断 partial/full -> derived
    - 无法判断（有金额但支付金额缺失）-> unknown
    """
    warnings = []
    text = clean_text(after_sale_text) or ""
    amt = refund_amount if refund_amount is not None else Decimal("0")
    paid = actual_paid_amount

    if "全额退款" in text:
        status, source = "full", "platform"
    elif "部分退款" in text:
        status, source = "partial", "platform"
    elif any(k in text for k in ("处理中", "退款中", "售后中", "待审核", "申请中")):
        status, source = "pending", "platform"
    elif "关闭" in text and ("退款" in text or "售后" in text):
        status, source = "closed", "platform"
    elif "退款成功" in text or "已退款" in text:
        if paid is not None:
            status, source = ("full" if amt >= paid else "partial"), "derived"
        elif amt > 0:
            status, source = "unknown", "derived"
        else:
            status, source = "pending", "derived"
    elif amt > 0:
        if paid is None:
            status, source = "unknown", "derived"
            warnings.append({
                "type": "REFUND_STATUS_UNKNOWN",
                "severity": "medium",
                "message": f"存在退款金额 {money_str(amt)} 但缺少支付金额，无法判断部分/全额退款",
            })
        else:
            status, source = ("full" if amt >= paid else "partial"), "derived"
    else:
        status, source = "none", "platform"

    if paid is not None and amt > paid:
        warnings.append({
            "type": "REFUND_EXCEEDS_PAID",
            "severity": "high",
            "message": f"退款金额 {money_str(amt)} 超过实际支付金额 {money_str(paid)}",
        })
    return {"refund_status": status, "status_source": source, "warnings": warnings}


def _money_field(cells, canonical, label, errors, warnings, required=False,
                 negative_warning_type=None):
    """解析单个金额字段并收集错误/警告"""
    d, err = parse_money(cells.get(canonical))
    if err:
        errors.append({"type": "INVALID_AMOUNT", "message": f"{label}{err}"})
        return None
    if d is None and required:
        errors.append({"type": "MISSING_AMOUNT", "message": f"{label}为空（必填）"})
        return None
    if d is not None and d < 0 and negative_warning_type:
        warnings.append({
            "type": negative_warning_type,
            "severity": "medium",
            "message": f"{label}为负数: {money_str(d)}，已保留原值待核对",
        })
    return d


def build_wechat_order_line(cells):
    """微信订单行 -> (标准记录 dict, errors[dict], warnings[dict])

    errors 非空 -> 该行进入错误行隔离（不生成标准记录，原始行仍留存）
    warnings 非空 -> 记录生成，同时生成异常记录
    """
    errors, warnings, issues = [], [], []

    order_id = clean_text(cells.get("order_id"))
    if not order_id:
        errors.append({"type": "ORDER_ID_EMPTY", "message": "订单号为空"})

    order_time_raw = cells.get("order_time")
    order_time = parse_datetime(order_time_raw)
    if order_time is None and order_time_raw not in (None, ""):
        warnings.append({
            "type": "UNPARSEABLE_ORDER_TIME",
            "severity": "medium",
            "message": f"订单下单时间无法解析，原值 {order_time_raw!r} 已留存待核对",
        })
        issues.append("unparseable_order_time")

    source_status = clean_text(cells.get("order_status"))
    std_status, known = normalize_order_status(source_status)
    if source_status and not known:
        warnings.append({
            "type": "UNKNOWN_ORDER_STATUS",
            "severity": "low",
            "message": f"未知订单状态: {source_status}",
        })
        issues.append("unknown_order_status")

    quantity, qerr = parse_quantity(cells.get("product_quantity"), "商品数量")
    if qerr:
        errors.append({"type": "INVALID_QUANTITY", "message": qerr})

    received = _money_field(cells, "actual_received_amount", "订单实际收款金额",
                            errors, warnings, required=True,
                            negative_warning_type="NEGATIVE_AMOUNT")
    paid = _money_field(cells, "actual_paid_amount", "订单实际支付金额",
                        errors, warnings, negative_warning_type="NEGATIVE_AMOUNT")
    shipping = _money_field(cells, "shipping_fee", "订单运费",
                            errors, warnings, negative_warning_type="NEGATIVE_AMOUNT")
    product_total = _money_field(cells, "product_total_price", "商品实际价格(总共)",
                                 errors, warnings, negative_warning_type="NEGATIVE_AMOUNT")

    tech_raw = cells.get("technical_service_fee")
    tech, tech_err = parse_money(tech_raw)
    if tech_err:
        errors.append({"type": "INVALID_AMOUNT", "message": f"技术服务费{tech_err}"})
        tech = None
    else:
        if tech is not None and tech < 0:
            warnings.append({
                "type": "TECH_FEE_NEGATIVE_NORMALIZED",
                "severity": "low",
                "message": f"技术服务费原始值为负 {money_str(tech)}，已标准化为费用正值 {money_str(-tech)}",
            })
            issues.append("technical_service_fee_negative_normalized")
            tech = -tech

    refund_amt, refund_err = parse_money(cells.get("refund_amount"))
    if refund_err:
        errors.append({"type": "INVALID_AMOUNT", "message": f"商品已退款金额{refund_err}"})
        refund_amt = None
    elif refund_amt is not None and refund_amt < 0:
        warnings.append({
            "type": "NEGATIVE_AMOUNT",
            "severity": "medium",
            "message": f"商品已退款金额为负数: {money_str(refund_amt)}",
        })

    refund_info = derive_refund_info(refund_amt, cells.get("after_sale"), paid)
    for w in refund_info["warnings"]:
        if w["type"] == "REFUND_EXCEEDS_PAID":
            issues.append("refund_exceeds_paid")
        warnings.append(w)

    if errors:
        return None, errors, warnings

    line = {
        "record_id": None,
        "source_record_id": None,
        "batch_id": None,
        "business_key": None,
        "order_id": order_id,
        "order_time": order_time,
        "order_status": std_status,
        "source_status": source_status,
        "product": {
            "title": clean_text(cells.get("product_title")),
            "sku_attribute": clean_text(cells.get("sku_attribute")),
            "sku_code": clean_text(cells.get("sku_code")),
            "quantity": quantity,
            "unit_price": None,
            "actual_unit_price": None,
            "actual_total_price": money_str(product_total),
        },
        "amounts": {
            "product_total": money_str(product_total),
            "actual_paid_amount": money_str(paid),
            "actual_received_amount": money_str(received),
            "shipping_fee": money_str(shipping),
            "technical_service_fee": money_str(tech),
            "platform_coupon_discount": None,
            "refund_amount": money_str(refund_amt),
        },
        "refund": {
            "refund_amount": money_str(refund_amt),
            "refund_status": refund_info["refund_status"],
            "status_source": refund_info["status_source"],
            "granularity": None,
            "deduction_status": "unknown",
        },
        "refund_deduction_status": "unknown",
        "match_key_normalized": (order_id or "").replace(" ", ""),
        "order_level_amounts_duplicated": None,
        "data_quality": {
            "status": "warning" if issues else "ok",
            "issues": issues,
        },
        "created_at": None,
        "updated_at": None,
    }
    return line, errors, warnings


def build_supplier_cost_line(cells):
    """供应商成本行 -> (标准记录, errors, warnings)

    口径：total_amount=综合发货成本（唯一扣减字段）；product_supply_cost=产品供货成本（仅分析）。
    """
    errors, warnings, issues = [], [], []

    main = clean_text(cells.get("main_order_id"))
    sub = clean_text(cells.get("supplier_order_id"))
    if not main and not sub:
        errors.append({"type": "MISSING_ORDER_KEYS",
                       "message": "主订单编号与订单编号均为空"})

    created_raw = cells.get("created_at")
    created_at = parse_datetime(created_raw)
    if created_at is None and created_raw not in (None, ""):
        warnings.append({
            "type": "UNPARSEABLE_ORDER_TIME",
            "severity": "medium",
            "message": f"订单创建时间无法解析，原值 {created_raw!r} 已留存待核对",
        })
        issues.append("unparseable_created_at")

    quantity, qerr = parse_quantity(cells.get("product_quantity"), "商品数量")
    if qerr:
        errors.append({"type": "INVALID_QUANTITY", "message": qerr})

    def part_money(canonical, label):
        """分项金额（仅分析用途）：非法值降级为警告 + None"""
        d, err = parse_money(cells.get(canonical))
        if err:
            warnings.append({
                "type": "INVALID_PARTIAL_AMOUNT",
                "severity": "low",
                "message": f"{label}{err}，该分项已置空（分项仅用于分析核对）",
            })
            return None
        if d is not None and d < 0:
            warnings.append({
                "type": "NEGATIVE_AMOUNT",
                "severity": "medium",
                "message": f"{label}为负数: {money_str(d)}",
            })
        return d

    total_amount, total_err = parse_money(cells.get("total_amount"))
    if total_err:
        errors.append({"type": "INVALID_AMOUNT", "message": f"总金额{total_err}"})
        total_amount = None
    elif total_amount is None:
        warnings.append({
            "type": "COST_MISSING",
            "severity": "high",
            "message": "总金额为空（综合成本缺失，该订单利润将不可计算，不会默认按 0 处理）",
        })
        issues.append("missing_total_amount")
    elif total_amount < 0:
        warnings.append({
            "type": "NEGATIVE_AMOUNT",
            "severity": "medium",
            "message": f"总金额为负数: {money_str(total_amount)}",
        })

    supply_cost = part_money("product_supply_cost", "成本")
    courier_total = part_money("courier_fee_total", "快递费合计")
    courier = part_money("courier_fee", "快递费")
    foam = part_money("foam_material_fee", "泡沫耗材费")
    surcharge = part_money("surcharge", "加收")
    fuel = part_money("fuel_or_tianjin_packing_fee", "燃油费/打包费")
    shipping = part_money("shipping_fee", "订单运费")
    subsidy = part_money("platform_subsidy", "平台补贴")
    supplier_received = part_money("supplier_actual_received", "订单实收款")

    if errors:
        return None, errors, warnings

    line = {
        "record_id": None,
        "source_record_id": None,
        "batch_id": None,
        "business_key": None,
        "main_order_id": main,
        "supplier_order_id": sub,
        "created_at": created_at,
        "product": {
            "title": clean_text(cells.get("product_title")),
            "product_code": clean_text(cells.get("product_code")),
            "sku_code": clean_text(cells.get("sku_code")),
            "sku_id": clean_text(cells.get("sku_id")),
            "sku_attribute": None,
            "quantity": quantity,
        },
        "cost_breakdown": {
            "product_supply_cost": money_str(supply_cost),
            "courier_fee": money_str(courier),
            "foam_material_fee": money_str(foam),
            "surcharge": money_str(surcharge),
            "fuel_or_tianjin_packing_fee": money_str(fuel),
            "courier_fee_total": money_str(courier_total),
            "total_amount": money_str(total_amount),
            "currency": "CNY",
        },
        "total_amount_semantics": "composite_cost",
        "cost_granularity": "supplier_line",
        "duplicate_total_amount_risk": None,
        "platform_subsidy": money_str(subsidy),
        "supplier_actual_received": money_str(supplier_received),
        "shipping_fee": money_str(shipping),
        "tracking": {
            "carrier": clean_text(cells.get("carrier")),
            "tracking_number": clean_text(cells.get("tracking_number")),
            "ship_time": None,
            "package_status": None,
        },
        "match_status": "unmatched",
        "data_quality": {
            "status": "warning" if issues else "ok",
            "issues": issues,
        },
        "created_at_ts": None,
        "updated_at": None,
    }
    return line, errors, warnings

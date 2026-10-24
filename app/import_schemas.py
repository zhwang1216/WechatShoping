"""Excel 模板字段映射定义（字段清单来自需求文档 1.2 / 6.3 / 6.4）

重要：字段名称相似不代表口径一致。每个 canonical 字段的语义以下方定义为准：
- 供应商表"总金额" = 综合发货成本（含产品成本、快递、泡沫耗材、加收、燃油/天津打包费），
  是利润计算的唯一直接扣减字段（total_amount_semantics=composite_cost）；
- 供应商表"成本" = 产品供货成本（product_supply_cost），仅用于拆分分析，禁止在利润中再次扣减；
- "快递费合计"与"快递费"等分项为包含关系候选（待确认口径 Q-05），仅用于核对展示。
"""

SOURCE_TYPES = ("wechat_orders", "supplier_cost")

TEMPLATE_VERSION = "wx_v1"
MAPPING_VERSION = "map_v1"


def normalize_header(text):
    """表头规范化：去空白、全角括号转半角、转小写，用于别名匹配"""
    if text is None:
        return ""
    s = str(text).strip()
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace(" ", "").replace("\u3000", "").replace("\t", "")
    return s.lower()


# 微信小店订单 Excel 字段（canonical -> 定义）
WECHAT_ORDER_FIELDS = {
    "order_id": {
        "aliases": ["订单号", "微信订单号", "订单编号"],
        "required_column": True, "label": "订单号",
    },
    "order_time": {
        "aliases": ["订单下单时间", "下单时间", "订单时间", "支付时间"],
        "required_column": True, "label": "订单下单时间",
    },
    "order_status": {
        "aliases": ["订单状态"],
        "required_column": True, "label": "订单状态",
    },
    "actual_received_amount": {
        "aliases": ["订单实际收款金额", "实际收款金额", "订单实收金额"],
        "required_column": True, "label": "订单实际收款金额",
    },
    "product_quantity": {
        "aliases": ["商品数量", "购买数量"],
        "required_column": True, "label": "商品数量",
    },
    "actual_paid_amount": {
        "aliases": ["订单实际支付金额", "实际支付金额"],
        "required_column": False, "label": "订单实际支付金额",
    },
    "shipping_fee": {
        "aliases": ["订单运费", "运费"],
        "required_column": False, "label": "订单运费",
    },
    "technical_service_fee": {
        "aliases": ["技术服务费"],
        "required_column": False, "label": "技术服务费",
    },
    "product_total_price": {
        "aliases": ["商品实际价格(总共)", "商品实际价格（总共）", "商品总价"],
        "required_column": False, "label": "商品实际价格(总共)",
    },
    "refund_amount": {
        "aliases": ["商品已退款金额", "已退款金额"],
        "required_column": False, "label": "商品已退款金额",
    },
    "after_sale": {
        "aliases": ["商品售后", "售后状态", "售后"],
        "required_column": False, "label": "商品售后",
    },
    "product_title": {
        "aliases": ["商品标题", "商品名称"],
        "required_column": False, "label": "商品标题",
    },
    "sku_attribute": {
        "aliases": ["sku属性", "SKU属性", "规格", "商品规格"],
        "required_column": False, "label": "sku属性",
    },
    "sku_code": {
        "aliases": ["sku编码", "SKU编码", "商品编码"],
        "required_column": False, "label": "sku编码",
    },
}

# 供应商及物流成本 Excel 字段（canonical -> 定义）
SUPPLIER_COST_FIELDS = {
    "total_amount": {
        "aliases": ["总金额", "综合总金额"],
        "required_column": True, "label": "总金额",
        "comment": "综合发货成本，利润计算唯一直接扣减字段",
    },
    "product_quantity": {
        "aliases": ["商品数量", "数量"],
        "required_column": True, "label": "商品数量",
    },
    "main_order_id": {
        "aliases": ["主订单编号", "主订单号"],
        "required_column": False, "label": "主订单编号",
        "comment": "与订单编号至少一项必填",
    },
    "supplier_order_id": {
        "aliases": ["订单编号", "子订单编号", "订单号"],
        "required_column": False, "label": "订单编号",
        "comment": "与主订单编号至少一项必填",
    },
    "created_at": {
        "aliases": ["订单创建时间", "创建时间", "下单时间"],
        "required_column": False, "label": "订单创建时间",
    },
    "product_supply_cost": {
        "aliases": ["成本", "供货成本"],
        "required_column": False, "label": "成本",
        "comment": "产品供货成本，仅用于拆分分析，禁止在利润中再次扣减",
    },
    "courier_fee_total": {
        "aliases": ["快递费合计"],
        "required_column": False, "label": "快递费合计",
    },
    "courier_fee": {
        "aliases": ["快递费"],
        "required_column": False, "label": "快递费",
    },
    "foam_material_fee": {
        "aliases": ["泡沫耗材费", "耗材费"],
        "required_column": False, "label": "泡沫耗材费",
    },
    "surcharge": {
        "aliases": ["加收"],
        "required_column": False, "label": "加收",
    },
    "fuel_or_tianjin_packing_fee": {
        "aliases": ["燃油费", "打包费", "燃油费(打包费)", "燃油费（打包费）"],
        "required_column": False, "label": "燃油费/打包费",
    },
    "shipping_fee": {
        "aliases": ["订单运费", "运费"],
        "required_column": False, "label": "订单运费",
    },
    "platform_subsidy": {
        "aliases": ["平台补贴", "补贴"],
        "required_column": False, "label": "平台补贴",
    },
    "supplier_actual_received": {
        "aliases": ["订单实收款", "实收款"],
        "required_column": False, "label": "订单实收款",
        "comment": "仅用于与微信侧金额交叉核验，禁止重复计入收入",
    },
    "product_code": {
        "aliases": ["商品编码"],
        "required_column": False, "label": "商品编码",
    },
    "sku_code": {
        "aliases": ["sku编码", "SKU编码", "sku"],
        "required_column": False, "label": "sku编码",
    },
    "sku_id": {
        "aliases": ["SKUID", "skuid", "SkuId", "sku_id"],
        "required_column": False, "label": "SKUID",
    },
    "product_title": {
        "aliases": ["商品标题", "商品名称"],
        "required_column": False, "label": "商品标题",
    },
    "carrier": {
        "aliases": ["快递公司", "物流公司"],
        "required_column": False, "label": "快递公司",
    },
    "tracking_number": {
        "aliases": ["快递单号", "物流单号", "运单号"],
        "required_column": False, "label": "快递单号",
    },
}


class Template:
    """某数据源类型的字段映射模板"""

    def __init__(self, source_type, fields):
        self.source_type = source_type
        self.fields = fields
        self.alias_lookup = {}
        for canonical, spec in fields.items():
            for alias in spec["aliases"]:
                norm = normalize_header(alias)
                if norm and norm not in self.alias_lookup:
                    self.alias_lookup[norm] = canonical
        self.required_columns = [
            c for c, spec in fields.items() if spec["required_column"]
        ]

    def label(self, canonical):
        return self.fields.get(canonical, {}).get("label", canonical)


TEMPLATES = {
    "wechat_orders": Template("wechat_orders", WECHAT_ORDER_FIELDS),
    "supplier_cost": Template("supplier_cost", SUPPLIER_COST_FIELDS),
}


def get_template(source_type):
    if source_type not in TEMPLATES:
        raise ValueError(f"未知数据源类型: {source_type}")
    return TEMPLATES[source_type]

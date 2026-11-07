"""导出 Excel 时写入单元格。

openpyxl 会把以 = 开头的字符串写成公式。订单号、商品名、运单号来自导入文件，
这里在写入后把 =、+、-、@ 以及制表符、回车开头的文本改回字符串单元格。
"""

_TEXT_PREFIXES = ("=", "+", "-", "@", "\t", "\r")


def append_row(ws, values):
    """追加一行。触发前缀的字符串强制为文本，数字和日期保持原类型。"""
    row = list(values)
    ws.append(row)
    idx = ws.max_row
    for col, value in enumerate(row, 1):
        if isinstance(value, str) and value[:1] in _TEXT_PREFIXES:
            ws.cell(row=idx, column=col).data_type = "s"
    return idx

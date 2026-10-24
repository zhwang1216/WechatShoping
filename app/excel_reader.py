"""Excel 读取与解析：SHA256、工作簿加载、工作表/表头识别、行迭代、预览"""
import hashlib
import io
from datetime import date, datetime

from openpyxl import load_workbook

from .errors import ExcelParseError
from .import_schemas import normalize_header


def compute_sha256(data):
    h = hashlib.sha256()
    h.update(data)
    return h.hexdigest()


def load_workbook_from_bytes(data):
    if not data:
        raise ExcelParseError("文件内容为空")
    if data[:2] != b"PK":
        raise ExcelParseError("不是有效的 .xlsx 文件（仅支持 Excel 2007+ 格式，不支持 .xls/CSV）")
    try:
        return load_workbook(io.BytesIO(data), data_only=True)
    except ExcelParseError:
        raise
    except Exception as e:
        raise ExcelParseError(f"无法解析 Excel 文件: {e}")


def cell_display(value):
    """单元格值转为展示字符串（原始值语义，不做业务清洗）"""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value == int(value) and abs(value) < 1e15:
        return str(int(value))
    return str(value)


class SheetAnalysis:
    """单个工作表的表头识别结果"""

    def __init__(self, name):
        self.name = name
        self.header_row = None      # 表头行号（1-based）
        self.column_map = {}        # canonical -> 列号（1-based）
        self.header_by_col = {}     # 列号 -> 原始表头文本
        self.columns = []           # 原始表头列表（按列序）
        self.missing_required = []  # 缺失的必填 canonical 字段
        self.matched_count = 0
        self.merged_cells = []      # 合并单元格区域描述

    @property
    def usable(self):
        return self.header_row is not None and not self.missing_required


def analyze_sheet(ws, template):
    """在表头前 10 行中寻找命中模板字段最多的行作为表头行"""
    sa = SheetAnalysis(ws.title)
    max_row = min(10, ws.max_row or 1)
    max_col = ws.max_column or 1
    best = None
    for r in range(1, max_row + 1):
        colmap = {}
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if v is None or str(v).strip() == "":
                continue
            canonical = template.alias_lookup.get(normalize_header(v))
            if canonical and canonical not in colmap:
                colmap[canonical] = c
        if len(colmap) >= 2 and (best is None or len(colmap) > len(best[1])):
            best = (r, colmap)
    if best is None:
        return sa
    sa.header_row, sa.column_map = best
    sa.matched_count = len(sa.column_map)
    for c in range(1, max_col + 1):
        v = ws.cell(row=sa.header_row, column=c).value
        text = "" if v is None else str(v)
        sa.columns.append(text)
        if text:
            sa.header_by_col[c] = text
    sa.missing_required = [f for f in template.required_columns if f not in sa.column_map]
    try:
        sa.merged_cells = [str(rng) for rng in ws.merged_cells.ranges]
    except Exception:
        sa.merged_cells = []
    return sa


def iter_data_rows(ws, sa):
    """迭代数据行：yield (行号, {canonical: 原始值})；全空行 yield (行号, None)"""
    if sa.header_row is None:
        return
    rev = {c: canonical for canonical, c in sa.column_map.items()}
    for row in ws.iter_rows(min_row=sa.header_row + 1):
        r = row[0].row
        cells = {}
        has_value = False
        for cell in row:
            canonical = rev.get(cell.column)
            if canonical is None:
                continue
            v = cell.value
            cells[canonical] = v
            if v is not None and str(v).strip() != "":
                has_value = True
        yield r, (cells if has_value else None)


def preview_rows(ws, sa, limit=20):
    """前 N 行预览（原始表头名 -> 展示值）"""
    rows = []
    for r, cells in iter_data_rows(ws, sa):
        if cells is None:
            rows.append({"row_number": r, "empty": True, "cells": {}})
        else:
            display = {}
            for canonical, v in cells.items():
                col = sa.column_map[canonical]
                header = sa.header_by_col.get(col, canonical)
                display[header] = cell_display(v)
            rows.append({"row_number": r, "empty": False, "cells": display})
        if len(rows) >= limit:
            break
    return rows


def raw_fields_of(sa, cells):
    """构造原始列名 -> 原始值展示（用于 raw_import_record.raw_fields）"""
    out = {}
    for canonical, v in cells.items():
        col = sa.column_map[canonical]
        header = sa.header_by_col.get(col, canonical)
        out[header] = cell_display(v)
    return out

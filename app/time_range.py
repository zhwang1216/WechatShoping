"""统一月份/日期筛选区间解析（所有列表接口共用，spec：月份筛选统一规则）

规则：
- 时间区间一律左闭右开 [start, end)：start = 当月第一天 00:00:00，
  end = 下个月第一天 00:00:00（排他上界）。禁止使用"当月最后一天 23:59:59"
  作为唯一结束条件，避免毫秒/秒级精度与跨日边界漏数。
- 参数优先级（同时存在时，从高到低）：
  1. from / to（自定义日期范围）
  2. month（自然月，格式 YYYY-MM）
  3. start / end（旧参数兼容）
- 边界统一输出 "YYYY-MM-DD HH:MM:SS"（空格分隔）：空格 ASCII(0x20) 小于
  'T'(0x54)，因此该边界与 "YYYY-MM-DDTHH:MM:SS"（ISO/T 分隔，normalizer
  标准格式）和空格分隔两种数据的字典序比较结果均正确，日期解析口径一致。
"""
import re
from datetime import datetime, timedelta

MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


class TimeRangeError(ValueError):
    """时间筛选参数格式非法"""


def _check_month(month):
    if not MONTH_RE.match(month or ""):
        raise TimeRangeError(f"month 格式必须为 YYYY-MM，例如 2026-10，收到: {month!r}")
    return month


def _check_date(value, name):
    if not DATE_RE.match(value or ""):
        raise TimeRangeError(f"{name} 格式必须为 YYYY-MM-DD，例如 2026-10-01，收到: {value!r}")
    try:
        datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        raise TimeRangeError(f"{name} 不是有效日期: {value!r}")
    return value


def _next_month_first_day(year, month):
    if month == 12:
        return f"{year + 1}-01-01 00:00:00"
    return f"{year}-{month + 1:02d}-01 00:00:00"


def month_bounds(month):
    """month=2026-10 -> ("2026-10-01 00:00:00", "2026-11-01 00:00:00")"""
    _check_month(month)
    year, mon = int(month[:4]), int(month[5:7])
    return (f"{month}-01 00:00:00", _next_month_first_day(year, mon))


def date_bounds(date_from=None, date_to=None):
    """自定义范围：from 当天零点（含），to 次日零点（排他）"""
    start = None
    end = None
    if date_from is not None:
        _check_date(date_from, "from")
        start = f"{date_from} 00:00:00"
    if date_to is not None:
        _check_date(date_to, "to")
        end = f"{datetime.strptime(date_to, '%Y-%m-%d') + timedelta(days=1):%Y-%m-%d} 00:00:00"
    if start and end and start >= end:
        raise TimeRangeError(f"from({date_from}) 不得晚于 to({date_to})")
    return start, end


def resolve_range(month=None, date_from=None, date_to=None,
                  start=None, end=None):
    """按优先级解析为左闭右开区间 (start, end_exclusive)；全部为空返回 (None, None)

    优先级：from/to > month > start/end（旧参数）。
    旧参数 end 原为闭区间语义，此处统一转换为次日零点排他上界。
    """
    if date_from is not None or date_to is not None:
        return date_bounds(date_from, date_to)
    if month:
        return month_bounds(month)
    return legacy_bounds(start, end)


def legacy_bounds(start=None, end=None):
    """兼容旧 start/end 参数：10 位日期补零点/次日零点，统一左闭右开"""
    if start is not None and DATE_RE.match(start or ""):
        _check_date(start, "start")
        start = f"{start} 00:00:00"
    if end is not None and DATE_RE.match(end or ""):
        _check_date(end, "end")
        end = f"{datetime.strptime(end, '%Y-%m-%d') + timedelta(days=1):%Y-%m-%d} 00:00:00"
    if start and end and start >= end:
        raise TimeRangeError(f"start({start}) 不得晚于 end({end})")
    return start, end


def norm_value(value):
    """数据值规范化：date-only 补 T00:00:00，便于与空格边界正确比较"""
    if value is None:
        return None
    v = str(value).strip()
    if len(v) == 10 and DATE_RE.match(v):
        v += "T00:00:00"
    return v


def in_range(value, start, end):
    """左闭右开判断：value ∈ [start, end)。None/空值在指定了边界时排除"""
    v = norm_value(value)
    if v is None or v == "":
        return start is None and end is None
    if start is not None and v < start:
        return False
    if end is not None and v >= end:
        return False
    return True


def range_meta(month=None, date_from=None, date_to=None, start=None, end=None):
    """解析区间并返回元信息（供接口响应附加，前端展示当前筛选条件）"""
    try:
        lo, hi = resolve_range(month, date_from, date_to, start, end)
    except TimeRangeError:
        raise
    return {
        "month": month or None,
        "from": date_from or None,
        "to": date_to or None,
        "start": lo,
        "end": hi,
        "label": describe(lo, hi),
    }


def describe(start, end):
    """区间中文描述（前端显示当前筛选条件）"""
    if start is None and end is None:
        return "全部时间"
    if start is None:
        return f"< {end}"
    if end is None:
        return f">= {start}"
    if start[:8] == end[:8]:
        # 同一年：2026-10-01 00:00:00 ~ 2026-11-01 00:00:00
        return f"{start[:10]} ~ {end[:10]}"
    return f"{start[:10]} ~ {end[:10]}"

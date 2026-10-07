"""清洗：文本型数字转数值、占位值转空值、多格式日期解析。

每个函数返回 (清洗后的值, 错误说明)。错误说明不为 None 时由调用方登记到问题清单，
清洗后的值置为空，绝不静默丢弃。
"""
from __future__ import annotations

import datetime as dt
import math
import re
import unicodedata
from typing import Any

TYPE_TEXT = "text"
TYPE_NUMBER = "number"
TYPE_DATE = "date"
TYPE_CODE = "code"   # 编码：全角转半角、去空格、字母大写，数字不丢前导零以外的信息
TYPE_AUTO = "auto"   # 未指定类型：只做占位值和空白处理，保留原始类型

# 统一视为空值的占位写法（比较前会做全角转半角、去空格、转小写）
PLACEHOLDERS = {"", "无", "/", "-", "—", "——", "--", "---", "n/a", "na", "null", "none",
                "空", "暂无"}
# Excel 错误值：通常是公式出错，必须报出来，不能当作空值
EXCEL_ERRORS = {"#N/A", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!",
                "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#错误"}
ERR_EXCEL = "单元格是 Excel 错误值（公式出错）"

_EXCEL_EPOCH = dt.datetime(1899, 12, 30)
_MAX_EXCEL_SERIAL = 2958465          # 9999-12-31
_CURRENCY = re.compile(r"(人民币|rmb|cny|元|¥|￥|\$)", re.IGNORECASE)
_DATE_YMD = re.compile(
    r"^(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?"
    r"(?:[ tT]+(\d{1,2}):(\d{1,2})(?::(\d{1,2})(?:\.\d+)?)?)?$")
_DATE_COMPACT = re.compile(r"^(\d{4})(\d{2})(\d{2})$")
_SERIAL = re.compile(r"^\d+(\.\d+)?$")
# 没写年份的日期：9.30 / 9/30 / 9-30 / 9月30日 / 9月30号
_DATE_MD = re.compile(r"^(\d{1,2})\s*(?:[./-]|月)\s*(\d{1,2})\s*[日号]?$")
_MIN_DATE_SERIAL = 367               # 小于这个数（1901 年以前）不当作 Excel 日期序列号
# Excel 里为保住前导零常用的 ="000123" 写法
_EXCEL_TEXT_WRAPPER = re.compile(r'^\s*=\s*"(.*)"\s*$', re.S)
# 数字后面跟的计量单位（如 12箱、8 瓶、3.5kg）；万、亿等倍数单位不自动换算
_NUM_WITH_UNIT = re.compile(r"^([-+]?\d+(?:\.\d+)?)\s*([a-zA-Z\u4e00-\u9fff]{1,4})$")
_MULTIPLIER_UNITS = set("万亿千百")
ERR_NO_YEAR = "日期没有写年份"


def unwrap_excel_text(value: Any) -> Any:
    """去掉 ="000123" 外壳，得到 000123。"""
    if isinstance(value, str):
        m = _EXCEL_TEXT_WRAPPER.match(value)
        if m:
            return m.group(1)
    return value


def _norm_str(s: str) -> str:
    return unicodedata.normalize("NFKC", s).strip()


def is_placeholder(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if isinstance(value, str):
        return re.sub(r"\s+", "", _norm_str(value)).lower() in PLACEHOLDERS
    return False


# ---------------------------------------------------------------- 数值
def to_number(value: Any) -> tuple[float | int | None, str | None]:
    value = unwrap_excel_text(value)
    if is_placeholder(value):
        return None, None
    if isinstance(value, bool):
        return None, "布尔值不能作为数值"
    if isinstance(value, int):
        return value, None
    if isinstance(value, float):
        if math.isinf(value):
            return None, "数值为无穷大"
        return (int(value) if value.is_integer() and abs(value) < 1e15 else value), None
    if isinstance(value, (dt.date, dt.datetime, dt.time)):
        return None, "单元格是日期/时间，不是数值"
    if isinstance(value, str) and value.strip() in EXCEL_ERRORS:
        return None, ERR_EXCEL
    s = re.sub(r"\s+", "", _norm_str(str(value)))
    s = _CURRENCY.sub("", s).replace(",", "")
    negative = False
    if s.startswith("(") and s.endswith(")"):        # 会计格式负数 (1,234.00)
        negative, s = True, s[1:-1]
    percent = s.endswith("%")
    if percent:
        s = s[:-1]
    try:
        num = float(s)
    except ValueError:
        m = _NUM_WITH_UNIT.match(s)
        if not m or _MULTIPLIER_UNITS & set(m.group(2)):
            return None, "无法转换为数值"
        num = float(m.group(1))                  # 12箱 → 12：去掉计量单位
    if math.isnan(num) or math.isinf(num):
        return None, "无法转换为数值"
    if percent:
        num /= 100
    if negative:
        num = -num
    if num.is_integer() and abs(num) < 1e15 and not percent:
        return int(num), None
    return num, None


# ---------------------------------------------------------------- 日期
def _from_serial(num: float) -> dt.date | dt.datetime | None:
    if not 1 <= num <= _MAX_EXCEL_SERIAL:
        return None
    value = _EXCEL_EPOCH + dt.timedelta(days=num)
    return value.date() if float(num).is_integer() else value.replace(microsecond=0)


def _from_ymd(y: int, m: int, d: int, hh: int = 0, mm: int = 0, ss: int = 0):
    try:
        value = dt.datetime(y, m, d, hh, mm, ss)
    except ValueError:
        return None
    return value.date() if (hh, mm, ss) == (0, 0, 0) else value


def is_month_day(value: Any) -> bool:
    """是否为没写年份的日期写法（9.30、9月30日 等）。"""
    return isinstance(value, str) and bool(_DATE_MD.match(_norm_str(unwrap_excel_text(value))))


def to_date(value: Any, default_year: int | None = None
            ) -> tuple[dt.date | dt.datetime | None, str | None]:
    """兼容 2026-10-02、2026/10/2、20261002、2026年10月2日、Excel 序列号。

    没写年份的写法（9.30、9/30、9月30日）用 default_year 补全；没有 default_year 时报错。
    """
    value = unwrap_excel_text(value)
    if is_placeholder(value):
        return None, None
    if isinstance(value, dt.datetime):
        return (value.date() if value.time() == dt.time(0) else value.replace(microsecond=0)), None
    if isinstance(value, dt.date):
        return value, None
    if isinstance(value, bool):
        return None, "布尔值不能作为日期"
    if isinstance(value, (int, float)):
        if float(value).is_integer() and 19000101 <= value <= 29991231:
            m = _DATE_COMPACT.match(str(int(value)))
            parsed = _from_ymd(int(m[1]), int(m[2]), int(m[3])) if m else None
            if parsed is not None:
                return parsed, None
            return None, "8 位数字不是有效的 年月日"
        if value < _MIN_DATE_SERIAL:
            return None, "数字太小，不像日期（如果想写“月.日”，请写上年份）"
        parsed = _from_serial(float(value))
        return (parsed, None) if parsed is not None else (None, "数字超出 Excel 日期序列号范围")
    if isinstance(value, dt.timedelta):
        return None, "单元格是时长，不是日期"
    if isinstance(value, dt.time):
        return None, "单元格只有时间，没有日期（若原值是 1900 年以前的日期，Excel 无法正确保存）"

    s = _norm_str(str(value))
    if s in EXCEL_ERRORS:
        return None, ERR_EXCEL
    m = _DATE_YMD.match(s)
    if m:
        parsed = _from_ymd(*(int(g) if g else 0 for g in m.groups()))
        return (parsed, None) if parsed is not None else (None, "年月日数值无效")
    m = _DATE_COMPACT.match(s)
    if m:
        parsed = _from_ymd(int(m[1]), int(m[2]), int(m[3]))
        return (parsed, None) if parsed is not None else (None, "8 位数字不是有效的 年月日")
    m = _DATE_MD.match(s)
    if m:
        if default_year is None:
            return None, ERR_NO_YEAR
        parsed = _from_ymd(default_year, int(m[1]), int(m[2]))
        return (parsed, None) if parsed is not None else (None, "月日数值无效")
    if _SERIAL.match(s):
        if float(s) < _MIN_DATE_SERIAL:
            return None, "数字太小，不像日期"
        parsed = _from_serial(float(s))
        return (parsed, None) if parsed is not None else (None, "数字超出 Excel 日期序列号范围")
    return None, "无法识别的日期格式"


# ---------------------------------------------------------------- 文本
def to_text(value: Any) -> tuple[str | None, str | None]:
    """文本列：统一转为字符串，整数型浮点去掉 .0，占位值转空。"""
    value = unwrap_excel_text(value)
    if is_placeholder(value):
        return None, None
    if isinstance(value, float):
        return (str(int(value)) if value.is_integer() else repr(value)), None
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S" if value.time() != dt.time(0) else "%Y-%m-%d"), None
    if isinstance(value, dt.date):
        return value.isoformat(), None
    return str(value).strip(), None


def to_code(value: Any) -> tuple[str | None, str | None]:
    """编码列：BX－003、 bx-003 、="BX-003" 都统一成 BX-003；数字 123.0 → "123"。"""
    value = unwrap_excel_text(value)
    if is_placeholder(value):
        return None, None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if isinstance(value, (dt.date, dt.datetime)):
        return None, "单元格是日期，不是编码"
    s = re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(value))).upper()
    return (s or None), None


def to_auto(value: Any) -> tuple[Any, str | None]:
    """未指定类型的列：只处理占位值和首尾空白。"""
    value = unwrap_excel_text(value)
    if is_placeholder(value):
        return None, None
    if isinstance(value, str):
        return value.strip(), None
    if isinstance(value, float) and value.is_integer() and abs(value) < 1e15:
        return int(value), None
    return value, None


CONVERTERS = {
    TYPE_NUMBER: to_number,
    TYPE_DATE: to_date,
    TYPE_TEXT: to_text,
    TYPE_CODE: to_code,
    TYPE_AUTO: to_auto,
}


def clean_value(value: Any, col_type: str, default_year: int | None = None) -> tuple[Any, str | None]:
    if col_type == TYPE_DATE:
        return to_date(value, default_year)
    return CONVERTERS.get(col_type, to_auto)(value)

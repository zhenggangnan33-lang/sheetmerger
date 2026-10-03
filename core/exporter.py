"""导出 xlsx：汇总 / 明细 / 问题清单 三个 Sheet。

使用 openpyxl 的 write_only 模式，数据量大时内存和速度都可控。
格式：表头加粗带底色、冻结首行、自动列宽、数值列千分位、日期列统一格式。
"""
from __future__ import annotations

import datetime as dt
import unicodedata
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .validator import ERROR, ISSUE_COLUMNS, WARNING, Issue

SHEET_SUMMARY = "汇总"
SHEET_DETAIL = "明细"
SHEET_ISSUES = "问题清单"

EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_CELL_CHARS = 32_767
WIDTH_SAMPLE_ROWS = 3000
MIN_WIDTH, MAX_WIDTH = 6, 60

FMT_INT = "#,##0"
FMT_FLOAT = "#,##0.00"
FMT_DATE = "yyyy-mm-dd"
FMT_DATETIME = "yyyy-mm-dd hh:mm:ss"

_HEADER_FONT = Font(bold=True)
_HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
_THIN = Side(style="thin", color="A6A6A6")
_HEADER_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=False)
_SEVERITY_FONT = {ERROR: Font(color="C00000", bold=True), WARNING: Font(color="C65911")}


def default_filename(now: dt.datetime | None = None) -> str:
    now = now or dt.datetime.now()
    return f"汇总结果_{now:%Y%m%d_%H%M%S}.xlsx"


def _display_width(v: Any) -> int:
    if v is None:
        return 0
    if isinstance(v, dt.datetime):
        return 19
    if isinstance(v, dt.date):
        return 10
    if isinstance(v, float):
        s = f"{v:,.2f}"
    elif isinstance(v, int) and not isinstance(v, bool):
        s = f"{v:,}"
    else:
        s = str(v)
    s = s.split("\n", 1)[0][:200]
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in s)


def _column_format(values: Iterable[Any]) -> str | None:
    """根据列中的值决定数字格式：整数/小数千分位，日期/日期时间统一格式。"""
    has_num = has_float = has_date = has_dt = has_other = False
    for v in values:
        if v is None or (isinstance(v, float) and v != v):
            continue
        if isinstance(v, bool):
            has_other = True
        elif isinstance(v, int):
            has_num = True
        elif isinstance(v, float):
            has_num = True
            has_float = has_float or not v.is_integer()
        elif isinstance(v, dt.datetime):
            has_dt = True
        elif isinstance(v, dt.date):
            has_date = True
        else:
            has_other = True
    if has_num and not (has_date or has_dt):
        return FMT_FLOAT if has_float else FMT_INT
    if (has_date or has_dt) and not has_num and not has_other:
        return FMT_DATETIME if has_dt else FMT_DATE
    return None


def _safe_value(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, float) and v != v:
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, "item") and not isinstance(v, (str, bytes)):   # numpy 标量
        try:
            return v.item()
        except (ValueError, AttributeError):
            pass
    if isinstance(v, str):
        v = ILLEGAL_CHARACTERS_RE.sub("", v)
        if len(v) > EXCEL_MAX_CELL_CHARS:
            v = v[:EXCEL_MAX_CELL_CHARS]
    return v


class _SheetWriter:
    """把一个 DataFrame/行列表写入 write_only 工作簿，超出 Excel 行数上限时自动续页。"""

    def __init__(self, wb: Workbook, title: str, columns: list[str],
                 rows: list[list[Any]], severity_col: int | None = None,
                 plain_columns: Iterable[str] = ()) -> None:
        self.wb, self.title, self.columns, self.rows = wb, title, columns, rows
        self.severity_col = severity_col
        self.plain_columns = set(plain_columns)   # 行号等不加千分位

    def write(self) -> None:
        formats = [None if name in self.plain_columns else _column_format(r[i] for r in self.rows)
                   for i, name in enumerate(self.columns)]
        widths = []
        for i, name in enumerate(self.columns):
            w = max([_display_width(name)] +
                    [_display_width(r[i]) for r in self.rows[:WIDTH_SAMPLE_ROWS]])
            widths.append(min(max(w + 2, MIN_WIDTH), MAX_WIDTH))

        per_sheet = EXCEL_MAX_ROWS - 1
        chunks = [self.rows[i:i + per_sheet] for i in range(0, len(self.rows), per_sheet)] or [[]]
        for n, chunk in enumerate(chunks, start=1):
            ws = self.wb.create_sheet(self.title if n == 1 else f"{self.title}_{n}")
            self._write_sheet(ws, chunk, formats, widths)

    def _write_sheet(self, ws, rows, formats, widths) -> None:
        from openpyxl.utils import get_column_letter
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        header = []
        for name in self.columns:
            c = WriteOnlyCell(ws, value=_safe_value(name))
            c.font, c.fill, c.border, c.alignment = (_HEADER_FONT, _HEADER_FILL,
                                                     _HEADER_BORDER, _HEADER_ALIGN)
            header.append(c)
        ws.append(header)
        for row in rows:
            out = []
            for i, v in enumerate(row):
                v = _safe_value(v)
                fmt = formats[i]
                if isinstance(v, str):
                    if v.startswith("="):
                        # 以 = 开头的文本不能被当成公式
                        c = WriteOnlyCell(ws, value=v)
                        c.data_type = "s"
                        out.append(c)
                    elif self.severity_col == i and v in _SEVERITY_FONT:
                        c = WriteOnlyCell(ws, value=v)
                        c.font = _SEVERITY_FONT[v]
                        out.append(c)
                    else:
                        out.append(v)
                elif fmt and v is not None and not isinstance(v, bool):
                    c = WriteOnlyCell(ws, value=v)
                    c.number_format = fmt
                    out.append(c)
                else:
                    out.append(v)
            ws.append(out)


def _frame_rows(df: pd.DataFrame) -> list[list[Any]]:
    return [list(r) for r in df.itertuples(index=False, name=None)]


def export_result(path: str | Path, summary: pd.DataFrame, detail: pd.DataFrame,
                  issues: list[Issue], plain_columns: Iterable[str] = ()) -> Path:
    """写出结果文件。issues 需已按严重程度排序。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook(write_only=True)
    _SheetWriter(wb, SHEET_SUMMARY, list(summary.columns), _frame_rows(summary)).write()
    _SheetWriter(wb, SHEET_DETAIL, list(detail.columns), _frame_rows(detail),
                 plain_columns=plain_columns).write()
    _SheetWriter(wb, SHEET_ISSUES, ISSUE_COLUMNS, [i.to_row() for i in issues],
                 severity_col=0, plain_columns={"行号"}).write()
    tmp = path.with_name(path.stem + ".~tmp.xlsx")
    wb.save(tmp)
    tmp.replace(path)
    return path

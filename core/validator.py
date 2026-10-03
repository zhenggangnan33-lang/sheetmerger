"""问题收集：统一记录 文件 / Sheet / 行号 / 列名 / 问题类型 / 说明。

所有模块遇到异常数据都只往这里登记问题，不抛异常中断流程。
"""
from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Any, Iterable

# 严重程度（排序：错误 > 警告 > 提示）
ERROR = "错误"
WARNING = "警告"
INFO = "提示"
SEVERITY_ORDER = {ERROR: 0, WARNING: 1, INFO: 2}

# 问题类型
T_READ_FAIL = "文件读取失败"
T_EMPTY = "空文件/空Sheet"
T_NO_DATA = "无数据行"
T_HEADER = "表头识别"
T_MAP_PENDING = "表头待确认"
T_MAP_UNMATCHED = "表头未匹配"
T_MAP_CONFLICT = "表头映射冲突"
T_NO_STD_COLUMN = "未识别到标准列"
T_MISSING_COLUMN = "缺少列"
T_SKIP_TOTAL = "跳过合计行"
T_SKIP_NOTE = "跳过说明行"
T_BAD_NUMBER = "数值无法转换"
T_BAD_DATE = "日期无法识别"
T_DUPLICATE = "重复记录"
T_DUPLICATE_DROPPED = "重复记录已删除"
T_AGG = "汇总设置"

ISSUE_COLUMNS = ["严重程度", "文件", "Sheet", "行号", "列名", "问题类型", "原始值", "说明"]


@dataclass
class Issue:
    severity: str
    file: str = ""
    sheet: str = ""
    row: int | None = None
    column: str = ""
    type: str = ""
    message: str = ""
    value: Any = None

    def to_row(self) -> list:
        value = "" if self.value is None else str(self.value)
        return [self.severity, self.file, self.sheet, self.row, self.column,
                self.type, value, self.message]

    def to_dict(self) -> dict:
        return asdict(self)


class IssueCollector:
    """问题清单收集器。"""

    def __init__(self) -> None:
        self._issues: list[Issue] = []

    def add(self, severity: str, type: str, message: str, file: str = "", sheet: str = "",
            row: int | None = None, column: str = "", value: Any = None) -> Issue:
        issue = Issue(severity=severity, file=file, sheet=sheet, row=row, column=column,
                      type=type, message=message, value=value)
        self._issues.append(issue)
        return issue

    def extend(self, issues: Iterable[Issue]) -> None:
        self._issues.extend(issues)

    @property
    def issues(self) -> list[Issue]:
        return list(self._issues)

    def sorted(self) -> list[Issue]:
        """按 严重程度 > 文件 > Sheet > 行号 排序。"""
        return sorted(
            self._issues,
            key=lambda i: (SEVERITY_ORDER.get(i.severity, 9), i.file, i.sheet,
                           i.row if i.row is not None else -1, i.column),
        )

    def count(self, severity: str | None = None) -> int:
        if severity is None:
            return len(self._issues)
        return sum(1 for i in self._issues if i.severity == severity)

    def __len__(self) -> int:
        return len(self._issues)

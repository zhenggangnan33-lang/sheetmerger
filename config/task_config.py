"""任务配置的保存与加载（JSON）。保存后下次可一键重跑。"""
from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass, field
from pathlib import Path

CONFIG_VERSION = 1


@dataclass
class AggSpec:
    column: str
    func: str = "求和"       # 求和 / 计数 / 平均 / 最大 / 最小


@dataclass
class TaskConfig:
    input_folder: str = ""
    recursive: bool = False
    # 排除的 Sheet，格式 "相对路径|Sheet名"；"相对路径|*" 表示排除整个文件
    excluded: list[str] = field(default_factory=list)
    # 手动指定表头行（从 1 开始），键为 "相对路径|Sheet名"
    header_rows: dict[str, int] = field(default_factory=dict)
    # 手动表头映射：原表头 -> 目标列名；目标为空字符串表示忽略该列
    column_mapping: dict[str, str] = field(default_factory=dict)
    # 列类型覆盖：列名 -> text / number / date
    column_types: dict[str, str] = field(default_factory=dict)
    accept_pending: bool = True      # "待确认"的映射是否采用（采用时在问题清单中提示）
    keep_unmatched: bool = True      # "未匹配"的列是否按原表头保留到明细
    group_by: list[str] = field(default_factory=list)
    aggregations: list[AggSpec] = field(default_factory=list)
    add_count_column: bool = False   # 汇总表是否附加“记录数”列（没有汇总列时总会附加）
    pivot_column: str = ""           # 交叉表：把这一列的每个值展开成汇总表的一列（空 = 不展开）
    split_by: str = ""               # 按这一列拆分输出（空 = 不拆分）
    split_mode: str = "sheet"        # sheet 每个值一个 Sheet / file 每个值一个文件
    dedup_mode: str = "mark"         # off 不检查 / mark 只标记 / drop 删除重复
    dedup_columns: list[str] = field(default_factory=list)   # 为空表示按全部列
    output_dir: str = ""             # 为空时与输入文件夹相同
    output_name: str = ""            # 为空时使用 汇总结果_YYYYMMDD_HHMMSS.xlsx
    version: int = CONFIG_VERSION

    # ------------------------------------------------------------ 序列化
    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "TaskConfig":
        known = {f for f in cls.__dataclass_fields__}
        kwargs = {k: v for k, v in (data or {}).items() if k in known}
        aggs = []
        for a in kwargs.get("aggregations", []) or []:
            if isinstance(a, dict) and a.get("column"):
                aggs.append(AggSpec(column=a["column"], func=a.get("func") or "求和"))
            elif isinstance(a, (list, tuple)) and a:
                aggs.append(AggSpec(column=a[0], func=a[1] if len(a) > 1 else "求和"))
        kwargs["aggregations"] = aggs
        kwargs["header_rows"] = {k: int(v) for k, v in (kwargs.get("header_rows") or {}).items()}
        return cls(**kwargs)

    def save(self, path: str | Path) -> Path:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(path.suffix + ".tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
        return path

    @classmethod
    def load(cls, path: str | Path) -> "TaskConfig":
        with open(path, encoding="utf-8-sig") as f:
            return cls.from_dict(json.load(f))

    # ------------------------------------------------------------ 便捷方法
    def is_excluded(self, rel: str, sheet: str) -> bool:
        return f"{rel}|{sheet}" in self.excluded or f"{rel}|*" in self.excluded

    def resolve_output_path(self) -> Path:
        from core.exporter import default_filename
        folder = Path(self.output_dir or self.input_folder or ".")
        name = self.output_name or default_filename()
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        return folder / name

"""表头映射：标准化 + 别名字典 + 模糊匹配，输出映射建议和置信度。

- 相似度 ≥ 90 自动映射，60–90 标为"待确认"，< 60 标为"未匹配"。
- 用户确认后的映射可通过 AliasStore.learn() 写回用户别名字典，越用越准。
"""
from __future__ import annotations

import json
import os
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from rapidfuzz import fuzz

AUTO_THRESHOLD = 90
PENDING_THRESHOLD = 60

STATUS_AUTO = "自动"
STATUS_PENDING = "待确认"
STATUS_UNMATCHED = "未匹配"
STATUS_MANUAL = "手动"
STATUS_IGNORED = "忽略"

TYPE_TEXT = "text"
TYPE_NUMBER = "number"
TYPE_DATE = "date"
TYPE_CODE = "code"
COLUMN_TYPES = (TYPE_TEXT, TYPE_NUMBER, TYPE_DATE, TYPE_CODE)

DEFAULT_ALIAS_PATH = Path(__file__).with_name("aliases_default.json")

_BRACKETS = re.compile(r"\([^)]*\)|\[[^\]]*\]|【[^】]*】|〔[^〕]*〕")


def app_data_dir() -> Path:
    """用户数据目录（别名字典等）。可用环境变量 SHEETMERGER_HOME 覆盖。"""
    env = os.environ.get("SHEETMERGER_HOME")
    if env:
        return Path(env)
    if sys.platform == "win32" and os.environ.get("APPDATA"):
        return Path(os.environ["APPDATA"]) / "SheetMerger"
    return Path.home() / ".sheetmerger"


def default_user_alias_path() -> Path:
    return app_data_dir() / "aliases.json"


def normalize_header(text: object) -> str:
    """标准化表头：全角转半角、去所有空白和换行、统一小写、去末尾冒号。"""
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text))
    s = re.sub(r"\s+", "", s).lower()
    return s.rstrip(":")


def strip_units(norm: str) -> str:
    """去掉括号内的单位说明，例如 金额(元) -> 金额。"""
    return _BRACKETS.sub("", norm)


def similarity(a: str, b: str) -> float:
    """两个已标准化表头的相似度（0–100）。"""
    if not a or not b:
        return 0.0
    if a == b:
        return 100.0
    a2, b2 = strip_units(a), strip_units(b)
    if a2 and a2 == b2:
        return 95.0
    score = max(fuzz.ratio(a, b), fuzz.ratio(a2, b2))
    short, long_ = sorted((len(a2), len(b2)))
    if short >= 2 and long_ <= short * 2.5:
        # 包含关系（如 门店 / 门店名称）给到"待确认"区间，但不自动映射；
        # 长度相差太大（如一整句说明文字里含"门店"）不算
        score = max(score, 0.85 * fuzz.partial_ratio(a2, b2))
    return float(score)


# ---------------------------------------------------------------- 别名字典
@dataclass
class ColumnSpec:
    name: str
    type: str = TYPE_TEXT
    aliases: list[str] = field(default_factory=list)


class AliasStore:
    """标准列 + 别名字典。内置字典只读，用户学习到的别名写入用户字典。"""

    def __init__(self, user_path: str | Path | None = None,
                 default_path: str | Path = DEFAULT_ALIAS_PATH) -> None:
        self.default_path = Path(default_path)
        self.user_path = Path(user_path) if user_path else default_user_alias_path()
        self.specs: dict[str, ColumnSpec] = {}
        self._user: dict[str, dict] = {}
        self._merge(self._load_json(self.default_path))
        self._user = self._load_json(self.user_path)
        self._merge(self._user)
        self._rebuild_index()

    @staticmethod
    def _load_json(path: Path) -> dict:
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def _merge(self, data: dict) -> None:
        for name, info in data.items():
            info = info if isinstance(info, dict) else {"aliases": list(info or [])}
            spec = self.specs.get(name) or ColumnSpec(name=name)
            if info.get("type") in COLUMN_TYPES:
                spec.type = info["type"]
            for a in [name, *info.get("aliases", [])]:
                if a not in spec.aliases:
                    spec.aliases.append(a)
            self.specs[name] = spec

    def _rebuild_index(self) -> None:
        # 标准化别名 -> 标准列名
        self.index: dict[str, str] = {}
        for spec in self.specs.values():
            for a in spec.aliases:
                self.index.setdefault(normalize_header(a), spec.name)

    @property
    def standard_names(self) -> list[str]:
        return list(self.specs)

    def column_type(self, name: str) -> str | None:
        spec = self.specs.get(name)
        return spec.type if spec else None

    def learn(self, source: str, target: str, col_type: str | None = None) -> None:
        """把用户确认的 原表头 -> 标准列 写入用户字典并保存。"""
        if not source or not target:
            return
        entry = self._user.setdefault(target, {"aliases": []})
        if col_type in COLUMN_TYPES:
            entry["type"] = col_type
        elif target not in self.specs and "type" not in entry:
            entry["type"] = TYPE_TEXT
        if source not in entry["aliases"] and normalize_header(source) != normalize_header(target):
            entry["aliases"].append(source)
        self._merge({target: entry})
        self._rebuild_index()
        self.save()

    def save(self) -> None:
        self.user_path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.user_path.with_suffix(".tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self._user, f, ensure_ascii=False, indent=2)
        os.replace(tmp, self.user_path)


# ---------------------------------------------------------------- 映射建议
@dataclass
class MappingSuggestion:
    source: str                 # 原表头
    target: str | None          # 映射到的列名（未匹配时为原表头本身或 None）
    score: float
    status: str
    candidates: list[tuple[str, float]] = field(default_factory=list)

    @property
    def normalized(self) -> str:
        return normalize_header(self.source)


def suggest(source: str, store: AliasStore, top_n: int = 3) -> MappingSuggestion:
    """为单个表头给出映射建议。"""
    norm = normalize_header(source)
    if norm in store.index:
        target = store.index[norm]
        return MappingSuggestion(source, target, 100.0, STATUS_AUTO, [(target, 100.0)])

    best: dict[str, float] = {}
    for alias_norm, std in store.index.items():
        sc = similarity(norm, alias_norm)
        if sc > best.get(std, -1):
            best[std] = sc
    ranked = sorted(best.items(), key=lambda kv: (-kv[1], store.standard_names.index(kv[0])))
    candidates = [(n, round(s, 1)) for n, s in ranked[:top_n] if s > 0]
    if not candidates or candidates[0][1] < PENDING_THRESHOLD:
        return MappingSuggestion(source, None, candidates[0][1] if candidates else 0.0,
                                 STATUS_UNMATCHED, candidates)
    target, score = candidates[0]
    status = STATUS_AUTO if score >= AUTO_THRESHOLD else STATUS_PENDING
    return MappingSuggestion(source, target, score, status, candidates)


def suggest_for_columns(columns: list[str], store: AliasStore,
                        overrides: dict[str, str | None] | None = None
                        ) -> tuple[list[MappingSuggestion], list[str]]:
    """为一张表的所有列给出映射建议，并处理"多列映射到同一标准列"的冲突。

    overrides: {原表头(任意写法): 目标列名}，目标为 "" 或 None 表示忽略该列。
    返回 (建议列表, 冲突说明列表)。
    """
    ov = {normalize_header(k): v for k, v in (overrides or {}).items()}
    result: list[MappingSuggestion] = []
    for col in columns:
        norm = normalize_header(col)
        if norm in ov:
            tgt = ov[norm] or None
            result.append(MappingSuggestion(col, tgt, 100.0,
                                            STATUS_MANUAL if tgt else STATUS_IGNORED,
                                            [(tgt, 100.0)] if tgt else []))
        else:
            result.append(suggest(col, store))

    conflicts: list[str] = []
    by_target: dict[str, list[MappingSuggestion]] = {}
    for s in result:
        if s.target and s.status in (STATUS_AUTO, STATUS_PENDING, STATUS_MANUAL):
            by_target.setdefault(s.target, []).append(s)
    for target, group in by_target.items():
        if len(group) < 2:
            continue
        # 手动 > 分数高 > 靠前
        winner = max(group, key=lambda s: (s.status == STATUS_MANUAL, s.score,
                                           -result.index(s)))
        for s in group:
            if s is winner:
                continue
            conflicts.append(f"“{s.source}”与“{winner.source}”都匹配到“{target}”，"
                             f"保留“{winner.source}”，“{s.source}”按未匹配处理")
            s.status, s.target = STATUS_UNMATCHED, None
    return result, conflicts

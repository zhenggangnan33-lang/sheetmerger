"""读取：扫描文件夹、读取文件、识别表头行、处理合并单元格、多 Sheet。

约定：
- 内部行号一律使用 Excel 中看到的行号（从 1 开始），方便用户定位。
- 任何单个文件 / Sheet 的失败都只登记问题，不向外抛异常。
"""
from __future__ import annotations

import csv
import io
import os
import re
import stat
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable

from .validator import (ERROR, INFO, WARNING, IssueCollector, T_EMPTY, T_HEADER,
                        T_NO_DATA, T_READ_FAIL, T_SKIP_NOTE, T_SKIP_TOTAL)

EXCEL_EXTS = {".xlsx", ".xlsm", ".xls", ".xlsb"}
CSV_EXTS = {".csv"}
SUPPORTED_EXTS = EXCEL_EXTS | CSV_EXTS

CSV_SHEET_NAME = "CSV"          # CSV 文件没有 Sheet，统一用这个名字
HEADER_SCAN_ROWS = 20           # 在前 20 行中找表头
TOTAL_KEYWORDS = ("合计", "总计", "小计")

ProgressCallback = Callable[[str], None]


# ---------------------------------------------------------------- 数据结构
@dataclass
class SheetRef:
    """一个可读取的 Sheet（CSV 视为只有一个 Sheet）。"""
    path: Path
    rel: str            # 相对输入文件夹的路径（/ 分隔），用于显示和配置
    sheet: str

    @property
    def key(self) -> str:
        return sheet_key(self.rel, self.sheet)


@dataclass
class SheetTable:
    """识别完表头后的一张表。rows 中每项为 (Excel 行号, 与 columns 对齐的值列表)。"""
    rel: str
    sheet: str
    header_row: int                      # 表头所在行号（从 1 开始）
    columns: list[str]
    rows: list[tuple[int, list[Any]]] = field(default_factory=list)
    auto_header_row: int | None = None   # 自动识别出的表头行（便于界面展示）

    @property
    def key(self) -> str:
        return sheet_key(self.rel, self.sheet)


def sheet_key(rel: str, sheet: str) -> str:
    return f"{rel}|{sheet}"


# ---------------------------------------------------------------- 扫描
def _is_hidden(path: Path) -> bool:
    if path.name.startswith("."):
        return True
    try:
        attrs = getattr(os.stat(path), "st_file_attributes", 0)
        return bool(attrs & getattr(stat, "FILE_ATTRIBUTE_HIDDEN", 0))
    except OSError:
        return False


def scan_folder(folder: str | Path, recursive: bool = False) -> list[Path]:
    """列出文件夹中支持的表格文件，跳过 ~$ 临时文件和隐藏文件。"""
    root = Path(folder)
    if not root.is_dir():
        return []
    candidates: Iterable[Path] = root.rglob("*") if recursive else root.iterdir()
    result = []
    for p in candidates:
        if not p.is_file() or p.suffix.lower() not in SUPPORTED_EXTS:
            continue
        if p.name.startswith("~$") or _is_hidden(p):
            continue
        # 递归时，隐藏目录下的文件也跳过
        if recursive and any(part.startswith(".") for part in p.relative_to(root).parts[:-1]):
            continue
        result.append(p)
    return sorted(result, key=lambda p: p.relative_to(root).as_posix())


def rel_path(path: Path, root: str | Path) -> str:
    try:
        return path.relative_to(root).as_posix()
    except ValueError:
        return path.name


# ---------------------------------------------------------------- 读取原始网格
def list_sheets(path: Path) -> list[str]:
    """返回文件中的工作表名；读取失败时抛出异常（由调用方登记问题）。"""
    if path.suffix.lower() in CSV_EXTS:
        return [CSV_SHEET_NAME]
    from python_calamine import CalamineWorkbook, SheetTypeEnum
    wb = CalamineWorkbook.from_path(str(path))
    try:
        names = []
        for meta in wb.sheets_metadata:
            # 图表 Sheet 等非工作表跳过
            if meta.typ in (SheetTypeEnum.WorkSheet,) or str(meta.typ).endswith("WorkSheet"):
                names.append(meta.name)
        return names
    finally:
        wb.close()


_CJK = re.compile(r"[\u4e00-\u9fff]")
# 东亚多字节编码：短文本时 charset-normalizer 容易在它们之间误判，面向中文用户统一优先按 GBK
_CHINESE_ENCODINGS = {"gb2312", "gbk", "gb18030", "hz", "cp936", "euc_cn", "big5", "big5hkscs",
                      "cp950", "cp949", "euc_kr", "johab", "shift_jis", "cp932", "euc_jp",
                      "iso2022_jp", "iso2022_kr"}


def decode_csv_bytes(data: bytes) -> tuple[str, str]:
    """自动识别 CSV 编码，返回 (文本, 编码名)。重点兼容 UTF-8(-BOM) 与 GBK。"""
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", errors="replace"), "utf-8-sig"
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16"), "utf-16"
    try:
        return data.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        pass
    # 只含常用汉字的 GBK 文件能按 GB2312 严格解码，这是最常见的情况，直接认定
    try:
        text = data.decode("gb2312")
        if _CJK.search(text):
            return data.decode("gb18030"), "gbk"
    except UnicodeDecodeError:
        pass
    # 其余交给 charset-normalizer；识别为中文编码时统一用 GB18030（GBK 超集）解码
    from charset_normalizer import from_bytes
    best = from_bytes(data).best()
    if best is not None and best.encoding:
        enc = best.encoding.lower().replace("-", "_")
        if enc in _CHINESE_ENCODINGS:
            try:
                text = data.decode("gb18030")
                if _CJK.search(text):
                    return text, "gbk"
            except UnicodeDecodeError:
                pass
        if not enc.startswith("utf"):
            # 单字节编码之间短文本难以区分，Windows 上最常见的是 cp1252（西欧）
            try:
                return data.decode("cp1252"), "cp1252"
            except UnicodeDecodeError:
                pass
        return str(best), best.encoding
    try:
        return data.decode("gb18030"), "gbk"
    except UnicodeDecodeError:
        pass
    return data.decode("latin-1"), "latin-1"


def _read_csv_grid(path: Path) -> list[list[Any]]:
    data = path.read_bytes()
    if not data.strip():
        return []
    text, _enc = decode_csv_bytes(data)
    sample = text[:8192]
    try:
        delimiter = csv.Sniffer().sniff(sample, delimiters=",\t;|").delimiter
    except csv.Error:
        # 各行列数不一致时 Sniffer 会失败，改按表头行中出现最多的分隔符判断
        first = sample.split("\n", 1)[0]
        counts = {d: first.count(d) for d in ",\t;|"}
        delimiter = max(counts, key=counts.get) if any(counts.values()) else ","
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    return [[(c if c != "" else None) for c in row] for row in reader]


def _merged_ranges(path: Path, sheet_obj, sheet_name: str) -> list[tuple[int, int, int, int]]:
    """返回合并区域 (r1, c1, r2, c2)，0 起始、闭区间。

    优先用 calamine 读取（快，且支持 xls）；取不到时 xlsx 回退到 openpyxl。
    """
    ranges = None
    try:
        ranges = sheet_obj.merged_cell_ranges
    except Exception:  # noqa: BLE001 - 不同格式支持程度不同
        ranges = None
    if ranges is not None:
        return [(a[0], a[1], b[0], b[1]) for a, b in ranges]
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
            try:
                ws = wb[sheet_name]
                return [(m.min_row - 1, m.min_col - 1, m.max_row - 1, m.max_col - 1)
                        for m in ws.merged_cells.ranges]
            finally:
                wb.close()
        except Exception:  # noqa: BLE001
            return []
    return []


_ERROR_CELL = re.compile(rb'<c\b(?=[^>]*\bt="e")[^>]*\br="([A-Z]+)(\d+)"[^>]*>(.*?)</c>', re.S)
_CELL_VALUE = re.compile(rb"<v>([^<]*)</v>")


def _col_index(letters: bytes) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ch - 64)
    return n - 1


def _xlsx_error_cells(path: Path, sheet_name: str) -> dict[tuple[int, int], str]:
    """找出 xlsx 中的 Excel 错误值单元格（#DIV/0!、#N/A 等），返回 {(行, 列): 错误文本}。

    calamine 会把错误值读成空，这里补读，避免"静默丢弃"。文件中没有错误值时几乎没有开销。
    """
    import zipfile
    from xml.etree import ElementTree as ET
    try:
        with zipfile.ZipFile(path) as zf:
            ns_main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            ns_rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"
            r_id = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            wb = ET.fromstring(zf.read("xl/workbook.xml"))
            rid = next((s.get(r_id) for s in wb.iter(f"{ns_main}sheet")
                        if s.get("name") == sheet_name), None)
            rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            target = next((r.get("Target") for r in rels.iter(f"{ns_rel}Relationship")
                           if r.get("Id") == rid), None)
            if not target:
                return {}
            name = target.lstrip("/") if target.startswith("/") else "xl/" + target
            data = zf.read(name)
    except (KeyError, OSError, ValueError, zipfile.BadZipFile, ET.ParseError):
        return {}
    if b't="e"' not in data:
        return {}
    out = {}
    for m in _ERROR_CELL.finditer(data):
        v = _CELL_VALUE.search(m.group(3))
        out[(int(m.group(2)) - 1, _col_index(m.group(1)))] = (
            v.group(1).decode("utf-8", "replace") if v else "#错误")
    return out


def fill_merged(grid: list[list[Any]], ranges: list[tuple[int, int, int, int]]) -> None:
    """把合并区域左上角的值向下、向右填充到整个区域。"""
    for r1, c1, r2, c2 in ranges:
        if r1 >= len(grid) or c1 >= len(grid[r1]):
            continue
        value = grid[r1][c1]
        if value is None:
            continue
        for r in range(r1, min(r2, len(grid) - 1) + 1):
            row = grid[r]
            if len(row) <= c2:
                row.extend([None] * (c2 + 1 - len(row)))
            for c in range(c1, c2 + 1):
                row[c] = value


def read_grid(path: Path, sheet: str) -> list[list[Any]]:
    """读取一个 Sheet 的原始二维网格（第 0 行即 Excel 第 1 行），并填充合并单元格。"""
    if path.suffix.lower() in CSV_EXTS:
        return _read_csv_grid(path)
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(str(path))
    try:
        sh = wb.get_sheet_by_name(sheet)
        raw = sh.to_python(skip_empty_area=False)
        grid = [[(None if v == "" else v) for v in row] for row in raw]
        if path.suffix.lower() in (".xlsx", ".xlsm"):
            for (r, c), err in _xlsx_error_cells(path, sheet).items():
                if r < len(grid):
                    row = grid[r]
                    if len(row) <= c:
                        row.extend([None] * (c + 1 - len(row)))
                    row[c] = err
        fill_merged(grid, _merged_ranges(path, sh, sheet))
        return grid
    finally:
        wb.close()


# ---------------------------------------------------------------- 表头识别
def _is_text(v: Any) -> bool:
    if not isinstance(v, str):
        return False
    s = v.strip()
    if not s:
        return False
    try:
        float(s.replace(",", ""))
        return False
    except ValueError:
        return True


def detect_header_row(grid: list[list[Any]], max_rows: int = HEADER_SCAN_ROWS) -> int | None:
    """在前 max_rows 行中找"非空文本单元格最多"的一行，返回 0 起始的行索引。

    合并单元格向右填充后标题行会出现多个相同值，因此按"不同文本值"计数，
    避免把横跨整行的大标题误判为表头。并列时取靠前的一行。
    """
    best_idx, best_count = None, 0
    for idx, row in enumerate(grid[:max_rows]):
        count = len({v.strip() for v in row if _is_text(v)})
        if count > best_count:
            best_idx, best_count = idx, count
    return best_idx


def _col_letter(idx: int) -> str:
    s = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def _cell_text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _is_total_row(row: list[Any]) -> bool:
    first = _cell_text(row[0]) if row else ""
    if not first:
        first = next((_cell_text(v) for v in row if _cell_text(v)), "")
    return any(k in first for k in TOTAL_KEYWORDS)


def build_table(grid: list[list[Any]], header_idx: int, rel: str, sheet: str,
                issues: IssueCollector) -> SheetTable:
    """以 header_idx 行为表头，整理出列名和数据行，跳过空行、合计行、说明行。"""
    width = max((len(r) for r in grid), default=0)
    header = list(grid[header_idx]) + [None] * (width - len(grid[header_idx]))

    # 列名：空表头命名为"未命名列X"，重名追加 _2、_3
    names: list[str] = []
    seen: dict[str, int] = {}
    for i, v in enumerate(header):
        name = _cell_text(v).replace("\r", "").replace("\n", "") or f"未命名列{_col_letter(i)}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        names.append(name)

    data_rows: list[tuple[int, list[Any]]] = []
    for idx in range(header_idx + 1, len(grid)):
        row = list(grid[idx]) + [None] * (width - len(grid[idx]))
        excel_row = idx + 1
        non_empty = [v for v in row if _cell_text(v)]
        if not non_empty:
            continue
        if _is_total_row(row):
            issues.add(INFO, T_SKIP_TOTAL, "识别为合计行，已跳过（不计入明细）",
                       file=rel, sheet=sheet, row=excel_row, value=_cell_text(non_empty[0]))
            continue
        if len(non_empty) == 1 and width >= 4 and _is_text(non_empty[0]):
            issues.add(INFO, T_SKIP_NOTE, "整行只有一个文字单元格，视为说明/备注行，已跳过",
                       file=rel, sheet=sheet, row=excel_row, value=_cell_text(non_empty[0]))
            continue
        data_rows.append((excel_row, row))

    # 去掉表头为空且整列无数据的列
    keep = [i for i, v in enumerate(header)
            if _cell_text(v) or any(_cell_text(r[i]) for _, r in data_rows)]
    columns = [names[i] for i in keep]
    rows = [(n, [r[i] for i in keep]) for n, r in data_rows]
    return SheetTable(rel=rel, sheet=sheet, header_row=header_idx + 1,
                      columns=columns, rows=rows)


# ---------------------------------------------------------------- 对外入口
def discover(folder: str | Path, recursive: bool, issues: IssueCollector) -> list[SheetRef]:
    """扫描文件夹并列出所有 Sheet；打不开的文件登记为错误。"""
    refs: list[SheetRef] = []
    for path in scan_folder(folder, recursive):
        rel = rel_path(path, folder)
        try:
            if path.stat().st_size == 0:
                issues.add(WARNING, T_EMPTY, "文件大小为 0，已跳过", file=rel)
                continue
            sheets = list_sheets(path)
        except Exception as e:  # noqa: BLE001 - 任何错误都不能中断整体流程
            issues.add(ERROR, T_READ_FAIL, f"无法打开文件（可能已损坏、加密或格式不符）：{_err(e)}",
                       file=rel)
            continue
        if not sheets:
            issues.add(WARNING, T_EMPTY, "文件中没有工作表", file=rel)
        refs.extend(SheetRef(path=path, rel=rel, sheet=s) for s in sheets)
    return refs


def load_sheet(ref: SheetRef, issues: IssueCollector,
               header_row: int | None = None) -> SheetTable | None:
    """读取一个 Sheet 并识别表头。header_row 为用户指定的表头行号（从 1 开始）。"""
    try:
        grid = read_grid(ref.path, ref.sheet)
    except Exception as e:  # noqa: BLE001
        issues.add(ERROR, T_READ_FAIL, f"读取失败：{_err(e)}", file=ref.rel, sheet=ref.sheet)
        return None
    if not any(_cell_text(v) for row in grid for v in row):
        issues.add(WARNING, T_EMPTY, "Sheet 为空，已跳过", file=ref.rel, sheet=ref.sheet)
        return None

    auto_idx = detect_header_row(grid)
    if header_row is not None:
        if not 1 <= header_row <= len(grid):
            issues.add(WARNING, T_HEADER,
                       f"指定的表头行 {header_row} 超出范围，改用自动识别结果",
                       file=ref.rel, sheet=ref.sheet)
            header_idx = auto_idx
        else:
            header_idx = header_row - 1
    else:
        header_idx = auto_idx
    if header_idx is None:
        issues.add(WARNING, T_HEADER, f"前 {HEADER_SCAN_ROWS} 行中未找到文字表头，已跳过",
                   file=ref.rel, sheet=ref.sheet)
        return None

    try:
        table = build_table(grid, header_idx, ref.rel, ref.sheet, issues)
    except Exception as e:  # noqa: BLE001
        issues.add(ERROR, T_READ_FAIL, f"整理表格失败：{_err(e)}", file=ref.rel, sheet=ref.sheet)
        return None
    table.auto_header_row = auto_idx + 1 if auto_idx is not None else None
    if header_idx > 0:
        issues.add(INFO, T_HEADER, f"表头位于第 {header_idx + 1} 行，上方内容视为标题已跳过",
                   file=ref.rel, sheet=ref.sheet, row=header_idx + 1)
    if not table.rows:
        issues.add(WARNING, T_NO_DATA, "只有表头，没有数据行", file=ref.rel, sheet=ref.sheet,
                   row=table.header_row)
    return table


def _err(e: Exception) -> str:
    msg = str(e).strip().splitlines()[0] if str(e).strip() else ""
    return f"{type(e).__name__}: {msg}" if msg else type(e).__name__

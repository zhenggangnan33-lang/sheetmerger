"""命令行入口，用于测试和调试。

示例：
    python cli.py --input ./tests/test_data --group 门店 --sum 金额 --output result.xlsx
    python cli.py --config 我的任务.json
    python cli.py --input ./tests/test_data --show-mapping
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from config.task_config import AggSpec, TaskConfig  # noqa: E402
from core import aggregator, pipeline  # noqa: E402
from core.header_mapper import AliasStore  # noqa: E402
from core.validator import ERROR, INFO, WARNING, IssueCollector  # noqa: E402


def _split(values: list[str] | None) -> list[str]:
    out: list[str] = []
    for v in values or []:
        out.extend(x.strip() for x in v.replace("，", ",").split(",") if x.strip())
    return out


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="SheetMerger 多表汇总工具（命令行版）")
    p.add_argument("--input", "-i", help="输入文件夹")
    p.add_argument("--recursive", "-r", action="store_true", help="包含子文件夹")
    p.add_argument("--config", "-c", help="加载任务配置 JSON（命令行参数会覆盖其中的同名设置）")
    p.add_argument("--save-config", help="把本次的完整配置保存到该 JSON 文件")
    p.add_argument("--group", "-g", action="append", help="分组列，可多次指定或用逗号分隔")
    p.add_argument("--sum", "-s", action="append", help="求和列，可多次指定或用逗号分隔")
    p.add_argument("--agg", "-a", action="append",
                   help="汇总列:方式，方式为 求和/计数/平均/最大/最小，例如 金额:平均")
    p.add_argument("--count-column", action="store_true", help="汇总表附加“记录数”列")
    p.add_argument("--pivot", help="交叉表：把这一列的每个值展开成汇总表的一列，如 --pivot 门店")
    p.add_argument("--split", help="按这一列拆分输出，如 --split 门店")
    p.add_argument("--split-mode", choices=["sheet", "file"],
                   help="拆分方式：sheet 每个值一个 Sheet（默认）/ file 每个值一个文件")
    p.add_argument("--dedup", choices=aggregator.DEDUP_MODES,
                   help="重复记录处理：off 不检查 / mark 只标记(默认) / drop 删除")
    p.add_argument("--dedup-cols", action="append", help="按哪些列判断重复，默认全部列")
    p.add_argument("--map", "-m", action="append",
                   help="手动表头映射 原表头=目标列，目标留空表示忽略该列，例如 销售额=金额")
    p.add_argument("--type", action="append", help="列类型 列名=text/number/date")
    p.add_argument("--header-row", action="append",
                   help="指定表头行 文件|Sheet=行号，例如 a.xlsx|Sheet1=3")
    p.add_argument("--exclude", action="append", help="排除 文件|Sheet，或 文件|* 排除整个文件")
    p.add_argument("--strict-mapping", action="store_true",
                   help="不采用“待确认”的映射（默认采用并在问题清单中提示）")
    p.add_argument("--drop-unmatched", action="store_true", help="丢弃未匹配的列（默认保留）")
    p.add_argument("--output", "-o", help="输出文件路径，默认在输入文件夹生成 汇总结果_时间.xlsx")
    p.add_argument("--alias-file", help="用户别名字典路径（默认在用户数据目录）")
    p.add_argument("--learn", action="store_true", help="把 --map 指定的映射写回用户别名字典")
    p.add_argument("--show-mapping", action="store_true", help="只显示识别结果和表头映射建议，不导出")
    return p


def config_from_args(args: argparse.Namespace) -> TaskConfig:
    cfg = TaskConfig.load(args.config) if args.config else TaskConfig()
    if args.input:
        cfg.input_folder = args.input
    if args.recursive:
        cfg.recursive = True
    if args.group:
        cfg.group_by = _split(args.group)
    aggs = [AggSpec(c, "求和") for c in _split(args.sum)]
    for item in _split(args.agg):
        col, _, func = item.partition(":")
        if not func:
            col, _, func = item.partition("：")
        aggs.append(AggSpec(col.strip(), (func or "求和").strip()))
    if aggs:
        cfg.aggregations = aggs
    if args.count_column:
        cfg.add_count_column = True
    if args.pivot is not None:
        cfg.pivot_column = args.pivot.strip()
    if args.split is not None:
        cfg.split_by = args.split.strip()
    if args.split_mode:
        cfg.split_mode = args.split_mode
    if args.dedup:
        cfg.dedup_mode = args.dedup
    if args.dedup_cols:
        cfg.dedup_columns = _split(args.dedup_cols)
    for item in args.map or []:
        src, _, tgt = item.partition("=")
        cfg.column_mapping[src.strip()] = tgt.strip()
    for item in args.type or []:
        col, _, t = item.partition("=")
        cfg.column_types[col.strip()] = t.strip()
    for item in args.header_row or []:
        key, _, row = item.rpartition("=")
        cfg.header_rows[key.strip()] = int(row)
    for item in args.exclude or []:
        if item not in cfg.excluded:
            cfg.excluded.append(item.strip())
    if args.strict_mapping:
        cfg.accept_pending = False
    if args.drop_unmatched:
        cfg.keep_unmatched = False
    if args.output:
        out = Path(args.output)
        cfg.output_dir = str(out.parent) if str(out.parent) not in ("", ".") else "."
        cfg.output_name = out.name
    return cfg


def show_mapping(cfg: TaskConfig, store: AliasStore) -> None:
    issues = IssueCollector()
    plans, _ = pipeline.plan_tables(cfg, store, issues)
    for p in plans:
        t = p.table
        print(f"\n[{t.rel} / {t.sheet}] 表头第 {t.header_row} 行，数据 {len(t.rows)} 行")
        for s in p.suggestions:
            tgt = s.target or "-"
            print(f"  {s.source:<16} -> {tgt:<10} {s.status}  {s.score:.0f}")
    for i in issues.sorted():
        if i.severity != INFO:
            print(f"{i.severity} {i.file} {i.sheet} {i.type}: {i.message}")


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    args = build_parser().parse_args(argv)
    try:
        cfg = config_from_args(args)
    except (OSError, ValueError) as e:
        print(f"配置有误：{e}", file=sys.stderr)
        return 2
    if not cfg.input_folder:
        print("请用 --input 指定输入文件夹，或用 --config 加载配置", file=sys.stderr)
        return 2
    store = AliasStore(user_path=args.alias_file)
    if args.learn:
        for src, tgt in cfg.column_mapping.items():
            if tgt:
                store.learn(src, tgt, cfg.column_types.get(tgt))
    if args.save_config:
        cfg.save(args.save_config)
        print(f"配置已保存：{args.save_config}")
    if args.show_mapping:
        show_mapping(cfg, store)
        return 0

    tty = sys.stdout.isatty()

    def progress(pct: int, msg: str) -> None:
        if tty:
            print(f"\r[{pct:3d}%] {msg[:50]:<50}", end="", flush=True)

    result = pipeline.run(cfg, store, progress=progress)
    if tty:
        print()
    c = result.issue_counts
    print(f"处理文件 {result.files_ok}/{result.files_total} 个，Sheet {result.sheets_read} 个，"
          f"明细 {result.rows_detail} 行，用时 {result.elapsed:.1f} 秒")
    print(f"问题：错误 {c[ERROR]}，警告 {c[WARNING]}，提示 {c[INFO]}")
    if result.output_path:
        print(f"结果文件：{result.output_path}")
        if result.split_dir:
            print(f"拆分文件：{result.split_dir}（{result.split_count} 个）")
        elif result.split_count:
            print(f"已按“{cfg.split_by}”拆分为 {result.split_count} 个 Sheet")
        return 0
    for i in result.issues.sorted()[:5]:
        print(f"{i.severity} {i.message}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())

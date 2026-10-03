# SheetMerger 多表汇总工具

把一个文件夹里格式不完全一致的多个表格（xlsx / xls / xlsb / csv）合并、清洗、汇总，
导出带"问题清单"的结果表。全部在本机处理，不联网。

> 当前进度：**阶段 1（核心引擎 + 命令行）**。图形界面、打包说明在后续阶段补充。

## 安装（开发环境）

```bash
python -m pip install -r requirements.txt
```

## 命令行用法

```bash
# 按门店汇总金额
python cli.py --input ./tests/test_data --group 门店 --sum 金额 --output result.xlsx

# 多个分组列、多种汇总方式
python cli.py -i 数据文件夹 -g 门店,商品 -s 金额 -a 数量:平均 -a 日期:最大

# 只看识别出的表头和映射建议，不导出
python cli.py -i 数据文件夹 --show-mapping

# 手动指定映射（目标留空表示忽略该列），并写回别名字典
python cli.py -i 数据文件夹 --map "单价/元=单价" --map "会员卡号=" --learn --show-mapping

# 保存配置，下次一键重跑
python cli.py -i 数据文件夹 -g 门店 -s 金额 --save-config 我的任务.json
python cli.py --config 我的任务.json
```

常用参数：

| 参数 | 说明 |
|---|---|
| `-r/--recursive` | 包含子文件夹 |
| `--dedup off/mark/drop` | 重复记录：不检查 / 只在问题清单标记（默认）/ 删除 |
| `--dedup-cols 列1,列2` | 按指定列判断重复，默认按全部业务列 |
| `--header-row "文件|Sheet=行号"` | 手动指定表头行 |
| `--exclude "文件|Sheet"` | 排除某个 Sheet，`文件|*` 排除整个文件 |
| `--type 列名=number` | 指定列类型（text / number / date） |
| `--strict-mapping` | 不采用"待确认"的表头映射 |
| `--drop-unmatched` | 丢弃未匹配到标准列的列 |

## 结果文件

一个 xlsx，三个 Sheet：

- **汇总**：分组列 + 各汇总列 + 记录数，最后一行为总计
- **明细**：合并后的全部记录，末尾附"来源文件 / 来源Sheet / 原始行号"便于追溯
- **问题清单**：按 错误 > 警告 > 提示 排序，每条带 文件 / Sheet / 行号 / 列名 / 原始值

默认文件名 `汇总结果_YYYYMMDD_HHMMSS.xlsx`，默认保存在输入文件夹（再次扫描时会自动跳过这类文件）。

## 处理规则摘要

- 跳过 `~$` 开头的临时文件和隐藏文件；打不开的文件记为"错误"后继续处理其他文件
- 表头：前 20 行中"不同文字单元格最多"的一行；上方标题行跳过
- 跳过：空行、首列含"合计/总计/小计"的行、整行只有一个文字单元格的说明行（如"制表人：xx"）
- 合并单元格：左上角的值向下、向右填充
- 表头映射：别名完全匹配或去掉括号单位后匹配 → 自动；相似度 ≥90 自动、60–90 待确认、<60 未匹配
- 数值：去千分位、"元"、¥、空格，支持全角数字、会计负数 `(26.00)`、百分比
- 空值：无、/、-、—、N/A、空字符串等
- 日期：2026-10-02、2026/10/2、2026.10.2、20261002、2026年10月2日、Excel 序列号
- 无法转换的值在明细中留空，并在问题清单中记录原始值

## 目录结构

```
core/         读取、表头映射、清洗、校验、汇总、导出、流程编排（pipeline.py）
config/       任务配置保存/加载
cli.py        命令行入口
tests/        测试与测试数据生成（python tests/generate_test_data.py）
```

## 测试

```bash
python tests/generate_test_data.py   # 重新生成 tests/test_data
python -m pytest
```

用户学习到的别名保存在 `%APPDATA%\SheetMerger\aliases.json`（Windows）或 `~/.sheetmerger/aliases.json`，
可用环境变量 `SHEETMERGER_HOME` 改到其他位置。

# 第三方依赖及许可协议

本工具只使用 MIT / BSD / Apache / LGPL 许可的依赖。打包发布前请按实际锁定的版本复核一遍。

## 运行时依赖（会打进 exe）

| 依赖 | 用途 | 许可协议 |
|---|---|---|
| pandas | 明细合并、分组汇总的数据结构 | BSD-3-Clause |
| numpy（pandas 依赖） | 数值计算 | BSD-3-Clause |
| python-dateutil（pandas 依赖） | 日期工具 | Apache-2.0 / BSD-3-Clause 双许可 |
| six（python-dateutil 依赖） | 兼容层 | MIT |
| tzdata（pandas 依赖，Windows 上需要） | 时区数据 | Apache-2.0 |
| python-calamine | 读取 xlsx / xls / xlsb（内含 Rust 库 calamine） | MIT |
| openpyxl | 写出 xlsx、读取合并单元格（备用） | MIT |
| et-xmlfile（openpyxl 依赖） | XML 写出 | MIT |
| lxml | openpyxl 的加速 XML 后端 | BSD-3-Clause |
| rapidfuzz | 表头模糊匹配 | MIT |
| charset-normalizer | CSV 编码识别 | MIT |
| PySide6 / shiboken6（阶段 2） | 图形界面 | LGPL-3.0 |

### LGPL（PySide6）注意事项
- 以动态链接方式使用 PySide6，不修改其源码。
- 发布时需附带 LGPL-3.0 许可文本，并说明用户可以替换 PySide6 / Qt 库。
  单文件 exe 会把 Qt 动态库打进包里，运行时解压后依然是独立的动态库，满足"可替换"要求；
  如需更稳妥，可改用 PyInstaller 的目录模式（onedir）发布。

## 开发 / 测试 / 打包工具（不打进 exe）

| 依赖 | 用途 | 许可协议 |
|---|---|---|
| pytest | 单元测试 | MIT |
| xlwt | 生成 .xls 测试样本 | BSD |
| PyInstaller（阶段 4） | 打包 exe | GPL-2.0 附带 Bootloader 例外条款，允许打包后的程序以任意协议发布 |

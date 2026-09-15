# WPS 表格按列拆分工具（统信 UOS / Linux）

按**指定列的内容**把一个 WPS / Excel 表格拆成多个表格，新文件**以分类名命名**，并且**完整沿用原表的模板格式**。

> 一张 3000 行的工资总表，按「部门」列拆分 → 得到 `财务部.xlsx`、`销售部.xlsx`、`技术部.xlsx`……
> 每个表只有本部门的数据，字体、边框、底色、行高、列宽、数字格式、冻结窗格、自动筛选跟原表一模一样。

---

## ✨ 核心能力

| 特点 | 说明 |
|---|---|
| **★ 保留原模板格式** | 「模板克隆」方式：直接复制原文件的全部样式部件，只替换数据行。字体/字号/颜色/边框/底色/行高/列宽/数字格式（日期·金额·百分比）/对齐/冻结窗格/自动筛选/合并单元格/条件格式/数据验证/超链接**全部保留** |
| **公式跟随** | 公式保留，行号引用自动改写（原第 12 行 → 新第 5 行，`=D12*E12` → `=D5*E5`） |
| **零依赖** | 只用 Python 标准库，**不装 pandas / openpyxl**，内网离线环境开箱即用 |
| **中文友好** | 文件名直接用中文分类名；清单 CSV 用 UTF-8 BOM，WPS 打开不乱码 |
| **数据保真** | 日期、数字、布尔、百分比等单元格类型原样保留 |
| **文件名安全** | 自动清理 `\ / : * ? " < > \|` 等非法字符，超长截断，重名加 `_2` |
| **三种用法** | 图形界面 / 交互式向导 / 命令行（可写进脚本批量跑） |
| **预演模式** | 先看分类和行数，确认无误再落盘 |

---

## 🚀 快速开始

```bash
git clone https://github.com/leifung/wps-split.git
cd wps-split
chmod +x 运行.sh 安装桌面快捷方式.sh

# 最常用：按“部门”列拆分
python3 wps_split.py 工资表.xlsx -c 部门

# 图形界面
python3 wps_split.py --gui        # 或双击「运行.sh」

# 交互式向导（一路回车）
python3 wps_split.py
```

统信 UOS 自带 Python3，一般无需安装任何东西。图形界面需要 tkinter，缺失时会自动退回命令行向导：

```bash
sudo apt install python3-tk
```

---

## 📖 常用命令

```bash
# 表头不在第一行（第 1 行是大标题，第 2 行才是表头）
python3 wps_split.py 工资表.xlsx -c 部门 --header-row 2

# 用列字母（C 列）或列序号（第 3 列）
python3 wps_split.py 工资表.xlsx -c C
python3 wps_split.py 工资表.xlsx -c 3

# 指定输出目录
python3 wps_split.py 工资表.xlsx -c 部门 -o ~/桌面/拆分结果

# 按两列组合拆分 → “财务部-2026.xlsx”
python3 wps_split.py 工资表.xlsx -c 部门,年份 --sep "-"

# 指定工作表 / 拆分所有工作表
python3 wps_split.py 工资表.xlsx -c 部门 --sheet 第二页
python3 wps_split.py 工资表.xlsx -c C --all-sheets

# 批量拆分整个目录
python3 wps_split.py --dir ./待拆分 -c 部门

# 查看表格结构 / 预演（不写文件）
python3 wps_split.py 工资表.xlsx --info
python3 wps_split.py 工资表.xlsx -c 部门 --dry-run
```

完整参数表见 [使用说明.md](使用说明.md)。

---

## 📦 输出示例

```
工资表_拆分结果/
├── 财务部.xlsx
├── 销售部.xlsx
├── 技术部_研发中心.xlsx     ← 原分类含 "/"，自动替换为 "_"
├── 未分类.xlsx              ← 该列为空的行
└── 拆分清单.csv             ← 分类名 / 行数 / 文件名 对照清单
```

每个分表结构：

```
┌─────────────────────────────────┐
│ 某某公司 2026 年员工工资台账       │ ← 模板行：原样保留
│ 序号 │ 姓名 │ 部门 │ 工资 │ 日期 │  ← 表头：原样式 + 冻结
├──────┼──────┼──────┼──────┼─────┤
│  1   │ 赵某 │ 财务部│ 8,000│2020-01-05│ ← 仅本分类数据行
│  8   │ 钱某 │ 财务部│ 9,200│2019-07-14│   边框/底色/格式沿用
└─────────────────────────────────┘
```

---

## 🔧 实现原理

关键在 `xlsx_template.py`：不用 XML 重新序列化（那会丢命名空间声明导致文件损坏），而是对原 `sheet*.xml` 做**字符串切片** ——

1. 头部（`<sheetPr>` ~ `<cols>`、`<sheetData>` 起始标签）与尾部（`<mergeCells>`、`<autoFilter>`、`<drawing>` 等）**原样保留**；
2. 数据行搬原文，只改行号 `r` 属性，因此**样式索引 `s` 天然有效**；
3. 自动重设 `<dimension>` / `<autoFilter ref>` / 隐藏筛选定义名称的范围；
4. 清理 `calcChain.xml`、其它工作表部件、指向已删表的定义名称。

输出文件只含目标工作表，可被 WPS / Excel / LibreOffice 正常打开。

---

## 📁 文件结构

| 文件 | 说明 |
|---|---|
| `wps_split.py` | 拆分主程序：命令行 + 交互式向导 |
| `xlsx_template.py` | **模板克隆引擎**（保留原格式的核心） |
| `xlsx_lite.py` | 零依赖 xlsx 读写模块 |
| `wps_split_gui.py` | tkinter 图形界面 |
| `运行.sh` / `安装桌面快捷方式.sh` | UOS 一键启动与桌面图标 |
| `使用说明.md` | 完整文档（含参数表与 FAQ） |
| `tests/` | 测试数据与回归校验脚本 |

---

## ✅ 自检

在有 openpyxl 的机器上可跑完整回归：

```bash
python3 tests/make_data.py        # 造数据
python3 tests/make_template.py    # 造带完整格式的模板样表
python3 wps_split.py tests/测试数据/员工工资表.xlsx -c 部门
python3 wps_split.py tests/测试数据/模板样表.xlsx -c 部门 --header-row 2
python3 tests/verify.py           # → 数据一致性校验
python3 tests/verify_template.py  # → 模板格式保留校验
```

已覆盖场景：普通表、带格式模板表、CSV 输入、多列组合、忽略大小写、跳过空值、批量目录、`--all-sheets`、预演、交互向导。

---

## ⚠️ 已知限制

- 只支持标准 `.xlsx`。老式 `.xls` / `.et` 需先用 WPS 另存为 `.xlsx`（程序会明确提示）。
- 数据行上的浮动图片、图表、批注、数据透视表不保留；模板行区域的图片（如 LOGO）会保留。
- 跨表引用、或引用了被分到其它分表的行的公式无法自动修正，建议抽查。

---

## 📄 许可

MIT License — 详见 [LICENSE](LICENSE)。

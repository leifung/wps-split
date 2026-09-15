#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wps_split.py —— 按指定列内容把 WPS 表格拆分（分割）成多个 WPS 表格

适用：统信 UOS / 麒麟 / deepin / Windows / macOS，只需 Python 3.6+，零第三方依赖。
功能：
    * 按某一列（或多列组合）的取值，把一张大表拆成 N 个小表
    * 每个小表自动以"分类名称"命名，例如  财务部.xlsx、销售部.xlsx
    * 每个小表都保留原表头（可关闭），保留列宽、日期、数字格式
    * 自动清理文件名中的非法字符，超长自动截断，重名自动编号
    * 生成"拆分清单.csv"，记录每个分类的行数与输出文件
    * 支持命令行、交互式向导、图形界面三种用法

用法示例（统信 UOS 终端）：
    python3 wps_split.py 工资表.xlsx -c 部门
    python3 wps_split.py 工资表.xlsx -c C -o ./拆分结果
    python3 wps_split.py 工资表.xlsx -c 部门,年份 --sep "-"         # 按两列组合拆分
    python3 wps_split.py 工资表.xlsx -c 部门 --sheet 第2页          # 指定工作表
    python3 wps_split.py 工资表.xlsx --info                         # 只看表格信息
    python3 wps_split.py --dir ./待拆分                             # 批量拆分整个目录
    python3 wps_split.py                                            # 交互式向导
    python3 wps_split.py --gui                                      # 图形界面
"""

import argparse
import csv
import datetime
import os
import re
import sys

from xlsx_lite import (
    XlsxReader, XlsxWriter, col_letter_to_index, index_to_col_letter,
    is_xlsx_file, detect_old_binary, MAX_COLS,
)
from xlsx_template import TemplateSplitter

__version__ = "1.0.0"

ILLEGAL_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]')
RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL"} | {"COM%d" % i for i in range(1, 10)} | \
                 {"LPT%d" % i for i in range(1, 10)}
DEFAULT_EMPTY_NAME = "未分类"
DEFAULT_MAX_NAME_LEN = 60


# --------------------------------------------------------------------------
# 通用工具
# --------------------------------------------------------------------------
def log(msg, callback=None):
    print(msg)
    sys.stdout.flush()
    if callback:
        callback(str(msg))


def value_to_text(val, empty_name=DEFAULT_EMPTY_NAME):
    """把单元格的值转成用于分类和命名的文本"""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "是" if val else "否"
    if isinstance(val, datetime.datetime):
        if val.hour or val.minute or val.second:
            return val.strftime("%Y-%m-%d %H:%M")
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime.date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime.time):
        return val.strftime("%H:%M")
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def safe_filename(name, max_len=DEFAULT_MAX_NAME_LEN, used=None, ext=".xlsx"):
    """把分类值转成安全的合法文件名"""
    name = ILLEGAL_FILENAME_CHARS.sub("_", name or "")
    name = re.sub(r"\s+", " ", name).strip().strip(". ")
    if not name:
        name = DEFAULT_EMPTY_NAME
    if name.upper().split(".")[0] in RESERVED_NAMES:
        name = "_" + name
    if len(name) > max_len:
        name = name[:max_len].rstrip(". ")
    if used is not None:
        base = name
        i = 2
        while name.lower() in used:
            suffix = "_%d" % i
            name = base[:max(1, max_len - len(suffix))] + suffix
            i += 1
        used.add(name.lower())
    return name + ext


def unique_dir(path):
    """目录已存在时自动追加 _2、_3"""
    if not os.path.exists(path):
        return path
    i = 2
    while True:
        candidate = "%s_%d" % (path, i)
        if not os.path.exists(candidate):
            return candidate
        i += 1


def resolve_columns(reader, sheet, spec, header_row, has_header, header_values=None):
    """
    把用户给的列描述解析成列序号列表（1 开始）
    spec 可以是：列名 / 列字母 / 列序号，多个用英文逗号分隔
    """
    if header_values is None:
        header_values = []
    specs = [s.strip() for s in str(spec).split(",") if s.strip()]
    if not specs:
        raise ValueError("未指定拆分列")

    header_map = {}
    if has_header:
        for i, v in enumerate(header_values, start=1):
            key = value_to_text(v).strip()
            if key and key not in header_map:
                header_map[key] = i
            key2 = re.sub(r"\s+", "", key)
            if key2 and key2 not in header_map:
                header_map[key2] = i

    cols = []
    for s in specs:
        if re.match(r"^\d+$", s):
            idx = int(s)
            if not 1 <= idx <= MAX_COLS:
                raise ValueError("列序号超出范围：%s" % s)
            cols.append(idx)
            continue
        if re.match(r"^[A-Za-z]{1,3}$", s):
            cols.append(col_letter_to_index(s))
            continue
        # 按表头名匹配
        if not has_header:
            raise ValueError("当前设置为无表头，无法按列名 %r 定位，请改用列字母或序号" % s)
        key = s.strip()
        if key in header_map:
            cols.append(header_map[key])
            continue
        key2 = re.sub(r"\s+", "", key)
        if key2 in header_map:
            cols.append(header_map[key2])
            continue
        # 模糊包含
        hit = [i for k, i in header_map.items() if key in k or k in key]
        if len(hit) == 1:
            cols.append(hit[0])
            continue
        if len(hit) > 1:
            raise ValueError("列名 %r 匹配到多列（%s），请写完整列名" %
                             (s, "、".join(index_to_col_letter(i) for i in hit)))
        raise ValueError("表头中找不到列 %r。可用列：%s" %
                         (s, "、".join(sorted(header_map, key=lambda x: header_map[x]))))
    return cols


# --------------------------------------------------------------------------
# 输入读取（xlsx / csv）
# --------------------------------------------------------------------------
def read_csv_rows(path, encoding=None):
    if encoding:
        encodings = [encoding]
    else:
        encodings = ["utf-8-sig", "gb18030", "utf-8", "big5"]
    last_err = None
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                rows = [list(r) for r in csv.reader(f)]
            return rows, enc
        except (UnicodeDecodeError, LookupError) as e:
            last_err = e
    raise RuntimeError("无法识别 CSV 编码：%s（%s）" % (path, last_err))


def check_input_file(path):
    """校验输入文件，返回 (类型, 提示)"""
    if not os.path.isfile(path):
        raise RuntimeError("文件不存在：%s" % path)
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        if not is_xlsx_file(path):
            if detect_old_binary(path):
                raise RuntimeError(
                    "《%s》是老式二进制表格（.xls/.et），本工具无法直接读取。\n"
                    "请用 WPS 打开后【文件 → 另存为 → Excel 工作簿(*.xlsx)】再拆分。" % os.path.basename(path))
            raise RuntimeError("《%s》不是有效的 xlsx 文件，可能已损坏。" % os.path.basename(path))
        return "xlsx", ""
    if ext == ".csv":
        return "csv", ""
    if ext in (".xls", ".et"):
        raise RuntimeError(
            "《%s》是老式二进制表格，本工具无法直接读取。\n"
            "请用 WPS 打开后【文件 → 另存为 → Excel 工作簿(*.xlsx)】再拆分。" % os.path.basename(path))
    raise RuntimeError("不支持的格式：%s（仅支持 .xlsx / .xlsm / .csv）" % ext)


# --------------------------------------------------------------------------
# 核心拆分
# --------------------------------------------------------------------------
def split_table(input_path, column, output_dir=None, sheet=0, header_row=1,
                keep_header=True, skip_empty=False, empty_name=DEFAULT_EMPTY_NAME,
                trim=True, ignore_case=False, sep="-", max_name_len=DEFAULT_MAX_NAME_LEN,
                sort_by_name=False, summary=True, dry_run=False, all_sheets=False,
                verbose=True, log_callback=None,
                keep_format=True, template_rows=None):
    """
    按列拆分表格。返回结果 dict：
        {output_dir, groups:[{name, rows, file}], total_rows, skipped}

    keep_format  : True=沿用原表模板格式（默认）；False=普通模式（只保留值与列宽）
    template_rows: 原样保留的前 N 行（标题/表头），默认等于表头行号
    """
    ftype, _ = check_input_file(input_path)
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    src_dir = os.path.dirname(os.path.abspath(input_path))

    if output_dir:
        output_dir = os.path.abspath(os.path.expanduser(output_dir))
    else:
        output_dir = os.path.join(src_dir, base_name + "_拆分结果")

    sheets_to_do = []
    if ftype == "csv":
        sheets_to_do = [(None, None)]
    else:
        with XlsxReader(input_path) as rd:
            names = rd.sheet_names
        if all_sheets:
            sheets_to_do = [(n, n) for n in names]
        else:
            sheets_to_do = [(names[0] if isinstance(sheet, int) else sheet, sheet)]

    all_result = {"output_dir": output_dir, "groups": [], "total_rows": 0, "skipped": 0,
                  "files": [], "sheets": []}

    for sheet_name, sheet_key in sheets_to_do:
        res = _split_one(input_path, ftype, sheet_key if ftype == "xlsx" else None,
                         sheet_name, column, output_dir, header_row, keep_header,
                         skip_empty, empty_name, trim, ignore_case, sep, max_name_len,
                         sort_by_name, dry_run, all_sheets and len(sheets_to_do) > 1,
                         verbose, log_callback, keep_format, template_rows)
        all_result["groups"].extend(res["groups"])
        all_result["total_rows"] += res["total_rows"]
        all_result["skipped"] += res["skipped"]
        all_result["files"].extend(res["files"])
        all_result["sheets"].append({"sheet": sheet_name, "groups": res["groups"]})

    # 拆分清单
    if summary and not dry_run and all_result["files"]:
        _write_summary(output_dir, all_result, log_callback)
    return all_result


def _split_one(input_path, ftype, sheet_key, sheet_name, column, output_dir, header_row,
               keep_header, skip_empty, empty_name, trim, ignore_case, sep, max_name_len,
               sort_by_name, dry_run, sub_dir, verbose, log_cb,
               keep_format=True, template_rows=None):
    result = {"groups": [], "total_rows": 0, "skipped": 0, "files": [], "mode": "template"}
    label = ("[%s] " % sheet_name) if sheet_name else ""

    # ---------- 读数据（连同原始行号，行号用于按模板原样搬运整行）----------
    rows_indexed = []      # [(原始行号, 行数据)]
    widths = {}
    if ftype == "csv":
        rows, enc = read_csv_rows(input_path)
        log("%s已读取 CSV（编码 %s），共 %d 行" % (label, enc, len(rows)), log_cb)
        rows_indexed = list(enumerate(rows, start=1))
    else:
        if sheet_key is None:
            sheet_key = 0
        reader = XlsxReader(input_path)
        try:
            widths = reader.read_cols_width(sheet_key)
            if isinstance(sheet_key, int):
                display = sheet_name or reader.sheet_names[sheet_key]
            else:
                display = sheet_name or str(sheet_key)
            rows_indexed = list(reader.iter_rows(sheet_key, with_index=True))
        finally:
            reader.close()
        log("%s已读取工作表《%s》，共 %d 行" % (label, display, len(rows_indexed)), log_cb)

    if not rows_indexed:
        raise RuntimeError("表格为空，无需拆分")

    # ---------- 表头 ----------
    has_header = bool(header_row) and header_row > 0
    header_values = []
    hidx = -1
    if has_header:
        for i, (no, _r) in enumerate(rows_indexed):
            if no == header_row:
                hidx = i
                break
        if hidx < 0:    # 表头行号在文件中不存在（前面有空行），退化为按序号取
            hidx = min(int(header_row) - 1, len(rows_indexed) - 1)
        header_values = rows_indexed[hidx][1]
        data_pairs = rows_indexed[hidx + 1:]
    else:
        data_pairs = rows_indexed

    cols = resolve_columns(None, sheet_key, column, header_row, has_header, header_values)
    col_desc = "、".join("%s列(%s)" % (index_to_col_letter(c),
                                    (value_to_text(header_values[c - 1]) if c <= len(header_values) else ""))
                        for c in cols)
    log("%s拆分依据：%s" % (label, col_desc), log_cb)
    max_data_col = max((len(r) for _no, r in rows_indexed if r), default=0)
    if max(cols) > max_data_col:
        log("%s警告：表格只有 %d 列，指定的 %s 列没有数据，所有行都会归入“%s”"
            % (label, max_data_col, index_to_col_letter(max(cols)), empty_name), log_cb)

    # ---------- 统计分类 ----------
    order = []
    counts = {}
    key_map = {}          # 归一化 key -> 展示用分类名
    row_ids = {}          # 归一化 key -> 原表行号列表
    header_width = len(header_values)

    def _classify(row):
        """返回 (分类显示名, 归一化key)；整行空或需跳过时返回 (None, None)"""
        if not row or all(v is None or v == "" for v in row):
            return None, None
        parts = []
        for c in cols:
            v = row[c - 1] if c <= len(row) else None
            t = value_to_text(v)
            if trim:
                t = t.strip()
            if not t and len(cols) > 1:
                t = empty_name      # 多列组合时空的部分用占位名，避免出现 "_否" 这类怪名
            parts.append(t)
        raw = sep.join(parts) if len(parts) > 1 else parts[0]
        if not raw:
            if skip_empty:
                return None, None
            disp = empty_name
        else:
            disp = raw
        return disp, (disp.lower() if ignore_case else disp)

    for row_no, row in data_pairs:
        disp, key = _classify(row)
        if key is None:
            continue
        if key not in key_map:
            key_map[key] = disp
            order.append(key)
            counts[key] = 0
            row_ids[key] = []
        counts[key] += 1
        row_ids[key].append(row_no)

    if not order:
        raise RuntimeError("没有可用于拆分的数据行")

    if sort_by_name:
        order.sort(key=lambda k: key_map[k])

    log("%s共发现 %d 个分类，数据行 %d 行" % (label, len(order), sum(counts.values())), log_cb)

    # ---------- 文件名分配 ----------
    used = set()
    file_of = {}
    for key in order:
        fname = safe_filename(key_map[key], max_name_len, used)
        file_of[key] = fname

    out_dir = output_dir
    if sub_dir:
        out_dir = os.path.join(output_dir, safe_filename(sheet_name or "Sheet1", 40, None, ""))
    if not dry_run:
        os.makedirs(out_dir, exist_ok=True)

    # ---------- 模板行（标题行/表头行：原样保留在原位置）----------
    if template_rows is None:
        tpl_limit = int(header_row) if has_header else 0
    else:
        tpl_limit = int(template_rows)
    template_ids = [no for no, _r in rows_indexed if no <= tpl_limit] if tpl_limit > 0 else []
    if not keep_header and has_header:      # 不保留表头时，模板行收缩到表头行之前
        template_ids = [no for no in template_ids if no < header_row]
    if has_header and tpl_limit > header_row:
        log("%s提示：保留前 %d 行作为模板，其中第 %d~%d 行是数据行，它们会出现在每一个分表里"
            % (label, tpl_limit, header_row + 1, tpl_limit), log_cb)

    # ---------- 输出模式 ----------
    splitter = None
    if keep_format and ftype == "xlsx":
        try:
            splitter = TemplateSplitter(input_path, sheet_key)
        except Exception as e:
            log("%s警告：无法沿用原表模板格式（%s），改用普通模式输出" % (label, e), log_cb)
            splitter = None
    if splitter is None:
        result["mode"] = "plain"
    elif verbose:
        log("%s输出模式：沿用原表模板格式（字体/边框/列宽/行高/数字格式全部保留）" % label, log_cb)

    # ---------- 写文件 ----------
    row_of = dict(rows_indexed)
    total = sum(counts.values())
    skipped = len(data_pairs) - total
    done = 0
    for key in order:
        g = {"name": key_map[key], "rows": counts[key], "file": file_of[key],
             "path": os.path.join(out_dir, file_of[key]), "sheet": sheet_name}
        result["groups"].append(g)
        if dry_run:
            done += counts[key]
            continue

        written = False
        if splitter is not None:
            try:
                splitter.build(g["path"], row_ids[key], template_ids,
                               sheet_name=sheet_name or None)
                written = True
            except Exception as e:
                log("%s警告：%s 沿用模板失败（%s），该文件改用普通模式" % (label, g["file"], e), log_cb)
                splitter = None
                result["mode"] = "plain"
        if not written:
            wb = XlsxWriter()
            sh = wb.add_sheet(sheet_name or "Sheet1")
            for c, w in widths.items():
                sh.set_col_width(c, w)
            if keep_header and has_header:
                sh.append_row(header_values, style=1)
            for no in row_ids[key]:
                values = list(row_of.get(no) or [])
                if len(values) < header_width:
                    values += [None] * (header_width - len(values))
                sh.append_row(values)
            wb.save(g["path"])
        result["files"].append(g["path"])
        done += counts[key]
        if verbose and total and done % 2000 == 0:
            log("%s  进度 %d/%d (%.0f%%)" % (label, done, total, done * 100.0 / total), log_cb)

    result["total_rows"] = done
    result["skipped"] = skipped
    result["output_dir"] = out_dir

    if verbose:
        log("%s拆分完成：%d 个文件，%d 行数据%s" %
            (label, len(order), done, ("，跳过 %d 行" % skipped) if skipped else ""), log_cb)
        if not dry_run:
            log("%s输出目录：%s" % (label, out_dir), log_cb)
    return result


def _write_summary(output_dir, result, log_cb=None):
    path = os.path.join(output_dir, "拆分清单.csv")
    try:
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["序号", "分类名称", "数据行数", "输出文件名", "输出路径"])
            for i, g in enumerate(result["groups"], start=1):
                w.writerow([i, g["name"], g["rows"], g["file"], g["path"]])
        log("已生成拆分清单：%s" % path, log_cb)
    except OSError as e:
        log("警告：拆分清单写入失败（%s）" % e, log_cb)


# --------------------------------------------------------------------------
# 表格信息
# --------------------------------------------------------------------------
def show_info(path):
    ftype, _ = check_input_file(path)
    print("文件：%s" % path)
    print("大小：%.1f KB" % (os.path.getsize(path) / 1024.0))
    if ftype == "csv":
        rows, enc = read_csv_rows(path)
        print("类型：CSV（%s），共 %d 行" % (enc, len(rows)))
        for i, r in enumerate(rows[:3], start=1):
            print("  第%d行：%s" % (i, " | ".join(value_to_text(v) for v in r[:12])))
        return
    with XlsxReader(path) as rd:
        print("工作表（%d 个）：%s" % (len(rd.sheet_names), "、".join(rd.sheet_names)))
        for idx, name in enumerate(rd.sheet_names):
            n = 0
            first = None
            for row in rd.iter_rows(idx):
                if first is None:
                    first = row
                n += 1
            print("\n[%d] %s —— 约 %d 行" % (idx + 1, name, n))
            if first:
                cols = []
                for i, v in enumerate(first, start=1):
                    t = value_to_text(v)
                    cols.append("%s:%s" % (index_to_col_letter(i), t or "(空)"))
                print("    表头：" + " | ".join(cols[:20]))


# --------------------------------------------------------------------------
# 交互向导
# --------------------------------------------------------------------------
def interactive_wizard():
    print("=" * 62)
    print("  WPS 表格按列拆分工具 v%s（统信 UOS 版）" % __version__)
    print("=" * 62)
    print("提示：可直接把文件拖进终端窗口，回车确认。\n")

    path = ""
    while True:
        path = input("① 请输入要拆分的表格路径：").strip().strip("'\"")
        if not path:
            print("   未输入，退出。")
            return 1
        try:
            ftype, _ = check_input_file(path)
            break
        except Exception as e:
            print("   × %s" % e)

    sheet = 0
    if ftype == "xlsx":
        with XlsxReader(path) as rd:
            names = rd.sheet_names
        if len(names) > 1:
            print("   工作表：" + "、".join("%d.%s" % (i + 1, n) for i, n in enumerate(names)))
            s = input("② 选择工作表序号（回车=1）：").strip()
            if s.isdigit() and 1 <= int(s) <= len(names):
                sheet = int(s) - 1

    header_row = 1
    s = input("③ 表头在第几行（回车=1，无表头填 0）：").strip()
    if s.isdigit():
        header_row = int(s)

    column = ""
    while True:
        column = input("④ 按哪一列拆分（列名/列字母/序号，多列用英文逗号分隔）：").strip()
        if column:
            break
        print("   必须指定拆分列。")

    out = input("⑤ 输出目录（回车=自动在源文件旁创建）：").strip().strip("'\"")
    out = out or None

    tpl = input("⑥ 每个分表原样保留前几行（标题/表头，回车=%d）：" % header_row).strip()
    template_rows = int(tpl) if tpl.isdigit() else None

    print("⑦ 其他选项（直接回车=使用默认值）")
    keep_header = input("   每个分表都保留表头？[Y/n]：").strip().lower() != "n"
    keep_format = input("   沿用原表的模板格式（字体/边框/列宽/数字格式）？[Y/n]：").strip().lower() != "n"
    skip_empty = input("   分类为空的行直接跳过，不生成'未分类'表？[y/N]：").strip().lower() == "y"
    summary = input("   生成拆分清单.csv？[Y/n]：").strip().lower() != "n"

    print()
    try:
        result = split_table(path, column, output_dir=out, sheet=sheet, header_row=header_row,
                             keep_header=keep_header, skip_empty=skip_empty, summary=summary,
                             keep_format=keep_format, template_rows=template_rows)
    except Exception as e:
        print("× 拆分失败：%s" % e)
        return 2
    print("\n完成！共 %d 个分类，%d 行数据 → %s" %
          (len(result["groups"]), result["total_rows"], result["output_dir"]))
    return 0


# --------------------------------------------------------------------------
# 命令行
# --------------------------------------------------------------------------
def build_parser():
    p = argparse.ArgumentParser(
        prog="wps_split.py",
        description="按指定列的内容把一个 WPS 表格拆分成多个 WPS 表格",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n"
               "  python3 wps_split.py 工资表.xlsx -c 部门\n"
               "  python3 wps_split.py 工资表.xlsx -c C -o ./out\n"
               "  python3 wps_split.py 工资表.xlsx -c 部门,年份 --sep -\n"
               "  python3 wps_split.py --dir ./待拆分 -c 部门\n")
    p.add_argument("input", nargs="?", help="要拆分的表格（.xlsx/.xlsm/.csv）")
    p.add_argument("-c", "--column", help="拆分依据列：列名 / 列字母(A,B,C) / 列序号(1,2,3)，多列用英文逗号分隔")
    p.add_argument("-o", "--output-dir", help="输出目录，默认在源文件旁创建“xxx_拆分结果”")
    p.add_argument("--sheet", default=0, help="工作表名或序号（从 0 开始），默认 0（第一个）")
    p.add_argument("--all-sheets", action="store_true", help="对工作簿中的每个工作表都执行拆分")
    p.add_argument("--header-row", type=int, default=1, help="表头所在行号，默认 1；填 0 表示无表头")
    p.add_argument("--no-header", action="store_true", help="拆分后的表不保留表头行")
    p.add_argument("--skip-empty", action="store_true", help="分类为空的行直接跳过，不生成“未分类”表")
    p.add_argument("--empty-name", default=DEFAULT_EMPTY_NAME, help="空值分类的命名，默认“未分类”")
    p.add_argument("--no-trim", action="store_true", help="不去除分类值首尾空格")
    p.add_argument("--ignore-case", action="store_true", help="分类时忽略英文大小写（合并 ABC 与 abc）")
    p.add_argument("--sep", default="-", help="多列组合时分类名的连接符，默认 -")
    p.add_argument("--max-name-len", type=int, default=DEFAULT_MAX_NAME_LEN, help="文件名最大字符数，默认 60")
    p.add_argument("--sort", action="store_true", help="按分类名称排序输出（默认按出现顺序）")
    p.add_argument("--template-rows", type=int, default=None,
                   help="每个分表原样保留的前 N 行（标题行/表头行），默认等于表头行号")
    p.add_argument("--plain", action="store_true",
                   help="普通模式：不沿用原表模板格式（一般不用，仅当模板复制异常时兜底）")
    p.add_argument("--no-summary", action="store_true", help="不生成拆分清单.csv")
    p.add_argument("--dry-run", action="store_true", help="预演：只统计分类和行数，不写文件")
    p.add_argument("--dir", dest="input_dir", help="批量处理：指定目录下所有表格")
    p.add_argument("--info", action="store_true", help="只显示表格信息（工作表、表头），不拆分")
    p.add_argument("--gui", action="store_true", help="启动图形界面")
    p.add_argument("-q", "--quiet", action="store_true", help="静默模式，减少输出")
    p.add_argument("-v", "--version", action="version", version="wps_split %s" % __version__)
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)

    if args.gui:
        try:
            import wps_split_gui
        except ImportError as e:
            print("无法启动图形界面：%s" % e)
            print("请先安装 tkinter：sudo apt install python3-tk")
            return 3
        return wps_split_gui.main()

    if args.info:
        if not args.input:
            print("× 请用 --info 时指定文件")
            return 2
        show_info(args.input)
        return 0

    if args.input_dir:
        if not args.column:
            print("× 批量模式必须用 -c 指定拆分列")
            return 2
        d = os.path.abspath(os.path.expanduser(args.input_dir))
        if not os.path.isdir(d):
            print("× 目录不存在：%s" % d)
            return 2
        files = [os.path.join(d, f) for f in sorted(os.listdir(d))
                 if f.lower().endswith((".xlsx", ".xlsm", ".csv")) and not os.path.basename(f).startswith("~$")
                 and not f.startswith("拆分清单")]
        if not files:
            print("× 目录中没有可拆分的表格：%s" % d)
            return 2
        base_out = args.output_dir or os.path.join(d, "拆分结果")
        ok = 0
        for i, f in enumerate(files, start=1):
            print("\n===== (%d/%d) %s =====" % (i, len(files), os.path.basename(f)))
            try:
                split_table(f, args.column, output_dir=os.path.join(base_out, os.path.splitext(os.path.basename(f))[0]),
                            sheet=args.sheet, header_row=args.header_row,
                            keep_header=not args.no_header, skip_empty=args.skip_empty,
                            empty_name=args.empty_name, trim=not args.no_trim,
                            ignore_case=args.ignore_case, sep=args.sep,
                            max_name_len=args.max_name_len, sort_by_name=args.sort,
                            summary=not args.no_summary, dry_run=args.dry_run,
                            all_sheets=args.all_sheets, verbose=not args.quiet,
                            keep_format=not args.plain, template_rows=args.template_rows)
                ok += 1
            except Exception as e:
                print("× 失败：%s" % e)
        print("\n批量完成：成功 %d / 共 %d，输出：%s" % (ok, len(files), base_out))
        return 0 if ok else 2

    if not args.input:
        return interactive_wizard()

    if not args.column:
        print("× 请用 -c 指定拆分列（列名/列字母/序号），或直接运行不带参数的交互式向导\n")
        build_parser().print_help()
        return 2

    sheet = args.sheet
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    try:
        split_table(args.input, args.column, output_dir=args.output_dir, sheet=sheet,
                    header_row=args.header_row, keep_header=not args.no_header,
                    skip_empty=args.skip_empty, empty_name=args.empty_name,
                    trim=not args.no_trim, ignore_case=args.ignore_case, sep=args.sep,
                    max_name_len=args.max_name_len, sort_by_name=args.sort,
                    summary=not args.no_summary, dry_run=args.dry_run,
                    all_sheets=args.all_sheets, verbose=not args.quiet,
                    keep_format=not args.plain, template_rows=args.template_rows)
    except Exception as e:
        print("× 拆分失败：%s" % e)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())

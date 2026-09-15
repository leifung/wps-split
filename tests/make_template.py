# -*- coding: utf-8 -*-
"""生成一份“带完整格式的模板样表”，用于验证拆分后是否保留原模板格式（需要 openpyxl）"""
import datetime
import os
import random

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE, "测试数据")
os.makedirs(OUT_DIR, exist_ok=True)

HEADERS = ["序号", "姓名", "部门", "基本工资", "绩效系数", "入职日期", "是否在职", "备注", "应发合计"]
DEPTS = ["财务部", "销售部", "市场部", "技术部/研发中心", "人力资源与行政管理部", "abc", "ABC", None, "重?名:部|门*测\"试<>"]
NAMES = "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜"

thin = Side(style="thin", color="9B9B9B")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
FILL_TITLE = PatternFill("solid", fgColor="DDEBF7")
FILL_HEADER = PatternFill("solid", fgColor="305496")
FILL_ODD = PatternFill("solid", fgColor="FFF2CC")
FILL_EVEN = PatternFill("solid", fgColor="FFFFFF")


def build():
    wb = Workbook()
    ws = wb.active
    ws.title = "员工工资"

    # 第 1 行：大标题（合并 + 居中 + 加粗 + 底色）
    ws["A1"] = "某某公司 2026 年员工工资台账"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    c = ws["A1"]
    c.font = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = FILL_TITLE
    ws.row_dimensions[1].height = 28

    # 第 2 行：表头
    for i, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=2, column=i, value=h)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[2].height = 22

    # 数据行
    random.seed(20260915)
    start = datetime.datetime(2015, 1, 1)
    for i in range(1, 201):
        r = i + 2
        dept = DEPTS[(i - 1) % len(DEPTS)]
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=NAMES[(i * 7) % len(NAMES)] + "某")
        ws.cell(row=r, column=3, value=dept)
        ws.cell(row=r, column=4, value=random.randint(4000, 15000)).number_format = "#,##0"
        ws.cell(row=r, column=5, value=round(random.uniform(0.6, 1.5), 2)).number_format = "0.00"
        ws.cell(row=r, column=6, value=start + datetime.timedelta(days=random.randint(0, 3000))
                ).number_format = "yyyy-mm-dd"
        ws.cell(row=r, column=7, value=(i % 3 != 0))
        ws.cell(row=r, column=8, value="" if i % 17 == 0 else "备注%d" % i)
        ws.cell(row=r, column=9, value="=D%d*E%d" % (r, r)).number_format = "#,##0.00"
        fill = FILL_ODD if i % 2 else FILL_EVEN
        for col in range(1, 10):
            cell = ws.cell(row=r, column=col)
            cell.border = BORDER
            cell.fill = fill
            if col in (3, 7):
                cell.alignment = Alignment(horizontal="center")
        ws.row_dimensions[r].height = 18

    # 列宽 / 冻结 / 自动筛选 / 条件格式
    widths = [6, 12, 24, 12, 10, 14, 10, 18, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = "A2:I201"

    # 第二页（用于 --all-sheets 测试）
    ws2 = wb.create_sheet("第二页")
    ws2.append(["品类", "数量", "单价"])
    for i in range(1, 6):
        ws2.append([random.choice(["盐", "纯碱", "氯化钙"]), random.randint(1, 50), round(random.uniform(1, 9), 2)])
    ws2.column_dimensions["A"].width = 16

    path = os.path.join(OUT_DIR, "模板样表.xlsx")
    wb.save(path)
    print("已生成：%s（200 行数据，含标题/表头样式/边框/交替底色/数字格式/冻结/筛选/公式）" % path)
    return path


if __name__ == "__main__":
    build()

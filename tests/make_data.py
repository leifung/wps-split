# -*- coding: utf-8 -*-
"""生成测试用表格（需要 openpyxl，仅用于开发验证，工具本身不需要）"""
import datetime
import os
import random

from openpyxl import Workbook

random.seed(20260915)

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "测试数据")

DEPTS = ["财务部", "销售部", "技术部/研发中心", "人力资源与行政管理部", "市场部", "abc", "ABC"]
NAMES = ["张伟", "李娜", "王强", "刘洋", "陈静", "赵磊", "孙芳", "周涛", "吴敏", "郑凯"]


def main():
    os.makedirs(OUT, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "员工工资"
    header = ["序号", "姓名", "部门", "基本工资", "绩效系数", "入职日期", "是否在职", "备注"]
    ws.append(header)

    rows = []
    for i in range(1, 301):
        dept = random.choice(DEPTS)
        if i % 37 == 0:
            dept = None
        if i % 53 == 0:
            dept = "  财务部  "            # 前后带空格，测 trim
        if i % 61 == 0:
            dept = "重?名:部|门*测\"试<>"   # 含非法文件名字符
        rows.append([
            i,
            random.choice(NAMES) + str(i),
            dept,
            random.randint(4000, 20000),
            round(random.uniform(0.6, 1.8), 2),
            datetime.datetime(2015, 1, 1) + datetime.timedelta(days=random.randint(0, 3000)),
            i % 5 != 0,
            "含,逗号 与 空格" if i % 11 == 0 else None,
        ])

    for r in rows:
        ws.append(r)

    for col, w in zip("ABCDEFGH", [6, 12, 22, 12, 10, 14, 10, 24]):
        ws.column_dimensions[col].width = w

    # 第二个工作表
    ws2 = wb.create_sheet("第二页")
    ws2.append(["品类", "数量"])
    for j, c in enumerate(["盐", "纯碱", "氯化钙", "硫酸钠"], start=1):
        ws2.append([c, j * 10])

    path = os.path.join(OUT, "员工工资表.xlsx")
    wb.save(path)
    print("已生成：%s（%d 行数据）" % (path, len(rows)))

    # 同时生成一份 csv
    import csv
    csv_path = os.path.join(OUT, "商品清单.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["编号", "品类", "单价"])
        for i in range(1, 51):
            w.writerow([i, random.choice(["盐", "纯碱", "氯化钙"]), random.randint(100, 999)])
    print("已生成：%s" % csv_path)


if __name__ == "__main__":
    main()

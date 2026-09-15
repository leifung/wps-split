# -*- coding: utf-8 -*-
"""
wps_split_gui.py —— 图形界面版（tkinter，Python 标准库自带）

启动：
    python3 wps_split_gui.py
    # 或
    python3 wps_split.py --gui
    # 或双击  运行.sh

若提示缺少 tkinter，在统信 UOS 终端执行： sudo apt install python3-tk
"""

import os
import queue
import sys
import threading
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, font as tkfont
except ImportError:  # pragma: no cover
    sys.stderr.write("缺少 tkinter，请执行：sudo apt install python3-tk\n")
    raise

import wps_split
from xlsx_lite import XlsxReader, index_to_col_letter


def pick_font(root):
    """挑一个系统中存在的中文字体"""
    try:
        families = set(tkfont.families(root))
    except Exception:
        return None
    for f in ("Noto Sans CJK SC", "WenQuanYi Micro Hei", "WenQuanYi Zen Hei",
              "Source Han Sans CN", "Microsoft YaHei", "SimHei"):
        if f in families:
            return f
    return None


class App(object):
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.headers = []
        self.sheet_names = []
        self.running = False

        fam = pick_font(root)
        if fam:
            try:
                root.option_add("*Font", "%s 10" % fam)
            except Exception:
                pass

        root.title("WPS 表格按列拆分工具  v%s" % wps_split.__version__)
        root.geometry("820x640")
        root.minsize(760, 580)
        self._build()
        self._poll()

    # ------------------------------------------------------------------ UI
    def _build(self):
        pad = {"padx": 8, "pady": 5}
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # 文件
        frm = ttk.LabelFrame(main, text=" 1. 选择要拆分的表格 ", padding=8)
        frm.pack(fill=tk.X, **pad)
        self.var_file = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_file).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(frm, text="浏览…", command=self.browse_file).pack(side=tk.LEFT)

        # 工作表 + 表头行
        frm2 = ttk.LabelFrame(main, text=" 2. 工作表与表头 ", padding=8)
        frm2.pack(fill=tk.X, **pad)
        ttk.Label(frm2, text="工作表：").grid(row=0, column=0, sticky=tk.W)
        self.var_sheet = tk.StringVar()
        self.cmb_sheet = ttk.Combobox(frm2, textvariable=self.var_sheet, width=22, state="readonly")

        self.cmb_sheet.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        self.cmb_sheet.bind("<<ComboboxSelected>>", lambda e: self.load_header())
        ttk.Label(frm2, text="表头行号：").grid(row=0, column=2, sticky=tk.W)
        self.var_header = tk.StringVar(value="1")
        ttk.Spinbox(frm2, from_=0, to=100, width=5, textvariable=self.var_header,
                    command=self.load_header).grid(row=0, column=3, sticky=tk.W)
        ttk.Button(frm2, text="读取列", command=self.load_header).grid(row=0, column=4, padx=10)

        # 拆分列
        frm3 = ttk.LabelFrame(main, text=" 3. 拆分依据列（新表格按该列内容命名）", padding=8)
        frm3.pack(fill=tk.X, **pad)
        self.var_col = tk.StringVar()
        self.cmb_col = ttk.Combobox(frm3, textvariable=self.var_col, width=30)
        self.cmb_col.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Label(frm3, text="多列用英文逗号分隔，如：部门,年份").pack(side=tk.LEFT)

        # 输出目录
        frm4 = ttk.LabelFrame(main, text=" 4. 输出目录（留空则自动建在源文件旁）", padding=8)
        frm4.pack(fill=tk.X, **pad)
        self.var_out = tk.StringVar()
        ttk.Entry(frm4, textvariable=self.var_out).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(frm4, text="选择…", command=self.browse_out).pack(side=tk.LEFT)

        # 选项
        frm5 = ttk.LabelFrame(main, text=" 5. 选项 ", padding=8)
        frm5.pack(fill=tk.X, **pad)
        self.var_keep_header = tk.BooleanVar(value=True)
        self.var_skip_empty = tk.BooleanVar(value=False)
        self.var_summary = tk.BooleanVar(value=True)
        self.var_trim = tk.BooleanVar(value=True)
        self.var_sort = tk.BooleanVar(value=False)
        self.var_all_sheets = tk.BooleanVar(value=False)
        self.var_keep_format = tk.BooleanVar(value=True)
        opts = [("★ 沿用原表模板格式", self.var_keep_format),
                ("每个分表保留表头", self.var_keep_header),
                ("跳过分类为空的行", self.var_skip_empty),
                ("生成拆分清单.csv", self.var_summary),
                ("去除分类值首尾空格", self.var_trim),
                ("按分类名排序", self.var_sort),
                ("拆分所有工作表", self.var_all_sheets)]
        for i, (text, var) in enumerate(opts):
            ttk.Checkbutton(frm5, text=text, variable=var).grid(row=i // 3, column=i % 3, sticky=tk.W, padx=6, pady=2)

        # 按钮 + 进度
        frm6 = ttk.Frame(main)
        frm6.pack(fill=tk.X, **pad)
        self.btn_run = ttk.Button(frm6, text="开始拆分", command=self.start)
        self.btn_run.pack(side=tk.LEFT)
        self.btn_dry = ttk.Button(frm6, text="预演（只看统计）", command=lambda: self.start(dry=True))
        self.btn_dry.pack(side=tk.LEFT, padx=8)
        ttk.Button(frm6, text="打开输出目录", command=self.open_out).pack(side=tk.LEFT)
        self.var_prog = tk.StringVar(value="就绪")
        ttk.Label(frm6, textvariable=self.var_prog).pack(side=tk.RIGHT)

        # 日志
        frm7 = ttk.LabelFrame(main, text=" 运行日志 ", padding=6)
        frm7.pack(fill=tk.BOTH, expand=True, **pad)
        self.txt = tk.Text(frm7, height=12, wrap=tk.WORD, relief=tk.FLAT)
        sb = ttk.Scrollbar(frm7, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # ------------------------------------------------------------- 事件
    def browse_file(self):
        path = filedialog.askopenfilename(
            title="选择表格文件",
            filetypes=[("表格文件", "*.xlsx *.xlsm *.csv"), ("Excel 工作簿", "*.xlsx"),
                       ("CSV 文件", "*.csv"), ("所有文件", "*.*")])
        if path:
            self.var_file.set(path)
            self.load_sheets()

    def browse_out(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.var_out.set(d)

    def load_sheets(self):
        path = self.var_file.get().strip()
        if not path or not os.path.isfile(path):
            return
        try:
            wps_split.check_input_file(path)
        except Exception as e:
            messagebox.showerror("文件无法处理", str(e))
            return
        if path.lower().endswith(".csv"):
            self.sheet_names = []
            self.cmb_sheet["values"] = ["(CSV)"]
            self.var_sheet.set("(CSV)")
            self.load_header()
            return
        try:
            with XlsxReader(path) as rd:
                self.sheet_names = rd.sheet_names
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        self.cmb_sheet["values"] = self.sheet_names
        self.var_sheet.set(self.sheet_names[0] if self.sheet_names else "")
        self.load_header()

    def load_header(self):
        path = self.var_file.get().strip()
        if not path or not os.path.isfile(path):
            return
        try:
            hr = int(self.var_header.get())
        except ValueError:
            hr = 1
        headers = []
        try:
            if path.lower().endswith(".csv"):
                rows, _ = wps_split.read_csv_rows(path)
                if hr >= 1 and len(rows) >= hr:
                    headers = rows[hr - 1]
            else:
                name = self.var_sheet.get()
                idx = self.sheet_names.index(name) if name in self.sheet_names else 0
                with XlsxReader(path) as rd:
                    for i, row in enumerate(rd.iter_rows(idx), start=1):
                        if i == hr:
                            headers = row
                            break
        except Exception as e:
            self.txt.insert(tk.END, "读取表头失败：%s\n" % e)
            return
        self.headers = headers
        values = []
        for i, v in enumerate(headers, start=1):
            t = wps_split.value_to_text(v)
            values.append(t if t else "%s列" % index_to_col_letter(i))
        self.cmb_col["values"] = values
        if values and not self.var_col.get():
            self.var_col.set(values[0])
        self.txt.insert(tk.END, "已读取 %d 列：%s\n" % (len(values), "、".join(values[:15])))

    def open_out(self):
        d = self.var_out.get().strip()
        if not d and self.var_file.get():
            d = os.path.join(os.path.dirname(os.path.abspath(self.var_file.get())),
                             os.path.splitext(os.path.basename(self.var_file.get()))[0] + "_拆分结果")
        if d and os.path.isdir(d):
            try:
                os.startfile(d)  # Windows
            except Exception:
                try:
                    import subprocess
                    subprocess.Popen(["xdg-open", d])
                except Exception:
                    pass
        else:
            messagebox.showinfo("提示", "输出目录还不存在：%s" % d)

    # ------------------------------------------------------------- 执行
    def start(self, dry=False):
        if self.running:
            return
        path = self.var_file.get().strip().strip("'\"")
        col = self.var_col.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择要拆分的表格")
            return
        if not col:
            messagebox.showwarning("提示", "请填写或选择拆分依据列")
            return
        self.running = True
        self.btn_run.config(state=tk.DISABLED)
        self.btn_dry.config(state=tk.DISABLED)
        self.var_prog.set("正在拆分…")
        self.txt.insert(tk.END, "\n===== 开始 =====\n")

        sheet_name = self.var_sheet.get()
        sheet_idx = self.sheet_names.index(sheet_name) if sheet_name in self.sheet_names else 0
        args = dict(
            input_path=path, column=col,
            output_dir=self.var_out.get().strip() or None,
            sheet=sheet_idx, header_row=int(self.var_header.get() or 1),
            keep_header=self.var_keep_header.get(), skip_empty=self.var_skip_empty.get(),
            trim=self.var_trim.get(), sort_by_name=self.var_sort.get(),
            summary=self.var_summary.get(), dry_run=dry,
            all_sheets=self.var_all_sheets.get(),
            keep_format=self.var_keep_format.get(),
            log_callback=self.q.put,
        )
        t = threading.Thread(target=self._worker, kwargs=args)
        t.daemon = True
        t.start()

    def _worker(self, **kw):
        try:
            result = wps_split.split_table(**kw)
            self.q.put(("__done__", result))
        except Exception:
            self.q.put(("__error__", traceback.format_exc()))

    def _poll(self):
        try:
            while True:
                item = self.q.get_nowait()
                if isinstance(item, tuple) and item[0] == "__done__":
                    res = item[1]
                    self.txt.insert(tk.END, "\n✔ 完成：%d 个分类，%d 行数据\n输出目录：%s\n"
                                    % (len(res["groups"]), res["total_rows"], res["output_dir"]))
                    self.var_prog.set("完成：%d 个文件" % len(res["files"]))
                    self.running = False
                    self.btn_run.config(state=tk.NORMAL)
                    self.btn_dry.config(state=tk.NORMAL)
                    if self.var_out.get().strip() == "":
                        self.var_out.set(res["output_dir"])
                    messagebox.showinfo("拆分完成",
                                        "共 %d 个分类，%d 行数据\n\n输出目录：\n%s"
                                        % (len(res["groups"]), res["total_rows"], res["output_dir"]))
                elif isinstance(item, tuple) and item[0] == "__error__":
                    self.txt.insert(tk.END, "\n✘ 出错：\n%s\n" % item[1])
                    self.var_prog.set("出错")
                    self.running = False
                    self.btn_run.config(state=tk.NORMAL)
                    self.btn_dry.config(state=tk.NORMAL)
                    messagebox.showerror("拆分失败", item[1].splitlines()[-1])
                else:
                    self.txt.insert(tk.END, str(item) + "\n")
                self.txt.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(120, self._poll)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())

# -*- coding: utf-8 -*-
"""
xlsx_template.py —— "模板克隆式"输出：以原表为模板，只替换数据行

原理：
    xlsx 本质上是一个 zip 包。本模块把原文件的所有部件（styles.xml 样式表、
    sharedStrings.xml 共享字符串、theme 主题、列宽、行高、边框、数字格式、
    冻结窗格、自动筛选、图片……）**原样复制**，只重写 xl/worksheets/sheetN.xml
    里的 <sheetData>，因此生成的新表与原表模板格式完全一致。

    行数据不是重新绘制，而是把原表中该行的 XML 原文搬过来，只改行号，
    所以单元格的样式索引 s、数字格式、字体、边框、公式全部天然保留。

处理细节：
    * 只保留目标工作表，同步清理 workbook.xml / rels / Content_Types
    * 删除 calcChain.xml（公式计算链，行号变了必须删，否则 Excel 会报错）
    * 行号重映射：<c r="D12">、公式里的 D12、shared 公式的 ref 都会跟着改
    * 合并单元格、条件格式、数据验证、超链接的引用范围重映射，无法映射的丢弃
    * 自动筛选范围重设为新表的数据范围
"""

import re
import zipfile

# ---------------------------------------------------------------- 正则
RE_SHEETDATA_OPEN = re.compile(r"<sheetData(?:\s[^>]*)?>")
RE_SHEETDATA_CLOSE = "</sheetData>"
RE_ROW = re.compile(r"<row\b[^>]*?/>|<row\b[^>]*>.*?</row>", re.S)
RE_ROW_R = re.compile(r'(<row\b[^>]*?\br=")(\d+)(")')
RE_CELL_R = re.compile(r'r="([A-Za-z]{1,3})(\d+)"')
RE_FORMULA_REF = re.compile(r"(\$?[A-Za-z]{1,3}\$?)(\d{1,7})(?![\d(])")
RE_F_OPEN = re.compile(r"<f\b[^>]*>")
RE_F_SHARED_EMPTY = re.compile(r"<f\b[^>]*t=\"shared\"[^>]*/>")
RE_MERGE_CELL = re.compile(r"<mergeCell\b[^>]*/>")
RE_CF_BLOCK = re.compile(r"<conditionalFormatting\b[^>]*>.*?</conditionalFormatting>", re.S)
RE_DV_BLOCK = re.compile(r"<dataValidation\b[^>]*>.*?</dataValidation>|<dataValidation\b[^>]*/>", re.S)
RE_HYPERLINK = re.compile(r"<hyperlink\b[^>]*/>")
RE_XML_ATTR_REF = re.compile(r'\b(ref|sqref)="([^"]*)"')
RE_RANGE = re.compile(r"^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$")
RE_SHEET_ELEM = re.compile(r"<sheet\b[^>]*/>|<sheet\b[^>]*>.*?</sheet>", re.S)
RE_REL_ELEM = re.compile(r"<Relationship\b[^>]*/>|<Relationship\b[^>]*>.*?</Relationship>", re.S)
RE_OVERRIDE_ELEM = re.compile(r"<Override\b[^>]*/>|<Override\b[^>]*>.*?</Override>", re.S)
RE_DIMENSION = re.compile(r'<dimension\b[^>]*/>')
RE_ACTIVE_TAB = re.compile(r'(<workbookView\b[^>]*?\bactiveTab=")(\d+)(")')
RE_ATTR = re.compile(r'(\w+)="([^"]*)"')
RE_AUTOFILTER = re.compile(r'<autoFilter\b[^>]*/>')
RE_DEFINED_NAME = re.compile(r"<definedName\b[^>]*/>|<definedName\b[^>]*>.*?</definedName>", re.S)
RE_RANGE_END_ROW = re.compile(r"(\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$)(\d+)")


def _attrs(text):
    """取一个 XML 标签里的属性字典"""
    head = text[:text.find(">") + 1]
    return dict(RE_ATTR.findall(head))


def _norm_path(target, base="xl/"):
    t = (target or "").replace("\\", "/")
    if t.startswith("/"):
        return t.lstrip("/")
    return base + t.lstrip("./")


class TemplateSplitter(object):
    """把一个工作表按行号切片，输出为若干"同模板"的新文件"""

    def __init__(self, src_path, sheet=0):
        self.src_path = src_path
        with zipfile.ZipFile(src_path, "r") as z:
            self.items = dict((i.filename, z.read(i.filename)) for i in z.infolist())

        self.sheet_part, self.sheet_name = self._locate_sheet(sheet)
        self.xml = self.items[self.sheet_part].decode("utf-8", "replace")
        self.rows = self._parse_rows()          # [(row_id, row_xml_text)]
        self.row_text = dict(self.rows)
        self.head, self.tail = self._split_head_tail()

    # ---------------------------------------------------------- 定位
    def _locate_sheet(self, sheet):
        wb = self.items.get("xl/workbook.xml", b"").decode("utf-8", "replace")
        rels = self.items.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8", "replace")
        rel_map = {}
        for m in RE_REL_ELEM.finditer(rels):
            a = _attrs(m.group(0))
            if a.get("Id"):
                rel_map[a["Id"]] = _norm_path(a.get("Target", ""))
        sheets = []
        for m in RE_SHEET_ELEM.finditer(wb):
            a = _attrs(m.group(0))
            rid = a.get("r:id") or a.get("id")
            part = rel_map.get(rid)
            if part and part in self.items:
                sheets.append((a.get("name") or "", part, rid))
        if not sheets:
            raise RuntimeError("无法定位工作表")
        if isinstance(sheet, int):
            idx = sheet if 0 <= sheet < len(sheets) else 0
        else:
            idx = 0
            for i, (name, _, _) in enumerate(sheets):
                if name == sheet:
                    idx = i
                    break
        name, part, rid = sheets[idx]
        self.keep_rid = rid
        self.all_sheets = sheets
        return part, name

    # ---------------------------------------------------------- 解析
    def _split_head_tail(self):
        m = RE_SHEETDATA_OPEN.search(self.xml)
        if not m:
            return self.xml, ""
        open_tag = m.group(0)
        if open_tag.endswith("/>"):
            return self.xml[:m.start()], self.xml[m.end():]
        close = self.xml.find(RE_SHEETDATA_CLOSE, m.end())
        if close < 0:
            return self.xml[:m.start()], ""
        return self.xml[:m.start()], self.xml[close + len(RE_SHEETDATA_CLOSE):]

    def _parse_rows(self):
        m = RE_SHEETDATA_OPEN.search(self.xml)
        if not m:
            return []
        if m.group(0).endswith("/>"):
            return []
        inner_start = m.end()
        inner_end = self.xml.find(RE_SHEETDATA_CLOSE, inner_start)
        inner = self.xml[inner_start:inner_end if inner_end > 0 else len(self.xml)]
        out = []
        for rm in RE_ROW.finditer(inner):
            text = rm.group(0)
            a = _attrs(text)
            try:
                rid = int(a.get("r"))
            except (TypeError, ValueError):
                rid = len(out) + 1
            out.append((rid, text))
        return out

    # ---------------------------------------------------------- 行处理
    @staticmethod
    def _remap_formula(text, mapping):
        """重写公式/引用里的行号"""
        if "<f" not in text:
            return text

        def repl(m):
            key = int(m.group(2))
            new = mapping.get(key)
            if not new or new == key:
                return m.group(0)
            return "%s%d" % (m.group(1), new)

        # shared / array 公式转成普通公式，避免 ref 与 si 关联失效
        text = RE_F_SHARED_EMPTY.sub("", text)

        def fix_f(m):
            tag = m.group(0)
            a = _attrs(tag)
            if a.get("t") in ("shared", "array"):
                tag = re.sub(r'\s*t="(shared|array)"', "", tag)
                tag = re.sub(r'\s*si="\d+"', "", tag)
                tag = re.sub(r'\s*ref="[^"]*"', "", tag)
            return tag

        text = RE_F_OPEN.sub(fix_f, text)
        # 公式正文中的引用
        out = []
        pos = 0
        for fm in re.finditer(r"<f\b[^>]*>(.*?)</f>", text, re.S):
            out.append(text[pos:fm.start(1)])
            out.append(RE_FORMULA_REF.sub(repl, fm.group(1)))
            pos = fm.end(1)
        out.append(text[pos:])
        return "".join(out)

    def _row_to_new(self, old_id, new_id, mapping):
        text = self.row_text.get(old_id)
        if text is None:
            return '<row r="%d"/>' % new_id
        # 1) 单元格引用 / shared 公式 ref
        def cell_repl(m):
            key = int(m.group(2))
            new = mapping.get(key, key)
            return 'r="%s%d"' % (m.group(1), new)

        text = RE_CELL_R.sub(cell_repl, text)
        # 2) 公式
        text = self._remap_formula(text, mapping)
        # 3) 行号本身
        text = RE_ROW_R.sub(lambda m: '%s%d%s' % (m.group(1), new_id, m.group(3)), text, count=1)
        if 'r="' not in text[:80]:  # 极少数 row 没写 r 属性
            text = text.replace("<row", '<row r="%d"' % new_id, 1)
        return text

    # ---------------------------------------------------------- 尾部处理
    def _remap_range(self, ref, mapping):
        """把 A5:C9 这样的范围映射到新行号；无法完整映射则返回 None"""
        m = RE_RANGE.match(ref or "")
        if not m:
            return None
        r1 = int(m.group(2))
        r2 = int(m.group(4) or m.group(2))
        if r1 in mapping and r2 in mapping and mapping[r2] - mapping[r1] == r2 - r1:
            new1, new2 = mapping[r1], mapping[r2]
            if m.group(3):
                return "%s%d:%s%d" % (m.group(1), new1, m.group(3), new2)
            return "%s%d" % (m.group(1), new1)
        return None

    def _remap_tail(self, mapping, last_row, last_col_letter, filter_start=1):
        tail = self.tail

        # 合并单元格
        def merge_repl(m):
            a = _attrs(m.group(0))
            new = self._remap_range(a.get("ref", ""), mapping)
            if not new:
                return ""
            return '<mergeCell ref="%s"/>' % new

        def merge_block(m):
            block = m.group(0)
            inner = RE_MERGE_CELL.sub(merge_repl, block)
            if "<mergeCell" not in inner:
                return ""
            return inner

        tail = re.sub(r"<mergeCells\b[^>]*>.*?</mergeCells>|<mergeCells\b[^>]*/>",
                      merge_block, tail, flags=re.S)

        # 超链接
        def hyp_repl(m):
            a = _attrs(m.group(0))
            new = self._remap_range(a.get("ref", ""), mapping)
            if not new:
                return ""
            out = m.group(0)
            return re.sub(r'ref="[^"]*"', 'ref="%s"' % new, out, count=1)

        tail = RE_HYPERLINK.sub(hyp_repl, tail)

        # 条件格式 / 数据验证：范围能映射就改，不能映射就整块丢弃
        def block_repl(m):
            block = m.group(0)

            def ref_repl(rm):
                new = self._remap_range(rm.group(2), mapping)
                if not new:
                    return ""
                return '%s="%s"' % (rm.group(1), new)

            new_block = RE_XML_ATTR_REF.sub(ref_repl, block)
            if not RE_XML_ATTR_REF.search(new_block):
                return ""
            return new_block

        tail = RE_CF_BLOCK.sub(block_repl, tail)
        tail = RE_DV_BLOCK.sub(block_repl, tail)

        # 自动筛选：重设为整张新表的范围
        if RE_AUTOFILTER.search(tail):
            tail = RE_AUTOFILTER.sub(
                '<autoFilter ref="A%d:%s%d"/>' % (filter_start, last_col_letter, last_row), tail)
        return tail

    # ---------------------------------------------------------- 输出
    def build(self, out_path, data_row_ids, template_row_ids=None, sheet_name=None):
        """
        data_row_ids    : 需要写入的原表行号列表（升序）
        template_row_ids: 作为模板原样保留的行号（通常是标题行/表头行）
        """
        template_row_ids = list(template_row_ids or [])
        data_row_ids = [r for r in data_row_ids if r in self.row_text or True]
        if not data_row_ids and not template_row_ids:
            raise RuntimeError("没有可写入的行")

        # 行号映射：模板行保持原行号，数据行依次排在后面
        mapping = {}
        body = []
        next_row = 1
        for rid in template_row_ids:
            if rid in self.row_text:
                mapping[rid] = rid
                body.append((rid, self.row_text[rid]))
                next_row = max(next_row, rid + 1)
        for rid in data_row_ids:
            mapping[rid] = next_row
            body.append((next_row, self._row_to_new(rid, next_row, mapping)))
            next_row += 1
        last_row = next_row - 1

        # 列数
        max_col = 1
        for _, text in body:
            for cm in RE_CELL_R.finditer(text):
                col = cm.group(1).upper()
                n = 0
                for ch in col:
                    n = n * 26 + (ord(ch) - 64)
                max_col = max(max_col, n)
        last_col_letter = ""
        n = max_col
        while n:
            n, rem = divmod(n - 1, 26)
            last_col_letter = chr(65 + rem) + last_col_letter

        head = self.head
        if RE_DIMENSION.search(head):
            head = RE_DIMENSION.sub('<dimension ref="A1:%s%d"/>' % (last_col_letter, last_row), head, count=1)

        filter_start = max(template_row_ids) if template_row_ids else 1
        tail = self._remap_tail(mapping, last_row, last_col_letter or "A", filter_start)
        sheet_xml = (head + "<sheetData>" + "".join(t for _, t in body) + "</sheetData>" + tail)

        self._write_zip(out_path, sheet_xml, sheet_name, last_row)

    # ---------------------------------------------------------- 打包
    def _write_zip(self, out_path, sheet_xml, sheet_name=None, last_row=1):
        items = dict(self.items)
        items[self.sheet_part] = sheet_xml.encode("utf-8")

        # 删除其他工作表部件与计算链
        drop = set()
        rels = items.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8", "replace")
        for m in RE_REL_ELEM.finditer(rels):
            a = _attrs(m.group(0))
            rid = a.get("Id")
            typ = a.get("Type", "")
            target = _norm_path(a.get("Target", ""))
            if rid == self.keep_rid:
                continue
            if typ.endswith("/worksheet") or typ.endswith("/chartsheet"):
                drop.add(target)
            elif "calcChain" in target:
                drop.add(target)
        if "xl/calcChain.xml" in items:
            drop.add("xl/calcChain.xml")

        new_rels = []
        for m in RE_REL_ELEM.finditer(rels):
            a = _attrs(m.group(0))
            target = _norm_path(a.get("Target", ""))
            if target in drop:
                continue
            new_rels.append(m.group(0))
        if new_rels:
            items["xl/_rels/workbook.xml.rels"] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + "".join(new_rels) + "</Relationships>").encode("utf-8")

        # workbook.xml 只保留目标 sheet
        wb = items.get("xl/workbook.xml", b"").decode("utf-8", "replace")

        def sheet_repl(m):
            a = _attrs(m.group(0))
            rid = a.get("r:id") or a.get("id")
            if rid == self.keep_rid:
                if sheet_name and 'name="' in m.group(0):
                    return re.sub(r'name="[^"]*"', 'name="%s"' % sheet_name, m.group(0), count=1)
                return m.group(0)
            return ""

        wb = RE_SHEET_ELEM.sub(sheet_repl, wb)
        wb = RE_ACTIVE_TAB.sub(lambda m: "%s0%s" % (m.group(1), m.group(3)), wb)

        # 定义名称：指向已删除工作表的直接删掉；其余把引用范围的末行收敛到新表末行
        dropped_names = set(n for n, _p, r in self.all_sheets if r != self.keep_rid)

        def dn_repl(m):
            blk = m.group(0)
            for nm in dropped_names:
                if nm and ("'" + nm + "'" in blk or nm + "!" in blk):
                    return ""
            mm = RE_RANGE_END_ROW.search(blk)
            if mm and int(mm.group(2)) > last_row:
                return blk[:mm.start(2)] + str(last_row) + blk[mm.end(2):]
            return blk

        wb = RE_DEFINED_NAME.sub(dn_repl, wb)
        items["xl/workbook.xml"] = wb.encode("utf-8")

        # Content_Types
        ct = items.get("[Content_Types].xml", b"").decode("utf-8", "replace")
        new_ct = []
        for m in RE_OVERRIDE_ELEM.finditer(ct):
            a = _attrs(m.group(0))
            part = (a.get("PartName") or "").lstrip("/")
            if part in drop:
                continue
            new_ct.append(m.group(0))
        if new_ct:
            head = ct[:RE_OVERRIDE_ELEM.search(ct).start()]
            tail_start = RE_OVERRIDE_ELEM.finditer(ct)
            last = None
            for x in tail_start:
                last = x
            items["[Content_Types].xml"] = (head + "".join(new_ct) + ct[last.end():]).encode("utf-8")

        for d in drop:
            items.pop(d, None)

        # 写盘（保持原条目的压缩方式）
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in items.items():
                if isinstance(data, str):
                    data = data.encode("utf-8")
                z.writestr(name, data)
        return out_path

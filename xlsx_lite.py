# -*- coding: utf-8 -*-
"""
xlsx_lite.py —— 零依赖 .xlsx 读写模块（只使用 Python 标准库）

设计目标：
    统信 UOS / 内网环境往往装不上 pandas、openpyxl，本模块只用 zipfile + xml
    标准库直接读写 Office Open XML (xlsx)，拿来即用。

支持：
    读：文本 / 数字 / 布尔 / 日期 / 错误值 / 公式结果、多工作表、列宽、合并判断前的原始值
    写：文本 / 整数 / 浮点 / 布尔 / 日期时间、表头加粗、冻结首行、列宽、共享字符串去重

已知限制：
    * 不支持老式 .xls / WPS 专有的 .et 二进制格式（需先用 WPS 另存为 .xlsx）
    * 不保留公式本身（保留公式的计算结果值）、不保留图片/图表/合并单元格
"""

import datetime
import re
import zipfile
from xml.etree import ElementTree as ET

__all__ = [
    "col_letter_to_index", "index_to_col_letter", "split_cell_ref",
    "XlsxReader", "XlsxWriter", "SheetWriter", "is_xlsx_file",
]

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

NS = "{%s}" % MAIN_NS
R_NS = "{%s}" % REL_NS

MAX_ROWS = 1048576
MAX_COLS = 16384

# Excel 内建的数字格式 ID 中，属于日期/时间的部分
_BUILTIN_DATE_FMT = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))

_ILLEGAL_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")
_BAD_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")


# --------------------------------------------------------------------------
# 工具函数
# --------------------------------------------------------------------------
def col_letter_to_index(letters):
    """列字母 -> 1 开始的列序号。  A->1, Z->26, AA->27"""
    letters = (letters or "").strip().upper()
    n = 0
    for ch in letters:
        if not ("A" <= ch <= "Z"):
            raise ValueError("非法列字母: %r" % letters)
        n = n * 26 + (ord(ch) - 64)
    if n <= 0:
        raise ValueError("非法列字母: %r" % letters)
    return n


def index_to_col_letter(index):
    """1 开始的列序号 -> 列字母。  1->A, 27->AA"""
    if index < 1 or index > MAX_COLS:
        raise ValueError("列序号超出范围: %s" % index)
    s = ""
    while index:
        index, rem = divmod(index - 1, 26)
        s = chr(65 + rem) + s
    return s


def split_cell_ref(ref):
    """'AB12' -> (28, 12)"""
    m = re.match(r"([A-Za-z]+)(\d+)", ref or "")
    if not m:
        return None, None
    return col_letter_to_index(m.group(1)), int(m.group(2))


def _escape(text):
    """XML 文本转义 + 剔除非法控制字符"""
    if text is None:
        return ""
    text = _ILLEGAL_XML.sub("", str(text))
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _safe_sheet_name(name, used=None):
    """工作表名合法化：去非法字符，最长 31 字符，避免重名"""
    name = _BAD_SHEET_CHARS.sub("_", (name or "Sheet").strip()) or "Sheet"
    name = name[:31]
    if used is not None:
        base = name
        i = 2
        while name in used:
            suffix = "_%d" % i
            name = base[:31 - len(suffix)] + suffix
            i += 1
        used.add(name)
    return name


def is_xlsx_file(path):
    """通过文件头判断是否为真正的 xlsx（zip 包）"""
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except OSError:
        return False
    if head[:4] == b"PK\x03\x04":
        return True
    return False


def detect_old_binary(path):
    """返回 True 表示这是 .xls / .et 之类的老二进制格式"""
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except OSError:
        return False
    return head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


# --------------------------------------------------------------------------
# 日期 <-> 序列号
# --------------------------------------------------------------------------
def datetime_to_serial(dt):
    """datetime/date -> Excel 序列号"""
    if isinstance(dt, datetime.datetime):
        base = datetime.datetime(1899, 12, 31)
        delta = dt - base
    elif isinstance(dt, datetime.date):
        base = datetime.date(1899, 12, 31)
        delta = dt - base
    else:
        return None
    serial = delta.days + delta.seconds / 86400.0 + getattr(delta, "microseconds", 0) / 86400.0 / 1e6
    # 兼容 Excel 的 1900 闰年 bug
    if serial >= 60:
        serial += 1
    return serial


def serial_to_datetime(serial):
    """Excel 序列号 -> datetime / date / time"""
    try:
        serial = float(serial)
    except (TypeError, ValueError):
        return None
    if serial < 0:
        return None
    # 1900 闰年 bug：60 表示不存在的 1900-02-29
    if serial >= 60:
        serial -= 1
    if serial < 1:  # 纯时间
        secs = int(round(serial * 86400))
        if secs >= 86400:
            secs -= 86400
        return datetime.time(secs // 3600, (secs % 3600) // 60, secs % 60)
    days = int(serial)
    frac = serial - days
    try:
        d = datetime.datetime(1899, 12, 31) + datetime.timedelta(days=days, seconds=frac * 86400)
    except OverflowError:
        return None
    # 消除浮点误差造成的 xx:xx:59.999999
    if frac == 0:
        return d.date()
    rounded = d + datetime.timedelta(microseconds=500000)
    d = rounded.replace(microsecond=0)
    if d.hour == 0 and d.minute == 0 and d.second == 0:
        return d.date()
    return d


# --------------------------------------------------------------------------
# 读取
# --------------------------------------------------------------------------
class XlsxReader(object):
    """流式读取 xlsx"""

    def __init__(self, path):
        self.path = path
        self._zip = zipfile.ZipFile(path, "r")
        self._names = set(self._zip.namelist())
        self.shared_strings = self._read_shared_strings()
        self.date_styles = self._read_date_styles()
        self.sheets = self._read_sheet_index()  # [(name, zip_path)]

    # -- 基础 -------------------------------------------------------------
    def close(self):
        try:
            self._zip.close()
        except Exception:
            pass

    def __enter__(self):
        return self

    def __exit__(self, *a):
        self.close()

    @property
    def sheet_names(self):
        return [n for n, _ in self.sheets]

    def _read_shared_strings(self):
        if "xl/sharedStrings.xml" not in self._names:
            return []
        out = []
        with self._zip.open("xl/sharedStrings.xml") as f:
            for _, si in ET.iterparse(f, events=("end",)):
                if si.tag == NS + "si":
                    out.append(self._si_text(si))
                    si.clear()
        return out

    @staticmethod
    def _si_text(si):
        """取 <si> 的纯文本，跳过日文注音 rPh"""
        parents = {}
        for p in si.iter():
            for c in p:
                parents[c] = p
        parts = []
        for node in si.iter():
            if node.tag != NS + "t":
                continue
            anc = parents.get(node)
            skip = False
            while anc is not None:
                if anc.tag == NS + "rPh":
                    skip = True
                    break
                anc = parents.get(anc)
            if not skip:
                parts.append(node.text or "")
        return "".join(parts)

    def _read_date_styles(self):
        """返回 list[bool]：cellXfs 中第 i 个样式是否为日期格式"""
        flags = []
        if "xl/styles.xml" not in self._names:
            return flags
        custom = {}
        with self._zip.open("xl/styles.xml") as f:
            root = ET.parse(f).getroot()
        for num_fmt in root.findall(NS + "numFmts/" + NS + "numFmt"):
            try:
                fmt_id = int(num_fmt.get("numFmtId"))
            except (TypeError, ValueError):
                continue
            custom[fmt_id] = num_fmt.get("formatCode") or ""
        cell_xfs = root.find(NS + "cellXfs")
        if cell_xfs is None:
            return flags
        for xf in cell_xfs.findall(NS + "xf"):
            try:
                fmt_id = int(xf.get("numFmtId", "0"))
            except ValueError:
                fmt_id = 0
            if fmt_id in _BUILTIN_DATE_FMT:
                flags.append(True)
            elif fmt_id in custom:
                flags.append(self._is_date_code(custom[fmt_id]))
            else:
                flags.append(False)
        return flags

    @staticmethod
    def _is_date_code(code):
        code = re.sub(r"\[[^\]]*\]", "", code or "")      # 去掉 [Red]、[$-409] 等
        code = re.sub(r'"[^"]*"', "", code)               # 去掉引号内文本
        code = code.replace("\\", "")
        return bool(re.search(r"[ymd]", code, re.IGNORECASE)) and bool(
            re.search(r"[yMd]", code)
        )

    def _read_sheet_index(self):
        rels = {}
        if "xl/_rels/workbook.xml.rels" in self._names:
            with self._zip.open("xl/_rels/workbook.xml.rels") as f:
                root = ET.parse(f).getroot()
            for rel in root:
                target = rel.get("Target", "")
                if target.startswith("/"):
                    target = target[1:]
                elif not target.startswith("xl/"):
                    target = "xl/" + target.lstrip("./")
                rels[rel.get("Id")] = target
        sheets = []
        with self._zip.open("xl/workbook.xml") as f:
            root = ET.parse(f).getroot()
        container = root.find(NS + "sheets")
        if container is None:
            return sheets
        for sh in container.findall(NS + "sheet"):
            name = sh.get("name") or "Sheet"
            rid = sh.get(R_NS + "id")
            path = rels.get(rid)
            if path and path in self._names:
                sheets.append((name, path))
        if not sheets:  # 兜底：直接找 worksheet
            for n in sorted(self._names):
                if n.startswith("xl/worksheets/") and n.endswith(".xml"):
                    sheets.append((n.split("/")[-1][:-4], n))
        return sheets

    # -- 行迭代 -----------------------------------------------------------
    def _sheet_path(self, sheet):
        if isinstance(sheet, int):
            if sheet < 0 or sheet >= len(self.sheets):
                raise IndexError("工作表序号 %s 不存在（共 %d 个）" % (sheet, len(self.sheets)))
            return self.sheets[sheet][1]
        for name, path in self.sheets:
            if name == sheet:
                return path
        raise KeyError("找不到工作表：%s（可选：%s）" % (sheet, "、".join(self.sheet_names)))

    def iter_rows(self, sheet=0, with_index=False):
        """逐行 yield list；with_index=True 时 yield (原始行号, list)"""
        path = self._sheet_path(sheet)
        seq = 0
        with self._zip.open(path) as f:
            context = ET.iterparse(f, events=("end",))
            root = getattr(context, "root", None)
            for _, elem in context:
                if elem.tag != NS + "row":
                    continue
                row = self._parse_row(elem)
                if with_index:
                    try:
                        no = int(elem.get("r"))
                    except (TypeError, ValueError):
                        no = seq + 1
                    seq += 1
                    yield no, row
                else:
                    yield row
                if root is not None:
                    root.clear()
                else:  # pragma: no cover
                    elem.clear()

    def _parse_row(self, row_elem):
        cells = {}
        for c in row_elem.findall(NS + "c"):
            col_idx, _ = split_cell_ref(c.get("r"))
            if col_idx is None:
                col_idx = len(cells) + 1
            cells[col_idx] = self._cell_value(c)
        if not cells:
            return []
        width = max(cells)
        return [cells.get(i) for i in range(1, width + 1)]

    def _cell_value(self, c):
        ctype = c.get("t") or "n"
        if ctype == "inlineStr":
            node = c.find(NS + "is")
            return self._collect_text(node) if node is not None else None
        v = c.find(NS + "v")
        if v is None or v.text is None:
            return None
        raw = v.text
        if ctype == "s":
            try:
                return self.shared_strings[int(raw)]
            except (ValueError, IndexError):
                return raw
        if ctype == "str":
            return raw
        if ctype == "b":
            return raw == "1" or raw.lower() == "true"
        if ctype == "e":
            return raw
        # 数字（可能带日期格式）
        try:
            style_idx = int(c.get("s") or 0)
        except ValueError:
            style_idx = 0
        if style_idx < len(self.date_styles) and self.date_styles[style_idx]:
            dt = serial_to_datetime(raw)
            if dt is not None:
                return dt
        try:
            num = float(raw)
        except ValueError:
            return raw
        if num.is_integer() and abs(num) < 1e15 and ("." not in raw and "e" not in raw.lower()):
            return int(num)
        return num

    @classmethod
    def _collect_text(cls, node):
        parents = {}
        for p in node.iter():
            for c in p:
                parents[c] = p
        parts = []
        for n in node.iter():
            if n.tag == NS + "t":
                anc = parents.get(n)
                skip = False
                while anc is not None:
                    if anc.tag == NS + "rPh":
                        skip = True
                        break
                    anc = parents.get(anc)
                if not skip:
                    parts.append(n.text or "")
        return "".join(parts)

    def read_cols_width(self, sheet=0):
        """读取列宽 {列序号: 宽度}"""
        path = self._sheet_path(sheet)
        widths = {}
        with self._zip.open(path) as f:
            for _, elem in ET.iterparse(f, events=("end",)):
                if elem.tag == NS + "col":
                    try:
                        w = float(elem.get("width"))
                    except (TypeError, ValueError):
                        continue
                    try:
                        lo = int(elem.get("min"))
                        hi = int(elem.get("max"))
                    except (TypeError, ValueError):
                        continue
                    for i in range(lo, hi + 1):
                        widths[i] = w
                elif elem.tag == NS + "sheetData":
                    break
        return widths

    def read_all(self, sheet=0):
        return list(self.iter_rows(sheet))


# --------------------------------------------------------------------------
# 写入
# --------------------------------------------------------------------------
class SheetWriter(object):
    """一张工作表的写入器（行数据暂存于内存，save 时统一落盘）"""

    def __init__(self, name, book):
        self.name = name
        self._book = book
        self._rows = []
        self._row_no = 0
        self._widths = {}
        self.freeze_header = True

    @property
    def row_count(self):
        return self._row_no

    def set_col_width(self, col_index, width):
        if width and width > 0:
            self._widths[col_index] = min(float(width), 255.0)

    def append_row(self, values, style=0):
        """追加一行；values 为 list。style: 0 常规 / 1 表头加粗"""
        self._row_no += 1
        if self._row_no > MAX_ROWS:
            raise RuntimeError("工作表行数超过 Excel 上限 1048576 行")
        if values is None:
            values = []
        row_no = self._row_no
        parts = []
        for i, val in enumerate(values, start=1):
            if val is None or val == "":
                if style:
                    parts.append('<c r="%s%d" s="%d"/>' % (index_to_col_letter(i), row_no, style))
                continue
            ref = "%s%d" % (index_to_col_letter(i), row_no)
            parts.append(self._cell_xml(ref, val, style))
        if not parts and not style:
            self._rows.append('<row r="%d"/>' % row_no)
        else:
            self._rows.append('<row r="%d">%s</row>' % (row_no, "".join(parts)))

    def _cell_xml(self, ref, val, style):
        s_attr = ' s="%d"' % style if style else ""
        if isinstance(val, bool):
            return '<c r="%s"%s t="b"><v>%d</v></c>' % (ref, s_attr, 1 if val else 0)
        if isinstance(val, (int, float)):
            if isinstance(val, float) and (val != val or val in (float("inf"), float("-inf"))):
                return '<c r="%s"%s t="e"><v>#NUM!</v></c>' % (ref, s_attr)
            return '<c r="%s"%s><v>%s</v></c>' % (ref, s_attr, repr(val) if isinstance(val, float) else val)
        if isinstance(val, (datetime.datetime, datetime.date, datetime.time)):
            serial = datetime_to_serial(val)
            if serial is None:  # 纯 time
                serial = (val.hour * 3600 + val.minute * 60 + val.second) / 86400.0
            date_style = self._book._style_of_date(val)
            txt = "%.10f" % serial
            txt = txt.rstrip("0").rstrip(".") or "0"
            return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, date_style, txt)
        # 文本
        idx = self._book._sst_index(val)
        return '<c r="%s"%s t="s"><v>%d</v></c>' % (ref, s_attr, idx)

    def to_xml(self):
        buf = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<worksheet xmlns="%s" xmlns:r="%s">' % (MAIN_NS, REL_NS),
               '<dimension ref="A1"/>']
        buf.append('<sheetViews><sheetView workbookViewId="0">')
        if self.freeze_header and self._row_no > 1:
            buf.append('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                       '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>')
        buf.append('</sheetView></sheetViews>')
        buf.append('<sheetFormatPr defaultRowHeight="15"/>')
        if self._widths:
            buf.append("<cols>")
            for i in sorted(self._widths):
                buf.append('<col min="%d" max="%d" width="%.2f" customWidth="1"/>' % (i, i, self._widths[i]))
            buf.append("</cols>")
        buf.append("<sheetData>")
        buf.extend(self._rows)
        buf.append("</sheetData>")
        buf.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        buf.append("</worksheet>")
        return "".join(buf)


class XlsxWriter(object):
    """生成一个 xlsx 文件"""

    def __init__(self):
        self._sheets = []
        self._sst = {}
        self._sst_order = []

    # -- 共享字符串 --------------------------------------------------------
    def _sst_index(self, text):
        text = "" if text is None else str(text)
        idx = self._sst.get(text)
        if idx is None:
            idx = len(self._sst_order)
            self._sst[text] = idx
            self._sst_order.append(text)
        return idx

    @staticmethod
    def _style_of_date(val):
        if isinstance(val, datetime.datetime) and (val.hour or val.minute or val.second):
            return STYLE_DATETIME
        return STYLE_DATE

    # -- 工作表 ------------------------------------------------------------
    def add_sheet(self, name):
        name = _safe_sheet_name(name, set(s.name for s in self._sheets))
        sh = SheetWriter(name, self)
        self._sheets.append(sh)
        return sh

    # -- 落盘 --------------------------------------------------------------
    def save(self, path):
        sheets = self._sheets or [self.add_sheet("Sheet1")]
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", self._content_types(len(sheets)))
            z.writestr("_rels/.rels", _ROOT_RELS)
            z.writestr("xl/workbook.xml", self._workbook_xml(sheets))
            z.writestr("xl/_rels/workbook.xml.rels", self._workbook_rels(len(sheets)))
            z.writestr("xl/styles.xml", _STYLES_XML)
            for i, sh in enumerate(sheets, start=1):
                z.writestr("xl/worksheets/sheet%d.xml" % i, sh.to_xml())
            z.writestr("xl/sharedStrings.xml", self._sst_xml())
            z.writestr("docProps/app.xml", _APP_XML)
            z.writestr("docProps/core.xml", _core_xml())
        return path

    # -- 各部件 XML --------------------------------------------------------
    @staticmethod
    def _content_types(n_sheets):
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                 '<Types xmlns="%s">' % CT_NS,
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                 '<Default Extension="xml" ContentType="application/xml"/>',
                 '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
                 '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
                 '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
                 '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
                 '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>']
        for i in range(1, n_sheets + 1):
            parts.append('<Override PartName="/xl/worksheets/sheet%d.xml" '
                         'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
        parts.append("</Types>")
        return "".join(parts)

    @staticmethod
    def _workbook_xml(sheets):
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                 '<workbook xmlns="%s" xmlns:r="%s">' % (MAIN_NS, REL_NS),
                 "<sheets>"]
        for i, sh in enumerate(sheets, start=1):
            parts.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (_escape(sh.name), i, i))
        parts.append("</sheets></workbook>")
        return "".join(parts)

    @staticmethod
    def _workbook_rels(n_sheets):
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                 '<Relationships xmlns="%s">' % PKG_REL_NS]
        for i in range(1, n_sheets + 1):
            parts.append('<Relationship Id="rId%d" Type="%s/worksheet" Target="worksheets/sheet%d.xml"/>'
                         % (i, REL_NS, i))
        parts.append('<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/>' % (n_sheets + 1, REL_NS))
        parts.append('<Relationship Id="rId%d" Type="%s/sharedStrings" Target="sharedStrings.xml"/>' % (n_sheets + 2, REL_NS))
        parts.append("</Relationships>")
        return "".join(parts)

    def _sst_xml(self):
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                 '<sst xmlns="%s" count="%d" uniqueCount="%d">' % (MAIN_NS, len(self._sst_order), len(self._sst_order))]
        for s in self._sst_order:
            if s != s.strip():
                parts.append('<si><t xml:space="preserve">%s</t></si>' % _escape(s))
            else:
                parts.append("<si><t>%s</t></si>" % _escape(s))
        parts.append("</sst>")
        return "".join(parts)


_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="%s">'
    '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/'
    'relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="%s/extended-properties" Target="docProps/app.xml"/>'
    '</Relationships>'
) % (PKG_REL_NS, REL_NS, REL_NS)

_APP_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    "<Application>WPS 表格拆分工具</Application></Properties>"
)


def _core_xml():
    ts = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:creator>WPS 表格拆分工具</dc:creator>"
        "<cp:lastModifiedBy>WPS 表格拆分工具</cp:lastModifiedBy>"
        '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
        "</cp:coreProperties>" % (ts, ts)
    )


# 样式索引（必须与下方 _STYLES_XML 中 cellXfs 的顺序严格一致）
STYLE_GENERAL = 0
STYLE_BOLD = 1
STYLE_DATE = 2
STYLE_DATETIME = 3

_STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="%s">'
    '<numFmts count="2">'
    '<numFmt numFmtId="176" formatCode="yyyy\\-mm\\-dd"/>'
    '<numFmt numFmtId="177" formatCode="yyyy\\-mm\\-dd\\ hh:mm"/>'
    "</numFmts>"
    '<fonts count="2">'
    '<font><sz val="11"/><color theme="1"/><name val="宋体"/><charset val="134"/></font>'
    '<font><b/><sz val="11"/><color theme="1"/><name val="宋体"/><charset val="134"/></font>'
    "</fonts>"
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    "<cellXfs count=\"4\">"
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="176" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="177" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    "</cellXfs>"
    '<cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>"
) % MAIN_NS

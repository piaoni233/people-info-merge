#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多表人员信息汇总助手 · 合并引擎 (people_merge.py)

纯 Python 标准库实现（零第三方依赖即可运行），可选使用 openpyxl 加速/兜底，
可选使用 xlrd 读取旧版二进制 .xls。

支持的输入格式（一律按文件内容判定，扩展名只作提示）：
  * .xlsx / .xlsm —— OOXML，标准库或 openpyxl 读取
  * .ods          —— OpenDocument，标准库（zipfile + ElementTree）读取
  * .xls          —— 旧版二进制（BIFF/OLE）需可选 xlrd；同时自动识别"其实是
                     HTML 表格 / Excel 2003 XML / 分隔符文本"的伪 .xls
  * .csv/.tsv/.txt —— 自动嗅探分隔符（, ; Tab |）与编码（UTF-8/GB18030…）
输出统一为 .xlsx（三个 sheet）。

设计遵循同目录 SKILL.md 与 references/ 中的约定：

  * inspect      只读每张表的表头 + N 行样本，输出 profile.json（不把全量数据卷入流程）
  * suggest-keys 逐对文件逐列打分，输出候选关联键 + 人类可读证据（供用户确认）
  * merge        按确认后的 plan.json 执行合并，输出「汇总结果 + 合并日志 + 未匹配记录」
  * selftest     端到端自测（全部在临时目录完成，不污染用户目录）
  * selfcheck    自证：扫描自身源码，确认无网络模块、无外部命令调用

安全声明：本文件只 import 标准库（+ 可选的 openpyxl / xlrd），不联网、
不执行外部命令，仅读写用户在命令行中明确指定的本地文件。
"""
from __future__ import annotations

import argparse
import ast
import csv
import datetime as _dt
import difflib
import io
import json
import os
import re
import sys
import tempfile
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from html.parser import HTMLParser
from xml.sax.saxutils import escape as _xml_escape

VERSION = "1.0.0"
TOOL_NAME = "people_merge"
MAX_FILE_BYTES = 200 * 1024 * 1024   # 单文件 200MB 上限
MAX_ZIP_XML_BYTES = 64 * 1024 * 1024  # 压缩包内单个 XML 解压后上限
MAX_CELL_REPEAT = 64                  # 单元格"重复列/行"展开上限（防稀疏文件膨胀）
MAX_ROWS = 200000                    # 单表 20 万行上限
ENCODINGS = ("utf-8-sig", "utf-8", "gb18030", "gbk", "big5")
ALLOWED_EXTS = (".xlsx", ".xlsm", ".xls", ".ods", ".csv", ".tsv", ".txt")
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"   # 旧版 .xls 的 OLE 复合文档头
ZIP_MAGIC = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")

try:  # Windows 控制台中文输出保护
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import openpyxl as _oxl  # type: ignore
    from openpyxl.styles import Font as _OxlFont  # type: ignore
    HAVE_OPENPYXL = True
    OPENPYXL_VERSION = str(getattr(_oxl, "__version__", "unknown"))
except Exception:
    _oxl = None
    _OxlFont = None
    HAVE_OPENPYXL = False
    OPENPYXL_VERSION = None

try:  # 可选：仅用于读取旧版二进制 .xls（xlrd 2.x 只支持 .xls，BSD 许可）
    import xlrd as _xlrd  # type: ignore
    HAVE_XLRD = True
    XLRD_VERSION = str(getattr(_xlrd, "__version__", "unknown"))
except Exception:
    _xlrd = None
    HAVE_XLRD = False
    XLRD_VERSION = None


class MergeError(Exception):
    """用户可理解的确定性错误（文件缺失、格式不支持、plan 缺键等）。"""


# ---------------- 基础：列名猜测与关联键归一化 ----------------

NAME_HINTS = ("姓名", "名字", "员工姓名", "人员姓名", "花名", "员工", "人员", "name")
ID_HINTS = ("工号", "员工号", "员工编号", "编号", "工卡", "备案号", "id")
CONTACT_HINTS = ("手机", "电话", "联系", "邮箱", "微信", "身份证")

# 同义逻辑字段归一化：不同表对同一语义用了不同列名时仍能识别为"同一逻辑字段"
#（冲突检测用；合并输出列仍保留原始列名 + 来源标注，不混淆出处）
FIELD_SYNONYMS = [
    (("手机", "手机号", "移动电话", "电话", "联系电话", "联系方式", "联络方式", "通讯方式"), "联系方式"),
    (("部门", "所属部门", "所在部门"), "部门"),
    (("姓名", "员工姓名", "人员姓名", "名字"), "姓名"),
    (("工号", "员工号", "员工编号", "编号", "工卡"), "工号"),
]


def logic_name(col: str) -> str:
    """列名 -> 逻辑字段名：去来源标注/括号备注，按同义词表归一。"""
    base = re.sub(r"[（(].*[）)]", "", col)
    base = re.sub(r"^(员工|人员)", "", base)
    base = squeeze(base).casefold()
    for words, canon in FIELD_SYNONYMS:
        for w in words:
            if w in col or w.casefold() in base:
                return canon
    return base

_ID_PREFIX_RE = re.compile(r"^(EMP|E|NO\.?|工号|员工号|编号|ID|工卡)\s*[:：\-_#]?\s*", re.IGNORECASE)
_ID_LIKE_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-_]*$")
_TRAIL_PAREN_RE = re.compile(r"[（(][^）)]{1,30}[）)]\s*$")
_ALIAS_SPLIT_RE = re.compile(r"[、/，,;；|]")
_ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_CELL_REF_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


def clean_xml(s: str) -> str:
    return _ILLEGAL_XML_RE.sub("", s)


def norm_text(v) -> str:
    if v is None:
        return ""
    return unicodedata.normalize("NFKC", str(v)).strip()


def squeeze(s) -> str:
    """NFKC + 删除全部空白（含全角空格 U+3000，NFKC 后变为普通空格）。"""
    return "".join(unicodedata.normalize("NFKC", str(s)).split())


def norm_key(value, kind: str = "auto") -> str:
    """关联键归一化：去空白/全角差异、去姓名备注后缀、工号类去前缀并小写。"""
    s = squeeze(value)
    s = _TRAIL_PAREN_RE.sub("", s)
    kl = kind
    if kl == "auto":
        kl = "id" if _ID_LIKE_RE.match(s or "") else "name"
    if kl == "id":
        s = _ID_PREFIX_RE.sub("", s)
        s = s.casefold()
    return s


def split_aliases(value) -> list:
    return [p.strip() for p in _ALIAS_SPLIT_RE.split(norm_text(value)) if p.strip()]


def guess_col_kind(header: str) -> str:
    h = norm_text(header)
    hl = h.lower()
    for hint in ID_HINTS:
        if hint.lower() in hl:
            return "id"
    for hint in NAME_HINTS:
        if hint.lower() in hl:
            return "name"
    for hint in CONTACT_HINTS:
        if hint in h:
            return "contact"
    return "other"


def col_letters(idx: int) -> str:
    n = idx + 1
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def now_str() -> str:
    return _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def short_name(path: str) -> str:
    return os.path.splitext(os.path.basename(path))[0]


# ---------------- 文件读取：格式嗅探 + 各格式解析器 ----------------


def _decode_text(raw: bytes, path: str) -> str:
    for enc in ENCODINGS:
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    raise MergeError("无法识别文本编码（已尝试 %s）：%s" % ("/".join(ENCODINGS), path))


def _num_text(v) -> str:
    """数字单元格文本化：13800000001.0 -> 13800000001（避免多余的 .0）。"""
    s = norm_text(v)
    if re.fullmatch(r"-?\d+\.0+", s):
        return s.split(".")[0]
    if re.fullmatch(r"-?\d+(?:\.\d+)?[eE][+-]?\d+", s):
        try:
            f = float(s)
        except ValueError:
            return s
        if f.is_integer():
            return str(int(f))
    return s


def _safe_xml(data: bytes, what: str):
    """统一的 XML 解析入口：拒绝 DTD/实体定义（防实体膨胀），并给出中文错误。"""
    head = data[:4096]
    if b"<!ENTITY" in head or b"<!DOCTYPE" in head:
        raise MergeError("%s 含 XML DTD/实体定义，出于安全考虑拒绝解析。" % what)
    try:
        return ET.fromstring(data)
    except ET.ParseError as e:
        raise MergeError("%s 解析失败：%s" % (what, e))


def sniff_format(path: str) -> str:
    """按文件内容判定格式：xlsx / ods / xls_ole / html / spreadsheetml / delimited。"""
    ext = os.path.splitext(path)[1].lower()
    if ext not in ALLOWED_EXTS:
        raise MergeError(
            "不支持的文件类型 %s：仅支持 %s（若文件其实是表格，请改回正确扩展名后重试）。"
            % (ext or "（无扩展名）", " / ".join(ALLOWED_EXTS)))
    with open(path, "rb") as f:
        head = f.read(8192)
    if head[:4] in ZIP_MAGIC:
        try:
            with zipfile.ZipFile(path) as zf:
                names = set(zf.namelist())
        except zipfile.BadZipFile:
            raise MergeError("压缩包已损坏或不是有效的表格文件：%s" % path)
        if "xl/workbook.xml" in names:
            return "xlsx"
        if "content.xml" in names:
            return "ods"
        raise MergeError("压缩包内未找到 xl/workbook.xml 或 content.xml，"
                         "不是支持的表格格式（.xlsx/.xlsm/.ods）：%s" % path)
    if head[:8] == OLE_MAGIC:
        return "xls_ole"
    probe = _decode_text(head, path).lstrip("\ufeff \t\r\n").lower()
    if "urn:schemas-microsoft-com:office:spreadsheet" in probe:
        return "spreadsheetml"          # Excel 2003 XML（常见的"伪 .xls"）
    if probe.startswith("<!doctype html") or probe.startswith("<html") or "<table" in probe[:4096]:
        return "html"                   # 系统网页导出（另一种常见"伪 .xls"）
    return "delimited"


def _read_delimited(path: str):
    """CSV / TSV / TXT：自动嗅探分隔符与编码。"""
    with open(path, "rb") as f:
        raw = f.read()
    if len(raw) > MAX_FILE_BYTES:
        raise MergeError("文件过大（>%dMB），拒绝读取：%s" % (MAX_FILE_BYTES // 1048576, path))
    text = _decode_text(raw, path)
    used = None
    for enc in ENCODINGS:
        try:
            raw.decode(enc)
            used = enc
            break
        except (UnicodeDecodeError, LookupError):
            continue
    try:
        dialect = csv.Sniffer().sniff(text.lstrip("\ufeff")[:4096], delimiters=[",", ";", "\t", "|"])
        delim = dialect.delimiter
    except Exception:
        delim = "\t" if os.path.splitext(path)[1].lower() == ".tsv" else ","
    reader = csv.reader(io.StringIO(text), delimiter=delim)
    rows = [[norm_text(c) for c in row] for row in reader]
    rows = [r for r in rows if any(c != "" for c in r)]
    return rows, {"encoding": used or "utf-8", "delimiter": delim}


class _HtmlTableParser(HTMLParser):
    """只取第一张（不含嵌套的）<table>；支持 colspan、<br> 视作空格，忽略脚本/样式。"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows: list = []
        self.n_tables = 0
        self._depth = 0        # 当前 table 嵌套深度
        self._taken = False    # 第一张表是否已采集完毕（后续表一律忽略）
        self._skip = 0
        self._row = None
        self._cell = None
        self._buf: list = []

    def _flush_cell(self):
        if self._cell is None or self._row is None:
            return
        self._row.append(squeeze("".join(self._buf)))
        try:
            span = max(1, int(self._cell.get("colspan", "1") or "1"))
        except ValueError:
            span = 1
        self._row.extend([""] * (min(span, MAX_CELL_REPEAT) - 1))
        self._cell = None
        self._buf = []

    def _collect(self) -> bool:
        return self._depth == 1 and not self._taken

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self._skip += 1
            return
        if tag == "table":
            self.n_tables += 1
            if self._depth == 0 and not self._taken:
                self._depth = 1
                self.rows = []
            else:
                self._depth += 1
            return
        if not self._collect():
            return
        if tag == "tr":
            self._flush_cell()
            self._row = []
        elif tag in ("td", "th"):
            self._flush_cell()
            self._cell = {k.lower(): (v or "") for k, v in attrs}
        elif tag == "br" and self._cell is not None:
            self._buf.append(" ")

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        self.handle_endtag(tag)

    def handle_data(self, data):
        if self._collect() and self._cell is not None and not self._skip:
            self._buf.append(data)

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self._skip = max(0, self._skip - 1)
            return
        if tag == "table":
            if self._depth == 1:
                self._taken = True
                self._flush_cell()
            self._depth = max(0, self._depth - 1)
            return
        if not self._collect():
            return
        if tag in ("td", "th"):
            self._flush_cell()
        elif tag == "tr":
            self._flush_cell()
            if self._row:
                self.rows.append(self._row)
            self._row = None

    def close(self):
        super().close()
        self._flush_cell()
        if self._row:                      # 容错：未闭合的 </tr>
            self.rows.append(self._row)
            self._row = None


def _read_html(path: str):
    """HTML 表格（含被改名成 .xls 的网页导出）。找不到表格时返回 None 由调用方回落。"""
    size = os.path.getsize(path)
    if size > MAX_FILE_BYTES:
        raise MergeError("文件过大（>%dMB），拒绝读取：%s" % (MAX_FILE_BYTES // 1048576, path))
    with open(path, "rb") as f:
        text = _decode_text(f.read(), path)
    parser = _HtmlTableParser()
    parser.feed(text)
    parser.close()
    if not parser.rows:
        return None, {"tables": parser.n_tables}
    return parser.rows, {"tables": parser.n_tables}


_SML_NS = "urn:schemas-microsoft-com:office:spreadsheet"


def _read_spreadsheetml(path: str):
    """Excel 2003 XML（SpreadsheetML）。找不到 Worksheet 时返回 None 由调用方回落。"""
    size = os.path.getsize(path)
    if size > MAX_FILE_BYTES:
        raise MergeError("文件过大（>%dMB），拒绝读取：%s" % (MAX_FILE_BYTES // 1048576, path))
    with open(path, "rb") as f:
        raw = f.read()
    root = _safe_xml(_decode_text(raw, path).encode("utf-8"), "Excel 2003 XML")

    def q(tag: str) -> str:
        return "{%s}%s" % (_SML_NS, tag)

    ws = root.find(".//" + q("Worksheet"))
    table = None if ws is None else ws.find(q("Table"))
    if table is None:
        return None, {"sheet": None}
    rows = []
    for tr in table.findall(q("Row")):
        cells: list = []
        idx = 0
        for tc in tr.findall(q("Cell")):
            try:
                idx = int(tc.get(q("Index")) or (idx + 1)) - 1
            except ValueError:
                idx = len(cells)
            if idx > len(cells):
                cells.extend([""] * min(idx - len(cells), MAX_CELL_REPEAT * 16))
            data = tc.find(q("Data"))
            val = "" if data is None else norm_text("".join(data.itertext()))
            if data is not None and (data.get(q("Type")) or "").lower() == "number":
                val = _num_text(val)
            cells.append(val)
            idx += 1
            try:
                span = int(tc.get(q("MergeAcross")) or 0)
            except ValueError:
                span = 0
            cells.extend([""] * min(max(span, 0), MAX_CELL_REPEAT))
        rows.append(cells)
    return (rows or None), {"sheet": ws.get(q("Name")) or "Sheet1"}


_TBL_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
_OFF_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
_TXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"


def _ods_para_text(el, x: str) -> str:
    """正文段落文本：按 ODF 规范还原 text:s（空格）、text:tab、text:line-break。"""
    parts = [el.text or ""]
    for child in el:
        if child.tag == x + "s":
            try:
                n = max(1, int(child.get(x + "c") or 1))
            except ValueError:
                n = 1
            parts.append(" " * n)
        elif child.tag == x + "tab":
            parts.append("\t")
        elif child.tag == x + "line-break":
            parts.append(" ")
        else:
            parts.append(_ods_para_text(child, x))
        parts.append(child.tail or "")
    return "".join(parts)


def _ods_cell_text(cell, t: str, o: str, x: str) -> str:
    vt = cell.get(o + "value-type")
    if vt in ("float", "currency", "percentage"):
        v = cell.get(o + "value")
        if v is not None:
            return _num_text(v)
    elif vt in ("date", "time"):
        v = cell.get(o + "date-value") or cell.get(o + "time-value") or ""
        if v:
            return norm_text(v)
    elif vt == "boolean":
        v = (cell.get(o + "boolean-value") or "").lower()
        if v:
            return "TRUE" if v == "true" else "FALSE"
    parts = [_ods_para_text(p, x) for p in cell.iter(x + "p")]
    return norm_text(" ".join(s for s in (p.strip() for p in parts) if s))


def _read_ods(path: str):
    """.ods（OpenDocument）：只读第一张工作表。"""
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise MergeError("不是有效的 .ods 文件（若被加密请先解密）：%s" % path)
    with zf:
        if "content.xml" not in zf.namelist():
            raise MergeError("不是有效的 .ods 文件（缺少 content.xml）：%s" % path)
        if zf.getinfo("content.xml").file_size > MAX_ZIP_XML_BYTES:
            raise MergeError(".ods 内容过大（>%dMB），请另存为 .xlsx 后重试：%s"
                             % (MAX_ZIP_XML_BYTES // 1048576, path))
        data = zf.read("content.xml")
    root = _safe_xml(data, ".ods 的 content.xml")
    t = "{%s}" % _TBL_NS          # 元素/属性名统一加命名空间前缀
    o = "{%s}" % _OFF_NS
    x = "{%s}" % _TXT_NS
    tables = list(root.iter(t + "table"))
    if not tables:
        return None, {"sheets": 0, "sheet": None}
    sh = tables[0]
    rows: list = []
    for tr in sh.findall(t + "table-row"):
        try:
            rrep = max(1, int(tr.get(t + "number-rows-repeated") or 1))
        except ValueError:
            rrep = 1
        cells: list = []
        for tc in tr:
            if tc.tag not in (t + "table-cell", t + "covered-table-cell"):
                continue
            try:
                crep = max(1, int(tc.get(t + "number-columns-repeated") or 1))
            except ValueError:
                crep = 1
            cells.extend([_ods_cell_text(tc, t, o, x)] * min(crep, MAX_CELL_REPEAT))
        while cells and cells[-1] == "":     # 去掉尾部填充（Calc 常见）
            cells.pop()
        if not cells and rrep > 1:
            continue
        for _ in range(min(rrep, MAX_ROWS + 1 - len(rows))):
            rows.append(list(cells))
        if len(rows) > MAX_ROWS:
            break
    return (rows or None), {"sheets": len(tables), "sheet": sh.get(t + "name") or "Sheet1"}


def _read_xls_ole(path: str):
    """旧版二进制 .xls（BIFF/OLE 复合文档）：需可选依赖 xlrd。"""
    if not HAVE_XLRD:
        raise MergeError(
            "检测到旧版二进制 .xls（BIFF/OLE 复合文档），本技能不内置二进制解析。请任选其一："
            "① 用 Excel / WPS 打开后「另存为」.xlsx（推荐，零依赖）；"
            "② 安装可选依赖后重试：pip install xlrd（BSD 许可，纯本地读取）。文件：%s" % path)
    try:
        book = _xlrd.open_workbook(path, on_demand=True)
        sh = book.sheet_by_index(0)
        rows = []
        for r in range(min(sh.nrows, MAX_ROWS + 1)):
            row = []
            for c in range(sh.ncols):
                v = sh.cell_value(r, c)
                row.append(_num_text(repr(v)) if isinstance(v, float) else norm_text(v))
            rows.append(row)
        info = {"sheets": book.nsheets, "sheet": sh.name}
        try:
            book.release_resources()
        except Exception:
            pass
    except MergeError:
        raise
    except Exception as e:                  # 加密/损坏的 .xls
        raise MergeError("旧版 .xls 解析失败（可能已加密或损坏）：%s（%s）" % (path, e))
    return (rows or None), info


def _ns(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag


def _xlsx_sheet_xml(zf: zipfile.ZipFile):
    """定位第一个数据工作表的 XML（跟随 workbook 中 sheet 顺序，跳过 chartsheet）。"""
    try:
        wb_xml = zf.read("xl/workbook.xml")
    except KeyError:
        raise MergeError("不是有效的 .xlsx 文件（缺少 xl/workbook.xml）")
    root = _safe_xml(wb_xml, "xl/workbook.xml")
    sheet_names = []
    for el in root.iter():
        if _ns(el.tag) == "sheet":
            name = el.get("name") or "Sheet1"
            sid = el.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            sheet_names.append((name, sid))
    try:
        rels_xml = zf.read("xl/_rels/workbook.xml.rels")
    except KeyError:
        rels_xml = b""
    target_of = {}
    if rels_xml:
        for el in _safe_xml(rels_xml, "xl/_rels/workbook.xml.rels").iter():
            if _ns(el.tag) == "Relationship" and el.get("Id") and el.get("Target"):
                target_of[el.get("Id")] = el.get("Target")
    for name, sid in sheet_names:
        target = target_of.get(sid or "", "")
        cand = "xl/" + str(target).lstrip("/")
        if cand.endswith(".xml") and "worksheets/sheet" in cand.replace("\\", "/"):
            try:
                return name, zf.read(cand)
            except KeyError:
                continue
    sheets = sorted([n for n in zf.namelist()
                     if re.match(r"xl/worksheets/sheet\d+\.xml$", n.replace("\\", "/"))])
    if not sheets:
        raise MergeError("xlsx 中未找到工作表数据")
    return short_name(sheets[0]), zf.read(sheets[0])


def _xlsx_shared_strings(zf: zipfile.ZipFile):
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    out = []
    for si in _safe_xml(data, "xl/sharedStrings.xml").iter():
        if _ns(si.tag) == "si":
            out.append("".join(t.text or "" for t in si.iter() if _ns(t.tag) == "t"))
    return out


def _xlsx_cell_value(cell, shared):
    t = cell.get("t", "")
    v_el = None
    for child in cell:
        if _ns(child.tag) == "v":
            v_el = child
            break
    if t == "inlineStr":
        return "".join(x.text or "" for x in cell.iter() if _ns(x.tag) == "t")
    if v_el is None or v_el.text is None:
        return ""
    v = v_el.text
    if t == "s":
        try:
            return shared[int(v)]
        except (ValueError, IndexError):
            return ""
    if t == "b":
        return "TRUE" if v == "1" else "FALSE"
    if t in ("e", "str"):
        return "" if t == "e" else v
    try:
        f = float(v)
        return str(int(f)) if f.is_integer() else str(f)
    except ValueError:
        return v


def _read_xlsx_stdlib(path: str):
    size = os.path.getsize(path)
    if size > MAX_FILE_BYTES:
        raise MergeError("文件过大")
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise MergeError("不是有效的 .xlsx 文件（若被加密请先解密）：%s" % path)
    with zf:
        if "xl/workbook.xml" not in zf.namelist():
            raise MergeError("不是有效的 .xlsx 文件：%s" % path)
        shared = _xlsx_shared_strings(zf)
        sheet_name, sheet_xml = _xlsx_sheet_xml(zf)
        grid: dict = {}
        max_row = max_col = -1
        for row in _safe_xml(sheet_xml, "xlsx 工作表 XML").iter():
            if _ns(row.tag) != "row":
                continue
            try:
                r = int(row.get("r", "1")) - 1
            except ValueError:
                continue
            for cell in row:
                if _ns(cell.tag) != "c":
                    continue
                m = _CELL_REF_RE.match(cell.get("r", ""))
                if not m:
                    continue
                c = col_index(m.group(1))
                grid[(r, c)] = norm_text(_xlsx_cell_value(cell, shared))
                max_row = max(max_row, r)
                max_col = max(max_col, c)
        rows = [[grid.get((r, c), "") for c in range(max_col + 1)] for r in range(max_row + 1)]
    rows = [r for r in rows if any(c != "" for c in r)]
    return rows, {"sheet": sheet_name}


def _read_xlsx_openpyxl(path: str):
    ws = _oxl.load_workbook(path, read_only=True, data_only=True).active
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append([norm_text("" if v is None else v) for v in row])
    rows = [r for r in rows if any(c != "" for c in r)]
    info = {"sheet": getattr(ws, "title", "Sheet1") or "Sheet1"}
    try:
        ws.parent.close()
    except Exception:
        pass
    return rows, info


def _normalize_rows(rows):
    width = max((len(r) for r in rows), default=0)
    return [r + [""] * (width - len(r)) for r in rows]


def _drop_empty_cols(rows):
    """去掉整列为空的列（稀疏表格的填充列、尾随空列）。"""
    if not rows:
        return rows
    width = len(rows[0])
    keep = [c for c in range(width) if any(r[c] != "" for r in rows)]
    if len(keep) == width:
        return rows
    return [[r[c] for c in keep] for r in rows]


def _find_header(rows):
    """表头行定位：首个含 >=2 个非空单元格的行；否则用第 1 行。"""
    for i, r in enumerate(rows[:10]):
        if sum(1 for c in r if c != "") >= 2:
            return i
    return 0


def read_table(path: str, backend: str = "auto"):
    """读取单个表格 -> (headers, data_rows, info)。data_rows 为 list[dict]。

    扩展名只作提示：先按内容嗅探格式（OOXML / ODS / 旧版 BIFF / HTML / Excel 2003 XML /
    分隔符文本），再派发到对应解析器；解析器返回 None 时回落为分隔符文本读取。
    """
    if not os.path.isfile(path):
        raise MergeError("文件不存在：%s" % path)
    if os.path.getsize(path) > MAX_FILE_BYTES:
        raise MergeError("文件过大（>%dMB），拒绝读取：%s" % (MAX_FILE_BYTES // 1048576, path))
    kind = sniff_format(path)
    if kind == "delimited":
        rows, info = _read_delimited(path)
    elif kind == "html":
        rows, info = _read_html(path)
        if rows is None:
            rows, info = _read_delimited(path)
            info["note"] = "HTML 中未找到表格，已按分隔符文本读取"
    elif kind == "spreadsheetml":
        rows, info = _read_spreadsheetml(path)
        if rows is None:
            rows, info = _read_delimited(path)
            info["note"] = "Excel 2003 XML 中未找到表数据，已按分隔符文本读取"
    elif kind == "ods":
        rows, info = _read_ods(path)
        if rows is None:
            raise MergeError("未在 .ods 中找到工作表：%s" % path)
    elif kind == "xls_ole":
        rows, info = _read_xls_ole(path)
        if rows is None:
            raise MergeError("未在 .xls 中找到工作表：%s" % path)
    else:                                  # xlsx / xlsm
        use = backend
        if use == "auto":
            use = "openpyxl" if HAVE_OPENPYXL else "stdlib"
        if use == "openpyxl":
            if not HAVE_OPENPYXL:
                raise MergeError("未安装 openpyxl（pip install openpyxl）")
            rows, info = _read_xlsx_openpyxl(path)
        else:
            rows, info = _read_xlsx_stdlib(path)
    info = dict(info or {})
    info["format"] = {"xlsx": "xlsx/xlsm", "xls_ole": "xls"}.get(kind, kind)
    if not rows:
        raise MergeError("文件中没有可读取的数据行：%s" % path)
    if len(rows) > MAX_ROWS + 10:
        raise MergeError("行数超过上限（%d 行）：%s" % (MAX_ROWS, path))
    rows = _drop_empty_cols(_normalize_rows(rows))
    hi = _find_header(rows)
    headers = [h if h else "列%d" % (i + 1) for i, h in enumerate(rows[hi])]
    seen: dict = {}
    fixed = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            fixed.append("%s（%d）" % (h, seen[h]))
        else:
            seen[h] = 1
            fixed.append(h)
    data = [dict(zip(fixed, r)) for r in rows[hi + 1:]]
    info.update({"path": path, "file": os.path.basename(path),
                 "headers": fixed, "n_rows": len(data), "n_cols": len(fixed),
                 "header_row": hi + 1, "backend": backend})
    return fixed, data, info


# ---------------- 文件写入：最小 XLSX（标准库）/ openpyxl ----------------

def _xlsx_esc(s: str) -> str:
    return _xml_escape(clean_xml(str(s)), {'"': "&quot;"})


def _col_widths(sheets) -> list:
    widths = []
    for sh in sheets:
        w = []
        for j in range(len(sh["headers"])):
            mx = len(sh["headers"][j] or "")
            for r in sh["rows"][:500]:
                mx = max(mx, len(str(r[j]) if j < len(r) else ""))
            w.append(min(60, max(8, int(mx * 1.6 + 2))))
        widths.append(w)
    return widths

def _write_xlsx_stdlib(path: str, sheets) -> None:
    widths = _col_widths(sheets)
    parts = {}
    parts["[Content_Types].xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        + "".join(
            '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
            for i in range(len(sheets)))
        + "</Types>")
    parts["_rels/.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>")
    rels = "".join(
        '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1)
        for i in range(len(sheets)))
    rels += '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (len(sheets) + 1)
    parts["xl/_rels/workbook.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + rels + "</Relationships>")
    sh_xml = "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (_xlsx_esc(sh["name"])[:31], i + 1, i + 1)
                     for i, sh in enumerate(sheets))
    parts["xl/workbook.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>" + sh_xml + "</sheets></workbook>")
    parts["xl/styles.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts><font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFDCE6F1"/><bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
        "</styleSheet>")
    for i, sh in enumerate(sheets):
        buf = ['<row r="1">']
        for j, h in enumerate(sh["headers"]):
            buf.append('<c r="%s1" t="inlineStr" s="1"><is><t>%s</t></is></c>'
                       % (col_letters(j), _xlsx_esc(h)))
        buf.append("</row>")
        for ri, row in enumerate(sh["rows"], start=2):
            buf.append('<row r="%d">' % ri)
            for j in range(len(sh["headers"])):
                v = "" if j >= len(row) or row[j] is None else str(row[j])
                buf.append('<c r="%s%d" t="inlineStr"><is><t>%s</t></is></c>'
                           % (col_letters(j), ri, _xlsx_esc(v)))
            buf.append("</row>")
        last = col_letters(len(sh["headers"]) - 1)
        cols = "".join('<col min="%d" max="%d" width="%.1f" customWidth="1"/>' % (j + 1, j + 1, float(w))
                       for j, w in enumerate(widths[i]))
        parts["xl/worksheets/sheet%d.xml" % (i + 1)] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            "<cols>%s</cols>" % cols
            + '<sheetViews><sheetView workbookViewId="0">'
            + '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            + "</sheetView></sheetViews>"
            + '<autoFilter ref="A1:%s%d"/>' % (last, len(sh["rows"]) + 1)
            + "<sheetData>" + "".join(buf) + "</sheetData></worksheet>")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data.encode("utf-8"))

def _write_xlsx_openpyxl(path: str, sheets) -> None:
    wb = _oxl.Workbook()
    wb.remove(wb.active)
    bold = _OxlFont(bold=True)
    try:
        from openpyxl.styles import PatternFill as _Fill
        hfill = _Fill("solid", fgColor="DCE6F1")
    except Exception:
        hfill = None
    widths = _col_widths(sheets)
    for i, sh in enumerate(sheets):
        ws = wb.create_sheet(sh["name"][:31])
        ws.append(list(sh["headers"]))
        for r in sh["rows"]:
            ext = (list(r) + [""] * len(sh["headers"]))[:len(sh["headers"])]
            ws.append(["" if v is None else v for v in ext])
        for cell in ws[1]:
            cell.font = bold
            if hfill is not None:
                cell.fill = hfill
        ws.freeze_panes = "A2"
        try:
            ws.auto_filter.ref = ws.dimensions
        except Exception:
            pass
        for j, w in enumerate(widths[i]):
            ws.column_dimensions[col_letters(j)].width = w
    wb.save(path)


def write_result_xlsx(path: str, sheets, backend: str = "auto") -> str:
    use = backend
    if use == "auto":
        use = "openpyxl" if HAVE_OPENPYXL else "stdlib"
    if use == "openpyxl":
        if not HAVE_OPENPYXL:
            raise MergeError("未安装 openpyxl（pip install openpyxl）")
        _write_xlsx_openpyxl(path, sheets)
    else:
        _write_xlsx_stdlib(path, sheets)
    return use


# ---------------- inspect / suggest-keys ----------------

def cmd_inspect(files, sample_n: int = 5) -> dict:
    if len(files) < 2:
        raise MergeError("至少需要 2 个表格文件。")
    seen = set()
    tables = []
    for f in files:
        p = os.path.abspath(f)
        if p in seen:
            raise MergeError("输入文件重复：%s" % f)
        seen.add(p)
        headers, data, info = read_table(p)
        tables.append({"file": info["file"], "path": p,
                       "headers": headers, "n_rows": info["n_rows"], "n_cols": info["n_cols"],
                       "header_row": info["header_row"], "format": info.get("format"),
                       "encoding": info.get("encoding"), "delimiter": info.get("delimiter"),
                       "sheet": info.get("sheet"),
                       "col_kinds": {h: guess_col_kind(h) for h in headers},
                       "sample": data[:max(0, sample_n)]})
    return {"tool": TOOL_NAME, "version": VERSION, "created": now_str(),
            "sample_n": sample_n, "tables": tables}


def _key_values(rows, col, kind):
    vals, nulls = [], 0
    for r in rows:
        raw = norm_text(r.get(col, ""))
        if raw == "":
            nulls += 1
            continue
        vals.append(norm_key(raw, kind))
    return vals, nulls


def _pair_stats(va, vb):
    sa, sb = set(va), set(vb)
    inter = sa & sb
    union = sa | sb
    return {
        "a_unique": len(sa), "b_unique": len(sb),
        "overlap": len(inter),
        "overlap_rate_vs_a": round(len(inter) / len(sa), 4) if sa else 0.0,
        "overlap_rate_vs_b": round(len(inter) / len(sb), 4) if sb else 0.0,
        "jaccard": round(len(inter) / len(union), 4) if union else 0.0,
    }

def cmd_suggest_keys(profile: dict) -> dict:
    tables = profile.get("tables", [])
    if len(tables) < 2:
        raise MergeError("profile 中表格不足 2 个。")
    full = []
    for t in tables:
        _, data, _ = read_table(t["path"])
        full.append(data)
    cands = []
    for i in range(len(tables)):
        for j in range(i + 1, len(tables)):
            ta, tb = tables[i], tables[j]
            for ca in ta["headers"]:
                ka = guess_col_kind(ca)
                if ka not in ("name", "id"):
                    continue
                va, na = _key_values(full[i], ca, ka)
                if not va:
                    continue
                for cb in tb["headers"]:
                    kb = guess_col_kind(cb)
                    if kb not in ("name", "id"):
                        continue
                    if ka != kb:
                        continue
                    vb, nb = _key_values(full[j], cb, kb)
                    if not vb:
                        continue
                    st = _pair_stats(va, vb)
                    a_only = sorted(set(va) - set(vb))[:5]
                    b_only = sorted(set(vb) - set(va))[:5]
                    dup_a = len(va) - len(set(va))
                    dup_b = len(vb) - len(set(vb))
                    cover = min(st["overlap_rate_vs_a"], st["overlap_rate_vs_b"])
                    uniq = 1.0 - (dup_a / len(va) + dup_b / len(vb)) / 2.0
                    score = round(0.6 * cover + 0.25 * st["jaccard"] + 0.15 * uniq, 4)
                    ev = ("“%s”[%s] 与 “%s”[%s] 清洗后：A 侧 %d/%d 对上（%.1f%%），"
                          "B 侧 %d/%d 对上（%.1f%%）" % (
                              ta["file"], ca, tb["file"], cb,
                              st["overlap"], st["a_unique"], st["overlap_rate_vs_a"] * 100,
                              st["overlap"], st["b_unique"], st["overlap_rate_vs_b"] * 100))
                    if na or nb:
                        ev += "；空值 A=%d/B=%d" % (na, nb)
                    if dup_a or dup_b:
                        ev += "；重复键 A=%d/B=%d（疑似重名或重复行）" % (dup_a, dup_b)
                    if a_only or b_only:
                        ev += "；疑似差异示例 A 独有%s / B 独有%s" % (a_only[:3], b_only[:3])
                    cands.append({
                        "a_file": ta["file"], "a_col": ca, "a_kind": ka,
                        "b_file": tb["file"], "b_col": cb, "b_kind": kb,
                        "score": score, "stats": st,
                        "nulls": {"a": na, "b": nb},
                        "dup_keys": {"a": dup_a, "b": dup_b},
                        "evidence": ev,
                    })
    cands.sort(key=lambda c: (-c["score"], c["a_file"], c["a_col"]))
    rec = None
    if cands:
        top = cands[0]
        note = "推荐用“%s”列互相比对；" % ("工号类" if top["a_kind"] == "id" else "姓名类")
        if top["dup_keys"]["a"] or top["dup_keys"]["b"]:
            note += "注意存在重复键，合并后请核对“表内重复键”清单；"
        kinds = {c["a_kind"] for c in cands[:3]} | {c["b_kind"] for c in cands[:3]}
        if "id" in kinds and top["a_kind"] == "name":
            note += "检测到工号类列，若姓名有一致率短板请优先改用工号。"
        else:
            note += "确认无误后即可写入 plan.json 执行合并。"
        rec = {"text": "首选：" + top["evidence"] + "。" + note,
               "top": {"a_file": top["a_file"], "a_col": top["a_col"],
                       "b_file": top["b_file"], "b_col": top["b_col"]}}
    return {"tool": TOOL_NAME, "version": VERSION, "created": now_str(),
            "candidates": cands, "recommendation": rec,
            "hint": "请用户确认关联键后再写入 plan.json；不要跳过确认直接合并。"}


# ---------------- merge ----------------

def _load_plan(plan_path: str) -> dict:
    try:
        with open(plan_path, "r", encoding="utf-8-sig") as f:
            plan = json.load(f)
    except (OSError, ValueError) as e:
        raise MergeError("plan.json 读取失败：%s" % e)
    if not isinstance(plan, dict) or not plan.get("tables"):
        raise MergeError("plan.json 格式错误：缺少 tables。")
    if len(plan["tables"]) < 2:
        raise MergeError("plan.json 至少需要 2 个表。")
    pol = plan.get("unmatched_policy", "separate_sheet")
    if pol not in ("separate_sheet", "keep_marked", "drop"):
        raise MergeError("unmatched_policy 非法：%s" % pol)
    for t in plan["tables"]:
        if not t.get("path") or not t.get("key_col"):
            raise MergeError("plan.json 每个表必须包含 path 与 key_col。")
    return plan


def _match_recipe(tables_info, recipe: dict):
    """配方命中判定：文件集合一致 + 每个文件的键列仍在表头中。

    兼容两种调用来源：配方里的条目用 `file`（文件名）标识，
    而 plan.json 里的条目通常只写了 `path`——这里统一按文件名比对。
    """
    def name_of(t: dict) -> str:
        return str(t.get("file") or os.path.basename(str(t.get("path", ""))))

    try:
        want = {r["file"]: r.get("key_col") for r in recipe.get("tables", [])}
        got = {name_of(t): t.get("key_col") for t in tables_info}
        if set(want) != set(got):
            return False, "文件集合不一致"
        for f, kc in want.items():
            if not kc:
                return False, "表 %s 的键列已不存在" % f
        return True, "配方命中：键映射与策略直接复用"
    except Exception as e:
        return False, "配方解析失败：%s" % e


def _closest(k: str, pool, n: int = 3) -> str:
    if not pool:
        return ""
    try:
        m = difflib.get_close_matches(k, pool, n=n, cutoff=0.5)
    except Exception:
        m = []
    return "、".join(m)

def cmd_merge(plan_path: str, out_xlsx: str, backend: str = "auto",
              recipe_path=None, save_recipe=None) -> dict:
    plan = _load_plan(plan_path)
    recipe_used = None
    if recipe_path:
        try:
            with open(recipe_path, "r", encoding="utf-8-sig") as f:
                recipe = json.load(f)
        except (OSError, ValueError) as e:
            raise MergeError("recipe 读取失败：%s" % e)
        ok, reason = _match_recipe(plan["tables"], recipe)
        if not ok:
            raise MergeError("配方未命中（%s），请重新确认关联键。" % reason)
        recipe_used = reason
    primary = plan.get("primary") or plan["tables"][0].get("file")
    policy = plan.get("unmatched_policy", "separate_sheet")

    loaded = []
    for t in plan["tables"]:
        headers, data, info = read_table(os.path.abspath(t["path"]), backend=backend)
        kc = t["key_col"]
        if kc not in headers:
            raise MergeError("表 %s 不存在键列“%s”（表头为：%s）" % (info["file"], kc, "、".join(headers)))
        kind = "id" if guess_col_kind(kc) == "id" else "name"
        dups: dict = {}
        nulls = 0
        for idx, r in enumerate(data):
            raw = norm_text(r.get(kc, ""))
            if raw == "":
                nulls += 1
                continue
            dups.setdefault(norm_key(raw, kind), []).append(idx)
        dup_keys = {k: v for k, v in dups.items() if len(v) > 1}
        loaded.append({"spec": t, "file": info["file"], "headers": headers,
                       "data": data, "info": info, "kind": kind,
                       "null_keys": nulls, "dup_keys": dup_keys})

    order = sorted(range(len(loaded)),
                   key=lambda i: 0 if loaded[i]["file"] == primary else 1)
    all_keys: list = []
    seen_keys = set()
    for i in order:
        L = loaded[i]
        for r in L["data"]:
            raw = norm_text(r.get(L["spec"]["key_col"], ""))
            if raw == "":
                continue
            k = norm_key(raw, L["kind"])
            if k not in seen_keys:
                seen_keys.add(k)
                all_keys.append(k)

    file_rows = {}
    for L in loaded:
        m: dict = {}
        for r in L["data"]:
            raw = norm_text(r.get(L["spec"]["key_col"], ""))
            if not raw:
                continue
            m.setdefault(norm_key(raw, L["kind"]), []).append(r)
        file_rows[L["file"]] = m

    out_headers = ["关联键"]
    col_map = []
    for L in loaded:
        kc = L["spec"]["key_col"]
        want = L["spec"].get("fields") or [h for h in L["headers"] if h != kc]
        for c in want:
            if c == kc or c not in L["headers"]:
                continue
            out = "%s（来源：%s）" % (c, L["file"])
            out_headers.append(out)
            col_map.append((L["file"], c, out))
    out_headers.append("数据状态")

    logic_groups: dict = {}
    for f, c, o in col_map:
        logic_groups.setdefault(logic_name(c), []).append((f, c, o))
    # 同一逻辑字段、且来自 >=2 个不同来源的列才成组
    logic_groups = {k: v for k, v in logic_groups.items()
                    if len({f for f, _, _ in v}) >= 2}

    result_rows, unmatched_rows, conflicts = [], [], []
    matched_per_file = {L["file"]: 0 for L in loaded}
    for k in all_keys:
        present = [L["file"] for L in loaded if k in file_rows[L["file"]]]
        full_hit = len(present) == len(loaded)
        row = {"关联键": k}
        flags = []
        if not full_hit:
            missing = [L["file"] for L in loaded if L["file"] not in present]
            flags.append("部分缺失（缺：%s）" % "、".join(missing))
            for L in loaded:
                if L["file"] not in present:
                    continue
                r0 = file_rows[L["file"]][k][0]
                unmatched_rows.append({
                    "来源表": L["file"], "原始键值": r0.get(L["spec"]["key_col"], ""),
                    "归一化键值": k, "缺席表": "、".join(missing),
                    "最相近候选": _closest(k, [kk for kk in all_keys if kk != k]),
                })
        for logic, cols in logic_groups.items():
            if len(cols) < 2:
                continue
            vals = {}
            for f, c, o in cols:
                if f in present:
                    v = norm_text(file_rows[f][k][0].get(c, ""))
                    if v:
                        vals[o] = v
            if len(set(vals.values())) > 1:
                flags.append("冲突：%s" % logic)
                conflicts.append({"关联键": k, "逻辑字段": logic,
                                  "取值": "；".join("%s=%s" % (o, v) for o, v in vals.items())})
        for f, c, o in col_map:
            row[o] = file_rows[f][k][0].get(c, "") if f in present else ""
        for f in present:
            matched_per_file[f] += 1
        row["数据状态"] = "；".join(flags) if flags else "完整匹配"
        if policy == "drop" and not full_hit:
            continue
        result_rows.append([row.get(h, "") for h in out_headers])

    log = [
        ["多表人员信息汇总助手 · 合并日志", ""],
        ["执行时间", now_str()],
        ["脚本版本", "%s %s（读取后端：%s）" % (TOOL_NAME, VERSION, backend)],
        ["关联键选择", "；".join("%s[%s]" % (L["file"], L["spec"]["key_col"]) for L in loaded)],
        ["主表", primary],
        ["未匹配处置策略", {"separate_sheet": "单独输出到“未匹配记录”sheet",
                        "keep_marked": "保留在结果表并标记",
                        "drop": "从结果表剔除（仍记入本日志备查）"}[policy]],
        ["配方", recipe_used or "未使用（本次为人工确认）"],
        ["", ""],
        ["输入文件", "行数 / 列数 / 说明"],
    ]
    for L in loaded:
        extra = []
        if L["info"].get("encoding"):
            extra.append("编码 " + str(L["info"]["encoding"]))
        if L["info"].get("sheet"):
            extra.append("工作表 " + str(L["info"]["sheet"]))
        log.append(["%s（键列：%s）" % (L["file"], L["spec"]["key_col"]),
                    "%d 行 / %d 列%s" % (L["info"]["n_rows"], L["info"]["n_cols"],
                                       "（" + "，".join(extra) + "）" if extra else "")])
    log += [["", ""], ["匹配统计", ""],
            ["全局不同关联键数", str(len(all_keys))],
            ["结果表行数", str(len(result_rows))],
            ["未匹配记录条数", str(len(unmatched_rows))],
            ["字段冲突条数", str(len(conflicts))]]
    for L in loaded:
        mr = matched_per_file[L["file"]]
        n = L["info"]["n_rows"]
        log.append(["表 %s 匹配覆盖" % L["file"],
                    "%d/%d（%.1f%%）" % (mr, n, mr * 100.0 / n if n else 0)])
    log.append(["表内重复键", ""])
    any_dup = False
    for L in loaded:
        if L["dup_keys"]:
            any_dup = True
            log.append(["表 %s" % L["file"], "%d 组重复（示例：%s），已全部保留请人工核对"
                        % (len(L["dup_keys"]), list(L["dup_keys"])[:5])])
    if not any_dup:
        log.append(["（无）", "各表键列均唯一"])
    log.append(["空键值（键列为空）", ""])
    for L in loaded:
        log.append(["表 %s" % L["file"], "%d 行键列为空，未参与匹配" % L["null_keys"]])
    log += [["", ""], ["字段冲突清单（均已保留原值、仅标记）", ""]]
    if conflicts:
        for c in conflicts[:50]:
            log.append([c["关联键"] + " / " + c["逻辑字段"], c["取值"]])
        if len(conflicts) > 50:
            log.append(["……", "共 %d 条，仅展示前 50 条" % len(conflicts)])
    else:
        log.append(["（无）", "多表同名字段取值一致或仅单表有值"])
    log += [["", ""], ["免责声明",
              "本日志为责任记录：Skill 不判断数据对错、不承诺一次合并永远正确；"
              "请先核对本日志与“未匹配记录”，再使用汇总结果。"]]

    un_headers = ["来源表", "原始键值", "归一化键值", "缺席表", "最相近候选"]
    un_rows = [[u[h] for h in un_headers] for u in unmatched_rows]
    sheets = [{"name": "汇总结果", "headers": out_headers, "rows": result_rows},
              {"name": "合并日志", "headers": ["项目", "内容"], "rows": log},
              {"name": "未匹配记录", "headers": un_headers, "rows": un_rows}]
    used_backend = write_result_xlsx(out_xlsx, sheets, backend=backend)
    report = {"tool": TOOL_NAME, "version": VERSION, "created": now_str(),
              "backend": used_backend, "output": os.path.abspath(out_xlsx),
              "n_keys": len(all_keys), "n_result_rows": len(result_rows),
              "n_unmatched": len(unmatched_rows), "n_conflicts": len(conflicts),
              "matched_per_file": matched_per_file,
              "policy": policy, "recipe": recipe_used}
    rep_path = os.path.splitext(out_xlsx)[0] + ".report.json"
    try:
        with open(rep_path, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        report["report_path"] = os.path.abspath(rep_path)
    except OSError:
        pass
    if save_recipe:
        rec = {"tool": TOOL_NAME, "version": VERSION, "created": now_str(),
               "tables": [{"file": L["file"], "path": L["spec"]["path"],
                           "key_col": L["spec"]["key_col"],
                           "fields": L["spec"].get("fields"),
                           "primary": (L["file"] == primary)} for L in loaded],
               "primary": primary, "unmatched_policy": policy}
        with open(save_recipe, "w", encoding="utf-8") as f:
            json.dump(rec, f, ensure_ascii=False, indent=2)
        report["recipe_saved"] = os.path.abspath(save_recipe)
    return report


# ---------------- selfcheck / selftest ----------------

NET_IMPORTS = ("urllib", "requests", "socket", "http", "ssl", "ftplib", "smtplib", "websocket")
# 以下仅用于向用户展示"检查了哪些危险模式"，实际检测走 AST（见 cmd_selfcheck），
# 因此本元组自身的字面量不会被误报。
DANGER_PATTERNS = ("os.system/os.popen/os.exec* 调用", "subprocess.* 调用", "eval()/exec() 调用")


def _ast_danger_calls(tree):
    """AST 级检测真正的危险调用点（字符串字面量不计入，避免自指误报）。"""
    hits = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        f = node.func
        if isinstance(f, ast.Name) and f.id in ("eval", "exec"):
            hits.add(f.id + "()")
        elif isinstance(f, ast.Attribute):
            v = f.value
            base = v.id if isinstance(v, ast.Name) else ""
            if base == "os" and f.attr in ("system", "popen", "execv", "execl",
                                           "execve", "execlp", "execvp", "spawnl",
                                           "spawnle", "spawnlp", "spawnv", "spawnve"):
                hits.add("os." + f.attr)
            elif base == "subprocess":
                hits.add("subprocess." + f.attr)
    return sorted(hits)


def cmd_selfcheck() -> dict:
    me = os.path.abspath(__file__)
    with open(me, "r", encoding="utf-8") as f:
        src = f.read()
    tree = ast.parse(src)
    imported = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            imported.add(node.module.split(".")[0])
    bad_imports = sorted(set(NET_IMPORTS) & imported)
    bad_calls = _ast_danger_calls(tree)
    ok = not bad_imports and not bad_calls
    return {"tool": TOOL_NAME, "version": VERSION, "ok": ok,
            "network_imports": bad_imports, "danger_calls": bad_calls,
            "checked_patterns": list(DANGER_PATTERNS),
            "backend_available": {"stdlib": True, "openpyxl": HAVE_OPENPYXL, "xlrd": HAVE_XLRD},
            "optional_versions": {"openpyxl": OPENPYXL_VERSION, "xlrd": XLRD_VERSION},
            "supported_formats": list(ALLOWED_EXTS),
            "conclusion": "通过：无网络模块、无外部命令调用，数据不出本地。"
            if ok else "未通过：发现可疑依赖/调用，请审查。"}


SELFTEST_FILES = {
    "recruit.csv": (
        "姓名,手机,应聘岗位\n"
        "张三,13800000001,产品经理\n"
        " 李四 ,13800000002,设计师\n"
        "王五,13800000003,工程师\n"
        "赵六,,运营\n"
        "钱七,13800000007,工程师\n"
    ),
    "attend.csv": (
        "员工姓名,部门,工号\n"
        "张三,产品部,EMP-1001\n"
        "李四,设计部,EMP-1002\n"
        "王　五,技术部,EMP-1003\n"
        "孙八,市场部,EMP-1008\n"
        "赵六,运营部,EMP-1006\n"
    ),
    "train.csv": (
        "姓名,完成课程,证书\n"
        "张三,入职培训,A\n"
        "李四,入职培训,A\n"
        "王五,安全培训,B\n"
        "周九,入职培训,A\n"
        "钱七,,C\n"
    ),
}

# 阶段 2 用到：把同一批人员分别写成"伪 .xls（HTML）"与".ods（OpenDocument）"，
# 用来证明跨格式合并可用（两种文件都由标准库现场生成/解析，不依赖 Office）。
SELFTEST_HTML_XLS = (
    "<html><head><meta charset='utf-8'><title>旧系统导出</title></head><body>"
    "<table border='1'>"
    "<tr><th>姓名</th><th>手机</th><th>入职日期</th></tr>"
    "<tr><td>张三</td><td>13800000001</td><td>2024-03-01</td></tr>"
    "<tr><td>李四</td><td>13800000002</td><td>2024-03-05</td></tr>"
    "</table></body></html>")

SELFTEST_ODS_XML = (
    "<?xml version='1.0' encoding='UTF-8'?>"
    "<office:document-content "
    "xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' "
    "xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0' "
    "xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0' office:version='1.2'>"
    "<office:body><office:spreadsheet><table:table table:name='培训'>"
    "<table:table-row>"
    "<table:table-cell office:value-type='string'><text:p>姓名</text:p></table:table-cell>"
    "<table:table-cell office:value-type='string'><text:p>完成课程</text:p></table:table-cell>"
    "</table:table-row>"
    "<table:table-row>"
    "<table:table-cell office:value-type='string'><text:p>张三</text:p></table:table-cell>"
    "<table:table-cell office:value-type='string'><text:p>入职培训</text:p></table:table-cell>"
    "</table:table-row>"
    "<table:table-row>"
    "<table:table-cell office:value-type='string'><text:p>李四</text:p></table:table-cell>"
    "<table:table-cell office:value-type='string'><text:p>安全培训</text:p></table:table-cell>"
    "<table:table-cell table:number-columns-repeated='10'/>"
    "</table:table-row>"
    "</table:table></office:spreadsheet></office:body></office:document-content>")

def cmd_selftest(backend: str = "stdlib") -> dict:
    tmp = tempfile.mkdtemp(prefix="people_merge_selftest_")
    try:
        paths = []
        for name, content in SELFTEST_FILES.items():
            p = os.path.join(tmp, name)
            with open(p, "w", encoding="utf-8-sig", newline="") as f:
                f.write(content)
            paths.append(p)
        prof = cmd_inspect(paths, sample_n=3)
        assert len(prof["tables"]) == 3, "inspect 应返回 3 张表"
        keys = cmd_suggest_keys(prof)
        assert keys["candidates"], "应找到候选关联键"
        top = keys["candidates"][0]
        assert top["score"] >= 0.5, "姓名键一致率过低：%s" % (top,)
        plan = {"primary": "recruit.csv", "unmatched_policy": "separate_sheet",
                "tables": [{"path": paths[0], "key_col": "姓名"},
                           {"path": paths[1], "key_col": "员工姓名"},
                           {"path": paths[2], "key_col": "姓名"}]}
        plan_p = os.path.join(tmp, "plan.json")
        with open(plan_p, "w", encoding="utf-8") as f:
            json.dump(plan, f, ensure_ascii=False)
        out = os.path.join(tmp, "汇总结果.xlsx")
        rep = cmd_merge(plan_p, out, backend=backend,
                        save_recipe=os.path.join(tmp, "recipe.json"))
        assert os.path.isfile(out) and os.path.getsize(out) > 0, "xlsx 未生成"
        _, back, _ = read_table(out, backend="stdlib")
        assert back, "写出的 xlsx 读回为空"
        assert rep["n_keys"] == 7, "期望 7 个不同人员，实际 %s" % rep["n_keys"]
        assert rep["n_unmatched"] == 6, "期望 6 条未匹配，实际 %s" % rep["n_unmatched"]

        # ---- 阶段 2：跨格式（伪 .xls=HTML / .ods / .tsv）也能识别与合并 ----
        xls_p = os.path.join(tmp, "旧系统导出.xls")
        with open(xls_p, "w", encoding="utf-8", newline="") as f:
            f.write(SELFTEST_HTML_XLS)
        ods_p = os.path.join(tmp, "培训平台导出.ods")
        with zipfile.ZipFile(ods_p, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
            zf.writestr("content.xml", SELFTEST_ODS_XML)
        tsv_p = os.path.join(tmp, "考勤导出.tsv")
        with open(tsv_p, "w", encoding="utf-8-sig", newline="") as f:
            f.write("姓名\t部门\n张三\t产品部\n李四\t设计部\n")
        prof2 = cmd_inspect([xls_p, ods_p, tsv_p], sample_n=2)
        kinds = [t["format"] for t in prof2["tables"]]
        assert kinds == ["html", "ods", "delimited"], "跨格式识别异常：%s" % kinds
        plan2 = {"primary": "旧系统导出.xls", "unmatched_policy": "separate_sheet",
                 "tables": [{"path": xls_p, "key_col": "姓名"},
                            {"path": ods_p, "key_col": "姓名"},
                            {"path": tsv_p, "key_col": "姓名"}]}
        plan2_p = os.path.join(tmp, "plan2.json")
        with open(plan2_p, "w", encoding="utf-8") as f:
            json.dump(plan2, f, ensure_ascii=False)
        out2 = os.path.join(tmp, "跨格式汇总.xlsx")
        rep2 = cmd_merge(plan2_p, out2, backend=backend)
        assert rep2["n_keys"] == 2 and rep2["n_result_rows"] == 2, \
            "跨格式合并异常：keys=%s rows=%s" % (rep2["n_keys"], rep2["n_result_rows"])
        h2, back2, _ = read_table(out2, backend="stdlib")
        row_zhang = [r for r in back2 if r.get("关联键") == "张三"][0]
        assert any("产品部" == v for v in row_zhang.values()), "TSV 字段未并入：%s" % row_zhang
        assert any("入职培训" == v for v in row_zhang.values()), "ODS 字段未并入：%s" % row_zhang
        assert all("冲突" not in (r.get("数据状态") or "") for r in back2), "跨格式用例不应有冲突"

        chk = cmd_selfcheck()
        assert chk["ok"], "selfcheck 未通过：%s" % chk
        return {"ok": True, "backend": rep["backend"], "report": rep,
                "formats_checked": kinds, "report_cross_format": rep2,
                "top_evidence": top["evidence"]}
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)


def _dump(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, indent=2)

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="people_merge",
                                 description="多表人员信息汇总助手 · 合并引擎（本地执行，不联网）")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("inspect", help="只读表头+样本 -> profile.json")
    p.add_argument("--files", nargs="+", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--sample", type=int, default=5)
    p.add_argument("--backend", default="auto", choices=("auto", "stdlib", "openpyxl"))
    p = sub.add_parser("suggest-keys", help="候选关联键 + 证据 -> keys.json")
    p.add_argument("--profile", required=True)
    p.add_argument("--out", required=True)
    p = sub.add_parser("merge", help="按 plan.json 合并 -> 汇总结果.xlsx")
    p.add_argument("--plan", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--backend", default="auto", choices=("auto", "stdlib", "openpyxl"))
    p.add_argument("--recipe", default=None)
    p.add_argument("--save-recipe", default=None)
    p = sub.add_parser("selftest", help="端到端自测（临时目录，不污染用户目录）")
    p.add_argument("--backend", default="stdlib", choices=("auto", "stdlib", "openpyxl"))
    sub.add_parser("selfcheck", help="自证无网络依赖/外部命令调用")
    a = ap.parse_args(argv)
    try:
        if a.cmd == "inspect":
            prof = cmd_inspect(a.files, a.sample)
            with open(a.out, "w", encoding="utf-8") as f:
                json.dump(prof, f, ensure_ascii=False, indent=2)
            print("inspect ok: %d tables -> %s" % (len(prof["tables"]), a.out))
        elif a.cmd == "suggest-keys":
            with open(a.profile, "r", encoding="utf-8-sig") as f:
                prof = json.load(f)
            keys = cmd_suggest_keys(prof)
            with open(a.out, "w", encoding="utf-8") as f:
                json.dump(keys, f, ensure_ascii=False, indent=2)
            print("suggest-keys ok: %d candidates -> %s" % (len(keys["candidates"]), a.out))
            if keys["recommendation"]:
                print(keys["recommendation"]["text"])
        elif a.cmd == "merge":
            rep = cmd_merge(a.plan, a.out, backend=a.backend,
                            recipe_path=a.recipe, save_recipe=a.save_recipe)
            print("merge ok: %d rows, %d unmatched, %d conflicts -> %s"
                  % (rep["n_result_rows"], rep["n_unmatched"], rep["n_conflicts"], a.out))
            print(_dump({k: rep[k] for k in ("backend", "n_keys", "matched_per_file", "policy")}))
        elif a.cmd == "selftest":
            r = cmd_selftest(backend=a.backend)
            print("selftest ok (backend=%s): %s" % (r["backend"], r["top_evidence"]))
            print("跨格式用例通过（识别为 %s）：%s 行 / %s 个关联键"
                  % (" + ".join(r["formats_checked"]),
                     r["report_cross_format"]["n_result_rows"],
                     r["report_cross_format"]["n_keys"]))
        elif a.cmd == "selfcheck":
            print(_dump(cmd_selfcheck()))
    except MergeError as e:
        print("错误：%s" % e, file=sys.stderr)
        return 2
    except AssertionError as e:
        print("自测断言失败：%s" % e, file=sys.stderr)
        return 3
    return 0


if __name__ == "__main__":
    sys.exit(main())

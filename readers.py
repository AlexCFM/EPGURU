"""Чтение любых табличных файлов в openpyxl.Workbook.

Формат определяется по СОДЕРЖИМОМУ, а не по расширению: многие системы выгружают
HTML или XML и называют его .xls.

Поддерживается:
  .xlsx .xlsm .xltx .xltm  — Excel 2007+ (форматирование сохраняется)
  .xls                     — Excel 97–2003 (двоичный)
  .xml                     — Excel 2003 XML (SpreadsheetML) и любой «табличный» XML
  .ods                     — LibreOffice / OpenOffice
  .csv .tsv .txt           — любые разделители (; , таб |), кодировки UTF-8 / Windows-1251
  .htm .html и «.xls»-HTML — HTML-таблицы
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import re
import zipfile
from html.parser import HTMLParser
from pathlib import Path

import openpyxl
from defusedxml import ElementTree as ET

SUPPORTED_EXT = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xml", ".ods",
                 ".csv", ".tsv", ".txt", ".htm", ".html"}
MAX_EMPTY_REPEAT = 1000      # защита от «миллиона пустых колонок» в ODS


class ReadError(Exception):
    pass


# ---------- определение формата ----------

def sniff(data: bytes) -> str:
    if data[:4] == b"PK\x03\x04":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                names = set(z.namelist())
                if "mimetype" in names and b"opendocument.spreadsheet" in z.read("mimetype"):
                    return "ods"
                if any(n.startswith("xl/") for n in names):
                    return "xlsx"
        except zipfile.BadZipFile:
            pass
        raise ReadError("Архив не похож ни на Excel, ни на ODS")
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "xls"
    head = _decode(data[:4096]).lstrip("﻿ \t\r\n").lower()
    if head.startswith("<?xml") or head.startswith("<workbook") or head.startswith("<ss:workbook"):
        if "urn:schemas-microsoft-com:office:spreadsheet" in head or "<workbook" in head:
            return "spreadsheetml"
        if "<html" in head or "<table" in head:
            return "html"
        return "xml"
    if head.startswith("<!doctype html") or head.startswith("<html") or "<table" in head[:2000]:
        return "html"
    if head.startswith("<"):
        return "xml"
    return "csv"


def _decode(data: bytes) -> str:
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1")


# ---------- читатели ----------

def _new_wb():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    return wb


def _add_sheet(wb, title, rows):
    title = re.sub(r"[\[\]\*\?/\\:]", "_", str(title or "Sheet"))[:31] or "Sheet"
    base, i = title, 2
    while title in wb.sheetnames:
        title = f"{base[:28]}_{i}"
        i += 1
    ws = wb.create_sheet(title)
    for r in rows:
        ws.append(list(r))
    return ws


def read_xlsx(data):
    return openpyxl.load_workbook(io.BytesIO(data))


def read_xls(data):
    import xlrd
    book = xlrd.open_workbook(file_contents=data)
    wb = _new_wb()
    for sh in book.sheets():
        rows = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                v = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    d = xlrd.xldate_as_datetime(v, book.datemode)
                    v = d.time() if 0 <= cell.value < 1 else d
                elif cell.ctype == xlrd.XL_CELL_NUMBER and float(v).is_integer():
                    v = int(v)
                elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    v = None
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    v = bool(v)
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    v = None
                row.append(v)
            rows.append(row)
        _add_sheet(wb, sh.name, rows)
    return wb


def read_csv(data):
    text = _decode(data)
    sample = text[:20000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,\t|")
        delim = dialect.delimiter
    except csv.Error:
        counts = {d: sample.count(d) for d in ";,\t|"}
        delim = max(counts, key=counts.get) if any(counts.values()) else ","
    rows = [[(c if c != "" else None) for c in row] for row in csv.reader(io.StringIO(text), delimiter=delim)]
    wb = _new_wb()
    _add_sheet(wb, "Sheet1", rows)
    return wb


SS = "urn:schemas-microsoft-com:office:spreadsheet"


def _ss_value(typ, text):
    if text is None:
        return None
    if typ == "Number":
        f = float(text)
        return int(f) if f.is_integer() else f
    if typ == "DateTime":
        d = dt.datetime.fromisoformat(text.split(".")[0])
        return d.time() if d.date() in (dt.date(1899, 12, 31), dt.date(1899, 12, 30)) else d
    if typ == "Boolean":
        return text.strip() in ("1", "true")
    return text


def read_spreadsheetml(data):
    root = ET.fromstring(data)
    ns = {"ss": SS}
    wb = _new_wb()
    for wsel in root.iter(f"{{{SS}}}Worksheet"):
        rows = []
        table = wsel.find("ss:Table", ns)
        if table is not None:
            for rowel in table.findall("ss:Row", ns):
                ridx = rowel.get(f"{{{SS}}}Index")
                if ridx:
                    while len(rows) < int(ridx) - 1:
                        rows.append([])
                row = []
                for cel in rowel.findall("ss:Cell", ns):
                    cidx = cel.get(f"{{{SS}}}Index")
                    if cidx:
                        while len(row) < int(cidx) - 1:
                            row.append(None)
                    d = cel.find("ss:Data", ns)
                    row.append(None if d is None else _ss_value(d.get(f"{{{SS}}}Type"), "".join(d.itertext())))
                rows.append(row)
        _add_sheet(wb, wsel.get(f"{{{SS}}}Name"), rows)
    if not wb.sheetnames:
        raise ReadError("В XML нет листов Excel")
    return wb


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def read_generic_xml(data):
    """Любой «табличный» XML: ищем элемент, у которого больше всего одинаковых детей — это строки."""
    root = ET.fromstring(data)
    best, best_n = None, 1
    for el in root.iter():
        kids = list(el)
        if not kids:
            continue
        tags = [_local(k.tag) for k in kids]
        top = max(set(tags), key=tags.count)
        n = tags.count(top)
        if n > best_n:
            best, best_n, best_tag = el, n, top
    if best is None:
        raise ReadError("Не удалось найти в XML повторяющиеся записи")
    records = [k for k in best if _local(k.tag) == best_tag]
    cols, rows = [], []
    for rec in records:
        row = {}
        for k, v in rec.attrib.items():
            row[_local(k)] = v
        for ch in rec:
            row[_local(ch.tag)] = ("".join(ch.itertext()).strip() or None) if len(ch) == 0 else \
                " ".join(t.strip() for t in ch.itertext() if t.strip()) or None
        if not row and (rec.text or "").strip():
            row[best_tag] = rec.text.strip()
        for k in row:
            if k not in cols:
                cols.append(k)
        rows.append(row)
    wb = _new_wb()
    _add_sheet(wb, best_tag, [cols] + [[r.get(c) for c in cols] for r in rows])
    return wb


class _Tables(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables, self._row, self._cell, self._span = [], None, None, 1

    def _close_cell(self):
        if self._cell is not None and self._row is not None:
            txt = re.sub(r"[ \t\r\f\v]+", " ", "".join(self._cell)).strip()
            self._row.append(txt or None)
            self._row.extend([None] * (self._span - 1))
        self._cell = None

    def _close_row(self):
        self._close_cell()
        if self._row is not None and self.tables:
            self.tables[-1].append(self._row)
        self._row = None

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.tables.append([])
        elif tag == "tr" and self.tables:
            self._close_row()                      # незакрытая предыдущая строка
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._close_cell()                     # <td>текст<td> — незакрытая ячейка, как в браузере
            self._cell = []
            try:
                self._span = max(1, int(dict(attrs).get("colspan") or 1))
            except ValueError:
                self._span = 1
        elif tag == "br" and self._cell is not None:
            self._cell.append("\n")

    def handle_endtag(self, tag):
        if tag in ("td", "th"):
            self._close_cell()                     # лишний </td> без открытой ячейки — игнорируется
        elif tag == "tr":
            self._close_row()
        elif tag == "table":
            self._close_row()

    def handle_data(self, d):
        if self._cell is not None:
            self._cell.append(d)


def read_html(data):
    p = _Tables()
    p.feed(_decode(data))
    tables = [t for t in p.tables if t]
    if not tables:
        raise ReadError("В HTML нет таблиц")
    wb = _new_wb()
    for i, t in enumerate(tables, 1):
        _add_sheet(wb, f"Table{i}", t)
    return wb


ODS = {"table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
       "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
       "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0"}


def _ods_value(cell):
    o = ODS["office"]
    typ = cell.get(f"{{{o}}}value-type")
    if typ in ("float", "percentage", "currency"):
        f = float(cell.get(f"{{{o}}}value"))
        return int(f) if f.is_integer() else f
    if typ == "date":
        v = cell.get(f"{{{o}}}date-value")
        return dt.datetime.fromisoformat(v) if "T" in v else dt.datetime.fromisoformat(v + "T00:00:00")
    if typ == "time":
        m = re.match(r"PT(\d+)H(\d+)M([\d.]+)S", cell.get(f"{{{o}}}time-value") or "")
        return dt.time(int(m.group(1)) % 24, int(m.group(2)), int(float(m.group(3)))) if m else None
    if typ == "boolean":
        return cell.get(f"{{{o}}}boolean-value") == "true"
    paras = cell.findall(f"{{{ODS['text']}}}p")
    txt = "\n".join("".join(p.itertext()) for p in paras)
    return txt or None


def read_ods(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        root = ET.fromstring(z.read("content.xml"))
    t = ODS["table"]
    wb = _new_wb()
    for tab in root.iter(f"{{{t}}}table"):
        rows = []
        for rowel in tab.iter(f"{{{t}}}table-row"):
            row = []
            for cell in rowel:
                if _local(cell.tag) not in ("table-cell", "covered-table-cell"):
                    continue
                rep = int(cell.get(f"{{{t}}}number-columns-repeated", "1"))
                v = _ods_value(cell)
                if v is None and rep > MAX_EMPTY_REPEAT:
                    break
                row.extend([v] * min(rep, MAX_EMPTY_REPEAT))
            while row and row[-1] is None:
                row.pop()
            rrep = int(rowel.get(f"{{{t}}}number-rows-repeated", "1"))
            if not row and rrep > MAX_EMPTY_REPEAT:
                continue
            rows.extend([row] * min(rrep, MAX_EMPTY_REPEAT))
        while rows and not rows[-1]:
            rows.pop()
        _add_sheet(wb, tab.get(f"{{{t}}}name"), rows)
    return wb


READERS = {"xlsx": read_xlsx, "xls": read_xls, "csv": read_csv, "spreadsheetml": read_spreadsheetml,
           "xml": read_generic_xml, "html": read_html, "ods": read_ods}


def load_any(path) -> tuple[openpyxl.Workbook, str]:
    """Возвращает (книга, найденный формат)."""
    data = Path(path).read_bytes()
    if not data:
        raise ReadError("Файл пустой")
    kind = sniff(data)
    try:
        wb = READERS[kind](data)
    except ReadError:
        raise
    except Exception as e:  # noqa: BLE001
        raise ReadError(f"Не удалось прочитать файл как {kind.upper()}: {e}") from e
    if not wb.sheetnames:
        raise ReadError("В файле нет данных")
    return wb, kind

"""Одни и те же данные EPG в разных форматах дают одинаковый результат."""
import datetime as dt
import io
import zipfile

import openpyxl
import pytest

from epg_processor import process
from readers import ReadError, load_any, sniff

HEAD = ["Дата", "Время", "Unnamed: 2", "Название (рус)", "Название эпизода", "Формат", "Возрастное ограничение"]
ROWS = [["09.11.2026", "10:03", "x", "Фильм", "Фильм", "Movie", "12"],
        ["10.11.2026", "23:59", "x", "Сериал", "Серия 1", "Series", "18 years"]]
EXPECTED = [("09/11/2026", "10:00", "Фильм", None, "х/ф", "12+"),
            ("10/11/2026", "23:55", "Сериал", "Серия 1", "Series", "18+")]
NAME = "EPG PLAN B 09.11-15.11"


def run(tmp_path, ext, data):
    src = tmp_path / f"{NAME}{ext}"
    src.write_bytes(data)
    out, log = process(src, out_dir=tmp_path / "out")
    assert out.suffix == ".xlsx"
    rows = list(openpyxl.load_workbook(out).active.iter_rows(values_only=True))
    norm = []
    for r in rows[1:]:
        d, t = r[0], r[1]
        d = d.strftime("%d/%m/%Y") if isinstance(d, dt.datetime) else d
        t = t.strftime("%H:%M") if isinstance(t, dt.time) else t
        norm.append((d, t) + tuple(r[2:]))
    assert rows[0] == tuple(h for h in HEAD if not h.startswith("Unnamed"))
    assert norm == EXPECTED
    return log


def csv_bytes(delim=";", enc="utf-8"):
    lines = [delim.join(HEAD)] + [delim.join(r) for r in ROWS]
    return ("\r\n".join(lines) + "\r\n").encode(enc)


@pytest.mark.parametrize("delim,enc", [(";", "utf-8"), (",", "utf-8-sig"), ("\t", "cp1251"), ("|", "utf-8")])
def test_csv(tmp_path, delim, enc):
    run(tmp_path, ".csv", csv_bytes(delim, enc))


def test_xlsx(tmp_path):
    wb = openpyxl.Workbook()
    for r in [HEAD] + ROWS:
        wb.active.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    run(tmp_path, ".xlsx", buf.getvalue())


def test_xls_binary(tmp_path):
    import xlwt
    wb = xlwt.Workbook()
    ws = wb.add_sheet("EPG")
    datefmt, timefmt = xlwt.easyxf(num_format_str="DD.MM.YYYY"), xlwt.easyxf(num_format_str="HH:MM")
    for c, h in enumerate(HEAD):
        ws.write(0, c, h)
    data = [[dt.datetime(2026, 11, 9), dt.time(10, 3), "x", "Фильм", "Фильм", "Movie", 12],
            [dt.datetime(2026, 11, 10), dt.time(23, 59), "x", "Сериал", "Серия 1", "Series", "18 years"]]
    for r, row in enumerate(data, 1):
        for c, v in enumerate(row):
            ws.write(r, c, v, datefmt if c == 0 else timefmt if c == 1 else xlwt.Style.default_style)
    buf = io.BytesIO()
    wb.save(buf)
    run(tmp_path, ".xls", buf.getvalue())


def test_spreadsheetml(tmp_path):
    def cell(v):
        return f'<Cell><Data ss:Type="String">{v}</Data></Cell>'
    rows = "".join("<Row>" + "".join(cell(v) for v in r) + "</Row>" for r in [HEAD] + ROWS)
    xml = ('<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>'
           '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" '
           'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
           f'<Worksheet ss:Name="EPG"><Table>{rows}</Table></Worksheet></Workbook>')
    assert sniff(xml.encode()) == "spreadsheetml"
    run(tmp_path, ".xml", xml.encode())


def test_spreadsheetml_types_and_index():
    xml = ('<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" '
           'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="A"><Table>'
           '<Row><Cell><Data ss:Type="DateTime">2026-11-09T00:00:00.000</Data></Cell>'
           '<Cell ss:Index="3"><Data ss:Type="DateTime">1899-12-31T10:03:00.000</Data></Cell>'
           '<Cell><Data ss:Type="Number">16</Data></Cell></Row></Table></Worksheet></Workbook>')
    import tempfile, pathlib
    p = pathlib.Path(tempfile.mkdtemp()) / "a.xml"
    p.write_text(xml)
    wb, kind = load_any(p)
    assert kind == "spreadsheetml"
    assert list(wb.active.iter_rows(values_only=True))[0] == (dt.datetime(2026, 11, 9), None, dt.time(10, 3), 16)


def test_generic_xml(tmp_path):
    tags = ["date", "time", "u", "title", "episode", "format", "age"]
    items = "".join("<program>" + "".join(f"<{t}>{v}</{t}>" for t, v in zip(tags, r)) + "</program>" for r in ROWS)
    xml = f'<?xml version="1.0" encoding="UTF-8"?><epg channel="PLAN B">{items}</epg>'
    src = tmp_path / "x.xml"
    src.write_text(xml, encoding="utf-8")
    wb, kind = load_any(src)
    assert kind == "xml"
    rows = list(wb.active.iter_rows(values_only=True))
    assert rows[0] == tuple(tags) and rows[1][3] == "Фильм" and len(rows) == 3


def test_html_disguised_as_xls(tmp_path):
    th = "".join(f"<th>{h}</th>" for h in HEAD)
    trs = "".join("<tr>" + "".join(f"<td>{v}</td>" for v in r) + "</tr>" for r in ROWS)
    html = f'<html><head><meta charset="utf-8"></head><body><table><tr>{th}</tr>{trs}</table></body></html>'
    assert sniff(html.encode()) == "html"
    run(tmp_path, ".xls", html.encode())


def test_ods(tmp_path):
    t, o, x = ("urn:oasis:names:tc:opendocument:xmlns:table:1.0", "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
               "urn:oasis:names:tc:opendocument:xmlns:text:1.0")

    def c(v):
        return f'<table:table-cell office:value-type="string"><text:p>{v}</text:p></table:table-cell>'
    rows = "".join("<table:table-row>" + "".join(c(v) for v in r) +
                   '<table:table-cell table:number-columns-repeated="16377"/></table:table-row>' for r in [HEAD] + ROWS)
    rows += '<table:table-row table:number-rows-repeated="1048000"><table:table-cell/></table:table-row>'
    content = (f'<?xml version="1.0"?><office:document-content xmlns:office="{o}" xmlns:table="{t}" xmlns:text="{x}">'
               f'<office:body><office:spreadsheet><table:table table:name="EPG">{rows}</table:table>'
               '</office:spreadsheet></office:body></office:document-content>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
        z.writestr("content.xml", content)
    run(tmp_path, ".ods", buf.getvalue())


def test_garbage(tmp_path):
    p = tmp_path / "a.xlsx"
    p.write_bytes(b"PK\x03\x04garbage")
    with pytest.raises(ReadError):
        load_any(p)
    p.write_bytes(b"")
    with pytest.raises(ReadError):
        load_any(p)

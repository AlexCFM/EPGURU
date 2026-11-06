import datetime as dt

import openpyxl

from fix_dates import fix_workbook

D = dt.datetime


def test_fix_dates(tmp_path):
    src = tmp_path / "EPG PLAN B 02.11-08.11.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in [["Дата", "Время", "Эфир"],
              [D(2026, 2, 11), "10:00", "11.02 вместо 02.11"],   # перепутано Excel'ем
              ["05.11.2026", "10:00", "текст"],
              ["11/06/2026", "10:00", "американский"],
              [46333, "10:00", "серийный номер = 07.11.2026"],
              [D(2026, 11, 8), "10:00", "нормальная"],
              ["мусор", "10:00", "не дата"]]:
        ws.append(r)
    wb.save(src)
    out = tmp_path / "out.xlsx"
    s = fix_workbook(src, out)
    vals = [c.value for c in openpyxl.load_workbook(out).active["A"][1:]]
    assert vals == [D(2026, 11, 2), "05/11/2026", "06/11/2026", D(2026, 11, 7), D(2026, 11, 8), "мусор"]
    assert s["swapped"] == 1 and s["serial"] == 1 and s["text"] == 2 and len(s["bad"]) == 1

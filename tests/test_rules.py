import datetime as dt
from pathlib import Path

import openpyxl
import pytest

from epg_processor import EPGError, auto_template, detect_week, process

D = dt.datetime
HEAD = ("канал", "дата", "время", "формат", "название (рус)", "название (ориг)", "сезон", "номер серии",
        "название эпизода", "маркер", "год", "жанр", "страна", "описание (короткое)", "описание (полное)", "Актёры")


def make(path, rows):
    wb = openpyxl.Workbook()
    for r in rows:
        wb.active.append(r)
    wb.save(path)
    return path


def read(path):
    return list(openpyxl.load_workbook(path).active.iter_rows(values_only=True))


# ---------- автовыбор и имя файла ----------

@pytest.mark.parametrize("name,tpl", [
    ("EPG PLAN B 09.11-15.11.xlsx", "make_sense"),
    ("EPG 9-1-1 09_11-15_11.xlsx", "make_sense"),
    ("VESNA 09.11-15.11.xlsx", "make_sense"),
    ("FIGHT CLUB EPG 09_11 - 15_11.xlsx", "fight_club"),
])
def test_auto_template(name, tpl):
    assert auto_template(name).stem == tpl


def test_week_with_year_in_name():
    assert detect_week("PLAN B 09_11_2026-15_11_2026.xlsx", "PLAN B") == ("09", "11", "15", "11")


# ---------- MAKE SENSE ----------

def test_make_sense(tmp_path):
    src = make(tmp_path / "EPG 9-1-1 09.11-15.11.xlsx", [
        ["Дата", "Время", "Unnamed: 2", "Название (рус)", "Название эпизода", "Формат", "Возрастное ограничение"],
        [D(2026, 11, 9), dt.time(10, 3), "x", "Фильм", "Фильм ", "Movie", "12"],
        [D(2026, 11, 9), dt.time(10, 19, 45), "x", "Сериал", "Серия 1. Начало", "Series", "18 years"],
        ["10.11.2026", "23:59", "x", "Шоу", "", "Movie ", "16+"],
    ])
    out, _ = process(src, out_dir=tmp_path / "out")
    assert out.name == "EPG 9-1-1 09.11 - 15.11.2026 (msk time).xlsx"
    rows = read(out)
    assert rows[0] == HEAD
    pick = lambda r: (r[1], r[2], r[4], r[8], r[3], r[9])      # дата, время, название, эпизод, формат, маркер
    assert all(r[0] == 911 for r in rows[1:])                   # столбец «канал» заполнен
    assert pick(rows[1])[1:] == (dt.time(10, 0), "Фильм", None, "х/ф", "12+")
    assert pick(rows[2])[1:] == (dt.time(10, 15), "Сериал", "Серия 1. Начало", "Series", "18+")
    assert pick(rows[3]) == ("10/11/2026", "23:55", "Шоу", None, "Movie ", "16+")   # "Movie " не точное совпадение


# ---------- FIGHT CLUB ----------

FC_HEAD = ["Date", "Start Time", "EPG_TITLE_RUS", "AGE RATING", "Note"]


def test_fight_club(tmp_path):
    src = make(tmp_path / "FIGHT CLUB EPG 09_11 - 15_11.xlsx", [
        FC_HEAD,
        [D(2026, 11, 8), "23:10", "Хвост", "16", "n"],          # дата из файла сохраняется
        ["09.11.2026", "00:05", "Программа А", 16, "n"],
        [D(2026, 11, 9), "00:30", "Программа А", 16, "n"],      # дубль -> удалить
        [D(2026, 11, 9), "01:00", "Программа Б", "18+", "n"],
        [D(2026, 11, 9), "02:00", "Программа А", 16, "n"],      # новый блок -> оставить
        [None, None, None, None, None],                         # пустая строка остаётся, граница
        [D(2026, 11, 9), "03:00", "Программа А", 16, "n"],      # после пустой -> оставить
        [D(2026, 11, 9), "04:00", " МИСС ВСЕЛЕННАЯ ", None, "n"],
    ])
    out, _ = process(src, out_dir=tmp_path / "out")
    assert out.name == "FIGHT CLUB EPG  09_11 - 15_11.xlsx"
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["FIGHT-CLUB"]
    rows = read(out)
    assert rows[0][:4] == ("Канал ", "Дата", "Время", "Название") and len(rows[0]) == 15
    assert len(rows) == 8
    assert [r[1] for r in rows[1:]] == [D(2026, 11, 8), "09/11/2026", D(2026, 11, 9), D(2026, 11, 9),
                                        None, D(2026, 11, 9), D(2026, 11, 9)]
    assert [r[3] for r in rows[1:]] == ["Хвост", "Программа А", "Программа Б", "Программа А", None,
                                        "Программа А", "ЗАГЛУШКА"]
    assert [r[9] for r in rows[1:]] == ["16+", "16+", "18+", "16+", None, "16+", None]
    assert rows[1][0] == "FIGHT CLUB" and rows[5][0] is None            # пустая строка осталась пустой
    assert rows[1][2] == dt.time(23, 10)                                  # время не округляется, тип «время»


def test_fight_club_csv_export(tmp_path):
    """CSV-выгрузка FIGHT-CLUB_20261116-20261122.csv: даты MM/DD/YYYY, куски одного боя схлопываются."""
    head = ["Date", "Start Time", "EPG_TITLE_RUS", "Episode Name", "EPISODE_NAME_2", "SEASON No.", "EPISODE No.",
            "TITLE_EN", "AGE RATING", "GENRE 1", "GENRE 2", "PROD_YEAR", "SHORT_DESCRIPTION", "LONG_DESCRIPTION"]
    fight = ["Лига EFC 135", "БОЙ ГОДА", "УОЛКЕР VS ДЖЕЙКОБС", "", "135", "EFC 135", "18", "Спорт", "Единоборства",
             "2026", "Коротко", "Длинно"]
    lines = [head, ["11/15/2026", "23:33"] + fight, ["11/16/2026", "00:04"] + fight, ["11/16/2026", "00:21"] + fight,
             ["11/16/2026", "01:58", "Лига EFC 132", "БОЙ", "А VS Б", "1", "132", "EFC 132", "18", "Спорт", "MMA",
              "2025", "К", "Д"]]
    import csv as _csv, io as _io
    buf = _io.StringIO()
    _csv.writer(buf).writerows(lines)
    src = tmp_path / "FIGHT-CLUB_20261116-20261122.csv"
    src.write_text("\ufeff" + buf.getvalue(), encoding="utf-8")
    out, log = process(src, out_dir=tmp_path / "out")
    assert out.name == "FIGHT CLUB EPG  16_11 - 22_11.xlsx"
    rows = read(out)[1:]
    assert [(r[1], r[2], r[3], r[6], r[7], r[9], r[12]) for r in rows] == [
        ("15/11/2026", dt.time(23, 33), "Лига EFC 135", None, 135, "18+", 2026),
        ("16/11/2026", dt.time(1, 58), "Лига EFC 132", 1, 132, "18+", 2025)]


def test_dates_any_input_format(tmp_path):
    src = make(tmp_path / "VESNA 09.11-15.11.xlsx", [
        ["Дата", "Время", "Формат"],
        ["11/15/2026", "10:00", "x"],      # американский, месяц первым
        ["11/09/2026", "10:00", "x"],      # тот же столбец -> тоже месяц первым = 9 ноября
        ["15.11.2026", "10:00", "x"],      # русский
        ["09.11.2026", "10:00", "x"],
        [46335, "10:00", "x"],             # серийный номер Excel = 09.11.2026
        [D(2026, 11, 9), "10:00", "x"],
    ])
    out, _ = process(src, out_dir=tmp_path)
    ws = openpyxl.load_workbook(out).active
    assert [ws.cell(r, 2).value for r in range(2, 6)] == ["15/11/2026", "09/11/2026", "15/11/2026", "09/11/2026"]
    assert ws.cell(6, 2).value == D(2026, 11, 9) and ws.cell(7, 2).value == D(2026, 11, 9)
    assert ws.cell(7, 2).number_format == r"DD\/MM\/YYYY"


def test_ambiguous_dates_checked_against_week(tmp_path):
    src = make(tmp_path / "VESNA 02.11-08.11.xlsx", [
        ["Дата", "Время", "Формат"],
        ["02/11/2026", "10:00", "x"], ["08/11/2026", "10:00", "x"],   # день первым: всё ноябрь
    ])
    out, _ = process(src, out_dir=tmp_path)
    assert [r[1] for r in read(out)[1:]] == ["02/11/2026", "08/11/2026"]

    src = make(tmp_path / "PLAN B 02.11-08.11.xlsx", [
        ["Дата", "Время", "Формат"], ["11/02/2026", "10:00", "x"]])  # месяц первым
    out, _ = process(src, out_dir=tmp_path)
    assert read(out)[1][1] == "02/11/2026"


def test_title_row_above_header(tmp_path):
    """Над шапкой строка «EPG 9-1-1 …» — шапка находится, заголовок отбрасывается."""
    src = make(tmp_path / "EPG 9-1-1 09.11-15.11.xlsx", [
        ["EPG 9-1-1  09.11 - 15.11.2026 msk time", None, None],
        [None, None, None],
        ["Дата", "Время", "Название (рус)", "Название эпизода", "Формат", "Маркер"],
        [D(2026, 11, 9), dt.time(10, 3), "Фильм", "1 серия", "Movie", "12"],
    ])
    out, log = process(src, out_dir=tmp_path / "out")
    rows = read(out)
    assert rows[0] == HEAD
    assert rows[1][4] == "Фильм" and rows[1][3] == "х/ф" and len(rows) == 2
    assert any("шапка найдена в строке 3" in l for l in log)


def test_prime_time_911(tmp_path):
    rows_in = [["Дата", "Время", "Название (рус)", "Формат"],
               ["09.11.2026", "21:27:40", "Дом Мэнсона", "Movie"],     # 21:25 -> 21:30
               ["10.11.2026", "21:30:10", "Пассаж", "Movie"],          # уже 21:30
               ["10.11.2026", "21:57:00", "Сило", "Movie"],            # прайм уже есть -> не трогаем
               ["11.11.2026", "22:06:00", "Лестэр", "Movie"],          # 22:05 -> 22:00
               ["12.11.2026", "21:40:00", "Эпплтон", "Movie"]]         # дальше 5 минут -> не трогаем
    src = make(tmp_path / "EPG 9-1-1 09.11-15.11.xlsx", rows_in)
    out, _ = process(src, out_dir=tmp_path / "out")
    assert [r[2] for r in read(out)[1:]] == ["21:30", "21:30", "21:55", "22:00", "21:40"]
    src = make(tmp_path / "EPG VESNA 09.11-15.11.xlsx", rows_in)        # на VESNA правило не действует
    out, _ = process(src, out_dir=tmp_path / "out2")
    assert read(out)[1][2] == "21:25"

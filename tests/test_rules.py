import datetime as dt
from pathlib import Path

import openpyxl
import pytest

from epg_processor import EPGError, auto_template, detect_week, process

D = dt.datetime


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
        [D(2026, 11, 9), dt.time(10, 19, 45), "x", "Сериал", "Серия 1", "Series", "18 years"],
        ["10.11.2026", "23:59", "x", "Шоу", "", "Movie ", "16+"],
    ])
    out, _ = process(src, out_dir=tmp_path / "out")
    assert out.name == "EPG_911_09_11_15_11_MAKE_SENSE.xlsx"
    rows = read(out)
    assert rows[0] == ("Дата", "Время", "Название (рус)", "Название эпизода", "Формат", "Возрастное ограничение")
    assert rows[1][1:] == (dt.time(10, 0), "Фильм", None, "х/ф", "12+")
    assert rows[2][1:] == (dt.time(10, 15), "Сериал", "Серия 1", "Series", "18+")
    assert rows[3] == ("10/11/2026", "23:55", "Шоу", None, "Movie ", "16+")   # "Movie " не точное совпадение


# ---------- FIGHT CLUB ----------

FC_HEAD = ["Date", "Start Time", "EPG_TITLE_RUS", "AGE RATING", "Note"]


def test_fight_club(tmp_path):
    src = make(tmp_path / "FIGHT CLUB EPG 09_11 - 15_11.xlsx", [
        FC_HEAD,
        [None, "20:00", "Программа А", 16, "n"],
        [None, "21:30", "Программа А", 16, "n"],      # дубль -> удалить
        [None, "23:55", "Программа А", 16, "n"],      # дубль -> удалить
        [None, "00:30", "Программа Б", "18+", "n"],   # переход через полночь
        [None, "02:00", "Программа А", 16, "n"],      # новый блок -> оставить
        [None, None, None, None, None],               # пустая строка остаётся, граница
        [None, "03:00", "Программа А", 16, "n"],      # после пустой -> оставить
        [None, "04:00", " МИСС ВСЕЛЕННАЯ ", None, "n"],
    ])
    out, _ = process(src, out_dir=tmp_path / "out", year=2026)
    assert out.name == "EPG_FIGHT_CLUB_09_11_15_11_processed.xlsx"
    rows = read(out)
    assert len(rows) == 7
    dates = [r[0] for r in rows[1:]]
    assert dates == [D(2026, 11, 9), D(2026, 11, 10), D(2026, 11, 10), None, D(2026, 11, 10), D(2026, 11, 10)]
    assert [r[2] for r in rows[1:]] == ["Программа А", "Программа Б", "Программа А", None, "Программа А", "ЗАГЛУШКА"]
    assert [r[3] for r in rows[1:]] == ["16+", "18+", "16+", None, "16+", None]

    out2, _ = process(src, out_dir=tmp_path / "out", year=2026)
    assert out2.name == "EPG_FIGHT_CLUB_09_11_15_11_processed_v2.xlsx"


def test_fight_club_year_from_data(tmp_path):
    src = make(tmp_path / "FIGHT CLUB EPG 09_11 - 15_11.xlsx", [
        FC_HEAD, [D(2027, 1, 1), "20:00", "А", 16, "n"]])
    out, _ = process(src, out_dir=tmp_path)
    assert read(out)[1][0] == D(2027, 11, 9)


def test_fight_club_no_year(tmp_path):
    src = make(tmp_path / "FIGHT CLUB EPG 09_11 - 15_11.xlsx", [FC_HEAD, [None, "20:00", "А", 16, "n"]])
    with pytest.raises(EPGError, match="год"):
        process(src, out_dir=tmp_path)


def test_fight_club_previous_day_tail(tmp_path):
    src = make(tmp_path / "FIGHT CLUB EPG 09_11 - 15_11.xlsx", [
        FC_HEAD,
        [None, "23:10", "Хвост", 16, "n"],
        [None, "23:50", "Хвост 2", 16, "n"],
        [None, "00:05", "Первая программа", 16, "n"],
        [None, "12:00", "Днём", 16, "n"],
        [None, "01:00", "Ночь", 16, "n"],
    ])
    out, _ = process(src, out_dir=tmp_path, year=2026)
    assert [r[0] for r in read(out)[1:]] == [
        D(2026, 11, 8), D(2026, 11, 8), D(2026, 11, 9), D(2026, 11, 9), D(2026, 11, 10)]


def test_fight_club_normal_start_not_shifted(tmp_path):
    src = make(tmp_path / "FIGHT CLUB EPG 09_11 - 15_11.xlsx", [
        FC_HEAD, [None, "06:00", "Утро", 16, "n"], [None, "23:30", "Вечер", 16, "n"]])
    out, _ = process(src, out_dir=tmp_path, year=2026)
    assert [r[0] for r in read(out)[1:]] == [D(2026, 11, 9), D(2026, 11, 9)]

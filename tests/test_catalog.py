import datetime as dt
import importlib
import io

import openpyxl
import pytest
import yaml
from fastapi.testclient import TestClient

from catalog import is_bare_episode, load_catalog, norm_title
from epg_processor import TEMPLATES_DIR, process

SPEC = yaml.safe_load(open(TEMPLATES_DIR / "make_sense.yaml", encoding="utf-8"))["catalogs"]


def book(sheets: dict) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


SUSPENSE = book({
    "FILMS": [["Name RU MOVIES", None], ["Девушка в поезде", 6.5], ["Пассаж", 6]],
    "SERIES": [[None], ["Name RU SERIES"], ["Береговая охранa"], ["Сило "], ["Полиция Вены: Спецотдел (12 и 13 сезоны)"]],
    "ODDS": [[None], ["Name RU ODDS"], ["Каина"]],
})
VESNA = book({
    "CONTENT ": [["Series Name (EN)", "Название (RU)"], ["12 Days", "12 дней подарков"]],
    "СЕРИАЛЫ ": [["Название (RU)", "Season", "Episode"], ["Коварная любовь", 1, 1], ["Коварная любовь", 1, 2]],
    "РУССКОЕ": [["Название (RU)", "Episode", "Duration"],
                ["Город Ромашек", 1, "00:43:58:04"], ["Город Ромашек", 2, "00:44:10:00"],
                ["Город Ромашек", 3, "00:44:00:00"], ["Город Ромашек", 4, "00:47:59:00"],     # 4×45 — фильм
                ["Пенелопа", 1, "01:28:00:00"], ["Пенелопа", 2, "01:28:00:00"],
                ["Пенелопа", 3, "01:28:00:00"], ["Пенелопа", 4, "01:28:00:00"],               # 4×88 — сериал
                ["Каждый за себя", 1, "01:33:00:00"], ["Каждый за себя", 2, "01:33:00:00"],
                ["Каждый за себя", 3, "01:33:00:00"],                                         # 3×93 — сериал
                ["Бабье лето", 1, "01:26:00:00"], ["Бабье лето", 2, "01:26:00:00"],           # 2×86 — фильм
                ["Наследница", 1, "00:52:00:00"], ["Наследница", 2, "00:52:00:00"],
                ["Наследница", 3, "00:52:00:00"], ["Наследница", 4, "00:52:00:00"],           # 4×52 — фильм
                ["Белое платье", 1, "01:30:00:00"]] +
               [["Длинный", i, "00:44:00:00"] for i in range(1, 7)],                         # 6 серий — сериал
})


@pytest.mark.parametrize("v,bare", [
    ("1 серия", True), ("Серия 2", True), ("Эпизод 3", True), ("2-я серия", True), ("Серия №5", True),
    ("Сезон 1. Серия 4", True), ("S01E05", True), ("Episode 7", True), ("  серия 12 ", True),
    ("Эпизод 3. Татьяна", False), ("Серия 1 (финал)", False), ("Татьяна", False), ("Серия", False),
    (None, False), ("", False), (3, False),
])
def test_bare_episode(v, bare):
    assert is_bare_episode(v) is bare


def test_norm_title_homoglyph_and_parens(tmp_path):
    p = tmp_path / "911.xlsx"
    p.write_bytes(SUSPENSE)
    cat = load_catalog(p, SPEC["911"])
    assert cat.lookup("Береговая охрана")["type"] == "series"      # латинская «a» в каталоге
    assert cat.lookup("СИЛО")["type"] == "series"
    assert cat.lookup("Полиция Вены: Спецотдел")["type"] == "series"
    assert cat.lookup("Девушка в поезде")["type"] == "film"
    assert cat.lookup("Каина")["type"] == "film"            # ODDS — фильмы
    assert cat.lookup("Неизвестный фильм") is None
    assert norm_title("Ёлка «Новая»") == norm_title("елка Новая")


def test_vesna_auto_sheet(tmp_path):
    p = tmp_path / "VESNA.xlsx"
    p.write_bytes(VESNA)
    cat = load_catalog(p, SPEC["VESNA"])
    assert cat.lookup("Город Ромашек")["type"] == "film"
    assert cat.lookup("Пенелопа")["type"] == "series"
    assert cat.lookup("Каждый за себя")["type"] == "series"
    assert cat.lookup("Бабье лето")["type"] == "film"
    assert cat.lookup("Наследница")["type"] == "film"
    assert cat.lookup("Длинный")["type"] == "series"
    assert cat.lookup("Белое платье")["type"] == "film"
    assert cat.lookup("Коварная любовь")["type"] == "series"
    assert cat.lookup("12 дней подарков")["type"] == "film"


def test_missing_sheet_error(tmp_path):
    p = tmp_path / "911.xlsx"
    p.write_bytes(book({"Лист1": [["что-то"]]}))
    with pytest.raises(ValueError, match="FILMS"):
        load_catalog(p, SPEC["911"])


EPG_HEAD = ["Дата", "Время", "Название (рус)", "Название эпизода", "Формат", "Возрастное ограничение"]


def run_epg(tmp_path, catalogs=True):
    cdir = tmp_path / "cat"
    cdir.mkdir()
    if catalogs:
        (cdir / "911.xlsx").write_bytes(SUSPENSE)
    wb = openpyxl.Workbook()
    for r in [EPG_HEAD,
              [dt.datetime(2026, 11, 9), dt.time(10, 3), "Береговая охрана", "1 серия", "Movie", "16"],
              [dt.datetime(2026, 11, 9), dt.time(11, 0), "Береговая охрана", "Эпизод 3. Татьяна", "Movie", "16"],
              [dt.datetime(2026, 11, 9), dt.time(12, 0), "Девушка в поезде", "Девушка в поезде", "Series", "18"],
              [dt.datetime(2026, 11, 9), dt.time(14, 0), "Каина", "Серия 2", "Movie", "16"],
              [dt.datetime(2026, 11, 9), dt.time(16, 0), "Девушка в поездe", None, "Movie", "18"],
              [dt.datetime(2026, 11, 9), dt.time(18, 0), "Совсем новый фильм", None, "Movie", "18"]]:
        wb.active.append(r)
    src = tmp_path / "EPG 9-1-1 09.11-15.11.xlsx"
    wb.save(src)
    out, log = process(src, out_dir=tmp_path / "out", catalog_dir=cdir)
    return list(openpyxl.load_workbook(out).active.iter_rows(values_only=True))[1:], log


def test_epg_with_catalog(tmp_path):
    rows, log = run_epg(tmp_path)
    assert [(r[4], r[8], r[3]) for r in rows] == [
        ("Береговая охрана", None, "сериал"),                  # «1 серия» вычищено, сериал по каталогу
        ("Береговая охрана", "Эпизод 3. Татьяна", "сериал"),   # с названием — не трогаем
        ("Девушка в поезде", None, "х/ф"),                     # фильм по каталогу
        ("Каина", None, "х/ф"),                                # ODDS — фильмы
        ("Девушка в поездe", None, "х/ф"),                     # латинская «e» в EPG тоже находится
        ("Совсем новый фильм", None, "х/ф"),                   # нет в каталоге — по общему правилу
    ]
    text = "\n".join(log)
    assert "нет в каталоге: «Совсем новый фильм» ×1" in text
    assert "Каталог: загружен" in text


def test_epg_without_catalog(tmp_path):
    rows, log = run_epg(tmp_path, catalogs=False)
    assert rows[0][3] == "х/ф"          # без каталога — только Movie -> х/ф
    assert rows[0][8] is None           # чистка серий работает и без каталога
    assert "Каталог: НЕ ЗАГРУЖЕН" in "\n".join(log)


@pytest.fixture()
def client(tmp_path, monkeypatch):
    monkeypatch.setenv("EPGURU_DATA", str(tmp_path / "data"))
    monkeypatch.setenv("EPGURU_PASSWORD", "secret")
    monkeypatch.setenv("EPGURU_SECURE_COOKIE", "0")
    import server.main as m
    importlib.reload(m)
    with TestClient(m.app) as c:
        c.post("/api/login", json={"password": "secret"})
        yield c


def test_catalog_api(client):
    assert {c["code"]: c["loaded"] for c in client.get("/api/catalogs").json()} == {"911": False, "VESNA": False, "PLAN_B": False, "FIGHT_CLUB": False}
    bad = client.post("/api/catalogs/911", files={"file": ("x.xlsx", book({"A": [["b"]]}))})
    assert bad.status_code == 400
    ok = client.post("/api/catalogs/911", files={"file": ("INVENTORY_SUSPENSE_RUS.xlsx", SUSPENSE)}).json()
    assert ok["films"] == 3 and ok["series"] == 3
    st = {c["code"]: c for c in client.get("/api/catalogs").json()}
    assert st["911"]["loaded"] and st["911"]["file"] == "INVENTORY_SUSPENSE_RUS.xlsx"
    assert client.post("/api/catalogs/NOPE", files={"file": ("a.xlsx", SUSPENSE)}).status_code == 404

    wb = openpyxl.Workbook()
    for r in [EPG_HEAD, [dt.datetime(2026, 11, 9), dt.time(10, 3), "Сило", "Серия 1", "Movie", "16"]]:
        wb.active.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    job = client.post("/api/process", files=[("files", ("EPG 911 09.11-15.11.xlsx", buf.getvalue()))]).json()[0]
    assert job["status"] == "done", job
    out = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/jobs/{job['id']}/download").content)).active
    assert out.cell(2, 4).value == "сериал" and out.cell(2, 9).value is None


def test_column_typed_catalog(tmp_path):
    p = tmp_path / "PLAN_B.xlsx"
    p.write_bytes(book({"КАТАЛОГ": [["Название (RU)", "Тип", "Название (EN)"],
                                    ["Хлоя", "сериал", "Khloe"], ["Крик Хаоса", "фильм", "Cry Havoc"],
                                    ["ПРЯМОЙ ЭФИР", "?", ""], ["Без типа", None, ""]]}))
    cat = load_catalog(p, SPEC["PLAN_B"])
    assert cat.lookup("Хлоя")["type"] == "series"
    assert cat.lookup("Крик Хаоса")["type"] == "film"
    assert cat.lookup("ПРЯМОЙ ЭФИР") is None and cat.lookup("Без типа") is None


def test_descriptions_and_original_title(tmp_path):
    cdir = tmp_path / "cat"
    cdir.mkdir()
    (cdir / "VESNA.xlsx").write_bytes(book({
        "CONTENT ": [["Series Name (EN)", "Название (RU)", "Short  EPG RU", "Long  EPG RU"],
                     ["12 Days", "12 дней подарков", "Коротко про фильм.", "Длинно про фильм."]],
        "СЕРИАЛЫ ": [["Название (RU)", "Season", "Episode", "Base EPG RU", "Description EPG RU"],
                     ["Коварная любовь", 1, 1, "Серия один.", "Серия один подробно."],
                     ["Коварная любовь", 1, 2, "Серия два.", "Серия два подробно."]]}))
    wb = openpyxl.Workbook()
    head = ["Channel", "Date", "Time", "Film Format", "Title Rus", "Title Eng", "Season #", "Episode #",
            "Episode Name", "Age Restrict", "Year", "Movie Genre", "Country", "Description Short",
            "Description Long", "Actors"]
    for r in [head,
              ["Romance_RUS", "2026-11-02", "10:00", "Movie", "12 дней подарков", "12 Days", "1", "1", "", "16",
               "2017", "", "", "Эшли Джонс,Дэвид Блю,,Джей Коннолли", "", ""],
              ["Romance_RUS", "2026-11-02", "12:00", "Movie", "Коварная любовь", "Deceiving Love", "1", "2", "", "16",
               "2022", "", "", "", "", ""],
              ["Romance_RUS", "2026-11-02", "14:00", "Movie", "Белое платье", "Белое платье", "1", "1", "", "12",
               "2010", "", "", "", "", ""]]:
        wb.active.append(r)
    src = tmp_path / "EPG_2026-11-022026-11-09.xlsx"
    wb.save(src)
    out, log = process(src, out_dir=tmp_path / "out", catalog_dir=cdir)
    assert out.name == "EPG VESNA 02.11 - 08.11.2026 (msk time).xlsx"
    ws = openpyxl.load_workbook(out).active
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert rows[0][0] == "VESNA" and rows[0][13] == "Коротко про фильм." and rows[0][14] == "Длинно про фильм."
    assert rows[0][15] == "Эшли Джонс,Дэвид Блю,Джей Коннолли"            # актёры на месте, двойная запятая убрана
    assert rows[0][6] is None and rows[0][7] is None                     # фильм: без сезона и серии
    assert rows[1][3] == "сериал" and rows[1][13] == "Серия два."         # описание конкретной серии
    assert rows[2][5] is None                                            # ориг = рус -> пусто
    assert ws.auto_filter.ref is None

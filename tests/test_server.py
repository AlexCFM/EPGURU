import datetime as dt
import importlib
import io

import openpyxl
import pytest
from fastapi.testclient import TestClient


@pytest.fixture()
def client(tmp_path, monkeypatch):
    monkeypatch.setenv("EPGURU_DATA", str(tmp_path / "data"))
    monkeypatch.setenv("EPGURU_PASSWORD", "secret")
    monkeypatch.setenv("EPGURU_SECURE_COOKIE", "0")
    import server.main as m
    importlib.reload(m)
    with TestClient(m.app) as c:
        yield c


def xlsx(rows) -> bytes:
    wb = openpyxl.Workbook()
    for r in rows:
        wb.active.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


MS = [["Дата", "Время", "Unnamed: 2", "Название (рус)", "Название эпизода", "Формат", "Возрастное ограничение"],
      [dt.datetime(2026, 11, 9), dt.time(10, 3), "x", "Фильм", "Фильм", "Movie", "12"]]


def login(c):
    assert c.post("/api/login", json={"password": "secret"}).status_code == 200


def test_auth_required(client):
    assert client.get("/api/jobs").status_code == 401
    assert client.get("/", follow_redirects=False).status_code in (302, 307)
    assert client.get("/login").status_code == 200
    assert client.post("/api/login", json={"password": "wrong"}).status_code == 401


def test_process_and_download(client):
    login(client)
    r = client.post("/api/process", files=[("files", ("EPG PLAN B 09.11-15.11.xlsx", xlsx(MS)))])
    job = r.json()[0]
    assert job["status"] == "done", job
    assert job["out"] == "EPG_PLAN_B_09_11_15_11_MAKE_SENSE.xlsx"
    d = client.get(f"/api/jobs/{job['id']}/download")
    assert d.status_code == 200
    ws = openpyxl.load_workbook(io.BytesIO(d.content)).active
    assert ws.cell(2, 2).value == dt.time(10, 0) and ws.cell(2, 5).value == "х/ф"
    assert [j["id"] for j in client.get("/api/jobs").json()] == [job["id"]]


def test_manual_channel(client):
    login(client)
    r = client.post("/api/process", data={"channel": "VESNA"},
                    files=[("files", ("расписание 09.11-15.11.xlsx", xlsx(MS)))])
    assert r.json()[0]["out"] == "EPG_VESNA_09_11_15_11_MAKE_SENSE.xlsx"


def test_errors_are_recorded(client):
    login(client)
    r = client.post("/api/process", files=[("files", ("bad.xlsx", b"not excel")),
                                           ("files", ("unknown 09.11-15.11.xlsx", xlsx(MS)))])
    res = r.json()
    assert [j["status"] for j in res] == ["error", "error"]
    assert res[0]["error"] and "шаблон" in res[1]["error"]


def test_fix_dates_endpoint(client):
    login(client)
    r = client.post("/api/fix-dates", files=[("files", ("PLAN B 02.11-08.11.xlsx",
                                                         xlsx([["Дата"], ["05.11.2026"]])))])
    j = r.json()[0]
    assert j["status"] == "done" and j["out"] == "PLAN B 02.11-08.11_dates_fixed.xlsx"


def test_config(client):
    login(client)
    names = [c["name"] for c in client.get("/api/config").json()["channels"]]
    assert {"PLAN B", "9-1-1", "VESNA", "FIGHT CLUB"} <= set(names)

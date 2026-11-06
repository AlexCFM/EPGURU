"""EPGURU — веб-сервис.

Запуск:  EPGURU_PASSWORD=... uvicorn server.main:app --host 0.0.0.0 --port 8000
"""
from __future__ import annotations

import hmac
import json
import os
import secrets
import shutil
import sqlite3
import sys
import threading
import time
import uuid
from contextlib import asynccontextmanager, closing
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from itsdangerous import BadSignature, SignatureExpired, TimestampSigner

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from epg_processor import list_templates, load_template, process  # noqa: E402
from fix_dates import fix_workbook  # noqa: E402

DATA = Path(os.environ.get("EPGURU_DATA", ROOT / "data"))
PASSWORD = os.environ.get("EPGURU_PASSWORD", "")
SECRET = os.environ.get("EPGURU_SECRET") or secrets.token_hex(32)
SECURE_COOKIE = os.environ.get("EPGURU_SECURE_COOKIE", "1") == "1"
MAX_MB = int(os.environ.get("EPGURU_MAX_MB", "20"))
SESSION_DAYS = 30
COOKIE = "epguru_session"

signer = TimestampSigner(SECRET, salt="epguru")
JOBS = DATA / "jobs"
DB = DATA / "epguru.db"
db_lock = threading.Lock()


# ---------- база ----------

def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    JOBS.mkdir(parents=True, exist_ok=True)
    with closing(db()) as con, con:
        con.execute("""CREATE TABLE IF NOT EXISTS jobs(
            id TEXT PRIMARY KEY, created REAL, kind TEXT, src TEXT, out TEXT,
            template TEXT, channel TEXT, status TEXT, log TEXT, error TEXT)""")


def save_job(**j):
    with db_lock, closing(db()) as con, con:
        con.execute("INSERT INTO jobs VALUES(:id,:created,:kind,:src,:out,:template,:channel,:status,:log,:error)", j)


def job_public(r) -> dict:
    return {"id": r["id"], "created": r["created"], "kind": r["kind"], "src": r["src"], "out": r["out"],
            "template": r["template"], "channel": r["channel"], "status": r["status"],
            "log": json.loads(r["log"] or "[]"), "error": r["error"]}


# ---------- вход ----------

_fails: dict[str, list[float]] = {}


def authed(request: Request) -> bool:
    if not PASSWORD:
        return True
    token = request.cookies.get(COOKIE)
    if not token:
        return False
    try:
        signer.unsign(token, max_age=SESSION_DAYS * 86400)
        return True
    except (BadSignature, SignatureExpired):
        return False


@asynccontextmanager
async def lifespan(_app):
    init_db()
    if not PASSWORD:
        print("ВНИМАНИЕ: EPGURU_PASSWORD не задан — вход без пароля", file=sys.stderr)
    yield


app = FastAPI(title="EPGURU", docs_url=None, redoc_url=None, openapi_url=None, lifespan=lifespan)


@app.middleware("http")
async def guard(request: Request, call_next):
    path = request.url.path
    open_paths = path in ("/login", "/login.html", "/api/login", "/healthz") or path.startswith("/assets/")
    if not open_paths and not authed(request):
        if path.startswith("/api/"):
            return JSONResponse({"detail": "Нужен вход"}, status_code=401)
        return RedirectResponse("/login")
    resp = await call_next(request)
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["Referrer-Policy"] = "same-origin"
    return resp


@app.get("/healthz")
def healthz():
    return {"ok": True}


@app.post("/api/login")
async def login(request: Request):
    ip = request.client.host if request.client else "?"
    now = time.time()
    recent = [t for t in _fails.get(ip, []) if now - t < 600]
    if len(recent) >= 10:
        raise HTTPException(429, "Слишком много попыток, подождите 10 минут")
    body = await request.json()
    if not hmac.compare_digest(str(body.get("password", "")), PASSWORD):
        _fails[ip] = recent + [now]
        time.sleep(1)
        raise HTTPException(401, "Неверный пароль")
    _fails.pop(ip, None)
    resp = JSONResponse({"ok": True})
    resp.set_cookie(COOKIE, signer.sign(uuid.uuid4().hex).decode(), max_age=SESSION_DAYS * 86400,
                    httponly=True, samesite="lax", secure=SECURE_COOKIE)
    return resp


@app.post("/api/logout")
def logout():
    resp = JSONResponse({"ok": True})
    resp.delete_cookie(COOKIE)
    return resp


# ---------- данные для интерфейса ----------

@app.get("/api/config")
def config():
    channels, templates = [], []
    for p in list_templates():
        t = load_template(p)
        templates.append({"id": p.stem, "name": t["name"]})
        for ch in t["channels"]:
            channels.append({"code": ch["code"], "name": ch["match"][0], "template": t["name"]})
    return {"channels": channels, "templates": templates}


@app.get("/api/jobs")
def jobs(limit: int = 50):
    with closing(db()) as con:
        rows = con.execute("SELECT * FROM jobs ORDER BY created DESC LIMIT ?", (min(limit, 500),)).fetchall()
    return [job_public(r) for r in rows]


@app.get("/api/jobs/{job_id}/download")
def download(job_id: str):
    with closing(db()) as con:
        r = con.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
    if not r or r["status"] != "done":
        raise HTTPException(404, "Файл не найден")
    path = JOBS / r["id"] / "out" / r["out"]
    if not path.is_file():
        raise HTTPException(404, "Файл не найден")
    return FileResponse(path, filename=r["out"],
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------- обработка ----------

async def _store(upload: UploadFile, job_dir: Path) -> Path:
    name = upload.filename or ""
    if not name.lower().endswith(".xlsx"):
        raise ValueError("Нужен файл .xlsx")
    data = await upload.read(MAX_MB * 1024 * 1024 + 1)
    if len(data) > MAX_MB * 1024 * 1024:
        raise ValueError(f"Файл больше {MAX_MB} МБ")
    if not data.startswith(b"PK"):
        raise ValueError("Это не Excel-файл (.xlsx)")
    job_dir.mkdir(parents=True)
    src = job_dir / "in.xlsx"
    src.write_bytes(data)
    return src


def _run(kind, upload_name, job_dir, src, template, channel):
    job = {"id": job_dir.name, "created": time.time(), "kind": kind, "src": upload_name, "out": None,
           "template": None, "channel": channel, "status": "error", "log": "[]", "error": None}
    try:
        if kind == "epg":
            tpl = next((p for p in list_templates() if p.stem == template), None) if template else None
            if template and not tpl:
                raise ValueError("Неизвестный шаблон")
            out, log = process(src, tpl, job_dir / "out", original_name=upload_name, channel=channel or None)
            job["template"] = log[0].split("·")[0].replace("Шаблон:", "").strip()
        else:
            out = job_dir / "out" / f"{Path(upload_name).stem}_dates_fixed.xlsx"
            s = fix_workbook(src, out, original_name=upload_name)
            log = [f"текст {s['text']}, числа {s['serial']}, перепутанные день/месяц {s['swapped']}, формат {s['format']}"]
            log += [f"не распознано: {b}" for b in s["bad"][:200]]
            job["template"] = "ДАТЫ"
        job.update(out=out.name, status="done", log=json.dumps(log, ensure_ascii=False))
    except Exception as e:  # noqa: BLE001
        job["error"] = str(e)
    save_job(**job)
    return job


async def _handle(kind, files, template="", channel=""):
    results = []
    for f in files:
        job_dir = JOBS / uuid.uuid4().hex
        try:
            src = await _store(f, job_dir)
        except ValueError as e:
            j = {"id": job_dir.name, "created": time.time(), "kind": kind, "src": f.filename, "out": None,
                 "template": None, "channel": channel, "status": "error", "log": "[]", "error": str(e)}
            save_job(**j)
            results.append(j)
            continue
        from starlette.concurrency import run_in_threadpool
        results.append(await run_in_threadpool(_run, kind, f.filename, job_dir, src, template, channel))
    for r in results:
        r["log"] = json.loads(r["log"])
    return results


@app.post("/api/process")
async def api_process(files: list[UploadFile] = File(...), template: str = Form(""), channel: str = Form("")):
    return await _handle("epg", files, template, channel)


@app.post("/api/fix-dates")
async def api_fix_dates(files: list[UploadFile] = File(...)):
    return await _handle("dates", files)


# ---------- статика ----------

WEB = ROOT / "web"


@app.get("/login")
def login_page():
    return FileResponse(WEB / "login.html")


@app.get("/")
def index():
    return FileResponse(WEB / "index.html")


app.mount("/assets", StaticFiles(directory=WEB / "assets"), name="assets")


def cleanup(days: int):
    """Удалить задания старше N дней (для cron)."""
    cutoff = time.time() - days * 86400
    with db_lock, closing(db()) as con, con:
        old = [r["id"] for r in con.execute("SELECT id FROM jobs WHERE created < ?", (cutoff,))]
        con.execute("DELETE FROM jobs WHERE created < ?", (cutoff,))
    for i in old:
        shutil.rmtree(JOBS / i, ignore_errors=True)
    return len(old)

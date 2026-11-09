"""Каталог контента канала: какие названия — фильмы, какие — сериалы.

Структура каталога (какие листы что содержат) описана в шаблоне, в блоке `catalogs`.
Сам файл каталога загружается через сайт и хранится на сервере, в git не попадает.
"""
from __future__ import annotations

import difflib
import re
import unicodedata
from collections import Counter
from pathlib import Path

from readers import load_any

# латинские буквы, похожие на русские (встречаются в каталогах: «охранa» с латинской a)
_HOMO = str.maketrans({"a": "а", "c": "с", "e": "е", "o": "о", "p": "р", "x": "х", "y": "у",
                       "k": "к", "m": "м", "t": "т", "h": "н", "b": "в"})
_CYR = re.compile(r"[а-яё]")


def norm_title(s) -> str:
    """Ключ для сравнения названий: регистр, ё/е, похожие латинские буквы, кавычки, пунктуация, пробелы."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s)).lower().replace("ё", "е")
    s = " ".join(w.translate(_HOMO) if _CYR.search(w) else w for w in s.split())
    s = re.sub(r"[^\w\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _strip_parens(s: str) -> str:
    return re.sub(r"\s*\([^)]*\)\s*", " ", str(s)).strip()


class Catalog:
    def __init__(self):
        self.items: dict[str, dict] = {}   # ключ -> {"title", "type"}

    def add(self, title, typ):
        if title is None or not str(title).strip():
            return
        for key in {norm_title(title), norm_title(_strip_parens(title))}:
            if key and key not in self.items:
                self.items[key] = {"title": str(title).strip(), "type": typ}

    def lookup(self, title):
        return self.items.get(norm_title(title)) or self.items.get(norm_title(_strip_parens(title or "")))

    def suggest(self, title, n=3):
        keys = difflib.get_close_matches(norm_title(title), list(self.items), n=n, cutoff=0.75)
        return [self.items[k]["title"] for k in keys]

    def counts(self):
        seen = {}
        for v in self.items.values():
            seen[v["title"]] = v["type"]
        return dict(Counter(seen.values()))


def _find_header(rows, col_names):
    wanted = {c.strip().lower() for c in col_names}
    for i, r in enumerate(rows[:15]):
        cells = [str(c).strip().lower() if c is not None else "" for c in r]
        if wanted & set(cells):
            return i, cells
    return None, None


def load_catalog(path, spec: dict) -> Catalog:
    """spec — блок шаблона: {"sheets": [{"name", "type": film|series|auto|keep, "title", "episode"?}]}"""
    wb, _ = load_any(path)
    by_name = {ws.title.strip().lower(): ws for ws in wb.worksheets}
    cat = Catalog()
    missing = []
    for sh in spec["sheets"]:
        ws = by_name.get(sh["name"].strip().lower())
        if ws is None:
            if not sh.get("optional"):
                missing.append(sh["name"])
            continue
        rows = list(ws.iter_rows(values_only=True))
        titles = sh["title"] if isinstance(sh["title"], list) else [sh["title"]]
        hi, header = _find_header(rows, titles)
        if hi is None:
            missing.append(f"{sh['name']} (нет столбца {titles[0]})")
            continue
        tcol = next(header.index(t.strip().lower()) for t in titles if t.strip().lower() in header)
        data = [r for r in rows[hi + 1:] if tcol < len(r) and r[tcol] not in (None, "")]
        if sh["type"] == "column":   # тип записан в самом каталоге: «фильм» / «сериал»
            tc = sh["type_column"].strip().lower()
            if tc not in header:
                missing.append(f"{sh['name']} (нет столбца {sh['type_column']})")
                continue
            ti = header.index(tc)
            values = {norm_title(k): v for k, v in sh.get("values", {"фильм": "film", "сериал": "series"}).items()}
            for r in data:
                typ = values.get(norm_title(r[ti] if ti < len(r) else None))
                if typ:
                    cat.add(r[tcol], typ)
        elif sh["type"] == "auto":
            dcol = None
            if sh.get("duration") and sh["duration"].strip().lower() in header:
                dcol = header.index(sh["duration"].strip().lower())
            groups: dict[str, dict] = {}
            for r in data:
                g = groups.setdefault(norm_title(r[tcol]), {"title": r[tcol], "n": 0, "max_min": 0.0})
                g["n"] += 1
                if dcol is not None and dcol < len(r):
                    g["max_min"] = max(g["max_min"], duration_minutes(r[dcol]) or 0.0)
            for g in groups.values():
                cat.add(g["title"], auto_type(g["n"], g["max_min"], sh))
        else:
            for r in data:
                cat.add(r[tcol], sh["type"])
    if missing:
        raise ValueError("В каталоге не найдены листы/столбцы: " + ", ".join(missing))
    if not cat.items:
        raise ValueError("Каталог пустой")
    return cat


def duration_minutes(v):
    """«00:43:58:04» (ЧЧ:ММ:СС:кадры), «00:43:58», time, timedelta, число минут -> минуты."""
    import datetime as dt
    if v is None or v == "":
        return None
    if isinstance(v, dt.timedelta):
        return v.total_seconds() / 60
    if isinstance(v, dt.time):
        return v.hour * 60 + v.minute + v.second / 60
    if isinstance(v, (int, float)):
        return float(v) * 1440 if v < 1 else float(v)
    m = re.match(r"^\s*(\d+):(\d{1,2})(?::(\d{1,2}))?", str(v))
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2)) + int(m.group(3) or 0) / 60


def auto_type(n_episodes: int, max_minutes: float, rule: dict) -> str:
    """Фильм или сериал по числу серий и хронометражу серии.
    Сериал: серий больше film_max_episodes, ИЛИ серий не меньше long_from_episodes
    и серия длиннее film_max_minutes. Иначе — фильм (в т.ч. 4 серии по ~45 мин, 2 серии по 90)."""
    max_eps = rule.get("film_max_episodes", 4)
    max_min = rule.get("film_max_minutes", 55)
    long_from = rule.get("long_from_episodes", 3)
    if n_episodes > max_eps:
        return "series"
    if n_episodes >= long_from and max_minutes and max_minutes > max_min:
        return "series"
    return "film"


def catalog_path(catalog_dir, code) -> Path | None:
    if not catalog_dir:
        return None
    from readers import SUPPORTED_EXT
    for p in sorted(Path(catalog_dir).glob(f"{code}.*")):
        if p.suffix.lower() in SUPPORTED_EXT:
            return p
    return None


# ---------- чистка названий серий ----------

_JUNK = re.compile(
    r"\b(?:серия|серии|эпизод|эпизода|episode|ep|season|сезон|часть|part|s\d+e\d+|\d+(?:-?(?:я|й|ая|ый|ой))?)\b|№",
    re.IGNORECASE)


def is_bare_episode(value) -> bool:
    """«1 серия», «Серия 2», «Эпизод 3», «Сезон 1. Серия 4», «S01E05» — да.
    «Эпизод 3. Татьяна» — нет (после номера есть название)."""
    if value is None:
        return False
    s = str(value).strip()
    if not s or not re.search(r"\d", s):
        return False
    if not re.search(r"(?i)серия|серии|эпизод|episode|\bep\b|часть|part|s\d+e\d+", s):
        return False
    rest = _JUNK.sub(" ", s)
    return not re.search(r"\w", rest)

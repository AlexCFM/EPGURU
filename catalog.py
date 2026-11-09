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


def _num(v):
    """«1», 1, 1.0, «01» -> «1»; пусто -> None."""
    if v is None or str(v).strip() == "":
        return None
    m = re.match(r"^\s*0*(\d+)(?:\.0+)?\s*$", str(v))
    return m.group(1) if m else None


class Catalog:
    def __init__(self):
        self.items: dict[str, dict] = {}   # ключ -> {"title", "type"}
        self.desc: dict[tuple, tuple] = {}  # (название, сезон, серия) -> (короткое, полное)
        self.desc_by_title: dict[str, set] = {}

    def add_desc(self, title, season, episode, short, long):
        short = str(short).strip() if short not in (None, "") else None
        long = str(long).strip() if long not in (None, "") else None
        if not (short or long) or title in (None, ""):
            return
        for t in {norm_title(title), norm_title(_strip_parens(title))}:
            if not t:
                continue
            s, e = _num(season), _num(episode)
            for key in ((t, s, e), (t, None, e), (t, None, None)):
                self.desc.setdefault(key, (short, long))
            self.desc_by_title.setdefault(t, set()).add((short, long))

    def describe(self, title, season=None, episode=None):
        hit = self._describe(norm_title(title), _num(season), _num(episode))
        if hit:
            return hit
        m = re.match(r"^(.*\D)\s+(\d{1,2})$", norm_title(title))      # «Тайны Лауры 3» — сезон в названии
        if m:
            return self._describe(m.group(1).strip(), _num(season) or m.group(2), _num(episode))
        return None

    def _describe(self, t, s, e):
        for key in ((t, s, e), (t, None, e)):
            if key[2] is not None and key in self.desc:
                return self.desc[key]
        # по одному названию — только если описание у названия одно (фильм или сериал с общим описанием)
        if len(self.desc_by_title.get(t, ())) == 1 or (e is None and (t, None, None) in self.desc):
            return self.desc.get((t, None, None))
        return None

    def add(self, title, typ):
        if title is None or not str(title).strip():
            return
        for key in {norm_title(title), norm_title(_strip_parens(title))}:
            if key and key not in self.items:
                self.items[key] = {"title": str(title).strip(), "type": typ}

    def lookup(self, title):
        hit = self.items.get(norm_title(title)) or self.items.get(norm_title(_strip_parens(title or "")))
        if hit:
            return hit
        # «Тайны Лауры 3» — номер сезона в конце: принимаем, только если без номера это сериал
        m = re.match(r"^(.*\D)\s+(\d{1,2})$", norm_title(title))
        if m:
            base = self.items.get(m.group(1).strip())
            if base and base["type"] == "series":
                return base
        return None

    def suggest(self, title, n=3):
        keys = difflib.get_close_matches(norm_title(title), list(self.items), n=n, cutoff=0.75)
        return [self.items[k]["title"] for k in keys]

    def counts(self):
        seen = {}
        for v in self.items.values():
            seen[v["title"]] = v["type"]
        c = dict(Counter(seen.values()))
        c["descriptions"] = len(self.desc_by_title)
        return c


def _h(x) -> str:
    return re.sub(r"\s+", " ", str(x)).strip().lower() if x is not None else ""


def _find_header(rows, col_names):
    wanted = {_h(c) for c in col_names}
    for i, r in enumerate(rows[:15]):
        cells = [_h(c) for c in r]
        if wanted & set(cells):
            return i, cells
    return None, None


def _col(header, name):
    if not name:
        return None
    for n in (name if isinstance(name, list) else [name]):
        if _h(n) in header:
            return header.index(_h(n))
    return None


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
        tcol = next(header.index(_h(t)) for t in titles if _h(t) in header)
        data = [r for r in rows[hi + 1:] if tcol < len(r) and r[tcol] not in (None, "")]
        # описания (если в листе есть такие столбцы)
        ds, dl = _col(header, sh.get("desc_short")), _col(header, sh.get("desc_long"))
        sc, ec = _col(header, sh.get("season")), _col(header, sh.get("episode"))
        if ds is not None or dl is not None:
            get = lambda r, i: r[i] if i is not None and i < len(r) else None
            for r in data:
                cat.add_desc(r[tcol], get(r, sc), get(r, ec), get(r, ds), get(r, dl))
        if sh["type"] == "descriptions":   # лист только с описаниями
            continue
        if sh["type"] == "column":   # тип записан в самом каталоге: «фильм» / «сериал»
            tc = _h(sh["type_column"])
            if tc not in header:
                if sh.get("default"):          # нет столбца «Тип» — тип по листу
                    for r in data:
                        cat.add(r[tcol], sh["default"])
                    continue
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
            if sh.get("duration") and _h(sh["duration"]) in header:
                dcol = header.index(_h(sh["duration"]))
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
    if not cat.items and not cat.desc:
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


import os as _os
# каталоги из репозитория (по умолчанию); в тестах отключаются через EPGURU_BUILTIN_CATALOGS=""
BUILTIN_DIR = Path(_os.environ.get("EPGURU_BUILTIN_CATALOGS", Path(__file__).resolve().parent / "catalogs"))


def _find(dirpath, code):
    from readers import SUPPORTED_EXT
    if not dirpath or not Path(dirpath).is_dir():
        return None
    for p in sorted(Path(dirpath).glob(f"{code}.*")):
        if p.suffix.lower() in SUPPORTED_EXT:
            return p
    return None


def catalog_path(catalog_dir, code, builtin=True) -> Path | None:
    """Сначала загруженный через сайт, потом встроенный из репозитория."""
    use_builtin = builtin and catalog_dir is not None and str(BUILTIN_DIR) not in ("", ".")
    return _find(catalog_dir, code) or (_find(BUILTIN_DIR, code) if use_builtin else None)


# ---------- чистка названий серий ----------

_ORD = (r"перв|втор|трет|четверт|пят|шест|седьм|восьм|девят|десят|одиннадцат|двенадцат|тринадцат|"
        r"четырнадцат|пятнадцат|шестнадцат|семнадцат|восемнадцат|девятнадцат|двадцат|последн|заключительн|финальн")
_ORD_WORD = rf"(?:{_ORD})(?:ая|ый|ой|ий|ья|ье|ое|ей)"
_JUNK = re.compile(
    rf"\b(?:серия|серии|эпизод|эпизода|episode|ep|season|сезон|часть|part|s\d+e\d+|{_ORD_WORD}|"
    r"\d+(?:-?(?:я|й|ая|ый|ой))?)\b|№",
    re.IGNORECASE)


def is_bare_episode(value) -> bool:
    """«1 серия», «Серия 2», «Эпизод 3», «Сезон 1. Серия 4», «S01E05» — да.
    «Эпизод 3. Татьяна» — нет (после номера есть название)."""
    if value is None:
        return False
    s = str(value).strip()
    if not s or not (re.search(r"\d", s) or re.search(rf"(?i)\b{_ORD_WORD}\b", s)):
        return False
    if not re.search(r"(?i)серия|серии|эпизод|episode|\bep\b|часть|part|s\d+e\d+", s):
        return False
    rest = _JUNK.sub(" ", s)
    return not re.search(r"\w", rest)

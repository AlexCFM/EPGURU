"""Детерминированный обработчик EPG по YAML-шаблонам. Без ИИ.

CLI:
  python epg_processor.py input.xlsx [ещё файлы...] [-t templates/x.yaml] [-o out_dir] [--year 2026]
Без -t шаблон выбирается автоматически по имени файла.
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from pathlib import Path

import openpyxl
import yaml

from catalog import catalog_path, is_bare_episode, load_catalog
from readers import load_any

TEMPLATES_DIR = Path(__file__).parent / "templates"


class EPGError(Exception):
    pass


# ---------- шаблоны ----------

def load_template(path) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def list_templates() -> list[Path]:
    return sorted(TEMPLATES_DIR.glob("*.yaml"))


def match_channel(filename: str, channels: list[dict]):
    up = filename.upper()
    for ch in channels:
        for token in ch["match"]:
            if token.upper() in up:
                return ch["code"], token
    return None


def match_channel_value(value: str, channels: list[dict]):
    """Значение столбца «Channel» в самом EPG (например, Suspenso_RUS)."""
    v = str(value or "").strip().lower()
    if not v:
        return None
    for ch in channels:
        if v in [str(x).strip().lower() for x in ch.get("values", [])]:
            return ch["code"]
    return None


CHANNEL_COLUMNS = ["channel", "канал"]


def channel_from_data(wb):
    """Самое частое значение столбца Channel/Канал на первом листе, где он есть."""
    from collections import Counter
    for ws in wb.worksheets:
        for r in range(1, min(10, ws.max_row) + 1):
            hdr = [_norm(c.value) for c in ws[r]]
            for name in CHANNEL_COLUMNS:
                if name in hdr:
                    col = hdr.index(name) + 1
                    vals = Counter(str(ws.cell(i, col).value).strip() for i in range(r + 1, ws.max_row + 1)
                                   if ws.cell(i, col).value not in (None, ""))
                    if vals:
                        return vals.most_common(1)[0][0]
    return None


def template_for_data_channel(value):
    for p in list_templates():
        code = match_channel_value(value, load_template(p)["channels"])
        if code:
            return p, code
    return None, None


def auto_template(filename: str) -> Path:
    found = [p for p in list_templates() if match_channel(filename, load_template(p)["channels"])]
    if not found:
        raise EPGError(f"Ни один шаблон не подходит к имени файла: {filename}")
    if len(found) > 1:
        names = ", ".join(p.stem for p in found)
        raise EPGError(f"К имени файла подходят несколько шаблонов ({names}), укажите шаблон явно")
    return found[0]


# ---------- имя файла: канал, даты, год ----------

def detect_channel(filename: str, channels: list[dict]) -> tuple[str, str]:
    m = match_channel(filename, channels)
    if not m:
        raise EPGError(f"Не удалось определить канал по имени файла: {filename}")
    return m


def detect_week(filename: str, channel_token: str) -> tuple[str, str, str, str]:
    iso = re.findall(r"(20\d{2})[-_.](\d{2})[-_.](\d{2})", Path(filename).stem)
    if len(iso) >= 2:
        a = dt.date(*map(int, iso[0]))
        b = dt.date(*map(int, iso[1]))
        if (b - a).days == 7:
            b -= dt.timedelta(days=1)
        return f"{a.day:02d}", f"{a.month:02d}", f"{b.day:02d}", f"{b.month:02d}"
    stem = Path(filename).stem.upper()
    if channel_token:
        stem = stem.replace(channel_token.upper(), " ")
    pairs = re.findall(r"(?<!\d)(\d{1,2})[._\-/](\d{1,2})(?:[._\-/]\d{4})?(?!\d)", stem)
    valid = [(d, m) for d, m in pairs if 1 <= int(d) <= 31 and 1 <= int(m) <= 12]
    if len(valid) < 2:
        raise EPGError(f"Не удалось найти две даты недели в имени файла: {filename}")
    (d1, m1), (d2, m2) = valid[0], valid[1]
    return d1.zfill(2), m1.zfill(2), d2.zfill(2), m2.zfill(2)


def year_from_filename(filename: str):
    iso = re.findall(r"(20\d{2})[-_.]\d{2}[-_.]\d{2}", Path(filename).stem)
    if iso:
        return int(iso[-1])
    m = re.search(r"(?<!\d)(20\d{2})(?!\d)", Path(filename).stem)
    return int(m.group(1)) if m else None


# ---------- значения ----------

def _norm(s) -> str:
    return str(s).strip().lower() if s is not None else ""


def _empty(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def parse_time(value):
    if _empty(value):
        return None
    if isinstance(value, dt.datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, dt.time):
        return value.replace(second=0, microsecond=0)
    if isinstance(value, (int, float)):            # доля суток Excel
        total = round(float(value) % 1 * 86400)
        return dt.time(total // 3600 % 24, total // 60 % 60)
    m = re.match(r"^\s*(\d{1,2}):(\d{2})(?::\d{2})?\s*$", str(value))
    return dt.time(int(m.group(1)) % 24, int(m.group(2))) if m else None


def floor_time(value, step: int):
    t = parse_time(value)
    return None if t is None else dt.time(t.hour, t.minute - t.minute % step)


_DATE_RE = re.compile(r"^\s*(\d{1,2})([./\-])(\d{1,2})\2(\d{2,4})\s*$")
DATE_FMT = "DD\\/MM\\/YYYY"   # экранированный «/», чтобы русский Excel не подменял его на «.»


def detect_order(values) -> dict:
    """Порядок день/месяц для текстовых дат по всему столбцу: {разделитель: 'dmy'|'mdy'}."""
    seen = {}
    for v in values:
        if not isinstance(v, str):
            continue
        m = _DATE_RE.match(v)
        if not m:
            continue
        a, sep, b = int(m.group(1)), m.group(2), int(m.group(3))
        if a > 12:
            seen.setdefault(sep, "dmy")
        elif b > 12:
            seen.setdefault(sep, "mdy")
    return seen


def parse_date(value, order=None, months=None):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool) and 20000 < value < 80000:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(value))   # серийный номер Excel
    if _empty(value):
        return None
    m = _DATE_RE.match(str(value))
    if m:
        a, sep, b, y = int(m.group(1)), m.group(2), int(m.group(3)), int(m.group(4))
        y += 2000 if y < 100 else 0
        mode = (order or {}).get(sep)
        if not mode and months:                       # сверка с месяцами недели из имени файла
            if b in months and a not in months:
                mode = "dmy"
            elif a in months and b not in months:
                mode = "mdy"
        mode = mode or ("mdy" if sep == "/" else "dmy")
        if a > 12:
            mode = "dmy"
        elif b > 12:
            mode = "mdy"
        d, mo = (a, b) if mode == "dmy" else (b, a)
        try:
            return dt.date(y, mo, d)
        except ValueError:
            return None
    m = re.match(r"^\s*(\d{4})-(\d{2})-(\d{2})", str(value))
    return dt.date(*(int(x) for x in m.groups())) if m else None


def age_plus(value):
    if _empty(value):
        return value
    m = re.search(r"\d+", str(value))
    return f"{m.group(0)}+" if m else value


def write_date(cell, d: dt.date, fmt: str):
    if isinstance(cell.value, str):
        cell.value = d.strftime("%d/%m/%Y")
    else:
        cell.value = dt.datetime(d.year, d.month, d.day)
        cell.number_format = DATE_FMT


# ---------- книга ----------

def find_header_row(ws, anchors, rows) -> int:
    anchors_n = {_norm(a) for a in anchors}
    for r in range(1, min(rows, ws.max_row) + 1):
        if {_norm(c.value) for c in ws[r]} & anchors_n:
            return r
    return 1


def header_map(ws, hr) -> dict[str, int]:
    return {_norm(c.value): c.column for c in ws[hr] if c.value is not None}


def cols_for(hmap, names) -> list[int]:
    if isinstance(names, str):
        names = [names]
    return [hmap[_norm(n)] for n in names if _norm(n) in hmap]


def row_is_empty(ws, r) -> bool:
    return all(_empty(ws.cell(r, c).value) for c in range(1, ws.max_column + 1))


def known_headers(tpl) -> set[str]:
    names = set(_norm(a) for a in tpl.get("header_anchors", []))
    for rule in tpl["rules"]:
        if rule["type"] == "layout":
            for c in rule["columns"]:
                names |= {_norm(x) for x in [c["name"]] + c.get("from", [])}
        for key in ("columns", "title", "format", "target", "compare_with", "date_columns", "time_columns"):
            v = rule.get(key)
            if isinstance(v, list) and v and isinstance(v[0], str):
                names |= {_norm(x) for x in v}
            elif isinstance(v, str):
                names.add(_norm(v))
    names.discard("")
    return names


def best_header_row(ws, tpl, rows=15):
    """Строка шапки — та, где больше всего знакомых названий столбцов (нужно хотя бы 2)."""
    known = known_headers(tpl)
    best, score = None, 1
    for r in range(1, min(rows, ws.max_row) + 1):
        hits = sum(1 for c in ws[r] if _norm(c.value) in known)
        if hits > score:
            best, score = r, hits
    return best


def process_sheet(ws, tpl, ctx, log):
    hr = best_header_row(ws, tpl, tpl.get("header_search_rows", 15))
    if hr is None:
        if ws.max_row > 1 or ws.max_column > 1:
            log.append(f"[{ws.title}] шапка не найдена (нет знакомых названий столбцов) — лист пропущен")
        return
    if hr > 1:
        log.append(f"[{ws.title}] шапка найдена в строке {hr}, строки выше неё отброшены")
    first = hr + 1

    for rule in tpl["rules"]:
        t = rule["type"]
        hmap = header_map(ws, hr)
        last = ws.max_row

        if t == "layout":
            apply_layout(ws, hr, rule, ctx, log)
            hr, first = 1, 2
            continue

        if t == "drop_columns_prefix":
            pref = rule["prefix"].lower()
            to_drop = sorted((c.column for c in ws[hr]
                              if c.value is not None and str(c.value).strip().lower().startswith(pref)),
                             reverse=True)
            for col in to_drop:
                ws.delete_cols(col)
            log.append(f"[{ws.title}] удалено столбцов Unnamed: {len(to_drop)}")

        elif t == "time_floor":
            n = 0
            for col in cols_for(hmap, rule["columns"]):
                for r in range(first, last + 1):
                    c = ws.cell(r, col)
                    new = floor_time(c.value, rule.get("step_minutes", 5))
                    if new is None:
                        continue
                    if isinstance(c.value, str):
                        c.value = new.strftime("%H:%M")
                    else:
                        c.value, c.number_format = new, "HH:MM"
                    n += 1
            log.append(f"[{ws.title}] время округлено: {n}")

        elif t == "date_format":
            n = 0
            for col in cols_for(hmap, rule["columns"]):
                order = detect_order(ws.cell(r, col).value for r in range(first, last + 1))
                for r in range(first, last + 1):
                    c = ws.cell(r, col)
                    d = parse_date(c.value, order, ctx.get("months"))
                    if d is not None:
                        write_date(c, d, rule.get("format", "DD/MM/YYYY"))
                        n += 1
            log.append(f"[{ws.title}] даты приведены: {n}")

        elif t == "rebuild_dates":
            dcols = cols_for(hmap, rule["date_columns"])
            tcols = cols_for(hmap, rule["time_columns"])
            if not dcols or not tcols:
                raise EPGError(f"[{ws.title}] нет столбца даты или времени для пересчёта дат")
            dcol, tcol = dcols[0], tcols[0]
            cur, prev, n = ctx["start_date"], None, 0
            # хвост предыдущих суток: первое время в файле позже порога -> дата предыдущего дня
            prev_from = rule.get("previous_day_from")
            if prev_from:
                threshold = parse_time(prev_from)
                first_t = next((t for r in range(first, last + 1)
                                if (t := parse_time(ws.cell(r, tcol).value)) is not None), None)
                if first_t is not None and first_t >= threshold:
                    cur -= dt.timedelta(days=1)
                    log.append(f"[{ws.title}] первая строка {first_t:%H:%M} — хвост предыдущих суток, дата {cur:%d/%m/%Y}")
            for r in range(first, last + 1):
                tm = parse_time(ws.cell(r, tcol).value)
                if tm is None:                 # пустое время сбрасывает сравнение
                    prev = None
                    continue
                if prev is not None and tm < prev:
                    cur += dt.timedelta(days=1)
                write_date(ws.cell(r, dcol), cur, rule.get("format", "DD/MM/YYYY"))
                prev, n = tm, n + 1
            log.append(f"[{ws.title}] даты пересчитаны: {n}, последняя {cur:%d/%m/%Y}")

        elif t == "dedupe_consecutive":
            ignore = set(cols_for(hmap, rule.get("ignore_columns", [])))
            cmp_cols = [c for c in range(1, ws.max_column + 1) if c not in ignore]
            to_delete, prev_key = [], None
            for r in range(first, last + 1):
                if row_is_empty(ws, r):        # пустая строка — граница блока
                    prev_key = None
                    continue
                key = tuple(ws.cell(r, c).value for c in cmp_cols)
                if key == prev_key:
                    to_delete.append(r)
                else:
                    prev_key = key
            for r in reversed(to_delete):
                ws.delete_rows(r)
            log.append(f"[{ws.title}] удалено подряд идущих дублей: {len(to_delete)}")

        elif t == "replace_exact":
            n, strip = 0, rule.get("strip", False)
            for col in cols_for(hmap, rule["columns"]):
                for r in range(first, last + 1):
                    c = ws.cell(r, col)
                    if not isinstance(c.value, str):
                        continue
                    key = c.value.strip() if strip else c.value
                    if key in rule["mapping"]:
                        c.value, n = rule["mapping"][key], n + 1
            log.append(f"[{ws.title}] замен: {n}")

        elif t == "age_rating":
            n = 0
            for col in cols_for(hmap, rule["columns"]):
                for r in range(first, last + 1):
                    c = ws.cell(r, col)
                    new = age_plus(c.value)
                    if new != c.value:
                        c.value, n = new, n + 1
            log.append(f"[{ws.title}] возрастных маркеров: {n}")

        elif t == "append_suffix":
            n, suf = 0, rule["suffix"]
            for col in cols_for(hmap, rule["columns"]):
                for r in range(first, last + 1):
                    c = ws.cell(r, col)
                    if _empty(c.value):
                        continue
                    s = str(c.value).strip()
                    s = re.sub(r"\.0$", "", s)  # 16.0 из Excel -> 16
                    if not s.endswith(suf):
                        s += suf
                    if s != c.value:
                        c.value, n = s, n + 1
            log.append(f"[{ws.title}] значений с «{suf}»: {n}")

        elif t == "catalog_format":
            cat = ctx.get("catalog")
            if cat is None:
                log.append(f"[{ws.title}] каталог не загружен — сверка пропущена")
                continue
            tcols, fcols = cols_for(hmap, rule["title"]), cols_for(hmap, rule["format"])
            if not tcols or not fcols:
                log.append(f"[{ws.title}] нет столбца названия или формата — сверка пропущена")
                continue
            mapping = rule["mapping"]
            changed, found, missing, borrowed = 0, 0, {}, {}
            for r in range(first, last + 1):
                title = ws.cell(r, tcols[0]).value
                if _empty(title):
                    continue
                item = cat.lookup(title)
                if not item:
                    for ocode, ocat in ctx.get("others", []):
                        item = ocat.lookup(title)
                        if item:
                            borrowed[str(title).strip()] = ocode
                            break
                if not item:
                    missing.setdefault(str(title).strip(), 0)
                    missing[str(title).strip()] += 1
                    continue
                found += 1
                new = mapping.get(item["type"])
                c = ws.cell(r, fcols[0])
                if new is not None and c.value != new:
                    c.value, changed = new, changed + 1
            log.append(f"[{ws.title}] сверка с каталогом: найдено {found}, формат исправлен {changed}, "
                       f"нет в каталоге {sum(missing.values())}")
            for title, oc in sorted(borrowed.items()):
                log.append(f"   «{title}» — нет в своём каталоге, тип взят из каталога {oc}")
            for title, n in sorted(missing.items()):
                sug = cat.suggest(title)
                log.append(f"   нет в каталоге: «{title}» ×{n}" + (f" — похоже: {', '.join(sug)}" if sug else ""))

        elif t == "clean_episode_name":
            n = 0
            tcols = cols_for(hmap, rule.get("title", []))
            for col in cols_for(hmap, rule["columns"]):
                for r in range(first, last + 1):
                    c = ws.cell(r, col)
                    v = c.value
                    # «Голова медузы. 1 серия» — название сериала + номер: убираем название и проверяем остаток
                    if tcols and isinstance(v, str):
                        title = ws.cell(r, tcols[0]).value
                        if not _empty(title):
                            from catalog import norm_title
                            nt, nv = norm_title(title), norm_title(v)
                            if nv.startswith(nt + " "):
                                v = nv[len(nt):]
                    if is_bare_episode(v):
                        c.value, n = None, n + 1
            log.append(f"[{ws.title}] очищено «пустых» названий серий: {n}")

        elif t == "fill_descriptions":
            cats = ([ctx["catalog"]] if ctx.get("catalog") else []) + [c for _, c in ctx.get("others", [])]
            tc, sc, ec = cols_for(hmap, rule["title"]), cols_for(hmap, rule["season"]), cols_for(hmap, rule["episode"])
            dsc, dlc = cols_for(hmap, rule["desc_short"]), cols_for(hmap, rule["desc_long"])
            if not cats or not tc or not (dsc or dlc):
                log.append(f"[{ws.title}] описания: нет каталога с описаниями — пропущено")
                continue
            filled, nodesc = 0, {}
            for r in range(first, last + 1):
                title = ws.cell(r, tc[0]).value
                if _empty(title):
                    continue
                need_s = dsc and _empty(ws.cell(r, dsc[0]).value)
                need_l = dlc and _empty(ws.cell(r, dlc[0]).value)
                if not (need_s or need_l):
                    continue
                season = ws.cell(r, sc[0]).value if sc else None
                ep = ws.cell(r, ec[0]).value if ec else None
                d = next((x for c in cats if (x := c.describe(title, season, ep))), None)
                if not d:
                    nodesc[str(title).strip()] = nodesc.get(str(title).strip(), 0) + 1
                    continue
                short, long = d
                short, long = short or long, long or short
                if need_s and short:
                    ws.cell(r, dsc[0]).value = short
                if need_l and long:
                    ws.cell(r, dlc[0]).value = long
                filled += 1
            log.append(f"[{ws.title}] описания подставлены из каталогов: {filled}"
                       + (f", не найдены для {sum(nodesc.values())}" if nodesc else ""))
            for title, n in sorted(nodesc.items()):
                log.append(f"   нет описания: «{title}» ×{n}")

        elif t == "clear_same_original":
            a, b = cols_for(hmap, rule["target"]), cols_for(hmap, rule["compare_with"])
            n = 0
            if a and b:
                for r in range(first, last + 1):
                    x, y = ws.cell(r, a[0]).value, ws.cell(r, b[0]).value
                    if not _empty(x) and not _empty(y):
                        from catalog import norm_title
                        if norm_title(x) == norm_title(y):
                            ws.cell(r, a[0]).value, n = None, n + 1
            log.append(f"[{ws.title}] оригинальное название совпадало с русским — очищено: {n}")

        elif t == "film_clear_episode":
            fcol = cols_for(hmap, rule["format"])
            clear = cols_for(hmap, rule["columns"])
            n = 0
            if fcol:
                for r in range(first, last + 1):
                    if str(ws.cell(r, fcol[0]).value or "").strip() == rule.get("film_value", "х/ф"):
                        hit = False
                        for c in clear:
                            if not _empty(ws.cell(r, c).value):
                                ws.cell(r, c).value, hit = None, True
                        n += hit
            log.append(f"[{ws.title}] у фильмов очищены сезон/серия/название серии: {n}")

        elif t == "style":
            apply_style(ws, hr, rule)
            log.append(f"[{ws.title}] оформление применено")

        elif t == "clear_if_equal":
            tc, cc = cols_for(hmap, rule["target"]), cols_for(hmap, rule["compare_with"])
            n = 0
            if tc and cc:
                for r in range(first, last + 1):
                    a, b = ws.cell(r, tc[0]).value, ws.cell(r, cc[0]).value
                    if not _empty(a) and not _empty(b) and str(a).strip() == str(b).strip():
                        ws.cell(r, tc[0]).value, n = None, n + 1
            log.append(f"[{ws.title}] очищено эпизодов: {n}")

        else:
            raise EPGError(f"Неизвестное правило: {t}")


_NAME = re.compile(r"^[^\W\d_][^\W\d_'’.\- ]*(?:[ '’.\-]+[^\W\d_][^\W\d_'’.\- ]*){0,4}\.?$")


def _looks_like_actors(v) -> bool:
    """«Иван Троян,Тереза Рамба,,Онджей Ветхи» — имена через запятую (пустые элементы допустимы), без предложений."""
    if not isinstance(v, str) or len(v) > 400 or re.search(r"[!?;:«»\"]|\.\s", v):
        return False
    parts = [p.strip() for p in v.split(",") if p.strip()]
    return len(parts) >= 2 and all(_NAME.match(p) and len(p) <= 60 for p in parts)


def apply_layout(ws, hr, rule, ctx, log):
    """Привести лист к целевой шапке и порядку столбцов. Неизвестные столбцы отбрасываются (пишутся в лог)."""
    src_hdr = [_norm(c.value) for c in ws[hr]]
    rows = [[c.value for c in r] for r in ws.iter_rows(min_row=hr + 1, max_row=ws.max_row)]
    target = rule["columns"]
    picks, used = [], set()
    for col in target:
        aliases = [_norm(a) for a in [col["name"]] + col.get("from", [])]
        idx = next((src_hdr.index(a) for a in aliases if a in src_hdr), None)
        picks.append(idx)
        if idx is not None:
            used.add(idx)
    dropped = [ws.cell(hr, i + 1).value for i, h in enumerate(src_hdr)
               if h and i not in used and not h.startswith("unnamed")]
    out, empty = [], 0
    for r in rows:
        vals = [r[i] if i is not None and i < len(r) else None for i in picks]
        if all(_empty(v) for v in vals):
            empty += 1
            continue
        out.append(vals)
    names = [c["name"] for c in target]

    def ci(key):
        return next((i for i, c in enumerate(target) if c.get("role") == key), None)

    # канал: значение из выгрузки -> как принято у нас (Suspenso_RUS -> 911)
    ch = ci("channel")
    if ch is not None and ctx.get("channel_display") is not None:
        for v in out:
            v[ch] = ctx["channel_display"]
    # числа: сезон, серия, год
    for i, c in enumerate(target):
        if c.get("int"):
            for v in out:
                if isinstance(v[i], str) and re.fullmatch(r"\s*\d+\s*", v[i]):
                    v[i] = int(v[i])
                elif isinstance(v[i], float) and v[i].is_integer():
                    v[i] = int(v[i])
    # сдвиг в выгрузке: актёры лежат в «описании», а «Актёры» пустые
    sd, ld, ac = ci("desc_short"), ci("desc_long"), ci("actors")
    if None not in (sd, ac):
        moved = 0
        for v in out:
            if _empty(v[ac]) and _looks_like_actors(v[sd]):
                v[ac], v[sd] = re.sub(r",\s*,+", ",", str(v[sd]).strip()).strip(", "), None
                moved += 1
        if moved:
            log.append(f"[{ws.title}] актёры перенесены из «{names[sd]}» в «{names[ac]}»: {moved} (сдвиг в выгрузке)")

    ws.delete_rows(1, ws.max_row)
    ws.append(names)
    for v in out:
        ws.append(v)
    if ws.max_column > len(names):
        ws.delete_cols(len(names) + 1, ws.max_column - len(names))
    if rule.get("sheet_title"):
        ws.title = rule["sheet_title"]
    log.append(f"[{ws.title}] шапка приведена к образцу: {len(names)} столбцов"
               + (f"; отброшены: {', '.join(map(str, dropped))}" if dropped else "")
               + (f"; убрано пустых строк: {empty}" if empty else ""))


def apply_style(ws, hr, rule):
    """Читаемый EPG: ширины, перенос, границы, шапка, закрепление, разделение по дням, печать."""
    import math
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    font = rule.get("font", "Arial")
    size = rule.get("size", 10)
    hdr = [str(c.value or "") for c in ws[hr]]
    widths = {_norm(k): v for k, v in rule.get("widths", {}).items()}
    center = {_norm(c) for c in rule.get("center", [])}
    thin = Side(style="thin", color=rule.get("border_color", "BFBFBF"))
    day_side = Side(style="medium", color=rule.get("day_border_color", "7F7F7F"))
    fills = [PatternFill("solid", fgColor=c) for c in rule.get("day_fills", ["FFFFFF", "F3F6F9"])]
    head_fill = PatternFill("solid", fgColor=rule.get("header_fill", "3B3835"))
    head_font = Font(name=font, size=size, bold=True, color=rule.get("header_font_color", "FFFFFF"))
    body_font = Font(name=font, size=size)
    title_font = Font(name=font, size=size, bold=True)
    bold_cols = {_norm(c) for c in rule.get("bold", [])}
    ncol = len(hdr)

    col_w = []
    for i, h in enumerate(hdr, 1):
        w = widths.get(_norm(h), 14)
        col_w.append(w)
        ws.column_dimensions[get_column_letter(i)].width = w

    for i, h in enumerate(hdr, 1):
        c = ws.cell(hr, i)
        c.font, c.fill = head_font, head_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.row_dimensions[hr].height = rule.get("header_height", 30)

    day_col = next((i for i, h in enumerate(hdr, 1) if _norm(h) == _norm(rule.get("day_column", ""))), None)
    line_h = size * 1.35
    prev_day, band = object(), -1
    for r in range(hr + 1, ws.max_row + 1):
        day = ws.cell(r, day_col).value if day_col else None
        new_day = day_col and day != prev_day
        if new_day:
            band += 1
            prev_day = day
        fill = fills[band % len(fills)] if day_col else None
        lines = 1
        for i in range(1, ncol + 1):
            c = ws.cell(r, i)
            h = _norm(hdr[i - 1])
            c.font = title_font if h in bold_cols else body_font
            c.alignment = Alignment(horizontal="center" if h in center else "left",
                                    vertical="top", wrap_text=True)
            c.border = Border(left=thin, right=thin, bottom=thin,
                              top=day_side if (new_day and r > hr + 1) else thin)
            if fill is not None:
                c.fill = fill
            if c.value is not None:
                text = str(c.value)
                chars = max(1.0, col_w[i - 1] * 1.2)    # Arial 10: ~1.2 символа на единицу ширины, с запасом
                need = sum(max(1, math.ceil(len(part) / chars)) for part in text.split("\n"))
                lines = max(lines, need)
        ws.row_dimensions[r].height = round(line_h * lines + 6, 1)

    ws.freeze_panes = rule.get("freeze", "A2")
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(ncol)}{ws.max_row}" if rule.get("filter") else None
    ws.sheet_view.zoomScale = rule.get("zoom", 100)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{hr}:{hr}"
    ws.page_margins.left = ws.page_margins.right = 0.3


def year_from_data(wb, tpl):
    """Год из уже заполненного столбца даты исходника (первая распознанная дата)."""
    names = []
    for rule in tpl["rules"]:
        names += rule.get("date_columns", []) + (rule.get("columns", []) if rule["type"] == "date_format" else [])
    for ws in wb.worksheets:
        hr = find_header_row(ws, tpl.get("header_anchors", []), tpl.get("header_search_rows", 10))
        for col in cols_for(header_map(ws, hr), names):
            for r in range(hr + 1, ws.max_row + 1):
                d = parse_date(ws.cell(r, col).value)
                if d:
                    return d.year
    return None


def versioned(path: Path) -> Path:
    if not path.exists():
        return path
    v = 2
    while (p := path.with_name(f"{path.stem}_v{v}{path.suffix}")).exists():
        v += 1
    return p


def template_for_channel(code: str) -> Path:
    for p in list_templates():
        if any(ch["code"] == code for ch in load_template(p)["channels"]):
            return p
    raise EPGError(f"Нет шаблона для канала {code}")


def process(input_path, template_path=None, out_dir="output", original_name=None, year=None, channel=None,
            catalog_dir=None):
    """channel — код канала, выбранный вручную (если в имени файла канала нет)."""
    name = original_name or Path(input_path).name
    wb, kind = load_any(input_path)
    data_channel = None
    if channel and not template_path:
        template_path = template_for_channel(channel)
    if not template_path:
        try:
            template_path = auto_template(name)
        except EPGError:
            data_channel = channel_from_data(wb)
            template_path, channel = template_for_data_channel(data_channel)
            if not template_path:
                raise EPGError(f"Канал не найден ни в имени файла «{name}», ни в столбце Channel"
                               + (f" (там «{data_channel}»)" if data_channel else ""))
    tpl = load_template(template_path)
    found = match_channel(name, tpl["channels"])
    if channel and (not found or found[0] != channel):
        if not any(ch["code"] == channel for ch in tpl["channels"]):
            raise EPGError(f"Канал {channel} не относится к шаблону {tpl['name']}")
        code, token = channel, ""
    elif found:
        code, token = found
    else:
        raise EPGError(f"Не удалось определить канал по имени файла: {name}")
    d1, m1, d2, m2 = detect_week(name, token)

    ctx = {"months": {int(m1), int(m2)}}
    ch_def = next((c for c in tpl["channels"] if c["code"] == code), {})
    ctx["channel_display"] = ch_def.get("display", code)
    if tpl.get("needs_year"):
        y = year or year_from_filename(name) or year_from_data(wb, tpl)
        if not y:
            raise EPGError("Не удалось определить год расписания — укажите его вручную (--year)")
        ctx["start_date"] = dt.date(int(y), int(m1), int(d1))

    spec = (tpl.get("catalogs") or {}).get(code)
    cpath = catalog_path(catalog_dir, code) if spec else None
    if cpath:
        ctx["catalog"] = load_catalog(cpath, spec)
    ctx["others"] = []
    if catalog_dir:
        for tp in list_templates():
            for ocode, ospec in (load_template(tp).get("catalogs") or {}).items():
                op = catalog_path(catalog_dir, str(ocode))
                if str(ocode) != str(code) and op:
                    try:
                        ctx["others"].append((str(ocode), load_catalog(op, ospec)))
                    except Exception:  # noqa: BLE001
                        pass

    log = [f"Шаблон: {tpl['name']} · канал {code} · неделя {d1}.{m1}–{d2}.{m2}"]
    if spec:
        log.append(f"Каталог: {'загружен' if cpath else 'НЕ ЗАГРУЖЕН'}")
    if kind != "xlsx":
        log.append(f"Исходный формат: {kind.upper()} → сохранено в XLSX")
    if data_channel:
        log.append(f"Канал взят из столбца Channel: «{data_channel}» → {code}")
    for ws in wb.worksheets:
        process_sheet(ws, tpl, ctx, log)

    y2 = ctx.get("year_end") or year_from_filename(name) or year_from_data(wb, tpl) or dt.date.today().year
    out = Path(out_dir) / tpl["output_name"].format(channel=code, name=ch_def.get("name", code),
                                                   d1=d1, m1=m1, d2=d2, m2=m2, y2=y2)
    out.parent.mkdir(parents=True, exist_ok=True)
    if tpl.get("versioning"):
        out = versioned(out)
    wb.save(out)

    if not out.exists() or out.stat().st_size == 0:
        raise EPGError("Выходной файл не создан или пустой")
    openpyxl.load_workbook(out).close()
    log.append(f"Готово: {out.name}, {out.stat().st_size:,} байт".replace(",", " "))
    return out, log


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("-t", "--template", help="YAML-шаблон; без него — автовыбор по имени файла")
    ap.add_argument("-o", "--out", default="output")
    ap.add_argument("--year", type=int, help="год расписания, если его нет в имени файла и данных")
    ap.add_argument("--catalogs", help="папка с каталогами: <КОД_КАНАЛА>.xlsx, например 911.xlsx, VESNA.xlsx")
    a = ap.parse_args()
    ok = True
    for p in a.inputs:
        try:
            _, log = process(p, a.template, a.out, year=a.year, catalog_dir=a.catalogs)
            print("\n".join(log), end="\n\n")
        except Exception as e:
            ok = False
            print(f"ОШИБКА {p}: {e}", file=sys.stderr)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()

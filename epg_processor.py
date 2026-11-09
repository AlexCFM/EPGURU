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


def process_sheet(ws, tpl, ctx, log):
    hr = find_header_row(ws, tpl.get("header_anchors", []), tpl.get("header_search_rows", 10))
    if hr == 1 and not ({_norm(c.value) for c in ws[1]} & {_norm(a) for a in tpl.get("header_anchors", [])}):
        log.append(f"[{ws.title}] заголовки не найдены, лист пропущен")
        return
    first = hr + 1

    for rule in tpl["rules"]:
        t = rule["type"]
        hmap = header_map(ws, hr)
        last = ws.max_row

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
            changed, found, missing = 0, 0, {}
            for r in range(first, last + 1):
                title = ws.cell(r, tcols[0]).value
                if _empty(title):
                    continue
                item = cat.lookup(title)
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
            for title, n in sorted(missing.items()):
                sug = cat.suggest(title)
                log.append(f"   нет в каталоге: «{title}» ×{n}" + (f" — похоже: {', '.join(sug)}" if sug else ""))

        elif t == "clean_episode_name":
            n = 0
            for col in cols_for(hmap, rule["columns"]):
                for r in range(first, last + 1):
                    c = ws.cell(r, col)
                    if is_bare_episode(c.value):
                        c.value, n = None, n + 1
            log.append(f"[{ws.title}] очищено «пустых» названий серий: {n}")

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
    if channel and not template_path:
        template_path = template_for_channel(channel)
    template_path = template_path or auto_template(name)
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

    wb, kind = load_any(input_path)
    ctx = {"months": {int(m1), int(m2)}}
    if tpl.get("needs_year"):
        y = year or year_from_filename(name) or year_from_data(wb, tpl)
        if not y:
            raise EPGError("Не удалось определить год расписания — укажите его вручную (--year)")
        ctx["start_date"] = dt.date(int(y), int(m1), int(d1))

    spec = (tpl.get("catalogs") or {}).get(code)
    cpath = catalog_path(catalog_dir, code) if spec else None
    if cpath:
        ctx["catalog"] = load_catalog(cpath, spec)

    log = [f"Шаблон: {tpl['name']} · канал {code} · неделя {d1}.{m1}–{d2}.{m2}"]
    if spec:
        log.append(f"Каталог: {'загружен' if cpath else 'НЕ ЗАГРУЖЕН'}")
    if kind != "xlsx":
        log.append(f"Исходный формат: {kind.upper()} → сохранено в XLSX")
    for ws in wb.worksheets:
        process_sheet(ws, tpl, ctx, log)

    out = Path(out_dir) / tpl["output_name"].format(channel=code, d1=d1, m1=m1, d2=d2, m2=m2)
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

"""Детерминированный обработчик EPG по YAML-шаблонам. Без ИИ.

CLI:  python epg_processor.py input.xlsx [-t templates/make_sense.yaml] [-o out_dir]
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from pathlib import Path

import openpyxl
import yaml


class EPGError(Exception):
    pass


# ---------- имя файла: канал и даты ----------

def detect_channel(filename: str, channels: list[dict]) -> tuple[str, str]:
    up = filename.upper()
    for ch in channels:
        for token in ch["match"]:
            if token.upper() in up:
                return ch["code"], token
    raise EPGError(f"Не удалось определить канал по имени файла: {filename}")


def detect_week(filename: str, channel_token: str) -> tuple[str, str, str, str]:
    stem = Path(filename).stem.upper().replace(channel_token.upper(), " ")
    # DD.MM / DD_MM / DD-MM, допускается год после
    pairs = re.findall(r"(?<!\d)(\d{1,2})[._\-/](\d{1,2})(?:[._\-/]\d{4})?(?!\d)", stem)
    valid = [(d, m) for d, m in pairs if 1 <= int(d) <= 31 and 1 <= int(m) <= 12]
    if len(valid) < 2:
        raise EPGError(f"Не удалось найти две даты недели в имени файла: {filename}")
    (d1, m1), (d2, m2) = valid[0], valid[1]
    return d1.zfill(2), m1.zfill(2), d2.zfill(2), m2.zfill(2)


# ---------- значения ----------

def _norm(s) -> str:
    return str(s).strip().lower() if s is not None else ""


def floor_time(value, step: int):
    """Возвращает dt.time, округлённое вниз, или None если не распознано."""
    if value is None or value == "":
        return None
    if isinstance(value, dt.datetime):
        t = value.time()
    elif isinstance(value, dt.time):
        t = value
    elif isinstance(value, (int, float)):            # доля суток Excel
        total = round(float(value) % 1 * 86400)
        t = dt.time(total // 3600 % 24, total // 60 % 60)
    else:
        m = re.match(r"^\s*(\d{1,2}):(\d{2})(?::\d{2})?\s*$", str(value))
        if not m:
            return None
        t = dt.time(int(m.group(1)) % 24, int(m.group(2)))
    return dt.time(t.hour, t.minute - t.minute % step)


def parse_date(value):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if value is None or value == "":
        return None
    m = re.match(r"^\s*(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})\s*$", str(value))
    if m:                                             # всегда день-первым
        d, mo, y = (int(x) for x in m.groups())
        y += 2000 if y < 100 else 0
        return dt.date(y, mo, d)
    m = re.match(r"^\s*(\d{4})-(\d{2})-(\d{2})", str(value))
    if m:
        return dt.date(*(int(x) for x in m.groups()))
    return None


def age_plus(value):
    if value is None or value == "":
        return value
    m = re.search(r"\d+", str(value))
    return f"{m.group(0)}+" if m else value


# ---------- книга ----------

def find_header_row(ws, anchors, rows) -> int:
    anchors_n = {_norm(a) for a in anchors}
    for r in range(1, min(rows, ws.max_row) + 1):
        vals = {_norm(c.value) for c in ws[r]}
        if vals & anchors_n:
            return r
    return 1


def header_map(ws, hr) -> dict[str, int]:
    return {_norm(c.value): c.column for c in ws[hr] if c.value is not None}


def cols_for(hmap, names) -> list[int]:
    return [hmap[_norm(n)] for n in names if _norm(n) in hmap]


def process_sheet(ws, tpl, log):
    hr = find_header_row(ws, tpl.get("header_anchors", []), tpl.get("header_search_rows", 10))
    first, last = hr + 1, ws.max_row

    for rule in tpl["rules"]:
        t = rule["type"]
        hmap = header_map(ws, hr)

        if t == "drop_columns_prefix":
            pref = rule["prefix"].lower()
            to_drop = sorted(
                (c.column for c in ws[hr]
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
                        c.value = new
                        c.number_format = "HH:MM"
                    n += 1
            log.append(f"[{ws.title}] время округлено: {n}")

        elif t == "date_format":
            n = 0
            for col in cols_for(hmap, rule["columns"]):
                for r in range(first, last + 1):
                    c = ws.cell(r, col)
                    d = parse_date(c.value)
                    if d is None:
                        continue
                    if isinstance(c.value, str):
                        c.value = d.strftime("%d/%m/%Y")
                    else:
                        c.value = dt.datetime(d.year, d.month, d.day)
                        c.number_format = rule.get("format", "DD/MM/YYYY")
                    n += 1
            log.append(f"[{ws.title}] даты приведены: {n}")

        elif t == "replace_exact":
            n = 0
            for col in cols_for(hmap, rule["columns"]):
                for r in range(first, last + 1):
                    c = ws.cell(r, col)
                    if isinstance(c.value, str) and c.value in rule["mapping"]:
                        c.value = rule["mapping"][c.value]
                        n += 1
            log.append(f"[{ws.title}] замен формата: {n}")

        elif t == "age_rating":
            n = 0
            for col in cols_for(hmap, rule["columns"]):
                for r in range(first, last + 1):
                    c = ws.cell(r, col)
                    new = age_plus(c.value)
                    if new != c.value:
                        c.value, n = new, n + 1
            log.append(f"[{ws.title}] возрастных маркеров: {n}")

        elif t == "clear_if_equal":
            tc = cols_for(hmap, [rule["target"]])
            cc = cols_for(hmap, [rule["compare_with"]])
            n = 0
            if tc and cc:
                for r in range(first, last + 1):
                    a, b = ws.cell(r, tc[0]).value, ws.cell(r, cc[0]).value
                    if a not in (None, "") and b not in (None, "") and str(a).strip() == str(b).strip():
                        ws.cell(r, tc[0]).value = None
                        n += 1
            log.append(f"[{ws.title}] очищено эпизодов: {n}")

        else:
            raise EPGError(f"Неизвестное правило: {t}")


def load_template(path) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def process(input_path, template_path, out_dir, original_name=None) -> tuple[Path, list[str]]:
    tpl = load_template(template_path)
    name = original_name or Path(input_path).name
    code, token = detect_channel(name, tpl["channels"])
    d1, m1, d2, m2 = detect_week(name, token)
    out_name = tpl["output_name"].format(channel=code, d1=d1, m1=m1, d2=d2, m2=m2)

    wb = openpyxl.load_workbook(input_path)
    log = [f"Канал: {code}, неделя {d1}.{m1}–{d2}.{m2}"]
    for ws in wb.worksheets:
        process_sheet(ws, tpl, log)

    out = Path(out_dir) / out_name
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)

    # проверки из раздела 6
    if not out.exists() or out.stat().st_size == 0:
        raise EPGError("Выходной файл не создан или пустой")
    openpyxl.load_workbook(out).close()
    log.append(f"Готово: {out.name}, {out.stat().st_size:,} байт".replace(",", " "))
    return out, log


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("-t", "--template", default=str(Path(__file__).parent / "templates/make_sense.yaml"))
    ap.add_argument("-o", "--out", default="output")
    a = ap.parse_args()
    ok = True
    for p in a.inputs:
        try:
            _, log = process(p, a.template, a.out)
            print("\n".join(log), end="\n\n")
        except Exception as e:
            ok = False
            print(f"ОШИБКА {p}: {e}", file=sys.stderr)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()

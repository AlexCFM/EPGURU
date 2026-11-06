"""Чинит все сломанные даты в xlsx: приводит к DD/MM/YYYY, исправляет перепутанные день/месяц.

  python fix_dates.py файл.xlsx [ещё файлы или папки...] [-o out_dir] [--in-place]

Что считается сломанным:
  - даты текстом в любом виде (09.11.2026, 11/09/2026, 2026-11-09, 9.11.26) -> 09/11/2026
  - даты числом (серийный номер Excel) -> настоящая дата
  - день и месяц перепутаны (Excel прочитал 04.11 как 11 апреля): если в имени файла есть
    неделя и дата в неё не попадает, а с переставленными днём и месяцем попадает — переставляем
  - формат ячейки, из-за которого русский Excel показывает 09.11.2026 вместо 09/11/2026

Столбцы дат: заголовок Дата/Date или больше половины значений похожи на даты.
Исходник не трогается (кроме --in-place), результат: <имя>_dates_fixed.xlsx
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from pathlib import Path

import openpyxl

from epg_processor import DATE_FMT, _empty, _norm, detect_order, parse_date

DATE_HEADERS = {"дата", "date", "дата эфира", "air date"}


def week_from_name(name: str):
    """(начало, конец) недели из имени файла, год — из имени или None."""
    stem = Path(name).stem
    pairs = re.findall(r"(?<!\d)(\d{1,2})[._\-/](\d{1,2})(?:[._\-/](\d{4}))?(?!\d)", stem)
    valid = [(int(d), int(m), int(y) if y else None) for d, m, y in pairs
             if 1 <= int(d) <= 31 and 1 <= int(m) <= 12]
    return valid[:2] if len(valid) >= 2 else None


def in_week(d: dt.date, week) -> bool:
    if not week:
        return True
    (d1, m1, _), (d2, m2, _) = week
    try:
        start = dt.date(d.year, m1, d1) - dt.timedelta(days=1)
        end = dt.date(d.year + (1 if m2 < m1 else 0), m2, d2) + dt.timedelta(days=1)
    except ValueError:
        return True
    if m2 < m1 and d.month == m2:             # неделя через Новый год
        start = start.replace(year=d.year - 1)
        end = end.replace(year=d.year)
    return start <= d <= end


def swapped(d: dt.date):
    if d.day > 12 or d.day == d.month:
        return None
    try:
        return dt.date(d.year, d.day, d.month)
    except ValueError:
        return None


def looks_like_date(v) -> bool:
    if isinstance(v, (dt.datetime, dt.date)):
        return True
    return isinstance(v, str) and parse_date(v) is not None


def date_columns(ws, hr):
    cols = []
    for c in ws[hr]:
        if _norm(c.value) in DATE_HEADERS:
            cols.append(c.column)
            continue
        vals = [ws.cell(r, c.column).value for r in range(hr + 1, ws.max_row + 1)]
        vals = [v for v in vals if not _empty(v)]
        if vals and sum(looks_like_date(v) for v in vals) / len(vals) > 0.5:
            cols.append(c.column)
    return cols


def header_row(ws) -> int:
    for r in range(1, min(10, ws.max_row) + 1):
        if any(_norm(c.value) in DATE_HEADERS for c in ws[r]):
            return r
    return 1


def fix_workbook(path: Path, out: Path, original_name=None):
    name = original_name or path.name
    week = week_from_name(name)
    months = {week[0][1], week[1][1]} if week else None
    wb = openpyxl.load_workbook(path)
    stats = {"text": 0, "serial": 0, "swapped": 0, "format": 0, "bad": []}

    for ws in wb.worksheets:
        hr = header_row(ws)
        for col in date_columns(ws, hr):
            order = detect_order(ws.cell(r, col).value for r in range(hr + 1, ws.max_row + 1))
            for r in range(hr + 1, ws.max_row + 1):
                c = ws.cell(r, col)
                v = c.value
                if _empty(v):
                    continue
                d = parse_date(v, order, months)
                if d is None:
                    stats["bad"].append(f"{ws.title}!{c.coordinate}: {v!r}")
                    continue
                if not in_week(d, week) and (s := swapped(d)) and in_week(s, week):
                    d = s
                    stats["swapped"] += 1
                elif isinstance(v, str):
                    stats["text"] += 1
                elif isinstance(v, (int, float)):
                    stats["serial"] += 1
                elif c.number_format == DATE_FMT:
                    continue
                else:
                    stats["format"] += 1
                if isinstance(v, str):
                    c.value = d.strftime("%d/%m/%Y")
                else:
                    c.value = dt.datetime(d.year, d.month, d.day)
                    c.number_format = DATE_FMT

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    openpyxl.load_workbook(out).close()
    return stats


def collect(inputs):
    for p in map(Path, inputs):
        if p.is_dir():
            yield from (f for f in sorted(p.rglob("*.xlsx"))
                        if not f.name.startswith("~$") and "_dates_fixed" not in f.stem)
        else:
            yield p


def main():
    ap = argparse.ArgumentParser(description="Чинит даты в xlsx")
    ap.add_argument("inputs", nargs="+", help="файлы или папки")
    ap.add_argument("-o", "--out", help="папка для результатов (по умолчанию рядом с исходником)")
    ap.add_argument("--in-place", action="store_true", help="перезаписать исходные файлы")
    a = ap.parse_args()
    ok = True
    for src in collect(a.inputs):
        out = src if a.in_place else (Path(a.out) if a.out else src.parent) / f"{src.stem}_dates_fixed.xlsx"
        try:
            s = fix_workbook(src, out)
            print(f"{src.name} -> {out.name}: текст {s['text']}, числа {s['serial']}, "
                  f"перепутанные день/месяц {s['swapped']}, формат {s['format']}")
            for b in s["bad"]:
                print(f"   не распознано: {b}")
        except Exception as e:
            ok = False
            print(f"ОШИБКА {src}: {e}", file=sys.stderr)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()

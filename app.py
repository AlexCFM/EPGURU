"""Веб-интерфейс: streamlit run app.py  → перетащил файлы, скачал результат."""
import tempfile
from pathlib import Path

import streamlit as st

from epg_processor import list_templates, load_template, process
from fix_dates import fix_workbook

st.set_page_config(page_title="EPGURU", page_icon="📺")
st.title("📺 EPGURU")

tab_epg, tab_dates = st.tabs(["Обработка EPG", "Починить даты"])

with tab_epg:
    AUTO = "Авто (по имени файла)"
    tpls = {AUTO: None} | {load_template(p)["name"]: p for p in list_templates()}
    choice = st.selectbox("Шаблон", list(tpls))
    files = st.file_uploader("EPG-файлы (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="epg")
    for f in files or []:
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "in.xlsx"
            src.write_bytes(f.getvalue())
            try:
                out, log = process(src, tpls[choice], Path(tmp) / "out", original_name=f.name)
                st.success(f"{f.name} → {log[-1]}")
                with st.expander("Лог"):
                    st.text("\n".join(log))
                st.download_button(f"⬇ {out.name}", out.read_bytes(), file_name=out.name, key="e" + f.name)
            except Exception as e:
                st.error(f"{f.name}: {e}")

with tab_dates:
    st.caption("Любые xlsx: даты текстом, числом, с перепутанными днём и месяцем → DD/MM/YYYY")
    files = st.file_uploader("Файлы (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="dates")
    for f in files or []:
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "in.xlsx"
            src.write_bytes(f.getvalue())
            name = f"{Path(f.name).stem}_dates_fixed.xlsx"
            out = Path(tmp) / name
            try:
                s = fix_workbook(src, out, original_name=f.name)
                st.success(f"{f.name}: текст {s['text']}, числа {s['serial']}, "
                           f"перепутанные день/месяц {s['swapped']}, формат {s['format']}")
                if s["bad"]:
                    st.warning("Не распознано:\n" + "\n".join(s["bad"][:50]))
                st.download_button(f"⬇ {name}", out.read_bytes(), file_name=name, key="d" + f.name)
            except Exception as e:
                st.error(f"{f.name}: {e}")

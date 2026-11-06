"""Веб-интерфейс: streamlit run app.py  → перетащил файлы, скачал результат."""
import datetime as dt
import tempfile
from pathlib import Path

import streamlit as st

from epg_processor import list_templates, load_template, process

st.set_page_config(page_title="EPGURU", page_icon="📺")
st.title("📺 EPGURU")

AUTO = "Авто (по имени файла)"
tpls = {AUTO: None} | {load_template(p)["name"]: p for p in list_templates()}
choice = st.selectbox("Шаблон", list(tpls))
year = st.number_input("Год (нужен FIGHT CLUB, если его нет в имени файла и в данных)",
                       min_value=2000, max_value=2100, value=None, placeholder="автоматически")
files = st.file_uploader("EPG-файлы (.xlsx)", type=["xlsx"], accept_multiple_files=True)

for f in files or []:
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "in.xlsx"
        src.write_bytes(f.getvalue())
        try:
            out, log = process(src, tpls[choice], Path(tmp) / "out", original_name=f.name, year=year)
            st.success(f"{f.name} → {log[-1]}")
            with st.expander("Лог"):
                st.text("\n".join(log))
            st.download_button(f"⬇ {out.name}", out.read_bytes(), file_name=out.name, key=f.name)
        except Exception as e:
            st.error(f"{f.name}: {e}")

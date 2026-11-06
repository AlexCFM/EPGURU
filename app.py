"""Веб-интерфейс: streamlit run app.py  → перетащил файлы, скачал результат."""
import tempfile
from pathlib import Path

import streamlit as st

from epg_processor import process

TEMPLATES = Path(__file__).parent / "templates"

st.set_page_config(page_title="EPG процессор", page_icon="📺")
st.title("📺 Обработка EPG")

tpls = sorted(TEMPLATES.glob("*.yaml"))
tpl = st.selectbox("Шаблон", tpls, format_func=lambda p: p.stem.replace("_", " ").upper())
files = st.file_uploader("EPG-файлы (.xlsx)", type=["xlsx"], accept_multiple_files=True)

for f in files or []:
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "in.xlsx"
        src.write_bytes(f.getvalue())
        try:
            out, log = process(src, tpl, Path(tmp) / "out", original_name=f.name)
            st.success(log[-1])
            with st.expander("Лог"):
                st.text("\n".join(log))
            st.download_button(f"⬇ {out.name}", out.read_bytes(), file_name=out.name, key=f.name)
        except Exception as e:
            st.error(f"{f.name}: {e}")

# EPG процессор

Обработка EPG по шаблонам без ИИ. Правила лежат в `templates/*.yaml`, движок — в `epg_processor.py`.

## Установка
    pip install openpyxl pyyaml streamlit

## Запуск
Из командной строки (можно передать сразу несколько файлов):

    python epg_processor.py "EPG PLAN B 09.11-15.11.xlsx" -o output

Веб-интерфейс в браузере:

    streamlit run app.py

## Новый шаблон
Скопируйте `templates/make_sense.yaml`, поменяйте каналы, правила и маску имени. Доступные правила:
`drop_columns_prefix`, `time_floor`, `date_format`, `replace_exact`, `age_rating`, `clear_if_equal`.

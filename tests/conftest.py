import os

# тесты не должны видеть встроенные каталоги из репозитория
os.environ.setdefault("EPGURU_BUILTIN_CATALOGS", "")

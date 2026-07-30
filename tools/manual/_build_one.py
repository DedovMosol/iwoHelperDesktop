# -*- coding: utf-8 -*-
"""Сборка ОДНОГО руководства: язык берётся из переменной окружения MANUAL_LANG.

Отдельным файлом, потому что и движок, и python-docx держат документ в состоянии модуля —
собрать два языка в одном процессе значит получить второй документ поверх первого.
Запускается из build_manual.py, руками звать незачем.
"""
import importlib
import os

import engine

LANG = os.environ.get("MANUAL_LANG", "ru")

engine.style_document()
engine.page_numbers()
engine.update_fields_on_open()

body = importlib.import_module("body_" + LANG)
body.build()

engine.attach_footnotes()
if not os.path.isdir(engine.OUT_DIR):
    raise SystemExit("нет папки назначения: " + engine.OUT_DIR)
engine.doc.save(engine.OUT)
engine.retheme(engine.OUT)   # шрифт темы — тоже Times New Roman
engine.finalize_toc(engine.OUT)

missing = [name for name in engine.FIG_ORDER if name not in engine._inserted]
print("рисунков вставлено:", len(engine._inserted), "из", len(engine.FIG_ORDER))
if missing:
    raise SystemExit("НЕ ВСТАВЛЕНЫ рисунки: " + ", ".join(missing))
print("таблиц:", len(engine.TABLES), "| сносок:", len(engine.FOOTNOTES))
print("сохранено:", engine.OUT)

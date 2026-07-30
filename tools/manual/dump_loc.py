# -*- coding: utf-8 -*-
"""Выгрузка русских подписей из каталога Loc.cs — чтобы текст инструкции называл
элементы ровно так же, как они подписаны на экране."""
import io
import json
import re

# Пути — ОТ САМОГО СКРИПТА, а не литералами с чужой машины: конвейер живёт в
# репозитории, и жёсткий путь работал бы только там, где его однажды написали.
import os

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "..", "src", "Loc.cs")
OUT = os.path.join(HERE, "loc_ru.json")

text = io.open(SRC, encoding="utf-8").read()
call = re.compile(r'A\(\s*"([^"]+)"\s*,\s*(.*?)\)\s*;', re.S)


def split_args(chunk):
    """Разбить аргументы по запятым ВЕРХНЕГО уровня (внутри строк запятые не считаются)."""
    out, cur, in_quotes, escaped = [], "", False, False
    for ch in chunk:
        if escaped:
            cur += ch
            escaped = False
            continue
        if ch == "\\":
            cur += ch
            escaped = True
            continue
        if ch == '"':
            in_quotes = not in_quotes
            cur += ch
            continue
        if ch == "," and not in_quotes:
            out.append(cur)
            cur = ""
            continue
        cur += ch
    out.append(cur)
    return out


def literal(expr):
    """Склеить строковое выражение вида "a" + "b" в готовый текст."""
    pieces = re.findall(r'"((?:[^"\\]|\\.)*)"', expr)
    joined = "".join(pieces)
    return joined.replace("\\n", "\n").replace('\\"', '"').replace("\\\\", "\\")


catalog = {}
for m in call.finditer(text):
    args = split_args(m.group(2))
    if args:
        catalog[m.group(1)] = literal(args[0])

io.open(OUT, "w", encoding="utf-8").write(json.dumps(catalog, ensure_ascii=False, indent=1))
print("ключей:", len(catalog))
for key in ["hub.subtitle", "pdf.header.subtitle", "split.header.subtitle",
            "ocr.header.subtitle", "excel.header.subtitle", "menu.shortcuts"]:
    print(" ", key, "=", catalog.get(key))

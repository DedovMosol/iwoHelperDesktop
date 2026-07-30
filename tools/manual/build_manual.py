# -*- coding: utf-8 -*-
"""Сборка руководства пользователя iwo Helper Desktop в .docx — по-русски и по-английски.

Оформление живёт в engine.py, текст — в body_ru.py и body_en.py. Разделение не ради
красоты: два руководства обязаны описывать ОДНУ программу, и единственный способ этого
добиться — собирать их одним и тем же кодом из одинаково устроенного текста. Расхождение
устройства (пропущенный раздел, лишний рисунок, таблица не там) ловится проверкой ниже,
а не вычиткой двух документов подряд.

Usage:
    python build_manual.py --lang ru
    python build_manual.py --lang en
    python build_manual.py --both
"""
import argparse
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def build(lang):
    """Собрать руководство на одном языке ОТДЕЛЬНЫМ процессом.

    Отдельным — потому что и движок, и python-docx держат состояние документа в модулях:
    собрать в одном процессе два документа значит получить второй поверх первого.
    """
    env = dict(os.environ, MANUAL_LANG=lang, PYTHONIOENCODING="utf-8")
    code = subprocess.call([sys.executable, os.path.join(HERE, "_build_one.py")], cwd=HERE, env=env)
    if code:
        raise SystemExit("сборка (%s) завершилась с кодом %d" % (lang, code))


def compare_structure():
    """Оба текста описывают одну программу — значит и устроены одинаково.

    Сравниваем ПОСЛЕДОВАТЕЛЬНОСТЬ вызовов строительных блоков: заголовки, рисунки,
    таблицы, примечания. Слова, разумеется, разные; порядок и состав — обязаны совпадать.
    """
    import re
    shapes = {}
    for lang in ("ru", "en"):
        text = open(os.path.join(HERE, "body_%s.py" % lang), encoding="utf-8").read()
        calls = re.findall(r"\b(h1|h2|picture|table|note|footnote|step|bullet|listed)\(", text)
        figures = re.findall(r'picture\("([\w-]+)"', text)
        shapes[lang] = (calls, figures)
    if shapes["ru"][0] != shapes["en"][0]:
        ru, en = shapes["ru"][0], shapes["en"][0]
        for i in range(min(len(ru), len(en))):
            if ru[i] != en[i]:
                raise SystemExit("тексты разошлись устройством на блоке %d: ru=%s, en=%s"
                                 % (i + 1, ru[i], en[i]))
        raise SystemExit("тексты разошлись длиной: ru=%d блоков, en=%d" % (len(ru), len(en)))
    if shapes["ru"][1] != shapes["en"][1]:
        raise SystemExit("рисунки идут в разном порядке:\n  ru=%s\n  en=%s"
                         % (shapes["ru"][1], shapes["en"][1]))
    print("устройство совпадает: блоков %d, рисунков %d"
          % (len(shapes["ru"][0]), len(shapes["ru"][1])))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--lang", choices=("ru", "en"))
    parser.add_argument("--both", action="store_true")
    args = parser.parse_args()
    langs = ("ru", "en") if args.both or not args.lang else (args.lang,)
    if len(langs) > 1:
        compare_structure()
    for lang in langs:
        print("--- %s" % lang)
        build(lang)

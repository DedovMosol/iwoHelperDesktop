# -*- coding: utf-8 -*-
"""Проверка готового документа по требованиям заказчика.

Смотрим не на то, что мы собирались сделать, а на то, что реально лежит в файле:
шрифты, кегли, интервал, красная строка, наличие полей оглавления и номера страницы,
сносок, картинок, а также что каждый рисунок упомянут в тексте.
"""
import io
import os
import re
import zipfile

from docx import Document
from docx.shared import Pt

# Путь — от самого скрипта: конвейер лежит в репозитории и не должен зависеть от того,
# куда его склонировали.
PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                    "docs", "Инструкция пользователя.docx")
FONT = "Times New Roman"
problems = []
log = []


def check(condition, ok_text, bad_text):
    (log if condition else problems).append(ok_text if condition else bad_text)


doc = Document(PATH)

# --- стиль по умолчанию
normal = doc.styles["Normal"]
check(normal.font.name == FONT, "основной шрифт: " + str(normal.font.name),
      "основной шрифт не Times New Roman: %s" % normal.font.name)
check(normal.font.size == Pt(14), "основной кегль: 14 пт",
      "основной кегль не 14 пт: %s" % normal.font.size)
check(abs(normal.paragraph_format.first_line_indent.cm - 1.25) < 0.01,
      "красная строка: 1,25 см",
      "красная строка не задана: %s" % normal.paragraph_format.first_line_indent)
check(str(normal.paragraph_format.line_spacing_rule).startswith("ONE_POINT_FIVE"),
      "интервал: полуторный",
      "интервал не полуторный: %s" % normal.paragraph_format.line_spacing_rule)

def inherited(style, attribute):
    """Свойство шрифта с учётом наследования стилей.

    Прямое чтение style.font.name врёт: Word, пересохраняя файл, убирает из стиля
    явную гарнитуру, если она совпадает с унаследованной от «Обычного». На бумаге
    стоит Times New Roman, а проверка видит None и поднимает ложную тревогу.
    """
    while style is not None:
        value = getattr(style.font, attribute)
        if value is not None:
            return value
        style = style.base_style
    return None


for name in ("Heading 1", "Heading 2"):
    style = doc.styles[name]
    size, font = inherited(style, "size"), inherited(style, "name")
    check(size == Pt(16), "%s: 16 пт" % name, "%s не 16 пт: %s" % (name, size))
    check(font == FONT, "%s: Times New Roman" % name,
          "%s не Times New Roman: %s" % (name, font))

# --- ни один прогон текста не должен быть набран чужим шрифтом или чужим кеглем
alien_font, alien_size = set(), set()
for paragraph in doc.paragraphs:
    for run in paragraph.runs:
        if run.font.name not in (None, FONT):
            alien_font.add(run.font.name)
        if run.font.size is not None and run.font.size not in (Pt(14), Pt(16), Pt(20), Pt(28)):
            alien_size.add(run.font.size.pt)
check(not alien_font, "чужих шрифтов в тексте нет", "чужие шрифты: %s" % alien_font)
check(not alien_size, "чужих кеглей в тексте нет", "чужие кегли: %s" % alien_size)

# --- разметка пакета
with zipfile.ZipFile(PATH) as zf:
    names = zf.namelist()
    document = zf.read("word/document.xml").decode("utf-8")
    types = zf.read("[Content_Types].xml").decode("utf-8")
    rels = zf.read("word/_rels/document.xml.rels").decode("utf-8")
    settings = zf.read("word/settings.xml").decode("utf-8")
    footnotes = zf.read("word/footnotes.xml").decode("utf-8") if "word/footnotes.xml" in names else ""

check("word/footnotes.xml" in names, "часть сносок на месте", "нет части word/footnotes.xml")
check("footnotes+xml" in types, "тип содержимого сносок объявлен",
      "в [Content_Types].xml нет типа для сносок")
check("relationships/footnotes" in rels, "связь со сносками объявлена",
      "в связях документа нет ссылки на сноски")
notes = re.findall(r'<w:footnote [^>]*w:id="(\d+)"', footnotes)
refs = re.findall(r"<w:footnoteReference w:id=\"(\d+)\"", document)
real_notes = [n for n in notes if int(n) > 0]
check(len(real_notes) == len(refs) and len(refs) > 0,
      "сносок %d, ссылок на них %d" % (len(real_notes), len(refs)),
      "сноски и ссылки не сходятся: %d и %d" % (len(real_notes), len(refs)))

check("TOC \\o" in document, "поле автооглавления вставлено", "нет поля TOC")
footers = "".join(zipfile.ZipFile(PATH).read(n).decode("utf-8")
                  for n in names if n.startswith("word/footer"))
check(" PAGE " in footers, "поле номера страницы стоит в колонтитуле",
      "в колонтитулах нет поля PAGE")
# Оглавление должно быть СОБРАНО, а не только объявлено полем. Раньше здесь стояла
# проверка флага «пересчитать поля при открытии» — она зеленела на файле с пустым
# оглавлением и просьбой к читателю нажать F9. Word этот флаг к тому же снимает,
# исполнив его. Проверяем результат: пункты на месте, у каждого есть номер страницы,
# и набор пунктов совпадает с набором заголовков — иначе оглавление устарело.
toc_entries = [par for par in doc.paragraphs if par.style.name in ("toc 1", "toc 2")]
check(len(toc_entries) >= 10, "оглавление собрано: %d пунктов" % len(toc_entries),
      "оглавление пустое или обрезано: %d пунктов" % len(toc_entries))
numbered = [par.text for par in toc_entries if re.search(r"\t\d+$", par.text)]
check(len(numbered) == len(toc_entries),
      "у всех пунктов оглавления есть номер страницы",
      "пунктов без номера страницы: %d" % (len(toc_entries) - len(numbered)))
# Стили оглавления Word строит от «Обычного», у которого красная строка 1,25 см,
# поэтому отсутствие отступа надо задать ЯВНО: унаследованный отступ сдвигает
# каждую строку оглавления вправо, и на глаз это принимают за так задуманное.
ragged = [par for par in toc_entries
          if par.paragraph_format.first_line_indent is None
          or par.paragraph_format.first_line_indent.cm > 0.01]
check(not ragged, "пункты оглавления идут без красной строки",
      "пунктов оглавления с красной строкой: %d (наследуют отступ основного текста)" % len(ragged))
in_toc = set(re.sub(r"\t\d+$", "", par.text).strip() for par in toc_entries)
in_body = set(par.text.strip() for par in doc.paragraphs
              if par.style.name in ("Heading 1", "Heading 2"))
check(in_toc == in_body, "оглавление совпадает с заголовками (%d)" % len(in_body),
      "оглавление разошлось с заголовками: нет в оглавлении %s, лишние в оглавлении %s"
      % (sorted(in_body - in_toc)[:3], sorted(in_toc - in_body)[:3]))
check("<w:titlePg" in document, "у первой страницы свой колонтитул (без номера)",
      "титульный лист не отделён от нумерации")

# Сколько рисунков должно быть — спрашиваем у самой сборки (FIG_ORDER в engine.py), а не
# держим отдельную константу: разошедшись, она превращает проверку в источник ложной тревоги.
FIGURES = len(re.findall(r'"[a-z0-9-]+"', io.open(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "engine.py"),
    encoding="utf-8").read().split("FIG_ORDER = [")[1].split("]")[0]))
images = [n for n in names if n.startswith("word/media/")]
check(len(images) == FIGURES, "картинок внедрено: %d" % len(images),
      "картинок внедрено %d, ожидалось %d" % (len(images), FIGURES))

# --- каждый рисунок обязан быть упомянут в тексте
text = "\n".join(par.text for par in doc.paragraphs)
captions = re.findall(r"Рис\. (\d+)\.", text)
mentions = set(int(n) for n in re.findall(r"рис\. (\d+)", text))
missing = [n for n in captions if int(n) not in mentions]
check(not missing, "все %d рисунков упомянуты в тексте" % len(captions),
      "рисунки без ссылки в тексте: %s" % missing)
check([int(n) for n in captions] == list(range(1, len(captions) + 1)),
      "нумерация рисунков сплошная",
      "нумерация рисунков нарушена: %s" % captions)

table_caps = re.findall(r"Таблица (\d+) —", text)
check([int(n) for n in table_caps] == list(range(1, len(table_caps) + 1)),
      "таблиц: %d, нумерация сплошная" % len(table_caps),
      "нумерация таблиц нарушена: %s" % table_caps)
check(len(doc.tables) == len(table_caps),
      "подписей к таблицам столько же, сколько таблиц",
      "таблиц %d, подписей %d" % (len(doc.tables), len(table_caps)))

# --- каждая подпись в «кавычках» обязана существовать в каталоге Loc
# Инструкция называет кнопки и поля ровно так, как они подписаны на экране. Стоит
# переименовать кнопку в программе — и текст начинает врать, а заметить это глазами
# на сотне упоминаний невозможно. Здесь перечислено то, что кавычки несут НЕ как
# подпись элемента: проза, примеры и куски шаблонов имени.
import json
import subprocess

subprocess.call(["python", os.path.join(os.path.dirname(os.path.abspath(__file__)), "dump_loc.py")])
loc = json.load(io.open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "loc_ru.json"),
                        encoding="utf-8"))
known = set()
for value in loc.values():
    s = value.strip()
    known.add(s)
    known.add(s.replace("◀", "").replace("▶", "").replace("⌂", "").replace("☰", "")
               .replace("⟳", "").strip())
    known.add(s.rstrip(".…:").strip())
    known.add(s.replace("◀", "").replace("▶", "").replace("⌂", "").replace("☰", "")
               .replace("⟳", "").strip().rstrip(".…:").strip())
PROSE = [
    ".xls", "1-3", "8-", "Microsoft Print to PDF", "[BASENAME]_часть_[FILENUMBER###]",
    "Блокнот", "Договор_часть_001", "Договор_часть_002", "взяв рукой",
    "каждая страница отдельным файлом", "плывут", "поедут", "склеивания",
    "том на триста страниц, смотрите раздел четыре", "файл повреждён", "Пуск",
    "Хорошо", "Нормально", "Инструкция по работе с программой: открыть",
    # Обороты речи, а не подписи на экране: кавычки здесь цитируют мысль, а не кнопку.
    "Настройках", "страница целиком картинкой", "только текста", "у вас последняя",
]
lower = set(k.lower() for k in known)


def is_caption(quote):
    """Подпись знакома, если совпала точно, с точностью до регистра, начинает подпись
    из каталога («Как назвать части (необязательно)») или процитирована внутри неё
    («Добавить лист «Содержание»…»)."""
    if quote in known or quote.lower() in lower:
        return True
    for value in loc.values():
        v = value.strip()
        if v.startswith(quote) or ("«%s»" % quote) in v:
            return True
    return False


unknown = [q for q in sorted(set(re.findall("«([^»]{2,60})»", text)))
           if not is_caption(q) and q not in PROSE]
check(not unknown, "все подписи в кавычках есть в каталоге программы",
      "подписи, которых нет в программе: %s" % ", ".join(unknown))

# --- рабочих упоминаний в инструкции быть не должно
#
# Слова проверяем ПО ОТПЕЧАТКАМ, а не списком: список из тех самых слов — это ровно то,
# чему в открытом репозитории делать нечего. Отпечаток берётся от начала слова (3–14 букв),
# поэтому одна запись накрывает все склонения и производные.
import hashlib

BANNED_FINGERPRINTS = {
    "610a4b77c77b3ffd",
    "2c0cda61e536f2c5",
    "fe4e39d2d1e3c6b5",
    "2821bc2655a8343f",
    "d74a2fa10ca4be11",
    "f3dcd5e6ce37de3e",
    "dab70bc2d883a287"
}


def _fingerprint(word):
    return hashlib.sha256(word.encode("utf-8")).hexdigest()[:16]


found = []
for raw in re.findall(r"[^\W\d_]{3,}", text.lower(), re.UNICODE):
    for size in range(3, min(len(raw), 14) + 1):
        if _fingerprint(raw[:size]) in BANNED_FINGERPRINTS:
            found.append(raw)
            break
check(not found, "рабочих упоминаний нет", "найдены рабочие упоминания: %s" % sorted(set(found)))

headings = [par.text for par in doc.paragraphs if par.style.name.startswith("Heading")]
print("Заголовков: %d" % len(headings))
for line in log:
    print("  ок    ", line)
for line in problems:
    print("  ПЛОХО ", line)
print()
print("итог:", "ВСЁ СОШЛОСЬ" if not problems else "ЕСТЬ ЗАМЕЧАНИЯ (%d)" % len(problems))

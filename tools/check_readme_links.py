# -*- coding: utf-8 -*-
"""Проверка ссылок README: внешние адреса, якоря разделов и пути в репозитории.

Зачем отдельная проверка. Битая ссылка в README — это первое, что видит человек,
пришедший на страницу проекта, и единственное, что никто никогда не перечитывает.
Три вида отказов ловятся здесь:

* **внешние адреса** — ссылки на файлы релиза прибиты к версии и после выпуска
  перестают отвечать, если ассет назван иначе;
* **якоря разделов** — заголовок с эмодзи даёт якорь, в котором остаётся
  невидимый селектор начертания (U+FE0F), и ссылка «#-download» молча ведёт
  в никуда, выглядя при этом правильной;
* **пути в репозитории** — переименованный документ оставляет ссылку, которая
  открывается страницей «404» уже внутри проекта.

Usage: python tools/check_readme_links.py [README.md]
Код возврата: 0 — всё цело, 1 — есть битые ссылки.
"""
import io
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) readme-link-check"}
TIMEOUT = 30


def anchor_of(heading):
    """Якорь, который GitHub делает из заголовка.

    Правило GitHub: в нижний регистр, пробелы в дефисы, выбрасывается всё, кроме
    букв, цифр, дефисов и подчёркиваний. Эмодзи выбрасывается, а вот селектор
    начертания U+FE0F, идущий за ним, — НЕТ: он не буква и не знак препинания,
    поэтому остаётся в якоре невидимым символом.
    """
    text = heading.strip().lower()
    text = re.sub(r"[ \t]+", "-", text)
    return "#" + "".join(ch for ch in text if ch.isalnum() or ch in "-_️")


def assembly_version(root):
    path = os.path.join(root, "src", "AssemblyInfo.cs")
    try:
        source = io.open(path, encoding="utf-8").read()
    except OSError:
        return None
    match = re.search(r'AssemblyFileVersion\("([0-9]+\.[0-9]+\.[0-9]+)', source)
    return match.group(1) if match else None


def is_pending_asset_url(target, version):
    if not version:
        return False
    prefix = "https://github.com/DedovMosol/iwoHelperDesktop/releases/download/v%s/" % version
    if not target.startswith(prefix):
        return False
    name = target[len(prefix):]
    return name in {
        "iwoHelperDesktop-%s.exe" % version,
        "iwoHelperDesktop-%s-x86.exe" % version,
        "iwoHelperDesktop-setup-%s.exe" % version,
        "iwoHelperDesktop-setup-%s-x86.exe" % version,
    }


def release_exists(version):
    url = "https://api.github.com/repos/DedovMosol/iwoHelperDesktop/releases/tags/v%s" % version
    try:
        with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=TIMEOUT):
            return True
    except urllib.error.HTTPError as error:
        return False if error.code == 404 else None
    except Exception:
        return None


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "README.md"
    root = os.path.dirname(os.path.abspath(path)) or "."
    version = assembly_version(root)
    published = None
    published_checked = False
    text = io.open(path, encoding="utf-8").read()

    anchors = set()
    for level, heading in re.findall(r"^(#{1,6})\s+(.+?)\s*$", text, re.M):
        anchors.add(anchor_of(heading))

    # Ссылки вида [подпись](адрес) и картинки ![подпись](адрес)
    links = re.findall(r"(?<!\!)\[[^\]]*\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)", text)
    images = re.findall(r"!\[[^\]]*\]\(([^)\s]+)\)", text)
    html_src = re.findall(r'<img[^>]+src="([^"]+)"', text)

    # Отказы делятся надвое. «Нет такой страницы» — наш дефект: ссылку правим.
    # Таймаут, 429 от лимита обращений или 5xx на стороне сервиса — не дефект
    # README, и ронять из-за них сборку значит приучить смотреть на красный цвет
    # как на норму. Поэтому такие случаи попадают в предупреждения.
    broken, flaky, checked = [], [], 0
    seen = set()
    for target in links + images + html_src:
        if target in seen:
            continue
        seen.add(target)
        checked += 1
        if target.startswith("#"):
            if target not in anchors:
                broken.append("якорь не найден: %s" % target)
        elif target.startswith(("http://", "https://")):
            try:
                # Адрес с кириллицей (например, имя файла в релизе) валит запрос
                # ошибкой кодировки вместо честного ответа сервера, поэтому путь
                # приводится к процентной записи, а уже закодированное не трогаем.
                safe = urllib.parse.quote(target, safe="/:?&=#%+,;@!$'()*~")
                request = urllib.request.Request(safe, headers=UA, method="GET")
                with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
                    if response.status >= 400:
                        broken.append("%s -> HTTP %s" % (target, response.status))
            except urllib.error.HTTPError as error:
                if error.code == 404 and is_pending_asset_url(target, version):
                    if not published_checked:
                        published = release_exists(version)
                        published_checked = True
                    if published is False:
                        flaky.append("%s -> pending release v%s (name is valid)" %
                                     (target, version))
                    elif published is None:
                        flaky.append("%s -> HTTP 404; release status unavailable" % target)
                    else:
                        broken.append("%s -> HTTP 404 (release exists, asset missing)" % target)
                else:
                    where = broken if error.code in (404, 410) else flaky
                    where.append("%s -> HTTP %s" % (target, error.code))
            except Exception as error:
                flaky.append("%s -> %s" % (target, error))
        elif target.startswith("mailto:"):
            pass
        else:
            local = os.path.join(root, target.split("#")[0])
            if not os.path.exists(local):
                broken.append("нет файла: %s" % target)

    print("проверено ссылок: %d" % checked)
    for item in broken:
        print("  БИТАЯ  " + item)
    for item in flaky:
        print("  не ответил (не ошибка README)  " + item)
    print("итог:", "все ссылки живы" if not broken else "битых ссылок: %d" % len(broken))
    return 1 if broken else 0


if __name__ == "__main__":
    sys.exit(main())

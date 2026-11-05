# -*- coding: utf-8 -*-
"""Проверка ссылок README: внешние адреса, якоря разделов и пути в репозитории.

Зачем отдельная проверка. Битая ссылка в README — это первое, что видит человек,
пришедший на страницу проекта, и единственное, что никто никогда не перечитывает.
Три вида отказов ловятся здесь:

* **внешние адреса** — ссылки на файлы релиза прибиты к версии и после выпуска
  перестают отвечать, если ассет назван иначе. Файлы релиза проверяются
  через GitHub API (имена ассетов), не GET `/releases/download/` — иначе
  каждый push CI рисует ложные скачивания;
* **якоря разделов** — заголовок с эмодзи даёт якорь, в котором остаётся
  невидимый селектор начертания (U+FE0F), и ссылка «#-download» молча ведёт
  в никуда, выглядя при этом правильной;
* **пути в репозитории** — переименованный документ оставляет ссылку, которая
  открывается страницей «404» уже внутри проекта.

Usage: python tools/check_readme_links.py [README.md]
Код возврата: 0 — всё цело, 1 — есть битые ссылки.
"""
import io
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) readme-link-check"}
TIMEOUT = 30
# GET /releases/download/… считает GitHub как скачивание ассета. Проверка
# живёт через API релиза (список имён) и HEAD для прочих URL.
_RELEASE_DOWNLOAD = re.compile(
    r"^https://github\.com/([^/]+)/([^/]+)/releases/download/(v[^/]+)/([^/?#]+)$")
_release_cache = {}


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
    # Имена ассетов — как их называет tools/make_release.ps1. Portable — zip: exe работает
    # в паре со своим .config (режим DPI), а одиночный portable-exe больше не выпускается.
    return name in {
        "iwoHelperDesktop-%s.zip" % version,
        "iwoHelperDesktop-%s-x86.zip" % version,
        "iwoHelperDesktop-setup-%s.exe" % version,
        "iwoHelperDesktop-setup-%s-x86.exe" % version,
    }


def github_release_download(url):
    """owner, repo, tag, filename из URL ассета, или None — не github download."""
    match = _RELEASE_DOWNLOAD.match(url)
    if not match:
        return None
    owner, repo, tag, name = match.groups()
    return owner, repo, tag, urllib.parse.unquote(name)


def fetch_release_assets(owner, repo, tag):
    """('ok', names) | ('missing', []) | ('error', []). Кэш на прогон. Не качает файлы."""
    key = (owner, repo, tag)
    if key in _release_cache:
        return _release_cache[key]
    url = "https://api.github.com/repos/%s/%s/releases/tags/%s" % (owner, repo, tag)
    try:
        request = urllib.request.Request(url, headers=UA, method="GET")
        with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
            data = json.loads(response.read().decode("utf-8"))
            names = [asset.get("name") for asset in data.get("assets") or [] if asset.get("name")]
            result = ("ok", names)
    except urllib.error.HTTPError as error:
        result = ("missing", []) if error.code == 404 else ("error", [])
    except Exception:
        result = ("error", [])
    _release_cache[key] = result
    return result


def extract_targets(text):
    """Все адреса README в порядке появления: ссылки, цели ссылок-бейджей, картинки.

    Ссылка-бейдж — это картинка внутри ссылки: [![подпись](картинка)](цель). Обычное
    выражение для ссылки видит в ней только первую пару скобок, то есть адрес
    КАРТИНКИ, а цель не проверялась вовсе. Так месяцами жили битыми кнопки
    скачивания и бейджи шапки: картинка с shields.io отвечает всегда, файл релиза —
    нет. Поэтому цели бейджей вынимаются отдельно.
    """
    badge_targets = re.findall(
        r"\[!\[[^\]]*\]\([^)\s]+(?:\s+\"[^\"]*\")?\)\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)", text)
    links = re.findall(r"(?<!\!)\[[^\]]*\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)", text)
    images = re.findall(r"!\[[^\]]*\]\(([^)\s]+)\)", text)
    html_src = re.findall(r'<img[^>]+src="([^"]+)"', text)
    return badge_targets + links + images + html_src


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "README.md"
    root = os.path.dirname(os.path.abspath(path)) or "."
    version = assembly_version(root)
    text = io.open(path, encoding="utf-8").read()

    anchors = set()
    for level, heading in re.findall(r"^(#{1,6})\s+(.+?)\s*$", text, re.M):
        anchors.add(anchor_of(heading))

    # Отказы делятся надвое. «Нет такой страницы» — наш дефект: ссылку правим.
    # Таймаут, 429 от лимита обращений или 5xx на стороне сервиса — не дефект
    # README, и ронять из-за них сборку значит приучить смотреть на красный цвет
    # как на норму. Поэтому такие случаи попадают в предупреждения.
    broken, flaky, checked = [], [], 0
    seen = set()
    for target in extract_targets(text):
        if target in seen:
            continue
        seen.add(target)
        checked += 1
        if target.startswith("#"):
            if target not in anchors:
                broken.append("якорь не найден: %s" % target)
        elif target.startswith(("http://", "https://")):
            try:
                asset = github_release_download(target)
                if asset is not None:
                    owner, repo, tag, filename = asset
                    status, names = fetch_release_assets(owner, repo, tag)
                    if status == "ok":
                        if filename not in names:
                            # Релиз уже есть, ассет ещё грузится (tag CI стартует
                            # сразу после push тега, upload setup.exe идёт следом).
                            if is_pending_asset_url(target, version):
                                flaky.append("%s -> pending asset on %s" % (target, tag))
                            else:
                                broken.append("%s -> asset not on %s" % (target, tag))
                    elif status == "missing":
                        if is_pending_asset_url(target, version):
                            flaky.append("%s -> pending release v%s (name is valid)" %
                                         (target, version))
                        else:
                            broken.append("%s -> no release %s" % (target, tag))
                    else:
                        flaky.append("%s -> release API unavailable" % target)
                    continue
                # Адрес с кириллицей валит запрос ошибкой кодировки; путь в проценты.
                # HEAD: тело не качается. GET на /releases/download/ накручивает счётчик.
                safe = urllib.parse.quote(target, safe="/:?&=#%+,;@!$'()*~")
                request = urllib.request.Request(safe, headers=UA, method="HEAD")
                with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
                    if response.status >= 400:
                        broken.append("%s -> HTTP %s" % (target, response.status))
            except urllib.error.HTTPError as error:
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

# -*- coding: utf-8 -*-
"""Проверка метаданных коммитов: в истории только автор проекта и ничьих подписей-соавторов.

Зачем. Список участников на странице проекта GitHub строит из авторов коммитов и из
трейлеров «Co-authored-by», а исправить его потом можно только переписав историю — и даже
после этого он долго держится в кэше. Дешевле не пустить такое в историю вообще.

Один источник правил для двух мест:

* **локальный хук** ``tools/githooks/commit-msg`` вызывает ``--commit-msg FILE``: убирает
  из сообщения строки атрибуции (это нормализация сообщения, ради которой хук и существует)
  и отказывает, если автор или коммитер — не автор проекта;
* **CI** вызывает без аргументов: вся достижимая история и все теги, любое нарушение —
  код возврата 1 со списком коммитов.

Usage:
  python tools/check_commit_metadata.py [--repo DIR] [REV ...]   # по умолчанию --all
  python tools/check_commit_metadata.py --commit-msg FILE        # режим хука commit-msg
Код возврата: 0 — чисто, 1 — нарушения, 2 — ошибка запуска.
"""
import io
import re
import subprocess
import sys

# Адреса автора проекта: основной и служебный адрес GitHub того же аккаунта (им подписаны
# правки, сделанные в веб-интерфейсе).
AUTHOR_EMAILS = frozenset([
    "andreyid@outlook.com",
    "73111581+dedovmosol@users.noreply.github.com",
])
# Коммитер правки из веб-интерфейса — сам GitHub; автор при этом всё равно владелец.
COMMITTER_EMAILS = AUTHOR_EMAILS | frozenset(["noreply@github.com"])

# Строки атрибуции: трейлер соавтора, ссылка на сеанс инструмента и подпись «сделано с …»,
# которую инструменты начинают значком. Всё привязано к началу строки, а подпись — ещё и к
# значку перед ней: «Generated files updated» в обычном сообщении — это текст, не подпись.
ATTRIBUTION_LINE = re.compile(
    r"^[ \t]*(?:"
    r"co-authored-by[ \t]*:"
    r"|[a-z0-9-]*-session[ \t]*:[ \t]*https?://"
    r"|[^\w\s]{1,4}[ \t]*generated (?:with|by)\b"
    r")",
    re.IGNORECASE)

FIELD = "\x1f"
RECORD = "\x1e"


def attribution_lines(message):
    """Строки сообщения, которые являются атрибуцией (в порядке появления)."""
    return [line for line in message.splitlines() if ATTRIBUTION_LINE.match(line)]


def strip_attribution(message):
    """Сообщение без строк атрибуции и без пустых строк, оставшихся от них в конце.

    Сам текст не трогаем: перевод строки в конце и пустые строки внутри сохраняются,
    иначе хук переписывал бы чистые сообщения.
    """
    if not attribution_lines(message):
        return message
    kept = [line for line in message.splitlines() if not ATTRIBUTION_LINE.match(line)]
    while kept and not kept[-1].strip():
        kept.pop()
    return "\n".join(kept) + "\n"


def is_author(email):
    return (email or "").strip().lower() in AUTHOR_EMAILS


def is_committer(email):
    return (email or "").strip().lower() in COMMITTER_EMAILS


def _git(repo, *args):
    result = subprocess.run(["git"] + list(args), cwd=repo, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE)
    if result.returncode != 0:
        # Только имя подкоманды: в аргументах формата — управляющие разделители полей.
        raise RuntimeError("git %s: %s" % (args[0], result.stderr.decode("utf-8", "replace").strip()))
    return result.stdout.decode("utf-8", "replace")


def history_violations(repo=".", revs=None):
    """Нарушения в коммитах revs (по умолчанию — вся достижимая история) и в тегах."""
    revs = list(revs) if revs else ["--all"]
    out = _git(repo, "log", "--format=%H" + FIELD + "%ae" + FIELD + "%ce" + FIELD + "%B" + RECORD, *revs)
    problems = []
    for record in out.split(RECORD):
        record = record.strip("\n")
        if not record:
            continue
        sha, author, committer, message = record.split(FIELD, 3)
        short = sha[:10]
        if not is_author(author):
            problems.append("%s: автор %s — не автор проекта" % (short, author))
        if not is_committer(committer):
            problems.append("%s: коммитер %s — не автор проекта" % (short, committer))
        for line in attribution_lines(message):
            problems.append("%s: строка атрибуции в сообщении: %s" % (short, line.strip()))
    tags = _git(repo, "for-each-ref", "refs/tags",
                "--format=%(refname:short)" + FIELD + "%(objecttype)" + FIELD + "%(taggeremail)")
    for line in tags.splitlines():
        name, kind, tagger = line.split(FIELD)
        if kind == "tag" and not is_author(tagger.strip("<>")):
            problems.append("тег %s: автор тега %s — не автор проекта" % (name, tagger))
    return problems


def _ident_email(repo, var):
    ident = _git(repo, "var", var).strip()
    match = re.search(r"<([^>]*)>", ident)
    return match.group(1) if match else ""


def commit_msg_hook(path, repo="."):
    """Режим хука: нормализует сообщение и проверяет, от чьего имени делается коммит."""
    problems = []
    for var, check in (("GIT_AUTHOR_IDENT", is_author), ("GIT_COMMITTER_IDENT", is_committer)):
        email = _ident_email(repo, var)
        if not check(email):
            problems.append("%s: %s — не автор проекта" % (var, email or "(не задан)"))
    if problems:
        return problems
    with io.open(path, "r", encoding="utf-8", newline="") as source:
        message = source.read()
    cleaned = strip_attribution(message)
    if cleaned != message:
        with io.open(path, "w", encoding="utf-8", newline="") as target:
            target.write(cleaned)
    return []


def main(argv):
    try:
        if len(argv) == 2 and argv[0] == "--commit-msg":
            problems = commit_msg_hook(argv[1])
        else:
            repo = "."
            if len(argv) >= 2 and argv[0] == "--repo":
                repo, argv = argv[1], argv[2:]
            problems = history_violations(repo, argv)
    except (OSError, RuntimeError) as error:
        sys.stderr.write("check_commit_metadata: %s\n" % error)
        return 2
    for problem in problems:
        sys.stderr.write(problem + "\n")
    if problems:
        sys.stderr.write("Коммиты проекта делает только его автор, без соавторов (CONTRIBUTING.md).\n")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

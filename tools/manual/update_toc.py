# -*- coding: utf-8 -*-
"""Пересобрать ОГЛАВЛЕНИЕ в готовом .docx, не трогая больше ничего.

Отдельно от сборки руководства: документ бывает доправлен руками, и пересобирать его
ради оглавления значило бы стереть правку. Здесь файл открывается как есть, Word
обновляет поля и сохраняет — содержимое остаётся тем же, меняются только пункты
оглавления и номера страниц.

Word водится из PowerShell (pywin32 в окружении нет), а путь ему передаётся ТОЛЬКО
ASCII: Windows PowerShell 5.1 портит не-ASCII, а имена файлов у нас и русские, и с
пробелами. Поэтому работаем по копии со служебным именем и возвращаем её на место
средствами Python.

Usage: python update_toc.py "путь\\к\\файлу.docx" [ещё файлы...]
"""
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def update(path):
    if not os.path.isfile(path):
        raise SystemExit("нет файла: " + path)
    before = os.path.getsize(path)
    ascii_copy = os.path.join(HERE, "_toc_update.docx")
    shutil.copyfile(path, ascii_copy)
    try:
        done = subprocess.run(
            ["powershell", "-NoProfile", "-File", os.path.join(HERE, "finalize_toc.ps1"),
             "-Path", ascii_copy],
            capture_output=True, text=True)
        out = (done.stdout or "") + (done.stderr or "")
        # Молчаливый провал здесь означает устаревшее оглавление в готовом документе,
        # поэтому проверяем и код возврата, и метку успеха: нулевой код сам по себе
        # ничего не доказывает, если Word не открылся.
        if done.returncode != 0 or "ok" not in out:
            raise SystemExit("оглавление собрать не удалось (код %s):\n%s" % (done.returncode, out.strip()))
        shutil.move(ascii_copy, path)
        marks = " ".join(w for w in out.split() if w.startswith(("pages=", "toc_entries=")))
        print("%s: %s (было %d Б, стало %d Б)"
              % (os.path.basename(path), marks, before, os.path.getsize(path)))
    finally:
        if os.path.exists(ascii_copy):
            os.remove(ascii_copy)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    for target in sys.argv[1:]:
        update(target)

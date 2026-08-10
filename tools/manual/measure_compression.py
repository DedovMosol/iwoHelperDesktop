# -*- coding: utf-8 -*-
"""Реальный замер сжатия для диаграммы в инструкции.

Берём документ со сканоподобными изображениями и прогоняем его тем же движком и
с теми же настройками, что использует программа, — цифры на диаграмме должны быть
измеренными, а не выдуманными.
"""
import io
import os
import random
import subprocess

from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(HERE, "compress")
os.makedirs(WORK, exist_ok=True)
# Ghostscript ИЩЕМ, а не прописываем: путь с чужой машины несёт на гит имя учётной
# записи и всё равно не совпадёт ни с одной другой установкой.
def _find_gs():
    from glob import glob
    found = sorted(glob(r"C:\Program Files\gs\gs*\bin\gswin64c.exe")) \
        + sorted(glob(os.path.expanduser(r"~\gs*\bin\gswin64c.exe")))
    if not found:
        raise SystemExit("не найден gswin64c.exe — укажите путь переменной GS_EXE")
    return found[-1]


GS = os.environ.get("GS_EXE") or _find_gs()

# «Скан» страницы: серый фон с зерном и текстом — так выглядит документ, снятый сканером,
# и именно на таких файлах сжатие даёт заметный эффект.
random.seed(7)
font = ImageFont.truetype(r"C:\Windows\Fonts\times.ttf", 34)
pages = []
for n in range(1, 5):
    img = Image.new("RGB", (2480, 3508), (247, 246, 242))  # A4 при 300 dpi — типичный скан
    px = img.load()
    for _ in range(200000):
        x, y = random.randrange(2480), random.randrange(3508)
        g = random.randrange(180, 240)
        px[x, y] = (g, g, g - 4)
    d = ImageDraw.Draw(img)
    d.text((220, 260), "Образец отсканированной страницы %d" % n, font=font, fill=(25, 25, 30))
    for line in range(40):
        width = random.randrange(1000, 2000)
        y = 420 + line * 74
        d.rectangle([220, y, 220 + width, y + 30], fill=(70 + line % 12, 70, 78))
    path = os.path.join(WORK, "scan%d.jpg" % n)
    img.save(path, quality=92)
    pages.append(path)

src = os.path.join(WORK, "Скан документа.pdf")
c = canvas.Canvas(src, pagesize=A4)
w, h = A4
for path in pages:
    c.drawImage(path, 10 * mm, 10 * mm, width=w - 20 * mm, height=h - 20 * mm)
    c.showPage()
c.save()


def compress(preset, out, downsample=True):
    """Те же аргументы, что строит программа (PdfCompression.BuildArguments).

    downsample=False — уровень «Очень хорошо»: пресет изображения не пересчитывает,
    и программа подпирает это обещание тремя ключами явно. Замер обязан идти с ними,
    иначе диаграмма покажет не то, что получит человек.
    """
    args = [GS, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
            "-dPDFSETTINGS=" + preset]
    if not downsample:
        args += ["-dDownsampleColorImages=false", "-dDownsampleGrayImages=false",
                 "-dDownsampleMonoImages=false"]
    args += ["-dNOPAUSE", "-dBATCH", "-dQUIET", "-dSAFER", "-sOutputFile=" + out, src]
    subprocess.run(args, check=True, capture_output=True)
    return os.path.getsize(out)


base = os.path.getsize(src)
verygood = compress("/default", os.path.join(WORK, "verygood.pdf"), downsample=False)
good = compress("/ebook", os.path.join(WORK, "good.pdf"))
small = compress("/screen", os.path.join(WORK, "small.pdf"))

# Порядок — как в окне программы (PdfCompression.Order), он же порядок столбиков.
result = {"none": base, "verygood": verygood, "good": good, "small": small}
io.open(os.path.join(HERE, "compression.txt"), "w", encoding="utf-8").write(
    "\n".join("%s=%d" % (k, v) for k, v in result.items()))
for name, size in result.items():
    print("%-6s %8.2f MB  %5.0f%%" % (name, size / 1048576.0, 100.0 * size / base))

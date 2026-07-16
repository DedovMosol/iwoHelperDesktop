using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

/// <summary>
/// Генератор app.ico из дизайна assets/icon.svg (GDI+ повторяет ту же графику).
/// Запуск: tools\make_icon.cmd. Результат коммитится в репозиторий,
/// перегенерация нужна только при смене дизайна.
/// </summary>
internal static class MakeIcon
{
    private static readonly Color BgLight = Color.FromArgb(33, 163, 102);  // #21A366
    private static readonly Color BgDark = Color.FromArgb(11, 92, 46);     // #0B5C2E
    private static readonly Color Head = Color.FromArgb(16, 124, 65);      // #107C41
    private static readonly Color Grid = Color.FromArgb(154, 200, 172);    // #9AC8AC

    private static int Main(string[] args)
    {
        string outPath = args.Length > 0 ? args[0] : "app.ico";
        int[] sizes = { 16, 20, 24, 32, 40, 48, 64, 128, 256 };

        var pngs = new List<byte[]>();
        foreach (int s in sizes)
            pngs.Add(RenderPng(s));

        // Формат ICO: заголовок, каталог, затем PNG-кадры (Windows Vista+).
        using (var w = new BinaryWriter(File.Create(outPath)))
        {
            w.Write((short)0);              // reserved
            w.Write((short)1);              // type: icon
            w.Write((short)sizes.Length);
            int offset = 6 + 16 * sizes.Length;
            for (int i = 0; i < sizes.Length; i++)
            {
                int s = sizes[i];
                w.Write((byte)(s >= 256 ? 0 : s)); // width (0 = 256)
                w.Write((byte)(s >= 256 ? 0 : s)); // height
                w.Write((byte)0);                  // palette
                w.Write((byte)0);                  // reserved
                w.Write((short)1);                 // planes
                w.Write((short)32);                // bpp
                w.Write(pngs[i].Length);
                w.Write(offset);
                offset += pngs[i].Length;
            }
            for (int i = 0; i < pngs.Count; i++)
                w.Write(pngs[i]);
        }
        Console.WriteLine("ICO OK: " + Path.GetFullPath(outPath));
        return 0;
    }

    private static byte[] RenderPng(int size)
    {
        using (var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb))
        {
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                Draw(g, size);
            }
            using (var ms = new MemoryStream())
            {
                bmp.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }
    }

    private static void Draw(Graphics g, int size)
    {
        float k = size / 256f;

        using (GraphicsPath bg = Rounded(16 * k, 16 * k, 224 * k, 224 * k, 44 * k))
        using (var grad = new LinearGradientBrush(
            new PointF(16 * k, 16 * k), new PointF(240 * k, 240 * k), BgLight, BgDark))
        {
            g.FillPath(grad, bg);
        }

        if (size <= 24)
        {
            // Мелкие размеры: только лист с шапкой, детали не читаются.
            DrawCard(g, k, 56, 56, 144, 144, 1f, true);
            return;
        }

        // Стопка листов позади: «многие файлы» — в один.
        DrawCard(g, k, 88, 44, 128, 112, 0.28f, false);
        DrawCard(g, k, 76, 58, 128, 112, 0.50f, false);
        // Итоговый лист с шапкой и сеткой.
        DrawCard(g, k, 64, 72, 128, 112, 1f, true);
        using (var p = new Pen(Grid, Math.Max(1f, 4 * k)))
        {
            p.StartCap = LineCap.Round;
            p.EndCap = LineCap.Round;
            g.DrawLine(p, 70 * k, 129 * k, 186 * k, 129 * k);
            g.DrawLine(p, 70 * k, 156 * k, 186 * k, 156 * k);
            g.DrawLine(p, 128 * k, 108 * k, 128 * k, 178 * k);
        }
    }

    private static void DrawCard(Graphics g, float k, float x, float y, float w, float h,
        float opacity, bool withHeader)
    {
        int alpha = (int)(255 * opacity);
        using (GraphicsPath card = Rounded(x * k, y * k, w * k, h * k, 12 * k))
        using (var b = new SolidBrush(Color.FromArgb(alpha, Color.White)))
        {
            g.FillPath(b, card);
        }
        if (!withHeader)
            return;
        float headH = h * 0.27f;
        using (GraphicsPath head = RoundedTop(x * k, y * k, w * k, headH * k, 12 * k))
        using (var hb = new SolidBrush(Head))
        {
            g.FillPath(hb, head);
        }
    }

    private static GraphicsPath Rounded(float x, float y, float w, float h, float r)
    {
        var p = new GraphicsPath();
        float d = r * 2;
        p.AddArc(x, y, d, d, 180, 90);
        p.AddArc(x + w - d, y, d, d, 270, 90);
        p.AddArc(x + w - d, y + h - d, d, d, 0, 90);
        p.AddArc(x, y + h - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }

    private static GraphicsPath RoundedTop(float x, float y, float w, float h, float r)
    {
        var p = new GraphicsPath();
        float d = r * 2;
        p.AddArc(x, y, d, d, 180, 90);
        p.AddArc(x + w - d, y, d, d, 270, 90);
        p.AddLine(x + w, y + h, x, y + h);
        p.CloseFigure();
        return p;
    }
}

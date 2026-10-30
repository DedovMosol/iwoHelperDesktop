using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Надпись под углом (WP-3): боковой текст полей, вертикальная шапка столбца, водяной знак
    /// по диагонали. В поток абзацев она не идёт (строки разных направлений не складываются в
    /// один абзац) и пишется отдельной повёрнутой надписью: в Word — плавающей надписью за
    /// текстом, в PowerPoint — надписью с поворотом. Координаты — страницы PDF (ось Y вверх),
    /// уже после пользовательского поворота.
    /// </summary>
    internal sealed class OcrRotatedText
    {
        public string Text;
        public double CenterXPt, CenterYPt; // центр рамки надписи
        public double LengthPt;             // длина строки вдоль базовой линии
        public double FontSizePt;
        public double AngleDeg;             // направление базовой линии против часовой, 0…360
        public bool Bold, Italic;
        public int ColorArgb;
        public string FontName;
    }

    /// <summary>Буква повёрнутой строки, снятая с PdfPig: начало и конец базовой линии и формат.</summary>
    internal struct RotatedLetter
    {
        public string Value;
        public double StartX, StartY, EndX, EndY;
        public double SizePt;
        public string FontName;
        public bool Bold, Italic;
        public int ColorArgb;
    }

    /// <summary>
    /// Сборка повёрнутых букв в надписи и геометрия их рамок — общая для обоих писателей
    /// (DRY: Word и PowerPoint ставят одну и ту же рамку и один и тот же поворот). Чистая
    /// логика без PdfPig — под тест.
    /// </summary>
    internal static class RotatedText
    {
        // Буква продолжает строку, если её направление отличается не больше чем на 3° и она
        // начата там, где кончилась предыдущая: поперёк — не дальше 0,3 кегля, вдоль — от
        // небольшого наложения кернинга до 2,5 кегля (больше — уже другая надпись). Зазор
        // вдоль больше 0,2 кегля без буквы-пробела — пробел между словами.
        private const double MaxTurnDeg = 3;
        private const double MaxAcrossEm = 0.3;
        private const double MaxOverlapEm = 0.3;
        private const double MaxGapEm = 2.5;
        private const double WordGapEm = 0.2;

        // Рамка надписи: высота — 1,2 кегля, базовая линия на 0,9 кегля ниже верха (как у
        // обычной строки), длина — с запасом на метрики чужого шрифта, чтобы строка не
        // переносилась; надпись пишется без переноса, и запас её не сдвигает: рамка
        // поворачивается вокруг центра, а текст прижат к её началу.
        internal const double FrameHeightEm = 1.2;
        internal const double BaselineFromTopEm = 0.9;
        internal const double LengthSlack = 0.15;

        public static double FrameHeight(OcrRotatedText t) { return FrameHeightEm * t.FontSizePt; }

        public static double FrameWidth(OcrRotatedText t) { return t.LengthPt * (1 + LengthSlack) + t.FontSizePt; }

        /// <summary>Поворот рамки по часовой (так считают Word и PowerPoint): 0…360. Чистая — под тест.</summary>
        public static double ClockwiseDegrees(OcrRotatedText t)
        {
            return Normalize(-t.AngleDeg);
        }

        internal static double Normalize(double deg)
        {
            double d = deg % 360;
            if (d < 0) d += 360;
            return d >= 359.95 ? 0 : d;
        }

        /// <summary>
        /// Левый верхний угол НЕповёрнутой рамки (pt, ось Y вверх): рамка той же длины и высоты
        /// с тем же центром, повёрнутая писателем вокруг центра, ложится на надпись.
        /// Рамка начинается у начала строки: центр сдвинут вдоль строки на половину запаса.
        /// </summary>
        public static void FrameTopLeft(OcrRotatedText t, out double left, out double top)
        {
            double w = FrameWidth(t), h = FrameHeight(t);
            left = t.CenterXPt - w / 2;
            top = t.CenterYPt + h / 2;
        }

        /// <summary>
        /// Буквы → надписи. Буквы идут в порядке потока содержимого: продолжение строки решает
        /// соседство с предыдущей буквой (см. константы выше). Пробелы и пустые строки по краям
        /// снимаются; надпись только из знаков частной области Юникода отбрасывается, как и в
        /// потоке. Центр рамки — середина базовой линии, смещённая «вверх» строки так, чтобы
        /// базовая линия легла на 0,9 кегля ниже верха рамки.
        /// </summary>
        public static List<OcrRotatedText> Group(IList<RotatedLetter> letters)
        {
            var result = new List<OcrRotatedText>();
            var run = new List<RotatedLetter>();
            var text = new StringBuilder();
            double dirX = 0, dirY = 0;
            double cosTurn = Math.Cos(MaxTurnDeg * Math.PI / 180);
            foreach (RotatedLetter l in letters)
            {
                double dx = l.EndX - l.StartX, dy = l.EndY - l.StartY;
                double len = Math.Sqrt(dx * dx + dy * dy);
                bool space = string.IsNullOrWhiteSpace(l.Value);
                if (!(len >= 0) || double.IsInfinity(len))
                    continue; // NaN/бесконечность из битой матрицы — не буква
                if (run.Count > 0)
                {
                    RotatedLetter p = run[run.Count - 1];
                    double gx = l.StartX - p.EndX, gy = l.StartY - p.EndY;
                    double along = gx * dirX + gy * dirY, across = -gx * dirY + gy * dirX;
                    double em = Math.Max(Math.Max(p.SizePt, l.SizePt), 1);
                    bool sameDir = len < 1e-6 || (dx * dirX + dy * dirY) / len >= cosTurn;
                    if (!sameDir || Math.Abs(across) > MaxAcrossEm * em || along < -MaxOverlapEm * em || along > MaxGapEm * em)
                    {
                        Flush(run, text, dirX, dirY, result);
                        run.Clear();
                        text.Length = 0;
                    }
                    else if (!space && along > WordGapEm * em && text.Length > 0 && text[text.Length - 1] != ' ')
                        text.Append(' ');
                }
                if (run.Count == 0)
                {
                    if (space || len < 1e-6)
                        continue; // строка начинается с буквы, задающей направление
                    dirX = dx / len;
                    dirY = dy / len;
                }
                run.Add(l);
                text.Append(space ? " " : l.Value);
            }
            Flush(run, text, dirX, dirY, result);
            return result;
        }

        private static void Flush(List<RotatedLetter> run, StringBuilder text, double dirX, double dirY, List<OcrRotatedText> result)
        {
            if (run.Count == 0)
                return;
            string s = CollapseSpaces(text.ToString());
            if (s.Length == 0 || PdfTextExtract.IsPrivateUseOnly(s))
                return;
            RotatedLetter first = run[0];
            RotatedLetter last = first;
            foreach (RotatedLetter l in run)
                if (!string.IsNullOrWhiteSpace(l.Value))
                    last = l;
            double length = (last.EndX - first.StartX) * dirX + (last.EndY - first.StartY) * dirY;
            if (!(length > 0))
                return;
            double size = first.SizePt > 0 ? first.SizePt : FontResolver.DefaultFontSize;
            double upX = -dirY, upY = dirX;
            // Центр рамки: середина ЗАПАСЁННОЙ длины от начала строки и середина высоты рамки.
            var t = new OcrRotatedText
            {
                Text = s, LengthPt = length, FontSizePt = size,
                AngleDeg = Normalize(Math.Atan2(dirY, dirX) * 180 / Math.PI),
                Bold = first.Bold, Italic = first.Italic, ColorArgb = first.ColorArgb, FontName = first.FontName
            };
            double half = FrameWidth(t) / 2, lift = (BaselineFromTopEm - FrameHeightEm / 2) * size;
            t.CenterXPt = first.StartX + dirX * half + upX * lift;
            t.CenterYPt = first.StartY + dirY * half + upY * lift;
            result.Add(t);
        }

        /// <summary>Начало базовой линии надписи — обратное к рамке из <see cref="Group"/>. Чистая — под тест.</summary>
        internal static void BaselineStart(OcrRotatedText t, out double x, out double y)
        {
            double a = t.AngleDeg * Math.PI / 180, dirX = Math.Cos(a), dirY = Math.Sin(a);
            double half = FrameWidth(t) / 2, lift = (BaselineFromTopEm - FrameHeightEm / 2) * t.FontSizePt;
            x = t.CenterXPt - dirX * half + dirY * lift;
            y = t.CenterYPt - dirY * half - dirX * lift;
        }

        private static string CollapseSpaces(string s)
        {
            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (sb.Length > 0 && sb[sb.Length - 1] != ' ')
                        sb.Append(' ');
                }
                else
                    sb.Append(c);
            }
            return sb.ToString().TrimEnd();
        }

        /// <summary>
        /// Пользовательский поворот страницы (по часовой, 0/90/180/270) для надписи: центр — как
        /// точка (<see cref="PageRotation.MapPoint"/>), направление — на тот же угол по часовой.
        /// pageW/pageH — размеры ДО поворота. Чистая — под тест.
        /// </summary>
        public static void Rotate(List<OcrRotatedText> texts, int rotation, double pageW, double pageH)
        {
            if (rotation == 0 || texts == null)
                return;
            foreach (OcrRotatedText t in texts)
            {
                double x, y;
                PageRotation.MapPoint(t.CenterXPt, t.CenterYPt, rotation, pageW, pageH, out x, out y);
                t.CenterXPt = x;
                t.CenterYPt = y;
                t.AngleDeg = Normalize(t.AngleDeg - rotation);
            }
        }

        /// <summary>Формат надписи как ран — писатели оформляют его тем же кодом, что обычный текст.</summary>
        public static OcrRun AsRun(OcrRotatedText t)
        {
            return new OcrRun
            {
                Text = t.Text, FontSizePt = t.FontSizePt, Bold = t.Bold, Italic = t.Italic,
                ColorArgb = t.ColorArgb, FontName = t.FontName
            };
        }
    }
}

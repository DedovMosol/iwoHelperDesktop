using System;

namespace ExcelMerger
{
    /// <summary>Горизонтальная или вертикальная линовка.</summary>
    internal enum LineOrientation { Horizontal, Vertical }

    /// <summary>
    /// Прямая линовка страницы из векторной графики PDF: строго горизонтальный или
    /// вертикальный отрезок в координатах PDF (pt, ось Y направлена вверх). Диагонали и
    /// кривые сюда не попадают — ориентация определяется при извлечении. Основа для
    /// восстановления таблиц (границы ячеек) и подчёркивания текста.
    /// </summary>
    internal sealed class PdfLine
    {
        public LineOrientation Orientation;
        public double X1;
        public double Y1;
        public double X2;
        public double Y2;
        public double Thickness; // толщина штриха (pt); 0 — неизвестна

        public double MinX { get { return Math.Min(X1, X2); } }
        public double MaxX { get { return Math.Max(X1, X2); } }
        public double MinY { get { return Math.Min(Y1, Y2); } }
        public double MaxY { get { return Math.Max(Y1, Y2); } }

        /// <summary>Постоянная координата линии: Y для горизонтали, X для вертикали.</summary>
        public double Position
        {
            get { return Orientation == LineOrientation.Horizontal ? (Y1 + Y2) / 2 : (X1 + X2) / 2; }
        }

        /// <summary>Длина вдоль своей оси: по X для горизонтали, по Y для вертикали.</summary>
        public double Length
        {
            get { return Orientation == LineOrientation.Horizontal ? MaxX - MinX : MaxY - MinY; }
        }
    }
}

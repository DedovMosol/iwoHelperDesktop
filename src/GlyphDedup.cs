using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Глиф для дедупликации: символ, центр рамки и кегль (координаты PDF, pt).</summary>
    internal struct GlyphInfo
    {
        public string Value;
        public double CenterX;
        public double CenterY;
        public double SizePt; // кегль; 0 — неизвестен (порог возьмёт запасной в pt)
    }

    /// <summary>
    /// Дедупликация сдвоенных глифов «псевдо-жирного»: в некоторых PDF (ж/д билеты) текст
    /// отрисован ДВАЖДЫ с микросмещением (~0.3 pt) — имитация полужирного без жирного шрифта.
    /// Извлечение иначе выдаёт «ЭЭллееккттрроонннныыйй»: обе копии каждой буквы попадают в одно
    /// слово. Дубль — тот же символ с почти совпадающим центром; порог — на порядок меньше шага
    /// настоящих соседних одинаковых символов («77» стоят не ближе ~0.35 кегля), поэтому ложных
    /// склеек не бывает. Массовая двойная отрисовка — признак жирного (вызывающий ставит Bold).
    /// Чистая логика — под тест.
    /// </summary>
    internal static class GlyphDedup
    {
        private const double DupTolFactor = 0.12; // |Δ центра| ≤ эта доля кегля по ОБЕИМ осям — дубль
        private const double MinTolPt = 0.45;     // запас при неизвестном/крошечном кегле (реальное смещение ~0.3 pt)

        /// <summary>
        /// Индексы глифов, остающихся после дедупликации (в исходном порядке); dropped — сколько
        /// дублей выброшено. Дубль ищется среди УЖЕ оставленных — тройная отрисовка тоже свернётся.
        /// </summary>
        public static List<int> Keep(IList<GlyphInfo> glyphs, out int dropped)
        {
            var keep = new List<int>(glyphs.Count);
            dropped = 0;
            for (int i = 0; i < glyphs.Count; i++)
            {
                GlyphInfo g = glyphs[i];
                double tol = Tolerance(g.SizePt);
                bool duplicate = false;
                foreach (int k in keep)
                {
                    GlyphInfo other = glyphs[k];
                    if (other.Value != g.Value)
                        continue;
                    double t = tol > Tolerance(other.SizePt) ? tol : Tolerance(other.SizePt);
                    if (Abs(g.CenterX - other.CenterX) <= t && Abs(g.CenterY - other.CenterY) <= t)
                    {
                        duplicate = true;
                        break;
                    }
                }
                if (duplicate)
                    dropped++;
                else
                    keep.Add(i);
            }
            return keep;
        }

        /// <summary>Дубли были массовыми (не меньше половины оставшихся): двойная отрисовка = жирный.</summary>
        public static bool LooksBold(int kept, int dropped)
        {
            return dropped * 2 >= kept && kept > 0;
        }

        private static double Tolerance(double sizePt)
        {
            double t = DupTolFactor * sizePt;
            return t > MinTolPt ? t : MinTolPt;
        }

        private static double Abs(double v)
        {
            return v < 0 ? -v : v;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Запоминание размера и положения окон между запусками (ключ — имя типа формы). Хранится в
    /// <see cref="UserSettings"/> строками «wnd.&lt;Форма&gt;=x,y,w,h,m» (m=1 — развёрнуто).
    /// Восстановление КЛАМПИТ прямоугольник в видимую рабочую область экранов (окно не откроется
    /// за краем экрана или на отключённом мониторе) и не меньше <c>MinimumSize</c>; развёрнутое
    /// окно восстанавливается развёрнутым. Единственный вход — <see cref="Attach"/>, одна строка
    /// в конструкторе окна: восстановление и сохранение он подключает событиями сам, поэтому
    /// окну не нужны перекрытия OnLoad и OnFormClosing и подключение нельзя развести по половинке.
    /// Чистые функции (Format/TryParse/ClampToWorkingArea/BestWorkArea) покрыты тестами, а сама
    /// проводка — живым тестом на настоящих окнах. Только UI-поток.
    /// </summary>
    internal static class WindowPlacement
    {
        /// <summary>Сериализация «x,y,w,h,m» (m=1 — развёрнуто). Чистая — под тест.</summary>
        internal static string Format(Rectangle r, bool maximized)
        {
            CultureInfo ci = CultureInfo.InvariantCulture;
            return string.Join(",", new[]
            {
                r.X.ToString(ci), r.Y.ToString(ci), r.Width.ToString(ci), r.Height.ToString(ci),
                maximized ? "1" : "0"
            });
        }

        /// <summary>Разбор «x,y,w,h,m»; false — неверный формат (размеры обязаны быть положительными). Чистая — под тест.</summary>
        internal static bool TryParse(string s, out Rectangle r, out bool maximized)
        {
            r = Rectangle.Empty;
            maximized = false;
            if (string.IsNullOrEmpty(s))
                return false;
            string[] p = s.Split(',');
            if (p.Length != 5)
                return false;
            CultureInfo ci = CultureInfo.InvariantCulture;
            int x, y, w, h, m;
            if (!int.TryParse(p[0], NumberStyles.Integer, ci, out x) ||
                !int.TryParse(p[1], NumberStyles.Integer, ci, out y) ||
                !int.TryParse(p[2], NumberStyles.Integer, ci, out w) ||
                !int.TryParse(p[3], NumberStyles.Integer, ci, out h) ||
                !int.TryParse(p[4], NumberStyles.Integer, ci, out m))
                return false;
            if (w <= 0 || h <= 0)
                return false;
            r = new Rectangle(x, y, w, h);
            maximized = m != 0;
            return true;
        }

        /// <summary>Рабочая область с наибольшим пересечением с desired; при нулевом пересечении — первая. Чистая — под тест.</summary>
        internal static Rectangle BestWorkArea(Rectangle desired, Rectangle[] workAreas)
        {
            Rectangle best = workAreas[0];
            long bestOverlap = -1;
            foreach (Rectangle a in workAreas)
            {
                Rectangle i = Rectangle.Intersect(desired, a);
                long overlap = (long)i.Width * i.Height;
                if (overlap > bestOverlap)
                {
                    bestOverlap = overlap;
                    best = a;
                }
            }
            return best;
        }

        /// <summary>
        /// Подогнать желаемый прямоугольник под видимую область: не меньше minSize, не больше
        /// экрана, целиком внутри той рабочей области, с которой он пересекается сильнее всего
        /// (при полном промахе — первой). Так восстановленное окно не окажется за краем экрана
        /// или на отключённом мониторе. Чистая — под тест.
        /// </summary>
        internal static Rectangle ClampToWorkingArea(Rectangle desired, Rectangle[] workAreas, Size minSize)
        {
            if (workAreas == null || workAreas.Length == 0)
                return desired; // экранов нет (не должно случаться) — как есть
            Rectangle area = BestWorkArea(desired, workAreas);
            int w = Math.Max(desired.Width, minSize.Width);
            int h = Math.Max(desired.Height, minSize.Height);
            if (w > area.Width) w = area.Width;
            if (h > area.Height) h = area.Height;
            int x = desired.X;
            int y = desired.Y;
            if (x + w > area.Right) x = area.Right - w;
            if (y + h > area.Bottom) y = area.Bottom - h;
            if (x < area.Left) x = area.Left;
            if (y < area.Top) y = area.Top;
            return new Rectangle(x, y, w, h);
        }

        // ---------- работа с формой (UI-поток) ----------

        private static Rectangle[] WorkAreas()
        {
            Screen[] screens = Screen.AllScreens;
            var areas = new Rectangle[screens.Length];
            for (int i = 0; i < screens.Length; i++)
                areas[i] = screens[i].WorkingArea;
            return areas;
        }

        /// <summary>
        /// Подключить запоминание размера и положения к окну: восстановление при показе и
        /// сохранение при закрытии. Одна точка подключения вместо пары перекрытий в каждом
        /// окне, поэтому новому окну достаточно одной строки в конструкторе.
        ///
        /// Восстановление идёт на событии Load, то есть ровно там, где раньше стоял вызов
        /// сразу после base.OnLoad. Сохранение — на событии FormClosing, которое поднимает
        /// base.OnFormClosing: окна с проверкой занятости выходят из своего перекрытия ДО
        /// вызова базы, поэтому при отказе в закрытии событие не поднимается и сохранения
        /// не будет. Проверка e.Cancel добавлена на случай, если закрытие отменит кто-то ещё.
        ///
        /// Только для окон с ИЗМЕНЯЕМЫМ размером: окну фиксированного размера габариты
        /// прошлой сессии могут не подойти после смены масштаба экрана.
        /// </summary>
        internal static void Attach(Form form)
        {
            form.Load += delegate { Restore(form); };
            form.FormClosing += delegate(object sender, FormClosingEventArgs e)
            {
                if (!e.Cancel)
                    Save(form);
            };
        }

        /// <summary>Восстановить размер/положение окна из настроек (клампя в видимую область). Только из Attach.</summary>
        private static void Restore(Form form)
        {
            string value;
            if (!UserSettings.Load().WindowBounds.TryGetValue(form.GetType().Name, out value))
                return;
            Rectangle saved;
            bool maximized;
            if (!TryParse(value, out saved, out maximized))
                return;
            form.StartPosition = FormStartPosition.Manual;
            form.Bounds = ClampToWorkingArea(saved, WorkAreas(), form.MinimumSize);
            if (maximized)
                form.WindowState = FormWindowState.Maximized;
        }

        /// <summary>Сохранить размер/положение окна в настройки. Только из Attach.</summary>
        private static void Save(Form form)
        {
            // Развёрнутое/свёрнутое окно: пишем НОРМАЛЬНЫЕ границы (RestoreBounds) — чтобы было
            // куда восстановиться, а флаг «развёрнуто» сохраняем отдельно.
            bool maximized = form.WindowState == FormWindowState.Maximized;
            Rectangle bounds = form.WindowState == FormWindowState.Normal ? form.Bounds : form.RestoreBounds;
            if (bounds.Width <= 0 || bounds.Height <= 0)
                return; // границ ещё нет (окно не показывалось) — не сохраняем
            UserSettings.SaveWindowBounds(form.GetType().Name, Format(bounds, maximized));
        }
    }
}

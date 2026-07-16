using System;
using System.Runtime.InteropServices;

namespace ExcelMerger
{
    public enum TaskbarProgressState
    {
        None = 0,          // TBPF_NOPROGRESS
        Indeterminate = 1, // TBPF_INDETERMINATE
        Normal = 2,        // TBPF_NORMAL
        Error = 4          // TBPF_ERROR
    }

    /// <summary>
    /// Прогресс на кнопке панели задач Windows (ITaskbarList3, Windows 7+).
    /// Все сбои проглатываются: индикатор — украшение, не причина падать.
    /// Вызывать только из UI-потока (COM-объект панели задач — STA).
    /// </summary>
    internal sealed class TaskbarProgress
    {
        private ITaskbarList3 _taskbar;
        private bool _unavailable;

        public void SetState(IntPtr hwnd, TaskbarProgressState state)
        {
            try
            {
                Ensure();
                if (_taskbar != null)
                    _taskbar.SetProgressState(hwnd, (int)state);
            }
            catch
            {
                _unavailable = true;
            }
        }

        public void SetValue(IntPtr hwnd, int completed, int total)
        {
            if (total <= 0)
                return;
            try
            {
                Ensure();
                if (_taskbar != null)
                    _taskbar.SetProgressValue(hwnd, (ulong)completed, (ulong)total);
            }
            catch
            {
                _unavailable = true;
            }
        }

        private void Ensure()
        {
            if (_taskbar == null && !_unavailable)
            {
                _taskbar = (ITaskbarList3)new TaskbarListRCW();
                _taskbar.HrInit();
            }
        }

        [ComImport, Guid("56FDF344-FD6D-11d0-958A-006097C9A090"), ClassInterface(ClassInterfaceType.None)]
        private class TaskbarListRCW
        {
        }

        [ComImport, Guid("EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface ITaskbarList3
        {
            // ITaskbarList — порядок vtable обязателен
            void HrInit();
            void AddTab(IntPtr hwnd);
            void DeleteTab(IntPtr hwnd);
            void ActivateTab(IntPtr hwnd);
            void SetActiveAlt(IntPtr hwnd);
            // ITaskbarList2
            void MarkFullscreenWindow(IntPtr hwnd, [MarshalAs(UnmanagedType.Bool)] bool fullscreen);
            // ITaskbarList3 — используются два метода; последующие слоты vtable не нужны
            void SetProgressValue(IntPtr hwnd, ulong completed, ulong total);
            void SetProgressState(IntPtr hwnd, int flags);
        }
    }
}

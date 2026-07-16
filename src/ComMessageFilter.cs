using System;
using System.Runtime.InteropServices;

namespace ExcelMerger
{
    /// <summary>
    /// OLE-фильтр сообщений: автоматический повтор COM-вызовов, отклонённых занятым
    /// Excel (SERVERCALL_RETRYLATER — типично, когда антивирус сканирует открываемый
    /// файл или Excel занят внутренней операцией). Без фильтра такой вызов сразу
    /// падает с «Вызов был отклонён» (RPC_E_CALL_REJECTED).
    /// Регистрация действует на текущем STA-потоке; Register/Revoke должны
    /// вызываться парой на том же потоке, что и COM-вызовы.
    /// </summary>
    internal sealed class ComMessageFilter : IOleMessageFilter
    {
        private const int ServerCallIsHandled = 0;   // SERVERCALL_ISHANDLED
        private const int ServerCallRetryLater = 2;  // SERVERCALL_RETRYLATER
        private const int PendingMsgWaitDefProcess = 2;
        private const int RetryDelayMs = 250;
        private const int GiveUpAfterMs = 20000;

        [ThreadStatic] private static IOleMessageFilter _previous;
        [ThreadStatic] private static bool _registered;

        public static void Register()
        {
            if (_registered)
                return;
            IOleMessageFilter old;
            // На MTA-потоке вернёт ошибку — тогда просто работаем без фильтра.
            if (CoRegisterMessageFilter(new ComMessageFilter(), out old) == 0)
            {
                _previous = old;
                _registered = true;
            }
        }

        public static void Revoke()
        {
            if (!_registered)
                return;
            IOleMessageFilter dummy;
            CoRegisterMessageFilter(_previous, out dummy);
            _previous = null;
            _registered = false;
        }

        int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return ServerCallIsHandled;
        }

        int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == ServerCallRetryLater && dwTickCount < GiveUpAfterMs)
                return RetryDelayMs; // повторить вызов через 250 мс
            return -1;               // отказаться — вызов завершится ошибкой
        }

        int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return PendingMsgWaitDefProcess;
        }

        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    }

    [ComImport, Guid("00000016-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }
}

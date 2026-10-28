using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Современный диалог выбора папки (IFileOpenDialog, Vista+)
    /// с фолбэком на стандартный FolderBrowserDialog.
    /// </summary>
    public static class FolderPicker
    {
        /// <summary>Выбранная папка или null (отмена).</summary>
        public static string Show(IWin32Window owner, string title, string initialDir)
        {
            string path;
            if (TryShowVistaDialog(owner, title, initialDir, out path))
                return path;
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.Description = title;
                if (!string.IsNullOrEmpty(initialDir) && Directory.Exists(initialDir))
                    fbd.SelectedPath = initialDir;
                return fbd.ShowDialog(owner) == DialogResult.OK ? fbd.SelectedPath : null;
            }
        }

        private const int HResultCancelled = unchecked((int)0x800704C7); // HRESULT_FROM_WIN32(ERROR_CANCELLED)

        /// <summary>
        /// Современный диалог. false — показать его не удалось (нет COM-класса, сбой до или при
        /// показе): тогда, и только тогда, — запасной диалог. Раньше запасной открывался после
        /// ЛЮБОГО сбоя, в том числе когда человек уже выбрал папку, а путь не прочитался, —
        /// второе окно выбора подряд. Человек ответил — true: путь или null (отмена, или выбран
        /// не файловый путь).
        /// </summary>
        private static bool TryShowVistaDialog(IWin32Window owner, string title, string initialDir, out string path)
        {
            path = null;
            IFileOpenDialog dialog = null;
            IShellItem folder = null;
            IShellItem item = null;
            IntPtr pszPath = IntPtr.Zero;
            try
            {
                int hr;
                try
                {
                    dialog = (IFileOpenDialog)new FileOpenDialogRCW();
                    uint options;
                    dialog.GetOptions(out options);
                    dialog.SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
                    dialog.SetTitle(title);

                    if (!string.IsNullOrEmpty(initialDir) && Directory.Exists(initialDir))
                    {
                        Guid shellItemGuid = typeof(IShellItem).GUID;
                        if (SHCreateItemFromParsingName(initialDir, IntPtr.Zero, ref shellItemGuid, out folder) == 0)
                            dialog.SetFolder(folder);
                    }

                    hr = dialog.Show(owner != null ? owner.Handle : IntPtr.Zero);
                }
                catch
                {
                    return false; // современный диалог недоступен
                }
                if (hr == HResultCancelled)
                    return true; // человек отказался
                if (hr != 0)
                    return false; // показать не удалось

                try
                {
                    dialog.GetResult(out item);
                    item.GetDisplayName(SIGDN_FILESYSPATH, out pszPath);
                    path = Marshal.PtrToStringUni(pszPath);
                }
                catch
                {
                    path = null; // выбор сделан, но пути нет — второе окно не открываем
                }
                return true;
            }
            finally
            {
                if (pszPath != IntPtr.Zero)
                    Marshal.FreeCoTaskMem(pszPath);
                // Диалог — одноразовый COM-объект: освобождаем детерминированно, не дожидаясь GC.
                ComSafe.Release(item);
                ComSafe.Release(folder);
                ComSafe.Release(dialog);
            }
        }

        private const uint FOS_PICKFOLDERS = 0x00000020;
        private const uint FOS_FORCEFILESYSTEM = 0x00000040;
        private const uint SIGDN_FILESYSPATH = 0x80058000;

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHCreateItemFromParsingName(
            string pszPath, IntPtr pbc, ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out IShellItem ppv);

        [ComImport, Guid("dc1c5a9c-e88a-4dde-a5a1-60f82a20aef7")]
        private class FileOpenDialogRCW { }

        [ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileOpenDialog
        {
            [PreserveSig]
            int Show(IntPtr hwndParent);
            void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
            void SetFileTypeIndex(uint iFileType);
            void GetFileTypeIndex(out uint piFileType);
            void Advise(IntPtr pfde, out uint pdwCookie);
            void Unadvise(uint dwCookie);
            void SetOptions(uint fos);
            void GetOptions(out uint pfos);
            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder(out IShellItem ppsi);
            void GetCurrentSelection(out IShellItem ppsi);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName(out IntPtr pszName);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            void GetResult(out IShellItem ppsi);
            void AddPlace(IShellItem psi, int fdap);
            void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            void Close(int hr);
            void SetClientGuid(ref Guid guid);
            void ClearClientData();
            void SetFilter(IntPtr pFilter);
            void GetResults(out IntPtr ppenum);
            void GetSelectedItems(out IntPtr ppsai);
        }

        [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
            void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
            void GetParent(out IShellItem ppsi);
            void GetDisplayName(uint sigdnName, out IntPtr ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        }
    }
}

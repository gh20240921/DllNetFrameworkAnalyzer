using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace DllNetFrameworkAnalyzer
{
    public class FolderBrowserDialog
    {
        public string Description { get; set; }
        public string SelectedPath { get; set; }

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SHBrowseForFolder(BROWSEINFO lpbi);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern bool SHGetPathFromIDList(IntPtr pidl, IntPtr pszPath);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct BROWSEINFO
        {
            public IntPtr hwndOwner;
            public IntPtr pidlRoot;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pszDisplayName;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpszTitle;
            public uint ulFlags;
            public IntPtr lpfn;
            public IntPtr lParam;
            public int iImage;
        }

        // フラグ定義
        private const uint BIF_RETURNONLYFSDIRS = 0x0001;
        private const uint BIF_NEWDIALOGSTYLE = 0x0040;

        public bool ShowDialog()
        {
            IntPtr pidl = IntPtr.Zero;
            bool result = false;

            try
            {
                var bi = new BROWSEINFO();
                bi.hwndOwner = new WindowInteropHelper(Application.Current.MainWindow).Handle;
                bi.pidlRoot = IntPtr.Zero;
                bi.pszDisplayName = new string(' ', 256);
                bi.lpszTitle = Description;
                bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
                bi.lpfn = IntPtr.Zero;
                bi.lParam = IntPtr.Zero;
                bi.iImage = 0;

                pidl = SHBrowseForFolder(bi);
                if (pidl != IntPtr.Zero)
                {
                    IntPtr pszPath = Marshal.AllocHGlobal(260 * Marshal.SystemDefaultCharSize);
                    result = SHGetPathFromIDList(pidl, pszPath);
                    if (result)
                    {
                        SelectedPath = Marshal.PtrToStringAuto(pszPath);
                    }
                    Marshal.FreeHGlobal(pszPath);
                }
            }
            finally
            {
                if (pidl != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(pidl);
                }
            }

            return result;
        }
    }
}

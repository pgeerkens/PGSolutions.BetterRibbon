using System;
using System.IO;
using System.Runtime.InteropServices;

using Microsoft.Vbe.Interop;
using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using Access = Microsoft.Office.Interop.Access;

namespace PGSolutions.ExcelRibbon.VbaSourceExport {
    internal class AccessWrapper : IDisposable {
        public static bool IsAccessSupported => true;

        public static AccessWrapper New() {
            AccessWrapper returnValue = null;
            try {
                Globals.ThisAddIn.Application.DisplayAlerts = false;

                returnValue = new AccessWrapper();
            } finally {
                Globals.ThisAddIn.Application.DisplayAlerts = false;
            }
            return returnValue;
        }

        private AccessWrapper() => AccessApp = new Access.Application();

        public Access.Application AccessApp { get; }

        /// <summary>Returns true exactly when the Project Object Model is trusted.</summary>
        public bool   IsProjectModelTrusted => AccessApp.VBE != null;
        public string CurrentProjectName    => AccessApp.CurrentProject.Name;
        public VBE    VBE                   => AccessApp.VBE;

        public void OpenDbWithuotAutoexec(string path, bool exclusive = false) =>
            Extensions.InvokeWithShiftKey(() => OpenDbAsCurrent(path,exclusive));

        public void OpenDbAsCurrent(string path, bool exclusive = false) =>
            AccessApp.OpenCurrentDatabase(path, exclusive);

        public void CloseCurrentDb() => AccessApp?.CloseCurrentDatabase();

        private string FullPath(string folder, string filename, string extension) =>
            Path.Combine(folder, Path.ChangeExtension(filename, extension));

        #region Standard IDisposable baseclass implementation
        private bool _isDisposed = false;
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing) {
            if (!_isDisposed) {

                // Dispose of managed resources (only!) here
                if (disposing) {
                    if (AccessApp?.CurrentDb() != null) { AccessApp?.CloseCurrentDatabase(); }
                }

                // Dispose of unmanaged resources here

                // Indicate that the instance has been disposed.
                _isDisposed = true;
            }
        }
        #endregion
    }

    internal static partial class Extensions {

        public static void InvokeWithShiftKey(this Action action) =>
            InvokeWithShiftKey<object>(() => {action(); return null;});

        public static T InvokeWithShiftKey<T>(this Func<T> func) {

            const byte VK_LSHIFT                = 0xA0;  // left shift key
            const uint KEYEVENTF_KEYUP          = 0x2;
            try {
                keybd_event(VK_LSHIFT, 0x10, 0, 0);
                return func();
            } finally {
                keybd_event(VK_LSHIFT, 0x10, KEYEVENTF_KEYUP, 0);
            }
        }

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);
    }
}

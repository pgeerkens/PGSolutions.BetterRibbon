////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Access = Microsoft.Office.Interop.Access;

namespace PGSolutions.BetterRibbon.VbaSourceExport {
    internal class AccessWrapper : IDisposable {
        public static bool IsAccessSupported => true;

        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
        public static AccessWrapper New() {
            try {
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                return new AccessWrapper();
            } finally {
                Globals.ThisAddIn.Application.DisplayAlerts = true;
            }
        }

        private AccessWrapper() => AccessApp = new Access.Application();

        public Access.Application AccessApp { get; }

        /// <summary>Returns true exactly when the Project Object Model is trusted.</summary>
        public bool   IsProjectModelTrusted => AccessApp.VBE != null;
        public string CurrentProjectName    => AccessApp.CurrentProject.Name;

        public void OpenDbWithuotAutoexec(string path, bool exclusive = false) =>
            Extensions.InvokeWithShiftKey(() => OpenDbAsCurrent(path,exclusive));

        public void OpenDbAsCurrent(string path, bool exclusive = false) =>
            AccessApp.OpenCurrentDatabase(path, exclusive);

        public void CloseCurrentDb() => AccessApp?.CloseCurrentDatabase();

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
        public static void InvokeWithShiftKey(this Action action) {

            const byte VK_LSHIFT = 0xA0;  // left shift key
            try {
                VK_LSHIFT.KeyDown();
                action();
            } finally {
                VK_LSHIFT.KeyUp();
            }
        }
    }
    internal static class NativeMethods {
        public static void KeyDown(this byte Vk) => Vk.keybd_event(0x10, KEYEVENTF_KEYDOWN, 0);
        public static void KeyUp(this byte Vk)   => Vk.keybd_event(0x10, KEYEVENTF_KEYUP, 0);

        private const uint KEYEVENTF_KEYDOWN = 0x0;
        private const uint KEYEVENTF_KEYUP   = 0x2;

        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "3")]
        [DllImport("user32.dll")]
        #pragma warning disable IDE1006 // Naming Styles - Matches name in external DLL
        private static extern void keybd_event(this byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);
        #pragma warning restore IDE1006 // Naming Styles
    }
}

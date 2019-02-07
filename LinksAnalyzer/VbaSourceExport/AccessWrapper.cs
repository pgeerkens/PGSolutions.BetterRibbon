////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using Access = Microsoft.Office.Interop.Access;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    internal class AccessWrapper : IDisposable {
        internal AccessWrapper() => AccessApp = new Access.Application();

        public static bool IsAccessSupported => true;

        public Access.Application AccessApp { get; }

        /// <summary>Returns true exactly when the Project Object Model is trusted.</summary>
        public bool   IsProjectModelTrusted => AccessApp.VBE != null;
        public string CurrentProjectName    => AccessApp.CurrentProject.Name;

        public void OpenDbWithuotAutoexec(string path, bool exclusive = false)
        => Extensions.InvokeWithShiftKey(() => AccessApp.OpenCurrentDatabase(path,exclusive));

        public void CloseCurrentDb() => AccessApp?.CloseCurrentDatabase();

        #region Standard IDisposable baseclass implementation w/ Finalizer
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
                    AccessApp.Quit();
                }

                // Dispose of unmanaged resources here

                // Indicate that the instance has been disposed.
                _isDisposed = true;
            }
        }
        #endregion
    }
}

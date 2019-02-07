////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

using Access = Microsoft.Office.Interop.Access;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    internal class AccessWrapper : IDisposable {
        public static bool IsAccessSupported => true;

        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
        public static AccessWrapper New(IApplication application) {
            try {
                application.DisplayAlerts = false;
                return new AccessWrapper();
            } finally {
                application.DisplayAlerts = true;
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
}

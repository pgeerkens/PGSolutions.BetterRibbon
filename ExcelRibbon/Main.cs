////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using System.Runtime.InteropServices;
using PGSolutions.RibbonDispatcher.ComClasses;
using System.Collections.Generic;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

namespace PGSolutions.ExcelRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [Serializable]
    [ComVisible(true)]
    [CLSCompliant(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMain))]
    [Guid(Guids.Main)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    public class Main : IMain {
        private IDictionary<string, ActivatableControl<IRibbonCommon>> AdaptorControls =>
                Globals.ThisAddIn.ViewModel.AdaptorControls;

        internal void WorkbookDeactivate(Excel.Workbook wb) =>
            DeactivateActivatableControls();
        internal void WindowDeactivate(Excel.Workbook wb, Excel.Window wn) =>
            DeactivateActivatableControls();

        private void DeactivateActivatableControls() {
            foreach (var c in AdaptorControls) c.Value.Detach();
        }

        public IRibbonFactory RibbonFactory => Globals.ThisAddIn.ViewModel.RibbonFactory;

        public IRibbonButton AttachProxy(string controlId, IRibbonTextLanguageControl strings) =>
            (AdaptorControls.FirstOrDefault(kv => kv.Key==controlId).Value as RibbonButtonAdaptor)?.Attach(strings);

        public void DetachProxy(string controlId) =>
            (AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonButtonAdaptor)?.Detach();

        /// <inheritdoc/>
        public void InvalidateControl(string ControlId) => Globals.ThisAddIn.ViewModel.InvalidateControl(ControlId);
    }
}

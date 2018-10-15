////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using Excel     = Microsoft.Office.Interop.Excel;
using Workbook  = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.Utilities;

using PGSolutions.LinksAnalyzer;
using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.ExcelRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [Serializable]
    [ComVisible(true)]
    [CLSCompliant(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMain))]
    [Guid(Guids.Main)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    public class Main : IMain, ILinksAnalyzer {
        private static IReadOnlyDictionary<string, IActivatable> AdaptorControls =>
                Globals.ThisAddIn.ViewModel.AdaptorControls;
        private static RibbonViewModel ViewModel = Globals.ThisAddIn.ViewModel;

        internal void WorkbookDeactivate(Workbook wb) =>
            DeactivateActivatableControls();
        internal void WindowDeactivate(Workbook wb, Excel.Window wn) =>
            DeactivateActivatableControls();

        private static void DeactivateActivatableControls() {
            foreach (var c in AdaptorControls) c.Value.Detach();
        }

        /// <inheritdoc/>
        public void InvalidateControl(string ControlId) => Globals.ThisAddIn.ViewModel.InvalidateControl(ControlId);

        public void DetachProxy(string controlId) =>
            (AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonButton)?.Detach();

        public void ShowInactive(bool showWhenInactive) {
            foreach (var ctrl in AdaptorControls) {
                ctrl.Value.ShowWhenInactive = showWhenInactive;
                ctrl.Value.Invalidate();
            }
            ViewModel.InvalidateControl(ViewModel.CustomButtonsViewMode.GroupId);
        }

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonControlStrings NewControlStrings(string label,
                string screenTip = "", string superTip = "",
                string keyTip = "", string alternateLabel = "", string description = "") =>
            Globals.ThisAddIn.ViewModel.RibbonFactory.NewControlStrings(label,
                    screenTip, superTip, keyTip, alternateLabel, description);

        public IRibbonButton AttachButton(string controlId, IRibbonControlStrings strings) {
            var ctrl = AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonButton;
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(controlId));
            ctrl?.Attach();
            return ctrl;
        }

        public IRibbonToggleButton AttachToggle(string controlId, IRibbonControlStrings strings,
                IBooleanSource source) {
            var ctrl = AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonToggleButton;
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(controlId));
            ctrl?.Attach(source.Getter);
            return ctrl;
        }

        public IRibbonCheckBox AttachCheckBox(string controlId, IRibbonControlStrings strings,
                IBooleanSource source) {
            var ctrl = AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonCheckBox;
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(controlId));
            ctrl?.Attach(source.Getter);
            return ctrl;
        }

        public IRibbonDropDown AttachDropDown(string controlId, IRibbonControlStrings strings,
                IIntegerSource source) {
            var ctrl = AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonDropDown;
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(controlId));
            ctrl?.Attach(source.Getter);
            return ctrl;
        }

        #region ILinksAnalyzer
        ILinksLexer ILinksAnalyzer.NewLinksLexer(ISourceCellRef cellRef, string formula) =>
             new LinksLexer(cellRef, formula);

        ISourceCellRef ILinksAnalyzer.NewSourceCellRef(Workbook wkbk, string tabName, string cellName) =>
            new SourceCellRef( wkbk, tabName, cellName );

        ISourceCellRef ILinksAnalyzer.NewSourceCellRef2(string wkBkPath, string wkBkName, string tabName, string cellName,
            bool isNamedRange = false
        ) => new SourceCellRef( wkBkPath, wkBkName, tabName, cellName );

        IExternalLinks ILinksAnalyzer.NewExternalLinksWB(Workbook wb, string excludedName) =>
            new ExternalLinks(wb, excludedName);

        IExternalLinks ILinksAnalyzer.NewExternalLinksWS(Worksheet ws) => new ExternalLinks(ws);

        IExternalLinks ILinksAnalyzer.Parse(ISourceCellRef cellRef, string formula) =>
            new ExternalLinks(cellRef, formula);
        #endregion
    }
}

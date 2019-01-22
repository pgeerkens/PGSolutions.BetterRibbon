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

namespace PGSolutions.BetterRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable" )]
    [Serializable, CLSCompliant(false)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonDispatcher))]
    [Guid(RibbonDispatcher.Guids.BettterRibbon)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    public sealed class Main : IRibbonDispatcher, ILinksAnalyzer {
        internal Main() { }

        private static RibbonViewModel ViewModel = Globals.ThisAddIn.ViewModel;
        private static IReadOnlyDictionary<string, IActivatable> AdaptorControls =>
                ViewModel.AdaptorControls;

        internal void WorkbookDeactivate(Workbook wb) => DetachActivatableControls();
        internal void WindowDeactivate(Workbook wb, Excel.Window wn) => DetachActivatableControls();

        private static void DetachActivatableControls() {
            foreach (var c in AdaptorControls) c.Value.Detach();
        }

        #region IRibbonDispatcher methods
        /// <inheritdoc/>
        public void InvalidateControl(string ControlId) => ViewModel.InvalidateControl( ControlId );

        public void DetachProxy(string controlId) =>
            (AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonButton)?.Detach();

        public void ShowInactive(bool showWhenInactive) {
            foreach (var ctrl in AdaptorControls) {
                ctrl.Value.ShowWhenInactive = showWhenInactive;
                ctrl.Value.Invalidate();
            }
            ViewModel.InvalidateControl(ViewModel.CustomButtonsViewModel.GroupId);
        }

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed",
                Justification = "Matches COM usage.")]
        public IRibbonControlStrings NewControlStrings(string label,
                string screenTip = "", string superTip = "",
                string keyTip = "", string alternateLabel = "", string description = "") =>
            ViewModel.RibbonFactory.NewControlStrings(label,
                    screenTip, superTip, keyTip, alternateLabel, description);

        public IRibbonButton AttachButton(string controlId, IRibbonControlStrings strings) {
            var ctrl = AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonButton;
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(controlId));
            ctrl?.Attach();
            return ctrl;
        }

        public IRibbonToggle AttachToggle(string controlId, IRibbonControlStrings strings,
                IBooleanSource source) {
            var ctrl = AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonToggleButton;
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(controlId));
            ctrl?.Attach(source.Getter);
            return ctrl;
        }

        public IRibbonToggle AttachCheckBox(string controlId, IRibbonControlStrings strings,
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
        #endregion

        #region ILinksAnalyzer methods
        /// <inheritdoc/>
        ILinksLexer ILinksAnalyzer.NewLinksLexer(ISourceCellRef cellRef, string formula)
             => new LinksLexer(cellRef, formula);

        /// <inheritdoc/>
        IExternalLinks ILinksAnalyzer.NewExternalLinks(Excel.Application excel, INameList nameList)
            => new ExternalLinks(Globals.ThisAddIn.Application, nameList);

        /// <inheritdoc/>
        IExternalLinks ILinksAnalyzer.NewExternalLinksWB(Workbook wb, string excludedName)
            => new ExternalLinks(wb, excludedName);

        /// <inheritdoc/>
        IExternalLinks ILinksAnalyzer.NewExternalLinksWS(Worksheet ws)
            => new ExternalLinks(ws);

        /// <inheritdoc/>
        IExternalLinks ILinksAnalyzer.Parse(ISourceCellRef cellRef, string formula)
            => new ExternalLinks(cellRef, formula);

        /// <inheritdoc/>
        void ILinksAnalyzer.WriteLinksAnalysisWB(Excel.Workbook wb)
            => wb.WriteLinks();

        /// <inheritdoc/>
        void ILinksAnalyzer.WriteLinksAnalysisFiles(Workbook wb, INameList nameList)
            => wb.WriteLinks(nameList);
        #endregion
    }
}

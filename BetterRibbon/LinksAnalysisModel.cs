////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonUtilities.LinksAnalysis;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.BetterRibbon {
    /// <summary>The TabModel for the Links Aalysis Group on the BetterRibbon.</summary>
    internal sealed class LinksAnalysisModel : AbstractRibbonGroupModel {
        public LinksAnalysisModel(IRibbonViewModel viewModel, IModelFactoryInternal factory, string viewModelName)
        : base(viewModel, viewModelName) {
            AnalyzeCurrentModel  = factory.NewButtonModel("AnalyzeLinksCurrent", AnalyzeCurrentClicked, "EditLinks");
            AnalyzeSelectedModel = factory.NewButtonModel("AnalyzeLinksSelected", AnalyzeSelectedClicked, "EditLinks");

            Invalidate();
        }

        public IButtonModel AnalyzeCurrentModel  { get; }

        public IButtonModel AnalyzeSelectedModel { get; }

        private void AnalyzeCurrentClicked(object sender)
        => DisplayAnalysis(new WorkbookParser(Application.ActiveWorkbook));

        private void AnalyzeSelectedClicked(object sender)
        => DisplayAnalysis(new WorkbookListParser(Application.Selection));

        private void DisplayAnalysis(IParser parser) {
            Application.Cursor = XlMousePointer.xlWait;
            try {
                parser.StatusAvailable += StatusAvailable;
                Application.ActiveWorkbook.WriteLinks(parser.Parse());
                parser.StatusAvailable -= StatusAvailable;
            }
            finally {
                Application.Cursor = XlMousePointer.xlDefault;
            }
        }

        private void StatusAvailable(object sender, RibbonUtilities.EventArgs<string> e)
        => Application.StatusBar = e.Value;

        private static Application Application => Globals.ThisAddIn.Application;
    }
}

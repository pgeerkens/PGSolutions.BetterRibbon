////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonUtilities.LinksAnalysis;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.BetterRibbon {
    internal sealed class LinksAnalysisModel : AbstractRibbonGroupModel {
        public LinksAnalysisModel(RibbonGroupViewModel viewModel) : base(viewModel) {
            AnalyzeCurrentModel  = GetModel<RibbonButton>("AnalyzeLinksCurrent", AnalyzeCurrentClicked,true, true, "EditLinks");
            AnalyzeSelectedModel = GetModel<RibbonButton>("AnalyzeLinksSelected",AnalyzeSelectedClicked,true, true, "EditLinks");

            Invalidate();
        }

        public RibbonButtonModel AnalyzeCurrentModel  { get; }
        public RibbonButtonModel AnalyzeSelectedModel { get; }

        private void AnalyzeCurrentClicked(object sender) {
            var parser = new LinksParser(Application.ActiveWorkbook);
            parser.StatusAvailable += StatusAvailable;
            DisplayAnalysis(parser);
            parser.StatusAvailable -= StatusAvailable;
        }

        private void AnalyzeSelectedClicked(object sender) {
            var parser = new LinksParser(Application.Selection);
            parser.StatusAvailable += StatusAvailable;
            DisplayAnalysis(parser);
            parser.StatusAvailable -= StatusAvailable;
        }

        private void StatusAvailable(object sender, EventArgs<string> e)
        => Application.StatusBar = e.Value;

        /// <inheritdoc/>
        private static void DisplayAnalysis(ILinksAnalysis externalLinks)
        => Application.ActiveWorkbook.WriteLinks(externalLinks);

        private static Application Application => Globals.ThisAddIn.Application;
    }
}

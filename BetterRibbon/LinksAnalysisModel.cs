////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonUtilities;
using PGSolutions.RibbonUtilities.LinksAnalysis;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.BetterRibbon {
    using IRibbonControlStrings = RibbonDispatcher.ComInterfaces.IRibbonControlStrings;

    internal sealed class LinksAnalysisModel : AbstractRibbonGroupModel {
        public LinksAnalysisModel(RibbonGroupViewModel viewModel) : this(viewModel,null) { }
        public LinksAnalysisModel(RibbonGroupViewModel viewModel, IRibbonControlStrings strings)
        : base(viewModel,strings) {
            AnalyzeCurrentModel  = NewButtonModel("AnalyzeLinksCurrent", AnalyzeCurrentClicked,true, true, "EditLinks");
            AnalyzeSelectedModel = NewButtonModel("AnalyzeLinksSelected",AnalyzeSelectedClicked,true, true, "EditLinks");

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

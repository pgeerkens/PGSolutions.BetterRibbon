////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher;
using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;
using PGSolutions.RibbonUtilities.LinksAnalysis;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.ToolsRibbon {
    /// <summary>The TabModel for the Links Aalysis Group on the ToolsRibbon.</summary>
    internal sealed class LinksAnalysisModel : AbstractRibbonGroupModel {
        public LinksAnalysisModel(IModelFactory factory, IGroupVM viewModel)
        : base(viewModel, factory.GetStrings(viewModel.ControlId)) {
            AnalyzeCurrentModel  = factory.NewButtonModel("AnalyzeLinksCurrent", AnalyzeCurrentClicked, "EditLinks".ToImageObject());
            AnalyzeSelectedModel = factory.NewButtonModel("AnalyzeLinksSelected", AnalyzeSelectedClicked, "EditLinks".ToImageObject());

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

        /// <summary>Posts status reports on the application StatusBar.</summary>
        /// <param name="sender">Originating object for the status post.</param>
        /// <param name="e">An <see cref="RibbonUtilities.EventArgs{T}"/> containing the status to be posted.</param>
        private void StatusAvailable(object sender, RibbonUtilities.EventArgs<string> e)
        => Application.StatusBar = e.Value;

        private static Application Application => Globals.ThisAddIn.Application;
    }
}

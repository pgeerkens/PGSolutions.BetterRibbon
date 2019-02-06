////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonUtilities.LinksAnalyzer;

namespace PGSolutions.BetterRibbon {
    internal sealed class LinksAnalysisModel {
        public LinksAnalysisModel(LinksAnalysisViewModel viewModel) {
            ViewModel = viewModel;
            ViewModel.AnalyzeCurrentClicked  += OnAnalyzeCurrentClicked;
            ViewModel.AnalyzeSelectedClicked += OnAnalyzeSelectedClicked;
        }

        public void Invalidate() => ViewModel.Invalidate();

        private LinksAnalysisViewModel ViewModel { get; set; }

        private void OnAnalyzeCurrentClicked(object sender) 
        => new LinksAnalyzer().WriteLinksAnalysisWB(Application.ActiveWorkbook);

        private void OnAnalyzeSelectedClicked(object sender)
        => Application.ActiveWorkbook.WriteLinks((Application.Selection as Range).GetNameList());

        static Application Application => Globals.ThisAddIn.Application;
    }
}

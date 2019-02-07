////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities.LinksAnalyzer {
    [CLSCompliant(false)]
    public sealed class LinksAnalysisModel {
        public LinksAnalysisModel(Application application, ILinksAnalysisViewModel viewModel) {
            Application = application;
            ViewModel   = viewModel;
            ViewModel.AnalyzeCurrentClicked  += OnAnalyzeCurrentClicked;
            ViewModel.AnalyzeSelectedClicked += OnAnalyzeSelectedClicked;
        }

        public void Invalidate() => ViewModel.Invalidate();

        private Application Application { get; }

        private ILinksAnalysisViewModel ViewModel { get; set; }

        private void OnAnalyzeCurrentClicked(object sender) 
        => new LinksAnalyzer(Application).WriteLinksAnalysisWB(Application.ActiveWorkbook);

        private void OnAnalyzeSelectedClicked(object sender)
        => Application.ActiveWorkbook.WriteLinks((Application.Selection as Range).GetNameList());
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    [CLSCompliant(false)]
    public sealed class LinksAnalysisModel {
        public LinksAnalysisModel(ILinksAnalysisViewModel viewModel) {
            ViewModel   = viewModel;
            ViewModel.AnalyzeCurrentClicked  += OnAnalyzeCurrentClicked;
            ViewModel.AnalyzeSelectedClicked += OnAnalyzeSelectedClicked;
        }

        public void Invalidate() => ViewModel.Invalidate();

        private ILinksAnalysisViewModel ViewModel { get; set; }

        private void OnAnalyzeCurrentClicked(object sender, WorkbookEventArgs e)
        => ViewModel.DisplayAnalysis(new ExternalLinks(e.Workbook, ""));

        private void OnAnalyzeSelectedClicked(object sender, RangeEventArgs e)
        => ViewModel.DisplayAnalysis(new ExternalLinks(ViewModel, e.Range));
    }

    [CLSCompliant(false)]
    public class RangeEventArgs : EventArgs {
        public RangeEventArgs(Range range) => Range = range;
        public Range Range { get; }
    }

    [CLSCompliant(false)]
    public class WorkbookEventArgs : EventArgs {
        public WorkbookEventArgs(Workbook workbook) => Workbook = workbook;
        public Workbook Workbook{ get; }
    }
}

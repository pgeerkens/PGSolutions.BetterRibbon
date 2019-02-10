////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    [CLSCompliant(false)]
    public interface ILinksAnalysisViewModel {
        event EventHandler<WorkbookEventArgs> AnalyzeCurrentClicked;
        event EventHandler<RangeEventArgs>    AnalyzeSelectedClicked;

        dynamic StatusBar { set; }

        void DisplayAnalysis(IExternalLinks externalLinks);

        void Invalidate();
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    public interface ILinksAnalysisViewModel {
        event EventHandler  AnalyzeCurrentClicked;
        event EventHandler  AnalyzeSelectedClicked;

        void Invalidate();
    }
}

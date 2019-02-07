////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalyzer {
    public interface ILinksAnalysisViewModel {
        event ClickedEventHandler  AnalyzeCurrentClicked;
        event ClickedEventHandler  AnalyzeSelectedClicked;

        void Invalidate();
    }
}

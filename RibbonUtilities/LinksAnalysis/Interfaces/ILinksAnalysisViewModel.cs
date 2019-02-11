////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    [CLSCompliant(false)]
    public interface ILinksAnalysisViewModel {
        event EventHandler<WorkbookEventArgs> AnalyzeCurrentClicked;
        event EventHandler<RangeEventArgs>    AnalyzeSelectedClicked;

        [SuppressMessage("Microsoft.Design", "CA1044:PropertiesShouldNotBeWriteOnly", Justification ="Match interface of Excel.")]
        dynamic StatusBar { set; }

        void DisplayAnalysis(ILinksAnalysis externalLinks);

        void Invalidate();
    }
}

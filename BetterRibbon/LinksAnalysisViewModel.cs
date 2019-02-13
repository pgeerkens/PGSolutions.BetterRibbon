////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;

using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonUtilities.LinksAnalysis;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.BetterRibbon {
    /// <summary>.</summary>
    [Description("")]
    [CLSCompliant(false)]
    public class LinksAnalysisViewModel : AbstractLinksAnalysisViewModel, ILinksAnalysisViewModel {
        /// <summary>.</summary>
        public LinksAnalysisViewModel(IRibbonFactory factory, bool isVisible = true, bool isEnabled = true)
        : base(factory, "LinksAnalysisGroup", isVisible, isEnabled) {
        }

        /// <inheritdoc/>
        public override void DisplayAnalysis(ILinksAnalysis externalLinks)
        => Globals.ThisAddIn.Application.ActiveWorkbook.WriteLinks(externalLinks);

        protected override Application Application => Globals.ThisAddIn.Application;
    }
}

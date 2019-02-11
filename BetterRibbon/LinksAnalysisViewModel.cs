////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;

using PGSolutions.BetterRibbon;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>.</summary>
    [Description("")]
    [CLSCompliant(false)]
    public class LinksAnalysisViewModel : AbstractRibbonGroupViewModel, ILinksAnalysisViewModel {
        /// <summary>.</summary>
        public LinksAnalysisViewModel(IRibbonFactory factory) : base(factory) {
            LinksAnalysisGroup    = Factory.NewRibbonGroup("LinksAnalysisGroup", true);
            AnalyzeCurrentButton  = Factory.NewRibbonButtonMso(itemId:"AnalyzeLinksCurrent",  imageMso:"EditLinks");
            AnalyzeSelectedButton = Factory.NewRibbonButtonMso(itemId:"AnalyzeLinksSelected", imageMso:"EditLinks");

            AnalyzeCurrentButton.Clicked += OnAnalyzeCurrentClicked;
            AnalyzeCurrentButton.Attach();

            AnalyzeSelectedButton.Clicked += OnAnalyzeSelectedClicked;
            AnalyzeSelectedButton.Attach();
        }

        /// <inheritdoc/>
        public event EventHandler<WorkbookEventArgs> AnalyzeCurrentClicked;
        /// <inheritdoc/>
        public event EventHandler<RangeEventArgs>    AnalyzeSelectedClicked;
        /// <inheritdoc/>
        public dynamic StatusBar { set => Application.StatusBar = value; }

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        public RibbonGroup  LinksAnalysisGroup    { get; }
        /// <inheritdoc/>
        public RibbonButton AnalyzeCurrentButton  { get; }
        /// <inheritdoc/>
        public RibbonButton AnalyzeSelectedButton { get; }

        /// <inheritdoc/>
        public void DisplayAnalysis(ILinksAnalysis externalLinks)
        => Globals.ThisAddIn.Application.ActiveWorkbook.WriteLinks(externalLinks);

        /// <inheritdoc/>
        public void Invalidate() => LinksAnalysisGroup.Invalidate();

        private void OnAnalyzeCurrentClicked(object sender)
        => AnalyzeCurrentClicked?.Invoke(sender, new WorkbookEventArgs(Application.ActiveWorkbook));
        private void OnAnalyzeSelectedClicked(object sender)
        => AnalyzeSelectedClicked?.Invoke(sender, new RangeEventArgs(Application.Selection));

        static Excel.Application Application => Globals.ThisAddIn.Application;
    }
}

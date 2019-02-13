////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    /// <summary>.</summary>
    [CLSCompliant(false)]
    public abstract class AbstractLinksAnalysisViewModel : AbstractRibbonGroupViewModel, ILinksAnalysisViewModel {
        /// <summary>.</summary>
        protected AbstractLinksAnalysisViewModel(IRibbonFactory factory) : base(factory) {
            LinksAnalysisGroup    = Factory.NewRibbonGroup("LinksAnalysisGroup", true);
            AnalyzeCurrentButton  = Factory.NewRibbonButtonMso(itemId: "AnalyzeLinksCurrent", imageMso: "EditLinks");
            AnalyzeSelectedButton = Factory.NewRibbonButtonMso(itemId: "AnalyzeLinksSelected", imageMso: "EditLinks");

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
        public void Invalidate() => LinksAnalysisGroup.Invalidate();

        public abstract void DisplayAnalysis(ILinksAnalysis externalLinks);

        protected virtual void OnAnalyzeCurrentClicked(object sender)
        => AnalyzeCurrentClicked?.Invoke(sender, new WorkbookEventArgs(Application.ActiveWorkbook));

        protected virtual void OnAnalyzeSelectedClicked(object sender)
        => AnalyzeSelectedClicked?.Invoke(sender, new RangeEventArgs(Application.Selection));

        protected abstract Application Application { get;}
    }
}

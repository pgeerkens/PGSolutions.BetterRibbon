////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComClasses;
using System;
using System.ComponentModel;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
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
        public event EventHandler  AnalyzeCurrentClicked;
        /// <inheritdoc/>
        public event EventHandler  AnalyzeSelectedClicked;

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        public RibbonGroup  LinksAnalysisGroup    { get; }
        /// <inheritdoc/>
        public RibbonButton AnalyzeCurrentButton  { get; }
        /// <inheritdoc/>
        public RibbonButton AnalyzeSelectedButton { get; }

        /// <inheritdoc/>
        public void Invalidate() => LinksAnalysisGroup.Invalidate();

        private void OnAnalyzeCurrentClicked(object sender) => AnalyzeCurrentClicked?.Invoke(sender, EventArgs.Empty);
        private void OnAnalyzeSelectedClicked(object sender) => AnalyzeSelectedClicked?.Invoke(sender, EventArgs.Empty);
    }
}

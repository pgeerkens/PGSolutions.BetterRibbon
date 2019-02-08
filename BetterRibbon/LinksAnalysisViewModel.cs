////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComClasses;
using System;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    [CLSCompliant(false)]
    public class LinksAnalysisViewModel : AbstractRibbonGroupViewModel, ILinksAnalysisViewModel {
        public LinksAnalysisViewModel(IRibbonFactory factory) : base(factory) {
            LinksAnalysisGroup    = Factory.NewRibbonGroup("LinksAnalysisGroup", true);
            AnalyzeCurrentButton  = Factory.NewRibbonButtonMso(itemId:"AnalyzeLinksCurrent",  imageMso:"EditLinks");
            AnalyzeSelectedButton = Factory.NewRibbonButtonMso(itemId:"AnalyzeLinksSelected", imageMso:"EditLinks");

            AnalyzeCurrentButton.Clicked += OnAnalyzeCurrentClicked;
            AnalyzeCurrentButton.Attach();

            AnalyzeSelectedButton.Clicked += OnAnalyzeSelectedClicked;
            AnalyzeSelectedButton.Attach();
        }

        public event EventHandler  AnalyzeCurrentClicked;
        public event EventHandler  AnalyzeSelectedClicked;

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        public RibbonGroup  LinksAnalysisGroup    { get; }
        public RibbonButton AnalyzeCurrentButton  { get; }
        public RibbonButton AnalyzeSelectedButton { get; }

        public void Invalidate() => LinksAnalysisGroup.Invalidate();

        private void OnAnalyzeCurrentClicked(object sender) => AnalyzeCurrentClicked?.Invoke(sender, EventArgs.Empty);
        private void OnAnalyzeSelectedClicked(object sender) => AnalyzeSelectedClicked?.Invoke(sender, EventArgs.Empty);
    }
}

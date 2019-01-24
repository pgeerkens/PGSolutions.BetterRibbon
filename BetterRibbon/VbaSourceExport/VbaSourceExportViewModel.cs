////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;
using static Microsoft.Office.Core.RibbonControlSize;

namespace PGSolutions.BetterRibbon.VbaSourceExport {
    internal class VbaSourceExportViewModel : AbstractRibbonGroupViewModel, IVbaSourceExportViewModel {
        public VbaSourceExportViewModel(IRibbonFactory factory, string suffix) : base(factory) {
            var defaultSize = suffix=="MS" ? RibbonControlSizeRegular : RibbonControlSizeLarge;
            VbASourceExportGroup  = Factory.NewRibbonGroup($"VbaExportGroup{suffix}");

            UseSrcFolderToggle    = Factory.NewRibbonToggleMso($"UseSrcFolderToggle{suffix}",
                                            size:defaultSize, imageMso:"MacroSecurity");
            SelectedProjectButton = Factory.NewRibbonButtonMso($"SelectedProjectButton{suffix}",
                                            size:defaultSize, imageMso:"RefreshAll", showImage:true);
            CurrentProjectButton = Factory.NewRibbonButtonMso($"CurrentProjectButton{suffix}",
                                            size:defaultSize, imageMso:"Refresh", showImage:true);

            UseSrcFolderToggle.Toggled    += OnSrcFolderToggled;
            SelectedProjectButton.Attach<RibbonButton>().Clicked += OnExportSelected;
            CurrentProjectButton.Attach<RibbonButton>().Clicked  += OnExportCurrent;
        }

        public void Attach(IBooleanSource srcToggleSource) =>
            UseSrcFolderToggle.Attach(srcToggleSource.Getter);

        public void Invalidate() => UseSrcFolderToggle.Invalidate();

        public event ToggledEventHandler UseSrcFolderToggled;
        public event ClickedEventHandler SelectedProjectsClicked;
        public event ClickedEventHandler CurrentProjectClicked;

        private void OnSrcFolderToggled(object sender, bool isPressed) =>
            UseSrcFolderToggled?.Invoke(sender,isPressed);
        private void OnExportSelected(object sender) =>
            SelectedProjectsClicked?.Invoke(sender);
        private void OnExportCurrent(object sender) =>
            CurrentProjectClicked?.Invoke(sender);

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private  RibbonGroup        VbASourceExportGroup  { get; }
        private  RibbonToggleButton UseSrcFolderToggle    { get; }
        public   RibbonButton       SelectedProjectButton { get; }
        public   RibbonButton       CurrentProjectButton  { get; }

        IRibbonToggle IVbaSourceExportViewModel.UseSrcFolderToggle    => UseSrcFolderToggle;
        IRibbonButton IVbaSourceExportViewModel.SelectedProjectButton => SelectedProjectButton;
        IRibbonButton IVbaSourceExportViewModel.CurrentProjectButton  => CurrentProjectButton;
    }
}

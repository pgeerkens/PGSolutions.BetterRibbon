////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using static Microsoft.Office.Core.RibbonControlSize;

namespace PGSolutions.BetterRibbon.VbaSourceExport {
    internal class VbaSourceExportViewModel : AbstractRibbonGroupViewModel, IVbaSourceExportGroupModel {
        public VbaSourceExportViewModel(IRibbonFactory factory, string suffix) : base(factory) {
            var defaultSize = suffix=="MS" ? RibbonControlSizeRegular : RibbonControlSizeLarge;
            VbASourceExportGroup  = Factory.NewRibbonGroup($"VbaExportGroup{suffix}");

            UseSrcFolderToggle    = Factory.NewRibbonToggleMso($"UseSrcFolderToggle{suffix}",
                                            size:defaultSize, imageMso:"MacroSecurity");
            SelectedProjectButton = Factory.NewRibbonButtonMso($"SelectedProjectButton{suffix}",
                                            size:defaultSize, imageMso:"RefreshAll", showImage:true);
            CurrentProjectButton = Factory.NewRibbonButtonMso($"CurrentProjectButton{suffix}",
                                            size:defaultSize, imageMso:"Refresh", showImage:true);
        }

        public void Attach(Func<bool> srcToggleSource) {
            UseSrcFolderToggle.Attach(srcToggleSource); UseSrcFolderToggle.Toggled    += OnToggled;
            SelectedProjectButton.Attach();             SelectedProjectButton.Clicked += OnExportSelected;
            CurrentProjectButton.Attach();              CurrentProjectButton.Clicked  += OnExportCurrent;
        }

        public void Detach() {
            UseSrcFolderToggle.Detach();       UseSrcFolderToggle.Toggled    -= OnToggled;
            SelectedProjectButton.Detach();    SelectedProjectButton.Clicked -= OnExportSelected;
            CurrentProjectButton.Detach();     CurrentProjectButton.Clicked  -= OnExportCurrent;
        }

        public void Invalidate() =>UseSrcFolderToggle.Invalidate();

        public event ToggledEventHandler UseSrcFolderToggled;
        public event ClickedEventHandler SelectedProjectsClicked;
        public event ClickedEventHandler CurrentProjectClicked;

        private void OnToggled(object sender, bool isPressed) => UseSrcFolderToggled(sender,isPressed);
        private void OnExportSelected(object sender) => SelectedProjectsClicked(sender);
        private void OnExportCurrent(object sender) => CurrentProjectClicked(sender);

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private RibbonGroup        VbASourceExportGroup  { get; }
        private RibbonToggleButton UseSrcFolderToggle    { get; }
        private RibbonButton       SelectedProjectButton { get; }
        private RibbonButton       CurrentProjectButton  { get; }
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ControlMixins;
using static PGSolutions.RibbonDispatcher.ComInterfaces.RdControlSize;

namespace PGSolutions.ExcelRibbon.VbaSourceExport {
    internal class VbaSourceExportViewModel : AbstractRibbonGroupViewModel, IVbaSourceExportGroupModel {
        public VbaSourceExportViewModel(IRibbonFactory factory, string suffix) : base(factory) {
            var defaultSize = suffix=="MS" ? rdRegular : rdLarge;
            VbASourceExportGroup  = Factory.NewRibbonGroup($"VbaExportGroup{suffix}");

            UseSrcFolderToggle    = Factory.NewRibbonToggleMso($"UseSrcFolderToggle{suffix}",
                    Size: defaultSize, ImageMso:"MacroSecurity");
            UseSrcFolderToggle.Toggled += UseSrcFolderToggled;

            SelectedProjectButton = Factory.NewRibbonButtonMso($"SelectedProjectButton{suffix}",
                    Size: defaultSize, ImageMso:"RefreshAll", ShowImage:true);
            SelectedProjectButton.Clicked += SelectedProjectsClicked;

            CurrentProjectButton = Factory.NewRibbonButtonMso($"CurrentProjectButton{suffix}",
                    Size: defaultSize, ImageMso:"Refresh", ShowImage:true);
            CurrentProjectButton.Clicked  += CurrentProjectClicked;
        }

        public void Detach() {
            //UseSrcFolderToggle.IsPressed = srcToggleSource(); // .Toggled -= UseSrcFolderToggled;
            SelectedProjectButton.Detach(); // .Clicked -= SelectedProjectsClicked; ;
            CurrentProjectButton.Detach(); // .Clicked  -= CurrentProjectClicked;
        }

        public void Attach(Func<bool> srcToggleSource) {
            UseSrcFolderToggle.IsPressed = srcToggleSource(); // .Toggled += UseSrcFolderToggled;
            SelectedProjectButton.Attach(null); // .Clicked += SelectedProjectsClicked; ;
            CurrentProjectButton.Attach(null); // .Clicked  += CurrentProjectClicked;
        }

        public event ToggledEventHandler UseSrcFolderToggled;
        public event ClickedEventHandler SelectedProjectsClicked;
        public event ClickedEventHandler CurrentProjectClicked;

        private RibbonGroup        VbASourceExportGroup  { get; }
        private RibbonToggleButton UseSrcFolderToggle    { get; }
        private RibbonButton       SelectedProjectButton { get; }
        private RibbonButton       CurrentProjectButton  { get; }
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ControlMixins;
using static PGSolutions.RibbonDispatcher.ComInterfaces.RdControlSize;

namespace PGSolutions.ExcelRibbon.VbaSourceExport {
    internal class VbaSourceExportViewModel : AbstractRibbonGroupViewModel {
        public VbaSourceExportViewModel(IRibbonFactory factory, string suffix, Func<bool> toggleSource) : base(factory) {
            var defaultSize = suffix=="MS" ? rdRegular : rdLarge;
            VbASourceExportGroup  = Factory.NewRibbonGroup($"VbaExportGroup{suffix}");
            UseSrcFolderToggle    = Factory.NewRibbonToggleMso($"UseSrcFolderToggle{suffix}", Size: defaultSize, ImageMso:"MacroSecurity");
            SelectedProjectButton = Factory.NewRibbonButtonMso($"SelectedProjectButton{suffix}", Size: defaultSize, ImageMso:"RefreshAll", ShowImage:true);
            CurrentProjectButton  = Factory.NewRibbonButtonMso($"CurrentProjectButton{suffix}", Size: defaultSize, ImageMso:"Refresh", ShowImage:true);

            UseSrcFolderToggle.Toggled    += UseSrcFolderToggled;
            UseSrcFolderToggle.IsPressed   = toggleSource();
            SelectedProjectButton.Clicked += SelectedProjectsClicked;
            CurrentProjectButton.Clicked  += CurrentProjectClicked;
        }

        public event ToggledEventHandler UseSrcFolderToggled;
        public event ClickedEventHandler SelectedProjectsClicked;
        public event ClickedEventHandler CurrentProjectClicked;

        public RibbonGroup        VbASourceExportGroup  { get; }
        public RibbonToggleButton UseSrcFolderToggle    { get; }
        public RibbonButton       SelectedProjectButton { get; }
        public RibbonButton       CurrentProjectButton  { get; }
        public RibbonDropDown     ButtonOptions         { get; }
    }
}

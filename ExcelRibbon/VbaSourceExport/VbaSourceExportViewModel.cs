////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using static PGSolutions.RibbonDispatcher.ComInterfaces.RdControlSize;

namespace PGSolutions.ExcelRibbon.VbaSourceExport {
    internal class VbaSourceExportViewModel : AbstractRibbonGroupViewModel {
        public VbaSourceExportViewModel(IRibbonFactory factory, string suffix) : base(factory) {
            var defaultSize = suffix=="MS" ? rdRegular : rdLarge;
            VbASourceExportGroup  = Factory.NewRibbonGroup($"VbaExportGroup{suffix}");
            UseSrcFolderToggle    = Factory.NewRibbonToggleMso($"UseSrcFolderToggle{suffix}", Size: defaultSize, ImageMso:"MacroSecurity");
            SelectedProjectButton = Factory.NewRibbonButtonMso($"SelectedProjectButton{suffix}", Size: defaultSize, ImageMso:"RefreshAll", ShowImage:true);
            CurrentProjecctButton = Factory.NewRibbonButtonMso($"CurrentProjectButton{suffix}", Size: defaultSize, ImageMso:"Refresh", ShowImage:true);

            SelectedProjectButton.Clicked += () => VbaSourceExportModel.ExportSelectedProject(UseSrcFolderToggle.IsPressed);
            CurrentProjecctButton.Clicked += () => VbaSourceExportModel.ExportCurrentProject(UseSrcFolderToggle.IsPressed);
            UseSrcFolderToggle.IsPressed = true;
        }

        public RibbonGroup        VbASourceExportGroup  { get; }
        public RibbonButton       SelectedProjectButton { get; }
        public RibbonButton       CurrentProjecctButton { get; }
        public RibbonDropDown     ButtonOptions         { get; }
        public RibbonToggleButton UseSrcFolderToggle    { get; }
    }
}

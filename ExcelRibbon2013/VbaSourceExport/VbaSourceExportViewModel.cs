////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.AbstractCOM;
using PGSolutions.RibbonDispatcher.ConcreteCOM;
using static PGSolutions.RibbonDispatcher.AbstractCOM.RdControlSize;

namespace PGSolutions.ExcelRibbon.VbaSourceExport {
    internal class VbaSourceExportViewModel : AbstractRibbonGroupViewModel {
        public VbaSourceExportViewModel(IRibbonFactory factory) : base(factory) {
            VbASourceExportGroup  = Factory.NewRibbonGroup("VbASourceExportGroup");
            UseSrcFolderToggle    = Factory.NewRibbonToggleMso("UseSrcFolderToggle", Size:rdLarge, ImageMso:"MacroSecurity");
            SelectedProjectButton = Factory.NewRibbonButtonMso("SelectedProjectButton", Size:rdLarge, ImageMso:"RefreshAll", ShowImage:true);
            CurrentProjecctButton = Factory.NewRibbonButtonMso("CurrentProjecctButton", Size:rdLarge, ImageMso:"Refresh", ShowImage:true);

            SelectedProjectButton.Clicked += () => VbaSourceExportModel.ExportSelectedProject(UseSrcFolderToggle.IsPressed);
            CurrentProjecctButton.Clicked += () => VbaSourceExportModel.ExportCurrentProject(UseSrcFolderToggle.IsPressed);
        }

        public RibbonGroup        VbASourceExportGroup  { get; }
        public RibbonButton       SelectedProjectButton { get; }
        public RibbonButton       CurrentProjecctButton { get; }
        public RibbonDropDown     ButtonOptions         { get; }
        public RibbonToggleButton UseSrcFolderToggle    { get; }
    }
}

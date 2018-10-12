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
                                        Size:defaultSize, ImageMso:"MacroSecurity");
            SelectedProjectButton = Factory.NewRibbonButtonMso($"SelectedProjectButton{suffix}",
                                        Size:defaultSize, ImageMso:"RefreshAll", ShowImage:true);
            CurrentProjectButton = Factory.NewRibbonButtonMso($"CurrentProjectButton{suffix}",
                                        Size:defaultSize, ImageMso:"Refresh", ShowImage:true);
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

        public event ToggledEventHandler UseSrcFolderToggled;
        public event ClickedEventHandler SelectedProjectsClicked;
        public event ClickedEventHandler CurrentProjectClicked;

        private void OnToggled(bool isPressed) => UseSrcFolderToggled(isPressed);
        private void OnExportSelected() => SelectedProjectsClicked();
        private void OnExportCurrent() => CurrentProjectClicked();

        private RibbonGroup        VbASourceExportGroup  { get; }
        private RibbonToggleButton UseSrcFolderToggle    { get; }
        private RibbonButton       SelectedProjectButton { get; }
        private RibbonButton       CurrentProjectButton  { get; }
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon.VbaSourceExport {
    internal interface IVbaSourceExportViewModel {
        event ToggledEventHandler UseSrcFolderToggled;
        event ClickedEventHandler SelectedProjectsClicked;
        event ClickedEventHandler CurrentProjectClicked;

        IRibbonToggle UseSrcFolderToggle  { get; }
        IRibbonButton SelectedProjectButton { get; }
        IRibbonButton CurrentProjectButton  { get; }

        void Attach(IBooleanSource srcToggleSource);
        void Invalidate();
    }
}

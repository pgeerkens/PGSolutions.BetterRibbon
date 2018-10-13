////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using PGSolutions.RibbonDispatcher.ControlMixins;

namespace PGSolutions.ExcelRibbon.VbaSourceExport {
    internal interface IVbaSourceExportGroupModel {
        event ToggledEventHandler UseSrcFolderToggled;
        event ClickedEventHandler SelectedProjectsClicked;
        event ClickedEventHandler CurrentProjectClicked;

        void Attach(Func<bool> useSrcFolderSource);
        void Detach();

        void Invalidate();
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using VbaExportEventHandler = EventHandler<VbaExportEventArgs>;

    [CLSCompliant(false)]
    public interface IVbaSourceExportViewModel {
        event ToggledEventHandler   UseSrcFolderToggled;
        event VbaExportEventHandler SelectedProjectsClicked;
        event VbaExportEventHandler CurrentProjectClicked;

        IRibbonToggle UseSrcFolderToggle  { get; }
        IRibbonButton SelectedProjectButton { get; }
        IRibbonButton CurrentProjectButton  { get; }

        void Attach(IBooleanSource srcToggleSource);
        void Invalidate();
    }
}

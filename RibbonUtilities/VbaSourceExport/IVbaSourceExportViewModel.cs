////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using VbaExportSelectedEventHandler = EventHandler<VbaExportSelectedEventArgs>;
    using VbaExportCurrentEventHandler = EventHandler<VbaExportCurrentEventArgs>;

    [CLSCompliant(false)]
    public interface IVbaSourceExportViewModel {
        [SuppressMessage("Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly")]
        event ToggledEventHandler   UseSrcFolderToggled;
        event VbaExportSelectedEventHandler SelectedProjectsClicked;
        event VbaExportCurrentEventHandler  CurrentProjectClicked;

        IRibbonToggle UseSrcFolderToggle  { get; }
        IRibbonButton SelectedProjectButton { get; }
        IRibbonButton CurrentProjectButton  { get; }

        void Attach(IBooleanSource srcToggleSource);
        void Invalidate();
    }
}

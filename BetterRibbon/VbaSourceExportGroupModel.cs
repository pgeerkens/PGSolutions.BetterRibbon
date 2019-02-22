////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonUtilities.LinksAnalysis;

namespace PGSolutions.BetterRibbon {
    internal sealed class VbaSourceExportGroupModel : AbstractRibbonGroupModel {
        public VbaSourceExportGroupModel(RibbonGroupViewModel viewModel, IRibbonFactory factory, string suffix)
        : base(viewModel) {
            Suffix = suffix;

            DestIsSrc      = factory.NewRibbonToggleModel($"UseSrcFolderToggle{suffix}",
                                OnUseSrcFolderToggled, false.ToggleImage());
            ExportSelected = factory.NewRibbonButtonModel($"SelectedProjectButton{suffix}",
                                OnExportSelected, "SaveAll");
            ExportCurrent  = factory.NewRibbonButtonModel($"CurrentProjectButton{suffix}",
                                OnExportCurrent, "FileSaveAs");

            Invalidate();
        }

        public event EventHandler<EventArgs<bool>> UseSrcFolderToggled;
        public event EventHandler ExportSelectedClicked;
        public event EventHandler ExportCurrentClicked;

        public IRibbonToggleModel DestIsSrc      { get; }

        public IRibbonButtonModel ExportSelected { get; }

        public IRibbonButtonModel ExportCurrent  { get; }

        public string             Suffix         { get; }

        private void OnUseSrcFolderToggled(object sender, bool isPressed)
        => UseSrcFolderToggled?.Invoke(sender, new EventArgs<bool>(isPressed));

        private void OnExportCurrent(object sender, EventArgs e)  => ExportCurrentClicked?.Invoke(sender,e);

        private void OnExportSelected(object sender, EventArgs e) => ExportSelectedClicked?.Invoke(sender,e);
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonUtilities.LinksAnalysis;

namespace PGSolutions.BetterRibbon {
    internal sealed class VbaSourceExportGroupModel : AbstractRibbonGroupModel {
        public VbaSourceExportGroupModel(GroupVM viewModel, IRibbonFactory factory, string suffix)
        : base(viewModel) {
            Suffix = suffix;

            DestIsSrc      = factory.NewToggleModel($"UseSrcFolderToggle{suffix}",
                                OnUseSrcFolderToggled, false.ToggleImage());
            ExportSelected = factory.NewButtonModel($"SelectedProjectButton{suffix}",
                                OnExportSelected, "SaveAll");
            ExportCurrent  = factory.NewButtonModel($"CurrentProjectButton{suffix}",
                                OnExportCurrent, "FileSaveAs");

            Invalidate();
        }

        public event EventHandler<EventArgs<bool>> UseSrcFolderToggled;
        public event EventHandler ExportSelectedClicked;
        public event EventHandler ExportCurrentClicked;

        public IToggleModel DestIsSrc      { get; }

        public IButtonModel ExportSelected { get; }

        public IButtonModel ExportCurrent  { get; }

        public string             Suffix         { get; }

        private void OnUseSrcFolderToggled(object sender, bool isPressed)
        => UseSrcFolderToggled?.Invoke(sender, new EventArgs<bool>(isPressed));

        private void OnExportCurrent(object sender, EventArgs e)  => ExportCurrentClicked?.Invoke(sender,e);

        private void OnExportSelected(object sender, EventArgs e) => ExportSelectedClicked?.Invoke(sender,e);
    }
}

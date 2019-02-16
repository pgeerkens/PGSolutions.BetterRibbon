////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon {
    internal sealed class VbaSourceExportGroupModel : AbstractRibbonGroupModel {
        public VbaSourceExportGroupModel(RibbonGroupViewModel viewModel, string suffix)
        : this(viewModel, suffix, null) { }

        public VbaSourceExportGroupModel(RibbonGroupViewModel viewModel, string suffix, IRibbonControlStrings strings)
        : base(viewModel, strings) {
            Suffix = suffix;

            DestIsSrc      = GetModel<RibbonToggleButton>($"UseSrcFolderToggle{suffix}", OnUseSrcFolderToggled, true, true, false.ToggleImage());
            ExportSelected = GetModel<RibbonButton>($"SelectedProjectButton{suffix}", OnExportSelected, true, true, "SaveAll");
            ExportCurrent  = GetModel<RibbonButton>($"CurrentProjectButton{suffix}", OnExportCurrent, true, true, "FileSaveAs");

            Invalidate();
        }

        public string Suffix { get; }

        public event EventHandler<EventArgs<bool>> UseSrcFolderToggled;
        public event ClickedEventHandler ExportSelectedClicked;
        public event ClickedEventHandler ExportCurrentClicked;

        public RibbonToggleModel DestIsSrc      { get; }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        public RibbonButtonModel ExportSelected { get; }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        public RibbonButtonModel ExportCurrent  { get; }

        private void OnUseSrcFolderToggled(object sender, bool isPressed)
        => UseSrcFolderToggled?.Invoke(sender, new EventArgs<bool>(isPressed));

        private void OnExportCurrent(object sender)  => ExportCurrentClicked?.Invoke(sender);

        private void OnExportSelected(object sender) => ExportSelectedClicked?.Invoke(sender);
    }
}

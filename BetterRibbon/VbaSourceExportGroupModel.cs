////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;
using PGSolutions.RibbonUtilities.LinksAnalysis;

namespace PGSolutions.BetterRibbon {
    internal sealed class VbaSourceExportGroupModel : AbstractRibbonGroupModel {
        public VbaSourceExportGroupModel(IModelFactory factory, IGroupVM viewModel, string suffix)
        : base(viewModel, factory.GetStrings(viewModel.Id)) {
            Suffix = suffix;

            DestIsSrc      = factory.NewToggleModel($"UseSrcFolderToggle{suffix}",
                                OnUseSrcFolderToggled, false.ToggleImage());
            ExportSelected = factory.NewButtonModel($"SelectedProjectButton{suffix}",
                                OnExportSelected, "SaveAll".ToImageObject());
            ExportCurrent  = factory.NewButtonModel($"CurrentProjectButton{suffix}",
                                OnExportCurrent, "FileSaveAs".ToImageObject());

            Invalidate();
        }

        public event ToggledEventHandler UseSrcFolderToggled;
        public event ClickedEventHandler ExportSelectedClicked;
        public event ClickedEventHandler ExportCurrentClicked;

        public IToggleModel DestIsSrc      { get; }

        public IButtonModel ExportSelected { get; }

        public IButtonModel ExportCurrent  { get; }

        public string       Suffix         { get; }

        private void OnUseSrcFolderToggled(IRibbonControl control, bool isPressed)
        => UseSrcFolderToggled?.Invoke(control, isPressed);

        private void OnExportCurrent(IRibbonControl control)  => ExportCurrentClicked?.Invoke(control);

        private void OnExportSelected(IRibbonControl control) => ExportSelectedClicked?.Invoke(control);
    }
}

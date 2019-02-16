////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;

namespace PGSolutions.BetterRibbon {
    [CLSCompliant(false)]
    public class Dispatcher : IRibbonDispatcher {
        public Dispatcher(BetterRibbonModel model) => Model = model;

        internal BetterRibbonModel     Model     { get; }

        /// <inheritdoc/>
        public void Invalidate() {
            Model.BrandingModel?.Invalidate();
            Model.LinksAnalysisModel?.Invalidate();
            Model.VbaSourceExportModel?.Invalidate();
            Model.CustomButtonsModel?.Invalidate();
        }

        /// <inheritdoc/>
        public void InvalidateCustomControlsGroup() => Model.CustomButtonsModel?.Invalidate();

        /// <inheritdoc/>
        public void InvalidateControl(string ControlId) => Model.ViewModel?.InvalidateControl(ControlId);

        /// <inheritdoc/>
        public void DetachProxy(string controlId)
        => Model.CustomButtonsModel.GetControl<IRibbonCommon>(controlId)?.Detach();

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed",
                Justification = "Matches COM usage.")]
        public IRibbonControlStrings NewControlStrings(string label,
                string screenTip = null, string superTip = null,
                string keyTip = null, string alternateLabel = null, string description = null) =>
            Model.ViewModel.RibbonFactory.NewControlStrings(label, screenTip,
                    superTip, keyTip, alternateLabel, description);

        /// <inheritdoc/>
        public ISelectableItemModel NewSelectableModel(string controlID, IRibbonControlStrings strings) {
            var vm = Model.ViewModel.RibbonFactory.NewSelectableItem(controlID);
            var model = new SelectableItemModel(id => vm, strings, true, true)
                        .Attach(controlID);
            return model;
        }

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonButtonModel NewRibbonButtonModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => Model.CustomButtonsModel.NewButtonModel(strings, image, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonButtonModel NewRibbonButtonModelMso(IRibbonControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => Model.CustomButtonsModel.NewButtonModel(strings, imageMso, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonToggleModel NewRibbonToggleModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => Model.CustomButtonsModel.NewToggleModel(strings, image, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonToggleModel NewRibbonToggleModelMso(IRibbonControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => Model.CustomButtonsModel.NewToggleModel(strings, imageMso, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonDropDownModel NewRibbonDropDownModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => Model.CustomButtonsModel.NewDropDownModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonGroupModel NewRibbonGroupModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => new RibbonGroupModel(id => Model.CustomButtonsModel.GetControl<RibbonGroupViewModel>(id),
                strings, isEnabled, isVisible, Model.CustomButtonsModel);
    }
}

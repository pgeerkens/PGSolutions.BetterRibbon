////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;

namespace PGSolutions.BetterRibbon {
    /// <summary>.</summary>
    [Description("The (top-level) ViewModel for the ribbon interface.")]
    [CLSCompliant(false)]
    public class Dispatcher : IRibbonDispatcher {
        public Dispatcher(BetterRibbonModel model) => Model = model;

        internal BetterRibbonModel     Model     { get; }

        /// <inheritdoc/>
        public void Invalidate() => Model.Invalidate();

        /// <inheritdoc/>
        public void DetachProxy(string controlId)
        => Model.GetCustomControl<IRibbonCommon>(controlId)?.Detach();

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonControlStrings NewControlStrings(string label,
                string screenTip = null, string superTip = null,
                string keyTip = null, string alternateLabel = null, string description = null)
        =>  Model.ViewModel.RibbonFactory.NewControlStrings(label, screenTip,
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
        => Model.NewRibbonButtonModel(strings, image, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonButtonModel NewRibbonButtonModelMso(IRibbonControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => Model.NewRibbonButtonModel(strings, imageMso, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonToggleModel NewRibbonToggleModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => Model.NewRibbonToggleModel(strings, image, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonToggleModel NewRibbonToggleModelMso(IRibbonControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => Model.NewRibbonToggleModel(strings, imageMso, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonDropDownModel NewRibbonDropDownModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => Model.NewRibbonDropDownModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonGroupModel NewRibbonGroupModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => Model.NewRibbonGroupModel(strings, isEnabled, isVisible);
    }
}

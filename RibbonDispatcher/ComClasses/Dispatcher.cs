////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>COM-visible implementation of the interface <see cref="IRibbonDispatcher"/>.</summary>
    [Description("The (top-level) ViewModel for the ribbon interface.")]
    [CLSCompliant(false)]
    public class Dispatcher : IRibbonDispatcher {
        /// <summary>.</summary>
        public Dispatcher(AbstractRibbonTabModel tabModel) => TabModel = tabModel;

        internal AbstractRibbonTabModel TabModel     { get; }

        /// <inheritdoc/>
        public void Invalidate() => TabModel.Invalidate();

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonControlStrings NewControlStrings(string label,
                string screenTip = null, string superTip = null,
                string keyTip = null, string alternateLabel = null, string description = null)
        =>  TabModel.ViewModel.RibbonFactory.NewControlStrings(label, screenTip,
                    superTip, keyTip, alternateLabel, description);

        /// <inheritdoc/>
        public ISelectableItemModel NewSelectableModel(string controlID, IRibbonControlStrings strings) {
            var vm = TabModel.ViewModel.RibbonFactory.NewSelectableItem(controlID);
            var model = new SelectableItemModel(id => vm, strings, true, true)
                        .Attach(controlID);
            return model;
        }

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonButtonModel NewRibbonButtonModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => TabModel.NewRibbonButtonModel(strings, image, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonButtonModel NewRibbonButtonModelMso(IRibbonControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => TabModel.NewRibbonButtonModel(strings, imageMso, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonToggleModel NewRibbonToggleModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => TabModel.NewRibbonToggleModel(strings, image, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonToggleModel NewRibbonToggleModelMso(IRibbonControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => TabModel.NewRibbonToggleModel(strings, imageMso, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonDropDownModel NewRibbonDropDownModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => TabModel.NewRibbonDropDownModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonGroupModel NewRibbonGroupModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => TabModel.NewRibbonGroupModel(strings, isEnabled, isVisible);
    }
}

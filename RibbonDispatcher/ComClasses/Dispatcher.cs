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
    public class Dispatcher : IDispatcher {
        /// <summary>.</summary>
        public Dispatcher(AbstractRibbonTabModel tabModel) => TabModel = tabModel;

        internal AbstractRibbonTabModel TabModel     { get; }

        /// <inheritdoc/>
        public void Invalidate() => TabModel.Invalidate();

        /// <inheritdoc/>
        public void DetachProxy(string controlId) => TabModel.DetachProxy(controlId);

        private IRibbonFactory RibbonFactory => TabModel.ViewModel.RibbonFactory;

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IControlStrings NewControlStrings(string label,
                string screenTip = null, string superTip = null,
                string keyTip = null, string alternateLabel = null, string description = null)
        =>  RibbonFactory.NewControlStrings(label, screenTip, superTip, keyTip, alternateLabel, description);

        /// <inheritdoc/>
        public SelectableItemModel NewSelectableModel(string controlID, IControlStrings strings) {
            var vm = RibbonFactory.NewSelectableItem(controlID);
            var model = new SelectableItemModel(id => vm, strings, true, true);
            model.Attach(controlID);
            return model;
        }

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ButtonModel NewButtonModel(IControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => RibbonFactory.NewButtonModel(strings, new ImageObject(image), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ButtonModel NewButtonModelMso(IControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => RibbonFactory.NewButtonModel(strings, new ImageObject(imageMso), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ToggleModel NewToggleModel(IControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => RibbonFactory.NewToggleModel(strings, new ImageObject(image), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ToggleModel NewToggleModelMso(IControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => RibbonFactory.NewToggleModel(strings, new ImageObject(imageMso), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public DropDownModel NewDropDownModel(IControlStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => RibbonFactory.NewDropDownModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public EditBoxModel NewEditBoxModel(IControlStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => RibbonFactory.NewEditBoxModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public GroupModel NewGroupModel(IControlStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => RibbonFactory.NewGroupModel(strings, isEnabled, isVisible);
    }
}

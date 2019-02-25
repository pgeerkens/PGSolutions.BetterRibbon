////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IControlStrings;

    /// <summary>COM-visible implementation of the interface <see cref="IRibbonDispatcher"/>.</summary>
    [Description("The (top-level) ViewModel for the ribbon interface.")]
    [CLSCompliant(false)]
    internal class ModelFactory : IModelFactory {
        /// <summary>.</summary>
        public ModelFactory(AbstractRibbonTabModel tabModel) => TabModel = tabModel;

        internal AbstractRibbonTabModel TabModel        { get; }

        private  ViewModelFactory       ViewModelFactory => TabModel.ViewModel.ViewModelFactory;

        /// <inheritdoc/>
        public void Invalidate() => TabModel.Invalidate();

        /// <inheritdoc/>
        public void DetachProxy(string controlId) => TabModel.DetachProxy(controlId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IStrings NewControlStrings(string label,
                string screenTip = null, string superTip = null,
                string keyTip = null, string alternateLabel = null, string description = null)
        =>  ViewModelFactory.NewControlStrings(label, screenTip, superTip, keyTip, alternateLabel, description);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IGroupModel NewGroupModel(IStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewGroupModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IButtonModel NewButtonModel(IStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewButtonModel(strings, new ImageObject(image), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IButtonModel NewButtonModelMso(IStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewButtonModel(strings, new ImageObject(imageMso), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IToggleModel NewToggleModel(IStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewToggleModel(strings, new ImageObject(image), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IToggleModel NewToggleModelMso(IStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewToggleModel(strings, new ImageObject(imageMso), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IEditBoxModel NewEditBoxModel(IStrings strings, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewEditBoxModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IDropDownModel NewDropDownModel(IStrings strings, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewDropDownModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IComboBoxModel NewComboBoxModel(IStrings strings, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewComboBoxModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ILabelModel NewLabelModel(IStrings strings, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewLabelModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IMenuModel NewMenuModel(IStrings strings, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewMenuModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ISplitButtonModel NewSplitButtonModel(IStrings splitStrings, IStrings buttonStrings,
                IStrings menuStrings, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewSplitButtonModel(splitStrings, buttonStrings, menuStrings, isEnabled, isVisible);

        /// <inheritdoc/>
        public ISelectableItemModel NewSelectableModel(string controlID, IStrings strings)
        => ViewModelFactory.NewSelectableModel(controlID, strings);
    }
}

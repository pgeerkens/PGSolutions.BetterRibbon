////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings  = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>COM-visible implementation of the interface <see cref="IModelFactory"/>.</summary>
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
        public IStrings NewControlStrings(string label, string screenTip, string superTip,
                string keyTip)
        => new ControlStrings(label, screenTip, superTip, keyTip);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IStrings2 NewControlStrings2(string label, string screenTip, string superTip,
                string keyTip, string description)
        =>  new ControlStrings2(label, screenTip, superTip, keyTip, description);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IGroupModel NewGroupModel(IStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewGroupModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IButtonModel NewButtonModel(IStrings2 strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewButtonModel(strings, new ImageObject(image), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IButtonModel NewButtonModelMso(IStrings2 strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewButtonModel(strings, new ImageObject(imageMso), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IToggleModel NewToggleModel(IStrings2 strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewToggleModel(strings, new ImageObject(image), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IToggleModel NewToggleModelMso(IStrings2 strings,
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
        public IMenuModel NewMenuModel(IStrings2 strings, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewMenuModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ISplitButtonModel NewSplitToggleButtonModel(IStrings splitStrings, IStrings2 buttonStrings,
                IStrings2 menuStrings, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewSplitToggleButtonModel(splitStrings, buttonStrings, menuStrings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ISplitButtonModel NewSplitPressButtonModel(IStrings splitStrings, IStrings2 buttonStrings,
                IStrings2 menuStrings, bool isEnabled = true, bool isVisible = true)
        => ViewModelFactory.NewSplitPressButtonModel(splitStrings, buttonStrings, menuStrings, isEnabled, isVisible);

        /// <inheritdoc/>
        public ISelectableItemModel NewSelectableModel(string controlID, IStrings strings)
        => ViewModelFactory.NewSelectableModel(controlID, strings);
    }
}

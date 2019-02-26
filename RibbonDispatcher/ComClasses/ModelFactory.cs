////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings  = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>COM-visible implementation of the interface <see cref="IModelFactory"/>.</summary>
    internal class ModelFactory : IModelFactory, IModelFactoryInternal {
        /// <summary>.</summary>
        public ModelFactory(ViewModelFactory viewModelFactory, IResourceLoader manager) {
            ViewModelFactory = viewModelFactory;
            ResourceManager = manager;
        }

        public IResourceLoader       ResourceManager  { get; }

        public ViewModelFactory       ViewModelFactory { get; }

        private IModelFactoryInternal _factory => this;

        /// <inheritdoc/>
        public void DetachProxy(string controlId) => ViewModelFactory.GetControl<IControlVM>(controlId).Detach();

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
        public IGroupModel NewGroupModel(string stringsId,
                bool isEnabled = true, bool isVisible = true)
        => _factory.NewGroupModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IButtonModel NewButtonModel(string stringsId,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => _factory.NewButtonModel(stringsId, new ImageObject(image), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IButtonModel NewButtonModelMso(string stringsId,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => _factory.NewButtonModel(stringsId, new ImageObject(imageMso), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IToggleModel NewToggleModel(string stringsId,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => _factory.NewToggleModel(stringsId, new ImageObject(image), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IToggleModel NewToggleModelMso(string stringsId,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => _factory.NewToggleModel(stringsId, new ImageObject(imageMso), isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IComboBoxModel NewComboBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => _factory.NewComboBoxModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IEditBoxModel NewEditBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => _factory.NewEditBoxModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IDropDownModel NewDropDownModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => _factory.NewDropDownModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ILabelModel NewLabelModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => ModelFactoryExtensions.NewLabelModel(this, stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IMenuModel NewMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => _factory.NewMenuModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ISplitButtonModel NewSplitToggleButtonModel(string splitStringId, string menuStringId,
                string toggleStringId, bool isEnabled = true, bool isVisible = true)
        => _factory.NewSplitToggleButtonModel(splitStringId, menuStringId, toggleStringId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ISplitButtonModel NewSplitPressButtonModel(string splitStringId, string menuStringId,
                string buttonStringId, bool isEnabled = true, bool isVisible = true)
        => _factory.NewSplitPressButtonModel(splitStringId, menuStringId, buttonStringId,  isEnabled, isVisible);

        /// <inheritdoc/>
        public ISelectableItemModel NewSelectableModel(string controlID)
        => _factory.NewSelectableModel(controlID);
    }
}

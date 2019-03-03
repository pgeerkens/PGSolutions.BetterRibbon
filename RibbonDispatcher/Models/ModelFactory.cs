////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>COM-visible implementation of the interface <see cref="IModelFactory"/>.</summary>
    public class ModelFactory : AbstractModelFactory, IModelFactory {
        /// <summary>.</summary>
        internal ModelFactory(ViewModelFactory viewModelFactory, IResourceLoader manager)
        : base(viewModelFactory, manager) { }

        /// <inheritdoc/>
        public void DetachProxy(string controlId) => ViewModelFactory.GetControl<IControlVM>(controlId).Detach();

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IStrings NewControlStrings(string label, string screenTip, string superTip, string keyTip=null)
        => new ControlStrings(label, screenTip, superTip, keyTip);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IStrings2 NewControlStrings2(string label, string screenTip, string superTip, string keyTip=null,
                string description=null)
        =>  new ControlStrings2(label, screenTip, superTip, keyTip, description);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IGroupModel NewGroupModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewGroupModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IButtonModel NewButtonModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewButtonModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IToggleModel NewToggleModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewToggleModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IEditBoxModel NewEditBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewEditBoxModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IDropDownModel NewDropDownModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewDropDownModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IStaticDropDownModel NewStaticDropDownModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewStaticDropDownModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IComboBoxModel NewComboBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewComboBoxModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IStaticComboBoxModel NewStaticComboBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewStaticComboBoxModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new ILabelControlModel NewLabelControlModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewLabelControlModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IMenuModel NewMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewMenuModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new ISplitToggleButtonModel NewSplitToggleButtonModel(string splitStringId, string menuStringId,
                string toggleStringId, bool isEnabled = true, bool isVisible = true)
        => base.NewSplitToggleButtonModel(splitStringId, menuStringId, toggleStringId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new ISplitPressButtonModel NewSplitPressButtonModel(string splitStringId, string menuStringId,
                string buttonStringId, bool isEnabled = true, bool isVisible = true)
        => base.NewSplitPressButtonModel(splitStringId, menuStringId, buttonStringId,  isEnabled, isVisible);

        /// <inheritdoc/>
        public new ISelectableItemModel NewSelectableModel(string controlID)
        => base.NewSelectableModel(controlID);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IGalleryModel NewGalleryModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewGalleryModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IStaticGalleryModel NewStaticGalleryModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewStaticGalleryModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IMenuSeparatorModel NewMenuSeparatorModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewMenuSeparatorModel(stringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        public IImageObject GetImage(IPictureDisp image) => new ImageObject(image);

        /// <inheritdoc/>
        public IImageObject GetImage(string imageMso) => new ImageObject(imageMso);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IDynamicMenuModel NewDynamicMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewDynamicMenuModel(stringsId, isEnabled, isVisible);
    }
}

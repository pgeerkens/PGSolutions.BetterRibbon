////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.Models {
    public partial class ModelFactory : AbstractModelFactory, IModelServer {
        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IGroupModel GetGroupModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewGroupModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IButtonModel GetButtonModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewButtonModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IToggleModel GetToggleModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewToggleModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IEditBoxModel GetEditBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewEditBoxModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IDropDownModel GetDropDownModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewDropDownModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IStaticDropDownModel GetStaticDropDownModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewStaticDropDownModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IComboBoxModel GetComboBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewComboBoxModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IStaticComboBoxModel GetStaticComboBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewStaticComboBoxModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ILabelControlModel GetLabelControlModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewLabelControlModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IMenuModel GetMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewMenuModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ISplitToggleButtonModel GetSplitToggleButtonModel(string stringsId, string menuStringsId,
                string toggleStringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewSplitToggleButtonModel(stringsId, menuStringsId, toggleStringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public ISplitPressButtonModel GetSplitPressButtonModel(string stringsId, string menuStringsId,
                string buttonStringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewSplitPressButtonModel(stringsId, menuStringsId, buttonStringsId,  isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        public ISelectableItemModel GetSelectableModel(string controlID)
        => base.NewSelectableModel(controlID).Attach(controlID);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IGalleryModel GetGalleryModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewGalleryModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IStaticGalleryModel GetStaticGalleryModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewStaticGalleryModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IMenuSeparatorModel GetMenuSeparatorModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewMenuSeparatorModel(stringsId, isEnabled, isVisible).Attach(stringsId);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IDynamicMenuModel GetDynamicMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewDynamicMenuModel(stringsId, isEnabled, isVisible).Attach(stringsId);
    }
}

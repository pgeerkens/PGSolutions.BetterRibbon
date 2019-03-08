////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

//** This (partial) implementation provides the COM-compatible, variable-parameter
//** wrapper to the base functionality defined in AbstractModelFactory: IModelFactory.
namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>COM-visible implementation of the interface <see cref="IModelFactory"/>.</summary>
    [Description("COM-visible implementation of the interface IModelFactory.")]
    public partial class ModelFactory : AbstractModelFactory {
        /// <summary>.</summary>
        public override IModelServer  AsServer  => this;

        internal ModelFactory(ViewModelFactory viewModelFactory, IResourceLoader manager)
        : base(viewModelFactory, manager) { }

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
        public new ISplitToggleButtonModel NewSplitToggleButtonModel(string splitStringsId, string menuStringsId,
                string toggleStringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewSplitToggleButtonModel(splitStringsId, menuStringsId, toggleStringsId, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new ISplitPressButtonModel NewSplitPressButtonModel(string splitStringsId, string menuStringsId,
                string buttonStringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewSplitPressButtonModel(splitStringsId, menuStringsId, buttonStringsId,  isEnabled, isVisible);

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
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public new IDynamicMenuModel NewDynamicMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true)
        => base.NewDynamicMenuModel(stringsId, isEnabled, isVisible);
    }
}

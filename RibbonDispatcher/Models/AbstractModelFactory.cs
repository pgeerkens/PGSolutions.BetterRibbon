////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>Internal implementation of the interface <see cref="IModelFactory"/>.</summary>
    /// <remarks>
    /// This class existsto expose the "evented" base classes to internal methods,
    /// while only the unevented COM-visible interfaces are exposed to VBA clients.
    /// </remarks>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability","CA1506:AvoidExcessiveClassCoupling")]
    public abstract class AbstractModelFactory: IModelFactory {
        /// <summary>.</summary>
        protected AbstractModelFactory(ViewModelFactory viewModelFactory, IResourceLoader manager) {
            ViewModelFactory = viewModelFactory;
            ResourceManager = manager;
        }

        internal IResourceLoader ResourceManager { get; }

        internal ViewModelFactory ViewModelFactory { get; }

        public abstract IModelFactory AsFactory { get; }

        public abstract IModelServer  AsServer  { get; }

        /// <inheritdoc/>
        public void DetachProxy(string controlId) => ViewModelFactory.GetControl<IControlVM>(controlId).Detach();

        /// <summary>Returns a new <see cref="IImageObject"/> from the supplied <see cref="IPictureDisp"/>.</summary>
        public IImageObject NewImageObject(IPictureDisp image) => new ImageObject(image);

        /// <summary>Returns a new <see cref="IImageObject"/> from the supplied MSO image name.</summary>
        public IImageObject NewImageObjectMso(string imageMso) => new ImageObject(imageMso);

        /// <summary>Creates, initializes and returns a new <see cref="GroupModel"/>.</summary>
        public IGroupModel NewGroupModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new GroupModel(GetControl<GroupVM>, GetStrings(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible };

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public IButtonModel NewButtonModel(string stringsId, bool isEnabled, bool isVisible)
        => new ButtonModel(GetControl<ButtonVM>, GetStrings2(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IButtonSource, IButtonVM, ButtonModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ToggleModel"/>.</summary>
        public IToggleModel NewToggleModel(string stringsId, bool isEnabled, bool isVisible)
        => new ToggleModel(GetControl<CheckBoxVM>, GetStrings2(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IToggleSource, IToggleVM, ToggleModel>();

        /// <summary>Creates, initializes and returns a new <see cref="EditBoxModel"/>.</summary>
        public IEditBoxModel NewEditBoxModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new EditBoxModel(GetControl<EditBoxVM>, GetStrings(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IEditBoxSource, IEditBoxVM, EditBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="DropDownModel"/>.</summary>
        public IDropDownModel NewDropDownModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new DropDownModel(GetControl<DropDownVM>, GetStrings(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IDropDownSource, IDropDownVM, DropDownModel>();

        /// <summary>Creates, initializes and returns a new <see cref="StaticDropDownModel"/>.</summary>
        public IStaticDropDownModel NewStaticDropDownModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new StaticDropDownModel(GetControl<StaticDropDownVM>, GetStrings(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IStaticDropDownSource, IDropDownVM, StaticDropDownModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ComboBoxModel"/>.</summary>
        public IComboBoxModel NewComboBoxModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new ComboBoxModel(GetControl<ComboBoxVM>, GetStrings(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IComboBoxSource, IComboBoxVM, ComboBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="StaticComboBoxModel"/>.</summary>
        public IStaticComboBoxModel NewStaticComboBoxModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new StaticComboBoxModel(GetControl<StaticComboBoxVM>, GetStrings(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IStaticComboBoxSource, IStaticComboBoxVM, StaticComboBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="LabelControlModel"/>.</summary>
        public ILabelControlModel NewLabelControlModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new LabelControlModel(GetControl<LabelControlVM>, GetStrings(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<ILabelControlSource, ILabelControlVM, LabelControlModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public IMenuModel NewMenuModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new MenuModel(GetControl<MenuVM>, GetStrings2(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IMenuSource, IMenuVM, MenuModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public ISplitToggleButtonModel NewSplitToggleButtonModel(
                string splitStringsId, string menuStringsId, string toggleStringsId,
                bool isEnabled, bool isVisible)
        => new SplitToggleButtonModel(GetControl<SplitToggleButtonVM>, GetStrings2(splitStringsId),
                new ToggleModel(GetControl<ToggleButtonVM>, GetStrings2(toggleStringsId)),
                new MenuModel(GetControl<MenuVM>, GetStrings2(menuStringsId)))
                { IsEnabled=isEnabled, IsVisible=isVisible }
            .InitializeModel<IToggleSource, ISplitToggleButtonVM, SplitToggleButtonModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public ISplitPressButtonModel NewSplitPressButtonModel(
                string splitStringsId, string menuStringsId, string buttonStringsId,
                bool isEnabled, bool isVisible)
        => new SplitPressButtonModel(GetControl<SplitPressButtonVM>, GetStrings2(splitStringsId),
                new ButtonModel(GetControl<ButtonVM>, GetStrings2(buttonStringsId)),
                new MenuModel(GetControl<MenuVM>, GetStrings2(menuStringsId)))
                { IsEnabled=isEnabled, IsVisible=isVisible }
            .InitializeModel<IButtonSource, ISplitPressButtonVM, SplitPressButtonModel>();

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        public ISelectableItemModel NewSelectableModel(string controlID)
        => new SelectableItemModel(GetStrings(controlID)).Attach(controlID);

        /// <summary>Creates, initializes and returns a new <see cref="GalleryModel"/>.</summary>
        public IGalleryModel NewGalleryModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new GalleryModel(GetControl<GalleryVM>, GetStrings2(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IGallerySource, IGalleryVM, GalleryModel>();

        /// <summary>Creates, initializes and returns a new <see cref="StaticGalleryModel"/>.</summary>
        public IStaticGalleryModel NewStaticGalleryModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new StaticGalleryModel(GetControl<StaticGalleryVM>, GetStrings2(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IStaticGallerySource, IStaticGalleryVM, StaticGalleryModel>();

        /// <summary>Creates, initializes and returns a new <see cref="MenuSeparatorModel"/>.</summary>
        public IMenuSeparatorModel NewMenuSeparatorModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new MenuSeparatorModel(GetControl<MenuSeparatorVM>, GetStrings(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IMenuSeparatorSource, IMenuSeparatorVM, MenuSeparatorModel>();

        /// <inheritdoc/>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design","CA1026:DefaultParametersShouldNotBeUsed")]
        public IStrings NewControlStrings(string label, string screenTip, string superTip, string keyTip=null)
        => new ControlStrings(label, screenTip, superTip, keyTip);

        /// <inheritdoc/>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design","CA1026:DefaultParametersShouldNotBeUsed")]
        public IStrings2 NewControlStrings2(string label, string screenTip, string superTip, string keyTip=null,
                string description=null)
        =>  new ControlStrings2(label, screenTip, superTip, keyTip, description);

        /// <inheritdoc/>
        public IStrings GetStrings(string id) => ResourceManager.GetControlStrings(id);

        /// <inheritdoc/>
        public IStrings2 GetStrings2(string id) => ResourceManager.GetControlStrings2(id);

        /// <inheritdoc/>
        public IImageObject GetImage(IPictureDisp image) => new ImageObject(image);

        /// <inheritdoc/>
        public IImageObject GetImageMso(string imageMso) => new ImageObject(imageMso);

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public IDynamicMenuModel NewDynamicMenuModel(string stringsId,
                bool isEnabled, bool isVisible)
        => new DynamicMenuModel(GetControl<DynamicMenuVM>, GetStrings2(stringsId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IDynamicMenuSource, IDynamicMenuVM, DynamicMenuModel>();

        /// <summary>Returns the control identifed by <paramref name="controlId"/> if one meeting the type constraints exists; else null.</summary>
        /// <typeparam name="TControl">Type filter for the returned control, constrained as a class,<see cref="IControlVM"/>.</typeparam>
        /// <param name="controlId"></param>
        /// <returns>A <typeparamref name="TControl"/></returns>
        internal TControl GetControl<TControl>(string controlId) where TControl: class,IControlVM
        => ViewModelFactory.GetControl<TControl>(controlId);
    }
}

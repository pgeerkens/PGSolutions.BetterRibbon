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

    /// <summary>Internal implementation of the interface <see cref="IModelFactory"/>.</summary>
    /// <remarks>
    /// This class existsto expose the "evented" base classes to internal methods,
    /// while only the unevented COM-visible interfaces are exposed to VBA clients.
    /// </remarks>
    public abstract class AbstractModelFactory {
        /// <summary>.</summary>
        protected AbstractModelFactory(ViewModelFactory viewModelFactory, IResourceLoader manager) {
            ViewModelFactory = viewModelFactory;
            ResourceManager = manager;
        }

        internal IResourceLoader ResourceManager { get; }

        internal ViewModelFactory ViewModelFactory { get; }

        /// <summary>Returns a new <see cref="IImageObject"/> from the supplied <see cref="IPictureDisp"/>.</summary>
        public IImageObject NewImageObject(IPictureDisp image) => new ImageObject(image);

        /// <summary>Returns a new <see cref="IImageObject"/> from the supplied MSO image name.</summary>
        public IImageObject NewImageObjectMso(string imageMso) => new ImageObject(imageMso);

        /// <summary>Creates, initializes and returns a new <see cref="GroupModel"/>.</summary>
        public IGroupModel NewGroupModel(string controlId,
                bool isEnabled, bool isVisible)
        => new GroupModel(GetControl<GroupVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible };

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public IButtonModel NewButtonModel(string controlId, bool isEnabled, bool isVisible)
        => new ButtonModel(GetControl<ButtonVM>, GetStrings2(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IButtonSource, IButtonVM, ButtonModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ToggleModel"/>.</summary>
        public IToggleModel NewToggleModel(string controlId, bool isEnabled, bool isVisible)
        => new ToggleModel(GetControl<CheckBoxVM>, GetStrings2(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IToggleSource, IToggleVM, ToggleModel>();

        /// <summary>Creates, initializes and returns a new <see cref="EditBoxModel"/>.</summary>
        public IEditBoxModel NewEditBoxModel(string controlId,
                bool isEnabled, bool isVisible)
        => new EditBoxModel(GetControl<EditBoxVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IEditBoxSource, IEditBoxVM, EditBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="DropDownModel"/>.</summary>
        public IDropDownModel NewDropDownModel(string controlId,
                bool isEnabled, bool isVisible)
        => new DropDownModel(GetControl<DropDownVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IDropDownSource, IDropDownVM, DropDownModel>();

        /// <summary>Creates, initializes and returns a new <see cref="StaticDropDownModel"/>.</summary>
        public IStaticDropDownModel NewStaticDropDownModel(string controlId,
                bool isEnabled, bool isVisible)
        => new StaticDropDownModel(GetControl<StaticDropDownVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IStaticDropDownSource, IDropDownVM, StaticDropDownModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ComboBoxModel"/>.</summary>
        public IComboBoxModel NewComboBoxModel(string controlId,
                bool isEnabled, bool isVisible)
        => new ComboBoxModel(GetControl<ComboBoxVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IComboBoxSource, IComboBoxVM, ComboBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="StaticComboBoxModel"/>.</summary>
        public IStaticComboBoxModel NewStaticComboBoxModel(string controlId,
                bool isEnabled, bool isVisible)
        => new StaticComboBoxModel(GetControl<StaticComboBoxVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IStaticComboBoxSource, IStaticComboBoxVM, StaticComboBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="LabelControlModel"/>.</summary>
        public ILabelControlModel NewLabelControlModel(string controlId,
                bool isEnabled, bool isVisible)
        => new LabelControlModel(GetControl<LabelControlVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<ILabelControlSource, ILabelControlVM, LabelControlModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public IMenuModel NewMenuModel(string controlId,
                bool isEnabled, bool isVisible)
        => new MenuModel(GetControl<MenuVM>, GetStrings2(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IMenuSource, IMenuVM, MenuModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public ISplitToggleButtonModel NewSplitToggleButtonModel(
                string splitStringId, string menuStringId, string toggleStringId,
                bool isEnabled, bool isVisible)
        => new SplitToggleButtonModel(GetControl<SplitToggleButtonVM>, GetStrings2(splitStringId),
                new ToggleModel(GetControl<ToggleButtonVM>, GetStrings2(toggleStringId)),
                new MenuModel(GetControl<MenuVM>, GetStrings2(menuStringId)))
                { IsEnabled=isEnabled, IsVisible=isVisible }
            .InitializeModel<IToggleSource, ISplitToggleButtonVM, SplitToggleButtonModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public ISplitPressButtonModel NewSplitPressButtonModel(
                string splitStringId, string menuStringId, string buttonStringId,
                bool isEnabled, bool isVisible)
        => new SplitPressButtonModel(GetControl<SplitPressButtonVM>, GetStrings2(splitStringId),
                new ButtonModel(GetControl<ButtonVM>, GetStrings2(buttonStringId)),
                new MenuModel(GetControl<MenuVM>, GetStrings2(menuStringId)))
                { IsEnabled=isEnabled, IsVisible=isVisible }
            .InitializeModel<IButtonSource, ISplitPressButtonVM, SplitPressButtonModel>();

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "factory")]
        public ISelectableItemModel NewSelectableModel(string controlID)
        => new SelectableItemModel(GetStrings(controlID)).Attach(controlID);

        /// <summary>Creates, initializes and returns a new <see cref="GalleryModel"/>.</summary>
        public IGalleryModel NewGalleryModel(string controlId,
                bool isEnabled, bool isVisible)
        => new GalleryModel(GetControl<GalleryVM>, GetStrings2(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IGallerySource, IGalleryVM, GalleryModel>();

        /// <summary>Creates, initializes and returns a new <see cref="StaticGalleryModel"/>.</summary>
        public IStaticGalleryModel NewStaticGalleryModel(string controlId,
                bool isEnabled, bool isVisible)
        => new StaticGalleryModel(GetControl<StaticGalleryVM>, GetStrings2(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IStaticGallerySource, IStaticGalleryVM, StaticGalleryModel>();

        /// <summary>Creates, initializes and returns a new <see cref="MenuSeparatorModel"/>.</summary>
        public IMenuSeparatorModel NewMenuSeparatorModel(string controlId,
                bool isEnabled, bool isVisible)
        => new MenuSeparatorModel(GetControl<MenuSeparatorVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IMenuSeparatorSource, IMenuSeparatorVM, MenuSeparatorModel>();

        internal TControl GetControl<TControl>(string controlId) where TControl : class, IControlVM
        => ViewModelFactory.GetControl<TControl>(controlId);

        public IStrings GetStrings(string id) => ResourceManager.GetControlStrings(id);

        public IStrings2 GetStrings2(string id) => ResourceManager.GetControlStrings2(id);

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public IDynamicMenuModel NewDynamicMenuModel(string controlId,
                bool isEnabled, bool isVisible)
        => new DynamicMenuModel(GetControl<DynamicMenuVM>, GetStrings2(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IDynamicMenuSource, IDynamicMenuVM, DynamicMenuModel>();
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>Internal implementation of the interface <see cref="IModelFactory"/>.</summary>
    public abstract class AbstractModelFactory {
        /// <summary>.</summary>
        protected AbstractModelFactory(ViewModelFactory viewModelFactory, IResourceLoader manager) {
            ViewModelFactory = viewModelFactory;
            ResourceManager = manager;
        }

        internal IResourceLoader ResourceManager { get; }

        internal ViewModelFactory ViewModelFactory { get; }

        /// <summary>Creates, initializes and returns a new <see cref="GroupModel"/>.</summary>
        public GroupModel NewGroupModel(string controlId,
                bool isEnabled, bool isVisible)
        => new GroupModel(GetControl<GroupVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible };

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public ButtonModel NewButtonModel(string controlId,
                ImageObject image, bool isEnabled, bool isVisible)
        => new ButtonModel(GetControl<ButtonVM>, GetStrings2(controlId))
                { Image=image, IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IButtonSource, IButtonVM, ButtonModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ToggleModel"/>.</summary>
        public ToggleModel NewToggleModel(string controlId,
                ImageObject image, bool isEnabled, bool isVisible)
        => new ToggleModel(GetControl<CheckBoxVM>, GetStrings2(controlId))
                { Image=image, IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IToggleSource, IToggleVM, ToggleModel>();

        /// <summary>Creates, initializes and returns a new <see cref="EditBoxModel"/>.</summary>
        public EditBoxModel NewEditBoxModel(string controlId,
                bool isEnabled, bool isVisible)
        => new EditBoxModel(GetControl<EditBoxVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IEditBoxSource, IEditBoxVM, EditBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ComboBoxModel"/>.</summary>
        public ComboBoxModel NewComboBoxModel(string controlId,
                bool isEnabled, bool isVisible)
        => new ComboBoxModel(GetControl<ComboBoxVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IComboBoxSource, IComboBoxVM, ComboBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="DropDownModel"/>.</summary>
        public DropDownModel NewDropDownModel(string controlId,
                bool isEnabled, bool isVisible)
        => new DropDownModel(GetControl<DropDownVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IDropDownSource, IDropDownVM, DropDownModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public LabelModel NewLabelModel(string controlId,
                bool isEnabled, bool isVisible)
        => new LabelModel(GetControl<LabelVM>, GetStrings(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<ILabelSource, ILabelVM, LabelModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public MenuModel NewMenuModel(string controlId,
                bool isEnabled, bool isVisible)
        => new MenuModel(GetControl<MenuVM>, GetStrings2(controlId))
                { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IMenuSource, IMenuVM, MenuModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public SplitToggleButtonModel NewSplitToggleButtonModel(
                string splitStringId, string menuStringId, string toggleStringId,
                bool isEnabled, bool isVisible)
        => new SplitToggleButtonModel(GetControl<SplitToggleButtonVM>, GetStrings(splitStringId),
                new ToggleModel(GetControl<ToggleButtonVM>, GetStrings2(toggleStringId)),
                new MenuModel(ViewModelFactory.GetControl<MenuVM>, GetStrings2(menuStringId)))
                { IsEnabled=isEnabled, IsVisible=isVisible }
            .InitializeModel<IToggleSource, ISplitToggleButtonVM, SplitToggleButtonModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public SplitPressButtonModel NewSplitPressButtonModel(
                string splitStringId, string menuStringId, string buttonStringId,
                bool isEnabled, bool isVisible)
        => new SplitPressButtonModel(GetControl<SplitPressButtonVM>, GetStrings(splitStringId),
                new ButtonModel(GetControl<ButtonVM>, GetStrings2(buttonStringId)),
                new MenuModel(GetControl<MenuVM>, GetStrings2(menuStringId)))
                { IsEnabled=isEnabled, IsVisible=isVisible }
            .InitializeModel<IButtonSource, ISplitPressButtonVM, SplitPressButtonModel>();

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "factory")]
        public ISelectableItemModel NewSelectableModel(string controlID)
        => new SelectableItemModel(GetStrings(controlID)).Attach(controlID);

        public TControl GetControl<TControl>(string controlId) where TControl : class, IControlVM
        => ViewModelFactory.GetControl<TControl>(controlId);

        public IStrings GetStrings(string id) => ResourceManager.GetControlStrings(id);

        public IStrings2 GetStrings2(string id) => ResourceManager.GetControlStrings2(id);
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings  = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>These extension methods on <see cref="ViewModelFactory"/> are the common link between <see cref="ControlModel{TSource, TCtrl}"/> objects created from VBA and C#.></summary>
    internal static partial class ViewModelFactoryExtensions {
        public static IStrings GetStrings(this IViewModelFactory vm, string controlId)
        => vm.ResourceManager.GetControlStrings(controlId);

        public static IStrings2 GetStrings2(this IViewModelFactory vm, string controlId)
        => vm.ResourceManager.GetControlStrings2(controlId);

        public static object LoadImage(this IViewModelFactory vm, string imageId)
        => vm.ResourceManager.GetImage(imageId);

        /// <summary>Creates, initializes and returns a new <see cref="GroupModel"/>.</summary>
        public static GroupModel NewGroupModel(this ViewModelFactory factory, IStrings strings,
                bool isEnabled, bool isVisible)
        => new GroupModel(factory.GetControl<GroupVM>, strings) { IsEnabled=isEnabled, IsVisible=isVisible };



        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public static ButtonModel NewButtonModel(this IModelFactoryInternal factory, string controlId,
                ImageObject image, bool isEnabled, bool isVisible)
        =>  new ButtonModel(factory.GetControl<ButtonVM>, factory.GetStrings2(controlId))
                            { Image=image, IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IButtonSource, IButtonVM, ButtonModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ToggleModel"/>.</summary>
        public static ToggleModel NewToggleModel(this IModelFactoryInternal factory, string controlId,
                ImageObject image, bool isEnabled, bool isVisible)
        =>  new ToggleModel(factory.GetControl<CheckBoxVM>, factory.GetStrings2(controlId))
                            { Image=image, IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IToggleSource, IToggleVM, ToggleModel>();

        /// <summary>Creates, initializes and returns a new <see cref="EditBoxModel"/>.</summary>
        public static EditBoxModel NewEditBoxModel(this IModelFactoryInternal factory, string controlId,
                bool isEnabled, bool isVisible)
        => new EditBoxModel(factory.GetControl<EditBoxVM>, factory.GetStrings(controlId)) 
                            { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IEditBoxSource, IEditBoxVM, EditBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ComboBoxModel"/>.</summary>
        public static ComboBoxModel NewComboBoxModel(this IModelFactoryInternal factory, string controlId,
                bool isEnabled, bool isVisible)
        => new ComboBoxModel(factory.GetControl<ComboBoxVM>, factory.GetStrings(controlId)) 
                            { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IComboBoxSource, IComboBoxVM, ComboBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="DropDownModel"/>.</summary>
        public static DropDownModel NewDropDownModel(this IModelFactoryInternal factory, string controlId,
                bool isEnabled, bool isVisible)
        => new DropDownModel(factory.GetControl<DropDownVM>, factory.GetStrings(controlId)) 
                            { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IDropDownSource, IDropDownVM, DropDownModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public static LabelModel NewLabelModel(this IModelFactoryInternal factory, string controlId,
                bool isEnabled, bool isVisible)
        => new LabelModel(factory.GetControl<LabelVM>, factory.GetStrings(controlId)) 
                         { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<ILabelSource, ILabelVM, LabelModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public static MenuModel NewMenuModel(this IModelFactoryInternal factory, string controlId,
                bool isEnabled, bool isVisible)
        => new MenuModel(factory.GetControl<MenuVM>, factory.GetStrings2(controlId)) 
                        { IsEnabled=isEnabled, IsVisible=isVisible }
                .InitializeModel<IMenuSource, IMenuVM, MenuModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public static SplitToggleButtonModel NewSplitToggleButtonModel(this IModelFactoryInternal factory,
                string splitStringId, string menuStringId, string toggleStringId,
                bool isEnabled, bool isVisible)
        => new SplitToggleButtonModel(factory.GetControl<SplitToggleButtonVM>, factory.GetStrings(splitStringId),
                new ToggleModel(factory.GetControl<ToggleButtonVM>, factory.GetStrings2(toggleStringId)),
                new MenuModel(factory.ViewModelFactory.GetControl<MenuVM>, factory.GetStrings2(menuStringId)))
                { IsEnabled=isEnabled, IsVisible=isVisible }
            .InitializeModel<IToggleSource, ISplitToggleButtonVM, SplitToggleButtonModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public static SplitPressButtonModel NewSplitPressButtonModel(this IModelFactoryInternal factory,
                string splitStringId, string menuStringId, string buttonStringId,
                bool isEnabled, bool isVisible)
        => new SplitPressButtonModel(factory.GetControl<SplitPressButtonVM>, factory.GetStrings(splitStringId),
                new ButtonModel(factory.GetControl<ButtonVM>, factory.GetStrings2(buttonStringId)),
                new MenuModel(factory.GetControl<MenuVM>, factory.GetStrings2(menuStringId)) )
                { IsEnabled=isEnabled, IsVisible=isVisible }
            .InitializeModel<IButtonSource, ISplitPressButtonVM, SplitPressButtonModel>();

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "factory")]
        public static ISelectableItemModel NewSelectableModel(this IModelFactoryInternal factory, string controlID)
        => new SelectableItemModel(factory.GetStrings(controlID)).Attach(controlID);

        public static TModel InitializeModel<TSource, TVM, TModel>(this TModel model)
            where TModel: ControlModel<TSource, TVM> where TSource: IControlSource where TVM: IControlVM {

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        private static TControl GetControl<TControl>(this IModelFactoryInternal factory, string controlId)
            where TControl: class,IControlVM
            => factory.ViewModelFactory.GetControl<TControl>(controlId);

        private static IStrings GetStrings(this IModelFactoryInternal factory, string id)
            => factory.ResourceManager.GetControlStrings(id);

        private static IStrings2 GetStrings2(this IModelFactoryInternal factory, string id)
            => factory.ResourceManager.GetControlStrings2(id);
    }
}

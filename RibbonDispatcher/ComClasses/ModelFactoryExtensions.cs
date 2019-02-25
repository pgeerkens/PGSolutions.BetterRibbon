////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IControlStrings;

    /// <summary>These extension methods on <see cref="ViewModelFactory"/> are the common link between <see cref="ControlModel{TSource, TCtrl}"/> objects created from VBA and C#.></summary>
    internal static partial class ViewModelFactoryExtensions {
        /// <summary>Creates, initializes and returns a new <see cref="GroupModel"/>.</summary>
        public static GroupModel NewGroupModel(this ViewModelFactory factory, IStrings strings,
                bool isEnabled, bool isVisible)
        => new GroupModel(factory.GetControl<GroupVM>, strings, isEnabled, isVisible);

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public static ButtonModel NewButtonModel(this ViewModelFactory factory, IStrings strings,
                ImageObject image, bool isEnabled, bool isVisible)
        => new ButtonModel(factory.GetControl<ButtonVM>, strings, image, isEnabled, isVisible)
                .InitializeModel<IButtonSource, IButtonVM, ButtonModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ToggleModel"/>.</summary>
        public static ToggleModel NewToggleModel(this ViewModelFactory factory, IStrings strings,
                ImageObject image, bool isEnabled, bool isVisible)
        => new ToggleModel(factory.GetControl<CheckBoxVM>, strings, image, isEnabled, isVisible)
                .InitializeModel<IToggleSource, IToggleControlVM, ToggleModel>();

        /// <summary>Creates, initializes and returns a new <see cref="EditBoxModel"/>.</summary>
        public static EditBoxModel NewEditBoxModel(this ViewModelFactory factory, IStrings strings,
                bool isEnabled, bool isVisible)
        => new EditBoxModel(factory.GetControl<EditBoxVM>, strings, isEnabled, isVisible)
                .InitializeModel<IEditBoxSource, IEditBoxVM, EditBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ComboBoxModel"/>.</summary>
        public static ComboBoxModel NewComboBoxModel(this ViewModelFactory factory, IStrings strings,
                bool isEnabled, bool isVisible)
        => new ComboBoxModel(factory.GetControl<ComboBoxVM>, strings, isEnabled, isVisible)
                .InitializeModel<IComboBoxSource, IComboBoxVM, ComboBoxModel>();

        /// <summary>Creates, initializes and returns a new <see cref="DropDownModel"/>.</summary>
        public static DropDownModel NewDropDownModel(this ViewModelFactory factory, IStrings strings,
                bool isEnabled, bool isVisible)
        => new DropDownModel(factory.GetControl<DropDownVM>, strings, isEnabled, isVisible)
                .InitializeModel<IDropDownSource, IDropDownVM, DropDownModel>();

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "factory")]
        public static ISelectableItemModel NewSelectableModel(this ViewModelFactory factory,
                string controlID, IStrings strings)
        => new SelectableItemModel(strings, true, true).Attach(controlID);

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public static LabelModel NewLabelModel(this ViewModelFactory factory, IStrings strings,
                bool isEnabled, bool isVisible)
        => new LabelModel(factory.GetControl<LabelVM>, strings, isEnabled, isVisible)
                .InitializeModel<ILabelSource, ILabelVM, LabelModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public static MenuModel NewMenuModel(this ViewModelFactory factory, IStrings strings,
                bool isEnabled, bool isVisible)
        => new MenuModel(factory.GetControl<MenuVM>, strings, isEnabled, isVisible)
                .InitializeModel<IMenuSource, IMenuVM, MenuModel>();

        /// <summary>Creates, initializes and returns a new <see cref="ButtonModel"/>.</summary>
        public static SplitButtonModel NewSplitButtonModel(this ViewModelFactory factory,
                IStrings splitStrings, IStrings buttonStrings, IStrings menuStrings,
                bool isEnabled, bool isVisible)
        => new SplitButtonModel(factory.GetControl<SplitButtonVM>, splitStrings,
            new ButtonModel(factory.GetControl<ButtonVM>, buttonStrings, ImageObject.Empty,true,true),
            new MenuModel(factory.GetControl<MenuVM>, menuStrings, true,true),
            isEnabled, isVisible)
                .InitializeModel<ISplitButtonSource, ISplitButtonVM, SplitButtonModel>();

        public static TModel InitializeModel<TSource, TVM, TModel>(this TModel model)
            where TModel : ControlModel<TSource, TVM> where TSource : IControlSource where TVM : IControlVM {

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }
    }
}

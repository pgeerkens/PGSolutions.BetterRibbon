////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IControlStrings;

    /// <summary>These extension methods on <see cref="ViewModelFactory"/> are the public API to C# for creation of objects subclassing <see cref="ControlModel{TSource, TCtrl}"/>.</summary>
    public static partial class PublicExtensions {
        /// <summary>Returns a new instance of an <see cref="IModelFactory"/>.</summary>
        /// <param name="model"></param>
        public static IModelFactory NewModelFactory(this AbstractRibbonTabModel model)
            => new ModelFactory(model);

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonButtonModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IButtonModel NewButtonModel(this ViewModelFactory factory, string id,
                ClickedEventHandler handler, ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewButtonModel(factory.GetStrings(id), image, isEnabled, isVisible);

            model.Clicked += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonToggleModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IToggleModel NewToggleModel(this ViewModelFactory factory, string id,
                ToggledEventHandler handler, ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewToggleModel(factory.GetStrings(id), image, isEnabled, isVisible);

            model.Toggled += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IEditBoxModel NewEditBoxModel(this ViewModelFactory factory, string id,
                EditedEventHandler handler, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewEditBoxModel(factory.GetStrings(id), isEnabled, isVisible);

            model.Edited += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IDropDownModel NewDropDownModel(this ViewModelFactory factory, string id,
                SelectionMadeEventHandler handler, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewDropDownModel(factory.GetStrings(id), isEnabled, isVisible);

            model.SelectionMade += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IComboBoxModel NewComboBoxModel(this ViewModelFactory factory, string id,
                EditedEventHandler handler,
                bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewComboBoxModel(factory.GetStrings(id), isEnabled, isVisible);

            model.Edited += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonLabelModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static ILabelModel NewLabelModel(this ViewModelFactory factory, string id,
                ClickedEventHandler handler, bool isEnabled = true, bool isVisible = true)
        => factory?.NewLabelModel(factory.GetStrings(id), isEnabled, isVisible)
                  ?.Attach(id);
    }

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

        public static TModel InitializeModel<TSource, TVM, TModel>(this TModel model)
            where TModel : ControlModel<TSource, TVM> where TSource : IControlSource where TVM : IControlVM {

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }
    }
}

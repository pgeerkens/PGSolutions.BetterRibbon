////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>These extension methods on <see cref="ViewModelFactory"/> are the public API to C# for creation of objects subclassing <see cref="ControlModel{TSource, TCtrl}"/>.</summary>
    public static partial class PublicFactoryExtensions {
        /// <summary>Returns a new instance of an <see cref="IModelFactory"/>.</summary>
        /// <param name="model"></param>
        public static IModelFactoryInternal NewModelFactory(this ViewModelFactory viewModelFactory, IResourceManager manager)
            => new ModelFactory(viewModelFactory, manager);

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonButtonModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IButtonModel NewButtonModel(this IModelFactoryInternal factory, string id,
                ClickedEventHandler handler, ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory.NewButtonModel(id, image, isEnabled, isVisible);

            model.Clicked += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonToggleModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IToggleModel NewToggleModel(this IModelFactoryInternal factory, string id,
                ToggledEventHandler handler, ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory.NewToggleModel(id, image, isEnabled, isVisible);

            model.Toggled += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IEditBoxModel NewEditBoxModel(this IModelFactoryInternal factory, string id,
                EditedEventHandler handler, bool isEnabled = true, bool isVisible = true) {
            var model = ViewModelFactoryExtensions.NewEditBoxModel(factory, id, isEnabled, isVisible);

            model.Edited += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IComboBoxModel NewComboBoxModel(this IModelFactoryInternal factory, string id,
                EditedEventHandler handler,
                bool isEnabled = true, bool isVisible = true) {
            var model = ViewModelFactoryExtensions.NewComboBoxModel(factory, id, isEnabled, isVisible);

            model.Edited += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IDropDownModel NewDropDownModel(this IModelFactoryInternal factory, string id,
                SelectionMadeEventHandler handler, bool isEnabled = true, bool isVisible = true) {
            var model = ViewModelFactoryExtensions.NewDropDownModel(factory, id, isEnabled, isVisible);

            model.SelectionMade += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonLabelModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static ILabelModel NewLabelModel(this IModelFactoryInternal factory, string id,
                ClickedEventHandler handler, bool isEnabled = true, bool isVisible = true)
        => factory?.NewLabelModel(id, isEnabled, isVisible)
                  ?.Attach(id);
    }
}

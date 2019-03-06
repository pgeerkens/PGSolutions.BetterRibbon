////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>These extension methods on <see cref="ViewModelFactory"/> are the public API to C# for creation of objects subclassing <see cref="ControlModel{TSource, TCtrl}"/>.</summary>
    internal static partial class PublicFactoryExtensions {
        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="IButtonModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IButtonModel NewButtonModel(this IModelFactory factory, string id,
                ClickedEventHandler handler, IImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewButtonModel(id, isEnabled, isVisible).SetImage(image);

            model.Clicked += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="IToggleModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IToggleModel NewToggleModel(this IModelFactory factory, string id,
                ToggledEventHandler handler, IImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewToggleModel(id, isEnabled, isVisible).SetImage(image);

            model.Toggled += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="IEditBoxModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IEditBoxModel NewEditBoxModel(this IModelFactory factory, string id,
                EditedEventHandler handler, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewEditBoxModel(id, isEnabled, isVisible);

            model.Edited += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="IComboBoxModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IComboBoxModel NewComboBoxModel(this IModelFactory factory, string id,
                EditedEventHandler handler,
                bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewComboBoxModel(id, isEnabled, isVisible);

            model.Edited += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="IDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IDropDownModel NewDropDownModel(this IModelFactory factory, string id,
                SelectionMadeEventHandler handler, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewDropDownModel(id, isEnabled, isVisible);

            model.SelectionMade += handler;
            return model?.Attach(id);
        }
    }
}

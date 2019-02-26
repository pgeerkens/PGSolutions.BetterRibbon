////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>These extension methods on <see cref="ViewModelFactory"/> are the public API to C# for creation of objects subclassing <see cref="ControlModel{TSource, TCtrl}"/>.</summary>
    public static partial class PublicFactoryExtensions {
        /// <summary>Returns a new instance of an <see cref="IModelFactory"/>.</summary>
        /// <param name="model"></param>
        public static IModelFactory NewModelFactory(this AbstractRibbonTabModel model)
            => new ModelFactory(model);

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonButtonModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IButtonModel NewButtonModel(this ViewModelFactory factory, string id,
                ClickedEventHandler handler, ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewButtonModel(factory.GetStrings2(id), image, isEnabled, isVisible);

            model.Clicked += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonToggleModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IToggleModel NewToggleModel(this ViewModelFactory factory, string id,
                ToggledEventHandler handler, ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewToggleModel(factory.GetStrings2(id), image, isEnabled, isVisible);

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
}

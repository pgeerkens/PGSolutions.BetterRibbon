////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>These extension methods on <see cref="ViewModelFactory"/> are the public API to C# for creation of objects subclassing <see cref="ControlModel{TSource, TCtrl}"/>.</summary>
    public static partial class PublicFactoryExtensions {
        /// <summary>Returns a new instance of an <see cref="IModelFactory"/>.</summary>
        /// <param name="model"></param>
        public static IModelFactory NewModelFactory(this ViewModelFactory viewModelFactory, IResourceLoader manager)
            => new ModelFactory(viewModelFactory, manager);

        public static AbstractModelFactory NewModelFactory2(this ViewModelFactory viewModelFactory, IResourceLoader manager)
            => new ModelFactory(viewModelFactory, manager);

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonButtonModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IButtonModel NewButtonModel(this AbstractModelFactory factory, string id,
                ClickedEventHandler handler, ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewButtonModel(id, image, isEnabled, isVisible);

            model.Clicked += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonToggleModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IToggleModel NewToggleModel(this AbstractModelFactory factory, string id,
                ToggledEventHandler handler, ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewToggleModel(id, image, isEnabled, isVisible);

            model.Toggled += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IEditBoxModel NewEditBoxModel(this AbstractModelFactory factory, string id,
                EditedEventHandler handler, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewEditBoxModel(id, isEnabled, isVisible);

            model.Edited += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IComboBoxModel NewComboBoxModel(this AbstractModelFactory factory, string id,
                EditedEventHandler handler,
                bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewComboBoxModel(id, isEnabled, isVisible);

            model.Edited += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static IDropDownModel NewDropDownModel(this AbstractModelFactory factory, string id,
                SelectionMadeEventHandler handler, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewDropDownModel(id, isEnabled, isVisible);

            model.SelectionMade += handler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonLabelModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static ILabelModel NewLabelModel(this AbstractModelFactory factory, string id,
                bool isEnabled = true, bool isVisible = true)
        => factory?.NewLabelModel(id, isEnabled, isVisible)
                  ?.Attach(id);
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IControlStrings;

    public static partial class PublicExtensions {
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
                SelectionMadeEventHandler selectedHandler, EditedEventHandler editedHandler,
                bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewComboBoxModel(factory.GetStrings(id), isEnabled, isVisible);

            model.Edited        += editedHandler;
            model.SelectionMade += selectedHandler;
            return model?.Attach(id);
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        public static SelectableItemModel NewSelectableModel(this ViewModelFactory factory,
                string controlID, IStrings strings) {
            var model = new SelectableItemModel(strings, true, true);

            model.Attach(controlID);
            return model;
        }
    }

    internal static partial class ViewModelFactoryExtensions {
        /// <summary>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</summary>
        /// <param name="ribbonXml"></param>
        public static IReadOnlyList<GroupVM> ParseXml(this ViewModelFactory factory, string ribbonXml) {
            //if (factory == null) throw new ArgumentNullException(nameof(factory));
            var groupModels = new List<GroupVM>();
            var doc = XDocument.Parse(ribbonXml);
            var root = doc.Root;
            XNamespace mso = ( from a in doc.Descendants().Attributes() 
                               where a.IsNamespaceDeclaration && a.Name.LocalName == "mso" 
                               select a
                             ).FirstOrDefault()?.Value;
            foreach (var group in root.Descendants(mso+"group")) {
                if (group.Attribute(mso+"idMso") != null  ||  group.Attribute(mso+"idQ") != null) continue;

                var viewModel = factory?.NewGroup(group.Attribute("id").Value);
                groupModels?.Add(viewModel);

                foreach (var element in group.Descendants()) {
                    if (element.Attribute(mso+"idMso") != null  ||  element.Attribute(mso+"idQ") != null) continue;

                    switch (element.Name) {
                        case XName name when name == mso+"toggleButton":
                            viewModel.Add<IToggleSource>(factory.NewToggleButton(element.Attribute("id").Value));
                            break;

                        case XName name when name == mso+"checkBox":
                            viewModel.Add<IToggleSource>(factory.NewCheckBox(element.Attribute("id").Value));
                            break;

                        case XName name when name == mso+"dropDown":
                            viewModel.Add<IDropDownSource>(factory.NewDropDown(element.Attribute("id").Value));
                            break;

                        case XName name when name == mso+"button":
                            viewModel.Add<IButtonSource>(factory.NewButton(element.Attribute("id").Value));
                            break;

                        case XName name when name == mso+"editBox":
                            viewModel.Add<IEditBoxSource>(factory.NewEditBox(element.Attribute("id").Value));
                            break;

                        case XName name when name == mso+"comboBox":
                            viewModel.Add<IComboBoxSource>(factory.NewComboBox(element.Attribute("id").Value));
                            break;

                        default:
                            break;
                    }
                }
            }

            return groupModels.AsReadOnly();
        }

        /// <summary>Creates, initializes and returns a new <see cref="GroupModel"/>.</summary>
        public static GroupModel NewGroupModel(this ViewModelFactory factory, IStrings strings,
                bool isEnabled, bool isVisible)
        => new GroupModel(factory.GetControl<GroupVM>, strings, isEnabled, isVisible);

        public static TModel InitializeModel<TSource, TVM, TModel>(this TModel model)
            where TModel : ControlModel<TSource, TVM> where TSource : IControlSource where TVM : IControlVM {

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

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
    }
}

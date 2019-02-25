////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

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

        public static TModel InitializeModel<TSource, TVM, TModel>(this TModel model)
            where TModel : ControlModel<TSource, TVM> where TSource : IControlSource where TVM : IControlVM {

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        /// <summary>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</summary>
        /// <param name="ribbonXml"></param>
        public static IReadOnlyList<TabVM> ParseXml(this ViewModelFactory factory, string ribbonXml) {
            var tabModels = new List<TabVM>();
            var doc = XDocument.Parse(ribbonXml);
            var root = doc.Root;
            XNamespace mso = ( from a in doc.Descendants().Attributes()
                               where a.IsNamespaceDeclaration && a.Name.LocalName == "mso"
                               select a
                             ).FirstOrDefault()?.Value;
            foreach (var tab in root.Descendants(mso+"tab")) {
                if (tab.Attribute("idMso") != null) {
                    tabModels?.Add(tab.ParseXmlChildren(mso, factory, factory?.NewTab(tab.Attribute("idMso").Value)));
                } else if(tab.Attribute("id") != null) {
                    tabModels?.Add(tab.ParseXmlChildren(mso, factory, factory?.NewTab(tab.Attribute("id").Value)));
                } else {
                    continue;
                }

            }

            return tabModels.AsReadOnly();
        }

        private static TCtrl ParseXmlChildren<TCtrl>(this XElement element, XNamespace mso,
                ViewModelFactory factory, TCtrl parent) where TCtrl: IContainerControl {
            foreach (var child in element.Elements()) {
                if (element.Attribute(mso+"idMso") != null  ||  element.Attribute(mso+"idQ") != null) continue;

                switch (child.Name) {
                    case XName name when name == mso+"toggleButton":
                        parent.Add(factory.NewToggleButton(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"checkBox":
                        parent.Add(factory.NewCheckBox(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"dropDown":
                        parent.Add(factory.NewDropDown(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"button":
                        parent.Add(factory.NewButton(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"editBox":
                        parent.Add(factory.NewEditBox(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"comboBox":
                        parent.Add(factory.NewComboBox(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"labelControl":
                        parent.Add(factory.NewLabel(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"box"
                                      || name == mso+"dialogBoxLauncher":
                        child.ParseXmlChildren(mso, factory, parent);
                        break;

                    case XName name when name == mso+"menu":
                        parent.Add(child.ParseXmlChildren(mso, factory,
                                factory.NewMenu(child.Attribute("id").Value)));
                        break;

                    case XName name when name == mso+"splitButton":
                        parent.Add(child.ParseXmlChildren(mso, factory,
                                factory.NewSplitButton(child.Attribute("id").Value)));
                        break;

                    case XName name when name == mso+"group":
                        parent.Add(child.ParseXmlChildren(mso, factory,
                                factory.NewGroup(child.Attribute("id").Value)) );
                        break;

                    case XName name when name == mso+"tab":
                        throw new InvalidOperationException($"Tab '{child.Name.LocalName}' found unexpectedly.");

                    default:
                        break;
                }
            }

            return parent;
        }
    }
}

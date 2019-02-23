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

    public static partial class RibbonFactoryExtensions {
        /// <summary>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</summary>
        /// <param name="ribbonXml"></param>
        public static IReadOnlyList<GroupVM> ParseXml(this IRibbonFactory factory, string ribbonXml) {
            if (factory == null) throw new ArgumentNullException(nameof(factory));
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
                            viewModel.Add<IRibbonToggleSource>(factory.NewToggleButton(element.Attribute("id").Value));
                            break;

                        case XName name when name == mso+"checkBox":
                            viewModel.Add<IRibbonToggleSource>(factory.NewCheckBox(element.Attribute("id").Value));
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

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonButtonModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static ButtonModel NewButtonModel(this IRibbonFactory factory, string id,
                ClickedEventHandler handler, ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewButtonModel(factory.GetStrings(id), image, isEnabled, isVisible);

            model?.Attach(id);
            model.Clicked += handler;
            return model;
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonToggleModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static ToggleModel NewToggleModel(this IRibbonFactory factory, string id,
                ToggledEventHandler handler, ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewToggleModel(factory.GetStrings(id), image, isEnabled, isVisible);

            model?.Attach(id);
            model.Toggled += handler;
            return model;
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static DropDownModel NewDropDownModel(this IRibbonFactory factory, string id,
                SelectedEventHandler handler, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewDropDownModel(factory.GetStrings(id), isEnabled, isVisible);

            model?.Attach(id);
            model.SelectionMade += handler;
            return model;
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public static EditBoxModel NewEditBoxModel(this IRibbonFactory factory, string id,
                EditedEventHandler handler, bool isEnabled = true, bool isVisible = true) {
            var model = factory?.NewEditBoxModel(factory.GetStrings(id), isEnabled, isVisible);

            model?.Attach(id);
            model.Edited += handler;
            return model;
        }


        /// <summary>Creates, initializes and returns a new <see cref="RibbonButtonModel"/>.</summary>
        internal static ButtonModel NewButtonModel(this IRibbonFactory factory, IStrings strings,
                ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = new ButtonModel(factory.GetControl<ButtonVM>, strings, image, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        /// <summary>Creates, initializes and returns a new <see cref="RibbonToggleModel"/>.</summary>
        internal static ToggleModel NewToggleModel(this IRibbonFactory factory, IStrings strings,
                ImageObject image, bool isEnabled = true, bool isVisible = true) {
            var model = new ToggleModel(factory.GetControl<CheckBoxVM>, strings, image, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        /// <summary>Creates, initializes and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        internal static DropDownModel NewDropDownModel(this IRibbonFactory factory, IStrings strings,
                bool isEnabled = true, bool isVisible = true) {
            var model = new DropDownModel(factory.GetControl<DropDownVM>, strings, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        /// <summary>Creates, initializes and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        internal static EditBoxModel NewEditBoxModel(this IRibbonFactory factory, IStrings strings,
                bool isEnabled = true, bool isVisible = true) {
            var model = new EditBoxModel(factory.GetControl<EditBoxVM>, strings, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        /// <summary>Creates, initializes and returns a new <see cref="RibbonGroupModel"/>.</summary>
        internal static GroupModel NewGroupModel(this IRibbonFactory factory, IStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => new GroupModel(factory.GetControl<GroupVM>, strings, isEnabled, isVisible);
    }
}

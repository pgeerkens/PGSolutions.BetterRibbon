////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Xml.Linq;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IRibbonControlStrings;

    public static partial class RibbonFactoryExtensions {
        /// <summary>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</summary>
        /// <param name="ribbonXml"></param>
        public static IReadOnlyList<RibbonGroupViewModel> ParseXml(this IRibbonFactory factory, string ribbonXml) {
            if (factory == null) throw new ArgumentNullException(nameof(factory));

            var groupModels = new List<RibbonGroupViewModel>();
            XNamespace mso2009 = "http://schemas.microsoft.com/office/2009/07/customui";
            foreach (var group in XDocument.Parse(ribbonXml).Root.Descendants(mso2009+"group")) {
                var viewModel = factory.AddGroupViewModel(groupModels, group.Attribute("id").Value);

                foreach (var element in group.Descendants()) {
                    switch (element.Name) {
                        case XName name when name == mso2009+"toggleButton":
                            viewModel.Add<IRibbonToggleSource>(factory.NewRibbonToggle(element.Attribute("id").Value));
                            break;
                        case XName name when name == mso2009+"checkBox":
                            viewModel.Add<IRibbonToggleSource>(factory.NewRibbonCheckBox(element.Attribute("id").Value));
                            break;
                        case XName name when name == mso2009+"dropDown":
                            viewModel.Add<IRibbonDropDownSource>(factory.NewRibbonDropDown(element.Attribute("id").Value));
                            break;
                        case XName name when name == mso2009+"button":
                            viewModel.Add<IRibbonButtonSource>(factory.NewRibbonButton(element.Attribute("id").Value));
                            break;
                        default:
                            break;
                    }
                }
            }

            return groupModels.AsReadOnly();
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonGroupViewModel"/>.</summary>
        public static RibbonGroupViewModel AddGroupViewModel(this IRibbonFactory factory,
                IList<RibbonGroupViewModel> groupModels, string groupName) {
            var viewModel = factory?.NewRibbonGroup(groupName);
            groupModels?.Add(viewModel);
            return viewModel;
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonButtonModel"/>.</summary>
        public static RibbonButtonModel NewRibbonButtonModel(this IRibbonFactory factory, string id,
                EventHandler handler, bool isEnabled, bool isVisible, ImageObject image) {
            var model = factory?.NewRibbonButtonModel(factory.GetStrings(id), image, isEnabled, isVisible);

            model?.Attach(id);
            model.Clicked += handler;
            return model;
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonToggleModel"/>.</summary>
        public static RibbonToggleModel NewRibbonToggleModel(this IRibbonFactory factory, string id,
                ToggledEventHandler handler, bool isEnabled, bool isVisible, ImageObject image) {
            var model = factory?.NewRibbonToggleModel(factory.GetStrings(id), image, isEnabled, isVisible);

            model?.Attach(id);
            model.Toggled += handler;
            return model;
        }

        /// <summary>Creates, initializes, attaches to the specified control view-model, and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        public static RibbonDropDownModel NewRibbonDropDownModel(this IRibbonFactory factory, string id,
                SelectedEventHandler handler, bool isEnabled, bool isVisible) {
            var model = factory?.NewRibbonDropDownModel(factory.GetStrings(id), isEnabled, isVisible);

            model?.Attach(id);
            model.SelectionMade += handler;
            return model;
        }

        /// <summary>Creates, initializes and returns a new <see cref="RibbonButtonModel"/>.</summary>
        internal static RibbonButtonModel NewRibbonButtonModel(this IRibbonFactory factory, IStrings strings,
                ImageObject image, bool isEnabled, bool isVisible) {
            var model = new RibbonButtonModel(factory.GetControl<RibbonButton>, strings, image, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        /// <summary>Creates, initializes and returns a new <see cref="RibbonToggleModel"/>.</summary>
        internal static RibbonToggleModel NewRibbonToggleModel(this IRibbonFactory factory, IStrings strings,
                ImageObject image, bool isEnabled, bool isVisible) {
            var model = new RibbonToggleModel(factory.GetControl<RibbonCheckBox>, strings, image, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        /// <summary>Creates, initializes and returns a new <see cref="RibbonDropDownModel"/>.</summary>
        internal static RibbonDropDownModel NewRibbonDropDownModel(this IRibbonFactory factory, IStrings strings,
                                 bool isEnabled, bool isVisible) {
            var model = new RibbonDropDownModel(factory.GetControl<RibbonDropDown>, strings, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        /// <summary>Creates, initializes and returns a new <see cref="RibbonGroupModel"/>.</summary>
        internal static RibbonGroupModel NewRibbonGroupModel(this IRibbonFactory factory, IStrings strings,
                bool isEnabled, bool isVisible)
        => new RibbonGroupModel(factory.GetControl<RibbonGroupViewModel>, strings, isEnabled, isVisible);
    }
}

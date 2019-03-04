////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Xml.Linq;
using PGSolutions.RibbonDispatcher.Models;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    using Trace = System.Diagnostics.Trace;

    /// <summary>.</summary>
    public static partial class XmParserExtensions {
        /// <summary>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</summary>
        /// <param name="ribbonXml"></param>
        public static ViewModelFactory ParseXmlTabs(this string ribbonXml)
        => XDocument.Parse(ribbonXml).Root.ParseXmlTabs();

        /// <summary>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</summary>
        /// <param name="ribbonXml"></param>
        public static ViewModelFactory ParseXmlTabs(this XElement root) {
            var factory = new ViewModelFactory();
            foreach (var tab in root.Descendants().Where(d => d.Name.LocalName == "tab")) {
                var name = tab.Attribute("idQ")?.Value?.Xns()
                        ?? tab.Attribute("idMso")?.Value
                        ?? tab.Attribute("id")?.Value;
                if (name != null) {
                    factory.TabViewModels.Add(tab.ParseXmlChildren(factory, factory?.NewTab(name)));
                }
            }
            return factory;
        }

        public static KeyedControls ParseXmlMenu(this XElement element, ViewModelFactory factory) {
            var controls = new KeyedControls();

            foreach (var child in element.Elements()) {
                var x = ParseXmlChild(child,factory);
                if (x != null) controls.Add(x);
            }

            return controls;
        }

        [SuppressMessage("Microsoft.Maintainability","CA1502:AvoidExcessiveComplexity")]
        public static TCtrl ParseXmlChildren<TCtrl>(this XElement element, ViewModelFactory factory,
                TCtrl parent) where TCtrl : IContainerControl {
            foreach (var child in element.Elements()) {
                var x = ParseXmlChild(child,factory);
                if (x != null) parent.Add(x);
            }
            return parent;
        }

        public static IControlVM ParseXmlChild(XElement child, ViewModelFactory factory) {
            string controlId = null;
            switch (child.Name.LocalName) {
                case string name when child.HasElements && StaticActions.TryGetValue(name,out var action):
                    if (TryGetControlId(child,ref controlId)) { return action(controlId,child,factory); }
                    break;

                case string name when Actions.TryGetValue(name,out var action):
                    if (TryGetControlId(child,ref controlId)) { return action(controlId,factory); }
                    break;

                case string name when name == "dialogBoxLauncher":
                    return ParseXmlChild(child.Elements().FirstOrDefault(), factory);

                case string name when name == "splitButton":
                    var menuId   = child.Elements().Last().Attribute("id").Value;
                    var menuVM   = child.Elements().Last().ParseXmlChildren(factory,factory.NewMenu(menuId));

                    var buttonId = child.Elements().First().Attribute("id").Value;

                    if (child.Elements().First().Name.LocalName == "button") {
                        return factory.NewSplitPressButton(child.Attribute("id").Value, menuVM,
                                    factory.NewButton(buttonId));
                    } else if (child.Elements().First().Name.LocalName == "toggleButton") {
                        return factory.NewSplitToggleButton(child.Attribute("id").Value, menuVM,
                                    factory.NewToggleButton(buttonId));
                    } else {
                        Trace.WriteLine($"Skipped a {child.Name.LocalName}: '{child.Attribute("id")}'");
                    }
                    break;

                default:
                    Trace.WriteLine($"Skipped a {child.Name.LocalName}: '{child.Attribute("id")}'");
                    break;
            }
            return null;
        }

        private static Dictionary<string,Func<string,XElement,ViewModelFactory,IControlVM>> StaticActions
            = new Dictionary<string,Func<string,XElement,ViewModelFactory,IControlVM>>() {
                {"dropDown",   (controlId,child,factory) => factory.NewStaticDropDown(controlId, child.Elements().ParseItemList()) },
                {"comboBox",   (controlId,child,factory) => factory.NewStaticComboBox(controlId, child.Elements().ParseItemList()) },
                {"gallery",    (controlId,child,factory) => factory.NewStaticGallery(controlId, child.Elements().ParseItemList()) },
                {"group",      (controlId,child,factory) => child.ParseXmlChildren(factory, factory.NewGroup(controlId)) },
                {"box",        (controlId,child,factory) => child.ParseXmlChildren(factory,factory.NewBoxControl(controlId))},
                {"menu",       (controlId,child,factory) => child.ParseXmlChildren(factory, factory.NewMenu(controlId)) }
            };

        private static Dictionary<string,Func<string,ViewModelFactory,IControlVM>> Actions
            = new Dictionary<string,Func<string,ViewModelFactory,IControlVM>>() {
                {"button",       (controlId,factory) => factory.NewButton(controlId) },
                {"gallery",      (controlId,factory) => factory.NewGallery(controlId) },
                {"editBox",      (controlId,factory) => factory.NewEditBox(controlId) },
                {"checkBox",     (controlId,factory) => factory.NewCheckBox(controlId) },
                {"dropDown",     (controlId,factory) => factory.NewDropDown(controlId) },
                {"comboBox",     (controlId,factory) => factory.NewComboBox(controlId) },
                {"dynamicMenu",  (controlId,factory) => factory.NewDynamicMenu(controlId) },
                {"toggleButton", (controlId,factory) => factory.NewToggleButton(controlId) },
                {"labelControl", (controlId,factory) => factory.NewLabelControl(controlId) },
                {"menuSeparator",(controlId,factory) => factory.NewMenuSeparator(controlId) }
            };

        private static bool TryGetControlId(XElement child, ref string controlId)
        => (controlId = child.Attribute("id")?.Value ?? child.Attribute("idQ")?.Value?.Xns())  !=  null;

        internal static IReadOnlyList<StaticItemVM> ParseItemList(this IEnumerable<XElement> elements) {
            var items = new List<StaticItemVM>();
            foreach (var child in elements) {
                switch (child.Name.LocalName) {
                    case string name when name == "item":
                        var id = child.Attribute("id").Value;
                        items.Add(new StaticItemVM(id,
                                new ControlStrings(child.Attribute("label")?.Value     ?? id,
                                                   child.Attribute("screentip")?.Value ?? "",
                                                   child.Attribute("supertip")?.Value  ?? "", null)
                        ));
                        break;
                    default:
                        Trace.WriteLine($"Skipped a {child.Name.LocalName}: '{child.Attribute("id")}' child of {child.Parent.Attribute("id")}");
                        break;
                }
            }
            return items;
        }
    }
}

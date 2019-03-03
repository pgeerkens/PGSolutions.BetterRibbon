////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Xml.Linq;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    using Trace = System.Diagnostics.Trace;

    /// <summary>.</summary>
    public static partial class XmParserExtensions {
        /// <summary>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</summary>
        /// <param name="ribbonXml"></param>
        public static ViewModelFactory ParseXmlTabs(this string ribbonXml)
        => XDocument.Parse(ribbonXml).ParseXmlTabs();

        /// <summary>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</summary>
        /// <param name="ribbonXml"></param>
        public static ViewModelFactory ParseXmlTabs(this XDocument doc) {
            var factory = new ViewModelFactory();
            var pg   = (XNamespace)"https://github.com/pgeerkens/PGSolutions.UtilityRibbon"; 
            var mso  = (XNamespace)( from a in doc.Descendants().Attributes()
                                     where a.IsNamespaceDeclaration && a.Name.LocalName == "mso"
                                     select a
                                   ).FirstOrDefault()?.Value;
            foreach (var tab in doc.Root.Descendants(mso+"tab")) {
                if (tab.Attribute("idMso") != null) {
                    factory.TabViewModels.Add(tab.ParseXmlChildren(mso, factory, factory?.NewTab(tab.Attribute("idMso").Value)));
                } else if (tab.Attribute("id") != null) {
                    factory.TabViewModels.Add(tab.ParseXmlChildren(mso, factory, factory?.NewTab(tab.Attribute("id").Value)));
                } else if (tab.Attribute("idQ") != null) {
                    factory.TabViewModels.Add(tab.ParseXmlChildren(mso, factory, factory?.NewTab(tab.Attribute("idQ").Value.Substring(3))));
                } else {
                    continue;
                }
            }
            return factory;
        }

        [SuppressMessage("Microsoft.Maintainability","CA1502:AvoidExcessiveComplexity")]
        public static TCtrl ParseXmlChildren<TCtrl>(this XElement element, XNamespace mso,
                ViewModelFactory factory, TCtrl parent) where TCtrl : IContainerControl {
            string controlId = null;
            foreach (var child in element.Elements()) {
                if (child.Attribute(mso+"idMso") != null  ||  child.Attribute(mso+"idQ") != null) continue;

                switch (child.Name) {
                    case XName name when child.HasElements && StaticActions.TryGetValue(name.LocalName,out var action):
                        if (TryGetControlId(child,ref controlId)) { parent.Add(action(controlId,child,factory,mso)); }
                        break;

                    case XName name when Actions.TryGetValue(name.LocalName,out var action):
                        if (TryGetControlId(child,ref controlId)) { parent.Add(action(controlId,factory)); }
                        break;

                    case XName name when name == mso+"dialogBoxLauncher":
                        child.ParseXmlChildren(mso, factory, parent);
                        break;

                    case XName name when name == mso+"splitButton":
                        var menuId   = child.Elements().Last().Attribute("id").Value;
                        var menuVM   = child.Elements().Last().ParseXmlChildren(mso, factory,
                                                            factory.NewMenu(menuId));

                        var buttonId = child.Elements().First().Attribute("id").Value;

                        if (child.Elements().First().Name == mso+"button") {
                            parent.Add(factory.NewSplitPressButton(child.Attribute("id").Value, menuVM,
                                        factory.NewButton(buttonId)));
                        } else if (child.Elements().First().Name == mso+"toggleButton") {
                            parent.Add(factory.NewSplitToggleButton(child.Attribute("id").Value, menuVM,
                                        factory.NewToggleButton(buttonId)));
                        } else throw new InvalidOperationException($"Encountered invalid control type: '{child.Elements().First().Name}'");
                        break;

                    default:
                        Trace.WriteLine($"Skipped a {child.Name.LocalName}: '{child.Attribute("id")}'");
                        break;
                }
            }
            return parent;
        }

        private static Dictionary<string,Func<string,XElement,ViewModelFactory,XNamespace,IControlVM>> StaticActions
            = new Dictionary<string,Func<string,XElement,ViewModelFactory,XNamespace,IControlVM>>() {
                {"dropDown",(controlId,child,factory,mso) => factory.NewStaticDropDown(controlId, child.ParseItemList(mso)) },
                {"comboBox",(controlId,child,factory,mso) => factory.NewStaticComboBox(controlId, child.ParseItemList(mso)) },
                {"gallery", (controlId,child,factory,mso) => factory.NewStaticGallery(controlId, child.ParseItemList(mso)) },
                {"group",   (controlId,child,factory,mso) => child.ParseXmlChildren(mso, factory, factory.NewGroup(controlId)) },
                {"box",     (controlId,child,factory,mso) => child.ParseXmlChildren(mso,factory,factory.NewBoxControl(controlId))},
                {"menu",    (controlId,child,factory,mso) => child.ParseXmlChildren(mso, factory, factory.NewMenu(controlId)) }
            };

        private static Dictionary<string,Func<string,ViewModelFactory,IControlVM>> Actions
            = new Dictionary<string,Func<string,ViewModelFactory,IControlVM>>() {
                {"button",       (controlId,factory) => factory.NewButton(controlId) },
                {"gallery",      (controlId,factory) => factory.NewGallery(controlId) },
                {"editBox",      (controlId,factory) => factory.NewEditBox(controlId) },
                {"checkBox",     (controlId,factory) => factory.NewCheckBox(controlId) },
                {"dropDown",     (controlId,factory) => factory.NewDropDown(controlId) },
                {"comboBox",     (controlId,factory) => factory.NewComboBox(controlId) },
                {"toggleButton", (controlId,factory) => factory.NewToggleButton(controlId) },
                {"labelControl", (controlId,factory) => factory.NewLabelControl(controlId) },
                {"menuSeparator",(controlId,factory) => factory.NewMenuSeparator(controlId) }
            };

        private static bool TryGetControlId(XElement child, ref string controlId)
        => (controlId = child.Attribute("id")?.Value ?? child.Attribute("idQ")?.Value)  !=  null;

        internal static IReadOnlyList<StaticItemVM> ParseItemList(this XElement parent, XNamespace mso) {
            var items = new List<StaticItemVM>();
            foreach (var child in parent.Elements()) {
                if (parent.Attribute(mso+"idMso") != null  ||  parent.Attribute(mso+"idQ") != null) continue;

                switch (child.Name) {
                    case XName name when name == mso+"item":
                        var id = child.Attribute("id").Value;
                        items.Add(new StaticItemVM(id,
                                new ControlStrings(child.Attribute("label")?.Value     ?? id,
                                                   child.Attribute("screentip")?.Value ?? "",
                                                   child.Attribute("supertip")?.Value  ?? "", null)
                        ));
                        break;
                    default:
                        Trace.WriteLine($"Skipped a {child.Name.LocalName}: '{child.Attribute("id")}' child of {parent.Attribute("id")}");
                        break;
                }
            }
            return items;
        }
    }
}

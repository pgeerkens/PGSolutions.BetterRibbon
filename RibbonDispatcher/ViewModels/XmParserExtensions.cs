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
    internal static partial class XmParserExtensions {
        /// <summary>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</summary>
        /// <param name="ribbonXml"></param>
        public static void ParseXmlTabs(this IList<TabVM> tabModels, ViewModelFactory factory, string ribbonXml) {
            var doc  = XDocument.Parse(ribbonXml);
            var root = doc.Root;
            var mso  = (XNamespace)( from a in doc.Descendants().Attributes()
                                     where a.IsNamespaceDeclaration && a.Name.LocalName == "mso"
                                     select a
                                   ).FirstOrDefault()?.Value;
            foreach (var tab in root.Descendants(mso+"tab")) {
                if (tab.Attribute("idMso") != null) {
                    tabModels?.Add(tab.ParseXmlChildren(mso, factory, factory?.NewTab(tab.Attribute("idMso").Value)));
                } else if (tab.Attribute("id") != null) {
                    tabModels?.Add(tab.ParseXmlChildren(mso, factory, factory?.NewTab(tab.Attribute("id").Value)));
                } else {
                    continue;
                }
            }
        }

        [SuppressMessage("Microsoft.Maintainability","CA1502:AvoidExcessiveComplexity")]
        public static TCtrl ParseXmlChildren<TCtrl>(this XElement element, XNamespace mso,
                ViewModelFactory factory, TCtrl parent) where TCtrl : IContainerControl {
            foreach (var child in element.Elements()) {
                if (element.Attribute(mso+"idMso") != null  ||  element.Attribute(mso+"idQ") != null) continue;

                switch (child.Name) {
                    case XName name when name == mso+"toggleButton":
                        parent.Add(factory.NewToggleButton(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"checkBox":
                        parent.Add(factory.NewCheckBox(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"dropDown"  &&  child.HasElements:
                        parent.Add(factory.NewStaticDropDown(child.Attribute("id").Value,
                                    child.ParseItemList(mso)));
                        break;

                    case XName name when name == mso+"dropDown":
                        parent.Add(factory.NewDropDown(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"comboBox"  &&  child.HasElements:
                        parent.Add(factory.NewStaticComboBox(child.Attribute("id").Value,
                                    child.ParseItemList(mso)));
                        break;

                    case XName name when name == mso+"comboBox":
                        parent.Add(factory.NewComboBox(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"gallery"  &&  child.HasElements:
                        parent.Add(factory.NewStaticGallery(child.Attribute("id").Value,
                                    child.ParseItemList(mso)));
                        break;

                    case XName name when name == mso+"gallery":
                        parent.Add(factory.NewGallery(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"button":
                        parent.Add(factory.NewButton(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"editBox":
                        parent.Add(factory.NewEditBox(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"labelControl":
                        parent.Add(factory.NewLabelControl(child.Attribute("id").Value));
                        break;

                    case XName name when name == mso+"menuSeparator":
                        parent.Add(factory.NewMenuSeparator(child.Attribute("id").Value));
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

                    case XName name when name == mso+"group":
                        parent.Add(child.ParseXmlChildren(mso, factory,
                                factory.NewGroup(child.Attribute("id").Value)));
                        break;

                    case XName name when name == mso+"tab":
                        throw new InvalidOperationException($"Tab '{child.Name.LocalName}' found unexpectedly.");

                    default:
                        Trace.WriteLine($"Skipped a {child.Name.LocalName}: '{child.Attribute("id")}'");
                        break;
                }
            }
            return parent;
        }

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

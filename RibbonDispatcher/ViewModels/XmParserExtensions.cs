////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    using Trace = System.Diagnostics.Trace;

    /// <summary>The methods to construct a View-Model hierarchy from an XML ribbon definition.</summary>
    public static partial class XmParserExtensions {
        /// <summary>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</summary>
        /// <param name="ribbonXml"></param>
        public static ViewModelFactory ParseXmlTabs(this string ribbonXml)
        => ViewModelFactory.ParseXmlDoc(XDocument.Parse(ribbonXml).Root);

        internal static KeyedControls ParseXmlMenu(this XElement element, ViewModelFactory factory) {
            var controls = new KeyedControls();

            foreach (var child in element.Elements()) {
                var x = ParseXmlChild(child,factory);
                if (x != null) controls.Add(x);
            }

            return controls;
        }

        internal static IReadOnlyList<IControlVM> ParseXmlChildren(this XElement element, ViewModelFactory factory)
        => (from child in element.Elements() select ParseXmlChild(child,factory)).ToList();

        private static IReadOnlyList<StaticItemVM> ParseItemList(this IEnumerable<XElement> elements)
        => (from child in elements select child.ParseXmlChild(null) as StaticItemVM).ToList();

        private static IControlVM ParseXmlChild(this XElement child, ViewModelFactory factory) {
            string controlId = null;
            switch (child.Name.LocalName) {
                case string name when Actions.TryGetValue(name,out var action):
                    if (TryGetControlId(child,ref controlId)) { return action(controlId,child,factory); }
                    break;

                // The dialogBoxLauncher control has no ControlId - so special handling needed
                case string name when name == "dialogBoxLauncher":
                    return child.Elements().FirstOrDefault().ParseXmlChild(factory);

                // And our friend the SplitButton is just very, very, special
                case string name when name == "splitButton":
                    var menu = child.Elements().Last().ParseXmlChild(factory) as IMenuVM;

                    switch(child.Elements().First().ParseXmlChild(factory)) {
                        case IButtonVM button:
                            return factory.NewSplitPressButton(child.Attribute("id").Value, menu, button);
                        case IToggleVM toggle:
                            return factory.NewSplitToggleButton(child.Attribute("id").Value, menu, toggle);
                    }
                    Trace.WriteLine($"Skipped a {child.Name.LocalName}: '{child.Attribute("id")}'");
                    break;

                default:
                    Trace.WriteLine($"Skipped a {child.Name.LocalName}: '{child.Attribute("id")}'");
                    break;
            }
            return null;
        }

        private static Dictionary<string,Func<string,XElement,ViewModelFactory,IControlVM>> Actions
            = new Dictionary<string,Func<string,XElement,ViewModelFactory,IControlVM>>() {
                {"dropDown",   (controlId,element,factory) => factory.NewDropDown(controlId,element.Elements().ParseItemList()) },
                {"comboBox",   (controlId,element,factory) => factory.NewComboBox(controlId,element.Elements().ParseItemList()) },
                {"gallery",    (controlId,element,factory) => factory.NewGallery(controlId,element.Elements().ParseItemList()) },
                {"group",      (controlId,element,factory) => factory.NewGroup(controlId,element.ParseXmlChildren(factory)) },
                {"menu",       (controlId,element,factory) => factory.NewMenu(controlId,element.ParseXmlChildren(factory)) },
                {"box",        (controlId,element,factory) => factory.NewBox(controlId,element.ParseXmlChildren(factory)) },
                {"tab",        (controlId,element,factory) => factory.NewTab(controlId,element.ParseXmlChildren(factory)) },
                {"buttonGroup",(controlId,element,factory) => factory.NewButtonGroup(controlId,element.ParseXmlChildren(factory)) },

                {"button",       (controlId,element,factory) => factory.NewButton(controlId) },
                {"editBox",      (controlId,element,factory) => factory.NewEditBox(controlId) },
                {"checkBox",     (controlId,element,factory) => factory.NewCheckBox(controlId) },
                {"dynamicMenu",  (controlId,element,factory) => factory.NewDynamicMenu(controlId) },
                {"toggleButton", (controlId,element,factory) => factory.NewToggleButton(controlId) },
                {"labelControl", (controlId,element,factory) => factory.NewLabelControl(controlId) },
                {"menuSeparator",(controlId,element,factory) => factory.NewMenuSeparator(controlId) },

                {"item", (controlId,element,factory) => new StaticItemVM(controlId,
                               new ControlStrings(element.Attribute("label")?.Value     ?? controlId,
                                                  element.Attribute("screentip")?.Value ?? "",
                                                  element.Attribute("supertip")?.Value  ?? "", null)) }
            };

        private static bool TryGetControlId(XElement child, ref string controlId)
        => (controlId = child.Attribute("id")?.Value
                     ?? child.Attribute("idMso")?.Value
                     ?? child.Attribute("idQ")?.Value?.XNS())  !=  null;
    }
}

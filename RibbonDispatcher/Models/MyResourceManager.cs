////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Reflection;
using System.Resources;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    public class MyResourceManager: IResourceLoader {
        public MyResourceManager() : this(Assembly.GetCallingAssembly()) { }

        private MyResourceManager(Assembly assembly)
        => ResourceManager = new ResourceManager(
                $"{assembly.GetName().Name}.Properties.Resources", assembly);

        protected ResourceManager ResourceManager { get; }

        /// <inheritdoc/>
        public IControlStrings GetControlStrings(string ControlId) =>
            new ControlStrings(
                    GetCurrentUIString($"{ControlId.XNS()}_Label")     ?? ControlId.Unknown(),
                    GetCurrentUIString($"{ControlId.XNS()}_ScreenTip") ?? ControlId.Unknown("ScreenTip"),
                    GetCurrentUIString($"{ControlId.XNS()}_SuperTip")  ?? ControlId.Unknown("SuperTip"),
                    GetCurrentUIString($"{ControlId.XNS()}_KeyTip")    ?? ""
            );

        /// <inheritdoc/>
        public IControlStrings2 GetControlStrings2(string ControlId) =>
            new ControlStrings2(
                    GetCurrentUIString($"{ControlId.XNS()}_Label")       ?? ControlId.Unknown(),
                    GetCurrentUIString($"{ControlId.XNS()}_ScreenTip")   ?? ControlId.Unknown("ScreenTip"),
                    GetCurrentUIString($"{ControlId.XNS()}_SuperTip")    ?? ControlId.Unknown("SuperTip"),
                    GetCurrentUIString($"{ControlId.XNS()}_KeyTip")      ?? "",
                    GetCurrentUIString($"{ControlId.XNS()}_Description") ?? ControlId.Unknown("Description")
            );

        /// <inheritdoc/>
        public IPictureDisp GetImage(string Name) => ResourceManager.GetResourceImage(Name);

        protected string GetCurrentUIString(string controlId) => ResourceManager.GetCurrentUIString(controlId);
    }
}

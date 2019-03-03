////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Reflection;
using System.Resources;
using PGSolutions.RibbonDispatcher.ComInterfaces;

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
                    GetCurrentUIString($"{ControlId.Xns()}_Label")     ?? ControlId.Unknown(),
                    GetCurrentUIString($"{ControlId.Xns()}_ScreenTip") ?? ControlId.Unknown("ScreenTip"),
                    GetCurrentUIString($"{ControlId.Xns()}_SuperTip")  ?? ControlId.Unknown("SuperTip"),
                    GetCurrentUIString($"{ControlId.Xns()}_KeyTip")    ?? ""
            );

        /// <inheritdoc/>
        public IControlStrings2 GetControlStrings2(string ControlId) =>
            new ControlStrings2(
                    GetCurrentUIString($"{ControlId.Xns()}_Label")       ?? ControlId.Unknown(),
                    GetCurrentUIString($"{ControlId.Xns()}_ScreenTip")   ?? ControlId.Unknown("ScreenTip"),
                    GetCurrentUIString($"{ControlId.Xns()}_SuperTip")    ?? ControlId.Unknown("SuperTip"),
                    GetCurrentUIString($"{ControlId.Xns()}_KeyTip")      ?? "",
                    GetCurrentUIString($"{ControlId.Xns()}_Description") ?? ControlId.Unknown("Description")
            );

        /// <inheritdoc/>
        public object GetImage(string Name) => ResourceManager.GetResourceImage(Name);

        protected string GetCurrentUIString(string controlId) => ResourceManager.GetCurrentUIString(controlId);
    }
}

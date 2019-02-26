////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Reflection;
using System.Resources;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    public class MyResourceManager: IResourceManager {
        public MyResourceManager() : this(Assembly.GetCallingAssembly()) { }

        private MyResourceManager(Assembly assembly)
        => ResourceManager = new ResourceManager(
                $"{assembly.GetName().Name}.Properties.Resources", assembly);

        protected ResourceManager ResourceManager { get; }

        /// <inheritdoc/>
        public IControlStrings GetControlStrings(string ControlId) =>
            new ControlStrings(
                    GetCurrentUIString($"{ControlId}_Label")          ?? ControlId.Unknown(),
                    GetCurrentUIString($"{ControlId}_ScreenTip")      ?? ControlId.Unknown("ScreenTip"),
                    GetCurrentUIString($"{ControlId}_SuperTip")       ?? ControlId.Unknown("SuperTip"),
                    GetCurrentUIString($"{ControlId}_KeyTip")         ?? ""
            );

        /// <inheritdoc/>
        public IControlStrings2 GetControlStrings2(string ControlId) =>
            new ControlStrings2(
                    GetCurrentUIString($"{ControlId}_Label")          ?? ControlId.Unknown(),
                    GetCurrentUIString($"{ControlId}_ScreenTip")      ?? ControlId.Unknown("ScreenTip"),
                    GetCurrentUIString($"{ControlId}_SuperTip")       ?? ControlId.Unknown("SuperTip"),
                    GetCurrentUIString($"{ControlId}_KeyTip")         ?? "",
                    GetCurrentUIString($"{ControlId}_Description")    ?? ControlId.Unknown("Description")
            );

        /// <inheritdoc/>
        public object GetImage(string Name) => ResourceManager.GetResourceImage(Name);

        protected string GetCurrentUIString(string controlId) => ResourceManager.GetCurrentUIString(controlId);
    }
}

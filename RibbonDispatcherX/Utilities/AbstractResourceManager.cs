////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Resources;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;

namespace PGSolutions.RibbonDispatcher.Utilities {
    public abstract class AbstractResourceManager : IResourceManager {
        public AbstractResourceManager(string assemblyName) => _assemblyName = assemblyName;

        protected readonly string _assemblyName;

        protected abstract Lazy<ResourceManager> ResourceManager { get; }
        
        /// <inheritdoc/>
        public IRibbonTextLanguageControl GetControlStrings(string ControlId) =>
            new RibbonTextLanguageControl(
                    GetCurrentUIString($"{ControlId}_Label")          ?? ControlId.Unknown(),
                    GetCurrentUIString($"{ControlId}_ScreenTip")      ?? ControlId.Unknown("ScreenTip"),
                    GetCurrentUIString($"{ControlId}_SuperTip")       ?? ControlId.Unknown("SuperTip"),
                    GetCurrentUIString($"{ControlId}_KeyTip")         ?? "",
                    GetCurrentUIString($"{ControlId}_AlternateLabel") ?? ControlId.Unknown("Alternate"),
                    GetCurrentUIString($"{ControlId}_Description")    ?? ControlId.Unknown("Description")
            );

        /// <inheritdoc/>
        public object GetImage(string Name) => ResourceManager.Value.GetResourceImage(Name);

        protected string GetCurrentUIString(string controlId) => ResourceManager.Value.GetCurrentUIString(controlId);
    }
}

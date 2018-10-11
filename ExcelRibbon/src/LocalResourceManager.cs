////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Reflection;
using System.Resources;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.ExcelRibbon {
    internal class LocalResourceManager : AbstractResourceManager {
        public  LocalResourceManager(string assemblyName) : base(assemblyName) { }

        protected override Lazy<ResourceManager> ResourceManager => new Lazy<ResourceManager>(
            () => new ResourceManager($"{_assemblyName}.Properties.Resources", Assembly.GetExecutingAssembly())
        );
    }
}

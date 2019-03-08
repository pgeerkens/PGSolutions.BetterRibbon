////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The publicly available entry points to the library.</summary>
    [SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
            Justification = "Public, Non-Creatable, class with exported methods.")]
    [Serializable, CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICustomRibbonComEntry))]
    [Guid(Guids.CustomRibbonComEntry)]
    public sealed class CustomRibbonComEntry : ICustomRibbonComEntry {
        public static ICustomRibbonComEntry New(CustomDispatcher dispatcher)
        => new CustomRibbonComEntry(dispatcher);

        private CustomRibbonComEntry(CustomDispatcher dispatcher) => Dispatcher = dispatcher;

        private CustomDispatcher Dispatcher { get; }

        /// <inheritdoc/>
        public IModelFactory NewBetterRibbon(IResourceLoader loader) => Dispatcher.NewModelFactory(loader);

        /// <inheritdoc/>
        public IModelServer NewModelServer(IResourceLoader loader)   => Dispatcher.NewModelServer(loader);

        /// <inheritdoc/>
        public void RegisterWorkbook(string workbookName) => Dispatcher.RegisterWorkbook(workbookName);
    }
}

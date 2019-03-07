////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Models;

namespace PGSolutions.BetterRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
            Justification = "Public, Non-Creatable, class with exported methods.")]
    [Serializable, CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICustomRibbonComEntry))]
    [Guid(RibbonDispatcher.Guids.CustomRibbonComEntry)]
    public sealed class ComEntry : ICustomRibbonComEntry {
        public ComEntry(AbstractDispatcher dispatcher) => Dispatcher = dispatcher;

        AbstractDispatcher Dispatcher { get; }

        /// <inheritdoc/>
        public IModelFactory NewBetterRibbon(IResourceLoader manager) => Dispatcher.NewModelFactory(manager);

        /// <inheritdoc/>
        public IModelServer NewModelServer(IResourceLoader manager) => Dispatcher.NewModelFactory(manager) as IModelServer;

        /// <inheritdoc/>
        public void RegisterWorkbook(string workbookName) => Dispatcher.RegisterWorkbook(workbookName);
    }
}

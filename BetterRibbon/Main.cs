////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonUtilities.LinksAnalysis;

namespace PGSolutions.BetterRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable" )]
    [Serializable, CLSCompliant(false)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IBetterRibbon))]
    [Guid(RibbonDispatcher.Guids.BetterRibbonMain)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    public sealed class Main : IBetterRibbon {
        internal Main(Func<IModelFactory> funcFactory) => FuncFactory = funcFactory;

        Func<IModelFactory> FuncFactory { get; }

         /// <inheritdoc/>
        public IModelFactory NewBetterRibbon() => FuncFactory();

        /// <inheritdoc/>
        public ILinksAnalyzer NewLinksAnalyzer() => new LinksAnalyzer();
    }
}

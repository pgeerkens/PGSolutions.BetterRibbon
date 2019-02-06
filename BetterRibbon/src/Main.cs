////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonUtilities.LinksAnalyzer;

namespace PGSolutions.BetterRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable" )]
    [Serializable, CLSCompliant(false)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComDefaultInterface(typeof(IBetterRibbon))]
    [Guid(RibbonDispatcher.Guids.BetterRibbonMain)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    public sealed class Main : IBetterRibbon {
        internal Main() { }

         /// <inheritdoc/>
        public IRibbonDispatcher NewBetterRibbon()  => Model as IRibbonDispatcher;
        /// <inheritdoc/>
        public ILinksAnalyzer NewLinksAnalyzer() => new LinksAnalyzer();

        private readonly BetterRibbonModel Model = new BetterRibbonModel(Globals.ThisAddIn.ViewModel);
    }
}

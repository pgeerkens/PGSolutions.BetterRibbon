////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonUtilities.LinksAnalysis;

namespace PGSolutions.UtilityRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable" )]
    [Serializable, CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IComEntry))]
    [Guid(RibbonDispatcher.Guids.IComEntry)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    public sealed class ComEntry : IComEntry {
        internal ComEntry() { }

        /// <inheritdoc/>
        [CLSCompliant(false)]
        public ILinksAnalyzer NewLinksAnalyzer() => new LinksAnalyzer();
    }
}

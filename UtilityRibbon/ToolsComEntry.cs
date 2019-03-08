////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonUtilities.LinksAnalysis;

namespace PGSolutions.ToolsRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable" )]
    [Serializable, CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IToolsComEntry))]
    [Guid(Guids.IToolsComEntry)]
    public sealed class ToolsComEntry : IToolsComEntry {
        internal static IToolsComEntry New() => new ToolsComEntry();

        internal ToolsComEntry() { }

        /// <inheritdoc/>
        [CLSCompliant(false)]
        public ILinksAnalyzer NewLinksAnalyzer() => new LinksAnalyzer();
    }
}

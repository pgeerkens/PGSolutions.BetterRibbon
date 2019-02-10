////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    using Range = Microsoft.Office.Interop.Excel.Range;
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;
    using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

    /// <summary>The publicly available entry points to the library.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Serializable, CLSCompliant(false)]
    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ILinksAnalyzer))]
    [Guid(Guids.LinksAnalyzer)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    public sealed class LinksAnalyzer : ILinksAnalyzer {
        public LinksAnalyzer() { }

        /// <inheritdoc/>
        public ILinksLexer NewLinksLexer(ISourceCellRef cellRef, string formula)
             => new LinksLexer(cellRef, formula);

        /// <inheritdoc/>
        public ILinksAnalysis NewExternalLinks(ILinksAnalysisViewModel viewModel, Range range)
            => new LinksParser(viewModel, range);

        /// <inheritdoc/>
        public ILinksAnalysis NewExternalLinksWB(Workbook wb, string excludedName)
            => new LinksParser(wb, excludedName);

        /// <inheritdoc/>
        public ILinksAnalysis NewExternalLinksWS(Worksheet ws)
            => new LinksParser(ws);

        /// <inheritdoc/>
        public ILinksAnalysis Parse(ISourceCellRef cellRef, string formula)
            => new LinksParser(cellRef, formula);
    }
}

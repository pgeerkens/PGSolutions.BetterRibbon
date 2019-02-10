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
    [ComVisible(true)]
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
        public IExternalLinks NewExternalLinks(ILinksAnalysisViewModel viewModel, Range range)
            => new ExternalLinks(viewModel, range);

        /// <inheritdoc/>
        public IExternalLinks NewExternalLinksWB(Workbook wb, string excludedName)
            => new ExternalLinks(wb, excludedName);

        /// <inheritdoc/>
        public IExternalLinks NewExternalLinksWS(Worksheet ws)
            => new ExternalLinks(ws);

        /// <inheritdoc/>
        public IExternalLinks Parse(ISourceCellRef cellRef, string formula)
            => new ExternalLinks(cellRef, formula);
    }
}

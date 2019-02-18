////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;
    using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

    /// <summary>The publicly available entry points to the library.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Serializable, CLSCompliant(false)]
    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ILinksAnalyzer))]
    [Guid(Guids.LinksAnalyzer)]
    [ProgId(ProgIds.RibbonUtilitiesProgId)]
    public sealed class LinksAnalyzer : ILinksAnalyzer {
        public LinksAnalyzer() { }

        /// <inheritdoc/>
        public ILinksLexer NewLinksLexer(ISourceCellRef cellRef, string formula)
             => new LinksLexer(cellRef, formula);

        /// <inheritdoc/>
        public ILinksAnalysis NewExternalLinksWB(Workbook wb, IList<string> excludedNames)
            => new FormulaParser().ParseWorkbook(wb, excludedNames);

        /// <inheritdoc/>
        public ILinksAnalysis NewExternalLinksWS(Worksheet ws)
            => new FormulaParser().ParseWorksheet(ws);

        /// <inheritdoc/>
        public ILinksAnalysis Parse(ISourceCellRef cellRef, string formula)
            => new FormulaParser().ParseFormula(cellRef, formula);
    }
}

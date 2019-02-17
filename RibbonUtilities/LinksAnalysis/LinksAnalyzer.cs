////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using PGSolutions.RibbonUtilities.LinksAnalysis;
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
            => new LinksParser(wb, excludedNames);

        /// <inheritdoc/>
        public ILinksAnalysis NewExternalLinksWS(Worksheet ws)
            => new LinksParser(ws);

        /// <inheritdoc/>
        public ILinksAnalysis Parse(ISourceCellRef cellRef, string formula)
            => new LinksParser(cellRef, formula);
    }
}
namespace PGSolutions.RibbonUtilities {
    using Microsoft.Office.Interop.Excel;
    using PGSolutions.RibbonUtilities.LinksAnalysis;
    using PGSolutions.RibbonUtilities.VbaSourceExport;

    /// <summary>.</summary>
    public class RibbonUtilitiesEntryPoint : IRibbonUtilities {
        /// <inheritdoc/>
        public ILinksAnalyzer NewLinksAnalyzer() => new LinksAnalyzer();

        public VbaSourceExporter NewVbaSourceExporter() => new VbaSourceExporter(ExcelApp());

        private static Application ExcelApp() {

            return new Application();
        }       
    }

    /// <summary>.</summary>
    public interface IRibbonUtilities {
        /// <summary>.</summary>
        ILinksAnalyzer NewLinksAnalyzer();

    }

    /// <summary>Static clas of ProgIds</summary>
    public static class ProgIds {
        /// <summary>ProgID for the Ribbon dispatcher.</summary>
        public const string RibbonUtilitiesProgId      = "PGSolutions.RibbonUtilities";
    }
}

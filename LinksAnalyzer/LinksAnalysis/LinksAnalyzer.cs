////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Range = Microsoft.Office.Interop.Excel.Range;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Application = Microsoft.Office.Interop.Excel.Application;
using PGSolutions.RibbonDispatcher.ComClasses;

using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    /// <summary>The publicly available entry points to the library.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Serializable, CLSCompliant(false)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ILinksAnalyzer))]
    [Guid(Guids.LinksAnalyzer)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    public sealed class LinksAnalyzer : ILinksAnalyzer {
        public LinksAnalyzer(Application application) => Application = application;

        Application Application { get; }

        /// <inheritdoc/>
        public ILinksLexer NewLinksLexer(ISourceCellRef cellRef, string formula)
             => new LinksLexer(cellRef, formula);

        /// <inheritdoc/>
        public IExternalLinks NewExternalLinks(Application excel, INameList nameList)
            => new ExternalLinks(Application, nameList);

        /// <inheritdoc/>
        public IExternalLinks NewExternalLinksWB(Workbook wb, string excludedName)
            => new ExternalLinks(wb, excludedName);

        /// <inheritdoc/>
        public IExternalLinks NewExternalLinksWS(Worksheet ws)
            => new ExternalLinks(ws);

        /// <inheritdoc/>
        public IExternalLinks Parse(ISourceCellRef cellRef, string formula)
            => new ExternalLinks(cellRef, formula);

        /// <inheritdoc/>
        public void WriteLinksAnalysisWB(Workbook wb)
            => wb.WriteLinks();

        /// <inheritdoc/>
        public void WriteLinksAnalysisFiles(Workbook wb, Range range)
            => wb.WriteLinks(range.GetNameList());
    }
}

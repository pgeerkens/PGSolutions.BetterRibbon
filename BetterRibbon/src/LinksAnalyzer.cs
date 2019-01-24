////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using PGSolutions.RibbonDispatcher.ComClasses;

using PGSolutions.LinksAnalyzer;
using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.BetterRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Serializable, CLSCompliant(false)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ILinksAnalyzer))]
    [Guid(Guids.LinksAnalyzer)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    public sealed class LinksAnalyzer : ILinksAnalyzer {
        internal LinksAnalyzer() {}

        #region ILinksAnalyzer methods
        /// <inheritdoc/>
        public ILinksLexer NewLinksLexer(ISourceCellRef cellRef, string formula)
             => new LinksLexer(cellRef, formula);

        /// <inheritdoc/>
        public IExternalLinks NewExternalLinks(Excel.Application excel, INameList nameList)
            => new ExternalLinks(Globals.ThisAddIn.Application, nameList);

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
        public void WriteLinksAnalysisFiles(Workbook wb, INameList nameList)
            => wb.WriteLinks(nameList);
        #endregion
    }
}
